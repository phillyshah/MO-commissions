"""Every 15 minutes, write the event_log chain head somewhere this box cannot rewrite (AUD-1).

Why Object Lock in COMPLIANCE mode: the threat is the database owner — the builder — editing the
log and recomputing the hashes. A chain alone cannot detect that. A copy of the head in a bucket
the builder cannot delete from (compliance mode: not even the root account, until retention ends)
can: any rewrite before an anchored row changes that row's hash and no longer matches the anchor.

Why hand-rolled SigV4: the worker image stays small and the whole request is one PUT. The signing
is ~40 lines and a pure function, so it is tested without the network.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import json
import logging
from pathlib import Path
from urllib.parse import quote

import httpx

from accounting import APP_VERSION
from accounting.config import settings
from accounting.db.pool import connection
from accounting.lifecycle import chain_head

log = logging.getLogger("accounting.audit.anchors")

LOCAL_DIR = Path("var/anchors")
RETENTION_YEARS = 7


def _sign(key: bytes, msg: str) -> bytes:
    return hmac.new(key, msg.encode(), hashlib.sha256).digest()


def sigv4_headers(*, method: str, host: str, path: str, payload: bytes, headers: dict[str, str],
                  region: str, service: str, access_key: str, secret_key: str, now: dt.datetime) -> dict[str, str]:
    """AWS Signature Version 4 for one request. Returns the headers to send (input headers plus
    Host, X-Amz-Date, X-Amz-Content-Sha256 and Authorization). Pure: no clock, no network."""
    amz_date = now.strftime("%Y%m%dT%H%M%SZ")
    date_stamp = now.strftime("%Y%m%d")
    payload_hash = hashlib.sha256(payload).hexdigest()
    all_headers = {k.lower(): v.strip() for k, v in headers.items()}
    all_headers.update({"host": host, "x-amz-date": amz_date, "x-amz-content-sha256": payload_hash})
    signed_names = sorted(all_headers)
    canonical_headers = "".join(f"{k}:{all_headers[k]}\n" for k in signed_names)
    signed_headers = ";".join(signed_names)
    canonical_uri = quote(path, safe="/~")
    canonical_request = "\n".join([method, canonical_uri, "", canonical_headers, signed_headers, payload_hash])
    scope = f"{date_stamp}/{region}/{service}/aws4_request"
    string_to_sign = "\n".join(["AWS4-HMAC-SHA256", amz_date, scope, hashlib.sha256(canonical_request.encode()).hexdigest()])
    k_date = _sign(("AWS4" + secret_key).encode(), date_stamp)
    k_region = _sign(k_date, region)
    k_service = _sign(k_region, service)
    k_signing = _sign(k_service, "aws4_request")
    signature = hmac.new(k_signing, string_to_sign.encode(), hashlib.sha256).hexdigest()
    out = dict(headers)
    out.update({
        "Host": host, "X-Amz-Date": amz_date, "X-Amz-Content-Sha256": payload_hash,
        "Authorization": f"AWS4-HMAC-SHA256 Credential={access_key}/{scope}, SignedHeaders={signed_headers}, Signature={signature}",
    })
    return out


def retention_until(now: dt.datetime) -> dt.datetime:
    """now + 7 years; a 29 February start lands on 1 March so the date always exists."""
    try:
        return now.replace(year=now.year + RETENTION_YEARS)
    except ValueError:
        return now.replace(year=now.year + RETENTION_YEARS, month=3, day=1)


def _object_key(now: dt.datetime) -> str:
    return f"{now:%Y-%m-%d}/{now:%Y%m%dT%H%M%SZ}.json"


def _put_s3(key: str, body: bytes, now: dt.datetime) -> None:
    host = f"{settings.anchor_bucket}.s3.{settings.aws_region}.amazonaws.com"
    path = "/" + key
    retain_until = retention_until(now).strftime("%Y-%m-%dT%H:%M:%SZ")
    headers = {
        "Content-Type": "application/json",
        "x-amz-object-lock-mode": "COMPLIANCE",
        "x-amz-object-lock-retain-until-date": retain_until,
    }
    signed = sigv4_headers(method="PUT", host=host, path=path, payload=body, headers=headers,
                           region=settings.aws_region, service="s3",
                           access_key=settings.aws_access_key_id, secret_key=settings.aws_secret_access_key, now=now)
    r = httpx.put(f"https://{host}{path}", content=body, headers=signed, timeout=20.0)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"S3 answered {r.status_code}: {r.text[:200]}")


def _put_local(key: str, body: bytes) -> str:
    path = LOCAL_DIR / key
    path.parent.mkdir(parents=True, exist_ok=True)
    # Write-once in spirit: never overwrite an existing anchor file.
    with open(path, "xb") as f:
        f.write(body)
    return str(path)


def anchor_now() -> dict | None:
    """Anchor the current chain head. Returns the anchor row, or None when the log is empty."""
    now = dt.datetime.now(dt.timezone.utc)
    with connection() as conn:
        head = chain_head(conn)
        if head is None:
            return None
        head_seq, head_hash = head
        key = _object_key(now)
        row = conn.execute(
            "insert into anchor (head_seq, head_hash, object_key) values (%s, %s, %s) returning *",
            (head_seq, head_hash, key),
        ).fetchone()
    body = json.dumps({"head_seq": head_seq, "head_hash_hex": head_hash.hex(),
                       "at": now.isoformat(), "portal_version": APP_VERSION}, sort_keys=True).encode()
    try:
        if settings.anchor_bucket:
            _put_s3(key, body, now)
            where = f"s3://{settings.anchor_bucket}/{key}"
        else:
            where = _put_local(key, body)
        status, error = "written", None
    except Exception as exc:
        log.exception("anchor write failed")
        where, status, error = key, "failed", str(exc)[:500]
    with connection() as conn:
        return conn.execute(
            "update anchor set status = %s, error = %s, object_key = %s where id = %s returning *",
            (status, error, where, row["id"]),
        ).fetchone()


def read_local(object_key: str) -> dict:
    return json.loads(Path(object_key).read_bytes())
