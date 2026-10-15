"""The seal (SIG-0, SIG-4): a signature over the payload hash from a key the builder does not hold.

Production: an AWS Lambda function URL in Lenny's account, IAM-authenticated with SigV4 (hand-rolled
below with hmac/sha256 so the portal carries no boto3). The Lambda re-verifies the WebAuthn assertion
against the passkey registry it holds and signs with KMS. Any failure — network, 4xx/5xx, a malformed
answer — raises SealUnavailable and the approval fails closed; the router turns that into "Approvals
are temporarily unavailable. Nothing was lost."

Development and tests: when settings.is_test AND a dev PEM is configured, a local ECDSA P-256 key
seals with key_id 'dev-local'. Both conditions are required so a stray DEV_SEAL_PRIVATE_KEY_PEM in a
production .env can never turn the portal into its own notary.
"""
from __future__ import annotations

import base64
import datetime as _dt
import hashlib
import hmac
import json
import logging
import urllib.parse
from dataclasses import dataclass

import httpx
from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.asymmetric.utils import Prehashed

from accounting.config import settings
from accounting.lifecycle import canonical

log = logging.getLogger("accounting.seal")

UNAVAILABLE_TEXT = "Approvals are temporarily unavailable. Nothing was lost."
DEV_KEY_ID = "dev-local"


class SealUnavailable(Exception):
    """The signing service could not produce a seal. Always fail closed."""


@dataclass(frozen=True)
class Seal:
    signature: bytes
    key_id: str


# --- SigV4, by hand -------------------------------------------------------------------------------

def _sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _hmac(key: bytes, msg: str) -> bytes:
    return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).digest()


def _canonical_query(query: str) -> str:
    pairs = urllib.parse.parse_qsl(query, keep_blank_values=True)
    enc = [(urllib.parse.quote(k, safe="-_.~"), urllib.parse.quote(v, safe="-_.~")) for k, v in pairs]
    return "&".join(f"{k}={v}" for k, v in sorted(enc))


def sigv4_sign(*, method: str, url: str, headers: dict[str, str], body: bytes, access_key: str, secret_key: str,
               region: str, service: str, amz_date: str, session_token: str | None = None) -> dict[str, str]:
    """Return the headers to send (input headers + x-amz-date + Authorization [+ x-amz-security-token]).

    `amz_date` is passed in (YYYYMMDD'T'HHMMSS'Z') so the AWS published test vector can be replayed in
    tests; callers use `_now_amz()`. Header names are lower-cased and values trimmed per the spec.
    """
    parts = urllib.parse.urlsplit(url)
    canonical_uri = urllib.parse.quote(parts.path or "/", safe="/-_.~")
    canonical_qs = _canonical_query(parts.query)
    hdrs = {k.lower().strip(): " ".join(v.strip().split()) for k, v in headers.items()}
    hdrs.setdefault("host", parts.netloc)
    hdrs["x-amz-date"] = amz_date
    if session_token:
        hdrs["x-amz-security-token"] = session_token
    signed_names = sorted(hdrs)
    canonical_headers = "".join(f"{k}:{hdrs[k]}\n" for k in signed_names)
    signed_headers = ";".join(signed_names)
    payload_hash = _sha256_hex(body)
    canonical_request = "\n".join([method.upper(), canonical_uri, canonical_qs, canonical_headers, signed_headers, payload_hash])
    date_stamp = amz_date[:8]
    scope = f"{date_stamp}/{region}/{service}/aws4_request"
    string_to_sign = "\n".join(["AWS4-HMAC-SHA256", amz_date, scope, _sha256_hex(canonical_request.encode("utf-8"))])
    k_date = _hmac(("AWS4" + secret_key).encode("utf-8"), date_stamp)
    k_region = _hmac(k_date, region)
    k_service = _hmac(k_region, service)
    k_signing = _hmac(k_service, "aws4_request")
    signature = hmac.new(k_signing, string_to_sign.encode("utf-8"), hashlib.sha256).hexdigest()
    out = dict(hdrs)
    out["authorization"] = (f"AWS4-HMAC-SHA256 Credential={access_key}/{scope}, SignedHeaders={signed_headers}, Signature={signature}")
    return out


def _now_amz() -> str:
    return _dt.datetime.now(_dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")


# --- the dev signer -------------------------------------------------------------------------------

def _dev_enabled() -> bool:
    return bool(settings.is_test and settings.dev_seal_private_key_pem)


def _dev_private_key() -> ec.EllipticCurvePrivateKey:
    key = serialization.load_pem_private_key(settings.dev_seal_private_key_pem.encode(), password=None)
    if not isinstance(key, ec.EllipticCurvePrivateKey):
        raise SealUnavailable("dev seal key is not an EC key")
    return key


def public_key_pem() -> str:
    """The dev signer's public key, for SIG-7 verification in tests. Empty when the dev signer is off."""
    if not _dev_enabled():
        return ""
    return _dev_private_key().public_key().public_bytes(
        serialization.Encoding.PEM, serialization.PublicFormat.SubjectPublicKeyInfo).decode()


def _dev_seal(payload_hash: bytes) -> Seal:
    sig = _dev_private_key().sign(payload_hash, ec.ECDSA(Prehashed(hashes.SHA256())))
    return Seal(signature=sig, key_id=DEV_KEY_ID)


# --- the production client -------------------------------------------------------------------------

def _service_seal(payload_hash: bytes, assertion: dict | None, passkey_public_key: bytes | None, payload: dict | None) -> Seal:
    if not settings.signing_service_url:
        raise SealUnavailable("signing service URL is not configured")
    if not (settings.aws_access_key_id and settings.aws_secret_access_key):
        raise SealUnavailable("signing service credentials are not configured")
    body_obj = {
        "action": "seal",
        "payload_hash_hex": payload_hash.hex(),
        "assertion": assertion,
        "passkey_public_key_b64": base64.b64encode(passkey_public_key).decode() if passkey_public_key else None,
        "payload": payload,
    }
    body = canonical(body_obj)
    headers = sigv4_sign(method="POST", url=settings.signing_service_url, headers={"content-type": "application/json"}, body=body,
                         access_key=settings.aws_access_key_id, secret_key=settings.aws_secret_access_key, region=settings.aws_region,
                         service="lambda", amz_date=_now_amz())
    try:
        resp = httpx.post(settings.signing_service_url, content=body, headers=headers, timeout=10.0)
    except Exception as e:  # network, TLS, timeout — all the same to the approver
        log.error("signing service unreachable: %s", e)
        raise SealUnavailable("signing service unreachable") from e
    if resp.status_code != 200:
        log.error("signing service answered %s: %s", resp.status_code, resp.text[:300])
        raise SealUnavailable(f"signing service answered {resp.status_code}")
    try:
        data = resp.json()
        sig = base64.b64decode(data["signature_b64"])
        key_id = str(data["key_id"])
    except Exception as e:
        raise SealUnavailable("signing service answer was malformed") from e
    if not sig or not key_id:
        raise SealUnavailable("signing service answer was empty")
    return Seal(signature=sig, key_id=key_id)


def seal(payload_hash: bytes, assertion: dict | None, passkey_public_key: bytes | None, payload: dict | None = None) -> Seal:
    """Seal a payload hash. Dev signer only when test mode AND a dev key are both set; otherwise the
    service, failing closed on any error."""
    if len(payload_hash) != 32:
        raise SealUnavailable("payload hash must be 32 bytes")
    if _dev_enabled():
        return _dev_seal(payload_hash)
    return _service_seal(payload_hash, assertion, passkey_public_key, payload)


# --- verification (SIG-7) -------------------------------------------------------------------------

def _load_public_key(public_key_pem_or_jwk) -> ec.EllipticCurvePublicKey:
    if isinstance(public_key_pem_or_jwk, dict):
        jwk = public_key_pem_or_jwk
        x = int.from_bytes(_b64url_decode(jwk["x"]), "big")
        y = int.from_bytes(_b64url_decode(jwk["y"]), "big")
        return ec.EllipticCurvePublicNumbers(x, y, ec.SECP256R1()).public_key()
    data = public_key_pem_or_jwk.encode() if isinstance(public_key_pem_or_jwk, str) else bytes(public_key_pem_or_jwk)
    key = serialization.load_pem_public_key(data)
    if not isinstance(key, ec.EllipticCurvePublicKey):
        raise ValueError("seal public key is not an EC key")
    return key


def _b64url_decode(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def verify_seal(payload_hash: bytes, signature: bytes, public_key_pem_or_jwk) -> bool:
    """True when `signature` is a valid ECDSA-P256 signature over the 32-byte digest. Used by the
    nightly SIG-7 job (with the KMS public key) and by tests (with the dev key)."""
    try:
        key = _load_public_key(public_key_pem_or_jwk)
        key.verify(bytes(signature), bytes(payload_hash), ec.ECDSA(Prehashed(hashes.SHA256())))
        return True
    except (InvalidSignature, ValueError, TypeError):
        return False
