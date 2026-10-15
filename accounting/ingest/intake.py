"""`ingest_raw_message(raw, *, imap_uid, source) -> [invoice ids]` — one message in, invoice rows out.

The order of decisions, and why it is this order:
  1. Store the raw .eml first (INT-1: nothing is ever discarded, including what we then drop).
  2. Idempotency on (message_id, raw_sha256): the IMAP poller re-reading a UID after a crash must not
     make a second set of invoices.
  3. Trust (INT-4). Our own mail, bounces and auto-replies are recorded as such and produce no invoice.
  4. Internal forwards: a Maxx employee forwarding a vendor's mail is judged by the ORIGINAL sender,
     with the forwarder recorded. If the original sender cannot be read, the sender is "unknown" and
     later matching treats it as unlisted.
  5. Classify (INT-2) and store attachments with the risky flag (INT-5), then split (INT-3).
  6. Every invoice row is created in RECEIVED and then moved with `transition()` — never inserted in
     another state — so every row's history starts the same way:
       direct sender      → QUARANTINED (quarantine_case trigger direct_sender [+ risky_file])
       risky file         → QUARANTINED (risky_file)
       statement/credit   → DIVERTED
       bank_change        → DIVERTED and a quarantine_case (bank_change_request): both, because the
                            Controller must see it and AP must not process it as an invoice
       inquiry/spam_other → DIVERTED (non_invoice); the message is kept, nothing is lost
       no attachment      → AP_REVIEW (validator_failed, "no file") so AP can chase the document
       otherwise          → stays RECEIVED for the pipeline
Uploads (INT-6) arrive here too with source="upload": the clerk is signed in, so the trusted-path
check does not apply and the message is recorded as forwarded_by_maxx with `upload: true` in headers.
"""
from __future__ import annotations

import email
import email.utils
import hashlib
import json
import logging
import re
from datetime import datetime, timezone
from email import policy
from email.message import EmailMessage
from uuid import UUID

from accounting import store
from accounting.config import settings
from accounting.db.pool import connection
from accounting.ingest import splitter, trusted_path
from accounting.ingest.classifier import classify
from accounting.lifecycle import log_event, transition
from accounting.parse.engine import assess_risk
from accounting.vendors.normalize import email_domain

log = logging.getLogger("accounting.ingest.intake")

FORWARD_MARKERS = re.compile(
    r"(?:^|\n)\s*(?:-{2,}\s*(?:Original Message|Forwarded message|Weitergeleitete Nachricht)\s*-{2,}|Begin forwarded message:|"
    r"From:\s*[^\n]+\n(?:Sent|Date):\s*[^\n]+\n(?:To:[^\n]+\n)?(?:Cc:[^\n]+\n)?Subject:)", re.I)
FORWARDED_FROM = re.compile(r"(?:^|\n)\s*\*?From:\*?\s*([^\n]+)", re.I)
MAX_BODY = 200_000


def _company_domains() -> set[str]:
    return {settings.forwarder_authserv_id.lower()}


def _text_body(msg: EmailMessage) -> str:
    """Plain text first; an HTML-only body is tag-stripped. Truncated — screening needs a few KB."""
    try:
        part = msg.get_body(preferencelist=("plain", "html"))
    except Exception:
        part = None
    if part is None:
        return ""
    try:
        text = part.get_content()
    except Exception:
        return ""
    if part.get_content_type() == "text/html":
        text = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", text, flags=re.S | re.I)
        text = re.sub(r"<br\s*/?>|</p>|</div>|</tr>", "\n", text, flags=re.I)
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"&nbsp;", " ", text)
    return text[:MAX_BODY]


def _received_at(msg: EmailMessage) -> datetime:
    """The Date header, else the first Received date — never our clock (the row says when the vendor
    sent it, which is what aging is measured from)."""
    for candidate in [msg.get("Date")] + [r.rsplit(";", 1)[-1] for r in (msg.get_all("Received") or [])]:
        if not candidate:
            continue
        try:
            dt = email.utils.parsedate_to_datetime(candidate.strip())
            if dt is not None:
                return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
        except (TypeError, ValueError):
            continue
    return datetime.now(timezone.utc)


def _headers_json(msg: EmailMessage) -> dict:
    out: dict = {}
    for k, v in msg.items():
        out.setdefault(k, []).append(str(v)[:2000])
    return out


def _addr(header_value: str | None) -> tuple[str, str]:
    name, addr = email.utils.parseaddr(header_value or "")
    return name.strip(), addr.strip().lower()


def detect_internal_forward(from_addr: str, body: str) -> str | None:
    """If a company address forwarded someone else's mail, the original sender's address, else None."""
    if email_domain(from_addr) not in _company_domains():
        return None
    m = FORWARD_MARKERS.search(body or "")
    if not m:
        return None
    tail = body[m.start():]
    fm = FORWARDED_FROM.search(tail)
    if not fm:
        return ""
    _, addr = _addr(fm.group(1).replace("[mailto:", "<").replace("]", ">"))
    if not addr or "@" not in addr:
        em = re.search(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", fm.group(1))
        addr = em.group(0).lower() if em else ""
    return addr


def _is_inline_image(part, html_body: str) -> bool:
    cid = (part.get("Content-ID") or "").strip("<>")
    disp = (part.get_content_disposition() or "").lower()
    if part.get_content_maintype() != "image":
        return False
    if cid and f"cid:{cid}" in html_body:
        return True
    return disp == "inline" and not part.get_filename()


def _html_body(msg: EmailMessage) -> str:
    try:
        part = msg.get_body(preferencelist=("html",))
        return part.get_content() if part is not None else ""
    except Exception:
        return ""


def _store_attachments(conn, message_row_id, msg: EmailMessage) -> list[dict]:
    rows: list[dict] = []
    html = _html_body(msg)
    for part in msg.iter_attachments():
        try:
            data = part.get_content()
        except Exception:
            continue
        if isinstance(data, str):
            data = data.encode("utf-8", errors="replace")
        if not isinstance(data, (bytes, bytearray)):
            # a nested message/rfc822 — keep its bytes
            data = bytes(part) if not hasattr(data, "as_bytes") else data.as_bytes()
        data = bytes(data)
        if not data:
            continue
        filename = (part.get_filename() or "").strip()[:255] or None
        ctype = (part.get_content_type() or "application/octet-stream").lower()
        inline = _is_inline_image(part, html)
        sha = hashlib.sha256(data).digest()
        risky, reason = assess_risk(filename, ctype, data)
        ref = store.put("originals", f"attachments/{sha.hex()}", data, ctype)
        row = conn.execute(
            """insert into attachment (message_id, filename, content_type, size_bytes, sha256, ref, risky, risk_reason)
               values (%s,%s,%s,%s,%s,%s,%s,%s) returning *""",
            (message_row_id, filename, ctype, len(data), sha, ref, risky, reason),
        ).fetchone()
        row = dict(row)
        row["inline"] = inline
        row["_data"] = data
        rows.append(row)
    return rows


def _new_invoice(conn, msg_row, unit: splitter.Unit, source: str) -> dict:
    att = unit.attachment
    return conn.execute(
        """insert into invoice (state, message_id, attachment_id, source, split_index, split_confirmed, page_from, page_to,
                                file_sha256, received_at, parse_detail)
           values ('RECEIVED', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s::jsonb) returning *""",
        (msg_row["id"], att["id"] if att else None, source, unit.split_index, unit.split_confirmed, unit.page_from,
         unit.page_to, att["sha256"] if att else None, msg_row["received_at"], json.dumps({"split": unit.detail})),
    ).fetchone()


def _quarantine(conn, inv, msg_row, triggers: list[str], detail: dict) -> None:
    conn.execute(
        "insert into quarantine_case (invoice_id, message_id, triggers, detail) values (%s, %s, %s, %s::jsonb)",
        (inv["id"], msg_row["id"], triggers, json.dumps(detail)),
    )
    transition(conn, inv["id"], "QUARANTINED", reason_code="hard_trigger", actor_type="system",
               payload={"triggers": triggers, "detail": detail})


def ingest_raw_message(raw: bytes, *, imap_uid: int | None, source: str = "email") -> list[UUID]:
    raw_sha = hashlib.sha256(raw).digest()
    msg: EmailMessage = email.message_from_bytes(raw, policy=policy.default)  # type: ignore[assignment]
    message_id = (msg.get("Message-ID") or "").strip() or None

    with connection() as conn:
        existing = conn.execute("select id from message where raw_sha256 = %s and message_id is not distinct from %s",
                                (raw_sha, message_id)).fetchone()
        if existing:
            return [r["id"] for r in conn.execute("select id from invoice where message_id = %s order by created_at",
                                                  (existing["id"],)).fetchall()]

        raw_ref = store.put("originals", f"messages/{raw_sha.hex()}.eml", raw, "message/rfc822")
        from_name, from_addr = _addr(msg.get("From"))
        body = _text_body(msg)
        subject = (msg.get("Subject") or "")[:998]

        def _is_ours(mid: str) -> bool:
            return conn.execute("select 1 from notification where message_id_header = %s limit 1", (mid,)).fetchone() is not None

        if source == "upload":
            trust = trusted_path.TrustResult("forwarded_by_maxx", {"upload": True, "reason": "manual upload by a signed-in user"})
        else:
            trust = trusted_path.decide(msg, is_portal_message_id=_is_ours)

        original_sender, forwarded_by = None, None
        if trust.trust in ("forwarded_by_maxx", "direct"):
            fwd = detect_internal_forward(from_addr, body)
            if fwd is not None:
                forwarded_by = from_addr
                original_sender = fwd or None   # "" = forwarder detected, original unreadable → unlisted

        names_for_class = [p.get_filename() for p in msg.iter_attachments() if p.get_filename()]
        cls = classify(subject, body, names_for_class)

        headers = _headers_json(msg)
        headers["_trust"] = trust.detail
        msg_row = conn.execute(
            """insert into message (imap_uid, message_id, raw_ref, raw_sha256, from_addr, from_name, original_sender, forwarded_by,
                                    subject, received_at, headers, auth_results, trust, classification, classifier_detail, body_text)
               values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s::jsonb,%s::jsonb,%s,%s,%s::jsonb,%s) returning *""",
            (imap_uid, message_id, raw_ref, raw_sha, from_addr or None, from_name or None, original_sender, forwarded_by,
             subject, _received_at(msg), json.dumps(headers), json.dumps(trust.auth_results or None), trust.trust,
             cls.label, json.dumps(cls.as_dict()), body),
        ).fetchone()
        log_event(conn, event_type="ingest", actor_type="system", reason_code=trust.trust,
                  payload={"message_id": str(msg_row["id"]), "classification": cls.label, "source": source,
                           "from": from_addr, "forwarded_by": forwarded_by, "original_sender": original_sender})

        if trust.trust in ("portal_own", "bounce"):
            return []

        attachments = _store_attachments(conn, msg_row["id"], msg)
        units = splitter.split(attachments, lambda a: a["_data"])

        ids: list[UUID] = []
        for unit in units:
            inv = _new_invoice(conn, msg_row, unit, source)
            ids.append(inv["id"])
            att = unit.attachment
            triggers: list[str] = []
            detail: dict = {}
            if trust.trust == "direct":
                triggers.append("direct_sender")
                detail["direct_sender"] = trust.detail
            if att and att["risky"]:
                triggers.append("risky_file")
                detail["risky_file"] = {"filename": att["filename"], "reason": att["risk_reason"]}
            if triggers:
                _quarantine(conn, inv, msg_row, triggers, detail)
                continue
            if cls.label in ("statement", "credit_memo"):
                transition(conn, inv["id"], "DIVERTED", reason_code=cls.label, actor_type="system", payload=cls.as_dict())
                conn.execute("update invoice set diverted_reason = %s where id = %s", (cls.label, inv["id"]))
            elif cls.label == "bank_change":
                conn.execute(
                    "insert into quarantine_case (invoice_id, message_id, triggers, detail) values (%s, %s, %s, %s::jsonb)",
                    (inv["id"], msg_row["id"], ["bank_change_request"], json.dumps(cls.as_dict())))
                transition(conn, inv["id"], "DIVERTED", reason_code="bank_change_request", actor_type="system", payload=cls.as_dict())
                conn.execute("update invoice set diverted_reason = 'bank_change_request' where id = %s", (inv["id"],))
            elif cls.label in ("inquiry", "spam_other"):
                transition(conn, inv["id"], "DIVERTED", reason_code="non_invoice", actor_type="system", payload=cls.as_dict())
                conn.execute("update invoice set diverted_reason = %s where id = %s", (cls.label, inv["id"]))
            elif att is None:
                transition(conn, inv["id"], "AP_REVIEW", reason_code="validator_failed", actor_type="system",
                           payload={"failures": [{"code": "no_attachment", "field": None, "text": "the message has no invoice file attached"}]})
        return ids


def build_upload_eml(*, filename: str, content_type: str, data: bytes, uploaded_by: str, note: str = "") -> bytes:
    """What `web/uploads.py` hands to ingest_raw_message for INT-6: a synthetic message that records who
    uploaded the file, so an upload has the same row shape as a mailed invoice."""
    m = EmailMessage()
    m["From"] = uploaded_by
    m["To"] = settings.imap_user
    m["Subject"] = f"Manual upload: {filename}"
    m["Date"] = email.utils.format_datetime(datetime.now(timezone.utc))
    m["Message-ID"] = email.utils.make_msgid(domain=settings.portal_host)
    m["X-Portal-Upload"] = "1"
    m.set_content(note or f"Uploaded through the portal by {uploaded_by}.")
    maintype, _, subtype = (content_type or "application/octet-stream").partition("/")
    m.add_attachment(data, maintype=maintype, subtype=subtype or "octet-stream", filename=filename)
    return m.as_bytes()
