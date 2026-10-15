"""`process_invoice(invoice_id) -> final state` — RECEIVED to wherever the facts put it.

    RECEIVED ─ file-hash duplicate? ─────────────────────────────▶ DUPLICATE
             ─ screen (sender, file, hidden text) ───────────────▶ QUARANTINED
             ─ parse → render PNGs → extract → store fields
             ─ validate, match vendor, duplicate-number check, screen again ─▶ QUARANTINED | DUPLICATE
             ─ failures / unconfirmed split / low confidence ────▶ AP_REVIEW
             ─ CAPTURED ─ no vendor ─────────────────────────────▶ VENDOR_SETUP
                        ─ else routing.route(conn, invoice_id) ──▶ PENDING_APPROVAL | APPROVED | AP_REVIEW

The invoice stays RECEIVED until the very end, so a crash mid-way leaves a row the next pass simply
picks up again; the one non-idempotent cost is a repeated model call, which is why the extracted
fields are written as soon as they exist. Every state change is `transition()`; this module never
touches `invoice.state` itself. Model output is written to the header columns as DATA — the vendor on
the row comes from `vendors.match`, the quarantine decision from `screening.run` — so nothing the
model says can pick a payee or lift a flag (CAP-5).

Transient extractor failures (rate limit, overload, network) leave the row RECEIVED with an attempt
count in parse_detail; after MAX_EXTRACT_ATTEMPTS the invoice goes to AP_REVIEW so it cannot wait in
RECEIVED forever while an outage is nobody's fault.

`accounting.routing.route` belongs to Package B and is imported inside the function so this module
imports cleanly before that package exists (tests monkeypatch it).
"""
from __future__ import annotations

import hashlib
import io
import json
import logging
from uuid import UUID

from accounting import store
from accounting.capture import extractor, transient
from accounting.capture.validators import Failure, parse_date, to_cents, validate
from accounting.db.pool import connection
from accounting.lifecycle import _default, log_event, transition
from accounting.parse import client as parser
from accounting.screening import duplicates
from accounting.screening.run import screen
from accounting.vendors import fingerprint as fp
from accounting.vendors.match import match_vendor
from accounting.vendors.normalize import invoice_number_norm

log = logging.getLogger("accounting.pipeline")

MAX_EXTRACT_ATTEMPTS = 3


def _json(obj) -> str:
    return json.dumps(obj, default=_default)


def _merge_parse_detail(conn, invoice_id, extra: dict) -> None:
    conn.execute("update invoice set parse_detail = parse_detail || %s::jsonb, updated_at = now() where id = %s",
                 (_json(extra), invoice_id))


def _slice(pages: list, page_from: int | None, page_to: int | None) -> list:
    if page_from is None:
        return pages
    return pages[max(page_from - 1, 0): page_to or len(pages)]


def _store_renders(conn, inv, pages_png: list[bytes]) -> None:
    conn.execute("delete from invoice_render where invoice_id = %s", (inv["id"],))
    for i, png in enumerate(pages_png, start=1):
        ref = store.put("renders", f"{inv['id']}/{i}-{hashlib.sha256(png).hexdigest()[:16]}.png", png, "image/png")
        w = h = None
        try:
            from PIL import Image

            with Image.open(io.BytesIO(png)) as im:
                w, h = im.size
        except Exception:
            pass
        conn.execute("insert into invoice_render (invoice_id, page, ref, width, height) values (%s,%s,%s,%s,%s)",
                     (inv["id"], i, ref, w, h))


def _write_fields(conn, inv, ex: dict) -> dict:
    """Extracted header fields → invoice columns. Returns the refreshed row."""
    v = lambda f: (ex.get(f) or {}).get("value")  # noqa: E731
    remit = v("remit_to")
    bank_last4 = v("bank_last4")
    conf = {f: (ex.get(f) or {}).get("confidence", 0.0) for f in extractor.FIELDS}
    row = conn.execute(
        """update invoice set vendor_name_raw = %s, invoice_number = %s, invoice_number_norm = %s, invoice_date = %s,
                              due_date = %s, currency = %s, gross_cents = %s, tax_cents = %s, line_total_cents = %s,
                              po_number = %s, remit_to = %s::jsonb, bank_last4 = %s, field_confidence = %s::jsonb,
                              updated_at = now()
           where id = %s returning *""",
        (v("vendor_name"), v("invoice_number"), invoice_number_norm(v("invoice_number")) or None,
         parse_date(v("invoice_date")), parse_date(v("due_date")), v("currency"),
         to_cents(v("gross")), to_cents(v("tax")), to_cents(v("line_total")), v("po_number"),
         _json(remit) if remit else None, bank_last4, _json(conf), inv["id"]),
    ).fetchone()
    conn.execute("delete from invoice_line where invoice_id = %s", (inv["id"],))
    for i, ln in enumerate(v("lines") or [], start=1):
        conn.execute(
            "insert into invoice_line (invoice_id, line_no, description, quantity, amount_cents) values (%s,%s,%s,%s,%s)",
            (inv["id"], i, (ln.get("description") or "")[:500] or None, _num(ln.get("quantity")), to_cents(ln.get("amount"))))
    return row


def _num(q):
    from decimal import Decimal, InvalidOperation

    if q in (None, ""):
        return None
    try:
        return Decimal(str(q).replace(",", ""))
    except InvalidOperation:
        return None


def _ap_review(conn, inv, reason: str, failures: list[Failure] | None = None, note: str | None = None) -> str:
    payload = {"failures": [f.__dict__ for f in failures or []]}
    if note:
        payload["note"] = note
    transition(conn, inv["id"], "AP_REVIEW", reason_code=reason, actor_type="system", payload=payload)
    return "AP_REVIEW"


def _duplicate(conn, inv, dup: duplicates.DupResult) -> str:
    transition(conn, inv["id"], "DUPLICATE", reason_code=dup.duplicate_reason, actor_type="system",
               payload={"matched_invoice_id": str(dup.matched_invoice_id), **dup.detail})
    conn.execute("update invoice set linked_invoice_id = %s where id = %s", (dup.matched_invoice_id, inv["id"]))
    return "DUPLICATE"


def process_invoice(invoice_id: UUID | str) -> str:
    with connection() as conn:
        inv = conn.execute("select * from invoice where id = %s for update", (invoice_id,)).fetchone()
        if inv is None:
            raise ValueError(f"invoice {invoice_id} not found")
        if inv["state"] != "RECEIVED":
            return inv["state"]

        # 1. the same file again
        dup = duplicates.check(conn, inv)
        if dup.duplicate_reason == "file_hash":
            return _duplicate(conn, inv, dup)

        # 2. sender- and file-level triggers before spending a model call
        s = screen(conn, invoice_id)
        if s.hard:
            return "QUARANTINED"

        if inv["attachment_id"] is None:
            return _ap_review(conn, inv, "validator_failed",
                              [Failure("no_attachment", None, "the message has no invoice file attached")])
        att = conn.execute("select * from attachment where id = %s", (inv["attachment_id"],)).fetchone()

        # 3. parse in the sandbox
        try:
            res = parser.parse(store.get(att["ref"]), att["content_type"] or "application/octet-stream", att["filename"] or "")
        except parser.ParseError as e:
            _merge_parse_detail(conn, invoice_id, {"parse_error": str(e)})
            return _ap_review(conn, inv, "validator_failed", [Failure("unparseable", None, str(e))])
        if res.risky:
            conn.execute("update attachment set risky = true, risk_reason = %s where id = %s", (res.risk_reason, att["id"]))
            screen(conn, invoice_id)
            return "QUARANTINED"
        pages = _slice(res.pages, inv["page_from"], inv["page_to"])
        texts = _slice(res.page_texts, inv["page_from"], inv["page_to"])
        text = "\n\f\n".join(texts)
        flags = [f for f in res.hidden_text_flags
                 if inv["page_from"] is None or inv["page_from"] <= int(f.get("page") or 0) <= (inv["page_to"] or res.page_count)]
        _merge_parse_detail(conn, invoice_id, {"page_count": len(pages), "hidden_text_flags": flags, "metadata": res.metadata,
                                               "text_sha256": hashlib.sha256(text.encode()).hexdigest()})
        _store_renders(conn, inv, pages)
        if flags:
            screen(conn, invoice_id)  # hidden_text is a hard trigger; screen() transitions
            return "QUARANTINED"

        # 4. extract (the only non-deterministic step; it proposes, nothing more)
        attempts = int((inv["parse_detail"] or {}).get("extract_attempts") or 0) + 1
        _merge_parse_detail(conn, invoice_id, {"extract_attempts": attempts})
        try:
            ex = extractor.extract(pages, text)
        except extractor.ExtractionError as e:
            _merge_parse_detail(conn, invoice_id, {"extract_error": str(e)})
            return _ap_review(conn, inv, "low_confidence", note=str(e))
        except Exception as e:
            if transient.is_transient(e) and attempts < MAX_EXTRACT_ATTEMPTS:
                log.warning("transient extractor failure on %s (attempt %d): %s", invoice_id, attempts, e)
                log_event(conn, event_type="extract_retry", actor_type="system", invoice_id=invoice_id,
                          payload={"attempt": attempts, "error": type(e).__name__})
                return "RECEIVED"
            log.exception("extractor failed on %s", invoice_id)
            _merge_parse_detail(conn, invoice_id, {"extract_error": type(e).__name__})
            return _ap_review(conn, inv, "low_confidence", note="The invoice could not be read automatically.")
        inv = _write_fields(conn, inv, ex)

        # 5. vendor, fingerprint, duplicates, full screen
        msg = conn.execute("select * from message where id = %s", (inv["message_id"],)).fetchone() if inv["message_id"] else None
        sender = (msg or {}).get("original_sender") or (msg or {}).get("from_addr")
        match = match_vendor(conn, sender_addr=sender, vendor_name=inv["vendor_name_raw"], tin=None,
                             remit_to=inv["remit_to"], message_text=((msg or {}).get("body_text") or "") + "\n" + text)
        conn.execute("update invoice set vendor_id = %s, vendor_match = %s::jsonb where id = %s",
                     (match.vendor_id, _json(match.as_dict()), inv["id"]))
        vendor = conn.execute("select * from vendor where id = %s", (match.vendor_id,)).fetchone() if match.vendor_id else None
        tags = [t for t, on in (("hcp", vendor and vendor["hcp"]), ("quality_critical", vendor and vendor["quality_critical"]),
                                ("fcpa_agent", vendor and vendor["fcpa_agent"]), ("intercompany", vendor and vendor["intercompany"]),
                                ("recurring", vendor and vendor["recurring_schedule"])) if on]
        # Only the last four bank digits are ever printed, so no fingerprint is computed from the invoice;
        # the screen compares last4 against the vendor's record. `invoice.bank_fingerprint` is set when AP
        # verifies a full account number, never from a model read.
        conn.execute("update invoice set tags = %s where id = %s", (tags, inv["id"]))
        inv = conn.execute("select * from invoice where id = %s", (inv["id"],)).fetchone()

        dup = duplicates.check(conn, inv)
        if dup.duplicate_reason:
            return _duplicate(conn, inv, dup)
        s = screen(conn, invoice_id)
        if s.hard:
            return "QUARANTINED"

        # 6. validators decide
        failures = validate({**inv, "vendor": vendor}, ex)
        if not inv["split_confirmed"]:
            return _ap_review(conn, inv, "split_unconfirmed", failures, note="the file appears to hold more than one invoice")
        if failures:
            codes = {f.code for f in failures}
            if codes == {"missing_po"}:
                return _ap_review(conn, inv, "missing_po", failures)
            if codes <= {"low_confidence"}:
                return _ap_review(conn, inv, "low_confidence", failures)
            return _ap_review(conn, inv, "validator_failed", failures)

        transition(conn, invoice_id, "CAPTURED", reason_code="capture_ok", actor_type="system",
                   payload={"fields": {f: (ex.get(f) or {}).get("value") for f in extractor.FIELDS if f != "lines"},
                            "vendor_match": match.as_dict(), "soft_signals": s.soft})
        if vendor is None:
            transition(conn, invoice_id, "VENDOR_SETUP", reason_code="vendor_unknown", actor_type="system",
                       payload={"vendor_name_raw": inv["vendor_name_raw"], "match": match.as_dict()})
            return "VENDOR_SETUP"
        if vendor["status"] == "setup":
            transition(conn, invoice_id, "VENDOR_SETUP", reason_code="vendor_unknown", actor_type="system",
                       payload={"vendor_id": str(vendor["id"]), "note": "vendor record is still being set up"})
            return "VENDOR_SETUP"

        from accounting.routing.route import route  # Package B

        return route(conn, invoice_id)


def process_pending(limit: int = 20) -> int:
    with connection() as conn:
        ids = [r["id"] for r in conn.execute(
            "select id from invoice where state = 'RECEIVED' order by received_at limit %s", (limit,)).fetchall()]
    done = 0
    for invoice_id in ids:
        try:
            if process_invoice(invoice_id) != "RECEIVED":
                done += 1
        except Exception:
            log.exception("pipeline failed on %s", invoice_id)
            try:
                with connection() as conn:
                    log_event(conn, event_type="pipeline_error", actor_type="system", invoice_id=invoice_id, payload={})
            except Exception:
                pass
    return done
