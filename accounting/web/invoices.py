"""/invoices — the AP working screens: list, detail, review edit, duplicate/divert, split confirm,
answer a question, record the SAP number, Controller void with a second person.

Every state change goes through `lifecycle.transition()`. Edits to header fields are the one
action with side effects outside this package (SIG-3): any existing signature is voided through
`approvals.void.void_all_for_invoice`, and the invoice is re-routed through `routing.route.route`.
Both are imported lazily so the three packages can be built side by side.

No error text from an exception ever reaches a page: every failure the reader can cause has a
sentence written for it, and everything else is "Something went wrong" plus a log line.
"""
from __future__ import annotations

import json
import logging
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation

from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import RedirectResponse, Response

from accounting.config import settings
from accounting.db.pool import connection
from accounting.lifecycle import TERMINAL, InvalidTransition, log_event, transition
from accounting.notify.send import queue as queue_notification
from accounting.reports import queue as approved_queue
from accounting.web.common import as_uuid, page, portal_number
from accounting.web.deps import NotFound, require, require_active

log = logging.getLogger("accounting.web.invoices")
router = APIRouter(prefix="/invoices")

VIEWERS = require("clerk", "supervisor", "controller", "admin", "auditor")
CLERKS = require("clerk", "supervisor", "controller", "admin")
CONTROLLERS = require("controller", "admin")

EDITABLE = ("vendor_name_raw", "invoice_number", "invoice_date", "due_date", "currency", "gross_cents", "tax_cents", "po_number")
STATES = ["RECEIVED", "CAPTURED", "AP_REVIEW", "QUARANTINED", "VENDOR_SETUP", "PENDING_APPROVAL", "QUESTIONED", "APPROVED",
          "ENTERED", "DIVERTED", "DUPLICATE", "FRAUD", "REJECTED", "VOID"]
DIVERT_REASONS = {"statement": "It is a statement, not an invoice", "credit_memo": "It is a credit memo",
                  "intercompany": "It is intercompany", "other_company_code": "It belongs to another company code",
                  "non_invoice": "It is not an invoice at all"}

# (event_type, reason_code) -> plain words for the timeline. Fallbacks below cover anything new.
EVENT_TEXT = {
    ("transition", "capture_ok"): "Fields were read from the document and passed every check.",
    ("transition", "validator_failed"): "A check failed, so AP needs to look at the fields.",
    ("transition", "low_confidence"): "Some fields were read with low confidence, so AP needs to confirm them.",
    ("transition", "split_unconfirmed"): "The file may hold more than one invoice; AP needs to confirm the split.",
    ("transition", "missing_po"): "No purchase order was found, so AP needs to review it.",
    ("transition", "manual_upload"): "Uploaded by hand, so AP reviews it first.",
    ("transition", "hard_trigger"): "A security trigger fired; held in quarantine for the Controller.",
    ("transition", "controller_release"): "Released from quarantine by the Controller.",
    ("transition", "controller_reject"): "Rejected from quarantine by the Controller.",
    ("transition", "controller_report"): "Reported as suspicious by the Controller; sender blocked.",
    ("transition", "clerk_edit"): "AP corrected the fields.",
    ("transition", "clerk_confirmed"): "AP confirmed the fields as read.",
    ("transition", "vendor_unknown"): "The vendor is not in the portal yet.",
    ("transition", "vendor_added"): "The vendor was added.",
    ("transition", "auto_release_recurring"): "Approved automatically as a recurring invoice within tolerance.",
    ("transition", "routed"): "Sent to the approvers.",
    ("transition", "asl_hard_stop"): "Stopped: the vendor's approved-supplier status does not allow it.",
    ("transition", "quality_no_po"): "Stopped: a quality-critical purchase without a purchase order.",
    ("transition", "field_edit"): "Fields were edited after routing; approvals were voided and it goes back to AP.",
    ("transition", "approver_question"): "An approver asked a question; the clock is paused.",
    ("transition", "ap_answered"): "AP answered the question; back with the approvers.",
    ("transition", "chain_complete"): "Every required approval is signed.",
    ("transition", "signed_rejection"): "Rejected by an approver.",
    ("transition", "reassign_exhausted"): "Nobody could be found to approve it; back to AP.",
    ("transition", "sap_entered"): "Keyed into SAP; the SAP document number was recorded.",
    ("transition", "controller_void"): "Voided by the Controller with a second person.",
    ("transition", "statement"): "Set aside: it is a statement, not an invoice.",
    ("transition", "credit_memo"): "Set aside: it is a credit memo.",
    ("transition", "intercompany"): "Set aside: intercompany.",
    ("transition", "other_company_code"): "Set aside: another company code.",
    ("transition", "non_invoice"): "Set aside: not an invoice.",
    ("transition", "bank_change_request"): "Set aside: a request to change bank details, which is never handled by email.",
    ("transition", "file_hash"): "Marked as a duplicate: the same file was seen before.",
    ("transition", "invoice_number"): "Marked as a duplicate: the same invoice number was seen before.",
    ("signature", None): "Signed.",
    ("question", None): "A question was asked.",
    ("answer", None): "A question was answered.",
    ("correction", None): "Fields were corrected.",
    ("split_confirmed", None): "AP confirmed how the file was split into invoices.",
}


def event_text(e: dict) -> str:
    key = (e["event_type"], e["reason_code"])
    if key in EVENT_TEXT:
        return EVENT_TEXT[key]
    if (e["event_type"], None) in EVENT_TEXT:
        return EVENT_TEXT[(e["event_type"], None)]
    if e["event_type"] == "transition":
        return f"Moved from {_words(e['from_state'])} to {_words(e['to_state'])}."
    return (e["reason_code"] or e["event_type"] or "").replace("_", " ").capitalize() + "."


def _words(state: str | None) -> str:
    return (state or "").replace("_", " ").title()


def _load(conn, invoice_id) -> dict:
    inv = conn.execute(
        """select i.*, v.legal_name as vendor, v.hcp as vendor_hcp, v.sap_vendor_number
           from invoice i left join vendor v on v.id = i.vendor_id where i.id = %s""", (invoice_id,)).fetchone()
    if inv is None:
        raise NotFound()
    return inv


def _detail_ctx(conn, inv: dict, user, error: str | None = None, note: str | None = None) -> dict:
    iid = inv["id"]
    pages = conn.execute("select page from invoice_render where invoice_id = %s order by page", (iid,)).fetchall()
    events = conn.execute(
        """select e.*, u.display_name as actor from event_log e left join ap_user u on u.id = e.actor_id
           where e.invoice_id = %s order by e.seq""", (iid,)).fetchall()
    for e in events:
        e["text"] = event_text(e)
    questions = conn.execute(
        """select q.*, a.display_name as asker, b.display_name as answerer from invoice_question q
           join ap_user a on a.id = q.asked_by left join ap_user b on b.id = q.answered_by
           where q.invoice_id = %s order by q.asked_at""", (iid,)).fetchall()
    signatures = conn.execute(
        """select s.signer_role, s.decision, s.signed_at, s.status, u.display_name as signer from approval_signature s
           join ap_user u on u.id = s.signer_id where s.invoice_id = %s order by s.signed_at""", (iid,)).fetchall()
    requests = conn.execute(
        """select r.role, r.tier, r.status, r.due_at, r.created_at, u.display_name as assignee from approval_request r
           join ap_user u on u.id = r.assignee_id where r.invoice_id = %s order by r.tier, r.created_at""", (iid,)).fetchall()
    second_people = conn.execute(
        """select id, display_name from ap_user where status = 'active' and id <> %s
           and roles && '{controller,admin}'::portal_role[] order by display_name""", (user.id,)).fetchall()
    lines = conn.execute("select * from invoice_line where invoice_id = %s order by line_no", (iid,)).fetchall()
    conf = inv["field_confidence"] or {}
    fields = []
    for name, label in (("vendor_name_raw", "Vendor (as printed)"), ("invoice_number", "Invoice number"), ("invoice_date", "Invoice date"),
                        ("due_date", "Due date"), ("currency", "Currency"), ("gross_cents", "Gross"), ("tax_cents", "Tax"),
                        ("line_total_cents", "Line total"), ("po_number", "PO number"), ("bank_last4", "Bank account, last 4")):
        c = conf.get(name.replace("_cents", ""), conf.get(name))
        fields.append({"name": name, "label": label, "value": inv[name], "confidence": c,
                       "low": c is not None and float(c) < settings.capture_confidence_threshold,
                       "money": name.endswith("_cents")})
    return {
        "inv": inv, "pn": portal_number(inv["portal_number"]), "pages": [p["page"] for p in pages], "events": events,
        "questions": questions, "signatures": signatures, "requests": requests, "second_people": second_people, "lines": lines,
        "fields": fields, "error": error, "note": note, "divert_reasons": DIVERT_REASONS,
        "can_act": user.status == "active" and user.has("clerk", "supervisor", "controller", "admin"),
        "is_controller": user.has("controller", "admin"), "terminal": inv["state"] in TERMINAL,
        "open_questions": [q for q in questions if q["answered_at"] is None],
        "cents_text": approved_queue.cents_text,
    }


# --- list -----------------------------------------------------------------------------------------

@router.get("")
def list_invoices(request: Request, state: str | None = None, vendor: str | None = None, q: str | None = None,
                  user=Depends(VIEWERS)):
    where, params = [], []
    if state and state in STATES:
        where.append("i.state = %s"); params.append(state)
    elif not state:
        where.append("i.state::text <> all(%s)"); params.append(list(TERMINAL))
    if vendor:
        where.append("(v.legal_name ilike %s or i.vendor_name_raw ilike %s)"); params += [f"%{vendor}%", f"%{vendor}%"]
    if q:
        like = f"%{q.strip()}%"
        where.append("(i.invoice_number ilike %s or i.po_number ilike %s or i.vendor_name_raw ilike %s or 'AP-' || lpad(i.portal_number::text, 6, '0') ilike %s)")
        params += [like, like, like, like]
    sql = f"""select i.id, i.portal_number, i.state, i.received_at, i.invoice_number, i.invoice_date, i.currency, i.gross_cents, i.tags,
                     coalesce(v.legal_name, i.vendor_name_raw) as vendor
              from invoice i left join vendor v on v.id = i.vendor_id
              {'where ' + ' and '.join(where) if where else ''} order by i.received_at limit 500"""
    with connection() as conn:
        rows = conn.execute(sql, params).fetchall()
    return page(request, user, "invoices/list.html", {"rows": rows, "states": STATES, "state": state or "", "vendor": vendor or "", "q": q or ""})


# --- approved queue (SAP-1). Declared before /{invoice_id} so "approved" is never read as an id. ----

@router.get("/approved")
def approved(request: Request, error: str | None = None, user=Depends(VIEWERS)):
    with connection() as conn:
        rows = approved_queue.approved_rows(conn)
    return page(request, user, "invoices/approved.html", {"rows": rows, "error": None, "cents_text": approved_queue.cents_text})


@router.get("/approved.csv")
def approved_csv(user=Depends(VIEWERS)):
    now = datetime.now(timezone.utc)
    with connection() as conn:
        rows = approved_queue.approved_rows(conn)
    return Response(approved_queue.csv_text(rows, now), media_type="text/csv; charset=utf-8",
                    headers={"Content-Disposition": f'attachment; filename="{approved_queue.csv_filename(now)}"'})


@router.post("/{invoice_id}/sap")
def record_sap(request: Request, invoice_id: str, sap_document_number: str = Form(""), back: str = Form("/invoices/approved"),
               user=Depends(CLERKS), _active=Depends(require_active)):
    iid = as_uuid(invoice_id)
    try:
        with connection() as conn:
            inv = _load(conn, iid)
            if inv["state"] != "APPROVED":
                raise approved_queue.SapNumberError("Only an approved invoice can be recorded as entered.")
            approved_queue.record_sap_document(conn, iid, sap_document_number, user.id)
    except approved_queue.SapNumberError as e:
        return _back_with_error(request, user, back, iid, str(e))
    except InvalidTransition:
        return _back_with_error(request, user, back, iid, "This invoice can no longer be recorded as entered.")
    return RedirectResponse(back if back.startswith("/") else "/invoices/approved", status_code=303)


def _back_with_error(request, user, back: str, iid, message: str):
    if back.startswith("/invoices/approved"):
        with connection() as conn:
            rows = approved_queue.approved_rows(conn)
        return page(request, user, "invoices/approved.html", {"rows": rows, "error": message, "cents_text": approved_queue.cents_text}, 400)
    with connection() as conn:
        inv = _load(conn, iid)
        return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error=message), 400)


# --- detail ---------------------------------------------------------------------------------------

@router.get("/{invoice_id}")
def detail(request: Request, invoice_id: str, user=Depends(VIEWERS)):
    iid = as_uuid(invoice_id)
    with connection() as conn:
        inv = _load(conn, iid)
        return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user))


@router.get("/{invoice_id}/page/{n}")
def render_page(invoice_id: str, n: int, user=Depends(VIEWERS)):
    """The rendered PNG, never the original file (DATA-2)."""
    iid = as_uuid(invoice_id)
    with connection() as conn:
        row = conn.execute("select ref from invoice_render where invoice_id = %s and page = %s", (iid, n)).fetchone()
    if row is None:
        raise NotFound()
    try:
        from accounting.store import get as store_get
        data = store_get(row["ref"])
    except Exception:
        log.exception("render %s page %s unavailable", iid, n)
        raise NotFound()
    return Response(data, media_type="image/png", headers={"Cache-Control": "no-store"})


# --- AP review edit (SIG-3) -------------------------------------------------------------------------

class FieldError(ValueError):
    pass


def _parse_money(raw: str, label: str) -> int | None:
    raw = (raw or "").strip().replace(",", "").replace("$", "")
    if raw == "":
        return None
    try:
        d = Decimal(raw)
    except InvalidOperation:
        raise FieldError(f"{label} must be a number like 1234.56.")
    if d != d.quantize(Decimal("0.01")):
        raise FieldError(f"{label} can have at most two decimal places.")
    return int(d * 100)


def _parse_date(raw: str, label: str) -> date | None:
    raw = (raw or "").strip()
    if not raw:
        return None
    try:
        return date.fromisoformat(raw)
    except ValueError:
        raise FieldError(f"{label} must be a date.")


def normalize_invoice_number(raw: str | None) -> str | None:
    """The same key the duplicate check uses (DUP-1): no spaces or punctuation, no leading zeros,
    upper-case — so 'inv-00123' and 'INV 123' are one number. Package A owns the rule; the fallback
    here matches it so an edit never produces a key the screening code would not."""
    if not raw:
        return None
    try:
        from accounting.vendors.normalize import invoice_number_norm
        return invoice_number_norm(raw) or None
    except ImportError:
        import re
        s = re.sub(r"[\s\-_./#:]", "", raw).upper().lstrip("0")
        return s or None


@router.post("/{invoice_id}/edit")
def edit(request: Request, invoice_id: str, vendor_name_raw: str = Form(""), invoice_number: str = Form(""),
         invoice_date: str = Form(""), due_date: str = Form(""), currency: str = Form(""), gross: str = Form(""),
         tax: str = Form(""), po_number: str = Form(""), user=Depends(CLERKS), _active=Depends(require_active)):
    iid = as_uuid(invoice_id)
    with connection() as conn:
        inv = _load(conn, iid)
        if inv["state"] != "AP_REVIEW":
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="Fields can only be edited while the invoice is in AP review."), 400)
        try:
            new = {
                "vendor_name_raw": vendor_name_raw.strip() or None, "invoice_number": invoice_number.strip() or None,
                "invoice_date": _parse_date(invoice_date, "Invoice date"), "due_date": _parse_date(due_date, "Due date"),
                "currency": (currency.strip().upper() or None), "gross_cents": _parse_money(gross, "Gross"),
                "tax_cents": _parse_money(tax, "Tax"), "po_number": po_number.strip() or None,
            }
        except FieldError as e:
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error=str(e)), 400)
        def _plain(v):  # dates and Decimals as text so the diff is JSON everywhere it is stored
            return v if v is None or isinstance(v, (int, str)) else str(v)
        diff = {k: {"from": _plain(inv[k]), "to": _plain(v)} for k, v in new.items() if inv[k] != v}
        if diff:
            conn.execute("insert into invoice_correction (invoice_id, user_id, diff) values (%s, %s, %s::jsonb)",
                         (iid, user.id, json.dumps(diff, default=str)))
            sets = ", ".join(f"{k} = %s" for k in new) + ", invoice_number_norm = %s, content_hash = null, updated_at = now()"
            conn.execute(f"update invoice set {sets} where id = %s", [*new.values(), normalize_invoice_number(new["invoice_number"]), iid])
            if conn.execute("select 1 from approval_signature where invoice_id = %s limit 1", (iid,)).fetchone():
                from accounting.approvals.void import void_all_for_invoice
                void_all_for_invoice(conn, iid, "field_edit", user.id)
            log_event(conn, event_type="correction", actor_type="user", actor_id=user.id, invoice_id=iid, payload={"diff": diff})
        transition(conn, iid, "CAPTURED", reason_code="clerk_edit" if diff else "clerk_confirmed", actor_type="user", actor_id=user.id,
                   payload={"fields": sorted(diff)})
        try:
            from accounting.routing.route import route
            route(conn, iid)
        except ImportError:
            log.warning("routing.route not available; %s left in CAPTURED", iid)
    return RedirectResponse(f"/invoices/{iid}", status_code=303)


# --- duplicate / divert / split -----------------------------------------------------------------------

def _simple_transition(request, user, iid, to_state: str, reason: str, payload: dict, actor_type: str = "user", allowed_from=()):
    with connection() as conn:
        inv = _load(conn, iid)
        if allowed_from and inv["state"] not in allowed_from:
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="That action is not available in this invoice's current status."), 400)
        try:
            transition(conn, iid, to_state, reason_code=reason, actor_type=actor_type, actor_id=user.id, payload=payload)
        except InvalidTransition:
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="That action is not available in this invoice's current status."), 400)
    return RedirectResponse(f"/invoices/{iid}", status_code=303)


@router.post("/{invoice_id}/duplicate")
def mark_duplicate(request: Request, invoice_id: str, of_portal_number: str = Form(""), user=Depends(CLERKS), _active=Depends(require_active)):
    iid = as_uuid(invoice_id)
    digits = "".join(ch for ch in of_portal_number if ch.isdigit())
    with connection() as conn:
        inv = _load(conn, iid)
        other = conn.execute("select id from invoice where portal_number = %s and id <> %s", (int(digits) if digits else -1, iid)).fetchone()
        if other is None:
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="Enter the portal number of the original invoice, like AP-000123."), 400)
        conn.execute("update invoice set linked_invoice_id = %s where id = %s", (other["id"], iid))
    return _simple_transition(request, user, iid, "DUPLICATE", "invoice_number", {"duplicate_of": str(other["id"])},
                              allowed_from=("RECEIVED", "AP_REVIEW"))


@router.post("/{invoice_id}/divert")
def divert(request: Request, invoice_id: str, reason: str = Form(""), note: str = Form(""), user=Depends(CLERKS), _active=Depends(require_active)):
    iid = as_uuid(invoice_id)
    if reason not in DIVERT_REASONS:
        with connection() as conn:
            inv = _load(conn, iid)
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="Choose a reason for setting this aside."), 400)
    with connection() as conn:
        conn.execute("update invoice set diverted_reason = %s where id = %s", (reason, iid))
    return _simple_transition(request, user, iid, "DIVERTED", reason, {"note": note.strip()[:500]}, allowed_from=("RECEIVED", "AP_REVIEW"))


@router.post("/{invoice_id}/confirm-split")
def confirm_split(request: Request, invoice_id: str, user=Depends(CLERKS), _active=Depends(require_active)):
    iid = as_uuid(invoice_id)
    with connection() as conn:
        _load(conn, iid)
        conn.execute("update invoice set split_confirmed = true, updated_at = now() where id = %s", (iid,))
        log_event(conn, event_type="split_confirmed", actor_type="user", actor_id=user.id, invoice_id=iid)
    return RedirectResponse(f"/invoices/{iid}", status_code=303)


# --- questions -----------------------------------------------------------------------------------------

@router.post("/{invoice_id}/answer/{question_id}")
def answer(request: Request, invoice_id: str, question_id: str, answer: str = Form(""), user=Depends(CLERKS), _active=Depends(require_active)):
    iid, qid = as_uuid(invoice_id), as_uuid(question_id)
    text = answer.strip()
    with connection() as conn:
        inv = _load(conn, iid)
        if not text:
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="Write an answer before sending it."), 400)
        q = conn.execute("select * from invoice_question where id = %s and invoice_id = %s and answered_at is null for update", (qid, iid)).fetchone()
        if q is None:
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="That question has already been answered."), 400)
        conn.execute("update invoice_question set answer = %s, answered_by = %s, answered_at = now() where id = %s", (text, user.id, qid))
        log_event(conn, event_type="answer", actor_type="user", actor_id=user.id, invoice_id=iid, payload={"question_id": str(qid)})
        still_open = conn.execute("select 1 from invoice_question where invoice_id = %s and answered_at is null limit 1", (iid,)).fetchone()
        if inv["state"] == "QUESTIONED" and not still_open:
            transition(conn, iid, "PENDING_APPROVAL", reason_code="ap_answered", actor_type="user", actor_id=user.id, payload={"question_id": str(qid)})
        queue_notification(conn, type="pending", user_id=q["asked_by"], invoice_id=iid, context={"count": 1})
    return RedirectResponse(f"/invoices/{iid}", status_code=303)


# --- Controller void with a second person -------------------------------------------------------------------

@router.post("/{invoice_id}/void")
def void(request: Request, invoice_id: str, reason: str = Form(""), second_person_id: str = Form(""),
         user=Depends(CONTROLLERS), _active=Depends(require_active)):
    iid = as_uuid(invoice_id)
    with connection() as conn:
        inv = _load(conn, iid)
        if not reason.strip():
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="Give a reason for voiding this invoice."), 400)
        try:
            second = as_uuid(second_person_id)
        except NotFound:
            second = None
        other = conn.execute(
            """select id, display_name from ap_user where id = %s and id <> %s and status = 'active'
               and roles && '{controller,admin}'::portal_role[]""", (second, user.id)).fetchone() if second else None
        if other is None:
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="Choose a second person (a Controller or administrator other than yourself)."), 400)
        try:
            transition(conn, iid, "VOID", reason_code="controller_void", actor_type="controller", actor_id=user.id,
                       payload={"reason": reason.strip()[:500], "by": str(user.id), "second_person": str(other["id"])})
        except InvalidTransition:
            return page(request, user, "invoices/detail.html", _detail_ctx(conn, inv, user, error="This invoice is already closed and cannot be voided."), 400)
        log_event(conn, event_type="second_person", actor_type="second_person", actor_id=other["id"], invoice_id=iid,
                  reason_code="controller_void", payload={"confirmed_for": str(user.id)})
    return RedirectResponse(f"/invoices/{iid}", status_code=303)
