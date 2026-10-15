"""The Controller's quarantine screen (QUA-1): one list, one case page, three buttons."""
from __future__ import annotations

from datetime import datetime, timezone
from urllib.parse import quote
from uuid import UUID

from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import RedirectResponse

from accounting.config import settings
from accounting.db.pool import connection
from accounting.identity.sessions import CurrentUser
from accounting.quarantine import service
from accounting.routing.common import load_invoice, portal_number
from accounting.web.deps import NotFound, require, require_active
from accounting.web.templating import render

router = APIRouter()
_view = require("controller", "cfo", "admin")


def _act(user: CurrentUser = Depends(require_active)) -> CurrentUser:
    if not user.has("controller", "cfo", "admin"):
        raise NotFound()
    return user


@router.get("/quarantine")
def quarantine_list(request: Request, user: CurrentUser = Depends(_view), show: str = "open"):
    with connection() as conn:
        state_filter = "q.state = 'open'" if show == "open" else "q.state <> 'open'"
        rows = conn.execute(
            f"""select q.*, i.portal_number, i.gross_cents, i.currency, i.invoice_number, i.received_at, v.legal_name as vendor_name,
                       m.from_addr, m.original_sender, m.subject
                from quarantine_case q left join invoice i on i.id = q.invoice_id left join vendor v on v.id = i.vendor_id
                left join message m on m.id = q.message_id
                where {state_filter} order by q.opened_at {'asc' if show == 'open' else 'desc'} limit 200"""
        ).fetchall()
    now = datetime.now(timezone.utc)
    for r in rows:
        r["portal"] = portal_number(r["portal_number"]) if r["portal_number"] else "—"
        r["escalated"] = service.is_escalated(r, now)
        r["trigger_labels"] = [t.replace("_", " ") for t in (r["triggers"] or [])]
    return render(request, "quarantine/list.html", {"rows": rows, "show": show, "as_of": now})


@router.get("/quarantine/{case_id}")
def case_page(request: Request, case_id: UUID, user: CurrentUser = Depends(_view), error: str | None = None, done: str | None = None):
    with connection() as conn:
        case = conn.execute("select * from quarantine_case where id = %s", (case_id,)).fetchone()
        if case is None:
            raise NotFound()
        inv = load_invoice(conn, case["invoice_id"]) if case["invoice_id"] else None
        msg = conn.execute("select id, from_addr, from_name, original_sender, forwarded_by, subject, received_at, trust from message where id = %s",
                           (case["message_id"],)).fetchone() if case["message_id"] else None
        pages = [p["page"] for p in conn.execute("select page from invoice_render where invoice_id = %s order by page", (case["invoice_id"],))] if inv else []
        history = service.vendor_history(conn, inv["vendor_id"], inv["id"]) if inv else []
        fp = service.fingerprint_comparison(conn, inv)
        reviewers = conn.execute(
            """select id, display_name from ap_user where status = 'active' and id <> %s and roles && '{controller,cfo}'::portal_role[]
               order by display_name""", (user.id,)).fetchall()
        reviewer = conn.execute("select display_name from ap_user where id = %s", (case["reviewer_id"],)).fetchone() if case["reviewer_id"] else None
        second = conn.execute("select display_name from ap_user where id = %s", (case["second_reviewer_id"],)).fetchone() if case["second_reviewer_id"] else None
        needs_callback = service.needs_callback(case)
        needs_second = service.needs_second_person(conn, case, inv)
        sender_domain = service.sender_domain(conn, case) if msg else None
    now = datetime.now(timezone.utc)
    return render(request, "quarantine/case.html", {
        "case": case, "inv": inv, "msg": msg, "pages": pages, "history": history, "fp": fp, "reviewers": reviewers,
        "portal": portal_number(inv["portal_number"]) if inv else "—", "escalated": service.is_escalated(case, now),
        "why": [(t, service.trigger_text(t)) for t in (case["triggers"] or [])],
        "needs_callback": needs_callback, "needs_second": needs_second, "callback_sources": service.CALLBACK_SOURCES,
        "i_am_second": case["second_reviewer_id"] == user.id and case["release_pending_at"] is not None,
        "reviewer_name": reviewer["display_name"] if reviewer else None, "second_name": second["display_name"] if second else None,
        "second_person_cents": settings.second_person_release_cents, "sender_domain": sender_domain,
        "error": error, "done": done, "as_of": now, "can_act": user.status == "active",
    })


@router.post("/quarantine/{case_id}/release")
def release(request: Request, case_id: UUID, callback_source: str = Form(""), callback_number: str = Form(""), note: str = Form(""),
            second_by: str = Form(""), user: CurrentUser = Depends(_act)):
    try:
        with connection() as conn:
            result = service.release(conn, case_id, user.id, UUID(second_by) if second_by else None, callback_source or None,
                                     callback_number, note)
    except service.QuarantineError as e:
        return RedirectResponse(f"/quarantine/{case_id}?error={quote(str(e), safe='')}", status_code=303)
    if result == "pending_second_person":
        return RedirectResponse(f"/quarantine/{case_id}?done=pending", status_code=303)
    return RedirectResponse("/quarantine?done=released", status_code=303)


@router.post("/quarantine/{case_id}/reject")
def reject(request: Request, case_id: UUID, reason: str = Form(""), user: CurrentUser = Depends(_act)):
    try:
        with connection() as conn:
            service.reject(conn, case_id, user.id, reason)
    except service.QuarantineError as e:
        return RedirectResponse(f"/quarantine/{case_id}?error={quote(str(e), safe='')}", status_code=303)
    return RedirectResponse("/quarantine?done=rejected", status_code=303)


@router.post("/quarantine/{case_id}/report")
def report(request: Request, case_id: UUID, user: CurrentUser = Depends(_act)):
    try:
        with connection() as conn:
            service.report(conn, case_id, user.id)
    except service.QuarantineError as e:
        return RedirectResponse(f"/quarantine/{case_id}?error={quote(str(e), safe='')}", status_code=303)
    return RedirectResponse("/quarantine?done=reported", status_code=303)
