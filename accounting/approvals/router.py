"""Approver-facing pages and the WebAuthn JSON endpoints.

Pages are plain forms; the few JavaScript steps (passkey create/get) live in static/webauthn.js and
POST JSON here. Every JSON endpoint answers {ok, ...} or {ok:false, error:"plain words"} so the page
can show the message without interpreting anything. Signing decisions never happen in a GET.
"""
from __future__ import annotations

import json
import logging
from datetime import date, datetime, timezone
from uuid import UUID

from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import JSONResponse, RedirectResponse

from accounting.approvals import sign, signing_sessions, webauthn_flow
from accounting.db.pool import connection
from accounting.identity import sessions
from accounting.identity.sessions import CurrentUser
from accounting.lifecycle import InvalidTransition, log_event, transition
from accounting.routing import reassign as reassign_mod
from accounting.routing.common import is_delegate_for, load_invoice, portal_number, step_up_cents
from accounting.web.deps import NotFound, current_user, require_active
from accounting.web.templating import render

log = logging.getLogger("accounting.approvals")
router = APIRouter()

SIGNER_ROLES = ("approver", "cfo", "compliance", "controller", "supervisor", "admin")


def _json_error(text: str, status: int = 400) -> JSONResponse:
    return JSONResponse({"ok": False, "error": text}, status_code=status)


async def _body(request: Request) -> dict:
    try:
        data = await request.json()
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def _my_open_requests(conn, user: CurrentUser) -> list[dict]:
    return conn.execute(
        """select r.*, i.portal_number, i.gross_cents, i.currency, i.due_date, i.received_at, i.soft_signals, i.tags, i.state,
                  i.invoice_number, v.legal_name as vendor_name, v.hcp as vendor_hcp
           from approval_request r join invoice i on i.id = r.invoice_id left join vendor v on v.id = i.vendor_id
           where r.status = 'open' and i.state = 'PENDING_APPROVAL'
             and (r.assignee_id = %s or r.assignee_id in (
                   select from_user_id from delegation where to_user_id = %s and revoked_at is null
                     and starts_on <= current_date and ends_on >= current_date))
           order by r.created_at""",
        (user.id, user.id),
    ).fetchall()


# --- pages ----------------------------------------------------------------------------------------

@router.get("/approvals")
def approvals_list(request: Request, user: CurrentUser = Depends(current_user)):
    with connection() as conn:
        rows = _my_open_requests(conn, user)
        session = signing_sessions.active(conn, user)
        has_passkey = bool(webauthn_flow.active_passkeys(conn, user.id))
    for r in rows:
        r["portal"] = portal_number(r["portal_number"])
    return render(request, "approvals/list.html", {"rows": rows, "signing_session": session, "has_passkey": has_passkey,
                                                   "as_of": datetime.now(timezone.utc)})


def _request_for(conn, request_id, user: CurrentUser) -> dict:
    req = conn.execute("select * from approval_request where id = %s", (request_id,)).fetchone()
    if req is None:
        raise NotFound()
    mine = req["assignee_id"] == user.id or is_delegate_for(conn, user.id, req["assignee_id"])
    if not mine and not user.has("controller", "admin", "auditor"):
        raise NotFound()
    return req


@router.get("/approvals/{request_id}")
def approval_detail(request: Request, request_id: UUID, user: CurrentUser = Depends(current_user), error: str | None = None):
    # ID-3: the directory is asked again on every invoice page load.
    if not sessions.background_check(user):
        return RedirectResponse("/login?next=/approvals", status_code=303)
    with connection() as conn:
        req = _request_for(conn, request_id, user)
        inv = load_invoice(conn, req["invoice_id"])
        pages = conn.execute("select page from invoice_render where invoice_id = %s order by page", (inv["id"],)).fetchall()
        history = conn.execute(
            """select id, portal_number, invoice_number, invoice_date, gross_cents, currency, state from invoice
               where vendor_id = %s and id <> %s order by received_at desc limit 6""",
            (inv["vendor_id"], inv["id"]),
        ).fetchall() if inv.get("vendor_id") else []
        questions = conn.execute("select * from invoice_question where invoice_id = %s order by asked_at", (inv["id"],)).fetchall()
        delegates = conn.execute(
            """select id, display_name, email from ap_user where status = 'active' and id <> %s
               and roles && %s::portal_role[] order by display_name""",
            (user.id, list(SIGNER_ROLES)),
        ).fetchall()
        attestation = sign.current_attestation(conn)
        needs_tap = signing_sessions.needs_tap(conn, inv, user)
        step_up = step_up_cents(conn)
        mine = req["assignee_id"] == user.id or is_delegate_for(conn, user.id, req["assignee_id"])
    fields = _fields(inv)
    return render(request, "approvals/detail.html", {
        "req": req, "inv": inv, "portal": portal_number(inv["portal_number"]), "pages": [p["page"] for p in pages],
        "fields": fields, "history": history, "questions": questions, "delegates": delegates, "attestation": attestation,
        "needs_tap": needs_tap, "step_up_cents": step_up, "can_act": mine and req["status"] == "open" and user.status == "active",
        "error": error, "as_of": datetime.now(timezone.utc),
    })


def _fields(inv: dict) -> list[dict]:
    conf = inv.get("field_confidence") or {}
    threshold = 0.85

    def f(label, key, value):
        c = conf.get(key)
        return {"label": label, "value": value, "confidence": c, "low": c is not None and float(c) < threshold}

    return [
        f("Vendor", "vendor_name", inv.get("vendor_legal_name") or inv.get("vendor_name_raw")),
        f("Invoice number", "invoice_number", inv.get("invoice_number")),
        f("Invoice date", "invoice_date", inv.get("invoice_date")),
        f("Due date", "due_date", inv.get("due_date")),
        f("Amount", "gross", inv.get("gross_cents")),
        f("Tax", "tax", inv.get("tax_cents")),
        f("Currency", "currency", inv.get("currency")),
        f("PO number", "po_number", inv.get("po_number")),
        f("GL / cost center", "gl", " / ".join(x for x in (inv.get("vendor_default_gl"), inv.get("vendor_default_cost_center")) if x) or None),
        f("Bank (last 4)", "bank_last4", inv.get("bank_last4")),
    ]


# --- form actions --------------------------------------------------------------------------------

@router.post("/approvals/{request_id}/approve")
def approve_click(request: Request, request_id: UUID, user: CurrentUser = Depends(require_active)):
    """The in-session click (no tap). Taps arrive through the JSON verify endpoint instead."""
    try:
        with connection() as conn:
            sign.approve(conn, user, request_id, None)
    except sign.SignRefused as e:
        return RedirectResponse(f"/approvals/{request_id}?error={_q(str(e))}", status_code=303)
    return RedirectResponse("/approvals?done=approved", status_code=303)


@router.post("/approvals/{request_id}/reject")
def reject_click(request: Request, request_id: UUID, reason: str = Form(""), user: CurrentUser = Depends(require_active)):
    try:
        with connection() as conn:
            sign.reject(conn, user, request_id, reason, None)
    except sign.SignRefused as e:
        return RedirectResponse(f"/approvals/{request_id}?error={_q(str(e))}", status_code=303)
    return RedirectResponse("/approvals?done=rejected", status_code=303)


@router.post("/approvals/{request_id}/reassign")
def reassign_action(request: Request, request_id: UUID, to_user: UUID = Form(...), reason: str = Form(""),
                    user: CurrentUser = Depends(require_active)):
    try:
        with connection() as conn:
            req = _request_for(conn, request_id, user)
            if req["assignee_id"] != user.id and not user.has("controller", "admin"):
                raise reassign_mod.ReassignError("Only the assigned approver or the Controller can reassign this.")
            reassign_mod.reassign(conn, request_id, to_user, user.id, reason.strip() or "reassigned by approver")
    except reassign_mod.ReassignError as e:
        return RedirectResponse(f"/approvals/{request_id}?error={_q(str(e))}", status_code=303)
    return RedirectResponse("/approvals?done=reassigned", status_code=303)


@router.post("/approvals/{request_id}/question")
def ask_question(request: Request, request_id: UUID, question: str = Form(...), user: CurrentUser = Depends(require_active)):
    q = question.strip()
    if not q:
        return RedirectResponse(f"/approvals/{request_id}?error={_q('Type your question first.')}", status_code=303)
    try:
        with connection() as conn:
            req = _request_for(conn, request_id, user)
            conn.execute("insert into invoice_question (invoice_id, asked_by, question) values (%s, %s, %s)", (req["invoice_id"], user.id, q))
            transition(conn, req["invoice_id"], "QUESTIONED", reason_code="approver_question", actor_type="user", actor_id=user.id,
                       payload={"request_id": str(request_id), "question": q[:500]})
    except InvalidTransition:
        return RedirectResponse(f"/approvals/{request_id}?error={_q('This invoice is not waiting for approval right now.')}", status_code=303)
    return RedirectResponse("/approvals?done=questioned", status_code=303)


# --- WebAuthn JSON: signing session --------------------------------------------------------------

@router.post("/approvals/webauthn/session/options")
async def session_options(request: Request, user: CurrentUser = Depends(require_active)):
    try:
        with connection() as conn:
            opts = webauthn_flow.authentication_options(conn, user, "session")
    except webauthn_flow.WebAuthnError as e:
        return _json_error(str(e))
    return {"ok": True, "options": opts}


@router.post("/approvals/webauthn/session/verify")
async def session_verify(request: Request, user: CurrentUser = Depends(require_active)):
    body = await _body(request)
    cred = body.get("credential")
    if not cred:
        return _json_error("No passkey response was received.")
    if not sessions.background_check(user):
        return _json_error("Your sign-in could not be confirmed. Please sign in again.", 401)
    try:
        with connection() as conn:
            pk, _ch = webauthn_flow.verify_authentication(conn, user, cred, "session")
            row = signing_sessions.start(conn, user, pk["id"], user.session_id, user.device_hash)
    except webauthn_flow.WebAuthnError as e:
        return _json_error(str(e))
    return {"ok": True, "signing_session_id": str(row["id"]), "expires_at": row["expires_at"].isoformat(),
            "redirect": body.get("next") or "/approvals"}


# --- WebAuthn JSON: one approval ------------------------------------------------------------------

@router.post("/approvals/webauthn/sign/{request_id}/options")
async def sign_options(request: Request, request_id: UUID, user: CurrentUser = Depends(require_active)):
    body = await _body(request)
    decision = body.get("decision", "approve")
    try:
        with connection() as conn:
            opts = sign.signing_options(conn, user, request_id, decision)
    except (sign.SignRefused, webauthn_flow.WebAuthnError) as e:
        return _json_error(str(e))
    return {"ok": True, "options": opts}


@router.post("/approvals/webauthn/sign/{request_id}/verify")
async def sign_verify(request: Request, request_id: UUID, user: CurrentUser = Depends(require_active)):
    body = await _body(request)
    cred, decision, reason = body.get("credential"), body.get("decision", "approve"), (body.get("reason") or "").strip()
    if not cred:
        return _json_error("No passkey response was received.")
    try:
        with connection() as conn:
            if decision == "reject":
                sid = sign.reject(conn, user, request_id, reason, cred)
            else:
                sid = sign.approve(conn, user, request_id, cred)
    except (sign.SignRefused, webauthn_flow.WebAuthnError) as e:
        return _json_error(str(e))
    return {"ok": True, "signature_id": str(sid), "redirect": f"/approvals?done={decision}d" if decision == "approve" else "/approvals?done=rejected"}


# --- passkeys -------------------------------------------------------------------------------------

@router.get("/passkeys")
def passkeys_page(request: Request, user: CurrentUser = Depends(current_user), error: str | None = None, done: str | None = None):
    with connection() as conn:
        keys = conn.execute("select * from passkey where user_id = %s order by status, created_at", (user.id,)).fetchall()
    return render(request, "approvals/passkeys.html", {"keys": keys, "error": error, "done": done})


@router.post("/passkeys/webauthn/register/options")
async def register_options(request: Request, user: CurrentUser = Depends(require_active)):
    body = await _body(request)
    try:
        with connection() as conn:
            opts = webauthn_flow.registration_options(conn, user, body.get("code", ""))
    except webauthn_flow.WebAuthnError as e:
        return _json_error(str(e))
    return {"ok": True, "options": opts}


@router.post("/passkeys/webauthn/register/verify")
async def register_verify(request: Request, user: CurrentUser = Depends(require_active)):
    body = await _body(request)
    cred = body.get("credential")
    if not cred:
        return _json_error("No passkey response was received.")
    try:
        with connection() as conn:
            row = webauthn_flow.verify_registration(conn, user, cred, body.get("code", ""), body.get("label"))
    except webauthn_flow.WebAuthnError as e:
        return _json_error(str(e))
    return {"ok": True, "passkey_id": str(row["id"]), "redirect": "/passkeys?done=enrolled"}


@router.post("/passkeys/{passkey_id}/revoke")
def revoke_passkey(request: Request, passkey_id: UUID, user: CurrentUser = Depends(require_active)):
    with connection() as conn:
        ok = webauthn_flow.revoke_passkey(conn, user, passkey_id)
    return RedirectResponse("/passkeys?done=revoked" if ok else "/passkeys?error=" + _q("That passkey was not found."), status_code=303)


# --- delegations ----------------------------------------------------------------------------------

@router.get("/delegations")
def delegations_page(request: Request, user: CurrentUser = Depends(current_user), error: str | None = None, done: str | None = None):
    with connection() as conn:
        mine = conn.execute(
            """select d.*, u.display_name as to_name from delegation d join ap_user u on u.id = d.to_user_id
               where d.from_user_id = %s order by d.revoked_at nulls first, d.starts_on desc""", (user.id,)).fetchall()
        for_me = conn.execute(
            """select d.*, u.display_name as from_name from delegation d join ap_user u on u.id = d.from_user_id
               where d.to_user_id = %s and d.revoked_at is null and d.ends_on >= current_date order by d.starts_on""", (user.id,)).fetchall()
        people = conn.execute(
            "select id, display_name from ap_user where status = 'active' and id <> %s and roles && %s::portal_role[] order by display_name",
            (user.id, list(SIGNER_ROLES))).fetchall()
    return render(request, "approvals/delegations.html", {"mine": mine, "for_me": for_me, "people": people, "error": error, "done": done,
                                                          "today": date.today()})


@router.post("/delegations")
def create_delegation(request: Request, to_user: UUID = Form(...), starts_on: date = Form(...), ends_on: date = Form(...),
                      user: CurrentUser = Depends(require_active)):
    if to_user == user.id:
        return RedirectResponse("/delegations?error=" + _q("You cannot delegate to yourself."), status_code=303)
    if ends_on < starts_on or ends_on < date.today():
        return RedirectResponse("/delegations?error=" + _q("The end date must be today or later, and not before the start."), status_code=303)
    with connection() as conn:
        target = conn.execute("select * from ap_user where id = %s and status = 'active'", (to_user,)).fetchone()
        if target is None:
            return RedirectResponse("/delegations?error=" + _q("That person cannot take approvals right now."), status_code=303)
        # APR-8: a delegate cannot sub-delegate — someone currently acting for others cannot pass that on,
        # and someone who has delegated away cannot be chosen as a delegate.
        if conn.execute("select 1 from delegation where to_user_id = %s and revoked_at is null and ends_on >= current_date", (user.id,)).fetchone():
            return RedirectResponse("/delegations?error=" + _q("You are currently acting as someone's delegate, so you cannot delegate your own approvals."), status_code=303)
        if conn.execute("select 1 from delegation where from_user_id = %s and revoked_at is null and ends_on >= %s and starts_on <= %s", (to_user, ends_on, starts_on)).fetchone():
            return RedirectResponse("/delegations?error=" + _q("That person has delegated their own approvals for those dates, so they cannot take yours."), status_code=303)
        row = conn.execute("insert into delegation (from_user_id, to_user_id, starts_on, ends_on) values (%s,%s,%s,%s) returning id",
                           (user.id, to_user, starts_on, ends_on)).fetchone()
        log_event(conn, event_type="routing", actor_type="user", actor_id=user.id, reason_code="delegation_created",
                  payload={"delegation_id": str(row["id"]), "to_user": str(to_user), "starts_on": starts_on.isoformat(), "ends_on": ends_on.isoformat()})
    return RedirectResponse("/delegations?done=created", status_code=303)


@router.post("/delegations/{delegation_id}/revoke")
def revoke_delegation(request: Request, delegation_id: UUID, user: CurrentUser = Depends(require_active)):
    with connection() as conn:
        n = conn.execute("update delegation set revoked_at = now() where id = %s and from_user_id = %s and revoked_at is null",
                         (delegation_id, user.id)).rowcount
        if n:
            log_event(conn, event_type="routing", actor_type="user", actor_id=user.id, reason_code="delegation_revoked",
                      payload={"delegation_id": str(delegation_id)})
    return RedirectResponse("/delegations?done=revoked", status_code=303)


def _q(text: str) -> str:
    from urllib.parse import quote
    return quote(text, safe="")
