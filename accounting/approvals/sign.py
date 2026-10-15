"""Approve or reject with a signature (APR-3, SIG-1..6, SOD-1/2/6).

Every decision runs the same gate, in this order, and each refusal says in plain words what to do:
  1. the person is active and their sign-in still checks out against the directory (ID-3),
  2. the request is open, for this invoice, and assigned to them (or they are the assignee's delegate),
  3. segregation of duties — never the person who captured, corrected or whose vendor it is,
  4. the content hash is unchanged since the invoice was routed (SIG-3); if not, every signature is
     voided, the invoice goes back to AP, and the approver is told,
  5. the tap rule: an assertion is required when needs_tap() says so, otherwise a live signing session,
  6. the seal — from the signing service, failing closed.
Only then is the signature row written and the request closed. The final approval of a chain moves
the invoice to APPROVED and tells AP.
"""
from __future__ import annotations

import json
from uuid import UUID

import psycopg

from accounting.approvals import payload as payload_mod
from accounting.approvals import signing_service, signing_sessions, void, webauthn_flow
from accounting.identity import sessions
from accounting.lifecycle import log_event, transition
from accounting.routing.common import excluded_users, is_delegate_for, load_invoice, notify_queue

ATTESTATION_ID = "att-v1"


class SignRefused(Exception):
    """A refusal the approver can read. Nothing was signed."""


def current_attestation(conn: psycopg.Connection) -> dict:
    return conn.execute("select * from attestation_text where effective_from <= now() order by effective_from desc limit 1").fetchone() \
        or {"id": ATTESTATION_ID, "text": ""}


def _gate(conn: psycopg.Connection, user, request_id) -> tuple[dict, dict]:
    if getattr(user, "status", "active") != "active":
        raise SignRefused("Your access is frozen. Contact the Controller.")
    if not sessions.background_check(user):
        raise SignRefused("Your sign-in could not be confirmed. Please sign in again; nothing was lost.")
    req = conn.execute("select * from approval_request where id = %s for update", (request_id,)).fetchone()
    if req is None or req["status"] != "open":
        raise SignRefused("This approval is no longer open.")
    if req["assignee_id"] != user.id and not is_delegate_for(conn, user.id, req["assignee_id"]):
        raise SignRefused("This approval is not assigned to you.")
    inv = load_invoice(conn, req["invoice_id"], lock=True)
    if inv["state"] != "PENDING_APPROVAL":
        raise SignRefused("This invoice is not waiting for approval right now.")
    if user.id in excluded_users(conn, inv):
        raise SignRefused("You handled this invoice in AP, so you cannot approve it. Ask AP to route it to someone else.")
    # SIG-3: the content must be what was routed. Compare against what the invoice recorded and against
    # what any earlier signer on this chain signed.
    h = payload_mod.content_hash(conn, inv["id"], inv)
    stored = bytes(inv["content_hash"]) if inv.get("content_hash") else h
    earlier = conn.execute("select content_hash from approval_signature where invoice_id = %s and status = 'valid' order by signed_at desc limit 1",
                           (inv["id"],)).fetchone()
    if h != stored or (earlier and bytes(earlier["content_hash"]) != h):
        void.void_all_for_invoice(conn, inv["id"], "content_changed", user.id)
        conn.execute("update invoice set content_hash = null where id = %s", (inv["id"],))
        transition(conn, inv["id"], "AP_REVIEW", reason_code="field_edit", actor_type="system",
                   payload={"detected_at": "signing", "by": str(user.id)})
        raise SignRefused("This invoice changed after it was sent for approval. Earlier approvals were cancelled and AP will send it again.")
    return req, inv


def _decide(conn: psycopg.Connection, user, request_id, decision: str, assertion: dict | None, reason: str | None) -> UUID:
    req, inv = _gate(conn, user, request_id)
    att = current_attestation(conn)
    passkey_id = None
    session_id = None
    passkey_pub = None
    if signing_sessions.needs_tap(conn, inv, user) or assertion is not None:
        if assertion is None:
            raise SignRefused("This approval needs a passkey tap.")
        pk, ch = webauthn_flow.verify_authentication(conn, user, assertion, "sign", request_id=req["id"])
        if ch["payload"] is None:
            raise SignRefused("This request has expired. Please start again.")
        pl = ch["payload"]
        if pl.get("decision") != decision or pl.get("signer_user_id") != str(user.id) or pl.get("invoice_id") != str(inv["id"]):
            raise SignRefused("The signed request does not match this decision. Please start again.")
        if pl.get("content_hash") != payload_mod.compute_content_hash(inv).hex():
            raise SignRefused("This invoice changed while you were signing. Please open it again.")
        passkey_id = pk["id"]
        passkey_pub = bytes(pk["public_key"])
        # A tap also refreshes the sitting.
        s = signing_sessions.active(conn, user)
        session_id = s["id"] if s else None
        if s:
            signing_sessions.touch(conn, s["id"])
    else:
        s = signing_sessions.active(conn, user)
        if s is None:
            raise SignRefused("Your signing session has ended. Tap your passkey to start a new one.")
        session_id = s["id"]
        passkey_id = s["passkey_id"]
        signing_sessions.touch(conn, s["id"])
        pl = payload_mod.build_payload(conn, inv["id"], user, req["role"], decision, att["id"],
                                       matrix_version_id=req["matrix_version_id"], inv=inv)
    ph = payload_mod.payload_hash(pl)
    try:
        sealed = signing_service.seal(ph, assertion, passkey_pub, pl)
    except signing_service.SealUnavailable as e:
        log_event(conn, event_type="signature", actor_type="system", invoice_id=inv["id"], reason_code="seal_unavailable",
                  payload={"request_id": str(req["id"]), "detail": str(e)[:200]})
        raise SignRefused(signing_service.UNAVAILABLE_TEXT) from e

    if decision == "reject":
        # Other signatures on this chain are cancelled first so the rejection record stands alone as valid.
        void.void_all_for_invoice(conn, inv["id"], "signed_rejection", user.id)
    sig = conn.execute(
        """insert into approval_signature (invoice_id, approval_request_id, signer_id, signer_role, decision, reason, content_hash,
                                           payload, payload_sha256, passkey_id, signing_session_id, assertion, seal, seal_key_id,
                                           attestation_text_id, matrix_version_id)
           values (%s,%s,%s,%s,%s,%s,%s,%s::jsonb,%s,%s,%s,%s::jsonb,%s,%s,%s,%s) returning id""",
        (inv["id"], req["id"], user.id, req["role"], decision, reason, bytes.fromhex(pl["content_hash"]), json.dumps(pl), ph,
         passkey_id, session_id, json.dumps(assertion) if assertion is not None else None, sealed.signature, sealed.key_id,
         att["id"], req["matrix_version_id"]),
    ).fetchone()
    new_status = "signed" if decision == "approve" else "rejected"
    conn.execute("update approval_request set status = %s, closed_at = now() where id = %s and status = 'open'", (new_status, req["id"]))
    log_event(conn, event_type="signature", actor_type="user", actor_id=user.id, invoice_id=inv["id"], reason_code=decision,
              payload={"signature_id": str(sig["id"]), "request_id": str(req["id"]), "role": req["role"], "with_tap": assertion is not None,
                       "signing_session_id": str(session_id) if session_id else None, "seal_key_id": sealed.key_id})

    if decision == "reject":
        transition(conn, inv["id"], "REJECTED", reason_code="signed_rejection", actor_type="user", actor_id=user.id,
                   payload={"signature_id": str(sig["id"]), "reason": reason})
        return sig["id"]

    remaining = conn.execute("select count(*) as n from approval_request where invoice_id = %s and status = 'open'", (inv["id"],)).fetchone()["n"]
    if remaining == 0:
        ids = [str(r["id"]) for r in conn.execute(
            "select id from approval_signature where invoice_id = %s and status = 'valid' and decision = 'approve' order by signed_at", (inv["id"],))]
        transition(conn, inv["id"], "APPROVED", reason_code="chain_complete", actor_type="user", actor_id=user.id,
                   payload={"signature_ids": ids})
        notify_queue(conn, type="approved_for_ap", user_id=None, invoice_id=inv["id"])
    return sig["id"]


def approve(conn: psycopg.Connection, user, request_id, assertion: dict | None = None) -> UUID:
    return _decide(conn, user, request_id, "approve", assertion, None)


def reject(conn: psycopg.Connection, user, request_id, reason: str, assertion: dict | None = None) -> UUID:
    reason = (reason or "").strip()
    if not reason:
        raise SignRefused("A reason is required to reject an invoice.")
    return _decide(conn, user, request_id, "reject", assertion, reason)


def signing_options(conn: psycopg.Connection, user, request_id, decision: str) -> dict:
    """Build the payload for a tap, store it with the challenge (= its hash) and return WebAuthn options.
    Runs the same gate as the decision so an approver is refused before tapping, not after."""
    if decision not in ("approve", "reject"):
        raise SignRefused("Unknown decision.")
    req, inv = _gate(conn, user, request_id)
    att = current_attestation(conn)
    pl = payload_mod.build_payload(conn, inv["id"], user, req["role"], decision, att["id"], matrix_version_id=req["matrix_version_id"], inv=inv)
    ph = payload_mod.payload_hash(pl)
    return webauthn_flow.authentication_options(conn, user, "sign", request_id=req["id"], challenge=ph, payload=pl)
