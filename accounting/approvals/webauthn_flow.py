"""WebAuthn enrollment and assertions (ID-4, D3) on top of py_webauthn.

Three kinds of ceremony share one challenge table:
  register — needs a live portal session AND an unused enrollment code the Controller handed over
             out-of-band; the code is compared by sha256, never stored or logged in clear.
  session  — a tap that opens a signing session (the challenge is random).
  sign     — a tap over one approval: the challenge IS the payload hash, so what the authenticator
             signs is what the approver saw, and the payload is stored beside it.

Challenges are written server-side when options are issued and consumed exactly once within two
minutes; the browser never chooses them. `userVerification: required` everywhere — a tap without PIN
or biometric would make a stolen laptop an approver.
"""
from __future__ import annotations

import hashlib
import json
import secrets
from datetime import datetime, timedelta, timezone
from uuid import UUID

import psycopg
from webauthn import (generate_authentication_options, generate_registration_options, options_to_json,
                      verify_authentication_response, verify_registration_response)
from webauthn.helpers import base64url_to_bytes, bytes_to_base64url
from webauthn.helpers.exceptions import InvalidAuthenticationResponse, InvalidRegistrationResponse
from webauthn.helpers.structs import (AuthenticatorSelectionCriteria, PublicKeyCredentialDescriptor, ResidentKeyRequirement,
                                      UserVerificationRequirement)

from accounting.config import settings
from accounting.lifecycle import log_event
from accounting.routing.common import notify_queue, users_with_role

CHALLENGE_TTL = timedelta(minutes=2)


class WebAuthnError(Exception):
    """Refused ceremony; the message is for the screen."""


def expected_origin() -> str:
    return settings.portal_base_url.rstrip("/")


def _store_challenge(conn: psycopg.Connection, user_id, kind: str, challenge: bytes, request_id=None, payload: dict | None = None) -> dict:
    # One live challenge per (user, kind, request): issuing new options kills the previous ones.
    conn.execute(
        """update webauthn_challenge set used_at = now() where user_id = %s and kind = %s and used_at is null
           and coalesce(approval_request_id::text, '') = coalesce(%s::text, '')""",
        (user_id, kind, request_id),
    )
    return conn.execute(
        """insert into webauthn_challenge (user_id, kind, challenge, approval_request_id, payload)
           values (%s,%s,%s,%s,%s::jsonb) returning *""",
        (user_id, kind, challenge, request_id, json.dumps(payload) if payload is not None else None),
    ).fetchone()


def _consume_challenge(conn: psycopg.Connection, user_id, kind: str, request_id=None) -> dict:
    row = conn.execute(
        """select * from webauthn_challenge where user_id = %s and kind = %s and used_at is null
           and coalesce(approval_request_id::text, '') = coalesce(%s::text, '')
           order by created_at desc limit 1 for update""",
        (user_id, kind, request_id),
    ).fetchone()
    if row is None:
        raise WebAuthnError("This request has expired. Please start again.")
    conn.execute("update webauthn_challenge set used_at = now() where id = %s", (row["id"],))
    if datetime.now(timezone.utc) - row["created_at"] > CHALLENGE_TTL:
        raise WebAuthnError("This request has expired. Please start again.")
    return row


def _options_dict(options) -> dict:
    return json.loads(options_to_json(options))


def active_passkeys(conn: psycopg.Connection, user_id) -> list[dict]:
    return conn.execute("select * from passkey where user_id = %s and status = 'active' order by created_at", (user_id,)).fetchall()


# --- registration ---------------------------------------------------------------------------------

def _valid_enrollment_code(conn: psycopg.Connection, user_id, code: str) -> dict | None:
    h = hashlib.sha256(code.strip().encode()).digest()
    return conn.execute(
        "select * from enrollment_code where user_id = %s and code_hash = %s and used_at is null and expires_at > now() for update",
        (user_id, h),
    ).fetchone()


def registration_options(conn: psycopg.Connection, user, code: str) -> dict:
    if not code or _valid_enrollment_code(conn, user.id, code) is None:
        raise WebAuthnError("That enrollment code is not valid. Codes come from the Controller and expire; ask for a new one.")
    challenge = secrets.token_bytes(32)
    opts = generate_registration_options(
        rp_id=settings.rp_id, rp_name=settings.rp_name,
        user_id=user.id.bytes, user_name=user.email, user_display_name=user.display_name,
        challenge=challenge,
        authenticator_selection=AuthenticatorSelectionCriteria(
            resident_key=ResidentKeyRequirement.PREFERRED, user_verification=UserVerificationRequirement.REQUIRED),
        exclude_credentials=[PublicKeyCredentialDescriptor(id=bytes(p["credential_id"])) for p in active_passkeys(conn, user.id)],
    )
    _store_challenge(conn, user.id, "register", challenge)
    return _options_dict(opts)


def verify_registration(conn: psycopg.Connection, user, credential: dict | str, code: str, label: str | None = None) -> dict:
    code_row = _valid_enrollment_code(conn, user.id, code or "")
    if code_row is None:
        raise WebAuthnError("That enrollment code is not valid. Codes come from the Controller and expire; ask for a new one.")
    ch = _consume_challenge(conn, user.id, "register")
    try:
        v = verify_registration_response(
            credential=credential, expected_challenge=bytes(ch["challenge"]), expected_rp_id=settings.rp_id,
            expected_origin=expected_origin(), require_user_verification=True,
        )
    except InvalidRegistrationResponse as e:
        raise WebAuthnError("The passkey could not be registered. Please try again.") from e
    cred_json = credential if isinstance(credential, dict) else json.loads(credential)
    transports = (cred_json.get("response") or {}).get("transports") or None
    conn.execute("update enrollment_code set used_at = now() where id = %s", (code_row["id"],))
    row = conn.execute(
        """insert into passkey (user_id, credential_id, public_key, sign_count, attestation, transports, label)
           values (%s,%s,%s,%s,%s::jsonb,%s,%s) returning *""",
        (user.id, v.credential_id, v.credential_public_key, int(v.sign_count),
         json.dumps({"fmt": str(v.fmt.value if hasattr(v.fmt, 'value') else v.fmt), "aaguid": v.aaguid,
                     "attestation_object": bytes_to_base64url(v.attestation_object),
                     "device_type": str(getattr(v.credential_device_type, 'value', v.credential_device_type)),
                     "backed_up": bool(v.credential_backed_up)}),
         transports, (label or "").strip()[:80] or None),
    ).fetchone()
    seq = log_event(conn, event_type="security", actor_type="user", actor_id=user.id, reason_code="passkey_enrolled",
                    payload={"passkey_id": str(row["id"]), "label": row["label"], "aaguid": v.aaguid, "code_id": str(code_row["id"])})
    conn.execute("insert into security_event (type, user_id, source, detail, event_seq) values ('passkey_enrolled', %s, 'portal', %s::jsonb, %s)",
                 (user.id, json.dumps({"passkey_id": str(row["id"])}), seq))
    # ID-4: enrollment alerts the user and the Controller.
    ctx = {"label": row["label"], "when": datetime.now(timezone.utc).isoformat(), "user_email": user.email}
    notify_queue(conn, type="enroll_alert", user_id=user.id, context=ctx)
    for c in users_with_role(conn, "controller"):
        if c["id"] != user.id:
            notify_queue(conn, type="enroll_alert", user_id=c["id"], context=ctx)
    return row


def revoke_passkey(conn: psycopg.Connection, user, passkey_id) -> bool:
    n = conn.execute("update passkey set status = 'revoked' where id = %s and user_id = %s and status = 'active'", (passkey_id, user.id)).rowcount
    if n:
        conn.execute("update signing_session set ended_at = now() where passkey_id = %s and ended_at is null", (passkey_id,))
        log_event(conn, event_type="security", actor_type="user", actor_id=user.id, reason_code="passkey_revoked", payload={"passkey_id": str(passkey_id)})
    return bool(n)


# --- authentication -------------------------------------------------------------------------------

def authentication_options(conn: psycopg.Connection, user, kind: str, *, request_id=None, challenge: bytes | None = None,
                           payload: dict | None = None) -> dict:
    keys = active_passkeys(conn, user.id)
    if not keys:
        raise WebAuthnError("You have no passkey yet. Enroll one under My passkeys first.")
    challenge = challenge or secrets.token_bytes(32)
    opts = generate_authentication_options(
        rp_id=settings.rp_id, challenge=challenge, user_verification=UserVerificationRequirement.REQUIRED,
        allow_credentials=[PublicKeyCredentialDescriptor(id=bytes(p["credential_id"])) for p in keys],
    )
    _store_challenge(conn, user.id, kind, challenge, request_id, payload)
    return _options_dict(opts)


def verify_authentication(conn: psycopg.Connection, user, credential: dict | str, kind: str, *, request_id=None) -> tuple[dict, dict]:
    """Verify an assertion against the stored challenge. Returns (passkey_row, challenge_row).
    Bumps sign_count and last_used_at — a sign_count going backwards is how a cloned authenticator
    shows itself, and py_webauthn refuses it."""
    ch = _consume_challenge(conn, user.id, kind, request_id)
    cred_json = credential if isinstance(credential, dict) else json.loads(credential)
    try:
        raw_id = base64url_to_bytes(cred_json["rawId"])
    except Exception as e:
        raise WebAuthnError("The passkey response was not understood.") from e
    pk = conn.execute("select * from passkey where credential_id = %s and user_id = %s and status = 'active' for update",
                      (raw_id, user.id)).fetchone()
    if pk is None:
        raise WebAuthnError("That passkey is not registered to you or has been revoked.")
    try:
        v = verify_authentication_response(
            credential=credential, expected_challenge=bytes(ch["challenge"]), expected_rp_id=settings.rp_id,
            expected_origin=expected_origin(), credential_public_key=bytes(pk["public_key"]),
            credential_current_sign_count=int(pk["sign_count"]), require_user_verification=True,
        )
    except InvalidAuthenticationResponse as e:
        log_event(conn, event_type="security", actor_type="user", actor_id=user.id, reason_code="assertion_failed",
                  payload={"kind": kind, "passkey_id": str(pk["id"]), "detail": str(e)[:200]})
        raise WebAuthnError("The passkey check failed. Please try again.") from e
    conn.execute("update passkey set sign_count = %s, last_used_at = now() where id = %s", (int(v.new_sign_count), pk["id"]))
    return pk, ch
