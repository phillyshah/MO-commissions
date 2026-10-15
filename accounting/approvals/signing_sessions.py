"""Signing sessions (D3): one passkey tap covers a sitting, not a single invoice.

A session starts on a successful 'session' assertion and lasts 15 minutes idle / 60 minutes at most
(settings). Inside it, an approval under the step-up amount is a click — still sealed, still recorded
with the session id so the audit trail shows which tap it rode on. At or above the step-up amount, and
for every HCP invoice, the approval needs its own tap regardless of the session.
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
from uuid import UUID

import psycopg

from accounting.config import settings
from accounting.lifecycle import log_event
from accounting.routing.common import step_up_cents


def start(conn: psycopg.Connection, user, passkey_id: UUID | str, portal_session_id: UUID | str, device_hash: str) -> dict:
    """Begin a signing session; any earlier open one for the user is ended so there is exactly one."""
    conn.execute("update signing_session set ended_at = now() where user_id = %s and ended_at is null", (user.id,))
    now = datetime.now(timezone.utc)
    row = conn.execute(
        """insert into signing_session (user_id, passkey_id, portal_session_id, started_at, last_activity_at, expires_at, device_hash)
           values (%s,%s,%s,%s,%s,%s,%s) returning *""",
        (user.id, passkey_id, portal_session_id, now, now, now + timedelta(minutes=settings.signing_session_max_minutes), device_hash),
    ).fetchone()
    log_event(conn, event_type="access", actor_type="user", actor_id=user.id, reason_code="signing_session_started",
              payload={"signing_session": str(row["id"]), "passkey_id": str(passkey_id)})
    return row


def active(conn: psycopg.Connection, user) -> dict | None:
    """The live session, or None. Idle and absolute limits are enforced here, and a stale row is ended
    so it cannot come back to life. The device hash must match the portal session that is asking —
    a signing session is not transferable between browsers."""
    now = datetime.now(timezone.utc)
    row = conn.execute(
        "select * from signing_session where user_id = %s and ended_at is null order by started_at desc limit 1", (user.id,)
    ).fetchone()
    if row is None:
        return None
    idle = now - row["last_activity_at"] > timedelta(minutes=settings.signing_session_idle_minutes)
    expired = now >= row["expires_at"]
    wrong_device = getattr(user, "device_hash", None) not in (None, row["device_hash"])
    wrong_portal_session = getattr(user, "session_id", None) not in (None, row["portal_session_id"])
    if idle or expired or wrong_device or wrong_portal_session:
        conn.execute("update signing_session set ended_at = now() where id = %s", (row["id"],))
        return None
    return row


def touch(conn: psycopg.Connection, session_id: UUID | str) -> None:
    conn.execute("update signing_session set last_activity_at = now() where id = %s and ended_at is null", (session_id,))


def end(conn: psycopg.Connection, user) -> None:
    conn.execute("update signing_session set ended_at = now() where user_id = %s and ended_at is null", (user.id,))


def needs_tap(conn: psycopg.Connection, inv: dict, user) -> bool:
    """True when this approval needs its own passkey tap. `inv` is load_invoice()'s row."""
    gross = int(inv.get("gross_cents") or 0)
    if gross >= step_up_cents(conn):
        return True
    if inv.get("vendor_hcp"):
        return True
    return active(conn, user) is None
