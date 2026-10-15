"""The notification queue: `queue()` writes a row, the worker's `flush()` sends it.

Why a queue rather than sending inline: a sign-off or an approval must not fail because the relay
is slow, and a relay outage must not lose a notice. The row is the record; the email is a
best-effort delivery of it, retried at most three times so a dead relay reads "failed" instead of
retrying forever.

Why the body is rendered at flush time: the "Freeze my access" token is minted for each email and
has a 14-day life. Minting it when the mail actually goes out — not when it was queued — keeps the
token's clock honest.
"""
from __future__ import annotations

import json
import logging
from uuid import UUID

import psycopg

from accounting.config import settings
from accounting.db.pool import connection
from accounting.notify import smtp, templates

log = logging.getLogger("accounting.notify.send")

MAX_ATTEMPTS = 3

# Types that have no single recipient: they go to a role.
ROLE_RECIPIENTS = {
    "approved_for_ap": ("clerk", "supervisor"),
    "vendor_setup": ("clerk", "supervisor"),
    "heartbeat_alert": ("controller",),
    "controller_daily": ("controller",),
    "security_report": ("controller",),
    "controller_reassign": ("controller",),
}


def _recipients(conn: psycopg.Connection, type_: str, user_id, to: str | None) -> list[tuple[UUID | None, str]]:
    """Who gets this: an explicit address, the user's address, or every active holder of the role.
    The Controller's configured address is the fallback so a control email is never silently dropped
    because nobody has the role yet."""
    if to:
        row = conn.execute("select id from ap_user where lower(email) = lower(%s)", (to,)).fetchone()
        return [(row["id"] if row else user_id, to)]
    if user_id is not None:
        row = conn.execute("select id, email from ap_user where id = %s", (user_id,)).fetchone()
        return [(row["id"], row["email"])] if row else []
    roles = ROLE_RECIPIENTS.get(type_)
    if roles:
        rows = conn.execute(
            "select id, email from ap_user where status = 'active' and roles && %s::portal_role[] order by email",
            (list(roles),),
        ).fetchall()
        if rows:
            return [(r["id"], r["email"]) for r in rows]
        if "controller" in roles and settings.controller_email:
            return [(None, settings.controller_email)]
    return []


def queue(conn: psycopg.Connection, *, type: str, user_id, invoice_id=None, to: str | None = None,
          context: dict | None = None):
    """Insert one queued row per recipient; returns the first notification id (None if nobody)."""
    first = None
    for rid, addr in _recipients(conn, type, user_id, to):
        row = conn.execute(
            """insert into notification (user_id, to_addr, type, template_version, related_invoice_id, context)
               values (%s, %s, %s, %s, %s, %s::jsonb) returning id""",
            (rid, addr, type, templates.TEMPLATE_VERSION, invoice_id, json.dumps(context or {}, default=str)),
        ).fetchone()
        first = first or row["id"]
    if first is None:
        log.warning("notification of type %s had no recipient", type)
    return first


def _freeze_url(conn: psycopg.Connection, user_id) -> str | None:
    if user_id is None:
        return None
    from accounting.identity.router import issue_freeze_token
    token = issue_freeze_token(conn, user_id)
    return f"{settings.portal_base_url}/freeze/{token}"


def _send_one(notification_id) -> bool:
    """One row, one transaction, so a failure here never rolls back a neighbour's success."""
    with connection() as conn:
        n = conn.execute("select * from notification where id = %s for update skip locked", (notification_id,)).fetchone()
        if n is None or n["delivery_status"] == "sent" or n["attempts"] >= MAX_ATTEMPTS:
            return False
        ctx = n["context"] if isinstance(n["context"], dict) else json.loads(n["context"] or "{}")
        subject, body = templates.render(n["type"], ctx, _freeze_url(conn, n["user_id"]))
        ok, reason, message_id = smtp.send(n["to_addr"], subject, body, headers={"Message-ID": smtp.new_message_id()})
        conn.execute(
            """update notification set attempts = attempts + 1, message_id_header = %s,
                   delivery_status = %s, error = %s, sent_at = case when %s then now() else sent_at end
               where id = %s""",
            (message_id, "sent" if ok else "failed", None if ok else reason[:500], ok, n["id"]),
        )
        return ok


def flush(limit: int = 50) -> int:
    """Send queued rows and retry failed ones (< MAX_ATTEMPTS). Returns how many were sent."""
    with connection() as conn:
        ids = [r["id"] for r in conn.execute(
            """select id from notification
               where delivery_status in ('queued', 'failed') and attempts < %s
               order by created_at limit %s""",
            (MAX_ATTEMPTS, limit),
        ).fetchall()]
    sent = 0
    for nid in ids:
        try:
            if _send_one(nid):
                sent += 1
        except Exception:  # a bad row must not stop the rest of the queue
            log.exception("notification %s could not be processed", nid)
    return sent
