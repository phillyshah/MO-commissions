"""Reminders and digests (REM-1 to REM-3, NOT-2, ID-5).

The escalation ladder itself (D+2 remind, D+3 manager, D+4 Controller) lives in
`accounting.routing.reassign.escalate_due`, which owns the approval_request clock. This module
turns what it returns — notification intents — into queued emails and stamps the request so the
same step is never sent twice. The import is lazy because the routing package is built alongside
this one; a missing module logs and sends nothing rather than crashing the scheduler.
"""
from __future__ import annotations

import logging
from datetime import datetime, timedelta, timezone

import psycopg

from accounting.business_days import business_days_between
from accounting.config import settings
from accounting.db.pool import connection
from accounting.notify.send import queue

log = logging.getLogger("accounting.notify.reminders")

# intent type -> approval_request column stamped when the notification is queued
STAMP = {
    "reminder": "reminded_at",
    "manager_copy": "escalated_manager_at",
    "controller_reassign": "escalated_controller_at",
}


def _escalate_due(conn: psycopg.Connection, now: datetime) -> list[dict]:
    try:
        from accounting.routing.reassign import escalate_due
    except ImportError:
        log.warning("routing.reassign is not available; no reminders this hour")
        return []
    return list(escalate_due(conn, now) or [])


def run(now: datetime | None = None) -> int:
    """Queue one notification per intent and stamp the request. Returns how many were queued."""
    now = now or datetime.now(timezone.utc)
    queued = 0
    with connection() as conn:
        for intent in _escalate_due(conn, now):
            type_ = intent.get("type") or "reminder"
            user_id = intent.get("user_id")
            invoice_id = intent.get("invoice_id")
            nid = queue(conn, type=type_, user_id=user_id, invoice_id=invoice_id,
                        context={"count": 1, "reason": intent.get("reason")})
            if nid is not None:
                queued += 1
            col = STAMP.get(type_)
            if col and invoice_id:
                if intent.get("request_id"):
                    conn.execute(f"update approval_request set {col} = %s where id = %s and {col} is null",
                                 (now, intent["request_id"]))
                else:
                    conn.execute(
                        f"update approval_request set {col} = %s where invoice_id = %s and status = 'open' and {col} is null",
                        (now, invoice_id))
    return queued


def _digest_context(conn: psycopg.Connection, user_id, now: datetime) -> dict | None:
    rows = conn.execute(
        """select r.created_at, r.due_at, v.legal_name, coalesce(v.hcp, false) or 'hcp' = any(i.tags) as hcp
           from approval_request r join invoice i on i.id = r.invoice_id
           left join vendor v on v.id = i.vendor_id
           where r.assignee_id = %s and r.status = 'open'""",
        (user_id,),
    ).fetchall()
    if not rows:
        return None
    oldest = min(r["created_at"] for r in rows)
    by_vendor: dict[str, int] = {}
    hcp = 0
    for r in rows:
        if r["hcp"]:
            hcp += 1
        if r["due_at"] < now:
            name = r["legal_name"] or "Unknown vendor"
            by_vendor[name] = by_vendor.get(name, 0) + 1
    return {
        "count": len(rows),
        "oldest_business_days": business_days_between(oldest, now),
        "delinquent_by_vendor": [{"vendor": k, "count": v} for k, v in sorted(by_vendor.items())],
        "hcp_count": hcp,
    }


def weekly_digest(now: datetime | None = None) -> int:
    """REM-3: one email per user who has anything open. Returns how many were queued."""
    now = now or datetime.now(timezone.utc)
    queued = 0
    with connection() as conn:
        users = conn.execute(
            "select distinct assignee_id as id from approval_request where status = 'open'").fetchall()
        for u in users:
            ctx = _digest_context(conn, u["id"], now)
            if ctx and queue(conn, type="digest", user_id=u["id"], context=ctx) is not None:
                queued += 1
    return queued


def _heartbeat_text() -> str:
    try:
        from accounting.ingest.heartbeat import check_daily
        result = check_daily()
        return result.get("text") or ("in range" if result.get("ok") else "out of range")
    except Exception as exc:  # the ingest package may not be present yet
        log.warning("heartbeat check unavailable: %s", exc)
        return "not checked (heartbeat unavailable)"


def controller_daily(now: datetime | None = None) -> int:
    """ID-5 / NOT-5: the Controller's one daily email — quarantine, escalations, every grant,
    removal and disable, freezes, and the mail heartbeat. Returns how many were queued."""
    now = now or datetime.now(timezone.utc)
    since = now - timedelta(hours=24)
    with connection() as conn:
        q_open = conn.execute("select count(*) as n from quarantine_case where state = 'open'").fetchone()["n"]
        esc = conn.execute("select count(*) as n from approval_request where escalated_controller_at >= %s",
                           (since,)).fetchone()["n"]
        freezes = conn.execute("select count(*) as n from security_event where type = 'freeze' and at >= %s",
                               (since,)).fetchone()["n"]
        changes = conn.execute(
            """select e.at, e.event_type, e.reason_code, e.payload, u.email as actor
               from event_log e left join ap_user u on u.id = e.actor_id
               where e.event_type in ('access', 'admin') and e.at >= %s
                 and coalesce(e.reason_code, '') not in ('sign_in', 'sign_out')
               order by e.at""",
            (since,),
        ).fetchall()
        lines = []
        for c in changes:
            who = (c["payload"] or {}).get("user_email") or (c["payload"] or {}).get("email") or ""
            lines.append(f"{c['reason_code'] or c['event_type']} {who} by {c['actor'] or 'system'}".strip())
        ctx = {
            "quarantine_open": int(q_open), "escalations": int(esc), "freezes": int(freezes),
            "access_changes": lines, "heartbeat": _heartbeat_text(),
        }
        nid = queue(conn, type="controller_daily", user_id=None, to=None, context=ctx)
        if nid is None and settings.controller_email:
            nid = queue(conn, type="controller_daily", user_id=None, to=settings.controller_email, context=ctx)
        return 1 if nid is not None else 0
