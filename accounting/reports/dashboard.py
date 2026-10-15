"""Home tiles (spec §9) and the sidebar counts.

Why the home page is role-shaped: a 55+ approver should see one number — how many things wait on
them — not a map of the whole system. The clerk sees the three queues they work; the Controller
sees quarantine and whether yesterday's mail arrived. Everything else is a tile to a sibling app.

"As of" is the server clock here: there is no published snapshot in this portal, the numbers are
live, and saying when they were read is still owed to the reader.
"""
from __future__ import annotations

import logging
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from accounting.business_days import add_business_days
from accounting.config import settings
from accounting.db.pool import connection

log = logging.getLogger("accounting.reports.dashboard")

AP_ROLES = ("clerk", "supervisor", "controller", "admin", "auditor")
APPROVER_ROLES = ("approver", "cfo", "compliance")


def _count(conn, sql: str, params=()) -> int:
    return int(conn.execute(sql, params).fetchone()["n"])


def break_glass_on(conn) -> bool:
    row = conn.execute("select value from config_current where key = 'break_glass'").fetchone()
    return bool(row and row["value"] is True)


def nav_counts(user) -> dict:
    """Badges for the sidebar plus the break-glass flag; one query each, nothing heavy."""
    out = {"approvals": 0, "quarantine": 0, "break_glass": False}
    if user is None:
        return out
    with connection() as conn:
        if user.has(*APPROVER_ROLES, "controller", "supervisor", "admin"):
            out["approvals"] = _count(conn, "select count(*) as n from approval_request where assignee_id = %s and status = 'open'", (user.id,))
        if user.has("controller", "cfo", "admin"):
            out["quarantine"] = _count(conn, "select count(*) as n from quarantine_case where state = 'open'")
        out["break_glass"] = break_glass_on(conn)
    return out


def nav_counts_for(user) -> dict:
    """The template-global version: never raises, because a sidebar badge must never take the page
    down with it."""
    try:
        return nav_counts(user)
    except Exception:
        log.exception("nav counts unavailable")
        return {"approvals": 0, "quarantine": 0, "break_glass": False}


def greeting_for(now: datetime) -> str:
    hour = now.astimezone(ZoneInfo(settings.timezone)).hour
    return "morning" if hour < 12 else "afternoon" if hour < 17 else "evening"


def _heartbeat_banner(conn, now: datetime) -> dict | None:
    yesterday = (now.astimezone(ZoneInfo(settings.timezone)) - timedelta(days=1)).date()
    row = conn.execute(
        """select coalesce(sum(n), 0) as n from forwarded_volume_by_day
           where trust = 'forwarded_by_maxx' and (day at time zone %s)::date = %s""",
        (settings.timezone, yesterday),
    ).fetchone()
    n = int(row["n"] or 0)
    if yesterday.weekday() >= 5:
        return None
    if n < settings.heartbeat_min_per_day or n > settings.heartbeat_max_per_day:
        return {"kind": "warn", "text": f"Mail heartbeat: {n} forwarded message(s) arrived yesterday; the expected range is "
                                        f"{settings.heartbeat_min_per_day} to {settings.heartbeat_max_per_day}. Check the Exchange redirect rule."}
    return None


def home_tiles(user) -> dict:
    now = datetime.now(timezone.utc)
    work: list[dict] = []
    banners: list[dict] = []
    with connection() as conn:
        if break_glass_on(conn):
            banners.append({"kind": "bad", "text": "BREAK-GLASS MODE IS ON. Normal controls are relaxed; every action is being recorded for review."})
        if user.has(*APPROVER_ROLES):
            n = _count(conn, "select count(*) as n from approval_request where assignee_id = %s and status = 'open'", (user.id,))
            two_days_ago = add_business_days(now, -2)
            old = _count(conn, "select count(*) as n from approval_request where assignee_id = %s and status = 'open' and created_at < %s",
                         (user.id, two_days_ago))
            work.append({"title": "Waiting for your approval", "count": n, "tone": "warn" if old else "",
                         "text": (f"{old} older than two business days" if old else "Oldest first when you open the list"),
                         "href": "/approvals"})
        if user.has("clerk", "supervisor", "controller", "admin"):
            work.append({"title": "Needs AP review", "count": _count(conn, "select count(*) as n from invoice where state = 'AP_REVIEW'"),
                         "tone": "", "text": "Fields to confirm or correct", "href": "/invoices?state=AP_REVIEW"})
            work.append({"title": "Unknown vendors", "count": _count(conn, "select count(*) as n from invoice where state = 'VENDOR_SETUP'"),
                         "tone": "", "text": "Invoices from vendors not set up yet", "href": "/invoices?state=VENDOR_SETUP"})
            work.append({"title": "Approved, waiting for SAP entry", "count": _count(conn, "select count(*) as n from invoice where state = 'APPROVED'"),
                         "tone": "", "text": "Ready to key into SAP", "href": "/invoices/approved"})
        if user.has("controller", "cfo", "admin"):
            q = _count(conn, "select count(*) as n from quarantine_case where state = 'open'")
            work.append({"title": "Quarantine", "count": q, "tone": "bad" if q else "", "text": "Messages held for your decision", "href": "/quarantine"})
            hb = _heartbeat_banner(conn, now)
            if hb:
                banners.append(hb)
    systems: list[dict] = []
    if user.has(*AP_ROLES):
        systems.append({"title": "Invoices (AP)", "text": "Intake, review, approvals and the approved queue", "href": "/invoices", "external": False})
    systems.append({"title": "Commission statements", "text": "The US Usage Dashboard", "href": settings.maxxdash_dashboard_url, "external": True})
    systems.append({"title": "Usage tickets", "text": "Usage ticket capture", "href": settings.usage_url, "external": True})
    if user.has("controller", "supervisor", "clerk", "admin"):
        systems.append({"title": "Statement generator", "text": "Commission statement generator", "href": settings.commissions_generator_url, "external": True})
    return {"greeting": greeting_for(now), "as_of": now, "banners": banners, "work": work, "systems": systems}
