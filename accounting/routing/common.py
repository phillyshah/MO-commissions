"""Helpers shared by routing, approvals and quarantine.

Everything here is about the OTHER packages' edges: the notification queue (Package C), the live
configuration values, and the people an invoice must never be routed to. They sit in one place so
the segregation-of-duties logic reads identically whether it runs at routing time or at signing
time — SOD-1 enforced twice with two different lists is a hole, not a belt and braces.
"""
from __future__ import annotations

import logging
from datetime import date
from typing import Any
from uuid import UUID

import psycopg

from accounting.config import settings

log = logging.getLogger("accounting.routing")


def notify_queue(conn: psycopg.Connection, **kw: Any) -> Any:
    """Queue a notification through Package C's `accounting.notify.send.queue`.

    Imported lazily because the module is being written in parallel; if it is missing the intent is
    logged and dropped rather than failing the approval — a lost email is recoverable (reminders and
    the digest repeat it), a lost signature is not.
    """
    try:
        from accounting.notify import send  # Package C
    except Exception:
        log.warning("notify.send unavailable; dropped %s", kw.get("type"))
        return None
    return send.queue(conn, **kw)


def step_up_cents(conn: psycopg.Connection) -> int:
    """D4: the step-up amount is a fixed, admin-configured figure under dual control.
    The approved value in config_current wins; the environment default is only the starting point."""
    row = conn.execute("select value from config_current where key = 'step_up_amount_cents'").fetchone()
    if row is None:
        return settings.default_step_up_amount_cents
    value = row["value"]
    try:
        return int(value if not isinstance(value, dict) else value.get("cents"))
    except (TypeError, ValueError):
        return settings.default_step_up_amount_cents


def config_bool(conn: psycopg.Connection, key: str, default: bool) -> bool:
    row = conn.execute("select value from config_current where key = %s", (key,)).fetchone()
    if row is None:
        return default
    v = row["value"]
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("1", "true", "yes", "on")


# ap_user columns with `roles` cast to text[]: psycopg has no adapter for the portal_role enum array and
# returns it as the raw string '{approver,cfo}', which `set()` would happily turn into characters.
USER_COLS = "u.id, u.subject, u.email, u.display_name, u.roles::text[] as roles, u.status, u.manager_id, u.backup_id, u.created_at"


def user_row(conn: psycopg.Connection, user_id: UUID | str | None) -> dict | None:
    if user_id is None:
        return None
    return conn.execute(f"select {USER_COLS} from ap_user u where u.id = %s", (user_id,)).fetchone()


def users_with_role(conn: psycopg.Connection, role: str) -> list[dict]:
    """Active users holding a portal role, least-loaded first so a role with several holders shares
    the work instead of always landing on the alphabetically first person."""
    return conn.execute(
        f"""select {USER_COLS}, (select count(*) from approval_request r where r.assignee_id = u.id and r.status = 'open') as open_count
           from ap_user u
           where %s = any(u.roles::text[]) and u.status = 'active'
           order by open_count, u.display_name""",
        (role,),
    ).fetchall()


def excluded_users(conn: psycopg.Connection, invoice: dict) -> set[UUID]:
    """People who may not sign this invoice.

    SOD-1: whoever captured, uploaded or corrected it. Corrections are rows in invoice_correction;
    uploads and manual captures are user-actor events on the invoice in event_log (Package A logs
    them; the invoice table itself has no uploader column). SOD-2: whoever edited this vendor's master
    data — Package A logs those as event_type 'vendor' with the vendor id in the payload.
    """
    ids: set[UUID] = set()
    for r in conn.execute("select user_id from invoice_correction where invoice_id = %s", (invoice["id"],)):
        ids.add(r["user_id"])
    for r in conn.execute(
        """select actor_id from event_log
           where invoice_id = %s and actor_type = 'user' and actor_id is not null
             and (event_type in ('upload', 'capture', 'correction') or reason_code in ('manual_upload', 'clerk_edit', 'clerk_confirmed'))""",
        (invoice["id"],),
    ):
        ids.add(r["actor_id"])
    if invoice.get("vendor_id"):
        for r in conn.execute(
            """select actor_id from event_log
               where actor_type = 'user' and actor_id is not null and event_type in ('vendor', 'vendor_edit')
                 and payload->>'vendor_id' = %s""",
            (str(invoice["vendor_id"]),),
        ):
            ids.add(r["actor_id"])
    return ids


def active_delegation(conn: psycopg.Connection, user_id: UUID | str, on: date | None = None) -> dict | None:
    """The delegation row that redirects `user_id`'s work today, if any (APR-8). Only direct
    delegations count — a delegate cannot sub-delegate, so the chain is never followed further."""
    on = on or date.today()
    return conn.execute(
        """select * from delegation where from_user_id = %s and revoked_at is null and starts_on <= %s and ends_on >= %s
           order by created_at desc limit 1""",
        (user_id, on, on),
    ).fetchone()


def is_delegate_for(conn: psycopg.Connection, delegate_id: UUID | str, principal_id: UUID | str, on: date | None = None) -> bool:
    d = active_delegation(conn, principal_id, on)
    return bool(d and d["to_user_id"] == delegate_id)


def load_invoice(conn: psycopg.Connection, invoice_id: UUID | str, *, lock: bool = False) -> dict | None:
    """Invoice row plus the vendor columns routing needs (prefixed vendor_*)."""
    row = conn.execute(
        f"""select i.*,
                   v.legal_name as vendor_legal_name, v.sap_vendor_number as vendor_sap_number, v.hcp as vendor_hcp,
                   v.quality_critical as vendor_quality_critical, v.asl_status as vendor_asl_status,
                   v.default_sponsor_id as vendor_default_sponsor_id, v.backup_sponsor_id as vendor_backup_sponsor_id,
                   v.default_gl as vendor_default_gl, v.default_cost_center as vendor_default_cost_center,
                   v.recurring_schedule as vendor_recurring_schedule, v.recurring_tolerance_pct as vendor_recurring_tolerance_pct,
                   v.status as vendor_status, v.created_at as vendor_created_at, v.po_required as vendor_po_required,
                   v.phone as vendor_phone, v.phone_since as vendor_phone_since
            from invoice i left join vendor v on v.id = i.vendor_id
            where i.id = %s {'for update of i' if lock else ''}""",
        (invoice_id,),
    ).fetchone()
    return row


def portal_number(n: int | None) -> str:
    return f"AP-{int(n):06d}" if n is not None else "AP-?"
