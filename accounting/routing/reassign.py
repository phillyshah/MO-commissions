"""Reassignment and the reminder clock (APR-2 Reassign, APR-7 leavers, REM-1).

An approval_request is never edited in place to point at somebody else: the old row is closed with
status 'reassigned' and a new row carries the work, keeping the original due date. The due date is
kept on purpose — sponsor aging never pauses (two clocks), and handing an invoice to a delegate must
not buy it two more days.

`escalate_due()` only decides and stamps; Package C's reminders job sends. Each step has its own
timestamp column so a step fires exactly once however often the job runs.
"""
from __future__ import annotations

from datetime import datetime
from uuid import UUID

import psycopg

from accounting.lifecycle import log_event, transition
from accounting.routing._days import bd, utcnow
from accounting.routing.common import excluded_users, load_invoice, notify_queue, user_row, users_with_role

REMIND_AFTER_BD = 2
MANAGER_AFTER_BD = 3
CONTROLLER_AFTER_BD = 4


class ReassignError(Exception):
    """Refused reassignment; message written for the approver's screen."""


def reassign(conn: psycopg.Connection, request_id: UUID | str, to_user: UUID | str, by: UUID | str, reason: str) -> UUID:
    req = conn.execute("select * from approval_request where id = %s for update", (request_id,)).fetchone()
    if req is None or req["status"] != "open":
        raise ReassignError("That approval is no longer open.")
    if str(to_user) == str(by):
        raise ReassignError("You cannot reassign an approval to yourself.")
    if str(to_user) == str(req["assignee_id"]):
        raise ReassignError("That person already has this approval.")
    target = user_row(conn, to_user)
    if target is None or target["status"] != "active":
        raise ReassignError("That person cannot take approvals right now.")
    inv = load_invoice(conn, req["invoice_id"])
    if target["id"] in excluded_users(conn, inv):
        raise ReassignError("That person handled this invoice in AP and cannot approve it.")
    if not target["roles"] or not set(target["roles"]) & {"approver", "cfo", "compliance", "controller", "supervisor", "admin"}:
        raise ReassignError("That person is not set up as an approver.")
    conn.execute(
        """update approval_request set status = 'reassigned', closed_at = now(), reassigned_to = %s, reassigned_by = %s,
                  reassign_reason = %s where id = %s""",
        (to_user, by, reason, request_id),
    )
    new = conn.execute(
        """insert into approval_request (invoice_id, matrix_version_id, assignee_id, role, tier, required_amount_cents, due_at,
                                         reminded_at, escalated_manager_at, escalated_controller_at)
           values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) returning id""",
        (req["invoice_id"], req["matrix_version_id"], to_user, req["role"], req["tier"], req["required_amount_cents"],
         req["due_at"], req["reminded_at"], req["escalated_manager_at"], req["escalated_controller_at"]),
    ).fetchone()
    log_event(conn, event_type="routing", actor_type="user", actor_id=by, invoice_id=req["invoice_id"], reason_code="reassigned",
              payload={"from_request_id": str(request_id), "to_request_id": str(new["id"]), "from_user": str(req["assignee_id"]),
                       "to_user": str(to_user), "reason": reason})
    notify_queue(conn, type="pending", user_id=to_user, invoice_id=req["invoice_id"])
    return new["id"]


def escalate_due(conn: psycopg.Connection, now: datetime | None = None) -> list[dict]:
    """Walk open requests and return the notification intents due now: D+2 reminder to the approver,
    D+3 copy to their manager, D+4 'controller_reassign' to every Controller. Stamps as it goes."""
    now = now or utcnow()
    days = bd()
    intents: list[dict] = []
    rows = conn.execute(
        """select r.*, u.manager_id from approval_request r join ap_user u on u.id = r.assignee_id
           where r.status = 'open' order by r.created_at"""
    ).fetchall()
    controllers = None
    for r in rows:
        age = days.business_days_between(r["created_at"], now)
        if age >= REMIND_AFTER_BD and r["reminded_at"] is None:
            intents.append({"type": "reminder", "user_id": r["assignee_id"], "invoice_id": r["invoice_id"], "request_id": r["id"]})
            conn.execute("update approval_request set reminded_at = %s where id = %s", (now, r["id"]))
        if age >= MANAGER_AFTER_BD and r["escalated_manager_at"] is None:
            if r["manager_id"]:
                intents.append({"type": "manager_copy", "user_id": r["manager_id"], "invoice_id": r["invoice_id"], "request_id": r["id"]})
            conn.execute("update approval_request set escalated_manager_at = %s where id = %s", (now, r["id"]))
        if age >= CONTROLLER_AFTER_BD and r["escalated_controller_at"] is None:
            if controllers is None:
                controllers = users_with_role(conn, "controller")
            for c in controllers:
                intents.append({"type": "controller_reassign", "user_id": c["id"], "invoice_id": r["invoice_id"], "request_id": r["id"]})
            conn.execute("update approval_request set escalated_controller_at = %s where id = %s", (now, r["id"]))
            log_event(conn, event_type="routing", actor_type="system", invoice_id=r["invoice_id"], reason_code="escalated_controller",
                      payload={"request_id": str(r["id"]), "assignee_id": str(r["assignee_id"]), "business_days": age})
    return intents


def reassign_open_items_for(conn: psycopg.Connection, user_id: UUID | str) -> list[dict]:
    """APR-7 leaver sweep. Every open request held by the user goes to the vendor's backup sponsor, or
    to another holder of the role; when nobody can take it the invoice goes back to AP with
    'reassign_exhausted' so it is visible rather than silently stuck on a departed person.
    Also ends the user's delegations in both directions — a leaver cannot keep acting for others, and
    work delegated TO them would otherwise land on an empty desk."""
    conn.execute("update delegation set revoked_at = now() where revoked_at is null and (from_user_id = %s or to_user_id = %s)",
                 (user_id, user_id))
    results: list[dict] = []
    for req in conn.execute("select * from approval_request where assignee_id = %s and status = 'open' for update", (user_id,)).fetchall():
        inv = load_invoice(conn, req["invoice_id"])
        excluded = excluded_users(conn, inv) | {req["assignee_id"]}
        target = None
        roles = req["role"].split(",")
        if "sponsor" in roles:
            for uid in (inv.get("vendor_backup_sponsor_id"), inv.get("vendor_default_sponsor_id")):
                u = user_row(conn, uid)
                if u and u["status"] == "active" and u["id"] not in excluded:
                    target = u
                    break
        if target is None:
            for role in roles:
                for portal_role in {"cfo": ["cfo"], "compliance": ["compliance"], "quality": ["compliance"]}.get(role, ["approver"]):
                    for u in users_with_role(conn, portal_role):
                        if u["id"] not in excluded:
                            target = u
                            break
                    if target:
                        break
                if target:
                    break
        if target is None:
            conn.execute("update approval_request set status = 'voided', closed_at = now(), reassign_reason = 'leaver_no_backup' where id = %s",
                         (req["id"],))
            if inv["state"] == "PENDING_APPROVAL":
                transition(conn, req["invoice_id"], "AP_REVIEW", reason_code="reassign_exhausted", actor_type="system",
                           payload={"request_id": str(req["id"]), "left_user": str(user_id)})
            results.append({"request_id": req["id"], "to": None})
            continue
        conn.execute(
            """update approval_request set status = 'reassigned', closed_at = now(), reassigned_to = %s, reassign_reason = 'leaver'
               where id = %s""", (target["id"], req["id"]))
        new = conn.execute(
            """insert into approval_request (invoice_id, matrix_version_id, assignee_id, role, tier, required_amount_cents, due_at)
               values (%s,%s,%s,%s,%s,%s,%s) returning id""",
            (req["invoice_id"], req["matrix_version_id"], target["id"], req["role"], req["tier"], req["required_amount_cents"], req["due_at"]),
        ).fetchone()
        log_event(conn, event_type="routing", actor_type="system", invoice_id=req["invoice_id"], reason_code="leaver_reassigned",
                  payload={"from_request_id": str(req["id"]), "to_request_id": str(new["id"]), "from_user": str(user_id), "to_user": str(target["id"])})
        notify_queue(conn, type="pending", user_id=target["id"], invoice_id=req["invoice_id"])
        results.append({"request_id": req["id"], "to": target["id"], "new_request_id": new["id"]})
    return results
