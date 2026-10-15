"""Build the approval chain for a CAPTURED invoice (APR-1, APR-5, APR-6, APR-8, SOD-1/2/6).

Order of decisions, each one logged:
  1. hard stops that send the invoice back to AP (ASL / quality-critical without a PO),
  2. auto-release for recurring invoices (APR-4),
  3. otherwise the chain: matrix roles -> people -> delegation -> segregation checks -> request rows.

Why people are resolved here and not at signing time: an approver must see their queue the moment the
invoice is routed, and the Controller must be able to read WHO was asked and why from the event log.
Signing re-checks the segregation rules anyway (sign.py), so a change in between is caught twice.
"""
from __future__ import annotations

import logging
from datetime import timedelta
from uuid import UUID

import psycopg

from accounting.lifecycle import InvalidTransition, log_event, transition
from accounting.routing import auto_release, matrix
from accounting.routing._days import bd, utcnow
from accounting.routing.common import (active_delegation, excluded_users, load_invoice, notify_queue,
                                       step_up_cents, user_row, users_with_role)

log = logging.getLogger("accounting.routing")

# Matrix role -> portal roles whose holders may fill it. 'quality' maps to Compliance because the
# portal has no Quality role in Stage 1 and the roles table gives Compliance the ASL/quality items.
ROLE_HOLDERS = {
    "cfo": ["cfo"],
    "compliance": ["compliance"],
    "quality": ["compliance"],
    "threshold": ["approver", "cfo"],
    "controller": ["controller"],
}

APPROVER_DUE_BUSINESS_DAYS = 2  # REM-1: D+2 is the first reminder, so that is the approver's due date


class Candidate:
    __slots__ = ("user_id", "via")

    def __init__(self, user_id: UUID, via: str):
        self.user_id = user_id
        self.via = via


def _usable(conn: psycopg.Connection, user_id: UUID | None, excluded: set[UUID]) -> Candidate | None:
    """A person who can take a request: exists, active, not excluded by SOD; redirected to an active
    delegate when they have one (and the delegate passes the same checks). A delegate who captured the
    invoice is excluded like anyone else — APR-8 says they cannot approve what they handled."""
    if user_id is None or user_id in excluded:
        return None
    row = user_row(conn, user_id)
    if row is None or row["status"] != "active":
        return None
    d = active_delegation(conn, user_id)
    if d is not None:
        delegate = user_row(conn, d["to_user_id"])
        if delegate is None or delegate["status"] != "active" or delegate["id"] in excluded:
            return None
        return Candidate(delegate["id"], f"delegate_of:{user_id}")
    return Candidate(row["id"], "direct")


def _resolve(conn: psycopg.Connection, role: str, inv: dict, excluded: set[UUID], taken: set[UUID]) -> Candidate | None:
    if role == "sponsor":
        for uid, via in ((inv.get("vendor_default_sponsor_id"), "sponsor"), (inv.get("vendor_backup_sponsor_id"), "backup_sponsor")):
            c = _usable(conn, uid, excluded)
            if c is not None:
                c.via = via if c.via == "direct" else f"{via}/{c.via}"
                return c
        return None
    # Prefer a holder who is not already on this chain (a second pair of eyes is the point of a
    # co-signer); when the only holder IS already on it, APR-6 applies and one signature covers both
    # roles, so fall back to them rather than sending the invoice back to AP.
    fallback: Candidate | None = None
    for portal_role in ROLE_HOLDERS.get(role, [role]):
        for row in users_with_role(conn, portal_role):
            c = _usable(conn, row["id"], excluded)
            if c is None:
                continue
            if c.user_id not in taken:
                return c
            fallback = fallback or c
    return fallback


def route(conn: psycopg.Connection, invoice_id: UUID | str) -> str:
    inv = load_invoice(conn, invoice_id, lock=True)
    if inv is None:
        raise InvalidTransition(f"invoice {invoice_id} not found")
    if inv["state"] != "CAPTURED":
        return inv["state"]
    if not inv.get("vendor_id"):
        return transition(conn, invoice_id, "VENDOR_SETUP", reason_code="vendor_unknown", actor_type="system")["state"]

    # 1. Hard stops (VEN-5, QMSR hook): never past AP.
    if inv.get("vendor_quality_critical") and (inv.get("vendor_asl_status") or "n/a") not in ("approved",):
        return transition(conn, invoice_id, "AP_REVIEW", reason_code="asl_hard_stop", actor_type="system",
                          payload={"asl_status": inv.get("vendor_asl_status")})["state"]
    if inv.get("vendor_quality_critical") and not inv.get("po_number"):
        return transition(conn, invoice_id, "AP_REVIEW", reason_code="quality_no_po", actor_type="system")["state"]

    # 2. Auto-release.
    ok, detail = auto_release.eligible(conn, inv)
    if ok:
        conn.execute("update invoice set tags = array_append(tags, 'auto_released') where id = %s and not ('auto_released' = any(tags))",
                     (invoice_id,))
        row = transition(conn, invoice_id, "APPROVED", reason_code="auto_release_recurring", actor_type="system", payload=detail)
        notify_queue(conn, type="approved_for_ap", user_id=None, invoice_id=invoice_id)
        return row["state"]
    log_event(conn, event_type="routing", actor_type="system", invoice_id=invoice_id, reason_code="auto_release_declined", payload=detail)

    # 3. The chain.
    version_id, rules = matrix.current(conn)
    step_up = step_up_cents(conn)
    roles = matrix.required_roles(rules, {
        "gross_cents": inv.get("gross_cents"), "gl": inv.get("vendor_default_gl"),
        "cost_center": inv.get("vendor_default_cost_center"), "hcp": inv.get("vendor_hcp"), "step_up_cents": step_up,
    })
    excluded = excluded_users(conn, inv)

    # person -> {tier, roles[], via}; APR-6: one person asked twice on identical amount/basis signs once.
    slots: dict[UUID, dict] = {}
    taken: set[UUID] = set()
    unfilled: list[str] = []
    for role, tier in roles:
        c = _resolve(conn, role, inv, excluded, taken)
        if c is None:
            unfilled.append(role)
            continue
        taken.add(c.user_id)
        slot = slots.setdefault(c.user_id, {"tier": tier, "roles": [], "via": c.via})
        slot["roles"].append(role)
        slot["tier"] = min(slot["tier"], tier)

    now = utcnow()
    if unfilled or not slots:
        # Nobody eligible for a required role. The lifecycle table only reaches AP_REVIEW with
        # 'reassign_exhausted' from PENDING_APPROVAL, so the invoice is routed (recorded as such) and
        # immediately handed back to AP with the reason — two honest events rather than one wrong code.
        transition(conn, invoice_id, "PENDING_APPROVAL", reason_code="routed", actor_type="system",
                   payload={"matrix_version_id": version_id, "chain": [], "unfilled_roles": unfilled,
                            "excluded": [str(u) for u in excluded]})
        return transition(conn, invoice_id, "AP_REVIEW", reason_code="reassign_exhausted", actor_type="system",
                          payload={"unfilled_roles": unfilled, "excluded": [str(u) for u in excluded]})["state"]

    # SIG-3 baseline: the hash every signer on this chain is compared against is the content AS ROUTED.
    # Overwritten (not merely set-if-null) because a re-route after an AP edit is new content.
    from accounting.approvals.payload import compute_content_hash
    conn.execute("update invoice set content_hash = %s where id = %s", (compute_content_hash(inv), invoice_id))

    due = bd().add_business_days(now, APPROVER_DUE_BUSINESS_DAYS)
    chain = []
    for user_id, slot in slots.items():
        r = conn.execute(
            """insert into approval_request (invoice_id, matrix_version_id, assignee_id, role, tier, required_amount_cents, due_at)
               values (%s, %s, %s, %s, %s, %s, %s) returning id""",
            (invoice_id, version_id, user_id, ",".join(slot["roles"]), slot["tier"], inv.get("gross_cents"), due),
        ).fetchone()
        chain.append({"request_id": str(r["id"]), "assignee_id": str(user_id), "role": ",".join(slot["roles"]),
                      "tier": slot["tier"], "via": slot["via"]})
    row = transition(conn, invoice_id, "PENDING_APPROVAL", reason_code="routed", actor_type="system",
                     payload={"matrix_version_id": version_id, "step_up_cents": str(step_up), "chain": chain,
                              "excluded": [str(u) for u in excluded]})
    for c in chain:
        notify_queue(conn, type="pending", user_id=c["assignee_id"], invoice_id=invoice_id)
    return row["state"]
