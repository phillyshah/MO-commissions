"""The approval matrix (APR-1, ADM-1): which roles must sign, by amount and coding.

Rules are a JSON list; each rule is an amount band and/or a coding match plus the roles it adds and
the tier they sit in. The matrix is versioned under dual control and each approval records the
version it was judged under, so an auditor can replay why a chain looked the way it did.

No version is ever created by the system: `current()` falls back to a built-in default when nothing
has been approved, and the first real version is proposed and approved by two people in the admin UI.
A system-authored "approved" row would be a dual-control bypass with a tidy audit trail.

Rule shape:
    {"min_cents": 0, "max_cents": 2500000, "roles": ["sponsor"], "tier": 1}
    {"min_cents": 2500000, "max_cents": null, "roles": ["sponsor", "cfo"], "tier": 1}
    {"gl_prefix": "6", "cost_center": "1000", "roles": ["quality"], "tier": 1}
"""
from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from typing import Any
from uuid import UUID

import psycopg

from accounting.lifecycle import log_event
from accounting.routing._days import bd

DEFAULT_RULES: list[dict] = [
    {"min_cents": 0, "max_cents": 2_500_000, "roles": ["sponsor"], "tier": 1},
    {"min_cents": 2_500_000, "max_cents": None, "roles": ["sponsor", "cfo"], "tier": 1},
]

KNOWN_ROLES = {"sponsor", "cfo", "compliance", "quality", "threshold", "controller"}


class MatrixError(Exception):
    """A rule set or an approval that the matrix refuses; the message is written for the admin page."""


def current(conn: psycopg.Connection) -> tuple[int | None, list[dict]]:
    """The approved version in effect now, or (None, DEFAULT_RULES)."""
    row = conn.execute(
        """select id, rules from matrix_version
           where approved_at is not null and effective_from is not null and effective_from <= now()
           order by effective_from desc, id desc limit 1"""
    ).fetchone()
    if row is None:
        return None, [dict(r) for r in DEFAULT_RULES]
    return int(row["id"]), list(row["rules"])


def validate_rules(rules: Any) -> list[dict]:
    if not isinstance(rules, list) or not rules:
        raise MatrixError("The matrix needs at least one rule.")
    out = []
    for i, r in enumerate(rules, start=1):
        if not isinstance(r, dict):
            raise MatrixError(f"Rule {i} is not an object.")
        roles = r.get("roles")
        if not isinstance(roles, list) or not roles or not all(isinstance(x, str) and x in KNOWN_ROLES for x in roles):
            raise MatrixError(f"Rule {i}: roles must be a list from {sorted(KNOWN_ROLES)}.")
        lo, hi = r.get("min_cents"), r.get("max_cents")
        for name, v in (("min_cents", lo), ("max_cents", hi)):
            if v is not None and (isinstance(v, bool) or not isinstance(v, int) or v < 0):
                raise MatrixError(f"Rule {i}: {name} must be a whole number of cents or null.")
        if lo is not None and hi is not None and hi <= lo:
            raise MatrixError(f"Rule {i}: max_cents must be greater than min_cents.")
        tier = r.get("tier", 1)
        if isinstance(tier, bool) or not isinstance(tier, int) or tier < 1:
            raise MatrixError(f"Rule {i}: tier must be 1 or higher.")
        has_amount = lo is not None or hi is not None
        has_coding = bool(r.get("gl_prefix")) or bool(r.get("cost_center"))
        if not has_amount and not has_coding:
            raise MatrixError(f"Rule {i}: needs an amount band or a GL prefix / cost center.")
        out.append({
            "min_cents": lo, "max_cents": hi,
            "gl_prefix": r.get("gl_prefix") or None, "cost_center": r.get("cost_center") or None,
            "roles": list(roles), "tier": tier,
        })
    return out


def propose(conn: psycopg.Connection, rules: Any, by: UUID | str) -> int:
    clean = validate_rules(rules)
    row = conn.execute(
        "insert into matrix_version (rules, proposed_by) values (%s::jsonb, %s) returning id",
        (json.dumps(clean), by),
    ).fetchone()
    log_event(conn, event_type="config", actor_type="user", actor_id=by, reason_code="matrix_proposed",
              payload={"matrix_version_id": int(row["id"]), "rules": clean})
    return int(row["id"])


def approve(conn: psycopg.Connection, version_id: int, by: UUID | str) -> dict:
    """SOD-4: the proposer cannot approve. ADM-1: effective the next business day, so a change made at
    4 pm cannot alter the chain of an invoice somebody is signing at 4:05."""
    row = conn.execute("select * from matrix_version where id = %s for update", (version_id,)).fetchone()
    if row is None:
        raise MatrixError("That matrix version does not exist.")
    if row["approved_at"] is not None:
        raise MatrixError("That matrix version is already approved.")
    if str(row["proposed_by"]) == str(by):
        raise MatrixError("The person who proposed a matrix cannot also approve it. Ask a second administrator.")
    now = datetime.now(timezone.utc)
    next_day = bd().add_business_days(now, 1)
    effective = datetime(next_day.year, next_day.month, next_day.day, 0, 0, tzinfo=timezone.utc)
    updated = conn.execute(
        "update matrix_version set approved_by = %s, approved_at = %s, effective_from = %s where id = %s returning *",
        (by, now, effective, version_id),
    ).fetchone()
    log_event(conn, event_type="config", actor_type="user", actor_id=by, reason_code="matrix_approved",
              payload={"matrix_version_id": int(version_id), "effective_from": effective})
    return updated


def _amount_matches(rule: dict, gross_cents: int) -> bool:
    lo, hi = rule.get("min_cents"), rule.get("max_cents")
    if lo is None and hi is None:
        return True
    if lo is not None and gross_cents < lo:
        return False
    if hi is not None and gross_cents >= hi:
        return False
    return True


def _coding_matches(rule: dict, gl: str | None, cost_center: str | None) -> bool:
    if rule.get("gl_prefix") and not (gl or "").startswith(str(rule["gl_prefix"])):
        return False
    if rule.get("cost_center") and (cost_center or "") != str(rule["cost_center"]):
        return False
    return True


def required_roles(rules: list[dict], inv: dict) -> list[tuple[str, int]]:
    """Roles (with tier) an invoice needs. `inv` carries gross_cents, gl, cost_center, hcp and
    step_up_cents (the live config value, passed in so this stays a pure function).

    The sponsor is always required: APR-1 says every invoice routes to the vendor's sponsor and the
    matrix only ADDS signers, so a rule set that forgets the sponsor cannot remove them. Compliance is
    added for HCP vendors (VEN-6) and the CFO at or above the step-up amount (D3/D4) for the same reason:
    those are requirements of the spec, not of whoever last edited the matrix.
    """
    gross = int(inv.get("gross_cents") or 0)
    gl, cc = inv.get("gl"), inv.get("cost_center")
    out: list[tuple[str, int]] = []

    def add(role: str, tier: int) -> None:
        for r, t in out:
            if r == role:
                return  # first (lowest tier among rules in order) wins; one role signs once
        out.append((role, tier))

    add("sponsor", 1)
    for rule in rules:
        if _amount_matches(rule, gross) and _coding_matches(rule, gl, cc):
            for role in rule.get("roles", []):
                add(role, int(rule.get("tier", 1)))
    if inv.get("hcp"):
        add("compliance", 1)
    step_up = inv.get("step_up_cents")
    if step_up is not None and gross >= int(step_up):
        add("cfo", 1)
    out.sort(key=lambda rt: rt[1])
    return out
