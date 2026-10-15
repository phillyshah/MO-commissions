"""/admin — users, configuration under dual control, approval matrix versions, delegations,
security events, the daily grants report and break-glass (ADM-1, ID-4, ID-5, NOT-5).

Dual control, as implemented: a config or matrix change is PROPOSED by one person and APPROVED by a
different one, and takes effect the next business day. The one deliberate exception: while the
`dual_control_enabled` setting is off or has never been set — the state a brand-new portal is in —
an administrator's own proposal is approved on the spot, so the portal can be configured at all.
That auto-approval is logged like any other change.

Every change writes an `admin` event with a before/after diff: the daily grants report (ID-5) is
nothing more than a reading of those rows.
"""
from __future__ import annotations

import hashlib
import json
import logging
import secrets
from datetime import datetime, timedelta, timezone

from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import RedirectResponse

from accounting.business_days import add_business_days
from accounting.db.pool import connection
from accounting.lifecycle import log_event
from accounting.web.common import as_uuid, page
from accounting.web.deps import require, require_active

log = logging.getLogger("accounting.web.admin")
router = APIRouter(prefix="/admin")

ADMINS = require("admin", "controller")
BREAK_GLASS = require("admin", "cfo")
ROLES = ["clerk", "supervisor", "controller", "approver", "cfo", "compliance", "admin", "auditor"]
CONFIG_KEYS = ["step_up_amount_cents", "dual_control_enabled", "soft_signals_enabled", "heartbeat_min_per_day", "heartbeat_max_per_day"]
ENROLLMENT_CODE_HOURS = 24


def _dual_control_on(conn) -> bool:
    row = conn.execute("select value from config_current where key = 'dual_control_enabled'").fetchone()
    return bool(row and row["value"] is True)


def _admin_event(conn, user, reason: str, diff: dict, target_user_id=None) -> int:
    payload = {"diff": diff}
    if target_user_id is not None:
        payload["user_id"] = str(target_user_id)
        row = conn.execute("select email from ap_user where id = %s", (target_user_id,)).fetchone()
        if row:
            payload["user_email"] = row["email"]
    return log_event(conn, event_type="admin", actor_type="user", actor_id=user.id, reason_code=reason, payload=payload)


def roles_list(value) -> list[str]:
    """`portal_role[]` comes back from psycopg as a list when the enum is registered and as the text
    '{clerk,admin}' when it is not. Accept both so a role check never compares characters."""
    if value is None:
        return []
    if isinstance(value, str):
        return [p.strip().strip('"') for p in value.strip("{}").split(",") if p.strip()]
    return [str(v) for v in value]


def _users(conn) -> list[dict]:
    rows = conn.execute(
        """select u.*, m.display_name as manager_name, b.display_name as backup_name,
                  (select count(*) from passkey p where p.user_id = u.id and p.status = 'active') as passkeys
           from ap_user u left join ap_user m on m.id = u.manager_id left join ap_user b on b.id = u.backup_id
           order by u.display_name""").fetchall()
    for r in rows:
        r["roles"] = roles_list(r["roles"])
    return rows


def _parse_value(raw: str):
    """Config values are JSON. Plain words are accepted too: 'true', '2500000', 'abc' → JSON string."""
    raw = (raw or "").strip()
    if raw == "":
        raise ValueError("Enter a value.")
    try:
        return json.loads(raw)
    except ValueError:
        return raw


def validate_rules(rules) -> list[dict]:
    """The approval-matrix rules shape: a list of {min_cents, max_cents|null, roles:[...], sequential, gl_prefix?, cost_center?}.
    Raises ValueError with a sentence."""
    if not isinstance(rules, list) or not rules:
        raise ValueError("Rules must be a list with at least one rule.")
    out = []
    for i, r in enumerate(rules, 1):
        if not isinstance(r, dict):
            raise ValueError(f"Rule {i} must be an object.")
        mn, mx = r.get("min_cents", 0), r.get("max_cents")
        if not isinstance(mn, int) or isinstance(mn, bool) or mn < 0:
            raise ValueError(f"Rule {i}: min_cents must be a whole number of cents.")
        if mx is not None and (not isinstance(mx, int) or isinstance(mx, bool) or mx < mn):
            raise ValueError(f"Rule {i}: max_cents must be a whole number of cents not below min_cents, or null.")
        roles = r.get("roles")
        if not isinstance(roles, list) or not roles or not all(isinstance(x, str) and x for x in roles):
            raise ValueError(f"Rule {i}: roles must be a non-empty list of role names.")
        if not isinstance(r.get("sequential", False), bool):
            raise ValueError(f"Rule {i}: sequential must be true or false.")
        for opt in ("gl_prefix", "cost_center"):
            if r.get(opt) is not None and not isinstance(r[opt], str):
                raise ValueError(f"Rule {i}: {opt} must be text.")
        out.append({"min_cents": mn, "max_cents": mx, "roles": roles, "sequential": r.get("sequential", False),
                    "gl_prefix": r.get("gl_prefix"), "cost_center": r.get("cost_center")})
    return out


# --- users ---------------------------------------------------------------------------------------------

@router.get("")
def users(request: Request, user=Depends(ADMINS)):
    return _users_page(request, user)


def _users_page(request, user, error=None, new_code=None, code_for=None, status=200):
    with connection() as conn:
        rows = _users(conn)
        bg = conn.execute("select value from config_current where key = 'break_glass'").fetchone()
    return page(request, user, "admin/users.html", {"users": rows, "roles": ROLES, "error": error, "new_code": new_code, "code_for": code_for,
                                                    "break_glass": bool(bg and bg["value"] is True), "can_break_glass": user.has("admin", "cfo")}, status)


@router.post("/users/{user_id}")
def update_user(request: Request, user_id: str, roles: list[str] = Form([]), manager_id: str = Form(""), backup_id: str = Form(""),
                user=Depends(ADMINS), _active=Depends(require_active)):
    uid = as_uuid(user_id)
    roles = [r for r in roles if r in ROLES]
    with connection() as conn:
        before = conn.execute("select roles, manager_id, backup_id from ap_user where id = %s for update", (uid,)).fetchone()
        if before is None:
            return _users_page(request, user, error="That user no longer exists.", status=404)
        mgr = as_uuid(manager_id) if manager_id else None
        bak = as_uuid(backup_id) if backup_id else None
        if mgr == uid or bak == uid:
            return _users_page(request, user, error="A person cannot be their own manager or backup.", status=400)
        old_roles = roles_list(before["roles"])
        current = {"roles": sorted(old_roles), "manager_id": before["manager_id"], "backup_id": before["backup_id"]}
        after = {"roles": sorted(roles), "manager_id": mgr, "backup_id": bak}
        diff = {k: {"from": current[k] if k == "roles" else (str(current[k]) if current[k] is not None else None),
                    "to": v if k == "roles" else (str(v) if v is not None else None)}
                for k, v in after.items() if current[k] != v}
        if diff:
            conn.execute("update ap_user set roles = %s::portal_role[], manager_id = %s, backup_id = %s where id = %s", (roles, mgr, bak, uid))
            reason = "grant" if ("roles" in diff and not old_roles) else "role_change"
            _admin_event(conn, user, reason, diff, uid)
    return RedirectResponse("/admin", status_code=303)


@router.post("/users/{user_id}/freeze")
def freeze_user(request: Request, user_id: str, user=Depends(ADMINS), _active=Depends(require_active)):
    uid = as_uuid(user_id)
    with connection() as conn:
        row = conn.execute("select status from ap_user where id = %s for update", (uid,)).fetchone()
        if row and row["status"] == "active":
            conn.execute("update ap_user set status = 'frozen' where id = %s", (uid,))
            from accounting.identity.sessions import revoke_all_for_user
            revoke_all_for_user(conn, uid, "frozen_by_admin")
            _admin_event(conn, user, "freeze", {"status": {"from": "active", "to": "frozen"}}, uid)
    return RedirectResponse("/admin", status_code=303)


@router.post("/users/{user_id}/unfreeze")
def unfreeze_user(request: Request, user_id: str, user=Depends(ADMINS), _active=Depends(require_active)):
    uid = as_uuid(user_id)
    with connection() as conn:
        row = conn.execute("select status from ap_user where id = %s for update", (uid,)).fetchone()
        if row and row["status"] == "frozen":
            conn.execute("update ap_user set status = 'active' where id = %s", (uid,))
            _admin_event(conn, user, "unfreeze", {"status": {"from": "frozen", "to": "active"}}, uid)
    return RedirectResponse("/admin", status_code=303)


@router.post("/users/{user_id}/disable")
def disable_user(request: Request, user_id: str, user=Depends(ADMINS), _active=Depends(require_active)):
    """ID-5 on the portal side: sessions die now; open approvals move to the backup."""
    uid = as_uuid(user_id)
    with connection() as conn:
        row = conn.execute("select status from ap_user where id = %s for update", (uid,)).fetchone()
        if row and row["status"] != "disabled":
            conn.execute("update ap_user set status = 'disabled' where id = %s", (uid,))
            from accounting.identity.sessions import revoke_all_for_user
            revoke_all_for_user(conn, uid, "disabled_by_admin")
            _admin_event(conn, user, "disable", {"status": {"from": row["status"], "to": "disabled"}}, uid)
            try:
                from accounting.routing.reassign import reassign_open_items_for
                reassign_open_items_for(conn, uid)
            except ImportError:
                log.warning("routing.reassign not available; open items for %s not reassigned", uid)
    return RedirectResponse("/admin", status_code=303)


@router.post("/users/{user_id}/enrollment-code")
def issue_code(request: Request, user_id: str, user=Depends(ADMINS), _active=Depends(require_active)):
    """ID-4: a six-digit code, shown ONCE to the Controller, who passes it on by phone or in person.
    Only its hash is stored, so the screen is the only place it ever appears."""
    uid = as_uuid(user_id)
    code = f"{secrets.randbelow(1_000_000):06d}"
    with connection() as conn:
        target = conn.execute("select display_name from ap_user where id = %s", (uid,)).fetchone()
        if target is None:
            return _users_page(request, user, error="That user no longer exists.", status=404)
        conn.execute("insert into enrollment_code (user_id, code_hash, issued_by, expires_at) values (%s, %s, %s, %s)",
                     (uid, hashlib.sha256(code.encode()).digest(), user.id, datetime.now(timezone.utc) + timedelta(hours=ENROLLMENT_CODE_HOURS)))
        _admin_event(conn, user, "enrollment_code_issued", {"expires_hours": ENROLLMENT_CODE_HOURS}, uid)
    return _users_page(request, user, new_code=code, code_for=target["display_name"])


# --- config under dual control -------------------------------------------------------------------------------

def _config_page(request, user, error=None, status=200):
    with connection() as conn:
        rows = conn.execute(
            """select c.*, p.display_name as proposer, a.display_name as approver from config_version c
               join ap_user p on p.id = c.proposed_by left join ap_user a on a.id = c.approved_by
               order by c.key, c.proposed_at desc""").fetchall()
        current = conn.execute("select * from config_current order by key").fetchall()
        dual = _dual_control_on(conn)
    return page(request, user, "admin/config.html", {"rows": rows, "current": current, "keys": CONFIG_KEYS, "dual_control": dual, "error": error}, status)


@router.get("/config")
def config(request: Request, user=Depends(ADMINS)):
    return _config_page(request, user)


@router.post("/config/propose")
def config_propose(request: Request, key: str = Form(""), value: str = Form(""), note: str = Form(""),
                   user=Depends(ADMINS), _active=Depends(require_active)):
    key = key.strip()
    if not key or not key.replace("_", "").isalnum():
        return _config_page(request, user, error="The setting name may only contain letters, numbers and underscores.", status=400)
    try:
        val = _parse_value(value)
    except ValueError as e:
        return _config_page(request, user, error=str(e), status=400)
    with connection() as conn:
        auto = (not _dual_control_on(conn)) and user.is_admin
        row = conn.execute(
            """insert into config_version (key, value, proposed_by, note, approved_by, approved_at, effective_from)
               values (%s, %s::jsonb, %s, %s, %s, %s, %s) returning id""",
            (key, json.dumps(val), user.id, note.strip() or None,
             user.id if auto else None, datetime.now(timezone.utc) if auto else None, datetime.now(timezone.utc) if auto else None),
        ).fetchone()
        _admin_event(conn, user, "config_auto_approved" if auto else "config_proposed",
                     {"key": key, "value": val, "version": row["id"], "note": "dual control is off; the administrator's proposal took effect at once" if auto else None})
    return RedirectResponse("/admin/config", status_code=303)


@router.post("/config/{version_id}/approve")
def config_approve(request: Request, version_id: int, user=Depends(ADMINS), _active=Depends(require_active)):
    with connection() as conn:
        row = conn.execute("select * from config_version where id = %s for update", (version_id,)).fetchone()
        if row is None or row["approved_at"] is not None:
            return _config_page(request, user, error="That proposal is not waiting for approval.", status=400)
        if row["proposed_by"] == user.id:
            return _config_page(request, user, error="The person who proposed a change cannot also approve it. Ask a second administrator or the Controller.", status=403)
        effective = add_business_days(datetime.now(timezone.utc), 1)
        conn.execute("update config_version set approved_by = %s, approved_at = now(), effective_from = %s where id = %s",
                     (user.id, effective, version_id))
        _admin_event(conn, user, "config_approved", {"key": row["key"], "version": version_id, "effective_from": effective.isoformat()})
    return RedirectResponse("/admin/config", status_code=303)


# --- approval matrix -----------------------------------------------------------------------------------------

def _matrix_page(request, user, error=None, status=200):
    with connection() as conn:
        rows = conn.execute(
            """select m.*, p.display_name as proposer, a.display_name as approver from matrix_version m
               join ap_user p on p.id = m.proposed_by left join ap_user a on a.id = m.approved_by order by m.id desc""").fetchall()
    for r in rows:
        r["rules_text"] = json.dumps(r["rules"], indent=2)
    return page(request, user, "admin/matrix.html", {"rows": rows, "error": error}, status)


@router.get("/matrix")
def matrix(request: Request, user=Depends(ADMINS)):
    return _matrix_page(request, user)


@router.post("/matrix/propose")
def matrix_propose(request: Request, rules: str = Form(""), user=Depends(ADMINS), _active=Depends(require_active)):
    try:
        parsed = validate_rules(json.loads(rules or "null"))
    except ValueError as e:
        msg = str(e) if not isinstance(e, json.JSONDecodeError) else "The rules are not valid JSON."
        return _matrix_page(request, user, error=msg, status=400)
    with connection() as conn:
        try:
            from accounting.routing.matrix import propose
            vid = propose(conn, parsed, user.id)
        except ImportError:
            vid = conn.execute("insert into matrix_version (rules, proposed_by) values (%s::jsonb, %s) returning id",
                               (json.dumps(parsed), user.id)).fetchone()["id"]
        _admin_event(conn, user, "matrix_proposed", {"version": int(vid) if isinstance(vid, int) else str(vid), "rules": parsed})
    return RedirectResponse("/admin/matrix", status_code=303)


@router.post("/matrix/{version_id}/approve")
def matrix_approve(request: Request, version_id: int, user=Depends(ADMINS), _active=Depends(require_active)):
    with connection() as conn:
        row = conn.execute("select * from matrix_version where id = %s for update", (version_id,)).fetchone()
        if row is None or row["approved_at"] is not None:
            return _matrix_page(request, user, error="That version is not waiting for approval.", status=400)
        if row["proposed_by"] == user.id and (_dual_control_on(conn) or not user.is_admin):
            return _matrix_page(request, user, error="The person who proposed a matrix cannot also approve it.", status=403)
        try:
            from accounting.routing.matrix import approve
            approve(conn, version_id, user.id)
        except ImportError:
            effective = add_business_days(datetime.now(timezone.utc), 1)
            conn.execute("update matrix_version set approved_by = %s, approved_at = now(), effective_from = %s where id = %s",
                         (user.id, effective, version_id))
        _admin_event(conn, user, "matrix_approved", {"version": version_id})
    return RedirectResponse("/admin/matrix", status_code=303)


# --- delegations, security events, grants report ------------------------------------------------------------------

@router.get("/delegations")
def delegations(request: Request, user=Depends(ADMINS)):
    with connection() as conn:
        rows = conn.execute(
            """select d.*, f.display_name as from_name, t.display_name as to_name from delegation d
               join ap_user f on f.id = d.from_user_id join ap_user t on t.id = d.to_user_id
               order by d.revoked_at is not null, d.ends_on desc""").fetchall()
    return page(request, user, "admin/delegations.html", {"rows": rows, "today": datetime.now(timezone.utc).date()})


@router.get("/security")
def security(request: Request, user=Depends(ADMINS)):
    with connection() as conn:
        rows = conn.execute(
            """select s.*, u.display_name, u.email from security_event s left join ap_user u on u.id = s.user_id
               order by s.at desc limit 200""").fetchall()
    return page(request, user, "admin/security.html", {"rows": rows})


@router.get("/grants")
def grants(request: Request, user=Depends(ADMINS)):
    """ID-5: every grant, removal, disable, freeze and sign-in event in the last seven days."""
    with connection() as conn:
        rows = conn.execute(
            """select e.at, e.event_type, e.reason_code, e.payload, u.display_name as actor from event_log e
               left join ap_user u on u.id = e.actor_id
               where e.event_type in ('admin', 'access') and e.at >= now() - interval '7 days' order by e.at desc""").fetchall()
    for r in rows:
        r["payload_text"] = json.dumps(r["payload"], default=str, sort_keys=True)
    return page(request, user, "admin/grants.html", {"rows": rows})


# --- break-glass ----------------------------------------------------------------------------------------------------

@router.post("/break-glass")
def break_glass(request: Request, on: str = Form("false"), reason: str = Form(""), user=Depends(BREAK_GLASS), _active=Depends(require_active)):
    """CFO or admin only. Takes effect at once and is itself a config version, so the history of every
    time it was switched on — and by whom — is in the same table as every other setting."""
    value = on.strip().lower() in ("1", "true", "yes", "on")
    with connection() as conn:
        conn.execute(
            """insert into config_version (key, value, proposed_by, note, approved_by, approved_at, effective_from)
               values ('break_glass', %s::jsonb, %s, %s, %s, now(), now())""",
            (json.dumps(value), user.id, reason.strip() or None, user.id))
        seq = _admin_event(conn, user, "break_glass_on" if value else "break_glass_off", {"reason": reason.strip()[:500]})
        conn.execute("insert into security_event (type, user_id, source, detail, event_seq) values (%s, %s, 'admin', %s::jsonb, %s)",
                     ("break_glass_on" if value else "break_glass_off", user.id, json.dumps({"reason": reason.strip()[:500]}), seq))
    return RedirectResponse("/admin", status_code=303)
