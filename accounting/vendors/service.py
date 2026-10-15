"""Vendor writes: every change is a `vendor_change` row, applied now or after a second person approves.

VEN-3 says sponsor, routing, allowed domains and recurring schedules change under dual control, and
that Andy may act alone until the switch is turned on before cutover. Both paths write the same
`vendor_change` row — the only difference is whether `applied_at` is set in the same transaction —
so the history reads identically whichever mode a change was made under.

Secrets in a change row: the TIN travels encrypted (AAD "vendor_change") and is re-encrypted under the
vendor id when applied; bank routing/account are reduced to (fingerprint, last4) BEFORE the row is
written, so a pending change never holds an account number either.
"""
from __future__ import annotations

import base64
import csv
import io
import json
from dataclasses import dataclass, field
from uuid import UUID

import psycopg

from accounting.config_flags import config_flag
from accounting.identity import crypto
from accounting.lifecycle import log_event
from accounting.vendors import fingerprint as fp
from accounting.vendors.normalize import digits, name_key

EDITABLE = (
    "sap_vendor_number", "legal_name", "country", "default_sponsor_id", "backup_sponsor_id", "default_gl",
    "default_cost_center", "payment_terms", "hcp", "quality_critical", "asl_status", "fcpa_agent",
    "intercompany", "recurring_schedule", "recurring_tolerance_pct", "po_required", "phone", "status",
)
BOOL_FIELDS = {"hcp", "quality_critical", "fcpa_agent", "intercompany", "po_required"}
REMIT_KEYS = ("name", "line1", "line2", "city", "state", "postal", "country")
_TIN_AAD = b"vendor_change"


class VendorError(Exception):
    """A message written for the person at the form. Never wraps a raw exception string."""


@dataclass
class ImportReport:
    created: int = 0
    updated: int = 0
    pending: int = 0
    errors: list[str] = field(default_factory=list)


def get(conn: psycopg.Connection, vendor_id: UUID | str) -> dict | None:
    return conn.execute("select * from vendor where id = %s", (vendor_id,)).fetchone()


def _diff(current: dict | None, fields: dict) -> dict:
    """{field: {from, to}} for what actually changes. `tin` and `bank` are handled by the caller."""
    out: dict = {}
    for k, v in fields.items():
        if k in ("tin", "bank_routing", "bank_account"):
            continue
        if k == "remit_to":
            v = {rk: (str(v.get(rk)).strip() or None) if v and v.get(rk) is not None else None for rk in REMIT_KEYS} if v else None
            if v and not any(v.values()):
                v = None
        elif k in BOOL_FIELDS:
            v = bool(v)
        elif k == "recurring_tolerance_pct":
            v = str(v).strip() if v not in (None, "") else None
        elif k in ("default_sponsor_id", "backup_sponsor_id"):
            v = str(v) if v else None
        else:
            v = (str(v).strip() or None) if v is not None else None
        cur = current.get(k) if current else None
        if isinstance(cur, UUID):
            cur = str(cur)
        if k == "recurring_tolerance_pct" and cur is not None:
            cur = str(cur)
        if cur != v:
            out[k] = {"from": cur, "to": v}
    return out


def _bank_to_change(fields: dict) -> dict | None:
    routing, account = fields.get("bank_routing"), fields.get("bank_account")
    if not digits(account):
        return None
    f, last4 = fp.fingerprint(routing or "", account)
    return {"fingerprint_hex": f.hex(), "last4": last4}


def propose_change(conn: psycopg.Connection, vendor_id: UUID | str | None, fields: dict, by: UUID | str,
                   *, source: str = "form", note: str | None = None) -> tuple[dict | None, bool]:
    """Returns (change_row, applied). `change_row` is None when nothing changed."""
    current = get(conn, vendor_id) if vendor_id else None
    if vendor_id and current is None:
        raise VendorError("That vendor no longer exists.")
    diff = _diff(current, fields)
    if fields.get("tin"):
        t = digits(fields["tin"])
        if len(t) != 9:
            raise VendorError("A TIN is nine digits.")
        diff["tin"] = {"from": "set" if current and current.get("tin_enc") else None,
                       "to_enc": base64.b64encode(crypto.encrypt(t, aad=_TIN_AAD)).decode("ascii"),
                       "to": f"set (ends {t[-4:]})"}
    bank = _bank_to_change(fields)
    if bank:
        cur_fp = fp.current(conn, vendor_id) if vendor_id else None
        if cur_fp is None or bytes(cur_fp["fingerprint"]).hex() != bank["fingerprint_hex"]:
            diff["bank"] = {"from": cur_fp["last4"] if cur_fp else None, "to": bank["last4"], "fingerprint_hex": bank["fingerprint_hex"]}
    if not diff:
        return None, False
    if current is None and not (fields.get("legal_name") or "").strip():
        raise VendorError("A legal name is required.")
    row = conn.execute(
        """insert into vendor_change (vendor_id, diff, proposed_by, note) values (%s, %s::jsonb, %s, %s) returning *""",
        (vendor_id, json.dumps(diff), by, note or source),
    ).fetchone()
    if config_flag(conn, "dual_control_enabled"):
        log_event(conn, event_type="vendor_change_proposed", actor_type="user", actor_id=by,
                  payload={"change_id": str(row["id"]), "vendor_id": str(vendor_id) if vendor_id else None,
                           "fields": sorted(diff.keys())})
        return row, False
    applied = _apply(conn, row, by, approved_by=None)
    return applied, True


def approve_change(conn: psycopg.Connection, change_id: UUID | str, by: UUID | str) -> dict:
    row = conn.execute("select * from vendor_change where id = %s for update", (change_id,)).fetchone()
    if row is None:
        raise VendorError("That change no longer exists.")
    if row["applied_at"] or row["rejected_at"]:
        raise VendorError("That change has already been decided.")
    if str(row["proposed_by"]) == str(by):
        raise VendorError("A change must be approved by someone other than the person who proposed it.")
    return _apply(conn, row, by, approved_by=by)


def reject_change(conn: psycopg.Connection, change_id: UUID | str, by: UUID | str, reason: str) -> None:
    row = conn.execute("select * from vendor_change where id = %s for update", (change_id,)).fetchone()
    if row is None or row["applied_at"] or row["rejected_at"]:
        raise VendorError("That change has already been decided.")
    conn.execute("update vendor_change set rejected_at = now(), approved_by = %s, note = %s where id = %s", (by, reason, change_id))
    log_event(conn, event_type="vendor_change_rejected", actor_type="user", actor_id=by,
              payload={"change_id": str(change_id), "reason": reason})


def _apply(conn: psycopg.Connection, change: dict, actor: UUID | str, approved_by: UUID | str | None) -> dict:
    diff: dict = change["diff"]
    vendor_id = change["vendor_id"]
    cols = {k: v["to"] for k, v in diff.items() if k in EDITABLE or k == "remit_to"}
    if "legal_name" in cols or vendor_id is None:
        name = cols.get("legal_name") or (get(conn, vendor_id) or {}).get("legal_name")
        if not name:
            raise VendorError("A legal name is required.")
        cols["legal_name"] = name
        cols["name_key"] = name_key(name)
    if vendor_id is None:
        cols.setdefault("status", "active")
        keys = list(cols.keys())
        placeholders = ", ".join("%s::jsonb" if k == "remit_to" else "%s" for k in keys)
        values = [json.dumps(cols[k]) if k == "remit_to" else cols[k] for k in keys]
        v = conn.execute(f"insert into vendor ({', '.join(keys)}) values ({placeholders}) returning *", values).fetchone()
        vendor_id = v["id"]
    elif cols:
        sets = ", ".join(f"{k} = %s::jsonb" if k == "remit_to" else f"{k} = %s" for k in cols)
        values = [json.dumps(cols[k]) if k == "remit_to" else cols[k] for k in cols]
        conn.execute(f"update vendor set {sets}, updated_at = now() where id = %s", values + [vendor_id])
    if "tin" in diff:
        t = crypto.decrypt(base64.b64decode(diff["tin"]["to_enc"]), aad=_TIN_AAD).decode()
        conn.execute("update vendor set tin_enc = %s, tin_status = coalesce(tin_status, 'pending'), updated_at = now() where id = %s",
                     (crypto.encrypt(t, aad=str(vendor_id).encode()), vendor_id))
    if "bank" in diff:
        fp.record(conn, vendor_id, bytes.fromhex(diff["bank"]["fingerprint_hex"]), diff["bank"]["to"],
                  "imported_vendor_list" if (change.get("note") == "csv_import") else "ap_verified_invoice", actor)
    if "phone" in diff:
        conn.execute("update vendor set phone_since = current_date where id = %s", (vendor_id,))
    row = conn.execute(
        "update vendor_change set vendor_id = %s, applied_at = now(), approved_by = %s where id = %s returning *",
        (vendor_id, approved_by, change["id"]),
    ).fetchone()
    safe_diff = {k: ({"from": v.get("from"), "to": v.get("to")}) for k, v in diff.items()}
    log_event(conn, event_type="vendor_change_applied", actor_type="user", actor_id=actor,
              payload={"change_id": str(change["id"]), "vendor_id": str(vendor_id), "diff": safe_diff,
                       "proposed_by": str(change["proposed_by"]), "approved_by": str(approved_by) if approved_by else None})
    return row


# --- allowed domains and platform map ------------------------------------------------------------
def add_domain(conn: psycopg.Connection, vendor_id: UUID | str, pattern: str, by: UUID | str) -> dict:
    p = (pattern or "").strip().lower().lstrip("@")
    if not p or " " in p or "." not in p:
        raise VendorError("Enter a domain (acme.com) or a full address (billing@acme.com).")
    dual = config_flag(conn, "dual_control_enabled")
    row = conn.execute(
        """insert into vendor_domain (vendor_id, pattern, status, added_by, approved_by)
           values (%s, %s, %s, %s, %s)
           on conflict (vendor_id, pattern) do update set status = excluded.status, approved_by = excluded.approved_by
           returning *""",
        (vendor_id, p, "proposed" if dual else "approved", by, None if dual else by),
    ).fetchone()
    log_event(conn, event_type="vendor_domain", actor_type="user", actor_id=by,
              reason_code="proposed" if dual else "approved", payload={"vendor_id": str(vendor_id), "pattern": p})
    return row


def approve_domain(conn: psycopg.Connection, domain_id: UUID | str, by: UUID | str) -> dict:
    row = conn.execute("select * from vendor_domain where id = %s for update", (domain_id,)).fetchone()
    if row is None:
        raise VendorError("That domain entry no longer exists.")
    if config_flag(conn, "dual_control_enabled") and str(row["added_by"]) == str(by):
        raise VendorError("A domain must be approved by someone other than the person who added it.")
    row = conn.execute("update vendor_domain set status = 'approved', approved_by = %s where id = %s returning *", (by, domain_id)).fetchone()
    log_event(conn, event_type="vendor_domain", actor_type="user", actor_id=by, reason_code="approved",
              payload={"vendor_id": str(row["vendor_id"]), "pattern": row["pattern"]})
    return row


def revoke_domain(conn: psycopg.Connection, domain_id: UUID | str, by: UUID | str) -> None:
    row = conn.execute("update vendor_domain set status = 'revoked' where id = %s returning *", (domain_id,)).fetchone()
    if row:
        log_event(conn, event_type="vendor_domain", actor_type="user", actor_id=by, reason_code="revoked",
                  payload={"vendor_id": str(row["vendor_id"]), "pattern": row["pattern"]})


def add_platform_map(conn: psycopg.Connection, vendor_id: UUID | str, platform_id: str, account_id: str, by: UUID | str) -> dict:
    account_id = (account_id or "").strip()
    if not account_id:
        raise VendorError("Enter the vendor's account id on that platform.")
    if conn.execute("select 1 from billing_platform where id = %s", (platform_id,)).fetchone() is None:
        raise VendorError("Unknown billing platform.")
    dual = config_flag(conn, "dual_control_enabled")
    try:
        row = conn.execute(
            """insert into billing_platform_map (platform_id, vendor_id, account_id, added_by, approved_by)
               values (%s,%s,%s,%s,%s) returning *""",
            (platform_id, vendor_id, account_id, by, None if dual else by),
        ).fetchone()
    except psycopg.errors.UniqueViolation:
        raise VendorError("That platform account id is already mapped to a vendor.")
    log_event(conn, event_type="vendor_platform_map", actor_type="user", actor_id=by,
              reason_code="proposed" if dual else "approved",
              payload={"vendor_id": str(vendor_id), "platform": platform_id, "account_id": account_id})
    return row


def approve_platform_map(conn: psycopg.Connection, map_id: UUID | str, by: UUID | str) -> dict:
    row = conn.execute("select * from billing_platform_map where id = %s for update", (map_id,)).fetchone()
    if row is None:
        raise VendorError("That mapping no longer exists.")
    if config_flag(conn, "dual_control_enabled") and str(row["added_by"]) == str(by):
        raise VendorError("A mapping must be approved by someone other than the person who added it.")
    row = conn.execute("update billing_platform_map set approved_by = %s where id = %s returning *", (by, map_id)).fetchone()
    log_event(conn, event_type="vendor_platform_map", actor_type="user", actor_id=by, reason_code="approved",
              payload={"vendor_id": str(row["vendor_id"]), "platform": row["platform_id"], "account_id": row["account_id"]})
    return row


# --- CSV import ---------------------------------------------------------------------------------
CSV_COLUMNS = (
    "sap_vendor_number", "legal_name", "tin", "remit_to_line1", "remit_to_city", "remit_to_state", "remit_to_postal",
    "country", "sponsor_email", "backup_sponsor_email", "default_gl", "default_cost_center", "payment_terms", "hcp",
    "quality_critical", "asl_status", "fcpa_agent", "intercompany", "recurring_schedule", "allowed_domains",
    "bank_routing", "bank_account", "phone",
)


def _truthy(s: str | None) -> bool:
    return (s or "").strip().lower() in ("1", "true", "yes", "y", "x", "on")


def _user_by_email(conn: psycopg.Connection, email: str | None) -> str | None:
    if not (email or "").strip():
        return None
    row = conn.execute("select id from ap_user where lower(email) = lower(%s)", (email.strip(),)).fetchone()
    return str(row["id"]) if row else None


def import_csv(conn: psycopg.Connection, text: str, by: UUID | str) -> ImportReport:
    """Upsert by SAP vendor number (or legal name when there is none). Each row goes through
    propose_change, so under dual control an import produces pending changes, not applied ones."""
    report = ImportReport()
    try:
        reader = csv.DictReader(io.StringIO(text))
        header = [h.strip().lower() for h in (reader.fieldnames or [])]
    except Exception:
        raise VendorError("The file could not be read as CSV.")
    missing = [c for c in ("legal_name",) if c not in header]
    if missing:
        raise VendorError("The CSV needs at least a legal_name column.")
    for n, raw in enumerate(reader, start=2):
        row = {(k or "").strip().lower(): (v or "").strip() for k, v in raw.items() if k}
        try:
            name = row.get("legal_name")
            if not name:
                report.errors.append(f"Row {n}: no legal name; skipped.")
                continue
            sap = row.get("sap_vendor_number") or None
            existing = None
            if sap:
                existing = conn.execute("select * from vendor where sap_vendor_number = %s", (sap,)).fetchone()
            if existing is None:
                existing = conn.execute("select * from vendor where name_key = %s", (name_key(name),)).fetchone()
            remit = {"name": name, "line1": row.get("remit_to_line1"), "line2": None, "city": row.get("remit_to_city"),
                     "state": row.get("remit_to_state"), "postal": row.get("remit_to_postal"), "country": row.get("country")}
            fields = {
                "sap_vendor_number": sap, "legal_name": name, "country": row.get("country") or None,
                "remit_to": remit if any(v for k, v in remit.items() if k != "name") else None,
                "default_sponsor_id": _user_by_email(conn, row.get("sponsor_email")),
                "backup_sponsor_id": _user_by_email(conn, row.get("backup_sponsor_email")),
                "default_gl": row.get("default_gl") or None, "default_cost_center": row.get("default_cost_center") or None,
                "payment_terms": row.get("payment_terms") or None, "hcp": _truthy(row.get("hcp")),
                "quality_critical": _truthy(row.get("quality_critical")), "asl_status": row.get("asl_status") or None,
                "fcpa_agent": _truthy(row.get("fcpa_agent")), "intercompany": _truthy(row.get("intercompany")),
                "recurring_schedule": row.get("recurring_schedule") or None, "phone": row.get("phone") or None,
                "tin": row.get("tin") or None, "bank_routing": row.get("bank_routing"), "bank_account": row.get("bank_account"),
            }
            if row.get("sponsor_email") and not fields["default_sponsor_id"]:
                report.errors.append(f"Row {n} ({name}): sponsor {row['sponsor_email']} has no portal account; sponsor left blank.")
            if existing is None:
                fields = {k: v for k, v in fields.items() if v not in (None, "", False) or k in ("legal_name",)}
            change, applied = propose_change(conn, existing["id"] if existing else None, fields, by, note="csv_import")
            if change is None:
                continue
            if not applied:
                report.pending += 1
            elif existing is None:
                report.created += 1
            else:
                report.updated += 1
            vendor_id = change["vendor_id"] if applied else None
            if vendor_id:
                for d in (row.get("allowed_domains") or "").split(";"):
                    if d.strip():
                        add_domain(conn, vendor_id, d, by)
        except VendorError as e:
            report.errors.append(f"Row {n}: {e}")
        finally:
            row.pop("bank_account", None)
            row.pop("bank_routing", None)
            row.pop("tin", None)
    log_event(conn, event_type="vendor_import", actor_type="user", actor_id=by,
              payload={"created": report.created, "updated": report.updated, "pending": report.pending, "errors": len(report.errors)})
    return report
