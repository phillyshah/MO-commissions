"""The invoice state machine (spec §Invoice lifecycle) and the hash-chained event log (AUD-1).

Every state change goes through `transition()`. It checks the (from, to) pair against TRANSITIONS,
updates the invoice row, and writes an event_log row whose hash chains to the previous row.
No other code path may change `invoice.state`.

Hash chain: hash = sha256(prev_hash || canonical_json({seq-less fields})). The first row chains to
32 zero bytes. `verify_chain()` re-walks it; the off-box verifier in Lenny's AWS account does the same
from the last anchor.
"""
from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from typing import Any
from uuid import UUID

import psycopg

from accounting.db.pool import connection

TERMINAL = {"DIVERTED", "DUPLICATE", "FRAUD", "REJECTED", "ENTERED", "VOID"}

# (from, to) -> set of allowed reason codes (None = any reason_code)
TRANSITIONS: dict[tuple[str, str], set[str] | None] = {
    ("RECEIVED", "DIVERTED"): {"statement", "credit_memo", "intercompany", "other_company_code", "non_invoice", "bank_change_request"},
    ("RECEIVED", "DUPLICATE"): {"file_hash", "invoice_number"},
    ("RECEIVED", "CAPTURED"): {"capture_ok"},
    ("RECEIVED", "AP_REVIEW"): {"validator_failed", "low_confidence", "split_unconfirmed", "missing_po", "manual_upload"},
    ("RECEIVED", "QUARANTINED"): {"hard_trigger"},
    ("CAPTURED", "QUARANTINED"): {"hard_trigger"},
    ("AP_REVIEW", "QUARANTINED"): {"hard_trigger"},
    ("QUARANTINED", "CAPTURED"): {"controller_release"},
    ("QUARANTINED", "FRAUD"): {"controller_reject", "controller_report"},
    ("AP_REVIEW", "CAPTURED"): {"clerk_edit", "clerk_confirmed"},
    ("AP_REVIEW", "DIVERTED"): {"statement", "credit_memo", "intercompany", "other_company_code", "non_invoice"},
    ("AP_REVIEW", "DUPLICATE"): {"invoice_number", "file_hash"},
    ("CAPTURED", "VENDOR_SETUP"): {"vendor_unknown"},
    ("VENDOR_SETUP", "CAPTURED"): {"vendor_added"},
    ("CAPTURED", "APPROVED"): {"auto_release_recurring"},
    ("CAPTURED", "PENDING_APPROVAL"): {"routed"},
    ("CAPTURED", "AP_REVIEW"): {"asl_hard_stop", "quality_no_po", "field_edit"},
    ("PENDING_APPROVAL", "QUESTIONED"): {"approver_question"},
    ("QUESTIONED", "PENDING_APPROVAL"): {"ap_answered"},
    ("PENDING_APPROVAL", "APPROVED"): {"chain_complete"},
    ("PENDING_APPROVAL", "REJECTED"): {"signed_rejection"},
    ("PENDING_APPROVAL", "AP_REVIEW"): {"field_edit", "reassign_exhausted"},
    ("QUESTIONED", "AP_REVIEW"): {"field_edit"},
    ("APPROVED", "AP_REVIEW"): {"field_edit"},
    ("APPROVED", "ENTERED"): {"sap_entered"},
}
for _s in ("RECEIVED", "CAPTURED", "AP_REVIEW", "QUARANTINED", "VENDOR_SETUP", "PENDING_APPROVAL", "QUESTIONED", "APPROVED"):
    TRANSITIONS[(_s, "VOID")] = {"controller_void"}

ZERO_HASH = b"\x00" * 32


class InvalidTransition(Exception):
    pass


def canonical(obj: Any) -> bytes:
    """RFC 8785-style canonical JSON: sorted keys, no whitespace, UTF-8, non-ASCII kept."""
    return json.dumps(obj, sort_keys=True, separators=(",", ":"), ensure_ascii=False, default=_default).encode("utf-8")


def _default(o: Any) -> Any:
    if isinstance(o, (datetime,)):
        return o.astimezone(timezone.utc).isoformat()
    if isinstance(o, UUID):
        return str(o)
    if isinstance(o, (bytes, memoryview)):
        return bytes(o).hex()
    raise TypeError(f"not JSON serialisable: {type(o)!r}")


def _row_hash(prev_hash: bytes, row: dict) -> bytes:
    return hashlib.sha256(prev_hash + canonical(row)).digest()


def _head(conn: psycopg.Connection) -> bytes:
    r = conn.execute("select hash from event_log order by seq desc limit 1 for update").fetchone()
    return bytes(r["hash"]) if r else ZERO_HASH


def log_event(
    conn: psycopg.Connection,
    *,
    event_type: str,
    actor_type: str,
    actor_id: UUID | str | None = None,
    invoice_id: UUID | str | None = None,
    from_state: str | None = None,
    to_state: str | None = None,
    reason_code: str | None = None,
    payload: dict | None = None,
) -> int:
    """Append one chained row inside the caller's transaction. Returns seq."""
    conn.execute("lock table event_log in exclusive mode")
    prev = _head(conn)
    at = datetime.now(timezone.utc)
    body = {
        "at": at,
        "invoice_id": str(invoice_id) if invoice_id else None,
        "from_state": from_state,
        "to_state": to_state,
        "event_type": event_type,
        "actor_id": str(actor_id) if actor_id else None,
        "actor_type": actor_type,
        "reason_code": reason_code,
        "payload": payload or {},
    }
    h = _row_hash(prev, body)
    r = conn.execute(
        """insert into event_log (at, invoice_id, from_state, to_state, event_type, actor_id, actor_type,
                                  reason_code, payload, prev_hash, hash)
           values (%s,%s,%s,%s,%s,%s,%s,%s,%s::jsonb,%s,%s) returning seq""",
        (at, invoice_id, from_state, to_state, event_type, actor_id, actor_type, reason_code,
         json.dumps(body["payload"], default=_default), prev, h),
    ).fetchone()
    return int(r["seq"])


def transition(
    conn: psycopg.Connection,
    invoice_id: UUID | str,
    to_state: str,
    *,
    reason_code: str,
    actor_type: str,
    actor_id: UUID | str | None = None,
    payload: dict | None = None,
) -> dict:
    """Move an invoice. Raises InvalidTransition if the table forbids it. Returns the updated row."""
    row = conn.execute("select * from invoice where id = %s for update", (invoice_id,)).fetchone()
    if row is None:
        raise InvalidTransition(f"invoice {invoice_id} not found")
    from_state = row["state"]
    allowed = TRANSITIONS.get((from_state, to_state), "missing")
    if allowed == "missing":
        raise InvalidTransition(f"{from_state} -> {to_state} is not a valid transition")
    if allowed is not None and reason_code not in allowed:
        raise InvalidTransition(f"{from_state} -> {to_state} does not accept reason {reason_code!r}")

    sets = ["state = %s", "updated_at = now()"]
    params: list[Any] = [to_state]
    # Two clocks: the SLA clock pauses while QUESTIONED; sponsor aging never pauses.
    if to_state == "QUESTIONED":
        sets.append("sla_paused_at = now()")
    if from_state == "QUESTIONED":
        sets.append("sla_paused_total = sla_paused_total + (now() - coalesce(sla_paused_at, now()))")
        sets.append("sla_paused_at = null")
    if to_state == "PENDING_APPROVAL" and row["routed_at"] is None:
        sets.append("routed_at = now()")
    if to_state == "APPROVED":
        sets.append("approved_at = now()")
    if to_state == "ENTERED":
        sets.append("entered_at = now()")
    params.append(invoice_id)
    updated = conn.execute(
        f"update invoice set {', '.join(sets)} where id = %s returning *", params
    ).fetchone()
    log_event(
        conn,
        event_type="transition",
        actor_type=actor_type,
        actor_id=actor_id,
        invoice_id=invoice_id,
        from_state=from_state,
        to_state=to_state,
        reason_code=reason_code,
        payload=payload,
    )
    return updated


def verify_chain(conn: psycopg.Connection | None = None, start_seq: int = 1) -> tuple[bool, int | None]:
    """Re-walk the chain. Returns (ok, first_bad_seq)."""
    def _walk(c: psycopg.Connection) -> tuple[bool, int | None]:
        prev_row = c.execute("select hash from event_log where seq < %s order by seq desc limit 1", (start_seq,)).fetchone()
        prev = bytes(prev_row["hash"]) if prev_row else ZERO_HASH
        cur = c.execute("select * from event_log where seq >= %s order by seq", (start_seq,))
        for r in cur:
            body = {
                "at": r["at"],
                "invoice_id": str(r["invoice_id"]) if r["invoice_id"] else None,
                "from_state": r["from_state"],
                "to_state": r["to_state"],
                "event_type": r["event_type"],
                "actor_id": str(r["actor_id"]) if r["actor_id"] else None,
                "actor_type": r["actor_type"],
                "reason_code": r["reason_code"],
                "payload": r["payload"],
            }
            if bytes(r["prev_hash"]) != prev or _row_hash(prev, body) != bytes(r["hash"]):
                return False, int(r["seq"])
            prev = bytes(r["hash"])
        return True, None

    if conn is not None:
        return _walk(conn)
    with connection() as c:
        return _walk(c)


def chain_head(conn: psycopg.Connection) -> tuple[int, bytes] | None:
    r = conn.execute("select seq, hash from event_log order by seq desc limit 1").fetchone()
    return (int(r["seq"]), bytes(r["hash"])) if r else None
