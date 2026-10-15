"""VEN-2: the bank fingerprint. A keyed HMAC over (routing, account) plus the last four digits.

The routing and account numbers are NEVER stored — not in this table, not in the event log, not in a
CSV import row that survives the request. Only the fingerprint (useless without the field key) and
the last four (what the Controller reads back on a callback) are kept. The dates on the rows are what
the 30-day (hard trigger) and 90-day (auto-release exclusion) rules read.
"""
from __future__ import annotations

from uuid import UUID

import psycopg

from accounting.identity import crypto
from accounting.lifecycle import log_event
from accounting.vendors.normalize import digits


def fingerprint(routing: str, account: str) -> tuple[bytes, str]:
    """Deterministic: same key, same digits, same bytes. Formatting ("021-000-021") does not matter."""
    r, a = digits(routing), digits(account)
    if not a:
        raise ValueError("an account number is required for a fingerprint")
    return crypto.hmac_fingerprint(r, a), a[-4:]


def current(conn: psycopg.Connection, vendor_id: UUID | str) -> dict | None:
    return conn.execute(
        "select * from vendor_bank_fingerprint where vendor_id = %s and superseded_at is null order by effective_from desc limit 1",
        (vendor_id,),
    ).fetchone()


def previous(conn: psycopg.Connection, vendor_id: UUID | str) -> dict | None:
    """The row the current one replaced — its superseded_at is the date of the last bank CHANGE."""
    return conn.execute(
        "select * from vendor_bank_fingerprint where vendor_id = %s and superseded_at is not null order by superseded_at desc limit 1",
        (vendor_id,),
    ).fetchone()


def record(conn: psycopg.Connection, vendor_id: UUID | str, fp: bytes, last4: str, source: str,
           verified_by: UUID | str | None) -> dict:
    """Supersede the current fingerprint (if different) and insert the new one. Recording the same
    fingerprint again is a no-op so a re-import does not look like a bank change."""
    cur = current(conn, vendor_id)
    if cur is not None and bytes(cur["fingerprint"]) == fp:
        return cur
    if cur is not None:
        conn.execute("update vendor_bank_fingerprint set superseded_at = now() where id = %s", (cur["id"],))
    row = conn.execute(
        """insert into vendor_bank_fingerprint (vendor_id, fingerprint, last4, source, verified_by)
           values (%s,%s,%s,%s,%s) returning *""",
        (vendor_id, fp, last4, source, verified_by),
    ).fetchone()
    log_event(conn, event_type="vendor_bank_fingerprint", actor_type="user" if verified_by else "system",
              actor_id=verified_by, reason_code="changed" if cur else "recorded",
              payload={"vendor_id": str(vendor_id), "last4": last4, "source": source, "previous_last4": cur["last4"] if cur else None})
    return row
