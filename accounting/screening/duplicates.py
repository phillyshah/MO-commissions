"""DUP-1, the portal half (SAP open/cleared items arrive in Stage 2).

Hard block (→ DUPLICATE): the same file (sha256), or the same normalised invoice number from the same
vendor — or from any vendor sharing a TIN or a bank fingerprint, because a fraudster re-submitting a
real invoice under a second vendor record is the case the TIN clause exists for.
Hard TRIGGER (→ QUARANTINED, not DUPLICATE): same number but a different amount or bank. That is not a
duplicate; it is the oldest invoice-fraud pattern there is, and the Controller sees it.
Flags (soft, AP review): same amount and date, same PO and amount, same amount within 30 days.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from uuid import UUID

import psycopg

from accounting.identity import crypto
from accounting.vendors.normalize import digits

LIVE_STATES = ("RECEIVED", "CAPTURED", "AP_REVIEW", "QUARANTINED", "VENDOR_SETUP", "PENDING_APPROVAL",
               "QUESTIONED", "APPROVED", "ENTERED")


@dataclass
class DupResult:
    duplicate_reason: str | None = None       # file_hash | invoice_number
    matched_invoice_id: UUID | None = None
    hard: list[str] = field(default_factory=list)
    soft: list[str] = field(default_factory=list)
    detail: dict = field(default_factory=dict)


def file_hash_duplicate(conn: psycopg.Connection, inv: dict) -> dict | None:
    if not inv.get("file_sha256"):
        return None
    # Only an OLDER row makes this one a duplicate: two copies arriving in the same poll must resolve to
    # one original and one duplicate, not two duplicates of each other.
    return conn.execute(
        """select id, portal_number, state from invoice where file_sha256 = %s and id <> %s and created_at < %s
           and coalesce(page_from, 0) = coalesce(%s, 0) and state = any(%s::invoice_state[]) order by created_at limit 1""",
        (inv["file_sha256"], inv["id"], inv["created_at"], inv.get("page_from"), list(LIVE_STATES)),
    ).fetchone()


def _related_vendor_ids(conn: psycopg.Connection, vendor_id) -> list:
    """The vendor plus any vendor sharing its TIN or current bank fingerprint."""
    ids = {str(vendor_id)}
    me = conn.execute("select tin_enc from vendor where id = %s", (vendor_id,)).fetchone()
    my_tin = None
    if me and me["tin_enc"]:
        try:
            my_tin = digits(crypto.decrypt(me["tin_enc"], aad=str(vendor_id).encode()).decode())
        except Exception:
            my_tin = None
    if my_tin:
        for v in conn.execute("select id, tin_enc from vendor where tin_enc is not null and id <> %s", (vendor_id,)).fetchall():
            try:
                if digits(crypto.decrypt(v["tin_enc"], aad=str(v["id"]).encode()).decode()) == my_tin:
                    ids.add(str(v["id"]))
            except Exception:
                continue
    for r in conn.execute(
        """select distinct b.vendor_id from vendor_bank_fingerprint a join vendor_bank_fingerprint b
             on a.fingerprint = b.fingerprint and a.vendor_id <> b.vendor_id
           where a.vendor_id = %s and a.superseded_at is null and b.superseded_at is null""", (vendor_id,)).fetchall():
        ids.add(str(r["vendor_id"]))
    return list(ids)


def check(conn: psycopg.Connection, inv: dict) -> DupResult:
    """`inv` is the invoice row with the extracted header fields already written to it."""
    out = DupResult()
    fh = file_hash_duplicate(conn, inv)
    if fh:
        out.duplicate_reason, out.matched_invoice_id = "file_hash", fh["id"]
        out.detail["matched_portal_number"] = fh["portal_number"]
        return out
    norm = inv.get("invoice_number_norm")
    vendor_id = inv.get("vendor_id")
    if norm and vendor_id:
        others = conn.execute(
            """select id, portal_number, gross_cents, bank_fingerprint, bank_last4, state from invoice
               where invoice_number_norm = %s and vendor_id::text = any(%s) and id <> %s and created_at < %s
                 and state = any(%s::invoice_state[])
               order by created_at""",
            (norm, _related_vendor_ids(conn, vendor_id), inv["id"], inv["created_at"], list(LIVE_STATES)),
        ).fetchall()
        for o in others:
            same_amount = o["gross_cents"] == inv.get("gross_cents")
            same_bank = (o["bank_fingerprint"] is None or inv.get("bank_fingerprint") is None
                         or bytes(o["bank_fingerprint"]) == bytes(inv["bank_fingerprint"]))
            if same_amount and same_bank:
                out.duplicate_reason, out.matched_invoice_id = "invoice_number", o["id"]
                out.detail["matched_portal_number"] = o["portal_number"]
                return out
            out.hard.append("duplicate_number_mismatch")
            out.detail["duplicate_number_mismatch"] = {
                "matched_portal_number": o["portal_number"], "their_gross_cents": o["gross_cents"],
                "our_gross_cents": inv.get("gross_cents"), "same_bank": same_bank}
            return out
    gross, inv_date, po = inv.get("gross_cents"), inv.get("invoice_date"), inv.get("po_number")
    if gross and vendor_id:
        if inv_date and conn.execute(
            "select 1 from invoice where vendor_id = %s and gross_cents = %s and invoice_date = %s and id <> %s and state = any(%s::invoice_state[]) limit 1",
            (vendor_id, gross, inv_date, inv["id"], list(LIVE_STATES))).fetchone():
            out.soft.append("same_amount_and_date")
        if po and conn.execute(
            "select 1 from invoice where vendor_id = %s and gross_cents = %s and po_number = %s and id <> %s and state = any(%s::invoice_state[]) limit 1",
            (vendor_id, gross, po, inv["id"], list(LIVE_STATES))).fetchone():
            out.soft.append("same_po_and_amount")
        if conn.execute(
            """select 1 from invoice where vendor_id = %s and gross_cents = %s and id <> %s and state = any(%s::invoice_state[])
               and received_at > %s::timestamptz - interval '30 days' and received_at < %s::timestamptz + interval '30 days' limit 1""",
            (vendor_id, gross, inv["id"], list(LIVE_STATES), inv["received_at"], inv["received_at"])).fetchone():
            out.soft.append("same_amount_within_30_days")
    return out
