"""Auto-release (APR-4): the only path on which an invoice is approved without a person.

It is deliberately narrow. The invoice must be from a vendor on a recurring schedule, must look like
the last one the company actually paid for a DIFFERENT service period, and nothing about the vendor or
the message may be out of the ordinary. Every negative returns a code the Controller can read in the
event log, because "why did this one not auto-release" is the question AP asks most.

Variance is integer arithmetic in basis points — a 2% tolerance on $12,345.67 is compared as
|variance| * 10000 <= prior * 200, never as a float.
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
from decimal import Decimal

import psycopg

from accounting.config import settings


def _tolerance_bp(vendor_pct) -> int:
    pct = Decimal(str(vendor_pct)) if vendor_pct is not None else Decimal(str(settings.recurring_tolerance_pct))
    return int(pct * 100)  # percent -> basis points, exact for two decimals


def prior_invoice(conn: psycopg.Connection, inv: dict) -> dict | None:
    """The most recent APPROVED/ENTERED invoice for this vendor whose invoice_date falls in a different
    calendar month — the service-period proxy until Stage 2 gives us real periods."""
    if not inv.get("vendor_id") or not inv.get("invoice_date"):
        return None
    return conn.execute(
        """select id, gross_cents, invoice_date, state from invoice
           where vendor_id = %s and id <> %s and state in ('APPROVED', 'ENTERED') and gross_cents is not null
             and invoice_date is not null
             and date_trunc('month', invoice_date) <> date_trunc('month', %s::date)
           order by invoice_date desc, created_at desc limit 1""",
        (inv["vendor_id"], inv["id"], inv["invoice_date"]),
    ).fetchone()


def eligible(conn: psycopg.Connection, inv: dict) -> tuple[bool, dict]:
    """(ok, detail). `inv` is `routing.common.load_invoice()`'s row (invoice + vendor_* columns)."""
    detail: dict = {"rule": "recurring_within_tolerance"}

    def no(code: str, **extra) -> tuple[bool, dict]:
        detail.update({"reason": code, **extra})
        return False, detail

    if not inv.get("vendor_id"):
        return no("no_vendor")
    if not inv.get("vendor_recurring_schedule"):
        return no("not_recurring")
    if inv.get("vendor_hcp"):
        return no("hcp_vendor")
    if inv.get("vendor_status") != "active":
        return no("vendor_not_active")
    if "released_from_quarantine" in (inv.get("tags") or []):
        return no("released_from_quarantine")
    if inv.get("soft_signals"):
        return no("soft_signals", signals=list(inv["soft_signals"]))
    if inv.get("gross_cents") is None:
        return no("no_amount")
    now = datetime.now(timezone.utc)
    created = inv.get("vendor_created_at")
    if created is not None and (now - created) < timedelta(days=30):
        return no("new_vendor")
    q = conn.execute("select count(*) as n from quarantine_case where invoice_id = %s", (inv["id"],)).fetchone()
    if q and q["n"]:
        return no("quarantine_history")
    bank = conn.execute(
        "select count(*) as n from vendor_bank_fingerprint where vendor_id = %s and effective_from > %s",
        (inv["vendor_id"], now - timedelta(days=90)),
    ).fetchone()
    if bank and bank["n"]:
        return no("bank_change_90d")
    prior = prior_invoice(conn, inv)
    if prior is None:
        return no("no_prior_period")
    prior_gross = int(prior["gross_cents"])
    gross = int(inv["gross_cents"])
    variance = gross - prior_gross
    tol_bp = _tolerance_bp(inv.get("vendor_recurring_tolerance_pct"))
    detail.update({
        "prior_invoice_id": str(prior["id"]), "prior_gross_cents": str(prior_gross),
        "gross_cents": str(gross), "variance_cents": str(variance), "tolerance_bp": tol_bp,
    })
    if prior_gross <= 0 or abs(variance) * 10_000 > prior_gross * tol_bp:
        return no("outside_tolerance")
    same_period = conn.execute(
        """select id from invoice where vendor_id = %s and id <> %s and state in ('APPROVED', 'ENTERED', 'PENDING_APPROVAL')
           and invoice_date is not null and date_trunc('month', invoice_date) = date_trunc('month', %s::date) limit 1""",
        (inv["vendor_id"], inv["id"], inv["invoice_date"]),
    ).fetchone()
    if same_period:
        return no("period_already_billed", other_invoice_id=str(same_period["id"]))
    detail["reason"] = "ok"
    return True, detail
