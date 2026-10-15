"""The approved queue (SAP-1): what AP keys into SAP, oldest first, and the one field that ends
the cycle clock — the SAP document number.

Money is integer cents in the database and a decimal STRING in the CSV ("1234.56"), never a float,
so what the clerk pastes into SAP is exactly what was approved.
"""
from __future__ import annotations

import csv
import io
from datetime import datetime, timezone

import psycopg

from accounting.lifecycle import transition
from accounting.web.templating import CONFIDENTIAL_FOOTER

CSV_COLUMNS = ["portal_number", "vendor", "sap_vendor_number", "invoice_number", "invoice_date", "due_date", "currency",
               "gross", "tax", "po_number", "gl", "cost_center", "approved_at", "approvers"]

SAP_NUMBER_MAX = 20


def cents_text(cents: int | None) -> str:
    if cents is None:
        return ""
    sign = "-" if cents < 0 else ""
    cents = abs(int(cents))
    return f"{sign}{cents // 100}.{cents % 100:02d}"


def approved_rows(conn: psycopg.Connection) -> list[dict]:
    return conn.execute(
        """select i.id, i.portal_number, i.invoice_number, i.invoice_date, i.due_date, i.currency, i.gross_cents, i.tax_cents,
                  i.po_number, i.approved_at, i.received_at, i.tags,
                  v.legal_name as vendor, v.sap_vendor_number, v.default_gl as gl, v.default_cost_center as cost_center,
                  coalesce((select string_agg(distinct u.display_name, '; ' order by u.display_name)
                            from approval_signature s join ap_user u on u.id = s.signer_id
                            where s.invoice_id = i.id and s.status = 'valid' and s.decision = 'approve'), '') as approvers,
                  (select count(*) from approval_signature s where s.invoice_id = i.id and s.status = 'valid' and s.decision = 'approve') as signature_count
           from invoice i left join vendor v on v.id = i.vendor_id
           where i.state = 'APPROVED'
           order by i.approved_at nulls first, i.received_at"""
    ).fetchall()


def csv_text(rows: list[dict], as_of: datetime | None = None) -> str:
    """The whole file as text. The last line is the confidential footer as a comment row so a
    printed or forwarded export still carries it."""
    as_of = as_of or datetime.now(timezone.utc)
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(CSV_COLUMNS)
    for r in rows:
        w.writerow([
            f"AP-{int(r['portal_number']):06d}", r["vendor"] or "", r["sap_vendor_number"] or "", r["invoice_number"] or "",
            r["invoice_date"].isoformat() if r["invoice_date"] else "", r["due_date"].isoformat() if r["due_date"] else "",
            r["currency"] or "", cents_text(r["gross_cents"]), cents_text(r["tax_cents"]), r["po_number"] or "",
            r["gl"] or "", r["cost_center"] or "", r["approved_at"].isoformat() if r["approved_at"] else "",
            r["approvers"] or ("auto-released" if not r.get("signature_count") else ""),
        ])
    w.writerow([f"# As of {as_of.isoformat(timespec='seconds')}"])
    w.writerow([f"# {CONFIDENTIAL_FOOTER}"])
    return buf.getvalue()


def csv_filename(as_of: datetime | None = None) -> str:
    as_of = as_of or datetime.now(timezone.utc)
    return f"MaxxDash-approved-invoices-{as_of:%Y-%m-%d}.csv"


class SapNumberError(ValueError):
    """Raised with a sentence fit for the screen."""


def record_sap_document(conn: psycopg.Connection, invoice_id, number: str, by) -> dict:
    number = (number or "").strip()
    if not number:
        raise SapNumberError("Enter the SAP document number.")
    if len(number) > SAP_NUMBER_MAX:
        raise SapNumberError(f"The SAP document number is at most {SAP_NUMBER_MAX} characters.")
    conn.execute("update invoice set sap_document_number = %s, entered_by = %s where id = %s", (number, by, invoice_id))
    return transition(conn, invoice_id, "ENTERED", reason_code="sap_entered", actor_type="user", actor_id=by,
                      payload={"sap_document_number": number})
