"""Deterministic validators (CAP-2). They decide; the model only proposed.

`validate(invoice_row, extracted) -> [Failure]`. An empty list means CAPTURED; anything else means
AP_REVIEW with the failing fields in the event payload. `invoice_row` is the invoice as stored plus,
when a vendor was matched, the vendor row under key "vendor" (po_required, country drive two checks).

Money is Decimal → integer cents here and nowhere is a float involved: a JSON number from the model
has already been turned into a string by the extractor.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation

from accounting.config import settings

REQUIRED = ("vendor_name", "invoice_number", "invoice_date", "gross", "currency")
_CURRENCY = re.compile(r"^[A-Z]{3}$")
US = {"US", "USA", "UNITED STATES", "UNITED STATES OF AMERICA"}


@dataclass(frozen=True)
class Failure:
    code: str
    field: str | None
    text: str


def to_cents(value) -> int | None:
    """"1234.56" -> 123456. None for anything that is not an amount. Never a float."""
    if value is None or value == "":
        return None
    try:
        d = Decimal(str(value).replace(",", "").replace("$", "").strip())
    except InvalidOperation:
        return None
    return int((d * 100).quantize(Decimal("1")))


def parse_date(value) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def _val(extracted: dict, f: str):
    cell = extracted.get(f) or {}
    return cell.get("value")


def _conf(extracted: dict, f: str) -> float:
    cell = extracted.get(f) or {}
    try:
        return float(cell.get("confidence") or 0.0)
    except (TypeError, ValueError):
        return 0.0


def validate(invoice_row: dict, extracted: dict) -> list[Failure]:
    out: list[Failure] = []
    today = date.today()
    threshold = settings.capture_confidence_threshold

    for f in REQUIRED:
        if _val(extracted, f) in (None, ""):
            out.append(Failure("missing_field", f, f"{f.replace('_', ' ')} could not be read"))
    low = [f for f in REQUIRED if _val(extracted, f) not in (None, "") and _conf(extracted, f) < threshold]
    for f in low:
        out.append(Failure("low_confidence", f, f"{f.replace('_', ' ')} was read with low confidence"))

    gross = to_cents(_val(extracted, "gross"))
    tax = to_cents(_val(extracted, "tax"))
    line_total = to_cents(_val(extracted, "line_total"))
    if gross is not None and gross <= 0:
        out.append(Failure("amount_invalid", "gross", "the total is zero or negative"))
    lines = _val(extracted, "lines") or []
    line_sum = None
    if lines:
        amounts = [to_cents(ln.get("amount")) for ln in lines if isinstance(ln, dict)]
        if amounts and all(a is not None for a in amounts):
            line_sum = sum(amounts)
    if line_total is not None and line_sum is not None and abs(line_total - line_sum) > 1:
        out.append(Failure("arithmetic", "line_total", "the line items do not add up to the subtotal"))
    if gross is not None and line_total is not None and tax is not None and abs(line_total + tax - gross) > 1:
        out.append(Failure("arithmetic", "gross", "subtotal plus tax does not equal the total"))
    elif gross is not None and line_total is None and line_sum is not None and tax is not None and abs(line_sum + tax - gross) > 1:
        out.append(Failure("arithmetic", "gross", "line items plus tax do not equal the total"))

    inv_date = parse_date(_val(extracted, "invoice_date"))
    due = parse_date(_val(extracted, "due_date"))
    if _val(extracted, "invoice_date") and inv_date is None:
        out.append(Failure("date_invalid", "invoice_date", "the invoice date is not a valid date"))
    if inv_date is not None:
        if inv_date > today + timedelta(days=1):
            out.append(Failure("date_future", "invoice_date", "the invoice date is in the future"))
        if inv_date < today - timedelta(days=366):
            out.append(Failure("date_old", "invoice_date", "the invoice is more than a year old"))
        if due is not None and due < inv_date:
            out.append(Failure("date_order", "due_date", "the due date is before the invoice date"))

    currency = _val(extracted, "currency")
    if currency and not _CURRENCY.match(str(currency)):
        out.append(Failure("currency_invalid", "currency", "the currency is not a three-letter code"))
    vendor = invoice_row.get("vendor") or {}
    vcountry = (vendor.get("country") or "").strip().upper()
    if currency and vcountry in US and currency != "USD":
        out.append(Failure("currency_vendor", "currency", f"{currency} does not match a US vendor"))
    if vendor.get("po_required") and not _val(extracted, "po_number"):
        out.append(Failure("missing_po", "po_number", "this vendor requires a PO number and none was found"))
    return out
