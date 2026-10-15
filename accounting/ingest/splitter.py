"""INT-3: one invoice row per invoice, not per message.

Default: one invoice per attachment (a PDF, an image). A PDF whose text layer shows MORE THAN ONE
distinct invoice number across its pages is split into page ranges, one per number, and every piece
is marked `split_confirmed=False` so AP confirms the boundaries before anything routes. Pages with no
number of their own belong to the invoice before them (page 2 of a two-page invoice rarely repeats the
header). A message with no document attachment yields one link-only invoice so the message is still
tracked and AP can chase the file.

The text comes from the parser (no rendering) — the splitter never opens the PDF itself (INT-5).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from accounting.ingest.classifier import DOC_EXT
from accounting.parse import client as parser
from accounting.vendors.normalize import invoice_number_norm

INVOICE_NO = re.compile(
    r"\b(?:invoice|inv\.?|bill|factura|rechnung)\s*(?:#|no\.?|nbr\.?|num(?:ber)?\.?|id)?\s*[:#\-]?\s*([A-Z0-9][A-Z0-9\-/]{2,24})",
    re.I)
_NOT_NUMBERS = {"DATE", "NUMBER", "TOTAL", "NO", "FOR", "FROM", "TO", "DUE", "AMOUNT", "COPY", "ORIGINAL", "PERIOD"}


@dataclass
class Unit:
    attachment: dict | None             # the stored attachment row (None for link-only)
    split_index: int | None
    page_from: int | None
    page_to: int | None
    split_confirmed: bool
    detail: dict


def numbers_on_page(text: str) -> list[str]:
    out = []
    for m in INVOICE_NO.finditer(text or ""):
        raw = m.group(1).strip(".:,-")
        norm = invoice_number_norm(raw)
        if len(norm) < 3 or norm in _NOT_NUMBERS or norm.isalpha():
            continue
        out.append(norm)
    return out


def page_groups(page_texts: list[str]) -> list[tuple[int, int, str | None]]:
    """[(page_from, page_to, number)] 1-based inclusive. Pages inherit the previous page's number."""
    groups: list[list] = []
    current: str | None = None
    for i, text in enumerate(page_texts, start=1):
        nums = numbers_on_page(text)
        number = nums[0] if nums else None   # the first number printed is the page's own; later ones are references
        if number is not None and number != current:
            current = number
            groups.append([i, i, current])
        elif groups:
            groups[-1][1] = i
        else:
            groups.append([i, i, None])
    return [(g[0], g[1], g[2]) for g in groups]


def split(attachments: list[dict], data_for) -> list[Unit]:
    """`attachments` are stored attachment rows (dicts with id, filename, content_type, risky);
    `data_for(att) -> bytes` fetches the bytes. Risky and non-document attachments are one unit each
    (the pipeline decides what to do with them); documents may split."""
    units: list[Unit] = []
    docs = [a for a in attachments if not a.get("inline")]
    for a in docs:
        name = a.get("filename") or ""
        ct = (a.get("content_type") or "").lower()
        is_pdf = ct == "application/pdf" or name.lower().endswith(".pdf")
        if a.get("risky") or not is_pdf:
            units.append(Unit(a, None, None, None, True, {"rule": "one_per_attachment"}))
            continue
        try:
            res = parser.page_texts(data_for(a), ct or "application/pdf", name)
        except Exception as e:  # unparseable PDFs still become one invoice; the pipeline reports why
            units.append(Unit(a, None, None, None, True, {"rule": "one_per_attachment", "note": str(e)[:200]}))
            continue
        groups = page_groups(res.page_texts)
        distinct = {g[2] for g in groups if g[2]}
        if len(distinct) > 1 and len(groups) > 1:
            for idx, (pf, pt, num) in enumerate(groups):
                units.append(Unit(a, idx, pf, pt, False, {"rule": "multi_invoice_pdf", "invoice_number_norm": num,
                                                          "page_count": res.page_count}))
        else:
            units.append(Unit(a, None, None, None, True, {"rule": "one_per_attachment", "page_count": res.page_count}))
    if not units:
        units.append(Unit(None, None, None, None, True, {"rule": "no_attachment"}))
    return units


def is_document_name(name: str | None) -> bool:
    return bool(name) and bool(DOC_EXT.search(name))
