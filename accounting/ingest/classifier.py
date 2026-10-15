"""INT-2: what kind of message is this? Deterministic keyword and structure rules, in a fixed order.

Order matters and is the whole design: a bank-change request often *also* contains the word
"invoice", and a statement often lists invoices, so the more specific class is tested first.
  bank_change  change-wording next to bank/remittance wording ("updated banking details", "our
               account number has changed"). An invoice that merely prints wire instructions does
               not match: the rules need the CHANGE verb.
  statement    "statement of account", "account statement", a subject or filename starting "statement"
  credit_memo  "credit memo/note/memorandum", filename containing "credit"
  invoice      invoice wording, "amount due", "bill", or a PDF/image attachment with no other class
  inquiry      questions about status/payment/W-9/remittance advice with nothing attached
  spam_other   none of the above
The result carries the rule that fired so a wrong class can be explained from the message row. The
model is never consulted here (CAP-5: a classifier that can be talked into "invoice" is not one).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

_CHANGE = r"(?:new|updated?|chang(?:e|ed|ing)|revis(?:e|ed)|switch(?:ed)?|replace(?:d)?|moved?|different|effective immediately)"
_BANKY = r"(?:bank(?:ing)?|account|wire|ach|remit(?:tance)?|payment|routing|iban|swift|beneficiary)"
BANK_CHANGE = [
    re.compile(rf"\b{_CHANGE}\b[^.\n]{{0,60}}\b{_BANKY}\b[^.\n]{{0,40}}\b(?:details?|information|instructions?|number|account)\b", re.I),
    re.compile(rf"\b{_BANKY}\b[^.\n]{{0,40}}\b(?:details?|information|instructions?|number)\b[^.\n]{{0,60}}\b(?:has|have|were|was|is|are)\b[^.\n]{{0,20}}\b(?:changed|updated|revised|moved)\b", re.I),
    re.compile(r"\bplease (?:update|change|use) (?:our|the|your records with)[^.\n]{0,40}\b(?:bank|account|wire|ach|remit)", re.I),
    re.compile(r"\b(?:update|change) of (?:bank|banking|remittance|payment) (?:details|information|instructions)\b", re.I),
]
STATEMENT = [
    re.compile(r"\b(?:statement of account|account statement|monthly statement|customer statement|open items? statement|aging statement)\b", re.I),
    re.compile(r"^\s*(?:re:|fw:|fwd:)?\s*statement\b", re.I),
]
CREDIT = [re.compile(r"\bcredit (?:memo(?:randum)?|note|adjustment)\b", re.I), re.compile(r"\bcm[-_ ]?\d{3,}\b")]
INVOICE = [
    re.compile(r"\binvoice\b", re.I), re.compile(r"\binv(?:oice)?\s*(?:#|no\.?|number)", re.I),
    re.compile(r"\b(?:amount|balance|payment|total) due\b", re.I), re.compile(r"\b(?:your|new|attached) bill\b", re.I),
    re.compile(r"\bbill(?:ing)? (?:statement|summary|notice)\b", re.I), re.compile(r"\bremit payment\b", re.I),
]
INQUIRY = [
    re.compile(r"\b(?:status of|has .{0,30} been paid|when (?:will|can) .{0,30} (?:paid|payment)|payment status|remittance advice|w-?9|vendor (?:form|setup|registration)|question about|following up|follow-up)\b", re.I),
    re.compile(r"\?\s*$", re.M),
]
INVOICE_FILENAME = re.compile(r"\b(?:inv(?:oice)?|bill|rechnung|factura|facture)\b", re.I)
STATEMENT_FILENAME = re.compile(r"\bstatement\b", re.I)
CREDIT_FILENAME = re.compile(r"\bcredit\b|\bcm\d", re.I)
DOC_EXT = re.compile(r"\.(pdf|jpe?g|png|tiff?|webp|heic)$", re.I)


@dataclass
class Classification:
    label: str
    rule: str
    detail: dict = field(default_factory=dict)

    def as_dict(self) -> dict:
        return {"label": self.label, "rule": self.rule, **self.detail}


def _hits(patterns, text) -> list[str]:
    return [m.group(0)[:60] for p in patterns for m in [p.search(text)] if m]


def classify(subject: str | None, body: str | None, attachment_names: list[str] | None) -> Classification:
    subject = subject or ""
    body = (body or "")[:50_000]
    names = [n for n in (attachment_names or []) if n]
    text = f"{subject}\n{body}"
    doc_attached = any(DOC_EXT.search(n) for n in names)

    hits = _hits(BANK_CHANGE, text)
    if hits:
        return Classification("bank_change", "bank_change_wording", {"matched": hits})
    hits = _hits(STATEMENT, subject) or _hits(STATEMENT, body[:2000])
    if hits or any(STATEMENT_FILENAME.search(n) and not INVOICE_FILENAME.search(n) for n in names):
        return Classification("statement", "statement_wording", {"matched": hits or [n for n in names if STATEMENT_FILENAME.search(n)]})
    hits = _hits(CREDIT, text)
    if hits or any(CREDIT_FILENAME.search(n) for n in names):
        return Classification("credit_memo", "credit_wording", {"matched": hits or [n for n in names if CREDIT_FILENAME.search(n)]})
    # A question ABOUT an invoice says "invoice" too, so with no document attached the inquiry rule runs first.
    inq = _hits(INQUIRY, text)
    if inq and not doc_attached:
        return Classification("inquiry", "inquiry_wording", {"matched": inq})
    hits = _hits(INVOICE, text)
    if hits:
        return Classification("invoice", "invoice_wording", {"matched": hits})
    if any(INVOICE_FILENAME.search(n) for n in names):
        return Classification("invoice", "invoice_filename", {"matched": [n for n in names if INVOICE_FILENAME.search(n)]})
    if doc_attached:
        return Classification("invoice", "document_attached", {"attachments": names})
    return Classification("spam_other", "no_rule_matched", {})
