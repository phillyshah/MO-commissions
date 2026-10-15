"""Which vendor record does this invoice belong to? `match_vendor(conn, ...) -> MatchResult`.

Tiers, strongest first; the first hit wins and the method is recorded on the invoice so an approver
can see WHY the portal thinks it is this vendor:
  domain_exact  — the sender address itself is an approved vendor_domain pattern
  domain        — the sender's domain is an approved vendor_domain pattern (subdomains included)
  platform      — the sender is a billing platform AND the platform's account-id regex finds an id
                  that billing_platform_map ties to one vendor. Without the id it is NOT a match: the
                  platform domain alone proves nothing (spec, billing-platform allowlist)
  tin           — the extracted TIN equals a vendor's decrypted TIN
  remit_to      — the normalised remit-to address equals a vendor's
  name          — name_key similarity ≥ 0.92 (difflib). Last because an attacker chooses the name
Unknown → MatchResult(None, "none", 0.0); the pipeline parks the invoice in VENDOR_SETUP.

`message_text` is an addition to the contract signature (keyword, defaulted) because the platform
tier needs the body to find the account id.
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from uuid import UUID

import psycopg

from accounting.identity import crypto
from accounting.vendors.normalize import digits, email_domain, name_key, remit_key

NAME_RATIO = 0.92


@dataclass
class MatchResult:
    vendor_id: UUID | None
    method: str
    confidence: float
    detail: dict = field(default_factory=dict)

    def as_dict(self) -> dict:
        return {"vendor_id": str(self.vendor_id) if self.vendor_id else None, "method": self.method,
                "confidence": self.confidence, **self.detail}


def platform_for_domain(conn: psycopg.Connection, domain: str) -> dict | None:
    if not domain:
        return None
    rows = conn.execute("select * from billing_platform").fetchall()
    for p in rows:
        for d in p["sending_domains"] or []:
            d = d.lower()
            if domain == d or domain.endswith("." + d):
                return p
    return None


def platform_account_match(conn: psycopg.Connection, platform: dict, text: str) -> tuple[dict | None, str | None]:
    """(billing_platform_map row, account_id) when the regex finds an id that is mapped and approved."""
    rx = platform.get("account_id_regex")
    if not rx or not text:
        return None, None
    try:
        pattern = re.compile(rx)
    except re.error:
        return None, None
    for m in pattern.finditer(text):
        account_id = (m.group(1) if m.groups() else m.group(0)).strip()
        row = conn.execute(
            "select * from billing_platform_map where platform_id = %s and lower(account_id) = lower(%s) and approved_by is not null",
            (platform["id"], account_id),
        ).fetchone()
        if row:
            return row, account_id
    return None, None


def match_vendor(conn: psycopg.Connection, *, sender_addr: str | None, vendor_name: str | None, tin: str | None,
                 remit_to: dict | None, message_text: str = "") -> MatchResult:
    addr = (sender_addr or "").strip().lower()
    domain = email_domain(addr)

    if addr:
        row = conn.execute(
            "select vendor_id from vendor_domain where status = 'approved' and pattern = %s limit 1", (addr,)
        ).fetchone()
        if row:
            return MatchResult(row["vendor_id"], "domain_exact", 1.0, {"pattern": addr})
    if domain:
        rows = conn.execute("select vendor_id, pattern from vendor_domain where status = 'approved' and pattern not like '%%@%%'").fetchall()
        hits = {r["vendor_id"] for r in rows if domain == r["pattern"] or domain.endswith("." + r["pattern"])}
        if len(hits) == 1:
            return MatchResult(hits.pop(), "domain", 0.98, {"pattern": domain})
        platform = platform_for_domain(conn, domain)
        if platform:
            mapped, account_id = platform_account_match(conn, platform, message_text or "")
            if mapped:
                return MatchResult(mapped["vendor_id"], "platform", 0.97,
                                   {"platform": platform["id"], "account_id": account_id})
    tin_digits = digits(tin)
    if len(tin_digits) >= 9:
        for v in conn.execute("select id, tin_enc from vendor where tin_enc is not null").fetchall():
            try:
                if digits(crypto.decrypt(v["tin_enc"], aad=str(v["id"]).encode()).decode()) == tin_digits:
                    return MatchResult(v["id"], "tin", 0.97, {})
            except Exception:
                continue
    rk = remit_key(remit_to)
    if rk and "|" in rk and all(rk.split("|")):
        for v in conn.execute("select id, remit_to from vendor where remit_to is not null").fetchall():
            if remit_key(v["remit_to"]) == rk:
                return MatchResult(v["id"], "remit_to", 0.9, {"remit_key": rk})
    key = name_key(vendor_name)
    if key:
        best, best_ratio = None, 0.0
        for v in conn.execute("select id, name_key from vendor where status <> 'blocked'").fetchall():
            ratio = difflib.SequenceMatcher(None, key, v["name_key"] or "").ratio()
            if ratio > best_ratio:
                best, best_ratio = v, ratio
        if best is not None and best_ratio >= NAME_RATIO:
            return MatchResult(best["id"], "name", round(best_ratio, 3), {"name_key": key})
    return MatchResult(None, "none", 0.0, {"name_key": key, "domain": domain})
