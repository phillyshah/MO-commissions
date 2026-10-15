"""`screen(conn, invoice_id) -> Screen(hard, soft, detail)` — the suspicious-invoice screen.

Hard triggers, exactly the spec's list, each with a code the Controller's screen can explain:
  lookalike_domain           sender domain is a lookalike of a vendor's known domain
  display_name_mismatch      display name matches a known vendor; address not on its list / not a mapped platform
  dmarc_fail                 DMARC failed and the domain's policy is quarantine/reject (read from the
                             Microsoft 365 Authentication-Results we trust, never looked up live)
  bank_mismatch              bank fingerprint differs from the vendor's
  remit_to_mismatch          remit-to address (street/postal) differs from the vendor's
  bank_change_request        the message asks to change bank details (classifier)
  bank_changed_recently      the vendor's bank fingerprint changed in the last 30 days
  duplicate_number_mismatch  same invoice number as an earlier invoice, different amount or bank
  payee_person / payee_crypto / payee_gift_card / payee_country
  risky_file                 the attachment is a risky type
  hidden_text                the parser flagged hidden text or AI instructions
  direct_sender              the message did not come through the Maxx forwarding path
Soft signals (banner only, behind config_current.soft_signals_enabled): urgency_language, unusual_hour,
new_contact, amount_above_history, round_amount, new_remit_wording, pdf_metadata_mismatch, plus the
DUP-1 flags. A soft signal never lowers a hard trigger, and a model never adds to or removes from either
list — everything here is computed from stored rows.

The screen is called before extraction (sender- and file-level triggers can save an API call) and
again after vendor matching (bank, remit-to, duplicate). Whatever it finds, it acts: a hard trigger
opens a quarantine_case and transitions to QUARANTINED in the same transaction, so a caller cannot
forget to.
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from uuid import UUID
from zoneinfo import ZoneInfo

import psycopg

from accounting.config import settings
from accounting.config_flags import config_flag
from accounting.lifecycle import transition
from accounting.screening import duplicates
from accounting.vendors import fingerprint as fp
from accounting.vendors.lookalike import is_lookalike
from accounting.vendors.match import platform_account_match, platform_for_domain
from accounting.vendors.normalize import email_domain, looks_like_person, name_key, remit_key, remit_wording_key

URGENCY = re.compile(r"\b(urgent(?:ly)?|immediately|asap|right away|final (?:notice|reminder|demand)|past due|overdue|"
                     r"within (?:24|48) hours|today only|last chance|suspend(?:ed|sion)|legal action|act now|time.sensitive)\b", re.I)
CRYPTO = re.compile(r"\b(bitcoin|btc|ethereum|usdt|tether|crypto(?:currency)?|wallet address|coinbase|binance)\b|\b0x[a-fA-F0-9]{40}\b|\bbc1[a-z0-9]{25,}\b", re.I)
GIFT_CARD = re.compile(r"\b(gift ?cards?|itunes|google play|steam cards?|prepaid cards?|vanilla visa|razer gold)\b", re.I)
US = {"US", "USA", "UNITED STATES", "UNITED STATES OF AMERICA"}
NAME_RATIO = 0.92
QUARANTINABLE = ("RECEIVED", "CAPTURED", "AP_REVIEW")


@dataclass
class Screen:
    hard: list[str] = field(default_factory=list)
    soft: list[str] = field(default_factory=list)
    detail: dict = field(default_factory=dict)


def _country(c: str | None) -> str:
    c = (c or "").strip().upper()
    return "US" if c in US else c


def _load(conn, invoice_id):
    inv = conn.execute("select * from invoice where id = %s for update", (invoice_id,)).fetchone()
    if inv is None:
        return None, None, None, []
    msg = conn.execute("select * from message where id = %s", (inv["message_id"],)).fetchone() if inv["message_id"] else None
    vendor = conn.execute("select * from vendor where id = %s", (inv["vendor_id"],)).fetchone() if inv["vendor_id"] else None
    atts = conn.execute("select * from attachment where message_id = %s", (inv["message_id"],)).fetchall() if inv["message_id"] else []
    return inv, msg, vendor, atts


def _sender(msg: dict | None) -> tuple[str, str]:
    """The address the screen judges (spec: forwarded invoices are judged by the ORIGINAL sender)."""
    if not msg:
        return "", ""
    addr = (msg.get("original_sender") or msg.get("from_addr") or "").lower()
    return addr, email_domain(addr)


def _sender_triggers(conn, inv, msg, vendor, s: Screen) -> None:
    addr, domain = _sender(msg)
    if msg and msg["trust"] == "direct":
        s.hard.append("direct_sender")
    if not domain:
        return
    known = [r["pattern"] for r in conn.execute(
        "select pattern from vendor_domain where status = 'approved' and pattern not like '%%@%%'").fetchall()]
    if domain not in known:
        look, why = is_lookalike(domain, known)
        if look:
            s.hard.append("lookalike_domain")
            s.detail["lookalike_domain"] = {"sender": domain, "reason": why}

    # Display name says a vendor we know; address is not on that vendor's list and not a mapped platform.
    display = name_key(msg.get("from_name")) if msg else ""
    claimed = []
    if display:
        for v in conn.execute("select id, legal_name, name_key from vendor").fetchall():
            if v["name_key"] and difflib.SequenceMatcher(None, display, v["name_key"]).ratio() >= NAME_RATIO:
                claimed.append(v)
    for v in claimed:
        allowed = conn.execute(
            "select pattern from vendor_domain where vendor_id = %s and status = 'approved'", (v["id"],)).fetchall()
        ok = any(p["pattern"] == addr or domain == p["pattern"] or domain.endswith("." + p["pattern"]) for p in allowed)
        if not ok:
            platform = platform_for_domain(conn, domain)
            if platform:
                mapped, _ = platform_account_match(conn, platform, (msg or {}).get("body_text") or "")
                ok = bool(mapped and str(mapped["vendor_id"]) == str(v["id"]))
        if not ok:
            s.hard.append("display_name_mismatch")
            s.detail["display_name_mismatch"] = {"display_name": msg.get("from_name"), "claims_vendor": v["legal_name"], "address": addr}
            break

    auth = (msg or {}).get("auth_results") or {}
    dmarc = (auth.get("dmarc") or {})
    if str(dmarc.get("result", "")).lower() == "fail" and str(dmarc.get("action", "")).lower() in ("quarantine", "reject", "oreject"):
        s.hard.append("dmarc_fail")
        s.detail["dmarc_fail"] = dmarc


def _file_triggers(inv, atts, s: Screen) -> None:
    for a in atts:
        if a["risky"] and (inv["attachment_id"] is None or str(a["id"]) == str(inv["attachment_id"])):
            s.hard.append("risky_file")
            s.detail["risky_file"] = {"filename": a["filename"], "reason": a["risk_reason"]}
            break
    flags = (inv.get("parse_detail") or {}).get("hidden_text_flags") or []
    if flags:
        s.hard.append("hidden_text")
        s.detail["hidden_text"] = flags[:10]


def _bank_triggers(conn, inv, msg, vendor, s: Screen) -> None:
    if msg and msg.get("classification") == "bank_change":
        s.hard.append("bank_change_request")
    if vendor is None:
        return
    cur = fp.current(conn, vendor["id"])
    if cur is not None and inv.get("bank_fingerprint") is not None and bytes(cur["fingerprint"]) != bytes(inv["bank_fingerprint"]):
        s.hard.append("bank_mismatch")
        s.detail["bank_mismatch"] = {"on_file_last4": cur["last4"], "invoice_last4": inv.get("bank_last4")}
    elif cur is not None and inv.get("bank_last4") and inv.get("bank_fingerprint") is None and inv["bank_last4"] != cur["last4"]:
        # Only the last four were printed; a different last four is still a different account.
        s.hard.append("bank_mismatch")
        s.detail["bank_mismatch"] = {"on_file_last4": cur["last4"], "invoice_last4": inv["bank_last4"]}
    prev = fp.previous(conn, vendor["id"])
    if prev is not None and prev["superseded_at"] is not None:
        # Judged at screening time, not receipt time: the change may have been recorded after the mail arrived.
        age = max(datetime.now(timezone.utc), inv["received_at"]) - prev["superseded_at"]
        if age.days < 30:
            s.hard.append("bank_changed_recently")
            s.detail["bank_changed_recently"] = {"changed_at": prev["superseded_at"].isoformat(), "days_ago": age.days}
    vr, ir = vendor.get("remit_to"), inv.get("remit_to")
    if vr and ir:
        vk, ik = remit_key(vr), remit_key(ir)
        if vk and ik and vk != ik:
            s.hard.append("remit_to_mismatch")
            s.detail["remit_to_mismatch"] = {"on_file": vr, "invoice": ir}
        elif remit_wording_key(vr) != remit_wording_key(ir):
            s.soft.append("new_remit_wording")


def _payee_triggers(conn, inv, msg, vendor, s: Screen) -> None:
    remit = inv.get("remit_to") or {}
    payee = remit.get("name") or ""
    if payee and looks_like_person(payee) and (vendor is None or name_key(payee) != vendor["name_key"]):
        s.hard.append("payee_person")
        s.detail["payee_person"] = {"payee": payee}
    text = " ".join(str(v) for v in remit.values() if v) + " " + ((msg or {}).get("body_text") or "")[:20000]
    if CRYPTO.search(text):
        s.hard.append("payee_crypto")
    if GIFT_CARD.search(text):
        s.hard.append("payee_gift_card")
    if vendor and remit.get("country") and vendor.get("country") and _country(remit["country"]) != _country(vendor["country"]):
        s.hard.append("payee_country")
        s.detail["payee_country"] = {"vendor": vendor["country"], "invoice": remit["country"]}


def _soft_signals(conn, inv, msg, vendor, s: Screen) -> None:
    body = ((msg or {}).get("body_text") or "") + " " + ((msg or {}).get("subject") or "")
    if URGENCY.search(body):
        s.soft.append("urgency_language")
    received = inv["received_at"]
    if received is not None:
        local = received.astimezone(ZoneInfo(settings.timezone)) if received.tzinfo else received.replace(tzinfo=timezone.utc).astimezone(ZoneInfo(settings.timezone))
        if local.hour < 6 or local.hour >= 22 or local.weekday() >= 5:
            s.soft.append("unusual_hour")
    gross = inv.get("gross_cents")
    if gross and gross >= 100_000 and gross % 10_000 == 0:
        s.soft.append("round_amount")
    if vendor is not None:
        addr, _ = _sender(msg)
        if addr and inv["message_id"]:
            seen = conn.execute(
                """select 1 from invoice i join message m on m.id = i.message_id
                   where i.vendor_id = %s and i.id <> %s and lower(coalesce(m.original_sender, m.from_addr)) = %s limit 1""",
                (vendor["id"], inv["id"], addr)).fetchone()
            prior = conn.execute("select 1 from invoice where vendor_id = %s and id <> %s limit 1", (vendor["id"], inv["id"])).fetchone()
            if prior and not seen:
                s.soft.append("new_contact")
        if gross:
            hist = conn.execute(
                """select max(gross_cents) as mx, count(*) as n from invoice where vendor_id = %s and id <> %s
                   and state in ('CAPTURED','PENDING_APPROVAL','QUESTIONED','APPROVED','ENTERED') and gross_cents is not null""",
                (vendor["id"], inv["id"])).fetchone()
            if hist and hist["n"] >= 3 and hist["mx"] and gross > 3 * hist["mx"]:
                s.soft.append("amount_above_history")
                s.detail["amount_above_history"] = {"gross_cents": gross, "prior_max_cents": hist["mx"]}
    meta = (inv.get("parse_detail") or {}).get("metadata") or {}
    author = name_key(meta.get("Author") or meta.get("Creator"))
    if author and vendor is not None and author != vendor["name_key"]:
        for v in conn.execute("select name_key from vendor where id <> %s", (vendor["id"],)).fetchall():
            if v["name_key"] and difflib.SequenceMatcher(None, author, v["name_key"]).ratio() >= NAME_RATIO:
                s.soft.append("pdf_metadata_mismatch")
                s.detail["pdf_metadata_mismatch"] = {"author": meta.get("Author") or meta.get("Creator")}
                break


def evaluate(conn: psycopg.Connection, invoice_id: UUID | str) -> Screen:
    """Compute without acting. `screen()` is this plus the quarantine."""
    inv, msg, vendor, atts = _load(conn, invoice_id)
    s = Screen()
    if inv is None:
        return s
    _sender_triggers(conn, inv, msg, vendor, s)
    _file_triggers(inv, atts, s)
    _bank_triggers(conn, inv, msg, vendor, s)
    _payee_triggers(conn, inv, msg, vendor, s)
    dup = duplicates.check(conn, inv)
    s.hard.extend(dup.hard)
    s.detail.update(dup.detail)
    if config_flag(conn, "soft_signals_enabled"):
        _soft_signals(conn, inv, msg, vendor, s)
        s.soft.extend(dup.soft)
    s.hard = list(dict.fromkeys(s.hard))
    s.soft = list(dict.fromkeys(s.soft))
    return s


def screen(conn: psycopg.Connection, invoice_id: UUID | str) -> Screen:
    s = evaluate(conn, invoice_id)
    inv = conn.execute("select state, message_id, soft_signals from invoice where id = %s", (invoice_id,)).fetchone()
    if inv is None:
        return s
    conn.execute("update invoice set soft_signals = %s where id = %s", (s.soft, invoice_id))
    if s.hard and inv["state"] in QUARANTINABLE:
        conn.execute(
            "insert into quarantine_case (invoice_id, message_id, triggers, detail) values (%s, %s, %s, %s::jsonb)",
            (invoice_id, inv["message_id"], s.hard, _json(s.detail)),
        )
        transition(conn, invoice_id, "QUARANTINED", reason_code="hard_trigger", actor_type="system",
                   payload={"triggers": s.hard, "soft": s.soft, "detail": s.detail})
    return s


def _json(obj) -> str:
    import json

    from accounting.lifecycle import _default

    return json.dumps(obj, default=_default)
