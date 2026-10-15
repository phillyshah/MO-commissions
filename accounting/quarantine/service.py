"""Quarantine decisions (QUA-1..5). The Controller's three verbs: Release, Reject, Report.

Release is the dangerous one, so it carries the controls: a sender or bank case needs a callback to a
number the email did not supply AND a second person; any release above the second-person amount needs
a second person too. The second person confirms on the same screen — the release is parked on the
case (release_pending_at) until they do, and the first reviewer cannot be the confirmer.

A released invoice is tagged 'released_from_quarantine' and auto-release checks that tag (QUA-3):
whatever the rule, a once-suspicious invoice goes through the full chain. Package A's pipeline, if it
re-runs eligibility, must respect the same tag.
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
from uuid import UUID

import psycopg

from accounting.config import settings
from accounting.lifecycle import log_event, transition
from accounting.routing._days import bd, utcnow
from accounting.routing.common import load_invoice, notify_queue, user_row, users_with_role

# Plain-English reasons, one per trigger code Package A's screening can raise.
TRIGGER_TEXT: dict[str, str] = {
    "direct_sender": "This email was sent straight to the portal mailbox instead of coming through the company's own mail system. "
                     "Genuine vendor mail always arrives through the company's forwarding path.",
    "lookalike_domain": "The sender's email domain looks almost — but not quite — like a known vendor's domain. "
                        "Swapped letters and added words are a common way to impersonate a vendor.",
    "unlisted_sender": "The sender's name matches a known vendor, but the email address is not one that vendor is known to use "
                       "and is not a recognised billing platform.",
    "dmarc_fail": "The sender's own domain says mail like this should be rejected, and this message failed that check. "
                  "It may not have come from where it claims.",
    "bank_mismatch": "The bank details or remit-to address on this invoice are different from the ones on file for this vendor.",
    "bank_change_request": "The message asks to change where payments are sent. Bank changes are never accepted by email.",
    "bank_changed_30d": "This vendor's bank details changed within the last 30 days. Invoices in that window get a second look.",
    "duplicate_number_changed": "An invoice with this same number was received before, but with a different amount or bank details.",
    "payee_person": "The payee looks like a private individual rather than the vendor company.",
    "payee_crypto": "The payment instructions refer to a cryptocurrency wallet.",
    "payee_gift_card": "The payment instructions refer to gift cards.",
    "payee_country": "The payee's country is different from the country on the vendor record.",
    "risky_file": "The attachment is a type that can carry a program (for example a macro-enabled Office file, a script or an archive "
                  "containing one). It was never opened.",
    "hidden_text": "The document contains hidden text — white-on-white, tiny, off the page, or instructions aimed at an automated reader.",
    "password_archive": "The attachment is a password-protected archive, which cannot be inspected.",
}

SENDER_TRIGGERS = {"direct_sender", "lookalike_domain", "unlisted_sender", "dmarc_fail"}
BANK_TRIGGERS = {"bank_mismatch", "bank_change_request", "bank_changed_30d", "payee_person", "payee_crypto", "payee_gift_card",
                 "payee_country", "duplicate_number_changed"}
CALLBACK_SOURCES = {
    "contract": "a number from the signed contract",
    "onboarding_file": "a number from the vendor onboarding file",
    "vendor_master_90d": "the vendor-master phone number, unchanged for at least 90 days",
}
ESCALATE_AFTER_BUSINESS_DAYS = 2


class QuarantineError(Exception):
    """Refused decision; message written for the Controller's screen."""


def trigger_text(code: str) -> str:
    return TRIGGER_TEXT.get(code, f"Screening rule '{code}' fired. Ask the system owner what it checks.")


def _load_case(conn: psycopg.Connection, case_id, *, lock: bool = True) -> dict:
    row = conn.execute(f"select * from quarantine_case where id = %s {'for update' if lock else ''}", (case_id,)).fetchone()
    if row is None:
        raise QuarantineError("That quarantine case does not exist.")
    return row


def needs_callback(case: dict) -> bool:
    return bool(set(case["triggers"] or []) & (SENDER_TRIGGERS | BANK_TRIGGERS))


def needs_second_person(conn: psycopg.Connection, case: dict, inv: dict | None) -> bool:
    if needs_callback(case):
        return True
    gross = int(inv["gross_cents"]) if inv and inv.get("gross_cents") is not None else 0
    return gross > settings.second_person_release_cents


def _is_reviewer_role(user: dict | None) -> bool:
    return bool(user and user["status"] == "active" and set(user["roles"] or []) & {"controller", "cfo", "admin"})


def release(conn: psycopg.Connection, case_id, by, second_by, callback_source: str | None, callback_number: str | None,
            note: str | None) -> str:
    """Returns 'released' or 'pending_second_person'.

    First call (by the reviewer): records the callback and, when a second person is needed, parks the
    release and names them. Second call (by that named person): completes it. The same person cannot
    do both — that is the whole point of the second person.
    """
    case = _load_case(conn, case_id)
    if case["state"] != "open":
        raise QuarantineError("This case has already been decided.")
    inv = load_invoice(conn, case["invoice_id"], lock=True) if case["invoice_id"] else None
    second_needed = needs_second_person(conn, case, inv)

    # Confirmation by the named second person.
    if case["release_pending_at"] is not None and str(by) == str(case["second_reviewer_id"]):
        return _complete_release(conn, case, inv, by=case["reviewer_id"], second_by=by)

    if needs_callback(case):
        if callback_source not in CALLBACK_SOURCES:
            raise QuarantineError("A callback to the vendor is required before release. Choose where the phone number came from — "
                                  "never the number in the email.")
        if not (callback_number or "").strip():
            raise QuarantineError("Enter the phone number you called.")
        if callback_source == "vendor_master_90d" and inv is not None:
            since = inv.get("vendor_phone_since")
            if since is None or (datetime.now(timezone.utc).date() - since) < timedelta(days=90):
                raise QuarantineError("The vendor-master number has not been on file for 90 days, so it cannot be used for the callback. "
                                      "Use the contract or the onboarding file.")
    if second_needed:
        if second_by is None or str(second_by) == str(by):
            raise QuarantineError("This release needs a second person. Choose another Controller or the CFO — not yourself.")
        second = user_row(conn, second_by)
        if not _is_reviewer_role(second):
            raise QuarantineError("The second person must be an active Controller or CFO.")
        conn.execute(
            """update quarantine_case set reviewer_id = %s, second_reviewer_id = %s, callback_source = %s, callback_number = %s,
                      callback_note = %s, release_pending_at = now() where id = %s""",
            (by, second_by, callback_source, (callback_number or "").strip() or None, (note or "").strip() or None, case_id),
        )
        log_event(conn, event_type="quarantine", actor_type="controller", actor_id=by, invoice_id=case["invoice_id"],
                  reason_code="release_pending_second_person",
                  payload={"case_id": str(case_id), "second_reviewer_id": str(second_by), "callback_source": callback_source})
        notify_queue(conn, type="pending", user_id=second_by, invoice_id=case["invoice_id"],
                     context={"what": "quarantine_second_person", "case_id": str(case_id)})
        return "pending_second_person"

    conn.execute(
        """update quarantine_case set reviewer_id = %s, callback_source = %s, callback_number = %s, callback_note = %s where id = %s""",
        (by, callback_source, (callback_number or "").strip() or None, (note or "").strip() or None, case_id),
    )
    case = _load_case(conn, case_id)
    return _complete_release(conn, case, inv, by=by, second_by=None)


def _complete_release(conn: psycopg.Connection, case: dict, inv: dict | None, *, by, second_by) -> str:
    conn.execute("update quarantine_case set state = 'released', decision_at = now(), release_pending_at = null, decision_reason = 'released' where id = %s",
                 (case["id"],))
    if inv is not None:
        conn.execute("update invoice set tags = array_append(tags, 'released_from_quarantine') where id = %s and not ('released_from_quarantine' = any(tags))",
                     (inv["id"],))
        if inv["state"] == "QUARANTINED":
            transition(conn, inv["id"], "CAPTURED", reason_code="controller_release",
                       actor_type="second_person" if second_by else "controller", actor_id=second_by or by,
                       payload={"case_id": str(case["id"]), "reviewer_id": str(by), "second_reviewer_id": str(second_by) if second_by else None,
                                "callback_source": case.get("callback_source"), "triggers": list(case["triggers"] or [])})
            # QUA-3: full chain, never auto-release. Route only when capture already produced an amount and a
            # vendor; otherwise the pipeline (Package A) must finish capture first and will route on CAPTURED.
            if inv.get("vendor_id") and inv.get("gross_cents") is not None:
                from accounting.routing.route import route
                route(conn, inv["id"])
    else:
        log_event(conn, event_type="quarantine", actor_type="controller", actor_id=by, reason_code="controller_release",
                  payload={"case_id": str(case["id"]), "message_id": str(case["message_id"]) if case["message_id"] else None})
    return "released"


def reject(conn: psycopg.Connection, case_id, by, reason: str) -> None:
    reason = (reason or "").strip()
    if not reason:
        raise QuarantineError("A reason is required.")
    case = _load_case(conn, case_id)
    if case["state"] != "open":
        raise QuarantineError("This case has already been decided.")
    conn.execute("update quarantine_case set state = 'rejected', reviewer_id = %s, decision_at = now(), decision_reason = %s, release_pending_at = null where id = %s",
                 (by, reason, case_id))
    if case["invoice_id"]:
        inv = load_invoice(conn, case["invoice_id"], lock=True)
        if inv["state"] == "QUARANTINED":
            transition(conn, inv["id"], "FRAUD", reason_code="controller_reject", actor_type="controller", actor_id=by,
                       payload={"case_id": str(case_id), "reason": reason})
    else:
        log_event(conn, event_type="quarantine", actor_type="controller", actor_id=by, reason_code="controller_reject",
                  payload={"case_id": str(case_id), "reason": reason})


def sender_domain(conn: psycopg.Connection, case: dict) -> str | None:
    if not case.get("message_id"):
        return None
    m = conn.execute("select from_addr, original_sender from message where id = %s", (case["message_id"],)).fetchone()
    if m is None:
        return None
    addr = (m["original_sender"] or m["from_addr"] or "").strip().lower()
    if "<" in addr and ">" in addr:
        addr = addr[addr.rfind("<") + 1:addr.rfind(">")]
    return addr.rsplit("@", 1)[1] if "@" in addr else None


def report(conn: psycopg.Connection, case_id, by) -> dict:
    """QUA-4: block the sender's domain and tell security. The invoice (if any) is marked FRAUD."""
    case = _load_case(conn, case_id)
    if case["state"] != "open":
        raise QuarantineError("This case has already been decided.")
    domain = sender_domain(conn, case)
    if domain:
        conn.execute("insert into blocked_sender (pattern, reason, blocked_by) values (%s, %s, %s) on conflict (pattern) do nothing",
                     (domain, f"reported from quarantine case {case_id}", by))
    conn.execute("update quarantine_case set state = 'reported', reviewer_id = %s, decision_at = now(), decision_reason = 'reported', release_pending_at = null where id = %s",
                 (by, case_id))
    if case["invoice_id"]:
        inv = load_invoice(conn, case["invoice_id"], lock=True)
        if inv["state"] == "QUARANTINED":
            transition(conn, inv["id"], "FRAUD", reason_code="controller_report", actor_type="controller", actor_id=by,
                       payload={"case_id": str(case_id), "blocked_domain": domain})
    seq = log_event(conn, event_type="security", actor_type="controller", actor_id=by, invoice_id=case["invoice_id"],
                    reason_code="security_report", payload={"case_id": str(case_id), "blocked_domain": domain})
    conn.execute("insert into security_event (type, user_id, source, detail, event_seq) values ('security_report', %s, 'quarantine', %s::jsonb, %s)",
                 (by, '{"case_id": "%s"}' % case_id, seq))
    notify_queue(conn, type="security_report", user_id=None, invoice_id=case["invoice_id"], to=settings.security_email or None,
                 context={"case_id": str(case_id), "blocked_domain": domain, "message_id": str(case["message_id"]) if case["message_id"] else None})
    return {"blocked_domain": domain}


def escalate_stale(conn: psycopg.Connection, now: datetime | None = None) -> list[dict]:
    """QUA-5: a case unreviewed for two business days goes to the CFO. Stamped so it fires once."""
    now = now or utcnow()
    days = bd()
    intents: list[dict] = []
    cfos = None
    for c in conn.execute("select * from quarantine_case where state = 'open' and escalated_cfo_at is null order by opened_at").fetchall():
        if days.business_days_between(c["opened_at"], now) >= ESCALATE_AFTER_BUSINESS_DAYS:
            if cfos is None:
                cfos = users_with_role(conn, "cfo")
            conn.execute("update quarantine_case set escalated_cfo_at = %s where id = %s", (now, c["id"]))
            for u in cfos:
                intents.append({"type": "quarantine_escalated", "user_id": u["id"], "invoice_id": c["invoice_id"], "case_id": c["id"]})
            log_event(conn, event_type="quarantine", actor_type="system", invoice_id=c["invoice_id"], reason_code="escalated_cfo",
                      payload={"case_id": str(c["id"])})
    return intents


def is_escalated(case: dict, now: datetime | None = None) -> bool:
    now = now or utcnow()
    return case["state"] == "open" and bd().business_days_between(case["opened_at"], now) >= ESCALATE_AFTER_BUSINESS_DAYS


def fingerprint_comparison(conn: psycopg.Connection, inv: dict | None) -> dict:
    """What the Controller compares: the last 4 on file (and since when) against the invoice's."""
    if inv is None or not inv.get("vendor_id"):
        return {"on_file_last4": None, "since": None, "invoice_last4": inv.get("bank_last4") if inv else None, "matches": None}
    fp = conn.execute(
        "select last4, effective_from, source from vendor_bank_fingerprint where vendor_id = %s and superseded_at is null order by effective_from desc limit 1",
        (inv["vendor_id"],),
    ).fetchone()
    on_file = fp["last4"] if fp else None
    inv_last4 = inv.get("bank_last4")
    matches = None if (on_file is None or inv_last4 is None) else (on_file == inv_last4)
    if inv.get("bank_fingerprint") and fp is not None:
        stored = conn.execute("select fingerprint from vendor_bank_fingerprint where vendor_id = %s and superseded_at is null order by effective_from desc limit 1",
                              (inv["vendor_id"],)).fetchone()
        if stored is not None:
            matches = bytes(stored["fingerprint"]) == bytes(inv["bank_fingerprint"])
    return {"on_file_last4": on_file, "since": fp["effective_from"] if fp else None, "source": fp["source"] if fp else None,
            "invoice_last4": inv_last4, "matches": matches}


def vendor_history(conn: psycopg.Connection, vendor_id, exclude_invoice_id=None, limit: int = 6) -> list[dict]:
    if not vendor_id:
        return []
    return conn.execute(
        """select id, portal_number, invoice_number, invoice_date, gross_cents, currency, state, received_at, bank_last4
           from invoice where vendor_id = %s and (%s::uuid is null or id <> %s)
           order by received_at desc limit %s""",
        (vendor_id, exclude_invoice_id, exclude_invoice_id, limit),
    ).fetchall()
