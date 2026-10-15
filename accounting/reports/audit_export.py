"""One invoice's full history as CSV or JSON (RPT-3), with each signature's verification status.

Why verification is optional here: the seal is checked by `approvals.signing_service.verify_seal`,
which needs the signing service's public key. When that module or key is not reachable the export
still goes out — but it says "not verified" rather than implying a check that did not happen.
"""
from __future__ import annotations

import csv
import io
import json
import logging
from datetime import datetime, timezone

import psycopg

from accounting.lifecycle import _default as _json_default
from accounting.web.templating import CONFIDENTIAL_FOOTER

log = logging.getLogger("accounting.reports.audit_export")


def seal_public_key() -> str:
    """The key the seals are checked against: the KMS public key when configured
    (SIGNING_PUBLIC_KEY_PEM), else the dev signer's key in test mode, else nothing."""
    import os
    pem = os.environ.get("SIGNING_PUBLIC_KEY_PEM", "")
    if pem:
        return pem
    try:
        from accounting.approvals.signing_service import public_key_pem
        return public_key_pem() or ""
    except Exception:
        return ""


def seal_status(sig: dict) -> str | None:
    """'verified', 'FAILED', or None when no check could be made (module or key unavailable)."""
    try:
        from accounting.approvals.signing_service import verify_seal
    except ImportError:
        return None
    key = seal_public_key()
    if not key:
        return None
    try:
        ok = verify_seal(bytes(sig["payload_sha256"]), bytes(sig["seal"]), key)
    except Exception as exc:
        log.warning("verify_seal failed for %s: %s", sig.get("id"), exc)
        return None
    return "verified" if ok else "FAILED"


def _verify(sig: dict) -> str:
    return seal_status(sig) or "not verified"


def history(conn: psycopg.Connection, invoice_id) -> dict | None:
    inv = conn.execute(
        """select i.*, v.legal_name as vendor from invoice i left join vendor v on v.id = i.vendor_id where i.id = %s""",
        (invoice_id,)).fetchone()
    if inv is None:
        return None
    events = conn.execute(
        """select e.seq, e.at, e.event_type, e.from_state, e.to_state, e.reason_code, e.actor_type, e.payload,
                  u.display_name as actor, encode(e.hash, 'hex') as hash, encode(e.prev_hash, 'hex') as prev_hash
           from event_log e left join ap_user u on u.id = e.actor_id where e.invoice_id = %s order by e.seq""",
        (invoice_id,)).fetchall()
    sigs = conn.execute(
        """select s.*, u.display_name as signer from approval_signature s join ap_user u on u.id = s.signer_id
           where s.invoice_id = %s order by s.signed_at""", (invoice_id,)).fetchall()
    signatures = []
    for s in sigs:
        signatures.append({
            "id": str(s["id"]), "signer": s["signer"], "role": s["signer_role"], "decision": s["decision"],
            "signed_at": s["signed_at"], "status": s["status"], "seal_key_id": s["seal_key_id"],
            "attestation_text_id": s["attestation_text_id"], "matrix_version_id": s["matrix_version_id"],
            "payload_sha256": bytes(s["payload_sha256"]).hex(),
            "verification": _verify(s) if s["status"] == "valid" else "not checked (" + s["status"] + ")",
        })
    return {
        "generated_at": datetime.now(timezone.utc),
        "invoice": {"id": str(inv["id"]), "portal_number": f"AP-{int(inv['portal_number']):06d}", "state": inv["state"],
                    "vendor": inv["vendor"], "invoice_number": inv["invoice_number"], "received_at": inv["received_at"],
                    "gross_cents": str(inv["gross_cents"]) if inv["gross_cents"] is not None else None,
                    "sap_document_number": inv["sap_document_number"]},
        "events": events,
        "signatures": signatures,
        "footer": CONFIDENTIAL_FOOTER,
    }


def to_json(h: dict) -> str:
    return json.dumps(h, indent=2, default=_json_default, ensure_ascii=False)


def to_csv(h: dict) -> str:
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(["# Audit history for", h["invoice"]["portal_number"], "generated", h["generated_at"].isoformat(timespec="seconds")])
    w.writerow(["seq", "at", "event_type", "from_state", "to_state", "reason_code", "actor", "actor_type", "payload", "hash", "prev_hash"])
    for e in h["events"]:
        w.writerow([e["seq"], e["at"].isoformat(), e["event_type"], e["from_state"] or "", e["to_state"] or "", e["reason_code"] or "",
                    e["actor"] or "", e["actor_type"], json.dumps(e["payload"], default=_json_default, sort_keys=True), e["hash"], e["prev_hash"]])
    w.writerow([])
    w.writerow(["signature_id", "signer", "role", "decision", "signed_at", "status", "seal_key_id", "payload_sha256", "verification"])
    for s in h["signatures"]:
        w.writerow([s["id"], s["signer"], s["role"], s["decision"], s["signed_at"].isoformat(), s["status"], s["seal_key_id"],
                    s["payload_sha256"], s["verification"]])
    w.writerow([f"# {CONFIDENTIAL_FOOTER}"])
    return buf.getvalue()
