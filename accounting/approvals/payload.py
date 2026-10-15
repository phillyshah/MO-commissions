"""What an approver signs (SIG-1) and the content hash that binds it to the invoice (SIG-3).

The content hash covers the original file AND the fields as approved, so a change to either after a
signature is detectable by recomputing it. It is stored on the invoice the first time it is computed
and recomputed on every approval; a difference voids every signature on the invoice. Amounts are
strings in the payload because the payload is JSON and JSON numbers are floats to half the parsers
in the world.
"""
from __future__ import annotations

import hashlib
from datetime import datetime, timezone
from uuid import UUID

import psycopg

from accounting.lifecycle import canonical
from accounting.routing.common import load_invoice


def approved_fields(inv: dict) -> dict:
    """The fields a signer is approving, in the shape that is hashed. `inv` is load_invoice()'s row."""
    return {
        "invoice_number": inv.get("invoice_number"),
        "invoice_date": inv["invoice_date"].isoformat() if inv.get("invoice_date") else None,
        "due_date": inv["due_date"].isoformat() if inv.get("due_date") else None,
        "currency": inv.get("currency"),
        "gross_cents": str(inv["gross_cents"]) if inv.get("gross_cents") is not None else None,
        "tax_cents": str(inv["tax_cents"]) if inv.get("tax_cents") is not None else None,
        "po_number": inv.get("po_number"),
        "remit_to": inv.get("remit_to"),
        "bank_last4": inv.get("bank_last4"),
        "vendor_id": str(inv["vendor_id"]) if inv.get("vendor_id") else None,
        "gl": inv.get("vendor_default_gl"),
        "cost_center": inv.get("vendor_default_cost_center"),
    }


def compute_content_hash(inv: dict) -> bytes:
    file_hash = bytes(inv["file_sha256"]) if inv.get("file_sha256") else b""
    return hashlib.sha256(file_hash + canonical(approved_fields(inv))).digest()


def content_hash(conn: psycopg.Connection, invoice_id: UUID | str, inv: dict | None = None) -> bytes:
    """Compute the hash; store it on the invoice the first time so later approvals compare against
    what the first signer saw rather than against themselves."""
    inv = inv or load_invoice(conn, invoice_id)
    if inv is None:
        raise LookupError(f"invoice {invoice_id} not found")
    h = compute_content_hash(inv)
    if inv.get("content_hash") is None:
        conn.execute("update invoice set content_hash = %s where id = %s and content_hash is null", (h, invoice_id))
    return h


def build_payload(conn: psycopg.Connection, invoice_id: UUID | str, signer, role: str, decision: str,
                  attestation_id: str, *, matrix_version_id: int | None = None, inv: dict | None = None) -> dict:
    """SIG-1 payload. `signer` is a CurrentUser or a user id. server_time is the one field that makes
    two payloads for the same decision differ, which is why the payload is stored beside its challenge
    (webauthn_challenge.payload) rather than rebuilt at verification."""
    inv = inv or load_invoice(conn, invoice_id)
    signer_id = getattr(signer, "id", signer)
    return {
        "invoice_id": str(inv["id"]),
        "content_hash": content_hash(conn, invoice_id, inv).hex(),
        "vendor_sap_number": inv.get("vendor_sap_number"),
        "bank_fingerprint_ref": bytes(inv["bank_fingerprint"]).hex() if inv.get("bank_fingerprint") else None,
        "gross_cents": str(inv["gross_cents"]) if inv.get("gross_cents") is not None else None,
        "currency": inv.get("currency"),
        "gl": inv.get("vendor_default_gl"),
        "cost_center": inv.get("vendor_default_cost_center"),
        "po_number": inv.get("po_number"),
        "matrix_version_id": matrix_version_id,
        "signer_user_id": str(signer_id),
        "signer_role": role,
        "attestation_text_id": attestation_id,
        "decision": decision,
        "server_time": datetime.now(timezone.utc).isoformat(),
    }


def payload_hash(payload: dict) -> bytes:
    return hashlib.sha256(canonical(payload)).digest()
