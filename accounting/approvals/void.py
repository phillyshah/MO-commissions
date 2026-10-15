"""Voiding signatures (SIG-3, SIG-5, DEL-1).

approval_signature is insert-only, so a void is a NEW row with status 'voided' whose `supersedes`
points at the record it cancels, carrying the same payload and seal so the audit trail still shows
exactly what had been signed. Open approval requests are closed as 'voided' at the same time: a
request that survived a content change would let somebody sign the new content under the old chain.
"""
from __future__ import annotations

from uuid import UUID

import psycopg

from accounting.lifecycle import log_event


def void_all_for_invoice(conn: psycopg.Connection, invoice_id: UUID | str, reason: str, by: UUID | str | None) -> int:
    rows = conn.execute(
        """select s.* from approval_signature s
           where s.invoice_id = %s and s.status = 'valid'
             and not exists (select 1 from approval_signature x where x.supersedes = s.id)
           order by s.signed_at""",
        (invoice_id,),
    ).fetchall()
    voided = []
    for s in rows:
        new = conn.execute(
            """insert into approval_signature (invoice_id, approval_request_id, signer_id, signer_role, decision, reason, content_hash,
                                               payload, payload_sha256, passkey_id, signing_session_id, assertion, seal, seal_key_id,
                                               attestation_text_id, matrix_version_id, status, supersedes)
               values (%s,%s,%s,%s,%s,%s,%s,%s::jsonb,%s,%s,%s,%s::jsonb,%s,%s,%s,%s,'voided',%s) returning id""",
            (s["invoice_id"], s["approval_request_id"], s["signer_id"], s["signer_role"], s["decision"], reason, s["content_hash"],
             _json(s["payload"]), s["payload_sha256"], s["passkey_id"], s["signing_session_id"], _json(s["assertion"]), s["seal"],
             s["seal_key_id"], s["attestation_text_id"], s["matrix_version_id"], s["id"]),
        ).fetchone()
        voided.append({"signature_id": str(s["id"]), "void_record_id": str(new["id"])})
    n_req = conn.execute(
        "update approval_request set status = 'voided', closed_at = now() where invoice_id = %s and status = 'open'", (invoice_id,)
    ).rowcount
    if voided or n_req:
        log_event(conn, event_type="signature", actor_type="user" if by else "system", actor_id=by, invoice_id=invoice_id,
                  reason_code="voided", payload={"reason": reason, "voided": voided, "requests_closed": n_req})
    return len(voided)


def _json(v):
    import json
    return json.dumps(v) if v is not None else None
