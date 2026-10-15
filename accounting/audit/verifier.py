"""Re-walk the audit chain from the last written anchor (AUD-1).

Two checks, because they catch different things. The anchor check — the row at the anchored seq
still carries the anchored hash — catches a rewrite of history BEFORE the anchor, which a chain
walk from that point could never see. The chain walk from that seq catches any rewrite after it.
The off-box verifier in Lenny's AWS account does the same from the bucket copy; this one runs
nightly on the box so a problem is seen within a day even if the off-box job is down.
"""
from __future__ import annotations

import json
import logging

from accounting.db.pool import connection
from accounting.lifecycle import log_event, verify_chain

log = logging.getLogger("accounting.audit.verifier")


def verify_since_last_anchor() -> tuple[bool, str]:
    with connection() as conn:
        anchor = conn.execute(
            "select * from anchor where status = 'written' order by head_seq desc, id desc limit 1").fetchone()
        if anchor is None:
            ok, bad = verify_chain(conn, 1)
            detail = "No anchor yet; the whole chain was checked." if ok else f"Chain broken at event {bad} (no anchor yet)."
        else:
            row = conn.execute("select hash from event_log where seq = %s", (anchor["head_seq"],)).fetchone()
            if row is None or bytes(row["hash"]) != bytes(anchor["head_hash"]):
                ok, detail = False, f"Event {anchor['head_seq']} no longer matches the anchor written at {anchor['at']:%Y-%m-%d %H:%M}."
            else:
                ok, bad = verify_chain(conn, int(anchor["head_seq"]))
                detail = (f"Chain intact from anchored event {anchor['head_seq']}." if ok
                          else f"Chain broken at event {bad} (after anchor {anchor['head_seq']}).")
        if not ok:
            log.error("audit chain verification failed: %s", detail)
            try:
                seq = log_event(conn, event_type="security", actor_type="system", reason_code="chain_verification_failed",
                                payload={"detail": detail})
                conn.execute("insert into security_event (type, source, detail, event_seq) values ('chain_broken', 'verifier', %s::jsonb, %s)",
                             (json.dumps({"detail": detail}), seq))
            except Exception:
                log.exception("could not record the verification failure")
    return ok, detail
