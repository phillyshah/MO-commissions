"""`poll_once() -> int`: fetch what is new in accounts@90ten.life and hand each message to intake.

The cursor is the IMAP UID, kept in `ingest_state` keyed by mailbox and folder. UIDs are only
meaningful for one UIDVALIDITY; when the server changes it the cursor is reset and every message is
offered again — which is safe, because intake is idempotent on (message_id, raw_sha256). Messages are
never deleted or moved from the mailbox: the mailbox is the second copy of the original.

One failure does not stop the batch, and the cursor only advances past a UID that was ingested or
recorded as failed in the log; a message that raises is retried next poll.
"""
from __future__ import annotations

import imaplib
import logging
import ssl

from accounting.config import settings
from accounting.db.pool import connection
from accounting.ingest.intake import ingest_raw_message
from accounting.lifecycle import log_event

log = logging.getLogger("accounting.ingest.imap")

BATCH = 50


def _key() -> str:
    return f"{settings.imap_user}/{settings.imap_folder}"


def _cursor(conn) -> dict:
    row = conn.execute("select * from ingest_state where key = %s", (_key(),)).fetchone()
    if row is None:
        row = conn.execute("insert into ingest_state (key) values (%s) returning *", (_key(),)).fetchone()
    return row


def _save(conn, last_uid: int, uidvalidity: int | None) -> None:
    conn.execute("update ingest_state set last_uid = %s, uidvalidity = %s, updated_at = now() where key = %s",
                 (last_uid, uidvalidity, _key()))


def _connect() -> imaplib.IMAP4_SSL:
    ctx = ssl.create_default_context()
    client = imaplib.IMAP4_SSL(settings.imap_host, settings.imap_port, ssl_context=ctx, timeout=60)
    client.login(settings.imap_user, settings.imap_password)
    return client


def poll_once() -> int:
    if not settings.imap_password:
        log.warning("IMAP_PASSWORD is not set; poll skipped")
        return 0
    with connection() as conn:
        cur = _cursor(conn)
    client = _connect()
    ingested = 0
    try:
        typ, data = client.select(settings.imap_folder, readonly=True)
        if typ != "OK":
            raise RuntimeError(f"cannot select folder {settings.imap_folder}")
        typ, uv = client.response("UIDVALIDITY")
        uidvalidity = int(uv[0]) if uv and uv[0] else None
        last_uid = int(cur["last_uid"] or 0)
        if cur["uidvalidity"] is not None and uidvalidity is not None and uidvalidity != cur["uidvalidity"]:
            log.warning("UIDVALIDITY changed %s -> %s; cursor reset", cur["uidvalidity"], uidvalidity)
            last_uid = 0
        typ, found = client.uid("search", None, f"UID {last_uid + 1}:*")
        uids = [int(u) for u in (found[0].split() if found and found[0] else []) if int(u) > last_uid]
        for uid in sorted(uids)[:BATCH]:
            typ, parts = client.uid("fetch", str(uid), "(BODY.PEEK[])")
            raw = next((p[1] for p in parts if isinstance(p, tuple) and len(p) > 1), None)
            if typ != "OK" or raw is None:
                log.error("fetch failed for uid %s", uid)
                break
            try:
                ingest_raw_message(bytes(raw), imap_uid=uid, source="email")
                ingested += 1
            except Exception:
                log.exception("ingest failed for uid %s", uid)
                with connection() as conn:
                    log_event(conn, event_type="ingest_error", actor_type="system", payload={"imap_uid": uid})
                break  # retried next poll; cursor stays before this uid
            last_uid = uid
            with connection() as conn:
                _save(conn, last_uid, uidvalidity)
        else:
            with connection() as conn:
                _save(conn, last_uid, uidvalidity)
    finally:
        try:
            client.logout()
        except Exception:
            pass
    return ingested
