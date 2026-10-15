"""The scheduler process (`accounting-worker` container).

    python -m accounting.worker              # run every job on its schedule
    python -m accounting.worker --once flush # run one job now and exit (for ops)

Why every job is wrapped: APScheduler drops a job that raises and keeps the others, but a job that
raises every time is a job that silently stopped. `_safe` logs the whole exception and swallows it,
so the scheduler never dies and the log shows the same failure every tick until somebody fixes it.

Why lazy imports: the ingest, pipeline, routing and quarantine packages are built by other
people at the same time as this one. A job whose package is missing logs "not available" and
the rest of the scheduler keeps running.
"""
from __future__ import annotations

import json
import logging
import sys
from datetime import datetime, timedelta, timezone
from functools import wraps
from typing import Callable

from accounting.config import settings
from accounting.db.pool import close_pool, connection

log = logging.getLogger("accounting.worker")


def _safe(name: str, fn: Callable[[], object]) -> Callable[[], object]:
    @wraps(fn)
    def run():
        try:
            result = fn()
            log.info("job %s finished: %r", name, result)
            return result
        except ImportError as exc:
            log.warning("job %s skipped: module not available (%s)", name, exc)
        except Exception:
            log.exception("job %s failed", name)
        return None
    return run


# --- the jobs -------------------------------------------------------------------------------------

def imap_poll():
    from accounting.ingest.imap_poller import poll_once
    return poll_once()


def process_pending():
    from accounting.pipeline import process_pending as _pp
    return _pp(limit=20)


def flush():
    from accounting.notify.send import flush as _flush
    return _flush(limit=50)


def reminders():
    from accounting.notify.reminders import run
    return run(datetime.now(timezone.utc))


def heartbeat():
    """INT-4: compare yesterday's forwarded volume with the expected range; out of range → Controller."""
    from accounting.ingest.heartbeat import check_daily
    from accounting.notify.send import queue
    result = check_daily()
    if not result.get("ok", True):
        with connection() as conn:
            queue(conn, type="heartbeat_alert", user_id=None, context={"text": result.get("text")})
    return result


def anchors():
    from accounting.audit.anchors import anchor_now
    row = anchor_now()
    return row["status"] if row else "empty log"


def nightly_verify():
    """AUD-1 chain re-walk plus SIG-7: every valid seal is re-verified. A seal that no longer
    verifies is a security event, not a log line."""
    from accounting.audit.verifier import verify_since_last_anchor
    from accounting.lifecycle import log_event
    ok, detail = verify_since_last_anchor()
    checked = bad = 0
    from accounting.reports.audit_export import seal_status
    with connection() as conn:
        sigs = conn.execute(
            "select id, invoice_id, payload_sha256, seal, seal_key_id from approval_signature where status = 'valid'").fetchall()
        for s in sigs:
            status = seal_status(s)
            if status is None:
                log.warning("SIG-7: no verification key available; %d seal(s) not re-verified", len(sigs))
                break
            checked += 1
            if status != "verified":
                bad += 1
                seq = log_event(conn, event_type="security", actor_type="system", invoice_id=s["invoice_id"],
                                reason_code="seal_verification_failed", payload={"signature_id": str(s["id"])})
                conn.execute(
                    "insert into security_event (type, source, detail, event_seq) values ('seal_invalid', 'nightly_verify', %s::jsonb, %s)",
                    (json.dumps({"signature_id": str(s["id"])}), seq))
    return {"chain_ok": ok, "detail": detail, "seals_checked": checked, "seals_bad": bad}


def sessions_sweep() -> dict:
    """ID-3 / ID-5: every open session older than the revalidation window is checked against GoTrue
    again. A ban or a removed grant shows up here within the hour even if the person never clicks;
    their sessions are revoked and, when the account is disabled, their open approvals move on."""
    from accounting.identity import crypto
    from accounting.identity.maxxdash import SignInError, provider
    from accounting.identity.sessions import revoke, revoke_all_for_user
    cutoff = datetime.now(timezone.utc) - timedelta(minutes=settings.session_revalidate_minutes)
    checked = revoked = reassigned = 0
    with connection() as conn:
        rows = conn.execute(
            "select * from portal_session where revoked_at is null and last_revalidated_at < %s", (cutoff,)).fetchall()
    for s in rows:
        checked += 1
        with connection() as conn:
            try:
                rt = crypto.decrypt(s["refresh_token_enc"], aad=str(s["user_id"]).encode()).decode()
                tokens = provider().refresh(rt)
                provider().get_user(tokens.access_token)
                ident = provider().identity(tokens.access_token)
                if "accounting" not in ident.portals:
                    raise SignInError("grant removed")
                conn.execute(
                    "update portal_session set refresh_token_enc = %s, access_token_exp = %s, last_revalidated_at = now() where id = %s",
                    (crypto.encrypt(tokens.refresh_token, aad=str(s["user_id"]).encode()),
                     datetime.now(timezone.utc) + timedelta(seconds=tokens.expires_in), s["id"]))
            except Exception as exc:
                revoked += 1
                revoke(conn, s["id"], "sweep_revalidation_failed")
                log.info("session %s revoked by sweep: %s", s["id"], exc)
                # A failed re-check means the directory no longer vouches for this person: treat the
                # account as disabled for routing purposes and move their open approvals on.
                conn.execute("update ap_user set status = 'disabled' where id = %s and status <> 'disabled'", (s["user_id"],))
                revoke_all_for_user(conn, s["user_id"], "sweep_revalidation_failed")
                try:
                    from accounting.routing.reassign import reassign_open_items_for
                    reassigned += int(reassign_open_items_for(conn, s["user_id"]) or 0)
                except ImportError:
                    log.warning("routing.reassign not available; open items for %s not reassigned", s["user_id"])
    return {"checked": checked, "revoked": revoked, "reassigned": reassigned}


def quarantine_escalate():
    from accounting.quarantine.service import escalate_stale
    with connection() as conn:
        return escalate_stale(conn)


def weekly_digest():
    from accounting.notify.reminders import weekly_digest as _wd
    return _wd()


def controller_daily():
    from accounting.notify.reminders import controller_daily as _cd
    return _cd()


JOBS: dict[str, Callable[[], object]] = {
    "imap_poll": imap_poll,
    "process_pending": process_pending,
    "flush": flush,
    "reminders": reminders,
    "heartbeat": heartbeat,
    "anchors": anchors,
    "nightly_verify": nightly_verify,
    "sessions_sweep": sessions_sweep,
    "quarantine_escalate": quarantine_escalate,
    "weekly_digest": weekly_digest,
    "controller_daily": controller_daily,
}


def build_scheduler():
    from apscheduler.schedulers.blocking import BlockingScheduler
    from apscheduler.triggers.cron import CronTrigger
    from apscheduler.triggers.interval import IntervalTrigger

    tz = settings.timezone
    sched = BlockingScheduler(timezone=tz, job_defaults={"coalesce": True, "max_instances": 1, "misfire_grace_time": 300})
    add = lambda name, trigger: sched.add_job(_safe(name, JOBS[name]), trigger, id=name, name=name)  # noqa: E731
    add("imap_poll", IntervalTrigger(seconds=max(15, settings.imap_poll_seconds)))
    add("process_pending", IntervalTrigger(seconds=30))
    add("flush", IntervalTrigger(seconds=30))
    add("reminders", IntervalTrigger(hours=1))
    add("heartbeat", CronTrigger(hour=8, minute=0, timezone=tz))
    add("anchors", IntervalTrigger(minutes=max(1, settings.anchor_interval_minutes)))
    add("nightly_verify", CronTrigger(hour=2, minute=15, timezone=tz))
    add("sessions_sweep", IntervalTrigger(hours=1))
    add("quarantine_escalate", IntervalTrigger(hours=1))
    add("weekly_digest", CronTrigger(day_of_week="mon", hour=7, minute=0, timezone=tz))
    add("controller_daily", CronTrigger(hour=7, minute=30, timezone=tz))
    return sched


def main(argv: list[str]) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    if "--once" in argv:
        i = argv.index("--once")
        name = argv[i + 1] if i + 1 < len(argv) else ""
        if name not in JOBS:
            print(f"unknown job {name!r}; one of: {', '.join(JOBS)}", file=sys.stderr)
            return 2
        result = _safe(name, JOBS[name])()
        print(f"{name}: {result!r}")
        return 0
    sched = build_scheduler()
    log.info("worker started with %d jobs", len(JOBS))
    try:
        sched.start()
    except (KeyboardInterrupt, SystemExit):
        pass
    return 0


if __name__ == "__main__":
    try:
        code = main(sys.argv[1:])
    finally:
        close_pool()
    sys.exit(code)
