"""INT-4 heartbeat: did yesterday's forwarded volume land inside the expected range?

A redirect rule that someone switched off looks exactly like a quiet day, and a flood looks like a
busy one; only the count tells them apart. Business days only for the lower bound — a Sunday with
nothing forwarded is not an alarm. The result is a dict the notify package turns into the Controller's
heartbeat_alert email: {ok, n, text, day, min, max}.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from accounting.config import settings
from accounting.db.pool import connection


def _is_business_day(d: date) -> bool:
    try:
        from accounting.business_days import is_business_day  # Package C; weekends only when absent
        return is_business_day(d)
    except Exception:
        return d.weekday() < 5


def check_daily(day: date | None = None) -> dict:
    tz = ZoneInfo(settings.timezone)
    day = day or (datetime.now(tz).date() - timedelta(days=1))
    start = datetime.combine(day, datetime.min.time(), tzinfo=tz).astimezone(timezone.utc)
    end = start + timedelta(days=1)
    with connection() as conn:
        n = conn.execute(
            "select count(*) as n from message where trust = 'forwarded_by_maxx' and ingested_at >= %s and ingested_at < %s",
            (start, end)).fetchone()["n"]
        direct = conn.execute(
            "select count(*) as n from message where trust = 'direct' and ingested_at >= %s and ingested_at < %s",
            (start, end)).fetchone()["n"]
    lo, hi = settings.heartbeat_min_per_day, settings.heartbeat_max_per_day
    business = _is_business_day(day)
    if business and n < lo:
        ok, text = False, (f"Only {n} forwarded message(s) arrived on {day:%A, %b %d}; at least {lo} expected. "
                           "The Microsoft 365 redirect rule may have stopped.")
    elif n > hi:
        ok, text = False, f"{n} forwarded messages arrived on {day:%A, %b %d}; the ceiling is {hi}. Check for a loop or a flood."
    else:
        ok, text = True, f"{n} forwarded message(s) on {day:%A, %b %d} — within the expected range."
    if direct:
        text += f" {direct} message(s) arrived directly and were quarantined."
    return {"ok": ok, "n": int(n), "direct": int(direct), "text": text, "day": day.isoformat(), "min": lo, "max": hi,
            "business_day": business}
