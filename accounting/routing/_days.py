"""Business-day arithmetic for routing, with a fallback.

Package C owns `accounting.business_days` (US federal holidays, America/New_York). Routing needs the
same arithmetic but must import and run before that module exists, so every caller goes through
`bd()`, which returns the real module when it is importable and a weekday-only stand-in otherwise.
The stand-in deliberately knows no holidays: when it is in use a reminder may fire a day early on a
holiday week, which is a nuisance, not a control failure.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from types import SimpleNamespace


def _is_business_day(d: date | datetime) -> bool:
    return d.weekday() < 5


def _add_business_days(dt: datetime, n: int) -> datetime:
    step = 1 if n >= 0 else -1
    remaining = abs(n)
    while remaining > 0:
        dt = dt + timedelta(days=step)
        if _is_business_day(dt):
            remaining -= 1
    return dt


def _business_days_between(a: datetime, b: datetime) -> int:
    """Whole business days from a to b (b later). Counts each weekday boundary crossed."""
    if b < a:
        return -_business_days_between(b, a)
    days = 0
    cur = a.date()
    end = b.date()
    while cur < end:
        cur = cur + timedelta(days=1)
        if _is_business_day(cur):
            days += 1
    return days


_FALLBACK = SimpleNamespace(
    add_business_days=_add_business_days,
    business_days_between=_business_days_between,
    is_business_day=_is_business_day,
)


def bd():
    try:
        from accounting import business_days  # Package C
        return business_days
    except Exception:  # ImportError, or a half-written module during parallel development
        return _FALLBACK


def utcnow() -> datetime:
    return datetime.now(timezone.utc)
