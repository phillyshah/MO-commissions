"""Business-day arithmetic in America/New_York with US federal holidays.

Why compute holidays instead of keeping a table: the reminder clock (D+2, D+3, D+4), the matrix
"effective next business day" rule and every cycle-time KPI depend on it, and a table somebody has
to remember to extend every December is a clock that silently runs fast in January. The federal
rules are fixed in statute (5 U.S.C. 6103) so they can be derived for any year.

Observed-day rule: a fixed-date holiday on a Saturday is observed the Friday before; on a Sunday,
the Monday after. That is why New Year's Day can land in the PREVIOUS calendar year (31 December)
and why holidays for a year are computed from both that year's and the next year's fixed dates.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from functools import lru_cache
from zoneinfo import ZoneInfo

TZ = ZoneInfo("America/New_York")


def _nth_weekday(year: int, month: int, weekday: int, n: int) -> date:
    """n-th (1-based) given weekday (Mon=0) of a month; n=-1 means the last one."""
    if n > 0:
        first = date(year, month, 1)
        offset = (weekday - first.weekday()) % 7
        return first + timedelta(days=offset + 7 * (n - 1))
    last = date(year + (month == 12), (month % 12) + 1, 1) - timedelta(days=1)
    offset = (last.weekday() - weekday) % 7
    return last - timedelta(days=offset)


def _observed(d: date) -> date:
    if d.weekday() == 5:
        return d - timedelta(days=1)
    if d.weekday() == 6:
        return d + timedelta(days=1)
    return d


def _fixed_for_year(year: int) -> list[date]:
    return [
        _observed(date(year, 1, 1)),     # New Year's Day
        _observed(date(year, 6, 19)),    # Juneteenth
        _observed(date(year, 7, 4)),     # Independence Day
        _observed(date(year, 11, 11)),   # Veterans Day
        _observed(date(year, 12, 25)),   # Christmas Day
    ]


@lru_cache(maxsize=64)
def federal_holidays(year: int) -> frozenset[date]:
    """Observed federal holidays that fall within the calendar `year`."""
    days = set(_fixed_for_year(year)) | set(_fixed_for_year(year + 1)) | set(_fixed_for_year(year - 1))
    days |= {
        _nth_weekday(year, 1, 0, 3),    # Martin Luther King Jr. Day — third Monday in January
        _nth_weekday(year, 2, 0, 3),    # Washington's Birthday — third Monday in February
        _nth_weekday(year, 5, 0, -1),   # Memorial Day — last Monday in May
        _nth_weekday(year, 9, 0, 1),    # Labor Day — first Monday in September
        _nth_weekday(year, 10, 0, 2),   # Columbus Day — second Monday in October
        _nth_weekday(year, 11, 3, 4),   # Thanksgiving — fourth Thursday in November
    }
    return frozenset(d for d in days if d.year == year)


def local_date(value: date | datetime) -> date:
    """The New York calendar date of a moment. Naive datetimes are taken as UTC, matching the DB."""
    if isinstance(value, datetime):
        if value.tzinfo is None:
            value = value.replace(tzinfo=ZoneInfo("UTC"))
        return value.astimezone(TZ).date()
    return value


def is_business_day(d: date | datetime) -> bool:
    day = local_date(d)
    return day.weekday() < 5 and day not in federal_holidays(day.year)


def add_business_days(dt: date | datetime, n: int):
    """Move `n` business days forward (or back for negative n). Datetimes keep their time of day,
    so a 'due in two business days' deadline lands at the same clock time, not at midnight."""
    step = 1 if n >= 0 else -1
    remaining = abs(n)
    cur = dt
    while remaining:
        cur = cur + timedelta(days=step)
        if is_business_day(cur):
            remaining -= 1
    return cur


def business_days_between(a: date | datetime, b: date | datetime) -> int:
    """Business days strictly after `a` up to and including `b` (New York dates). Same day is 0;
    the next business morning is 1. Negative when b is before a."""
    da, db = local_date(a), local_date(b)
    if db < da:
        return -business_days_between(db, da)
    n = 0
    cur = da
    while cur < db:
        cur += timedelta(days=1)
        if is_business_day(cur):
            n += 1
    return n
