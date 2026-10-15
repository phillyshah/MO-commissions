"""Jinja rendering with the helpers every page needs. Money is always formatted here, from cents."""
from __future__ import annotations

from datetime import date, datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from fastapi import Request
from fastapi.templating import Jinja2Templates

from accounting import APP_NAME, APP_VERSION
from accounting.config import settings

TEMPLATES_DIR = Path(__file__).parent / "templates"
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

CONFIDENTIAL_FOOTER = "CONFIDENTIAL INFORMATION. NOT FOR DISTRIBUTION."


def money(cents: int | None, currency: str | None = "USD") -> str:
    if cents is None:
        return "—"
    sign = "-" if cents < 0 else ""
    cents = abs(int(cents))
    symbol = {"USD": "$", "EUR": "€", "GBP": "£"}.get(currency or "USD", (currency or "") + " ")
    return f"{sign}{symbol}{cents // 100:,}.{cents % 100:02d}"


def local(dt: datetime | None, fmt: str = "%b %d, %Y %I:%M %p") -> str:
    if dt is None:
        return "—"
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(ZoneInfo(settings.timezone)).strftime(fmt).replace(" 0", " ")


def day(d: date | datetime | None) -> str:
    if d is None:
        return "—"
    if isinstance(d, datetime):
        return local(d, "%b %d, %Y")
    return d.strftime("%b %d, %Y")


def age_days(since: datetime | None) -> str:
    if since is None:
        return "—"
    delta = datetime.now(timezone.utc) - (since if since.tzinfo else since.replace(tzinfo=timezone.utc))
    d = delta.days
    return "today" if d == 0 else f"{d} day{'s' if d != 1 else ''}"


templates.env.filters.update({"money": money, "local": local, "day": day, "age": age_days})
templates.env.globals.update({
    "APP_NAME": APP_NAME,
    "APP_VERSION": APP_VERSION,
    "CONFIDENTIAL_FOOTER": CONFIDENTIAL_FOOTER,
    "DASHBOARD_URL": settings.maxxdash_dashboard_url,
    "USAGE_URL": settings.usage_url,
    "COMMISSIONS_GENERATOR_URL": settings.commissions_generator_url,
    # Sidebar badges and the break-glass flag, looked up by base.html itself so every page gets them
    # without each router remembering to pass them. Imported lazily: reports imports this module.
    "nav_counts_for": lambda user: __import__("accounting.reports.dashboard", fromlist=["nav_counts_for"]).nav_counts_for(user),
})


def render(request: Request, name: str, context: dict | None = None, status_code: int = 200):
    ctx = {"request": request, "user": getattr(request.state, "user", None)}
    ctx.update(context or {})
    return templates.TemplateResponse(request, name, ctx, status_code=status_code)
