"""Small helpers shared by the Package C routers (invoices, uploads, admin, reports)."""
from __future__ import annotations

from datetime import datetime, timezone
from uuid import UUID

from fastapi import Request

from accounting.web.deps import NotFound
from accounting.web.templating import render


def page(request: Request, user, name: str, ctx: dict | None = None, status_code: int = 200):
    """Render with the signed-in user on `request.state`, which base.html reads for the sidebar."""
    request.state.user = user
    ctx = dict(ctx or {})
    ctx.setdefault("as_of", datetime.now(timezone.utc))
    return render(request, name, ctx, status_code=status_code)


def as_uuid(value: str) -> UUID:
    """A malformed id is a 404 like any other unknown path — never a stack trace."""
    try:
        return UUID(str(value))
    except (ValueError, AttributeError, TypeError):
        raise NotFound()


def portal_number(n: int | None) -> str:
    return f"AP-{int(n):06d}" if n is not None else "AP-?"
