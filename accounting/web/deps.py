"""FastAPI dependencies: who is signed in, and what they may do.

ID-2: an unauthenticated or un-granted visitor gets a generic 404 from every portal path except /login
and the auth endpoints. Role checks raise 404 too — the portal never reveals what exists.
"""
from __future__ import annotations

from typing import Callable

from fastapi import Depends, HTTPException, Request

from accounting.identity.sessions import CurrentUser, resolve


class NotFound(HTTPException):
    def __init__(self):
        super().__init__(status_code=404)


def current_user(request: Request) -> CurrentUser:
    user = getattr(request.state, "user", None)
    if user is None:
        user = resolve(request)
        request.state.user = user
    if user is None or user.status == "disabled":
        raise NotFound()
    return user


def require(*roles: str) -> Callable[[CurrentUser], CurrentUser]:
    def _dep(user: CurrentUser = Depends(current_user)) -> CurrentUser:
        if "admin" in user.roles or user.has(*roles):
            return user
        raise NotFound()
    return _dep


def require_active(user: CurrentUser = Depends(current_user)) -> CurrentUser:
    """Frozen users may look but not act."""
    if user.status != "active":
        raise HTTPException(status_code=403, detail="Your access is frozen. Contact the Controller.")
    return user
