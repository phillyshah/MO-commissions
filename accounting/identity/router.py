"""Sign-in, sign-out, ForwardAuth for the sibling apps, and the one email link that changes state."""
from __future__ import annotations

import hashlib
import secrets
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

from fastapi import APIRouter, Form, Request, Response, status
from fastapi.responses import JSONResponse, RedirectResponse

from accounting.config import settings
from accounting.db.pool import connection
from accounting.identity import ratelimit, sessions
from accounting.identity.maxxdash import SignInError, provider
from accounting.lifecycle import log_event
from accounting.web.templating import render

router = APIRouter()


@router.get("/login")
def login_page(request: Request, next: str = "/"):
    if sessions.resolve(request):
        return RedirectResponse("/", status_code=303)
    return render(request, "login.html", {"next": _safe_next(next), "error": None, "mode": "signin"})


@router.post("/auth/session")
def sign_in(request: Request, email: str = Form(...), password: str = Form(...), next: str = Form("/")):
    address = sessions.client_ip(request)
    email_n = email.strip().lower()
    ok1, retry1 = ratelimit.per_account.hit(f"{address}|{email_n}")
    ok2, retry2 = ratelimit.per_address.hit(address)
    if not (ok1 and ok2):
        resp = render(request, "login.html", {"next": _safe_next(next), "error": ratelimit.TOO_MANY, "mode": "signin"},
                      status_code=429)
        resp.headers["Retry-After"] = str(max(retry1, retry2))
        return resp
    try:
        tokens = provider().password_grant(email_n, password, address)
        ident = provider().identity(tokens.access_token)
    except SignInError as e:
        return render(request, "login.html", {"next": _safe_next(next), "error": e.message, "mode": "signin"}, status_code=e.status)
    except Exception:
        return render(request, "login.html", {"next": _safe_next(next), "error": "Connection failed — try again.", "mode": "signin"},
                      status_code=502)
    ratelimit.per_account.clear(f"{address}|{email_n}")
    if "accounting" not in ident.portals:
        # ID-2: no hint the system exists.
        return render(request, "404.html", {}, status_code=404)
    try:
        with connection() as conn:
            raw, _row = sessions.create_session(conn, request, ident, tokens)
    except SignInError as e:
        return render(request, "login.html", {"next": _safe_next(next), "error": e.message, "mode": "signin"}, status_code=e.status)
    resp = RedirectResponse(_safe_next(next), status_code=303)
    sessions.set_cookie(resp, raw, _row.get("sso_raw"))
    return resp


@router.post("/auth/recover")
def recover(request: Request, email: str = Form(...)):
    try:
        provider().recover(email)
    except Exception:
        pass
    return render(request, "login.html", {
        "next": "/", "error": None, "mode": "reset",
        "note": "If an account exists for that email, a reset link has been sent from the dashboard. "
                "If you don't receive it, contact your administrator."})


@router.post("/auth/signout")
def sign_out(request: Request):
    user = sessions.resolve(request)
    if user:
        with connection() as conn:
            sessions.revoke(conn, user.session_id, "sign_out")
            log_event(conn, event_type="access", actor_type="user", actor_id=user.id, reason_code="sign_out")
    resp = RedirectResponse("/login", status_code=303)
    sessions.clear_cookie(resp)
    return resp


@router.get("/auth/verify")
def forward_auth(request: Request):
    """Traefik ForwardAuth for usage.accounting.* and commissions.accounting.*. Traefik returns a non-2xx
    answer to the browser as-is, so a 302 to the login page is what sends a stranger to sign in. The sibling
    hosts present the Domain-scoped SSO cookie, not the __Host- one."""
    user = sessions.resolve(request, allow_sso=True)
    if user is None:
        wanted = request.headers.get("x-forwarded-proto", "https") + "://" + request.headers.get("x-forwarded-host", settings.portal_host) \
            + request.headers.get("x-forwarded-uri", "/")
        return RedirectResponse(f"{settings.portal_base_url}/login?next={quote(wanted, safe='')}", status_code=302)
    return Response(status_code=200, headers={"X-Portal-User": user.email, "X-Portal-User-Id": str(user.id)})


# --- Freeze my access (NOT-5) -------------------------------------------------------------------

def issue_freeze_token(conn, user_id) -> str:
    raw = secrets.token_urlsafe(16)  # 128 bits
    conn.execute(
        "insert into freeze_token (user_id, token_hash, expires_at) values (%s, %s, %s)",
        (user_id, hashlib.sha256(raw.encode()).digest(), datetime.now(timezone.utc) + timedelta(days=14)),
    )
    return raw


@router.get("/freeze/{token}")
def freeze_confirm(request: Request, token: str):
    # GET shows a confirmation form; only the POST changes state (so a link scanner cannot freeze anyone).
    return render(request, "freeze.html", {"token": token})


@router.post("/freeze/{token}")
def freeze(request: Request, token: str):
    h = hashlib.sha256(token.encode()).digest()
    with connection() as conn:
        t = conn.execute("select * from freeze_token where token_hash = %s and used_at is null and expires_at > now() for update",
                         (h,)).fetchone()
        if t is None:
            return render(request, "freeze_done.html", {"ok": False}, status_code=400)
        conn.execute("update freeze_token set used_at = now() where id = %s", (t["id"],))
        conn.execute("update ap_user set status = 'frozen' where id = %s and status = 'active'", (t["user_id"],))
        sessions.revoke_all_for_user(conn, t["user_id"], "freeze_link")
        conn.execute("update passkey set status = 'revoked' where user_id = %s and status = 'active'", (t["user_id"],))
        seq = log_event(conn, event_type="security", actor_type="user", actor_id=t["user_id"], reason_code="freeze_my_access",
                        payload={"ip": sessions.client_ip(request)})
        conn.execute("insert into security_event (type, user_id, source, detail, event_seq) values ('freeze', %s, 'email_link', '{}', %s)",
                     (t["user_id"], seq))
    return render(request, "freeze_done.html", {"ok": True})


def _safe_next(next_path: str) -> str:
    """Only a path on this host, or an absolute URL on one of OUR hosts (the sibling apps under the
    portal host) — never anywhere else, so the login page cannot be used as an open redirect."""
    if not next_path:
        return "/"
    if next_path.startswith("/") and not next_path.startswith("//"):
        return next_path
    try:
        from urllib.parse import urlsplit
        parts = urlsplit(next_path)
        host = (parts.hostname or "").lower()
        if parts.scheme == "https" and (host == settings.portal_host or host.endswith("." + settings.portal_host)):
            return next_path
    except ValueError:
        pass
    return "/"
