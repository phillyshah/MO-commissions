"""Server-side portal sessions (ID-6).

The browser holds one opaque cookie, `__Host-ap_session` (host-only by construction, Secure, HttpOnly,
SameSite=Strict, no Max-Age so it dies with the browser). The row holds the Supabase refresh token,
encrypted, so the portal can silently refresh the one-hour access token until the session hits its
own limits: 30 minutes idle, 8 hours absolute, re-validated against GoTrue every hour (ID-3).

The portal never reads a `maxxdash-*` cookie and never sets a refresh cookie.
"""
from __future__ import annotations

import hashlib
import secrets
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from uuid import UUID

import psycopg
from fastapi import Request, Response

from accounting.config import settings
from accounting.db.pool import connection
from accounting.identity import crypto
from accounting.identity.maxxdash import Identity, SignInError, Tokens, provider
from accounting.lifecycle import log_event

COOKIE = "__Host-ap_session" if settings.secure_cookies else "ap_session"
# The sibling apps (usage.accounting.*, commissions.accounting.*) sit behind Traefik ForwardAuth, which
# calls /auth/verify with the browser's cookies for THEIR host. A __Host- cookie never travels there, so a
# second opaque token is issued for Domain=<portal_host>. It is never scoped wider than the portal host.
SSO_COOKIE = "ap_sso"


@dataclass
class CurrentUser:
    id: UUID
    subject: UUID
    email: str
    display_name: str
    roles: list[str]
    status: str
    maxxdash_role: str | None
    session_id: UUID
    device_hash: str

    def has(self, *roles: str) -> bool:
        return any(r in self.roles for r in roles)

    @property
    def is_admin(self) -> bool:
        return "admin" in self.roles


def device_hash(request: Request) -> str:
    ua = request.headers.get("user-agent", "")
    lang = request.headers.get("accept-language", "")
    return hashlib.sha256(f"{ua}|{lang}".encode()).hexdigest()[:32]


def client_ip(request: Request) -> str:
    xff = request.headers.get("x-forwarded-for")
    if xff:
        return xff.split(",")[0].strip()
    return request.client.host if request.client else "0.0.0.0"


def ip_country(request: Request) -> str | None:
    # Traefik/Cloudflare style header if present; otherwise unknown (never blocks).
    return request.headers.get("cf-ipcountry") or request.headers.get("x-country")


def upsert_ap_user(conn: psycopg.Connection, ident: Identity) -> dict:
    row = conn.execute(
        """insert into ap_user (subject, email, display_name, last_directory_sync)
           values (%s, %s, %s, now())
           on conflict (subject) do update set email = excluded.email,
                                               display_name = excluded.display_name,
                                               last_directory_sync = now()
           returning *""",
        (ident.subject, ident.email, ident.display_name),
    ).fetchone()
    return row


def create_session(conn: psycopg.Connection, request: Request, ident: Identity, tokens: Tokens) -> tuple[str, dict]:
    """Returns (raw_cookie_value, row). The row carries `sso_raw` (not a column) for the SSO cookie."""
    user = upsert_ap_user(conn, ident)
    if user["status"] == "disabled":
        raise SignInError("This account is disabled.", 403)
    raw = secrets.token_urlsafe(32)
    sso_raw = secrets.token_urlsafe(32)
    now = datetime.now(timezone.utc)
    row = conn.execute(
        """insert into portal_session (user_id, token_hash, sso_token_hash, refresh_token_enc, access_token_exp, device_hash,
                                       ip_country, absolute_expiry)
           values (%s,%s,%s,%s,%s,%s,%s,%s) returning *""",
        (
            user["id"],
            hashlib.sha256(raw.encode()).digest(),
            hashlib.sha256(sso_raw.encode()).digest(),
            crypto.encrypt(tokens.refresh_token, aad=str(user["id"]).encode()),
            now + timedelta(seconds=tokens.expires_in),
            device_hash(request),
            ip_country(request),
            now + timedelta(hours=settings.session_absolute_hours),
        ),
    ).fetchone()
    log_event(conn, event_type="access", actor_type="user", actor_id=user["id"],
              reason_code="sign_in", payload={"session": str(row["id"]), "device": row["device_hash"]})
    row["sso_raw"] = sso_raw
    return raw, row


def set_cookie(response: Response, raw: str, sso_raw: str | None = None) -> None:
    response.set_cookie(
        COOKIE, raw, httponly=True, secure=settings.secure_cookies, samesite="strict", path="/",
        # no max_age: session cookie, dies with the browser
    )
    if sso_raw:
        response.set_cookie(
            SSO_COOKIE, sso_raw, httponly=True, secure=settings.secure_cookies, samesite="lax", path="/",
            domain=settings.portal_host if settings.secure_cookies else None,
        )


def clear_cookie(response: Response) -> None:
    response.delete_cookie(COOKIE, path="/", secure=settings.secure_cookies, httponly=True, samesite="strict")
    response.delete_cookie(SSO_COOKIE, path="/", secure=settings.secure_cookies, httponly=True, samesite="lax",
                           domain=settings.portal_host if settings.secure_cookies else None)


def revoke(conn: psycopg.Connection, session_id: UUID | str, reason: str) -> None:
    conn.execute("update portal_session set revoked_at = now(), revoked_reason = %s where id = %s and revoked_at is null",
                 (reason, session_id))
    conn.execute("update signing_session set ended_at = now() where portal_session_id = %s and ended_at is null", (session_id,))


def revoke_all_for_user(conn: psycopg.Connection, user_id: UUID | str, reason: str) -> int:
    n = conn.execute("update portal_session set revoked_at = now(), revoked_reason = %s where user_id = %s and revoked_at is null",
                     (reason, user_id)).rowcount
    conn.execute("update signing_session set ended_at = now() where user_id = %s and ended_at is null", (user_id,))
    return n


def resolve(request: Request, *, allow_sso: bool = False) -> CurrentUser | None:
    """Load the session for this request, enforcing idle/absolute/device rules and the hourly re-check.
    Returns None when there is no valid session (callers then show the generic 404 or the login)."""
    raw = request.cookies.get(COOKIE)
    column = "token_hash"
    if not raw and allow_sso:
        raw = request.cookies.get(SSO_COOKIE)
        column = "sso_token_hash"
    if not raw:
        return None
    token_hash = hashlib.sha256(raw.encode()).digest()
    now = datetime.now(timezone.utc)
    with connection() as conn:
        # roles::text[] — psycopg returns the custom enum array as its literal text otherwise.
        s = conn.execute(
            f"""select s.*, u.subject, u.email, u.display_name, u.roles::text[] as roles, u.status
               from portal_session s join ap_user u on u.id = s.user_id
               where s.{column} = %s and s.revoked_at is null for update of s""",
            (token_hash,),
        ).fetchone()
        if s is None:
            return None
        if s["status"] == "disabled":
            revoke(conn, s["id"], "disabled")
            return None
        if now > s["absolute_expiry"]:
            revoke(conn, s["id"], "absolute_expiry")
            return None
        if now - s["last_seen_at"] > timedelta(minutes=settings.session_idle_minutes):
            revoke(conn, s["id"], "idle")
            return None
        if s["device_hash"] != device_hash(request):
            revoke(conn, s["id"], "device_changed")
            return None
        country = ip_country(request)
        if country and s["ip_country"] and country != s["ip_country"]:
            revoke(conn, s["id"], "country_changed")
            return None

        # Hourly re-validation with GoTrue (ban/disable shows up here), refreshing the access token.
        if now - s["last_revalidated_at"] > timedelta(minutes=settings.session_revalidate_minutes) or now >= s["access_token_exp"]:
            try:
                rt = crypto.decrypt(s["refresh_token_enc"], aad=str(s["user_id"]).encode()).decode()
                tokens = provider().refresh(rt)
                provider().get_user(tokens.access_token)
                ident = provider().identity(tokens.access_token)
                if "accounting" not in ident.portals:
                    revoke(conn, s["id"], "grant_removed")
                    return None
                conn.execute(
                    """update portal_session set refresh_token_enc = %s, access_token_exp = %s,
                                                 last_revalidated_at = now() where id = %s""",
                    (crypto.encrypt(tokens.refresh_token, aad=str(s["user_id"]).encode()),
                     now + timedelta(seconds=tokens.expires_in), s["id"]),
                )
                conn.execute("update ap_user set email = %s, display_name = %s, last_directory_sync = now() where id = %s",
                             (ident.email, ident.display_name, s["user_id"]))
            except SignInError:
                revoke(conn, s["id"], "revalidation_failed")
                return None
        conn.execute("update portal_session set last_seen_at = now() where id = %s", (s["id"],))
        return CurrentUser(
            id=s["user_id"], subject=s["subject"], email=s["email"], display_name=s["display_name"],
            roles=list(s["roles"] or []), status=s["status"], maxxdash_role=None,
            session_id=s["id"], device_hash=s["device_hash"],
        )


def background_check(user: CurrentUser) -> bool:
    """ID-3: before every signature and on each invoice page load, ask GoTrue again."""
    with connection() as conn:
        s = conn.execute("select * from portal_session where id = %s and revoked_at is null", (user.session_id,)).fetchone()
        if s is None:
            return False
        try:
            rt = crypto.decrypt(s["refresh_token_enc"], aad=str(s["user_id"]).encode()).decode()
            tokens = provider().refresh(rt)
            provider().get_user(tokens.access_token)
            conn.execute(
                "update portal_session set refresh_token_enc = %s, access_token_exp = %s, last_revalidated_at = now() where id = %s",
                (crypto.encrypt(tokens.refresh_token, aad=str(s["user_id"]).encode()),
                 datetime.now(timezone.utc) + timedelta(seconds=tokens.expires_in), s["id"]),
            )
            return True
        except SignInError:
            revoke(conn, s["id"], "background_check_failed")
            return False
