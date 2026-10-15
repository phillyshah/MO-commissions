"""MaxxDash identity: the Supabase project behind dashboard.90ten.life, used for IDENTITY ONLY.

What this module does, and nothing else:
  - password grant     POST /auth/v1/token?grant_type=password   (same call MaxxDash's own login makes)
  - refresh            POST /auth/v1/token?grant_type=refresh_token
  - JWKS verification  GET  /auth/v1/.well-known/jwks.json       (asymmetric signing keys — ID-1)
  - user check         GET  /auth/v1/user                          (reflects a ban; a JWT does not — ID-3)
  - portal grant       GET  /rest/v1/portal_access?select=portal   (RLS: the user's own rows)
  - profile            GET  /rest/v1/user_profiles?select=role,display_name

The portal never holds the MaxxDash service-role key.
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any

import httpx
import jwt
from jwt import PyJWKClient

from accounting.config import settings

KNOWN_REASONS = {
    "email_not_confirmed": "This account's email address has not been confirmed yet. Ask a Maxx admin to confirm it for you.",
    "user_banned": "This account is locked. Ask a Maxx admin to unlock it.",
    "over_request_rate_limit": "Too many sign-in attempts. Wait a minute and try again.",
    "over_email_send_rate_limit": "Too many attempts. Wait a minute and try again.",
    "signup_disabled": "Sign-in is switched off for this site. Ask a Maxx admin.",
}
GENERIC = "Unable to sign in. Check your credentials and try again."


class SignInError(Exception):
    def __init__(self, message: str, status: int = 401):
        super().__init__(message)
        self.message = message
        self.status = status


@dataclass
class Tokens:
    access_token: str
    refresh_token: str
    expires_in: int


@dataclass
class Identity:
    subject: str
    email: str
    display_name: str
    maxxdash_role: str | None
    portals: list[str]


class MaxxDashIdentity:
    def __init__(self, base_url: str | None = None, anon_key: str | None = None):
        self.base_url = (base_url or settings.maxxdash_supabase_url).rstrip("/")
        self.anon_key = anon_key or settings.maxxdash_anon_key
        self._jwks: PyJWKClient | None = None
        self._jwks_loaded_at = 0.0
        self._client = httpx.Client(timeout=10.0)

    # -- transport ---------------------------------------------------------------
    def _headers(self, bearer: str | None = None) -> dict[str, str]:
        h = {"apikey": self.anon_key, "Content-Type": "application/json"}
        if bearer:
            h["Authorization"] = f"Bearer {bearer}"
        return h

    def password_grant(self, email: str, password: str, client_ip: str) -> Tokens:
        r = self._client.post(
            f"{self.base_url}/auth/v1/token?grant_type=password",
            headers={**self._headers(), "X-Forwarded-For": client_ip},
            json={"email": email.strip().lower(), "password": password},
        )
        if r.status_code != 200:
            raise SignInError(self._failure_message(r))
        body = r.json()
        return Tokens(body["access_token"], body["refresh_token"], int(body.get("expires_in", 3600)))

    def refresh(self, refresh_token: str) -> Tokens:
        r = self._client.post(
            f"{self.base_url}/auth/v1/token?grant_type=refresh_token",
            headers=self._headers(),
            json={"refresh_token": refresh_token},
        )
        if r.status_code != 200:
            raise SignInError("Your session has ended. Please sign in again.")
        body = r.json()
        return Tokens(body["access_token"], body["refresh_token"], int(body.get("expires_in", 3600)))

    @staticmethod
    def _failure_message(r: httpx.Response) -> str:
        try:
            body = r.json()
            code = body.get("error_code") or body.get("error") or ""
            return KNOWN_REASONS.get(code, GENERIC)
        except Exception:
            return GENERIC

    # -- verification -----------------------------------------------------------
    def _jwks_client(self) -> PyJWKClient:
        now = time.time()
        if self._jwks is None or now - self._jwks_loaded_at > settings.jwks_cache_seconds:
            self._jwks = PyJWKClient(f"{self.base_url}/auth/v1/.well-known/jwks.json", cache_keys=True)
            self._jwks_loaded_at = now
        return self._jwks

    def verify_jwt(self, access_token: str) -> dict[str, Any]:
        """Signature + expiry + audience. Raises jwt exceptions on failure.
        Pre-build check (spec §3): the project must be on an asymmetric signing key; an HS256
        token has no JWKS entry and fails here, which is the right outcome."""
        key = self._jwks_client().get_signing_key_from_jwt(access_token)
        return jwt.decode(
            access_token,
            key.key,
            algorithms=["ES256", "RS256"],
            audience="authenticated",
            options={"require": ["exp", "sub"]},
        )

    def get_user(self, access_token: str) -> dict[str, Any]:
        """GoTrue's own view of the account. 401/403 here means banned, deleted or expired."""
        r = self._client.get(f"{self.base_url}/auth/v1/user", headers=self._headers(access_token))
        if r.status_code != 200:
            raise SignInError("Your session is no longer valid. Please sign in again.")
        user = r.json()
        if user.get("banned_until"):
            raise SignInError(KNOWN_REASONS["user_banned"])
        return user

    def identity(self, access_token: str) -> Identity:
        claims = self.verify_jwt(access_token)
        user = self.get_user(access_token)
        profile = self._rest(access_token, "user_profiles", "select=role,display_name")
        grants = self._rest(access_token, "portal_access", "select=portal")
        prof = profile[0] if profile else {}
        return Identity(
            subject=claims["sub"],
            email=(user.get("email") or claims.get("email") or "").lower(),
            display_name=prof.get("display_name") or user.get("email") or claims["sub"],
            maxxdash_role=prof.get("role"),
            portals=[g["portal"] for g in grants],
        )

    def _rest(self, access_token: str, table: str, query: str) -> list[dict]:
        r = self._client.get(f"{self.base_url}/rest/v1/{table}?{query}", headers=self._headers(access_token))
        if r.status_code != 200:
            return []
        return r.json()

    def recover(self, email: str) -> None:
        self._client.post(f"{self.base_url}/auth/v1/recover", headers=self._headers(), json={"email": email.strip().lower()})


_provider: MaxxDashIdentity | None = None


def provider() -> MaxxDashIdentity:
    global _provider
    if _provider is None:
        _provider = MaxxDashIdentity()
    return _provider


def set_provider(p: MaxxDashIdentity) -> None:
    """Tests inject a fake."""
    global _provider
    _provider = p
