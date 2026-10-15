"""Sign-in, sessions and the generic-404 posture, with a fake MaxxDash identity provider.

Self-contained fixtures (no reliance on the shared conftest) so this file runs alone.
"""
from __future__ import annotations

import os

os.environ.setdefault("ACCOUNTING_ENV", "test")
os.environ.setdefault("SECURE_COOKIES", "false")

import hashlib  # noqa: E402
from datetime import timedelta  # noqa: E402

import pytest  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

from accounting.db.migrate import apply_all  # noqa: E402
from accounting.db.pool import connection  # noqa: E402
from accounting.identity import maxxdash, sessions  # noqa: E402
from accounting.identity.maxxdash import Identity, SignInError, Tokens  # noqa: E402

SUBJECT = "11111111-1111-1111-1111-111111111111"


class FakeProvider:
    def __init__(self):
        self.banned = False
        self.granted = True
        self.refreshes = 0

    def password_grant(self, email, password, client_ip):
        if password != "right":
            raise SignInError(maxxdash.GENERIC)
        return Tokens("access-1", "refresh-1", 3600)

    def refresh(self, refresh_token):
        self.refreshes += 1
        if self.banned:
            raise SignInError("banned")
        return Tokens("access-2", f"refresh-{self.refreshes + 1}", 3600)

    def get_user(self, access_token):
        if self.banned:
            raise SignInError("banned")
        return {"id": SUBJECT, "email": "lenny@maxxortho.com"}

    def identity(self, access_token):
        return Identity(SUBJECT, "lenny@maxxortho.com", "Lenny Knobbs", "maxx_accounting",
                        ["accounting"] if self.granted else [])

    def recover(self, email):
        pass


@pytest.fixture()
def fake():
    p = FakeProvider()
    maxxdash.set_provider(p)  # type: ignore[arg-type]
    return p


@pytest.fixture(autouse=True)
def clean_db():
    apply_all()
    with connection() as c:
        c.execute("alter table event_log disable trigger all")
        c.execute("truncate signing_session, freeze_token, portal_session, enrollment_code, passkey, event_log, security_event, ap_user cascade")
        c.execute("alter table event_log enable trigger all")
    yield


@pytest.fixture()
def client(monkeypatch):
    # The feature packages are built separately; this file tests identity alone, so when a package is
    # absent the routers are stubbed. With every package present the real app is exercised.
    import sys, types
    from accounting.web import routers
    try:
        routers.all_routers()
    except ModuleNotFoundError:
        monkeypatch.setattr(routers, "all_routers", lambda: [])
        if "accounting.reports.dashboard" not in sys.modules:
            pkg = types.ModuleType("accounting.reports"); mod = types.ModuleType("accounting.reports.dashboard")
            mod.home_tiles = lambda user: {"greeting": "morning", "as_of": __import__("datetime").datetime.now(__import__("datetime").timezone.utc),
                                           "banners": [], "work": [], "systems": []}
            monkeypatch.setitem(sys.modules, "accounting.reports", pkg)
            monkeypatch.setitem(sys.modules, "accounting.reports.dashboard", mod)
    sys.modules.pop("accounting.main", None)
    from accounting.main import app
    return TestClient(app, base_url="http://accounting.test")


def _sign_in(client, password="right"):
    return client.post("/auth/session", data={"email": "Lenny@MaxxOrtho.com", "password": password, "next": "/"},
                       follow_redirects=False)


def test_login_page_renders_without_session(client, fake):
    r = client.get("/login")
    assert r.status_code == 200
    assert "Welcome back" in r.text
    assert 'name="password"' in r.text


def test_wrong_password_is_generic(client, fake):
    r = _sign_in(client, "wrong")
    assert r.status_code == 401
    assert "Check your credentials" in r.text
    assert sessions.COOKIE not in r.cookies


def test_sign_in_sets_host_only_session_cookie_and_logs(client, fake):
    r = _sign_in(client)
    assert r.status_code == 303 and r.headers["location"] == "/"
    cookie = r.headers["set-cookie"].lower()
    assert "httponly" in cookie and "samesite=strict" in cookie and "max-age" not in cookie and "expires" not in cookie
    with connection() as c:
        assert c.execute("select count(*) from portal_session where revoked_at is null").fetchone()["count"] == 1
        ev = c.execute("select * from event_log order by seq desc limit 1").fetchone()
        assert ev["event_type"] == "access" and ev["reason_code"] == "sign_in"
        # the cookie value itself is never stored
        raw = r.cookies[sessions.COOKIE]
        s = c.execute("select token_hash from portal_session").fetchone()
        assert bytes(s["token_hash"]) == hashlib.sha256(raw.encode()).digest()


def test_no_grant_means_generic_404(client, fake):
    fake.granted = False
    r = _sign_in(client)
    assert r.status_code == 404
    assert "Page not found" in r.text
    assert sessions.COOKIE not in r.cookies


def test_home_requires_session_and_unknown_paths_are_404(client, fake):
    assert client.get("/", follow_redirects=False).status_code == 303
    assert client.get("/invoices", follow_redirects=False).status_code == 404
    _sign_in(client)
    r = client.get("/", follow_redirects=False)
    assert r.status_code == 200
    assert "Lenny" in r.text
    assert r.headers["cache-control"] == "no-store"


def test_idle_and_absolute_expiry_revoke(client, fake):
    _sign_in(client)
    with connection() as c:
        c.execute("update portal_session set last_seen_at = now() - interval '31 minutes'")
    assert client.get("/", follow_redirects=False).status_code == 303
    with connection() as c:
        assert c.execute("select revoked_reason from portal_session").fetchone()["revoked_reason"] == "idle"
    _sign_in(client)
    with connection() as c:
        c.execute("update portal_session set absolute_expiry = now() - interval '1 second' where revoked_at is null")
    assert client.get("/", follow_redirects=False).status_code == 303


def test_hourly_revalidation_refreshes_and_catches_a_ban(client, fake):
    _sign_in(client)
    with connection() as c:
        c.execute("update portal_session set last_revalidated_at = now() - interval '61 minutes'")
    assert client.get("/", follow_redirects=False).status_code == 200
    assert fake.refreshes == 1
    fake.banned = True
    with connection() as c:
        c.execute("update portal_session set last_revalidated_at = now() - interval '61 minutes'")
    assert client.get("/", follow_redirects=False).status_code == 303
    with connection() as c:
        assert c.execute("select revoked_reason from portal_session where revoked_at is not null").fetchone()["revoked_reason"] == "revalidation_failed"


def test_device_change_revokes(client, fake):
    _sign_in(client)
    r = client.get("/", headers={"user-agent": "something-else"}, follow_redirects=False)
    assert r.status_code == 303


def test_sign_out_clears_cookie(client, fake):
    _sign_in(client)
    r = client.post("/auth/signout", follow_redirects=False)
    assert r.status_code == 303
    assert client.get("/", follow_redirects=False).status_code == 303


def test_forward_auth(client, fake):
    r = client.get("/auth/verify", headers={"x-forwarded-host": "usage.accounting.90ten.life", "x-forwarded-uri": "/x"},
                   follow_redirects=False)
    assert r.status_code == 302 and "/login?next=https%3A%2F%2Fusage.accounting.90ten.life%2Fx" in r.headers["location"]
    r = _sign_in(client)
    sso = r.cookies[sessions.SSO_COOKIE]
    # the sibling host presents only the SSO cookie
    bare = TestClient(client.app, base_url="http://accounting.test")
    r = bare.get("/auth/verify", cookies={sessions.SSO_COOKIE: sso})
    assert r.status_code == 200 and r.headers["x-portal-user"] == "lenny@maxxortho.com"
    # the SSO cookie alone does not open the portal itself
    assert bare.get("/", cookies={sessions.SSO_COOKIE: sso}, follow_redirects=False).status_code == 303


def test_roles_come_back_as_a_real_list(client, fake):
    _sign_in(client)
    with connection() as c:
        c.execute("update ap_user set roles = '{approver,cfo}'::portal_role[]")
    r = client.get("/", follow_redirects=False)
    assert r.status_code == 200 and "My approvals" in r.text


def test_safe_next_only_our_hosts():
    from accounting.identity.router import _safe_next
    assert _safe_next("/invoices") == "/invoices"
    assert _safe_next("//evil.com") == "/"
    assert _safe_next("https://evil.com/") == "/"
    assert _safe_next("https://usage.accounting.90ten.life/a") == "https://usage.accounting.90ten.life/a"


def test_freeze_link_is_post_only_and_single_use(client, fake):
    _sign_in(client)
    with connection() as c:
        uid = c.execute("select id from ap_user").fetchone()["id"]
        from accounting.identity.router import issue_freeze_token
        token = issue_freeze_token(c, uid)
    # GET only shows the confirmation; nothing changes
    assert client.get(f"/freeze/{token}").status_code == 200
    with connection() as c:
        assert c.execute("select status from ap_user").fetchone()["status"] == "active"
    r = client.post(f"/freeze/{token}")
    assert r.status_code == 200 and "Your access is frozen" in r.text
    with connection() as c:
        assert c.execute("select status from ap_user").fetchone()["status"] == "frozen"
        assert c.execute("select count(*) from portal_session where revoked_at is null").fetchone()["count"] == 0
        assert c.execute("select count(*) from security_event where type='freeze'").fetchone()["count"] == 1
    assert client.post(f"/freeze/{token}").status_code == 400


def test_rate_limit_after_repeated_failures(client, fake):
    from accounting.identity import ratelimit
    ratelimit.per_account._hits.clear(); ratelimit.per_address._hits.clear()
    last = None
    for _ in range(ratelimit.per_account.limit + 1):
        last = _sign_in(client, "wrong")
    assert last.status_code == 429 and "Retry-After" in last.headers
