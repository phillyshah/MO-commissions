"""Real Postgres, real migrations, clean tables between tests.

Environment first: ACCOUNTING_ENV=test and a scratch object-store directory are set BEFORE anything
imports `accounting.config`, because Settings reads the environment at import. Then migrations apply
(idempotent — the runner records what it has done). Between tests every table is truncated. The audit
tables refuse UPDATE/DELETE by trigger (and `message`/`attachment`/`invoice` refuse DELETE), which is
exactly right in production and in the way here, so the fixture disables triggers on those tables for
the truncate and re-enables them — re-running the migrations would be far too slow per test.
"""
from __future__ import annotations

import os
import tempfile
import uuid

import pytest

os.environ.setdefault("ACCOUNTING_ENV", "test")
os.environ.setdefault("ACCOUNTING_DATABASE_URL", "postgresql://ap_test:ap_test@localhost/ap_test")
os.environ.setdefault("ACCOUNTING_STORE_DIR", tempfile.mkdtemp(prefix="ap-store-"))
os.environ.setdefault("PARSER_SOCKET", os.path.join(tempfile.mkdtemp(prefix="ap-sock-"), "absent.sock"))
os.environ.setdefault("SECURE_COOKIES", "0")
os.environ["ANTHROPIC_API_KEY"] = os.environ.get("ANTHROPIC_API_KEY") or "test-key-never-used"

from accounting.db import migrate  # noqa: E402
from accounting.db.pool import close_pool, connection  # noqa: E402

GUARDED = ("event_log", "approval_signature", "security_event", "message", "attachment", "invoice")
TABLES = (
    "ingest_state", "vendor_change", "anchor", "security_event", "notification", "quarantine_case", "approval_signature",
    "delegation", "approval_request", "matrix_version", "invoice_question", "invoice_correction", "invoice_render",
    "invoice_line", "invoice", "attachment", "message", "blocked_sender", "billing_platform_map", "vendor_domain",
    "vendor_bank_fingerprint", "vendor", "config_version", "freeze_token", "signing_session", "enrollment_code", "passkey",
    "portal_session", "event_log", "ap_user",
)


@pytest.fixture(scope="session", autouse=True)
def _migrated():
    migrate.apply_all()
    yield
    close_pool()


def _truncate_all() -> None:
    with connection() as conn:
        for t in GUARDED:
            conn.execute(f"alter table {t} disable trigger all")
        conn.execute("truncate " + ", ".join(TABLES) + " cascade")
        for t in GUARDED:
            conn.execute(f"alter table {t} enable trigger all")


@pytest.fixture(autouse=True)
def clean_db():
    _truncate_all()
    yield


@pytest.fixture
def db():
    """A transaction per use: `with db() as conn:`."""
    return connection


def make_user(conn, email: str, roles=("clerk",), display_name: str | None = None) -> dict:
    return conn.execute(
        "insert into ap_user (subject, email, display_name, roles) values (%s,%s,%s,%s::portal_role[]) returning *",
        (uuid.uuid4(), email, display_name or email.split("@")[0].title(), list(roles)),
    ).fetchone()


def make_vendor(conn, name: str, *, domains=(), country="US", remit=None, sap=None, status="active", **cols) -> dict:
    from accounting.vendors.normalize import name_key
    import json

    extra_cols = ", ".join(f", {k}" for k in cols)
    extra_vals = "".join(", %s" for _ in cols)
    v = conn.execute(
        f"insert into vendor (legal_name, name_key, country, remit_to, sap_vendor_number, status{extra_cols}) "
        f"values (%s,%s,%s,%s::jsonb,%s,%s{extra_vals}) returning *",
        (name, name_key(name), country, json.dumps(remit) if remit else None, sap, status, *cols.values()),
    ).fetchone()
    if domains:
        admin = conn.execute("select id from ap_user limit 1").fetchone() or make_user(conn, f"seed-{uuid.uuid4().hex[:6]}@maxxortho.com", ("admin",))
        for d in domains:
            conn.execute("insert into vendor_domain (vendor_id, pattern, status, added_by, approved_by) values (%s,%s,'approved',%s,%s)",
                         (v["id"], d, admin["id"], admin["id"]))
    return v
