"""Package C: business days, notifications, reminders, anchors/verifier, dashboard tiles, approved
queue, month-end, invoice screens, admin dual control, upload.

Sibling packages (A, B) are replaced by small fake modules in sys.modules so these tests pass
whether or not the real modules are finished; what is asserted here is that Package C calls them
with the agreed signatures and acts on what they return.
"""
from __future__ import annotations

import csv
import email
import hashlib
import io
import json
import sys
import types
from datetime import date, datetime, timedelta, timezone
from uuid import uuid4

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from starlette.exceptions import HTTPException as StarletteHTTPException

from accounting import business_days as bd
from accounting.audit import anchors, verifier
from accounting.db.pool import connection
from accounting.lifecycle import log_event, transition
from accounting.notify import reminders, send, smtp, templates
from accounting.reports import dashboard, queue as approved_queue
from accounting.web import deps
from accounting.identity.sessions import CurrentUser
from accounting.web.templating import render

# --- fixtures local to this file (the shared conftest resets the database) -------------------------------

class Person:
    def __init__(self, id, email, display_name, roles):
        self.id, self.email, self.display_name, self.roles = id, email, display_name, roles

    def current(self, status: str = "active") -> CurrentUser:
        return CurrentUser(id=self.id, subject=self.id, email=self.email, display_name=self.display_name, roles=list(self.roles),
                           status=status, maxxdash_role=None, session_id=uuid4(), device_hash="test-device")


def make_user(conn, name: str, roles: list[str], manager_id=None) -> Person:
    uid = uuid4()
    email_addr = f"{name.lower().replace(' ', '.')}@maxxortho.com"
    conn.execute("insert into ap_user (id, subject, email, display_name, roles, manager_id) values (%s, %s, %s, %s, %s::portal_role[], %s)",
                 (uid, uuid4(), email_addr, name, roles, manager_id))
    return Person(uid, email_addr, name, roles)


@pytest.fixture()
def people():
    with connection() as conn:
        clerk = make_user(conn, "Carla Clerk", ["clerk"])
        supervisor = make_user(conn, "Sam Supervisor", ["supervisor"])
        controller = make_user(conn, "Corey Controller", ["controller"])
        controller2 = make_user(conn, "Chris Second", ["controller"])
        admin = make_user(conn, "Andy Admin", ["admin"])
        cfo = make_user(conn, "Fiona CFO", ["cfo"])
        approver = make_user(conn, "Al Approver", ["approver"], manager_id=controller.id)
        auditor = make_user(conn, "Ava Auditor", ["auditor"])
    return {"clerk": clerk, "supervisor": supervisor, "controller": controller, "controller2": controller2, "admin": admin,
            "cfo": cfo, "approver": approver, "auditor": auditor}


def make_vendor(conn, name: str, **kw):
    vid = uuid4()
    cols = {"id": vid, "legal_name": name, "name_key": name.lower().replace(" ", ""), **kw}
    conn.execute(f"insert into vendor ({', '.join(cols)}) values ({', '.join('%s' for _ in cols)})", list(cols.values()))
    return vid


def make_invoice(conn, *, state: str = "RECEIVED", vendor_id=None, gross_cents: int | None = 123456, received_at=None, **kw):
    """A row inserted directly in `state`; tests that need the audit trail call lifecycle.transition themselves."""
    iid = uuid4()
    cols = {"id": iid, "state": state, "vendor_id": vendor_id, "gross_cents": gross_cents, "currency": "USD",
            "received_at": received_at or datetime.now(timezone.utc), **kw}
    conn.execute(f"insert into invoice ({', '.join(cols)}) values ({', '.join('%s' for _ in cols)})", list(cols.values()))
    return iid


# --- app under test ---------------------------------------------------------------------------------

def build_app(user) -> TestClient:
    from accounting.reports.router import router as reports
    from accounting.web.admin import router as admin
    from accounting.web.invoices import router as invoices
    from accounting.web.uploads import router as uploads
    app = FastAPI()
    for r in (invoices, uploads, admin, reports):
        app.include_router(r)

    @app.exception_handler(StarletteHTTPException)
    async def http_error(request, exc):
        if exc.status_code == 404:
            return render(request, "404.html", {}, status_code=404)
        return render(request, "error.html", {"message": exc.detail or "Not allowed."}, status_code=exc.status_code)

    app.dependency_overrides[deps.current_user] = lambda: user
    return TestClient(app, follow_redirects=False)


def fake_module(monkeypatch, name: str, **attrs):
    mod = types.ModuleType(name)
    for k, v in attrs.items():
        setattr(mod, k, v)
    monkeypatch.setitem(sys.modules, name, mod)
    return mod


def mock_smtp(monkeypatch, fail_with: Exception | None = None):
    sent: list = []

    def _deliver(msg):
        if fail_with:
            raise fail_with
        sent.append(msg)
    monkeypatch.setattr(smtp, "_deliver", _deliver)
    monkeypatch.setattr(smtp, "configuration_problem", lambda: None)
    return sent


def add_signature(conn, invoice_id, signer_id, role="sponsor"):
    conn.execute(
        """insert into approval_signature (invoice_id, signer_id, signer_role, decision, content_hash, payload, payload_sha256,
                                           seal, seal_key_id, attestation_text_id)
           values (%s, %s, %s, 'approve', %s, '{}'::jsonb, %s, %s, 'test-key', 'att-v1')""",
        (invoice_id, signer_id, role, b"\x01" * 32, b"\x02" * 32, b"\x03" * 64))


# --- business days ------------------------------------------------------------------------------------

def test_federal_holidays_and_observed_shifts():
    h2026 = bd.federal_holidays(2026)
    assert date(2026, 1, 1) in h2026            # New Year's
    assert date(2026, 1, 19) in h2026           # MLK, third Monday
    assert date(2026, 2, 16) in h2026           # Presidents
    assert date(2026, 5, 25) in h2026           # Memorial, last Monday
    assert date(2026, 6, 19) in h2026           # Juneteenth (Friday)
    assert date(2026, 7, 3) in h2026            # Independence Day falls on Saturday -> observed Friday
    assert date(2026, 7, 4) not in h2026
    assert date(2026, 9, 7) in h2026            # Labor
    assert date(2026, 10, 12) in h2026          # Columbus
    assert date(2026, 11, 11) in h2026          # Veterans
    assert date(2026, 11, 26) in h2026          # Thanksgiving
    assert date(2026, 12, 25) in h2026          # Christmas
    # Christmas 2027 is a Saturday -> observed Friday 24th; New Year's 2028 is a Saturday -> observed 31 Dec 2027.
    h2027 = bd.federal_holidays(2027)
    assert date(2027, 12, 24) in h2027 and date(2027, 12, 31) in h2027
    assert date(2028, 1, 1) not in bd.federal_holidays(2028)
    assert not bd.is_business_day(date(2026, 7, 3))
    assert not bd.is_business_day(date(2026, 7, 5))  # Sunday
    assert bd.is_business_day(date(2026, 7, 6))


def test_add_and_between_business_days():
    # Thu 2 July 2026 + 2 business days skips the observed holiday (Fri 3rd) and the weekend -> Tue 7th.
    assert bd.add_business_days(date(2026, 7, 2), 2) == date(2026, 7, 7)
    assert bd.business_days_between(date(2026, 7, 2), date(2026, 7, 7)) == 2
    assert bd.business_days_between(date(2026, 7, 7), date(2026, 7, 2)) == -2
    assert bd.business_days_between(date(2026, 7, 7), date(2026, 7, 7)) == 0
    # A UTC moment late on a Friday evening is still Friday in New York; one business day later is Monday.
    friday_late_utc = datetime(2026, 10, 17, 2, 30, tzinfo=timezone.utc)  # Fri 16 Oct 22:30 New York
    assert bd.local_date(friday_late_utc) == date(2026, 10, 16)
    assert bd.add_business_days(friday_late_utc, 1).astimezone(bd.TZ).date() == date(2026, 10, 19)


# --- templates --------------------------------------------------------------------------------------------

FORBIDDEN = ("$", "/invoices/", "Acme", "INV-")


@pytest.mark.parametrize("type_", sorted(templates.BUILDERS))
def test_templates_carry_footer_and_freeze_link_and_nothing_sensitive(type_):
    ctx = {"count": 3, "oldest_business_days": 4, "delinquent_by_vendor": [{"vendor": "Acme Supply", "count": 2}], "hcp_count": 1,
           "text": "heartbeat text", "quarantine_open": 2, "escalations": 1, "freezes": 0, "access_changes": ["grant a@b.com by admin"],
           "heartbeat": "in range", "amount": "$1,234.00", "vendor": "Acme Supply", "invoice_number": "INV-99"}
    subject, body = templates.render(type_, ctx, "https://accounting.90ten.life/freeze/tok123")
    assert templates.FOOTER in body
    assert "Freeze my access: https://accounting.90ten.life/freeze/tok123" in body
    assert "You have 3 items waiting" in body or type_ in ("manager_copy", "enroll_alert", "freeze_link", "heartbeat_alert",
                                                            "controller_daily", "security_report")
    for bad in FORBIDDEN:
        if type_ == "digest" and bad == "Acme":
            continue  # REM-3 allows vendor names in the weekly digest's delinquent list
        assert bad not in body, f"{type_} leaked {bad!r}"
    assert "INV-99" not in body and "$1,234" not in body
    assert "/invoices/" not in body


def test_digest_lists_delinquent_vendors_and_hcp_as_count_only():
    _, body = templates.render("digest", {"count": 2, "oldest_business_days": 5,
                                          "delinquent_by_vendor": [{"vendor": "Acme Supply", "count": 1}], "hcp_count": 1}, None)
    assert "Acme Supply: 1 item" in body and "5 business days" in body
    assert "1 item involve a healthcare professional" in body
    assert "Freeze my access" not in body  # no user, no freeze line


# --- queue / flush ------------------------------------------------------------------------------------------

def test_queue_and_flush_success_records_message_id(people, monkeypatch):
    sent = mock_smtp(monkeypatch)
    with connection() as conn:
        nid = send.queue(conn, type="pending", user_id=people["approver"].id, context={"count": 1})
    assert send.flush() == 1
    with connection() as conn:
        n = conn.execute("select * from notification where id = %s", (nid,)).fetchone()
        tokens = conn.execute("select count(*) as n from freeze_token where user_id = %s", (people["approver"].id,)).fetchone()["n"]
    assert n["delivery_status"] == "sent" and n["attempts"] == 1 and n["sent_at"] is not None
    assert n["message_id_header"].startswith("<") and n["message_id_header"].endswith("@accounting.90ten.life>")
    assert tokens == 1
    msg = sent[0]
    assert msg["From"] == smtp.settings.email_from and msg["To"] == people["approver"].email
    assert msg["Message-ID"] == n["message_id_header"]
    body = msg.get_content()
    assert templates.FOOTER in body and "/freeze/" in body


def test_flush_failure_retries_at_most_three_times(people, monkeypatch):
    mock_smtp(monkeypatch, fail_with=OSError("relay down"))
    with connection() as conn:
        nid = send.queue(conn, type="reminder", user_id=people["approver"].id)
    for _ in range(5):
        assert send.flush() == 0
    with connection() as conn:
        n = conn.execute("select * from notification where id = %s", (nid,)).fetchone()
    assert n["delivery_status"] == "failed" and n["attempts"] == 3
    assert "relay down" in n["error"] and "Could not reach" in n["error"]


def test_queue_role_recipients_for_ap_wide_types(people):
    with connection() as conn:
        send.queue(conn, type="approved_for_ap", user_id=None, invoice_id=None)
        rows = conn.execute("select to_addr from notification order by to_addr").fetchall()
    assert [r["to_addr"] for r in rows] == sorted([people["clerk"].email, people["supervisor"].email])


def test_smtp_send_never_raises_and_requires_from_equals_user(monkeypatch):
    import dataclasses
    monkeypatch.setattr(smtp, "settings", dataclasses.replace(smtp.settings, smtp_password="x", email_from="other@90ten.life"))
    ok, reason, mid = smtp.send("a@b.com", "s", "t")
    assert ok is False and "SMTP_USER" in reason and mid.startswith("<")


# --- reminders -----------------------------------------------------------------------------------------------

def test_reminders_turn_intents_into_notifications_and_stamp_requests(people, monkeypatch):
    with connection() as conn:
        inv = make_invoice(conn, state="PENDING_APPROVAL")
        rid = conn.execute(
            "insert into approval_request (invoice_id, assignee_id, role, due_at) values (%s, %s, 'sponsor', now() - interval '1 day') returning id",
            (inv, people["approver"].id)).fetchone()["id"]
    now = datetime.now(timezone.utc)
    intents = [{"type": "reminder", "user_id": people["approver"].id, "invoice_id": inv, "request_id": rid},
               {"type": "manager_copy", "user_id": people["controller"].id, "invoice_id": inv, "request_id": rid}]
    calls = []

    def fake_escalate(conn, when):
        calls.append(when)
        return intents
    fake_module(monkeypatch, "accounting.routing.reassign", escalate_due=fake_escalate)
    assert reminders.run(now) == 2
    assert calls == [now]
    with connection() as conn:
        r = conn.execute("select reminded_at, escalated_manager_at, escalated_controller_at from approval_request where id = %s", (rid,)).fetchone()
        types_ = sorted(x["type"] for x in conn.execute("select type from notification").fetchall())
    assert r["reminded_at"] is not None and r["escalated_manager_at"] is not None and r["escalated_controller_at"] is None
    assert types_ == ["manager_copy", "reminder"]


def test_weekly_digest_and_controller_daily(people, monkeypatch):
    with connection() as conn:
        v = make_vendor(conn, "Acme Supply")
        inv = make_invoice(conn, state="PENDING_APPROVAL", vendor_id=v)
        conn.execute("insert into approval_request (invoice_id, assignee_id, role, due_at, created_at) values (%s, %s, 'sponsor', now() - interval '2 days', now() - interval '10 days')",
                     (inv, people["approver"].id))
        conn.execute("insert into quarantine_case (triggers) values ('{direct-sender}')")
        log_event(conn, event_type="admin", actor_type="user", actor_id=people["admin"].id, reason_code="grant",
                  payload={"user_email": "new@maxxortho.com"})
    fake_module(monkeypatch, "accounting.ingest.heartbeat", check_daily=lambda: {"ok": True, "n": 12, "text": "12 messages, in range"})
    assert reminders.weekly_digest() == 1
    assert reminders.controller_daily() == 1
    with connection() as conn:
        digest = conn.execute("select * from notification where type = 'digest'").fetchone()
        dailies = conn.execute("select * from notification where type = 'controller_daily' order by to_addr").fetchall()
    assert digest["user_id"] == people["approver"].id
    assert digest["context"]["delinquent_by_vendor"] == [{"vendor": "Acme Supply", "count": 1}]
    assert digest["context"]["oldest_business_days"] >= 6
    # every Controller gets the daily
    assert [d["to_addr"] for d in dailies] == sorted([people["controller"].email, people["controller2"].email])
    daily = dailies[0]
    assert daily["context"]["quarantine_open"] == 1 and daily["context"]["heartbeat"] == "12 messages, in range"
    assert any("grant new@maxxortho.com" in c for c in daily["context"]["access_changes"])


# --- anchors and verifier -------------------------------------------------------------------------------------

def test_anchor_local_backend_and_verifier_detects_tampering(people, monkeypatch, tmp_path):
    monkeypatch.setattr(anchors, "LOCAL_DIR", tmp_path / "anchors")
    with connection() as conn:
        inv = make_invoice(conn)
        transition(conn, inv, "AP_REVIEW", reason_code="manual_upload", actor_type="system")
        transition(conn, inv, "CAPTURED", reason_code="clerk_confirmed", actor_type="user", actor_id=people["clerk"].id)
        head_seq = conn.execute("select max(seq) as s from event_log").fetchone()["s"]
    row = anchors.anchor_now()
    assert row["status"] == "written" and row["head_seq"] == head_seq
    body = anchors.read_local(row["object_key"])
    assert body["head_seq"] == head_seq and body["head_hash_hex"] == bytes(row["head_hash"]).hex() and body["portal_version"]
    ok, detail = verifier.verify_since_last_anchor()
    assert ok, detail
    # Tamper with the anchored row the way only the database owner could.
    with connection() as conn:
        conn.execute("alter table event_log disable trigger all")
        conn.execute("update event_log set reason_code = 'clerk_edit' where seq = %s", (head_seq,))
        conn.execute("alter table event_log enable trigger all")
    ok, detail = verifier.verify_since_last_anchor()
    assert not ok and str(head_seq) in detail
    with connection() as conn:
        assert conn.execute("select count(*) as n from security_event where type = 'chain_broken'").fetchone()["n"] == 1


def test_sigv4_is_deterministic_and_carries_object_lock_headers():
    now = datetime(2026, 10, 14, 12, 0, 0, tzinfo=timezone.utc)
    h = anchors.sigv4_headers(method="PUT", host="b.s3.us-east-1.amazonaws.com", path="/2026-10-14/x.json", payload=b"{}",
                              headers={"x-amz-object-lock-mode": "COMPLIANCE", "x-amz-object-lock-retain-until-date": "2033-10-14T12:00:00Z"},
                              region="us-east-1", service="s3", access_key="AKIA", secret_key="secret", now=now)
    assert h["Authorization"].startswith("AWS4-HMAC-SHA256 Credential=AKIA/20261014/us-east-1/s3/aws4_request, SignedHeaders=")
    assert "x-amz-object-lock-mode" in h["Authorization"] and h["X-Amz-Date"] == "20261014T120000Z"
    assert h == anchors.sigv4_headers(method="PUT", host="b.s3.us-east-1.amazonaws.com", path="/2026-10-14/x.json", payload=b"{}",
                                      headers={"x-amz-object-lock-mode": "COMPLIANCE", "x-amz-object-lock-retain-until-date": "2033-10-14T12:00:00Z"},
                                      region="us-east-1", service="s3", access_key="AKIA", secret_key="secret", now=now)
    assert anchors.retention_until(datetime(2028, 2, 29, tzinfo=timezone.utc)) == datetime(2035, 3, 1, tzinfo=timezone.utc)


# --- home tiles ---------------------------------------------------------------------------------------------------

def test_home_tiles_by_role(people):
    with connection() as conn:
        make_invoice(conn, state="AP_REVIEW")
        make_invoice(conn, state="APPROVED")
        inv = make_invoice(conn, state="PENDING_APPROVAL")
        conn.execute("insert into approval_request (invoice_id, assignee_id, role, due_at, created_at) values (%s, %s, 'sponsor', now(), now() - interval '5 days')",
                     (inv, people["approver"].id))
        conn.execute("insert into quarantine_case (triggers) values ('{direct-sender}')")
    t = dashboard.home_tiles(people["approver"].current())
    assert t["greeting"] in ("morning", "afternoon", "evening") and isinstance(t["as_of"], datetime)
    assert [w["title"] for w in t["work"]] == ["Waiting for your approval"]
    assert t["work"][0]["count"] == 1 and t["work"][0]["tone"] == "warn"
    assert [s["title"] for s in t["systems"]] == ["Commission statements", "Usage tickets"]

    t = dashboard.home_tiles(people["clerk"].current())
    titles = {w["title"]: w["count"] for w in t["work"]}
    assert titles == {"Needs AP review": 1, "Unknown vendors": 0, "Approved, waiting for SAP entry": 1}
    assert [s["title"] for s in t["systems"]] == ["Invoices (AP)", "Commission statements", "Usage tickets", "Statement generator"]

    t = dashboard.home_tiles(people["controller"].current())
    titles = {w["title"]: w["count"] for w in t["work"]}
    assert titles["Quarantine"] == 1 and "Needs AP review" in titles
    # yesterday had no forwarded mail: the heartbeat banner shows on a weekday
    yesterday = (datetime.now(timezone.utc).astimezone(bd.TZ) - timedelta(days=1)).date()
    assert (yesterday.weekday() >= 5) or any("heartbeat" in b["text"].lower() for b in t["banners"])

    t = dashboard.home_tiles(people["auditor"].current())
    assert t["work"] == [] and t["systems"][0]["title"] == "Invoices (AP)"

    counts = dashboard.nav_counts_for(people["controller"].current())
    assert counts["quarantine"] == 1 and counts["break_glass"] is False


# --- approved queue ------------------------------------------------------------------------------------------------

def test_approved_queue_csv_and_record_sap_number(people):
    with connection() as conn:
        v = make_vendor(conn, "Acme Supply", sap_vendor_number="100234", default_gl="6100", default_cost_center="CC-1")
        older = make_invoice(conn, state="APPROVED", vendor_id=v, gross_cents=123456, tax_cents=1000, invoice_number="INV-7",
                             invoice_date=date(2026, 9, 1), received_at=datetime(2026, 9, 2, tzinfo=timezone.utc),
                             approved_at=datetime(2026, 9, 5, tzinfo=timezone.utc))
        newer = make_invoice(conn, state="APPROVED", vendor_id=v, gross_cents=5, approved_at=datetime(2026, 9, 9, tzinfo=timezone.utc))
        add_signature(conn, older, people["approver"].id)
        add_signature(conn, older, people["cfo"].id, "cfo")
    client = build_app(people["clerk"].current())
    r = client.get("/invoices/approved.csv")
    assert r.status_code == 200
    assert r.headers["content-disposition"].startswith('attachment; filename="MaxxDash-approved-invoices-')
    lines = r.text.strip().split("\n")
    rows = list(csv.reader(io.StringIO("\n".join(lines[:3]))))
    assert rows[0] == approved_queue.CSV_COLUMNS
    assert rows[1][0].startswith("AP-") and rows[1][1] == "Acme Supply" and rows[1][2] == "100234" and rows[1][3] == "INV-7"
    assert rows[1][7] == "1234.56" and rows[1][8] == "10.00" and rows[1][10] == "6100" and rows[1][11] == "CC-1"
    assert rows[1][13] == "Al Approver; Fiona CFO"
    assert rows[2][7] == "0.05" and rows[2][13] == "auto-released"
    assert lines[-1] == "# CONFIDENTIAL INFORMATION. NOT FOR DISTRIBUTION."
    assert lines[-2].startswith("# As of ")

    page = client.get("/invoices/approved")
    assert page.status_code == 200 and "As of" in page.text and "Download CSV" in page.text and "Al Approver" in page.text

    r = client.post(f"/invoices/{older}/sap", data={"sap_document_number": "   "})
    assert r.status_code == 400 and "Enter the SAP document number" in r.text
    r = client.post(f"/invoices/{older}/sap", data={"sap_document_number": "X" * 21})
    assert r.status_code == 400 and "at most 20 characters" in r.text
    r = client.post(f"/invoices/{older}/sap", data={"sap_document_number": "5100012345"})
    assert r.status_code == 303 and r.headers["location"] == "/invoices/approved"
    with connection() as conn:
        inv = conn.execute("select state, sap_document_number, entered_at, entered_by from invoice where id = %s", (older,)).fetchone()
        ev = conn.execute("select * from event_log where invoice_id = %s and reason_code = 'sap_entered'", (older,)).fetchone()
    assert inv["state"] == "ENTERED" and inv["sap_document_number"] == "5100012345" and inv["entered_at"] and inv["entered_by"] == people["clerk"].id
    assert ev["payload"] == {"sap_document_number": "5100012345"} and ev["to_state"] == "ENTERED"


# --- month end -------------------------------------------------------------------------------------------------------

def test_month_end_listing_and_sign_off(people):
    cutoff = date(2026, 9, 30)
    with connection() as conn:
        v = make_vendor(conn, "Acme Supply")
        listed = make_invoice(conn, state="PENDING_APPROVAL", vendor_id=v, received_at=datetime(2026, 9, 20, tzinfo=timezone.utc))
        entered_late = make_invoice(conn, state="ENTERED", vendor_id=v, received_at=datetime(2026, 9, 25, tzinfo=timezone.utc),
                                    entered_at=datetime(2026, 10, 3, tzinfo=timezone.utc))
        make_invoice(conn, state="ENTERED", vendor_id=v, received_at=datetime(2026, 9, 10, tzinfo=timezone.utc),
                     entered_at=datetime(2026, 9, 28, tzinfo=timezone.utc))  # entered before cutoff: not listed
        make_invoice(conn, state="RECEIVED", vendor_id=v, received_at=datetime(2026, 10, 2, tzinfo=timezone.utc))  # after cutoff
        make_invoice(conn, state="DIVERTED", vendor_id=v, received_at=datetime(2026, 9, 2, tzinfo=timezone.utc))  # never posted
    client = build_app(people["controller"].current())
    r = client.get(f"/reports/month-end?cutoff={cutoff}")
    assert r.status_code == 200 and "As of" in r.text and "Controller sign-off" in r.text
    assert r.text.count('class="clickable"') == 2 and str(listed) in r.text and str(entered_late) in r.text
    r = client.post("/reports/month-end/sign-off", data={"cutoff": str(cutoff)})
    assert r.status_code == 303
    with connection() as conn:
        ev = conn.execute("select * from event_log where event_type = 'month_end_signoff'").fetchone()
    assert ev["payload"]["cutoff"] == "2026-09-30" and ev["payload"]["count"] == 2
    assert set(ev["payload"]["invoices"]) == {str(listed), str(entered_late)}
    r = client.get(f"/reports/month-end?cutoff={cutoff}")
    assert "Signed off by Corey Controller" in r.text
    # A clerk may not sign off.
    assert build_app(people["clerk"].current()).post("/reports/month-end/sign-off", data={"cutoff": str(cutoff)}).status_code == 404


def test_reports_overview_and_audit_export(people):
    with connection() as conn:
        inv = make_invoice(conn, state="RECEIVED")
        transition(conn, inv, "AP_REVIEW", reason_code="manual_upload", actor_type="system")
        add_signature(conn, inv, people["approver"].id)
    client = build_app(people["supervisor"].current())
    r = client.get("/reports")
    assert r.status_code == 200 and "As of" in r.text and "Touchless rate" in r.text and "AP REVIEW" in r.text
    r = client.get(f"/reports/audit/{inv}.csv")
    # The fixture seal is random bytes: with the dev key present it FAILS, without one it is "not verified" — never "verified".
    assert r.status_code == 200 and "manual_upload" in r.text and ("not verified" in r.text or "FAILED" in r.text)
    assert r.text.strip().endswith("# CONFIDENTIAL INFORMATION. NOT FOR DISTRIBUTION.")
    r = client.get(f"/reports/audit/{inv}.json")
    body = json.loads(r.text)
    assert body["invoice"]["state"] == "AP_REVIEW" and len(body["events"]) == 1
    assert body["signatures"][0]["verification"] in ("not verified", "FAILED")
    assert build_app(people["clerk"].current()).get("/reports").status_code == 404


# --- invoice screens ---------------------------------------------------------------------------------------------------

def test_invoice_edit_voids_signatures_and_reroutes(people, monkeypatch):
    voided, routed = [], []
    fake_module(monkeypatch, "accounting.approvals.void",
                void_all_for_invoice=lambda conn, iid, reason, by: voided.append((iid, reason, by)) or 1)
    fake_module(monkeypatch, "accounting.routing.route", route=lambda conn, iid: routed.append(iid) or "CAPTURED")
    with connection() as conn:
        inv = make_invoice(conn, state="AP_REVIEW", invoice_number="inv 00123", gross_cents=100, content_hash=b"\x09" * 32,
                           field_confidence=json.dumps({"gross": 0.5, "invoice_number": 0.99}))
        add_signature(conn, inv, people["approver"].id)
    client = build_app(people["clerk"].current())
    detail = client.get(f"/invoices/{inv}")
    assert detail.status_code == 200 and "Correct the fields" in detail.text and 'class="low"' in detail.text and "As of" in detail.text
    r = client.post(f"/invoices/{inv}/edit", data={"vendor_name_raw": "Acme", "invoice_number": "INV-00123", "invoice_date": "2026-09-01",
                                                   "due_date": "", "currency": "usd", "gross": "1,234.56", "tax": "", "po_number": "PO9"})
    assert r.status_code == 303
    assert voided == [(inv, "field_edit", people["clerk"].id)] and routed == [inv]
    with connection() as conn:
        row = conn.execute("select * from invoice where id = %s", (inv,)).fetchone()
        corr = conn.execute("select diff from invoice_correction where invoice_id = %s", (inv,)).fetchone()
        ev = conn.execute("select reason_code from event_log where invoice_id = %s and event_type = 'transition'", (inv,)).fetchone()
    assert row["state"] == "CAPTURED" and row["gross_cents"] == 123456 and row["content_hash"] is None
    assert row["invoice_number_norm"] == "INV00123" and row["currency"] == "USD"  # Package A's DUP-1 key
    assert corr["diff"]["gross_cents"] == {"from": 100, "to": 123456} and ev["reason_code"] == "clerk_edit"
    # Bad money never becomes a float or a stack trace.
    with connection() as conn:
        inv2 = make_invoice(conn, state="AP_REVIEW")
    r = client.post(f"/invoices/{inv2}/edit", data={"gross": "twelve"})
    assert r.status_code == 400 and "Gross must be a number" in r.text


def test_invoice_list_answer_void_and_divert(people, monkeypatch):
    with connection() as conn:
        v = make_vendor(conn, "Acme Supply")
        inv = make_invoice(conn, state="QUESTIONED", vendor_id=v, invoice_number="Q-1")
        qid = conn.execute("insert into invoice_question (invoice_id, asked_by, question) values (%s, %s, 'Is this the right PO?') returning id",
                           (inv, people["approver"].id)).fetchone()["id"]
        other = make_invoice(conn, state="RECEIVED", vendor_id=v)
    client = build_app(people["clerk"].current())
    r = client.get("/invoices?vendor=acme&q=Q-1")
    assert r.status_code == 200 and "state-QUESTIONED" in r.text and "As of" in r.text
    r = client.post(f"/invoices/{inv}/answer/{qid}", data={"answer": "Yes, PO 9 covers it."})
    assert r.status_code == 303
    with connection() as conn:
        assert conn.execute("select state from invoice where id = %s", (inv,)).fetchone()["state"] == "PENDING_APPROVAL"
        n = conn.execute("select type, user_id from notification").fetchone()
    assert n["type"] == "pending" and n["user_id"] == people["approver"].id
    # Divert needs a reason; then the invoice is closed.
    assert client.post(f"/invoices/{other}/divert", data={"reason": "nope"}).status_code == 400
    assert client.post(f"/invoices/{other}/divert", data={"reason": "statement"}).status_code == 303
    with connection() as conn:
        assert conn.execute("select state, diverted_reason from invoice where id = %s", (other,)).fetchone() == {"state": "DIVERTED", "diverted_reason": "statement"}
    # Void: controller only, second person required and must be someone else.
    assert client.post(f"/invoices/{inv}/void", data={"reason": "x", "second_person_id": str(people["controller2"].id)}).status_code == 404
    ctrl = build_app(people["controller"].current())
    r = ctrl.post(f"/invoices/{inv}/void", data={"reason": "Entered twice", "second_person_id": str(people["controller"].id)})
    assert r.status_code == 400 and "second person" in r.text
    r = ctrl.post(f"/invoices/{inv}/void", data={"reason": "Entered twice", "second_person_id": str(people["controller2"].id)})
    assert r.status_code == 303
    with connection() as conn:
        assert conn.execute("select state from invoice where id = %s", (inv,)).fetchone()["state"] == "VOID"
        ev = conn.execute("select payload from event_log where invoice_id = %s and reason_code = 'controller_void' and event_type = 'transition'", (inv,)).fetchone()
    assert ev["payload"]["second_person"] == str(people["controller2"].id) and ev["payload"]["by"] == str(people["controller"].id)
    # Timeline is in plain words.
    r = ctrl.get(f"/invoices/{inv}")
    assert "Voided by the Controller with a second person." in r.text and "AP answered the question" in r.text


def test_frozen_user_can_look_but_not_act(people):
    with connection() as conn:
        inv = make_invoice(conn, state="RECEIVED")
    client = build_app(people["clerk"].current(status="frozen"))
    assert client.get(f"/invoices/{inv}").status_code == 200
    r = client.post(f"/invoices/{inv}/divert", data={"reason": "statement"})
    assert r.status_code == 403 and "frozen" in r.text


# --- admin -------------------------------------------------------------------------------------------------------------

def test_admin_dual_control(people):
    admin = build_app(people["admin"].current())
    controller = build_app(people["controller"].current())
    # Dual control unset: the admin's proposal takes effect at once, and is logged as such.
    r = admin.post("/admin/config/propose", data={"key": "step_up_amount_cents", "value": "2500000", "note": "initial"})
    assert r.status_code == 303
    with connection() as conn:
        cur = conn.execute("select value from config_current where key = 'step_up_amount_cents'").fetchone()
        ev = conn.execute("select reason_code from event_log where event_type = 'admin' order by seq desc limit 1").fetchone()
    assert cur["value"] == 2500000 and ev["reason_code"] == "config_auto_approved"
    # A non-admin's proposal is never auto-approved.
    r = controller.post("/admin/config/propose", data={"key": "soft_signals_enabled", "value": "true"})
    assert r.status_code == 303
    with connection() as conn:
        assert conn.execute("select value from config_current where key = 'soft_signals_enabled'").fetchone() is None
    # Switch dual control on; from here the admin's own proposals wait for someone else.
    admin.post("/admin/config/propose", data={"key": "dual_control_enabled", "value": "true"})
    r = admin.post("/admin/config/propose", data={"key": "step_up_amount_cents", "value": "5000000"})
    with connection() as conn:
        pending = conn.execute("select id from config_version where key = 'step_up_amount_cents' and approved_at is null").fetchone()
        assert conn.execute("select value from config_current where key = 'step_up_amount_cents'").fetchone()["value"] == 2500000
    r = admin.post(f"/admin/config/{pending['id']}/approve")
    assert r.status_code == 403 and "cannot also approve" in r.text
    r = controller.post(f"/admin/config/{pending['id']}/approve")
    assert r.status_code == 303
    with connection() as conn:
        row = conn.execute("select * from config_version where id = %s", (pending["id"],)).fetchone()
    assert row["approved_by"] == people["controller"].id and row["effective_from"] > datetime.now(timezone.utc)  # next business day
    assert bd.is_business_day(row["effective_from"])
    page = admin.get("/admin/config")
    assert page.status_code == 200 and "Dual control is on" in page.text and "As of" in page.text


def test_admin_users_matrix_breakglass_and_grants(people):
    admin = build_app(people["admin"].current())
    target = people["approver"]
    r = admin.post(f"/admin/users/{target.id}", data={"roles": ["approver", "compliance"], "manager_id": str(people["controller"].id), "backup_id": ""})
    assert r.status_code == 303
    r = admin.post(f"/admin/users/{target.id}/enrollment-code")
    assert r.status_code == 200 and "Passkey enrollment code for Al Approver" in r.text
    import re
    code = re.search(r'letter-spacing:\.2em">(\d{6})<', r.text).group(1)
    with connection() as conn:
        u = conn.execute("select roles, manager_id from ap_user where id = %s", (target.id,)).fetchone()
        ec = conn.execute("select * from enrollment_code where user_id = %s", (target.id,)).fetchone()
    from accounting.web.admin import roles_list
    assert sorted(roles_list(u["roles"])) == ["approver", "compliance"] and u["manager_id"] == people["controller"].id
    assert bytes(ec["code_hash"]) == hashlib.sha256(code.encode()).digest() and ec["issued_by"] == people["admin"].id
    assert admin.post(f"/admin/users/{target.id}/freeze").status_code == 303
    with connection() as conn:
        assert conn.execute("select status from ap_user where id = %s", (target.id,)).fetchone()["status"] == "frozen"
    assert admin.post(f"/admin/users/{target.id}/unfreeze").status_code == 303
    with connection() as conn:
        assert conn.execute("select status from ap_user where id = %s", (target.id,)).fetchone()["status"] == "active"
        reasons = [r["reason_code"] for r in conn.execute("select reason_code from event_log where event_type = 'admin' order by seq").fetchall()]
    assert reasons == ["role_change", "enrollment_code_issued", "freeze", "unfreeze"]
    # Matrix: shape is validated; the proposer may not approve under dual control.
    r = admin.post("/admin/matrix/propose", data={"rules": "not json"})
    assert r.status_code == 400 and "not valid JSON" in r.text
    r = admin.post("/admin/matrix/propose", data={"rules": json.dumps([{"min_cents": 0, "roles": []}])})
    assert r.status_code == 400 and "roles must be" in r.text
    r = admin.post("/admin/matrix/propose", data={"rules": json.dumps([{"min_cents": 0, "max_cents": None, "roles": ["sponsor"], "sequential": False}])})
    assert r.status_code == 303
    # Break-glass: CFO only (and admin); banner on every page; off again.
    assert build_app(people["controller"].current()).post("/admin/break-glass", data={"on": "true", "reason": "x"}).status_code == 404
    cfo = build_app(people["cfo"].current())
    assert cfo.post("/admin/break-glass", data={"on": "true", "reason": "Auditors on site"}).status_code == 303
    page = admin.get("/invoices")
    assert "BREAK-GLASS MODE IS ON" in page.text
    assert cfo.post("/admin/break-glass", data={"on": "false", "reason": "done"}).status_code == 303
    assert "BREAK-GLASS MODE IS ON" not in admin.get("/invoices").text
    grants = admin.get("/admin/grants")
    assert grants.status_code == 200 and "role change" in grants.text and "break glass on" in grants.text
    assert admin.get("/admin/security").status_code == 200 and admin.get("/admin/delegations").status_code == 200
    assert build_app(people["clerk"].current()).get("/admin").status_code == 404


# --- upload ------------------------------------------------------------------------------------------------------------

def test_upload_builds_message_and_calls_intake(people, monkeypatch):
    seen = {}
    new_id = uuid4()

    def fake_ingest(raw, *, imap_uid, source):
        seen.update(raw=raw, imap_uid=imap_uid, source=source)
        return [new_id]
    fake_module(monkeypatch, "accounting.ingest.intake", ingest_raw_message=fake_ingest)
    client = build_app(people["clerk"].current())
    assert client.get("/upload").status_code == 200
    r = client.post("/upload", files={"file": ("scan 1.pdf", b"%PDF-1.4 fake", "application/pdf")}, data={"note": "From the mail room"})
    assert r.status_code == 303 and r.headers["location"] == f"/invoices/{new_id}"
    assert seen["imap_uid"] is None and seen["source"] == "upload"
    import email.policy
    msg = email.message_from_bytes(seen["raw"], policy=email.policy.default)
    assert msg["From"] == people["clerk"].email and msg["Subject"] == "Manual upload: scan 1.pdf"
    parts = [p for p in msg.walk() if p.get_filename()]
    assert parts[0].get_filename() == "scan 1.pdf" and parts[0].get_payload(decode=True) == b"%PDF-1.4 fake"
    assert "From the mail room" in msg.get_body(preferencelist=("plain",)).get_content()
    r = client.post("/upload", files={"file": ("notes.docx", b"x", "application/octet-stream")})
    assert r.status_code == 400 and "PDF, PNG or JPG" in r.text
