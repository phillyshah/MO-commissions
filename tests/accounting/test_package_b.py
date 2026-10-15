"""Package B: routing, approvals, signing, quarantine.

Everything runs against the real schema in the local Postgres. Package A and C functions are stubbed
(notify queue) or fall back (business days). A software authenticator built on `cryptography` plays the
passkey so the WebAuthn ceremonies are exercised end to end, not just their option generation.
"""
from __future__ import annotations

import hashlib
import json
import os
import secrets
import uuid
from datetime import date, datetime, timedelta, timezone

import cbor2
import pytest
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import ec
from webauthn.helpers import bytes_to_base64url

from accounting.approvals import payload as payload_mod
from accounting.approvals import sign, signing_service, signing_sessions, void, webauthn_flow
from accounting.config import settings
from accounting.db.pool import connection
from accounting.identity.sessions import CurrentUser
from accounting.lifecycle import canonical
from accounting.quarantine import service as quarantine
from accounting.routing import auto_release, matrix, reassign, route

# The dev seal signer switches on only when is_test AND a PEM are both set. Settings is a frozen
# dataclass read at import (the shared conftest imported it first), so when the environment did not
# carry DEV_SEAL_PRIVATE_KEY_PEM a key is generated here and placed on the live settings object.
if not settings.dev_seal_private_key_pem:
    from cryptography.hazmat.primitives import serialization as _s

    _pem = ec.generate_private_key(ec.SECP256R1()).private_bytes(_s.Encoding.PEM, _s.PrivateFormat.PKCS8, _s.NoEncryption()).decode()
    os.environ["DEV_SEAL_PRIVATE_KEY_PEM"] = _pem
    object.__setattr__(settings, "dev_seal_private_key_pem", _pem)


@pytest.fixture()
def notify(monkeypatch):
    """A stand-in for Package C's accounting.notify.send that records every queue() call.
    Installed in sys.modules so the lazy `from accounting.notify import send` in the code under test
    finds it, whatever state the real module is in."""
    import sys
    import types

    calls: list[dict] = []
    mod = types.ModuleType("accounting.notify.send")

    def queue(conn, *, type, user_id=None, invoice_id=None, to=None, context=None):
        calls.append({"type": type, "user_id": user_id, "invoice_id": invoice_id, "to": to, "context": context})
        return len(calls)

    mod.queue = queue
    pkg = types.ModuleType("accounting.notify")
    pkg.send = mod
    monkeypatch.setitem(sys.modules, "accounting.notify", pkg)
    monkeypatch.setitem(sys.modules, "accounting.notify.send", mod)
    return calls


@pytest.fixture()
def background_ok(monkeypatch):
    """ID-3 directory re-check, answered 'still fine' without talking to Supabase."""
    from accounting.identity import sessions
    monkeypatch.setattr(sessions, "background_check", lambda user: True)


pytestmark = pytest.mark.usefixtures("notify", "background_ok")

STEP_UP = settings.default_step_up_amount_cents  # $25,000.00


# --- builders -------------------------------------------------------------------------------------

def mk_user(conn, name: str, roles: list[str], *, status="active", manager=None, backup=None) -> dict:
    row = conn.execute(
        """insert into ap_user (subject, email, display_name, roles, status, manager_id, backup_id)
           values (%s, %s, %s, %s::portal_role[], %s, %s, %s)
           returning id, subject, email, display_name, roles::text[] as roles, status, manager_id, backup_id""",
        (uuid.uuid4(), f"{name.lower().replace(' ', '.')}.{secrets.token_hex(3)}@maxxortho.com", name, roles, status, manager, backup),
    ).fetchone()
    sess = conn.execute(
        """insert into portal_session (user_id, token_hash, refresh_token_enc, access_token_exp, device_hash, absolute_expiry)
           values (%s, %s, %s, now() + interval '1 hour', 'dev-hash', now() + interval '8 hours') returning id""",
        (row["id"], secrets.token_bytes(32), b"x"),
    ).fetchone()
    row["session_id"] = sess["id"]
    return row


def cu(row: dict) -> CurrentUser:
    return CurrentUser(id=row["id"], subject=row["subject"], email=row["email"], display_name=row["display_name"],
                       roles=list(row["roles"]), status=row["status"], maxxdash_role=None, session_id=row["session_id"],
                       device_hash="dev-hash")


def mk_vendor(conn, sponsor=None, backup=None, **kw) -> dict:
    cols = {"legal_name": kw.pop("legal_name", "Acme Supplies LLC"), "name_key": "acme", "sap_vendor_number": kw.pop("sap", secrets.token_hex(4)),
            "default_sponsor_id": sponsor, "backup_sponsor_id": backup, "default_gl": kw.pop("gl", "6100"),
            "default_cost_center": kw.pop("cost_center", "CC10"), "status": "active",
            "created_at": kw.pop("created_at", datetime.now(timezone.utc) - timedelta(days=400))}
    cols.update(kw)
    names = ", ".join(cols)
    return conn.execute(f"insert into vendor ({names}) values ({', '.join(['%s'] * len(cols))}) returning *", list(cols.values())).fetchone()


def mk_invoice(conn, vendor=None, gross=120_000, *, state="CAPTURED", invoice_date=None, **kw) -> dict:
    cols = {"state": state, "vendor_id": vendor["id"] if vendor else None, "invoice_number": kw.pop("invoice_number", f"INV-{secrets.token_hex(3)}"),
            "invoice_date": invoice_date or date(2026, 10, 1), "due_date": date(2026, 10, 31), "currency": "USD", "gross_cents": gross,
            "tax_cents": 0, "file_sha256": hashlib.sha256(secrets.token_bytes(8)).digest(), "received_at": datetime.now(timezone.utc),
            "bank_last4": "4321"}
    cols.update(kw)
    names = ", ".join(cols)
    return conn.execute(f"insert into invoice ({names}) values ({', '.join(['%s'] * len(cols))}) returning *", list(cols.values())).fetchone()


def mk_passkey(conn, user) -> dict:
    return conn.execute("insert into passkey (user_id, credential_id, public_key, label) values (%s, %s, %s, 'test') returning *",
                        (user["id"], secrets.token_bytes(16), b"\x01\x02")).fetchone()


def open_session(conn, user_row) -> dict:
    pk = mk_passkey(conn, user_row)
    return signing_sessions.start(conn, cu(user_row), pk["id"], user_row["session_id"], "dev-hash")


def requests_for(conn, invoice_id):
    return conn.execute("select * from approval_request where invoice_id = %s order by created_at", (invoice_id,)).fetchall()


def state_of(conn, invoice_id) -> str:
    return conn.execute("select state from invoice where id = %s", (invoice_id,)).fetchone()["state"]


# --- a software passkey ---------------------------------------------------------------------------

class SoftAuthenticator:
    """Enough of a FIDO2 authenticator to satisfy py_webauthn: ES256 key, 'none' attestation, UV flag set."""

    def __init__(self):
        self.key = ec.generate_private_key(ec.SECP256R1())
        self.credential_id = secrets.token_bytes(32)
        self.sign_count = 0
        self.origin = settings.portal_base_url
        self.rp_hash = hashlib.sha256(settings.rp_id.encode()).digest()

    def _cose_key(self) -> bytes:
        nums = self.key.public_key().public_numbers()
        return cbor2.dumps({1: 2, 3: -7, -1: 1, -2: nums.x.to_bytes(32, "big"), -3: nums.y.to_bytes(32, "big")})

    def _client_data(self, typ: str, challenge_b64url: str) -> bytes:
        return json.dumps({"type": typ, "challenge": challenge_b64url, "origin": self.origin, "crossOrigin": False}).encode()

    def register(self, options: dict) -> dict:
        self.sign_count += 1
        client_data = self._client_data("webauthn.create", options["challenge"])
        attested = b"\x00" * 16 + len(self.credential_id).to_bytes(2, "big") + self.credential_id + self._cose_key()
        auth_data = self.rp_hash + bytes([0x45]) + self.sign_count.to_bytes(4, "big") + attested
        att_obj = cbor2.dumps({"fmt": "none", "attStmt": {}, "authData": auth_data})
        return {"id": bytes_to_base64url(self.credential_id), "rawId": bytes_to_base64url(self.credential_id), "type": "public-key",
                "response": {"clientDataJSON": bytes_to_base64url(client_data), "attestationObject": bytes_to_base64url(att_obj),
                             "transports": ["internal"]}}

    def authenticate(self, options: dict, *, challenge_override: str | None = None) -> dict:
        self.sign_count += 1
        client_data = self._client_data("webauthn.get", challenge_override or options["challenge"])
        auth_data = self.rp_hash + bytes([0x05]) + self.sign_count.to_bytes(4, "big")
        sig = self.key.sign(auth_data + hashlib.sha256(client_data).digest(), ec.ECDSA(hashes.SHA256()))
        return {"id": bytes_to_base64url(self.credential_id), "rawId": bytes_to_base64url(self.credential_id), "type": "public-key",
                "response": {"clientDataJSON": bytes_to_base64url(client_data), "authenticatorData": bytes_to_base64url(auth_data),
                             "signature": bytes_to_base64url(sig), "userHandle": None}}


def issue_code(conn, user, controller, code="ABCD-1234") -> str:
    conn.execute("insert into enrollment_code (user_id, code_hash, issued_by, expires_at) values (%s, %s, %s, now() + interval '1 day')",
                 (user["id"], hashlib.sha256(code.encode()).digest(), controller["id"]))
    return code


# --- matrix ----------------------------------------------------------------------------------------

def test_required_roles_default_step_up_and_hcp():
    base = {"gross_cents": 100_00, "gl": "6100", "cost_center": "CC10", "hcp": False, "step_up_cents": STEP_UP}
    assert matrix.required_roles(matrix.DEFAULT_RULES, base) == [("sponsor", 1)]
    assert matrix.required_roles(matrix.DEFAULT_RULES, {**base, "gross_cents": STEP_UP}) == [("sponsor", 1), ("cfo", 1)]
    assert ("compliance", 1) in matrix.required_roles(matrix.DEFAULT_RULES, {**base, "hcp": True})
    # A config step-up lower than the rule band still forces the CFO (D3/D4 are not negotiable by the matrix).
    assert ("cfo", 1) in matrix.required_roles(matrix.DEFAULT_RULES, {**base, "gross_cents": 1_000_00, "step_up_cents": 500_00})
    # An empty rule set cannot remove the sponsor.
    assert matrix.required_roles([], base) == [("sponsor", 1)]


def test_required_roles_coding_rules_and_tiers():
    rules = matrix.DEFAULT_RULES + [{"gl_prefix": "6", "cost_center": "CC10", "roles": ["quality"], "tier": 2},
                                    {"gl_prefix": "7", "roles": ["controller"], "tier": 1}]
    got = matrix.required_roles(rules, {"gross_cents": 500_00, "gl": "6100", "cost_center": "CC10", "hcp": False, "step_up_cents": STEP_UP})
    assert got == [("sponsor", 1), ("quality", 2)]  # sorted by tier; GL 7 rule does not match
    got = matrix.required_roles(rules, {"gross_cents": 500_00, "gl": "6100", "cost_center": "CC99", "hcp": False, "step_up_cents": STEP_UP})
    assert ("quality", 2) not in got  # cost center must match too


def test_matrix_dual_control_and_default():
    with connection() as c:
        a, b = mk_user(c, "Admin One", ["admin"]), mk_user(c, "Admin Two", ["admin"])
        assert matrix.current(c) == (None, matrix.DEFAULT_RULES)
        with pytest.raises(matrix.MatrixError):
            matrix.propose(c, [{"roles": ["sponsor"]}], a["id"])  # no band, no coding
        vid = matrix.propose(c, [{"min_cents": 0, "max_cents": None, "roles": ["sponsor", "cfo"], "tier": 1}], a["id"])
        with pytest.raises(matrix.MatrixError):
            matrix.approve(c, vid, a["id"])  # SOD-4
        row = matrix.approve(c, vid, b["id"])
        assert row["effective_from"] > datetime.now(timezone.utc)  # next business day
        assert matrix.current(c)[0] is None  # not yet in force
        c.execute("update matrix_version set effective_from = now() - interval '1 minute' where id = %s", (vid,))
        assert matrix.current(c)[0] == vid


# --- routing --------------------------------------------------------------------------------------

def test_route_parallel_chain_above_step_up(notify):
    with connection() as c:
        sponsor, cfo = mk_user(c, "Sue Sponsor", ["approver"]), mk_user(c, "Carl Cfo", ["cfo"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, STEP_UP)
        assert route.route(c, inv["id"]) == "PENDING_APPROVAL"
        reqs = requests_for(c, inv["id"])
        assert {r["assignee_id"] for r in reqs} == {sponsor["id"], cfo["id"]}
        assert {r["tier"] for r in reqs} == {1}  # APR-5: parallel
        assert all(r["required_amount_cents"] == STEP_UP for r in reqs)
        assert c.execute("select content_hash from invoice where id = %s", (inv["id"],)).fetchone()["content_hash"] is not None
    assert sorted(n["type"] for n in notify) == ["pending", "pending"]


def test_route_sod_excludes_corrector_then_backup_then_exhausted():
    with connection() as c:
        sponsor, backup = mk_user(c, "Sue Sponsor", ["approver"]), mk_user(c, "Bob Backup", ["approver"])
        v = mk_vendor(c, sponsor["id"], backup["id"])
        inv = mk_invoice(c, v, 500_00)
        c.execute("insert into invoice_correction (invoice_id, user_id, diff) values (%s, %s, '{}')", (inv["id"], sponsor["id"]))
        assert route.route(c, inv["id"]) == "PENDING_APPROVAL"
        (req,) = requests_for(c, inv["id"])
        assert req["assignee_id"] == backup["id"]

        inv2 = mk_invoice(c, v, 500_00)
        c.execute("insert into invoice_correction (invoice_id, user_id, diff) values (%s, %s, '{}')", (inv2["id"], sponsor["id"]))
        c.execute("insert into invoice_correction (invoice_id, user_id, diff) values (%s, %s, '{}')", (inv2["id"], backup["id"]))
        assert route.route(c, inv2["id"]) == "AP_REVIEW"
        assert requests_for(c, inv2["id"]) == []
        ev = c.execute("select reason_code from event_log where invoice_id = %s and to_state = 'AP_REVIEW'", (inv2["id"],)).fetchone()
        assert ev["reason_code"] == "reassign_exhausted"


def test_route_delegation_redirects_and_delegate_cannot_approve_own_capture():
    with connection() as c:
        sponsor, dele, backup = mk_user(c, "Sue Sponsor", ["approver"]), mk_user(c, "Dan Delegate", ["approver"]), mk_user(c, "Bob Backup", ["approver"])
        c.execute("insert into delegation (from_user_id, to_user_id, starts_on, ends_on) values (%s, %s, current_date - 1, current_date + 5)",
                  (sponsor["id"], dele["id"]))
        v = mk_vendor(c, sponsor["id"], backup["id"])
        inv = mk_invoice(c, v, 500_00)
        route.route(c, inv["id"])
        (req,) = requests_for(c, inv["id"])
        assert req["assignee_id"] == dele["id"]
        inv2 = mk_invoice(c, v, 500_00)
        c.execute("insert into invoice_correction (invoice_id, user_id, diff) values (%s, %s, '{}')", (inv2["id"], dele["id"]))
        route.route(c, inv2["id"])
        (req2,) = requests_for(c, inv2["id"])
        assert req2["assignee_id"] == backup["id"]


def test_route_overlap_one_person_signs_once_and_hard_stops():
    with connection() as c:
        cfo = mk_user(c, "Carl Cfo", ["cfo", "approver"])
        v = mk_vendor(c, cfo["id"])
        inv = mk_invoice(c, v, STEP_UP + 1)
        route.route(c, inv["id"])
        (req,) = requests_for(c, inv["id"])  # APR-6: sponsor and cfo collapse into one request
        assert set(req["role"].split(",")) == {"sponsor", "cfo"}
        # VEN-5 / QMSR hard stops
        q = mk_vendor(c, cfo["id"], legal_name="Quality Parts", sap="2", quality_critical=True, asl_status="conditional")
        assert route.route(c, mk_invoice(c, q, 100_00)["id"]) == "AP_REVIEW"
        q2 = mk_vendor(c, cfo["id"], legal_name="Quality Parts 2", sap="3", quality_critical=True, asl_status="approved")
        assert route.route(c, mk_invoice(c, q2, 100_00, po_number=None)["id"]) == "AP_REVIEW"
        assert route.route(c, mk_invoice(c, q2, 100_00, po_number="PO-1")["id"]) == "PENDING_APPROVAL"
        assert route.route(c, mk_invoice(c, None, 100_00)["id"]) == "VENDOR_SETUP"


# --- auto-release ------------------------------------------------------------------------------------

def _recurring_setup(c, **vendor_kw):
    sponsor = mk_user(c, "Sue Sponsor", ["approver"])
    mk_user(c, "Connie Compliance", ["compliance"])  # so an HCP chain can be built when auto-release declines
    vendor_kw.setdefault("recurring_schedule", "monthly")
    v = mk_vendor(c, sponsor["id"], **vendor_kw)
    prior = mk_invoice(c, v, 100_000, state="ENTERED", invoice_date=date(2026, 9, 1), invoice_number="SEP")
    return sponsor, v, prior


def test_auto_release_positive(notify):
    with connection() as c:
        _, v, prior = _recurring_setup(c)
        inv = mk_invoice(c, v, 101_500, invoice_date=date(2026, 10, 1))  # +1.5%
        ok, detail = auto_release.eligible(c, route.load_invoice(c, inv["id"]))
        assert ok and detail["prior_invoice_id"] == str(prior["id"]) and detail["variance_cents"] == "1500"
        assert route.route(c, inv["id"]) == "APPROVED"
        ev = c.execute("select reason_code, payload from event_log where invoice_id = %s and to_state = 'APPROVED'", (inv["id"],)).fetchone()
        assert ev["reason_code"] == "auto_release_recurring" and ev["payload"]["rule"] == "recurring_within_tolerance"
    assert [n["type"] for n in notify] == ["approved_for_ap"]


@pytest.mark.parametrize("case", ["outside_tolerance", "soft_signals", "hcp_vendor", "bank_change_90d", "quarantine_history",
                                  "no_prior_period", "not_recurring", "new_vendor", "released_from_quarantine"])
def test_auto_release_negatives(case):
    with connection() as c:
        kw = {}
        if case == "hcp_vendor":
            kw["hcp"] = True
        if case == "not_recurring":
            kw["recurring_schedule"] = None
        if case == "new_vendor":
            kw["created_at"] = datetime.now(timezone.utc) - timedelta(days=3)
        _, v, prior = _recurring_setup(c, **kw)
        inv_kw = {}
        if case == "outside_tolerance":
            inv_kw["gross"] = 103_000  # +3%
        if case == "soft_signals":
            inv_kw["soft_signals"] = ["urgency_language"]
        if case == "no_prior_period":
            inv_kw["invoice_date"] = date(2026, 9, 15)  # same month as the prior invoice
        if case == "released_from_quarantine":
            inv_kw["tags"] = ["released_from_quarantine"]
        gross = inv_kw.pop("gross", 100_500)
        inv = mk_invoice(c, v, gross, invoice_date=inv_kw.pop("invoice_date", date(2026, 10, 1)), **inv_kw)
        if case == "bank_change_90d":
            c.execute("insert into vendor_bank_fingerprint (vendor_id, fingerprint, last4, source) values (%s, %s, '4321', 'sap')", (v["id"], b"fp"))
        if case == "quarantine_history":
            c.execute("insert into quarantine_case (invoice_id, triggers, state) values (%s, '{direct_sender}', 'released')", (inv["id"],))
        ok, detail = auto_release.eligible(c, route.load_invoice(c, inv["id"]))
        assert not ok and detail["reason"] == case
        assert route.route(c, inv["id"]) == "PENDING_APPROVAL"


# --- escalation / reassign ---------------------------------------------------------------------------

def test_escalate_due_fires_each_step_once():
    with connection() as c:
        mgr = mk_user(c, "Mary Manager", ["supervisor"])
        sponsor = mk_user(c, "Sue Sponsor", ["approver"], manager=mgr["id"])
        ctrl1, ctrl2 = mk_user(c, "Cora Controller", ["controller"]), mk_user(c, "Colin Controller", ["controller"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, 500_00)
        route.route(c, inv["id"])
        monday = datetime(2026, 10, 5, 13, 0, tzinfo=timezone.utc)
        c.execute("update approval_request set created_at = %s where invoice_id = %s", (monday, inv["id"]))
        assert reassign.escalate_due(c, monday + timedelta(days=1)) == []
        wed = reassign.escalate_due(c, monday + timedelta(days=2))
        assert [i["type"] for i in wed] == ["reminder"] and wed[0]["user_id"] == sponsor["id"]
        assert reassign.escalate_due(c, monday + timedelta(days=2, hours=3)) == []  # already stamped
        thu = reassign.escalate_due(c, monday + timedelta(days=3))
        assert [i["type"] for i in thu] == ["manager_copy"] and thu[0]["user_id"] == mgr["id"]
        fri = reassign.escalate_due(c, monday + timedelta(days=4))
        assert sorted(i["type"] for i in fri) == ["controller_reassign", "controller_reassign"]
        assert {i["user_id"] for i in fri} == {ctrl1["id"], ctrl2["id"]}
        assert reassign.escalate_due(c, monday + timedelta(days=10)) == []
        # Weekend does not count: created Friday Oct 2, Monday Oct 5 is only 1 business day, Tuesday is 2.
        # (Not the following Friday: Oct 12 is Columbus Day and the holiday-aware calendar skips it.)
        inv2 = mk_invoice(c, v, 500_00)
        route.route(c, inv2["id"])
        friday = datetime(2026, 10, 2, 13, 0, tzinfo=timezone.utc)
        c.execute("update approval_request set created_at = %s where invoice_id = %s", (friday, inv2["id"]))
        assert reassign.escalate_due(c, friday + timedelta(days=3)) == []
        assert [i["type"] for i in reassign.escalate_due(c, friday + timedelta(days=4))] == ["reminder"]


def test_reassign_rules_and_leaver_sweep():
    with connection() as c:
        sponsor, other, backup = mk_user(c, "Sue Sponsor", ["approver"]), mk_user(c, "Olga Other", ["approver"]), mk_user(c, "Bob Backup", ["approver"])
        clerk = mk_user(c, "Kim Clerk", ["clerk"])
        v = mk_vendor(c, sponsor["id"], backup["id"])
        inv = mk_invoice(c, v, 500_00)
        c.execute("insert into invoice_correction (invoice_id, user_id, diff) values (%s, %s, '{}')", (inv["id"], clerk["id"]))
        route.route(c, inv["id"])
        (req,) = requests_for(c, inv["id"])
        with pytest.raises(reassign.ReassignError):
            reassign.reassign(c, req["id"], sponsor["id"], sponsor["id"], "me")  # never to self
        with pytest.raises(reassign.ReassignError):
            reassign.reassign(c, req["id"], clerk["id"], sponsor["id"], "clerk")  # SOD-1
        new_id = reassign.reassign(c, req["id"], other["id"], sponsor["id"], "holiday")
        rows = requests_for(c, inv["id"])
        assert [r["status"] for r in rows] == ["reassigned", "open"] and rows[1]["id"] == new_id and rows[1]["due_at"] == rows[0]["due_at"]
        # Leaver: Olga goes; her open item lands on the vendor's backup sponsor.
        res = reassign.reassign_open_items_for(c, other["id"])
        assert res[0]["to"] == backup["id"]
        assert c.execute("select assignee_id from approval_request where invoice_id = %s and status = 'open'", (inv["id"],)).fetchone()["assignee_id"] == backup["id"]


# --- payload, content hash, seal ------------------------------------------------------------------------

def test_payload_canonical_form_and_content_hash():
    with connection() as c:
        sponsor = mk_user(c, "Sue Sponsor", ["approver"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, 123_45, bank_fingerprint=b"\xab\xcd")
        full = route.load_invoice(c, inv["id"])
        expected = hashlib.sha256(bytes(inv["file_sha256"]) + canonical(payload_mod.approved_fields(full))).digest()
        assert payload_mod.content_hash(c, inv["id"]) == expected
        assert c.execute("select content_hash from invoice where id = %s", (inv["id"],)).fetchone()["content_hash"] == expected
        pl = payload_mod.build_payload(c, inv["id"], cu(sponsor), "sponsor", "approve", "att-v1", matrix_version_id=None)
        assert set(pl) == {"invoice_id", "content_hash", "vendor_sap_number", "bank_fingerprint_ref", "gross_cents", "currency", "gl",
                           "cost_center", "po_number", "matrix_version_id", "signer_user_id", "signer_role", "attestation_text_id",
                           "decision", "server_time"}
        assert pl["gross_cents"] == "12345" and pl["bank_fingerprint_ref"] == "abcd" and pl["content_hash"] == expected.hex()
        assert canonical(pl) == json.dumps(pl, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()
        assert payload_mod.payload_hash(pl) == hashlib.sha256(canonical(pl)).digest()


def test_dev_seal_and_verify():
    h = hashlib.sha256(b"payload").digest()
    s = signing_service.seal(h, None, None)
    assert s.key_id == "dev-local"
    pem = signing_service.public_key_pem()
    assert signing_service.verify_seal(h, s.signature, pem)
    assert not signing_service.verify_seal(hashlib.sha256(b"other").digest(), s.signature, pem)
    assert not signing_service.verify_seal(h, s.signature[:-1] + bytes([s.signature[-1] ^ 1]), pem)
    with pytest.raises(signing_service.SealUnavailable):
        signing_service.seal(b"short", None, None)


def test_production_seal_fails_closed(monkeypatch):
    object.__setattr__(settings, "dev_seal_private_key_pem", "")  # frozen dataclass: bypass on purpose, restored below
    try:
        with pytest.raises(signing_service.SealUnavailable):
            signing_service.seal(hashlib.sha256(b"x").digest(), None, None)
        object.__setattr__(settings, "signing_service_url", "https://example.invalid/seal")
        object.__setattr__(settings, "aws_access_key_id", "AKIA")
        object.__setattr__(settings, "aws_secret_access_key", "secret")
        monkeypatch.setattr(signing_service.httpx, "post", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("down")))
        with pytest.raises(signing_service.SealUnavailable):
            signing_service.seal(hashlib.sha256(b"x").digest(), None, None)
    finally:
        object.__setattr__(settings, "dev_seal_private_key_pem", os.environ["DEV_SEAL_PRIVATE_KEY_PEM"])
        object.__setattr__(settings, "signing_service_url", "")
        object.__setattr__(settings, "aws_access_key_id", "")
        object.__setattr__(settings, "aws_secret_access_key", "")


def test_sigv4_matches_aws_published_vector():
    # AWS "Signature Version 4 signing process" worked example (GET iam ListUsers, 2015-08-30).
    headers = signing_service.sigv4_sign(
        method="GET", url="https://iam.amazonaws.com/?Action=ListUsers&Version=2010-05-08",
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=utf-8", "Host": "iam.amazonaws.com"}, body=b"",
        access_key="AKIDEXAMPLE", secret_key="wJalrXUtnFEMI/K7MDENG+bPxRfiCYEXAMPLEKEY", region="us-east-1", service="iam",
        amz_date="20150830T123600Z")
    assert headers["authorization"] == (
        "AWS4-HMAC-SHA256 Credential=AKIDEXAMPLE/20150830/us-east-1/iam/aws4_request, "
        "SignedHeaders=content-type;host;x-amz-date, Signature=5d672d79c15b13162d9279b0855cfba6789a8edb4c82c400e06b5924a6f2b5d7")


# --- approve / reject -----------------------------------------------------------------------------------

def test_approve_chain_complete_and_notify(notify):
    with connection() as c:
        sponsor, cfo = mk_user(c, "Sue Sponsor", ["approver"]), mk_user(c, "Carl Cfo", ["cfo"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, 500_00)
        c.execute("insert into config_version (key, value, proposed_by, approved_by, approved_at, effective_from) values "
                  "('step_up_amount_cents', '40000', %s, %s, now(), now() - interval '1 day')", (sponsor["id"], cfo["id"]))
        route.route(c, inv["id"])  # $500 ≥ configured $400 step-up → sponsor + cfo
        reqs = {r["assignee_id"]: r for r in requests_for(c, inv["id"])}
        assert set(reqs) == {sponsor["id"], cfo["id"]}
        # Click without any session → refused.
        with pytest.raises(sign.SignRefused, match="passkey"):
            sign.approve(c, cu(sponsor), reqs[sponsor["id"]]["id"], None)
        # Even inside a session, at/above step-up a click is refused.
        open_session(c, sponsor)
        with pytest.raises(sign.SignRefused, match="passkey"):
            sign.approve(c, cu(sponsor), reqs[sponsor["id"]]["id"], None)
        # Lower the bar for the rest of the test: step-up back to the default.
        c.execute("update config_version set value = %s where key = 'step_up_amount_cents'", (str(STEP_UP),))
        sid = sign.approve(c, cu(sponsor), reqs[sponsor["id"]]["id"], None)
        srow = c.execute("select * from approval_signature where id = %s", (sid,)).fetchone()
        assert srow["decision"] == "approve" and srow["assertion"] is None and srow["signing_session_id"] is not None
        assert srow["seal_key_id"] == "dev-local" and signing_service.verify_seal(bytes(srow["payload_sha256"]), bytes(srow["seal"]), signing_service.public_key_pem())
        assert state_of(c, inv["id"]) == "PENDING_APPROVAL"
        # Wrong person.
        with pytest.raises(sign.SignRefused, match="not assigned"):
            sign.approve(c, cu(sponsor), reqs[cfo["id"]]["id"], None)
        open_session(c, cfo)
        sign.approve(c, cu(cfo), reqs[cfo["id"]]["id"], None)
        assert state_of(c, inv["id"]) == "APPROVED"
        ev = c.execute("select payload from event_log where invoice_id = %s and to_state = 'APPROVED'", (inv["id"],)).fetchone()
        assert len(ev["payload"]["signature_ids"]) == 2
        assert c.execute("select count(*) as n from event_log where event_type = 'signature' and invoice_id = %s", (inv["id"],)).fetchone()["n"] == 2
    assert [n["type"] for n in notify if n["type"] == "approved_for_ap"] == ["approved_for_ap"]


def test_reject_voids_others_and_requires_reason():
    with connection() as c:
        sponsor, cfo = mk_user(c, "Sue Sponsor", ["approver"]), mk_user(c, "Carl Cfo", ["cfo"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, STEP_UP)
        route.route(c, inv["id"])
        reqs = {r["assignee_id"]: r for r in requests_for(c, inv["id"])}
        # Make the sponsor's slot a click (under step-up for the sponsor would need a lower amount; use HCP-free click by lowering step-up)
        c.execute("insert into config_version (key, value, proposed_by, approved_by, approved_at, effective_from) values "
                  "('step_up_amount_cents', %s, %s, %s, now(), now() - interval '1 day')", (str(STEP_UP * 10), sponsor["id"], cfo["id"]))
        open_session(c, sponsor)
        open_session(c, cfo)
        sign.approve(c, cu(sponsor), reqs[sponsor["id"]]["id"], None)
        with pytest.raises(sign.SignRefused, match="reason"):
            sign.reject(c, cu(cfo), reqs[cfo["id"]]["id"], "  ", None)
        sign.reject(c, cu(cfo), reqs[cfo["id"]]["id"], "Duplicate of last month", None)
        assert state_of(c, inv["id"]) == "REJECTED"
        sigs = c.execute("select decision, status, supersedes from approval_signature where invoice_id = %s order by signed_at", (inv["id"],)).fetchall()
        assert [(s["decision"], s["status"]) for s in sigs] == [("approve", "valid"), ("approve", "voided"), ("reject", "valid")]
        assert sigs[1]["supersedes"] is not None


def test_content_change_voids_and_refuses(background_ok):
    with connection() as c:
        sponsor, cfo = mk_user(c, "Sue Sponsor", ["approver"]), mk_user(c, "Carl Cfo", ["cfo"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, 500_00)
        route.route(c, inv["id"])
        (req,) = requests_for(c, inv["id"])
        assert req["assignee_id"] == sponsor["id"]
        # Change the amount underneath the routed invoice (as an AP edit outside the correction path would).
        open_session(c, sponsor)
        c.execute("update invoice set gross_cents = gross_cents + 1 where id = %s", (inv["id"],))
        with pytest.raises(sign.SignRefused, match="changed after"):
            sign.approve(c, cu(sponsor), req["id"], None)
        assert state_of(c, inv["id"]) == "AP_REVIEW"
        assert c.execute("select status from approval_request where id = %s", (req["id"],)).fetchone()["status"] == "voided"
        assert c.execute("select count(*) as n from approval_signature where invoice_id = %s", (inv["id"],)).fetchone()["n"] == 0


def test_seal_unavailable_fails_closed(monkeypatch):
    with connection() as c:
        sponsor = mk_user(c, "Sue Sponsor", ["approver"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, 500_00)
        route.route(c, inv["id"])
        (req,) = requests_for(c, inv["id"])
        open_session(c, sponsor)
        monkeypatch.setattr(signing_service, "seal", lambda *a, **k: (_ for _ in ()).throw(signing_service.SealUnavailable("down")))
        with pytest.raises(sign.SignRefused, match="temporarily unavailable"):
            sign.approve(c, cu(sponsor), req["id"], None)
        assert c.execute("select count(*) as n from approval_signature").fetchone()["n"] == 0
        assert c.execute("select status from approval_request where id = %s", (req["id"],)).fetchone()["status"] == "open"


def test_background_check_failure_refuses(monkeypatch):
    from accounting.identity import sessions
    monkeypatch.setattr(sessions, "background_check", lambda user: False)
    with connection() as c:
        sponsor = mk_user(c, "Sue Sponsor", ["approver"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, 500_00)
        route.route(c, inv["id"])
        (req,) = requests_for(c, inv["id"])
        with pytest.raises(sign.SignRefused, match="sign in again"):
            sign.approve(c, cu(sponsor), req["id"], None)


# --- signing sessions -------------------------------------------------------------------------------------

def test_needs_tap_matrix():
    with connection() as c:
        sponsor = mk_user(c, "Sue Sponsor", ["approver"])
        v = mk_vendor(c, sponsor["id"])
        hcp = mk_vendor(c, sponsor["id"], legal_name="Dr Smith", sap="9", hcp=True)
        small = route.load_invoice(c, mk_invoice(c, v, 500_00)["id"])
        big = route.load_invoice(c, mk_invoice(c, v, STEP_UP)["id"])
        hcp_inv = route.load_invoice(c, mk_invoice(c, hcp, 500_00)["id"])
        u = cu(sponsor)
        assert signing_sessions.needs_tap(c, small, u)  # no session
        s = open_session(c, sponsor)
        assert not signing_sessions.needs_tap(c, small, u)
        assert signing_sessions.needs_tap(c, big, u)
        assert signing_sessions.needs_tap(c, hcp_inv, u)
        c.execute("update signing_session set last_activity_at = now() - interval '16 minutes' where id = %s", (s["id"],))
        assert signing_sessions.needs_tap(c, small, u)  # idle → ended
        assert signing_sessions.active(c, u) is None
        s = open_session(c, sponsor)
        c.execute("update signing_session set started_at = now() - interval '61 minutes', expires_at = now() - interval '1 minute' where id = %s", (s["id"],))
        assert signing_sessions.needs_tap(c, small, u)  # max age
        s = open_session(c, sponsor)
        other_device = CurrentUser(**{**u.__dict__, "device_hash": "someone-else"})
        assert signing_sessions.active(c, other_device) is None  # not transferable between browsers


# --- WebAuthn ----------------------------------------------------------------------------------------------

def test_webauthn_enroll_session_and_signed_approval(notify):
    with connection() as c:
        ctrl = mk_user(c, "Cora Controller", ["controller"])
        sponsor = mk_user(c, "Sue Sponsor", ["approver"])
        u = cu(sponsor)
        auth = SoftAuthenticator()
        with pytest.raises(webauthn_flow.WebAuthnError, match="enrollment code"):
            webauthn_flow.registration_options(c, u, "WRONG")
        code = issue_code(c, sponsor, ctrl)
        opts = webauthn_flow.registration_options(c, u, code)
        assert opts["rp"]["id"] == settings.rp_id and opts["authenticatorSelection"]["userVerification"] == "required"
        pk = webauthn_flow.verify_registration(c, u, auth.register(opts), code, "Sue's phone")
        assert bytes(pk["credential_id"]) == auth.credential_id and pk["label"] == "Sue's phone" and pk["transports"] == ["internal"]
        assert c.execute("select used_at from enrollment_code").fetchone()["used_at"] is not None
        with pytest.raises(webauthn_flow.WebAuthnError):  # code is single-use
            webauthn_flow.registration_options(c, u, code)
        alerts = [n for n in notify if n["type"] == "enroll_alert"]
        assert {n["user_id"] for n in alerts} == {sponsor["id"], ctrl["id"]}

        # Session tap.
        sopts = webauthn_flow.authentication_options(c, u, "session")
        assert sopts["userVerification"] == "required" and sopts["allowCredentials"][0]["id"] == bytes_to_base64url(auth.credential_id)
        pk2, _ch = webauthn_flow.verify_authentication(c, u, auth.authenticate(sopts), "session")
        assert pk2["id"] == pk["id"]
        assert c.execute("select sign_count, last_used_at from passkey where id = %s", (pk["id"],)).fetchone()["sign_count"] == auth.sign_count
        sess = signing_sessions.start(c, u, pk["id"], sponsor["session_id"], "dev-hash")

        # A replayed assertion (same challenge) is refused: the challenge was consumed.
        with pytest.raises(webauthn_flow.WebAuthnError, match="expired"):
            webauthn_flow.verify_authentication(c, u, auth.authenticate(sopts), "session")

        # Signed approval on an HCP invoice: challenge == payload hash, payload stored with it.
        comp = mk_user(c, "Connie Compliance", ["compliance"])
        v = mk_vendor(c, sponsor["id"], hcp=True)
        inv = mk_invoice(c, v, 500_00)
        route.route(c, inv["id"])
        req = next(r for r in requests_for(c, inv["id"]) if r["assignee_id"] == sponsor["id"])
        with pytest.raises(sign.SignRefused, match="passkey tap"):
            sign.approve(c, u, req["id"], None)  # HCP → click refused even inside a session
        o = sign.signing_options(c, u, req["id"], "approve")
        ch = c.execute("select * from webauthn_challenge where approval_request_id = %s and used_at is null", (req["id"],)).fetchone()
        assert bytes_to_base64url(bytes(ch["challenge"])) == o["challenge"]
        assert bytes(ch["challenge"]) == payload_mod.payload_hash(ch["payload"]) and ch["payload"]["decision"] == "approve"
        # A tap over the wrong challenge is refused and logged.
        with pytest.raises(webauthn_flow.WebAuthnError):
            sign.approve(c, u, req["id"], auth.authenticate(o, challenge_override=bytes_to_base64url(b"\x00" * 32)))
        o = sign.signing_options(c, u, req["id"], "approve")  # new options (old challenge consumed)
        sid = sign.approve(c, u, req["id"], auth.authenticate(o))
        srow = c.execute("select * from approval_signature where id = %s", (sid,)).fetchone()
        assert srow["passkey_id"] == pk["id"] and srow["assertion"]["rawId"] == bytes_to_base64url(auth.credential_id)
        assert srow["signing_session_id"] == sess["id"]  # the tap rode on the open sitting and is recorded against it
        assert bytes(srow["payload_sha256"]) == payload_mod.payload_hash(srow["payload"])
        assert srow["payload"]["decision"] == "approve" and srow["payload"]["signer_user_id"] == str(sponsor["id"])
        assert signing_service.verify_seal(bytes(srow["payload_sha256"]), bytes(srow["seal"]), signing_service.public_key_pem())
        assert c.execute("select status from approval_request where id = %s", (req["id"],)).fetchone()["status"] == "signed"
        assert state_of(c, inv["id"]) == "PENDING_APPROVAL"  # Compliance still has to co-sign
        req2 = next(r for r in requests_for(c, inv["id"]) if r["assignee_id"] == comp["id"])
        assert req2["status"] == "open"
        # Sue cannot act on Connie's slot, tap or no tap.
        with pytest.raises(sign.SignRefused, match="not assigned"):
            sign.reject(c, u, req2["id"], "no", None)
        # Options sealed for 'approve' cannot be turned in as a rejection: the stored payload says 'approve'.
        comp_u, comp_auth = cu(comp), SoftAuthenticator()
        ccode = issue_code(c, comp, ctrl, "WXYZ-9876")
        webauthn_flow.verify_registration(c, comp_u, comp_auth.register(webauthn_flow.registration_options(c, comp_u, ccode)), ccode)
        o2 = sign.signing_options(c, comp_u, req2["id"], "approve")
        with pytest.raises(sign.SignRefused, match="does not match"):
            sign.reject(c, comp_u, req2["id"], "changed my mind", comp_auth.authenticate(o2))



def test_webauthn_expired_challenge_and_revoked_key():
    with connection() as c:
        ctrl, sponsor = mk_user(c, "Cora Controller", ["controller"]), mk_user(c, "Sue Sponsor", ["approver"])
        u, auth = cu(sponsor), SoftAuthenticator()
        code = issue_code(c, sponsor, ctrl)
        opts = webauthn_flow.registration_options(c, u, code)
        c.execute("update webauthn_challenge set created_at = now() - interval '3 minutes'")
        with pytest.raises(webauthn_flow.WebAuthnError, match="expired"):
            webauthn_flow.verify_registration(c, u, auth.register(opts), code)
        opts = webauthn_flow.registration_options(c, u, code)
        pk = webauthn_flow.verify_registration(c, u, auth.register(opts), code)
        assert webauthn_flow.revoke_passkey(c, u, pk["id"])
        with pytest.raises(webauthn_flow.WebAuthnError, match="no passkey"):
            webauthn_flow.authentication_options(c, u, "session")


# --- quarantine ---------------------------------------------------------------------------------------------

def _quarantined(c, triggers, gross=500_00, with_message=True):
    ctrl1, ctrl2 = mk_user(c, "Cora Controller", ["controller"]), mk_user(c, "Colin Controller", ["controller"])
    sponsor = mk_user(c, "Sue Sponsor", ["approver"])
    v = mk_vendor(c, sponsor["id"], phone="215-555-0100", phone_since=date(2025, 1, 1))
    inv = mk_invoice(c, v, gross, state="QUARANTINED")
    mid = None
    if with_message:
        mid = c.execute(
            """insert into message (raw_ref, raw_sha256, from_addr, original_sender, subject, received_at, headers, trust)
               values ('r', %s, 'billing@acme-supplies.co', null, 'Invoice', now(), '{}', 'direct') returning id""",
            (secrets.token_bytes(32),)).fetchone()["id"]
    case = c.execute("insert into quarantine_case (invoice_id, message_id, triggers) values (%s, %s, %s) returning *",
                     (inv["id"], mid, triggers)).fetchone()
    return ctrl1, ctrl2, sponsor, v, inv, case


def test_quarantine_release_sender_case_needs_callback_and_second_person(notify):
    with connection() as c:
        ctrl1, ctrl2, sponsor, v, inv, case = _quarantined(c, ["lookalike_domain"])
        with pytest.raises(quarantine.QuarantineError, match="callback"):
            quarantine.release(c, case["id"], ctrl1["id"], ctrl2["id"], None, None, None)
        with pytest.raises(quarantine.QuarantineError, match="phone number"):
            quarantine.release(c, case["id"], ctrl1["id"], ctrl2["id"], "contract", "", None)
        with pytest.raises(quarantine.QuarantineError, match="second person"):
            quarantine.release(c, case["id"], ctrl1["id"], None, "contract", "215-555-0199", "spoke to Jan")
        with pytest.raises(quarantine.QuarantineError, match="not yourself"):
            quarantine.release(c, case["id"], ctrl1["id"], ctrl1["id"], "contract", "215-555-0199", "spoke to Jan")
        with pytest.raises(quarantine.QuarantineError, match="Controller or CFO"):
            quarantine.release(c, case["id"], ctrl1["id"], sponsor["id"], "contract", "215-555-0199", "spoke to Jan")
        assert quarantine.release(c, case["id"], ctrl1["id"], ctrl2["id"], "contract", "215-555-0199", "spoke to Jan") == "pending_second_person"
        row = c.execute("select * from quarantine_case where id = %s", (case["id"],)).fetchone()
        assert row["state"] == "open" and row["release_pending_at"] is not None and row["second_reviewer_id"] == ctrl2["id"]
        assert state_of(c, inv["id"]) == "QUARANTINED"
        # The first reviewer calling again does not complete it; the second person does.
        assert quarantine.release(c, case["id"], ctrl1["id"], ctrl2["id"], "contract", "215-555-0199", "again") == "pending_second_person"
        assert quarantine.release(c, case["id"], ctrl2["id"], None, None, None, None) == "released"
        row = c.execute("select * from quarantine_case where id = %s", (case["id"],)).fetchone()
        assert row["state"] == "released" and row["callback_source"] == "contract"
        invrow = c.execute("select state, tags from invoice where id = %s", (inv["id"],)).fetchone()
        assert "released_from_quarantine" in invrow["tags"]
        assert invrow["state"] == "PENDING_APPROVAL"  # QUA-3: released → captured → full chain, never auto-release
        ev = c.execute("select actor_type, reason_code from event_log where invoice_id = %s and to_state = 'CAPTURED'", (inv["id"],)).fetchone()
        assert (ev["actor_type"], ev["reason_code"]) == ("second_person", "controller_release")
        with pytest.raises(quarantine.QuarantineError, match="already"):
            quarantine.reject(c, case["id"], ctrl1["id"], "late")


def test_quarantine_release_amount_rule_and_vendor_master_phone():
    with connection() as c:
        ctrl1, ctrl2, sponsor, v, inv, case = _quarantined(c, ["risky_file"], gross=500_00)
        assert not quarantine.needs_callback(case)
        assert quarantine.release(c, case["id"], ctrl1["id"], None, None, None, "clean after review") == "released"
        # Above $10,000: second person even without a sender/bank trigger.
        ctrl1, ctrl2, sponsor, v, inv, case = _quarantined(c, ["hidden_text"], gross=1_000_001)
        with pytest.raises(quarantine.QuarantineError, match="second person"):
            quarantine.release(c, case["id"], ctrl1["id"], None, None, None, None)
        assert quarantine.release(c, case["id"], ctrl1["id"], ctrl2["id"], None, None, None) == "pending_second_person"
        assert quarantine.release(c, case["id"], ctrl2["id"], None, None, None, None) == "released"
        # Bank case with the vendor-master number: refused when the number is younger than 90 days.
        ctrl1, ctrl2, sponsor, v, inv, case = _quarantined(c, ["bank_mismatch"])
        c.execute("update vendor set phone_since = current_date - 10 where id = %s", (v["id"],))
        with pytest.raises(quarantine.QuarantineError, match="90 days"):
            quarantine.release(c, case["id"], ctrl1["id"], ctrl2["id"], "vendor_master_90d", "215-555-0100", "ok")
        c.execute("update vendor set phone_since = current_date - 120 where id = %s", (v["id"],))
        assert quarantine.release(c, case["id"], ctrl1["id"], ctrl2["id"], "vendor_master_90d", "215-555-0100", "ok") == "pending_second_person"


def test_quarantine_reject_report_and_escalation(notify):
    with connection() as c:
        ctrl1, ctrl2, sponsor, v, inv, case = _quarantined(c, ["lookalike_domain"])
        quarantine.reject(c, case["id"], ctrl1["id"], "Not our vendor")
        assert state_of(c, inv["id"]) == "FRAUD"
        ctrl1, ctrl2, sponsor, v, inv2, case2 = _quarantined(c, ["lookalike_domain"])
        out = quarantine.report(c, case2["id"], ctrl1["id"])
        assert out["blocked_domain"] == "acme-supplies.co"
        assert c.execute("select pattern from blocked_sender").fetchone()["pattern"] == "acme-supplies.co"
        assert state_of(c, inv2["id"]) == "FRAUD"
        assert [n for n in notify if n["type"] == "security_report"][0]["context"]["blocked_domain"] == "acme-supplies.co"
        # Escalation after two business days, once.
        cfo = mk_user(c, "Carl Cfo", ["cfo"])
        ctrl1, ctrl2, sponsor, v, inv3, case3 = _quarantined(c, ["direct_sender"])
        monday = datetime(2026, 10, 5, 13, 0, tzinfo=timezone.utc)
        c.execute("update quarantine_case set opened_at = %s where id = %s", (monday, case3["id"]))
        assert quarantine.escalate_stale(c, monday + timedelta(days=1)) == []
        intents = quarantine.escalate_stale(c, monday + timedelta(days=2))
        assert [(i["type"], i["user_id"]) for i in intents] == [("quarantine_escalated", cfo["id"])]
        assert quarantine.escalate_stale(c, monday + timedelta(days=5)) == []
        assert quarantine.is_escalated(c.execute("select * from quarantine_case where id = %s", (case3["id"],)).fetchone(), monday + timedelta(days=3))
        for code in ("lookalike_domain", "bank_mismatch", "hidden_text", "something_new"):
            assert len(quarantine.trigger_text(code)) > 20


def test_void_all_creates_new_rows_and_closes_requests():
    with connection() as c:
        sponsor = mk_user(c, "Sue Sponsor", ["approver"])
        v = mk_vendor(c, sponsor["id"])
        inv = mk_invoice(c, v, 500_00)
        route.route(c, inv["id"])
        (req,) = requests_for(c, inv["id"])
        c.execute("update approval_request set status = 'signed' where id = %s", (req["id"],))
        c.execute("insert into approval_request (invoice_id, assignee_id, role, due_at) values (%s, %s, 'cfo', now())", (inv["id"], sponsor["id"]))
        h = payload_mod.content_hash(c, inv["id"])
        c.execute("""insert into approval_signature (invoice_id, approval_request_id, signer_id, signer_role, decision, content_hash, payload,
                     payload_sha256, seal, seal_key_id, attestation_text_id) values (%s, %s, %s, 'sponsor', 'approve', %s, '{}', %s, %s, 'k', 'att-v1')""",
                  (inv["id"], req["id"], sponsor["id"], h, h, b"seal"))
        assert void.void_all_for_invoice(c, inv["id"], "field_edit", sponsor["id"]) == 1
        rows = c.execute("select status, supersedes from approval_signature where invoice_id = %s order by signed_at", (inv["id"],)).fetchall()
        assert [r["status"] for r in rows] == ["valid", "voided"] and rows[1]["supersedes"] is not None
        assert c.execute("select count(*) as n from approval_request where invoice_id = %s and status = 'open'", (inv["id"],)).fetchone()["n"] == 0
        assert void.void_all_for_invoice(c, inv["id"], "again", None) == 0  # idempotent: nothing left to void
