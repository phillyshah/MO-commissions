"""/vendors — list, detail, new/edit under dual control, CSV import, domains, platform map, and the
VENDOR_SETUP queue where AP attaches a parked invoice to a vendor (VEN-4).

Reads are open to every AP-side role plus the auditor; writes need clerk/supervisor/admin AND an
active (not frozen) account. Form errors are our own `VendorError` text; anything else is logged
and shown as a generic line — a raw exception never reaches a page.
"""
from __future__ import annotations

import logging
from uuid import UUID

from fastapi import APIRouter, Depends, Request
from fastapi.responses import RedirectResponse

from accounting.config_flags import config_flag
from accounting.db.pool import connection
from accounting.identity.sessions import CurrentUser
from accounting.lifecycle import InvalidTransition, log_event, transition
from accounting.vendors import fingerprint as fp
from accounting.vendors import service
from accounting.web.deps import NotFound, require, require_active
from accounting.web.templating import render

log = logging.getLogger("accounting.vendors.router")

router = APIRouter(prefix="/vendors")
READ = Depends(require("clerk", "supervisor", "controller", "admin", "auditor"))
WRITE = Depends(require("clerk", "supervisor", "admin"))
ACTIVE = Depends(require_active)

GENERIC_ERROR = "Something went wrong saving that. Nothing was changed — try again."


def _form_fields(form) -> dict:
    """Turn the vendor form into the dict `service.propose_change` takes. Unchecked boxes are absent
    from a POST, so every boolean is read explicitly."""
    g = lambda k: (form.get(k) or "").strip()  # noqa: E731
    remit = {k: g(f"remit_{k}") or None for k in service.REMIT_KEYS}
    fields = {k: g(k) or None for k in ("sap_vendor_number", "legal_name", "country", "default_gl", "default_cost_center",
                                        "payment_terms", "asl_status", "recurring_schedule", "phone", "status")}
    fields.update({
        "default_sponsor_id": g("default_sponsor_id") or None,
        "backup_sponsor_id": g("backup_sponsor_id") or None,
        "recurring_tolerance_pct": g("recurring_tolerance_pct") or None,
        "remit_to": remit if any(remit.values()) else None,
        "tin": g("tin") or None,
        "bank_routing": g("bank_routing"), "bank_account": g("bank_account"),
    })
    for b in service.BOOL_FIELDS:
        fields[b] = form.get(b) in ("on", "true", "1")
    if fields["status"] not in ("active", "blocked", "setup"):
        fields["status"] = None
    return fields


def _people(conn) -> list[dict]:
    return conn.execute("select id, display_name, email from ap_user where status <> 'disabled' order by display_name").fetchall()


def _platforms(conn) -> list[dict]:
    return conn.execute("select id, display_name from billing_platform order by display_name").fetchall()


@router.get("")
def list_vendors(request: Request, q: str = "", status: str = "", user: CurrentUser = READ):
    with connection() as conn:
        where, params = ["1=1"], []
        if q.strip():
            where.append("(legal_name ilike %s or sap_vendor_number ilike %s)")
            params += [f"%{q.strip()}%", f"%{q.strip()}%"]
        if status in ("active", "blocked", "setup"):
            where.append("status = %s")
            params.append(status)
        vendors = conn.execute(
            f"""select v.*, (select count(*) from invoice i where i.vendor_id = v.id) as invoice_count,
                       (select last4 from vendor_bank_fingerprint f where f.vendor_id = v.id and superseded_at is null limit 1) as last4
                from vendor v where {' and '.join(where)} order by legal_name limit 500""", params).fetchall()
        pending = conn.execute("select count(*) as n from vendor_change where applied_at is null and rejected_at is null").fetchone()["n"]
        setup_n = conn.execute("select count(*) as n from invoice where state = 'VENDOR_SETUP'").fetchone()["n"]
    return render(request, "vendors/list.html", {"vendors": vendors, "q": q, "status": status, "pending": pending, "setup_n": setup_n})


@router.get("/new")
def new_form(request: Request, invoice: UUID | None = None, user: CurrentUser = WRITE):
    prefill: dict = {}
    with connection() as conn:
        if invoice:
            inv = conn.execute("select * from invoice where id = %s", (invoice,)).fetchone()
            if inv:
                prefill = {"legal_name": inv["vendor_name_raw"], "remit_to": inv["remit_to"], "invoice_id": str(inv["id"]),
                           "country": (inv["remit_to"] or {}).get("country")}
        ctx = {"vendor": prefill, "people": _people(conn), "dual": config_flag(conn, "dual_control_enabled"), "mode": "new"}
    return render(request, "vendors/form.html", ctx)


@router.post("/new")
async def create(request: Request, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    form = await request.form()
    fields = _form_fields(form)
    invoice_id = (form.get("invoice_id") or "").strip() or None
    try:
        with connection() as conn:
            change, applied = service.propose_change(conn, None, fields, user.id)
            if change is None:
                raise service.VendorError("Nothing to save.")
            if applied:
                for d in (form.get("allowed_domains") or "").split(";"):
                    if d.strip():
                        service.add_domain(conn, change["vendor_id"], d, user.id)
                if invoice_id:
                    _attach(conn, invoice_id, change["vendor_id"], user)
                return RedirectResponse(f"/vendors/{change['vendor_id']}", status_code=303)
            return RedirectResponse("/vendors?proposed=1", status_code=303)
    except service.VendorError as e:
        message = str(e)
    except Exception:
        log.exception("vendor create failed")
        message = GENERIC_ERROR
    with connection() as conn:
        ctx = {"vendor": {**fields, "invoice_id": invoice_id}, "people": _people(conn), "error": message, "mode": "new",
               "dual": config_flag(conn, "dual_control_enabled")}
    return render(request, "vendors/form.html", ctx, status_code=400)


@router.get("/import")
def import_form(request: Request, user: CurrentUser = WRITE):
    return render(request, "vendors/import.html", {"columns": service.CSV_COLUMNS})


@router.post("/import")
async def import_post(request: Request, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    form = await request.form()
    upload = form.get("file")
    report, message = None, None
    try:
        if upload is None or not hasattr(upload, "read"):
            raise service.VendorError("Choose a CSV file first.")
        data = await upload.read()
        if len(data) > 5 * 1024 * 1024:
            raise service.VendorError("The file is larger than 5 MB.")
        text = data.decode("utf-8-sig", errors="replace")
        with connection() as conn:
            report = service.import_csv(conn, text, user.id)
    except service.VendorError as e:
        message = str(e)
    except Exception:
        log.exception("vendor import failed")
        message = GENERIC_ERROR
    return render(request, "vendors/import.html", {"columns": service.CSV_COLUMNS, "report": report, "error": message},
                  status_code=400 if message else 200)


@router.get("/setup")
def setup_queue(request: Request, user: CurrentUser = READ):
    with connection() as conn:
        rows = conn.execute(
            """select i.*, m.from_addr, m.from_name from invoice i left join message m on m.id = i.message_id
               where i.state = 'VENDOR_SETUP' order by i.received_at""").fetchall()
        vendors = conn.execute("select id, legal_name, sap_vendor_number from vendor where status <> 'blocked' order by legal_name").fetchall()
    return render(request, "vendors/setup.html", {"rows": rows, "vendors": vendors})


def _attach(conn, invoice_id, vendor_id, user: CurrentUser) -> str:
    inv = conn.execute("select * from invoice where id = %s for update", (invoice_id,)).fetchone()
    if inv is None:
        raise service.VendorError("That invoice no longer exists.")
    conn.execute("update invoice set vendor_id = %s, vendor_match = %s::jsonb, updated_at = now() where id = %s",
                 (vendor_id, '{"method": "ap_attached", "confidence": 1.0}', invoice_id))
    log_event(conn, event_type="vendor_attached", actor_type="user", actor_id=user.id, invoice_id=invoice_id,
              payload={"vendor_id": str(vendor_id)})
    if inv["state"] != "VENDOR_SETUP":
        return inv["state"]
    transition(conn, invoice_id, "CAPTURED", reason_code="vendor_added", actor_type="user", actor_id=user.id,
               payload={"vendor_id": str(vendor_id)})
    from accounting.routing.route import route  # another package's module; imported where it is needed

    return route(conn, invoice_id)


@router.post("/setup/{invoice_id}/attach")
async def attach(request: Request, invoice_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    form = await request.form()
    vendor_id = (form.get("vendor_id") or "").strip()
    try:
        with connection() as conn:
            if not vendor_id or service.get(conn, vendor_id) is None:
                raise service.VendorError("Choose a vendor.")
            _attach(conn, invoice_id, vendor_id, user)
        return RedirectResponse("/vendors/setup", status_code=303)
    except (service.VendorError, InvalidTransition) as e:
        message = str(e) if isinstance(e, service.VendorError) else "That invoice is no longer waiting for a vendor."
    except Exception:
        log.exception("attach failed")
        message = GENERIC_ERROR
    with connection() as conn:
        rows = conn.execute(
            """select i.*, m.from_addr, m.from_name from invoice i left join message m on m.id = i.message_id
               where i.state = 'VENDOR_SETUP' order by i.received_at""").fetchall()
        vendors = conn.execute("select id, legal_name, sap_vendor_number from vendor where status <> 'blocked' order by legal_name").fetchall()
    return render(request, "vendors/setup.html", {"rows": rows, "vendors": vendors, "error": message}, status_code=400)


@router.post("/changes/{change_id}/approve")
def approve_change(request: Request, change_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    try:
        with connection() as conn:
            row = service.approve_change(conn, change_id, user.id)
        return RedirectResponse(f"/vendors/{row['vendor_id']}", status_code=303)
    except service.VendorError as e:
        return render(request, "error.html", {"message": str(e)}, status_code=400)


@router.post("/changes/{change_id}/reject")
async def reject_change(request: Request, change_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    form = await request.form()
    try:
        with connection() as conn:
            service.reject_change(conn, change_id, user.id, (form.get("reason") or "").strip()[:500] or "rejected")
        return RedirectResponse("/vendors", status_code=303)
    except service.VendorError as e:
        return render(request, "error.html", {"message": str(e)}, status_code=400)


@router.get("/{vendor_id}")
def detail(request: Request, vendor_id: UUID, user: CurrentUser = READ):
    with connection() as conn:
        v = service.get(conn, vendor_id)
        if v is None:
            raise NotFound()
        ctx = {
            "vendor": v,
            "fingerprint": fp.current(conn, vendor_id),
            "domains": conn.execute(
                """select d.*, u.display_name as added_by_name from vendor_domain d join ap_user u on u.id = d.added_by
                   where d.vendor_id = %s order by d.status, d.pattern""", (vendor_id,)).fetchall(),
            "platforms": conn.execute(
                """select m.*, p.display_name from billing_platform_map m join billing_platform p on p.id = m.platform_id
                   where m.vendor_id = %s order by p.display_name""", (vendor_id,)).fetchall(),
            "all_platforms": _platforms(conn),
            "changes": conn.execute(
                """select c.*, p.display_name as proposed_by_name, a.display_name as approved_by_name
                   from vendor_change c join ap_user p on p.id = c.proposed_by left join ap_user a on a.id = c.approved_by
                   where c.vendor_id = %s order by c.proposed_at desc limit 50""", (vendor_id,)).fetchall(),
            "invoices": conn.execute(
                """select id, portal_number, state, invoice_number, invoice_date, gross_cents, currency, received_at
                   from invoice where vendor_id = %s order by received_at desc limit 25""", (vendor_id,)).fetchall(),
            "sponsor": conn.execute("select display_name, email from ap_user where id = %s", (v["default_sponsor_id"],)).fetchone() if v["default_sponsor_id"] else None,
            "backup": conn.execute("select display_name, email from ap_user where id = %s", (v["backup_sponsor_id"],)).fetchone() if v["backup_sponsor_id"] else None,
            "dual": config_flag(conn, "dual_control_enabled"),
            "can_write": user.has("clerk", "supervisor", "admin") and user.status == "active",
        }
    return render(request, "vendors/detail.html", ctx)


@router.get("/{vendor_id}/edit")
def edit_form(request: Request, vendor_id: UUID, user: CurrentUser = WRITE):
    with connection() as conn:
        v = service.get(conn, vendor_id)
        if v is None:
            raise NotFound()
        ctx = {"vendor": v, "people": _people(conn), "dual": config_flag(conn, "dual_control_enabled"), "mode": "edit",
               "fingerprint": fp.current(conn, vendor_id)}
    return render(request, "vendors/form.html", ctx)


@router.post("/{vendor_id}/edit")
async def edit(request: Request, vendor_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    form = await request.form()
    fields = _form_fields(form)
    try:
        with connection() as conn:
            change, applied = service.propose_change(conn, vendor_id, fields, user.id)
        suffix = "" if (change is None or applied) else "?proposed=1"
        return RedirectResponse(f"/vendors/{vendor_id}{suffix}", status_code=303)
    except service.VendorError as e:
        message = str(e)
    except Exception:
        log.exception("vendor edit failed")
        message = GENERIC_ERROR
    with connection() as conn:
        v = service.get(conn, vendor_id) or {}
        ctx = {"vendor": {**v, **{k: val for k, val in fields.items() if k not in ("tin", "bank_routing", "bank_account")}},
               "people": _people(conn), "error": message, "mode": "edit", "dual": config_flag(conn, "dual_control_enabled"),
               "fingerprint": fp.current(conn, vendor_id)}
    return render(request, "vendors/form.html", ctx, status_code=400)


@router.post("/{vendor_id}/domains")
async def add_domain(request: Request, vendor_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    form = await request.form()
    try:
        with connection() as conn:
            service.add_domain(conn, vendor_id, form.get("pattern") or "", user.id)
    except service.VendorError as e:
        return render(request, "error.html", {"message": str(e)}, status_code=400)
    return RedirectResponse(f"/vendors/{vendor_id}", status_code=303)


@router.post("/{vendor_id}/domains/{domain_id}/approve")
def approve_domain(request: Request, vendor_id: UUID, domain_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    try:
        with connection() as conn:
            service.approve_domain(conn, domain_id, user.id)
    except service.VendorError as e:
        return render(request, "error.html", {"message": str(e)}, status_code=400)
    return RedirectResponse(f"/vendors/{vendor_id}", status_code=303)


@router.post("/{vendor_id}/domains/{domain_id}/revoke")
def revoke_domain(request: Request, vendor_id: UUID, domain_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    with connection() as conn:
        service.revoke_domain(conn, domain_id, user.id)
    return RedirectResponse(f"/vendors/{vendor_id}", status_code=303)


@router.post("/{vendor_id}/platforms")
async def add_platform(request: Request, vendor_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    form = await request.form()
    try:
        with connection() as conn:
            service.add_platform_map(conn, vendor_id, form.get("platform_id") or "", form.get("account_id") or "", user.id)
    except service.VendorError as e:
        return render(request, "error.html", {"message": str(e)}, status_code=400)
    return RedirectResponse(f"/vendors/{vendor_id}", status_code=303)


@router.post("/{vendor_id}/platforms/{map_id}/approve")
def approve_platform(request: Request, vendor_id: UUID, map_id: UUID, user: CurrentUser = WRITE, _: CurrentUser = ACTIVE):
    try:
        with connection() as conn:
            service.approve_platform_map(conn, map_id, user.id)
    except service.VendorError as e:
        return render(request, "error.html", {"message": str(e)}, status_code=400)
    return RedirectResponse(f"/vendors/{vendor_id}", status_code=303)
