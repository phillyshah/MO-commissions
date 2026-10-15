"""/reports — RPT-1 (queue by status, AP workload, approver backlog, touchless rate, quarantine),
cycle time by stage, aging by sponsor, month-end (RPT-2) and audit export (RPT-3).

Cycle time is the primary KPI: receipt to SAP entry, in business days, median. Medians, not means,
because one invoice stuck for a quarter would otherwise hide a healthy month.
"""
from __future__ import annotations

import statistics
from datetime import date, datetime, timedelta, timezone

from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import RedirectResponse, Response

from accounting.business_days import business_days_between
from accounting.db.pool import connection
from accounting.lifecycle import TERMINAL
from accounting.reports import audit_export, month_end
from accounting.web.common import as_uuid, page
from accounting.web.deps import NotFound, require, require_active

router = APIRouter(prefix="/reports")
VIEWERS = require("controller", "supervisor", "cfo", "admin", "auditor")
WINDOW_DAYS = 30


def _median_bd(pairs: list[tuple[datetime, datetime]]) -> int | None:
    vals = [business_days_between(a, b) for a, b in pairs if a and b]
    return int(statistics.median(vals)) if vals else None


def _overview(conn, now: datetime) -> dict:
    since = now - timedelta(days=WINDOW_DAYS)
    by_status = conn.execute(
        """select state::text as state, count(*) as n, min(received_at) as oldest from invoice
           where state::text <> all(%s) group by state order by min(received_at)""", (list(TERMINAL),)).fetchall()
    ap_workload = sum(int(r["n"]) for r in by_status if r["state"] in ("AP_REVIEW", "VENDOR_SETUP"))
    backlog = conn.execute(
        """select u.display_name, count(*) as n, min(r.created_at) as oldest
           from approval_request r join ap_user u on u.id = r.assignee_id
           where r.status = 'open' group by u.display_name order by min(r.created_at)""").fetchall()
    for b in backlog:
        b["oldest_bd"] = business_days_between(b["oldest"], now)
    past_received = conn.execute(
        "select count(*) as n from invoice where received_at >= %s and state <> 'RECEIVED'", (since,)).fetchone()["n"]
    touchless = conn.execute(
        """select count(*) as n from invoice i
           where i.received_at >= %s and (i.routed_at is not null or i.approved_at is not null)
             and not exists (select 1 from invoice_correction c where c.invoice_id = i.id)""", (since,)).fetchone()["n"]
    quarantine = conn.execute("select count(*) as n from quarantine_case where state = 'open'").fetchone()["n"]
    rate = (int(touchless) * 100 // int(past_received)) if past_received else None

    done = conn.execute(
        """select received_at, routed_at, approved_at, entered_at from invoice
           where received_at >= %s - interval '60 days'""", (since,)).fetchall()
    cycle = {
        "receipt_to_routed": _median_bd([(r["received_at"], r["routed_at"]) for r in done if r["routed_at"]]),
        "routed_to_approved": _median_bd([(r["routed_at"], r["approved_at"]) for r in done if r["routed_at"] and r["approved_at"]]),
        "approved_to_entered": _median_bd([(r["approved_at"], r["entered_at"]) for r in done if r["approved_at"] and r["entered_at"]]),
        "receipt_to_entered": _median_bd([(r["received_at"], r["entered_at"]) for r in done if r["entered_at"]]),
    }
    aging = conn.execute(
        """select coalesce(u.display_name, 'No sponsor') as sponsor, count(*) as n,
                  count(*) filter (where i.received_at < %s) as over_7,
                  count(*) filter (where i.received_at < %s) as over_14, min(i.received_at) as oldest
           from invoice i left join vendor v on v.id = i.vendor_id left join ap_user u on u.id = v.default_sponsor_id
           where i.state::text <> all(%s) group by 1 order by min(i.received_at)""",
        (now - timedelta(days=7), now - timedelta(days=14), list(TERMINAL))).fetchall()
    return {"by_status": by_status, "ap_workload": ap_workload, "backlog": backlog, "touchless_rate": rate,
            "touchless_n": int(touchless), "past_received_n": int(past_received), "quarantine": int(quarantine),
            "cycle": cycle, "aging": aging, "window_days": WINDOW_DAYS}


@router.get("")
def overview(request: Request, user=Depends(VIEWERS)):
    now = datetime.now(timezone.utc)
    with connection() as conn:
        data = _overview(conn, now)
    return page(request, user, "reports/overview.html", {"r": data, "as_of": now})


def _parse_cutoff(raw: str | None) -> date:
    if not raw:
        today = datetime.now(timezone.utc).date()
        return today.replace(day=1) - timedelta(days=1)  # last day of the previous month
    try:
        return date.fromisoformat(raw)
    except ValueError:
        raise NotFound()


@router.get("/month-end")
def month_end_page(request: Request, cutoff: str | None = None, user=Depends(VIEWERS)):
    c = _parse_cutoff(cutoff)
    with connection() as conn:
        rows = month_end.received_not_entered(conn, c)
        signed = month_end.signoffs(conn, c)
    total = sum(int(r["gross_cents"] or 0) for r in rows)
    return page(request, user, "reports/month_end.html", {"cutoff": c, "rows": rows, "total_cents": total, "signoffs": signed,
                                                          "can_sign": user.has("controller", "admin")})


@router.post("/month-end/sign-off")
def month_end_sign(request: Request, cutoff: str = Form(...), user=Depends(require("controller", "admin")),
                   _active=Depends(require_active)):
    c = _parse_cutoff(cutoff)
    with connection() as conn:
        rows = month_end.received_not_entered(conn, c)
        month_end.sign_off(conn, c, user.id, rows)
    return RedirectResponse(f"/reports/month-end?cutoff={c.isoformat()}", status_code=303)


@router.get("/audit/{invoice_id}.csv")
def audit_csv(invoice_id: str, user=Depends(VIEWERS)):
    with connection() as conn:
        h = audit_export.history(conn, as_uuid(invoice_id))
    if h is None:
        raise NotFound()
    name = f"MaxxDash-audit-{h['invoice']['portal_number']}.csv"
    return Response(audit_export.to_csv(h), media_type="text/csv; charset=utf-8",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})


@router.get("/audit/{invoice_id}.json")
def audit_json(invoice_id: str, user=Depends(VIEWERS)):
    with connection() as conn:
        h = audit_export.history(conn, as_uuid(invoice_id))
    if h is None:
        raise NotFound()
    name = f"MaxxDash-audit-{h['invoice']['portal_number']}.json"
    return Response(audit_export.to_json(h), media_type="application/json",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})
