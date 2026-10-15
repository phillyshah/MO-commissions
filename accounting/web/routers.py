"""Every feature router, in one list, so main.py stays short and a missing module fails loudly at import."""
from __future__ import annotations

from fastapi import APIRouter


def all_routers() -> list[APIRouter]:
    from accounting.approvals.router import router as approvals
    from accounting.quarantine.router import router as quarantine
    from accounting.reports.router import router as reports
    from accounting.vendors.router import router as vendors
    from accounting.web.admin import router as admin
    from accounting.web.invoices import router as invoices
    from accounting.web.uploads import router as uploads
    return [invoices, approvals, quarantine, vendors, reports, admin, uploads]
