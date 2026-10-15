"""Received-not-entered as of a cutoff (RPT-2, CUT-2, CUT-3).

Why "as of a cutoff" rather than "right now": the accrual asks what had arrived by the last day of
the month and had NOT yet been keyed into SAP at that moment. An invoice entered on the 3rd of the
next month still belongs on the list for the month it arrived in. The receive date is shown so the
expense lands in the right period.
"""
from __future__ import annotations

import json
from datetime import date, datetime, time, timezone
from zoneinfo import ZoneInfo

import psycopg

from accounting.config import settings
from accounting.lifecycle import log_event

# States that mean "this will never be posted": they never belong on the accrual list.
NEVER_POSTED = ("DIVERTED", "DUPLICATE", "FRAUD", "REJECTED", "VOID")


def cutoff_moment(cutoff: date) -> datetime:
    """The end of the cutoff day in New York, as an aware timestamp."""
    return datetime.combine(cutoff, time(23, 59, 59, 999999), tzinfo=ZoneInfo(settings.timezone)).astimezone(timezone.utc)


def received_not_entered(conn: psycopg.Connection, cutoff: date) -> list[dict]:
    moment = cutoff_moment(cutoff)
    return conn.execute(
        """select i.id, i.portal_number, i.state, i.received_at, i.invoice_date, i.invoice_number, i.currency, i.gross_cents,
                  i.entered_at, i.sap_document_number, v.legal_name as vendor
           from invoice i left join vendor v on v.id = i.vendor_id
           where i.received_at <= %s
             and (i.entered_at is null or i.entered_at > %s)
             and i.state::text <> all(%s)
           order by i.received_at""",
        (moment, moment, list(NEVER_POSTED)),
    ).fetchall()


def sign_off(conn: psycopg.Connection, cutoff: date, by, rows: list[dict]) -> int:
    total = sum(int(r["gross_cents"] or 0) for r in rows)
    return log_event(conn, event_type="month_end_signoff", actor_type="controller", actor_id=by,
                     payload={"cutoff": cutoff.isoformat(), "count": len(rows), "total_cents": str(total),
                              "invoices": [str(r["id"]) for r in rows]})


def signoffs(conn: psycopg.Connection, cutoff: date) -> list[dict]:
    return conn.execute(
        """select e.at, e.payload, u.display_name from event_log e left join ap_user u on u.id = e.actor_id
           where e.event_type = 'month_end_signoff' and e.payload->>'cutoff' = %s order by e.at desc""",
        (cutoff.isoformat(),),
    ).fetchall()
