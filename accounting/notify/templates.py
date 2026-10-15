"""The fixed email template set (NOT-1, NOT-4, NOT-5, REM-3).

Every email the portal sends looks the same, so a user can learn it once. None of them carries an
amount, a vendor name (except the weekly digest's "delinquent items by vendor", which REM-3 asks
for), an invoice number or a link to an invoice. The only link that changes anything is "Freeze
my access", and it only ever reduces access. A phishing email that imitates these has to add a
button or a link — and the footer tells the reader exactly what to do when they see one.

Builders return `(subject, body)` from a context dict. `render(type, ctx, freeze_url)` adds the
common frame: the portal link, the freeze line and the footer.
"""
from __future__ import annotations

from accounting.config import settings

TEMPLATE_VERSION = "v1"

FOOTER = "This mailbox never sends approval buttons or invoice links. If you receive one, forward it to security."
SUBJECT_PREFIX = "[Maxx AP]"


def _plural(n: int, word: str) -> str:
    n = int(n or 0)
    return f"{n} {word}{'' if n == 1 else 's'}"


def _waiting(n: int) -> str:
    return f"You have {_plural(n, 'item')} waiting in the AP dashboard."


# --- builders: (subject, body) --------------------------------------------------------------------

def pending(ctx: dict) -> tuple[str, str]:
    return ("Something is waiting for you in the AP dashboard",
            _waiting(ctx.get("count", 1)) + "\n\nSign in to the portal to see it.")


def reminder(ctx: dict) -> tuple[str, str]:
    return ("Reminder: an approval is waiting for you",
            _waiting(ctx.get("count", 1)) + "\n\nOne of them has been waiting for more than two business days. "
            "Sign in to the portal to act on it.")


def manager_copy(ctx: dict) -> tuple[str, str]:
    return ("An approval assigned to someone on your team is overdue",
            "An approval assigned to a member of your team has been waiting for more than three business days.\n\n"
            "No action is needed from you yet; if it is still open tomorrow the Controller will reassign it.")


def controller_reassign(ctx: dict) -> tuple[str, str]:
    return ("An approval needs reassigning",
            _waiting(ctx.get("count", 1)) + "\n\nAn approval has been waiting for more than four business days. "
            "Please reassign it or decide it yourself from the dashboard.")


def digest(ctx: dict) -> tuple[str, str]:
    """REM-3: counts, oldest age, delinquent items by vendor. Amounts never; HCP items as a count."""
    lines = [_waiting(ctx.get("count", 0))]
    oldest = ctx.get("oldest_business_days")
    if oldest is not None:
        lines.append(f"The oldest has been waiting {_plural(oldest, 'business day')}.")
    by_vendor = ctx.get("delinquent_by_vendor") or []
    if by_vendor:
        lines.append("")
        lines.append("Past due, by vendor:")
        for row in by_vendor:
            lines.append(f"  - {row.get('vendor', 'Unknown vendor')}: {_plural(row.get('count', 0), 'item')}")
    hcp = int(ctx.get("hcp_count") or 0)
    if hcp:
        lines.append("")
        lines.append(f"{_plural(hcp, 'item')} involve a healthcare professional; details are shown only in the portal.")
    return ("Your weekly AP summary", "\n".join(lines))


def enroll_alert(ctx: dict) -> tuple[str, str]:
    return ("A passkey was added to your account",
            "A new passkey was enrolled on your Maxx Accounting Portal account.\n\n"
            "If you did not do this, use the freeze link below right away and tell the Controller.")


def freeze_link(ctx: dict) -> tuple[str, str]:
    return ("Your 'Freeze my access' link",
            "You asked for a link to freeze your own access. Use the line below if you ever think "
            "someone else has your sign-in. It can only reduce access; it cannot approve anything.")


def heartbeat_alert(ctx: dict) -> tuple[str, str]:
    return ("Forwarded mail volume is out of range",
            ctx.get("text") or "Yesterday's forwarded invoice volume was outside the expected range. "
            "Check the Exchange redirect rule and the mailbox.")


def controller_daily(ctx: dict) -> tuple[str, str]:
    lines = ["Your daily AP control summary.", ""]
    lines.append(f"Quarantine, open: {int(ctx.get('quarantine_open') or 0)}")
    lines.append(f"Escalated to you in the last 24 hours: {int(ctx.get('escalations') or 0)}")
    lines.append(f"Access freezes in the last 24 hours: {int(ctx.get('freezes') or 0)}")
    changes = ctx.get("access_changes") or []
    lines.append(f"Access grants, removals and disables in the last 24 hours: {len(changes)}")
    for c in changes:
        lines.append(f"  - {c}")
    lines.append("")
    lines.append(f"Mail heartbeat: {ctx.get('heartbeat') or 'not checked'}")
    return ("Daily AP control summary", "\n".join(lines))


def security_report(ctx: dict) -> tuple[str, str]:
    return ("A quarantined message was reported",
            "The Controller reported a quarantined message as suspicious and blocked its sender. "
            "Details are in the portal's quarantine screen.")


def approved_for_ap(ctx: dict) -> tuple[str, str]:
    return ("An invoice is approved and ready for SAP entry",
            _waiting(ctx.get("count", 1)) + "\n\nIt is in the Approved queue, ready to be keyed into SAP.")


def vendor_setup(ctx: dict) -> tuple[str, str]:
    return ("An invoice needs a vendor set up",
            _waiting(ctx.get("count", 1)) + "\n\nIt is from a vendor the portal does not know yet. "
            "Add the vendor from the Invoices screen to release it.")


BUILDERS = {
    "pending": pending,
    "reminder": reminder,
    "manager_copy": manager_copy,
    "controller_reassign": controller_reassign,
    "digest": digest,
    "enroll_alert": enroll_alert,
    "freeze_link": freeze_link,
    "heartbeat_alert": heartbeat_alert,
    "controller_daily": controller_daily,
    "security_report": security_report,
    "approved_for_ap": approved_for_ap,
    "vendor_setup": vendor_setup,
}


def render(type_: str, ctx: dict | None, freeze_url: str | None) -> tuple[str, str]:
    """The finished subject and body. Unknown types fall back to the generic 'pending' wording,
    so a new intent from another package still produces a safe, link-free email."""
    builder = BUILDERS.get(type_, pending)
    subject, body = builder(ctx or {})
    parts = [body, "", f"Open the AP dashboard: {settings.portal_base_url}/"]
    if freeze_url:
        parts += ["", f"Freeze my access: {freeze_url}",
                  "(Use this only if you think someone else has your sign-in. It signs you out everywhere and "
                  "pauses your approvals; it cannot approve anything.)"]
    parts += ["", "--", FOOTER]
    return f"{SUBJECT_PREFIX} {subject}", "\n".join(parts)
