"""The one SMTP sender. Ported from usage/app/email.py: implicit TLS on 465, returns a reason
instead of raising.

Why implicit TLS (465) and not STARTTLS: STARTTLS opens in plaintext and upgrades, so a downgrade
is possible and every failure mode doubles. Hostinger offers 465; nothing here needs the other.

Why From must equal SMTP_USER: Hostinger refuses mail whose From header is not the authenticated
mailbox. Catching that here, with a sentence, beats a relay rejection nobody reads.

Why the Message-ID is ours: the ingestor reads the same mailbox this sends from. Recording the
Message-ID we generated lets it drop our own mail, bounces and auto-replies so nothing loops
(NOT-3).
"""
from __future__ import annotations

import logging
import smtplib
import ssl
import uuid
from email.message import EmailMessage
from email.utils import formatdate

from accounting.config import settings

log = logging.getLogger("accounting.notify.smtp")


def new_message_id() -> str:
    """`<uuid@accounting.90ten.life>` — the host is ours, so the ingestor can match on it."""
    return f"<{uuid.uuid4()}@{settings.portal_host}>"


def configuration_problem() -> str | None:
    """A sentence naming the missing value, or None when mail can go out."""
    if not settings.smtp_host:
        return "SMTP_HOST is not set, so nothing can be sent yet."
    if not (settings.smtp_user and settings.smtp_password):
        return "SMTP_USER and SMTP_PASSWORD are not both set."
    if not settings.email_from or "@" not in settings.email_from:
        return "EMAIL_FROM is not an email address."
    if settings.email_from.strip().lower() != settings.smtp_user.strip().lower():
        return "EMAIL_FROM must be the same mailbox as SMTP_USER; the relay refuses anything else."
    return None


def build_message(to: str, subject: str, text: str, headers: dict[str, str] | None = None) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = settings.email_from
    msg["To"] = to
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=False)
    msg["Auto-Submitted"] = "auto-generated"
    for k, v in (headers or {}).items():
        if k in msg:
            del msg[k]
        msg[k] = v
    if "Message-ID" not in msg:
        msg["Message-ID"] = new_message_id()
    msg.set_content(text)
    return msg


def _deliver(msg: EmailMessage) -> None:
    with smtplib.SMTP_SSL(settings.smtp_host, settings.smtp_port, context=ssl.create_default_context(), timeout=30) as smtp:
        smtp.login(settings.smtp_user, settings.smtp_password)
        smtp.send_message(msg)


def send(to: str, subject: str, text: str, *, headers: dict[str, str] | None = None) -> tuple[bool, str, str]:
    """`(ok, reason, message_id)`. Never raises.

    The relay's own words come back on refusal: "login denied" and "mailbox unavailable" are
    different problems with different fixes, and only the relay knows which one happened.
    """
    msg = build_message(to, subject, text, headers)
    message_id = str(msg["Message-ID"])
    problem = configuration_problem()
    if problem:
        return False, problem, message_id
    if not to or "@" not in to:
        return False, "No recipient address, so there is nobody to notify.", message_id
    try:
        _deliver(msg)
    except smtplib.SMTPAuthenticationError as exc:
        log.warning("smtp login rejected: %s", exc)
        return False, f"The relay rejected the login: {exc}", message_id
    except smtplib.SMTPException as exc:
        log.warning("smtp refused: %s", exc)
        return False, f"The relay refused the message: {exc}", message_id
    except OSError as exc:
        log.warning("smtp unreachable: %s", exc)
        return False, f"Could not reach {settings.smtp_host}:{settings.smtp_port} — {exc}", message_id
    return True, "Sent.", message_id
