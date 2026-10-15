"""/upload — a hand upload (INT-6) travels the same road as an emailed invoice.

Why wrap the file in an email message: the ingestor is the ONE entry point, with its idempotency,
trust decision, attachment risk check and splitting. Building a small RFC 5322 message from the
upload means none of that is duplicated here, and the audit trail reads the same for both paths.
`source='upload'` records which road it came by.
"""
from __future__ import annotations

import logging
from email.message import EmailMessage
from email.utils import formatdate, make_msgid

from fastapi import APIRouter, Depends, File, Form, Request, UploadFile
from fastapi.responses import RedirectResponse

from accounting.config import settings
from accounting.web.common import page
from accounting.web.deps import require, require_active

log = logging.getLogger("accounting.web.uploads")
router = APIRouter(prefix="/upload")

UPLOADERS = require("clerk", "supervisor", "controller", "admin")
MAX_BYTES = 25 * 1024 * 1024
ALLOWED = {".pdf": "application/pdf", ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}


def build_raw_message(*, from_addr: str, filename: str, data: bytes, content_type: str, note: str = "") -> bytes:
    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = settings.imap_user
    msg["Subject"] = f"Manual upload: {filename}"
    msg["Date"] = formatdate(localtime=False)
    msg["Message-ID"] = make_msgid(domain=settings.portal_host)
    msg["X-Portal-Source"] = "upload"
    msg.set_content(note.strip() or f"Uploaded through the portal by {from_addr}.")
    maintype, subtype = content_type.split("/", 1)
    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
    return msg.as_bytes()


@router.get("")
def form(request: Request, user=Depends(UPLOADERS)):
    return page(request, user, "uploads/upload.html", {"error": None, "max_mb": MAX_BYTES // (1024 * 1024)})


@router.post("")
async def upload(request: Request, file: UploadFile = File(...), note: str = Form(""), user=Depends(UPLOADERS),
                 _active=Depends(require_active)):
    name = (file.filename or "").strip()
    ext = ("." + name.rsplit(".", 1)[-1].lower()) if "." in name else ""
    ctx = {"max_mb": MAX_BYTES // (1024 * 1024)}
    if ext not in ALLOWED:
        return page(request, user, "uploads/upload.html", {**ctx, "error": "Choose a PDF, PNG or JPG file."}, 400)
    data = await file.read()
    if not data:
        return page(request, user, "uploads/upload.html", {**ctx, "error": "The file is empty."}, 400)
    if len(data) > MAX_BYTES:
        return page(request, user, "uploads/upload.html", {**ctx, "error": f"The file is larger than {ctx['max_mb']} MB."}, 400)
    raw = build_raw_message(from_addr=user.email, filename=name, data=data, content_type=ALLOWED[ext], note=note)
    try:
        from accounting.ingest.intake import ingest_raw_message
        ids = ingest_raw_message(raw, imap_uid=None, source="upload")
    except ImportError:
        log.warning("ingest.intake not available")
        return page(request, user, "uploads/upload.html", {**ctx, "error": "Uploads are not switched on yet. Nothing was saved."}, 503)
    except Exception:
        log.exception("upload intake failed")
        return page(request, user, "uploads/upload.html", {**ctx, "error": "The file could not be taken in. Nothing was saved — try again, or email it to the AP mailbox."}, 500)
    if not ids:
        return page(request, user, "uploads/upload.html", {**ctx, "error": "The file was received but no invoice could be made from it. Check the Invoices list; it may have been set aside or held for review."}, 200)
    return RedirectResponse(f"/invoices/{ids[0]}", status_code=303)
