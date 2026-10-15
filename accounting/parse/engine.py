"""What the parser container does to a file: rasterize, read the text layer, flag hidden text, refuse
risky types. Pure functions over bytes — no database, no network, no settings beyond the limits the
caller passes in — so the same code runs inside the container and in-process under pytest.

HIDDEN TEXT IS DETECTED DETERMINISTICALLY (CAP-5). The model is never asked "is there anything
suspicious here"; pdfium tells us the fill colour, font size and position of every text object, and
a regex list names the instruction patterns. Each flag is a hard quarantine trigger downstream. Four
checks:
  white_text   — fill colour within WHITE_TOLERANCE of white (the page background we assume; invoices
                 are printed on white and an attacker who paints a dark background first is caught by
                 the rendered PNG the Controller sees anyway);
  tiny_text    — font size under MIN_FONT_PT (4pt is unreadable on paper and on a 150-dpi render);
  offpage_text — text whose bounds fall entirely outside the media box;
  ai_instruction — a string matching AI_INSTRUCTION_PATTERNS anywhere in the text layer, visible or not.
Invisible render mode (Tr 3) is deliberately NOT a flag: every OCR'd scan carries its text layer that
way, and flagging it would quarantine most real paper invoices.

RISKY TYPES are judged by extension AND content, because a renamed .exe is still an .exe: MZ header,
macro-enabled Office (zip containing vbaProject.bin or a .docm/.xlsm/.pptm name), scripts, archives
containing any of those, and password-protected archives (general-purpose flag bit 0 in a zip entry).
Archives are inspected by listing, never extracted: a 100:1 decompression bomb is refused on its
declared size ratio before a byte is inflated.
"""
from __future__ import annotations

import ctypes
import io
import re
import zipfile
from dataclasses import dataclass, field

WHITE_TOLERANCE = 15          # 0..255 per channel; 240+ reads as white on paper and on screen
MIN_FONT_PT = 4.0
RENDER_DPI = 150              # approvers read a PNG; 150 dpi on Letter is 1275x1650, legible and small
MAX_RENDER_PX = 2200          # long edge cap so a poster-sized page cannot exhaust memory
MAX_ARCHIVE_RATIO = 100       # INT-5: 100:1 decompression limit
MAX_IMAGE_PIXELS = 50_000_000

AI_INSTRUCTION_PATTERNS = [
    r"ignore (?:all |any )?(?:previous|prior|above|earlier) (?:instructions?|prompts?|directions?)",
    r"disregard (?:all |any )?(?:previous|prior|above|earlier) (?:instructions?|prompts?)",
    r"\bas an ai\b",
    r"\bas a language model\b",
    r"\byou are (?:an? )?(?:ai|assistant|chatgpt|claude|llm|language model)\b",
    r"\bassistant\s*:",
    r"\bsystem\s*(?:prompt|message)\s*:",
    r"\b(?:please )?approve this (?:invoice|payment|request)\b",
    r"\bmark (?:this |the )?(?:invoice |payment )?as (?:approved|paid|verified|safe)\b",
    r"\b(?:set|report|return) (?:the )?confidence (?:to|as|of) (?:high|1(?:\.0)?|100)",
    r"\bdo not (?:flag|quarantine|report)\b",
    r"\bthis (?:invoice|document) (?:is|has been) (?:verified|pre-?approved|safe)\b",
    r"\bnew (?:bank|wire|ach|account) (?:details|instructions)\b.{0,80}\b(?:use|update|replace)\b",
    r"<\s*/?\s*(?:system|instructions?|prompt)\s*>",
    r"\[\s*(?:system|inst|instructions?)\s*\]",
]
_AI_RE = re.compile("|".join(f"(?:{p})" for p in AI_INSTRUCTION_PATTERNS), re.IGNORECASE | re.DOTALL)

RISKY_EXTENSIONS = {
    "exe", "dll", "js", "jse", "vbs", "vbe", "ps1", "psm1", "bat", "cmd", "scr", "hta", "com", "pif",
    "msi", "msp", "jar", "wsf", "wsh", "lnk", "reg", "cpl", "docm", "xlsm", "pptm", "dotm", "xltm", "potm",
    "ppam", "xlam", "sldm",
}
ARCHIVE_EXTENSIONS = {"zip", "7z", "rar", "gz", "tgz", "tar", "bz2", "xz", "iso", "img", "cab", "arj", "ace"}
PDF_TYPES = {"application/pdf", "application/x-pdf"}
IMAGE_TYPES = {"image/jpeg", "image/png", "image/tiff", "image/webp", "image/bmp", "image/gif"}


class ParseError(Exception):
    """The file could not be turned into pages. The message is written for a person and is safe to show."""


@dataclass
class Analysis:
    pages: list[bytes] = field(default_factory=list)          # PNG per page
    page_texts: list[str] = field(default_factory=list)
    text: str = ""
    page_count: int = 0
    hidden_text_flags: list[dict] = field(default_factory=list)  # {kind, page, text}
    risky: bool = False
    risk_reason: str | None = None
    metadata: dict = field(default_factory=dict)

    def as_dict(self) -> dict:
        return {
            "pages": self.pages, "page_texts": self.page_texts, "text": self.text, "page_count": self.page_count,
            "hidden_text_flags": self.hidden_text_flags, "risky": self.risky, "risk_reason": self.risk_reason,
            "metadata": self.metadata,
        }


# --------------------------------------------------------------------------------------------------
# Risky types
# --------------------------------------------------------------------------------------------------
def _ext(filename: str | None) -> str:
    name = (filename or "").lower().rstrip(". ")
    return name.rsplit(".", 1)[-1] if "." in name else ""


def is_pdf(data: bytes) -> bool:
    return data[:1024].lstrip().startswith(b"%PDF-")


def assess_risk(filename: str | None, content_type: str | None, data: bytes) -> tuple[bool, str | None]:
    """(risky, reason). Reason text is for the Controller's quarantine screen."""
    ext = _ext(filename)
    if ext in RISKY_EXTENSIONS:
        return True, f"file type .{ext} is not accepted (executable, script or macro-enabled document)"
    if data[:2] == b"MZ":
        return True, "file is a Windows executable regardless of its name"
    if data[:4] == b"\x7fELF":
        return True, "file is an executable regardless of its name"
    if data[:4] == b"PK\x03\x04":
        risky, reason = _inspect_zip(data)
        if risky:
            return True, reason
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" and ext in {"doc", "xls", "ppt", "dot", "xlt", "pot"}:
        # Legacy binary Office. Macros live inside the OLE container and we do not parse OLE here, so
        # the honest answer is "could carry a macro" — the Controller decides.
        return True, "legacy binary Office file; may contain macros"
    if data[:6] in (b"Rar!\x1a\x07",) or data[:6] == b"7z\xbc\xaf\x27\x1c":
        return True, "compressed archive (RAR/7z) cannot be inspected and is not accepted"
    if ext in ARCHIVE_EXTENSIONS and data[:4] != b"PK\x03\x04":
        return True, f"compressed archive .{ext} is not accepted"
    if ext == "pdf" or (content_type or "").lower() in PDF_TYPES:
        if is_pdf(data) and _pdf_has_javascript(data):
            return True, "PDF contains JavaScript or a launch action"
    return False, None


def _inspect_zip(data: bytes) -> tuple[bool, str | None]:
    """Office documents are zips too, so a zip is only risky for what it contains."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
        infos = zf.infolist()
    except Exception:
        return True, "archive is damaged or not readable"
    compressed = sum(i.compress_size for i in infos) or 1
    declared = sum(i.file_size for i in infos)
    if declared > compressed * MAX_ARCHIVE_RATIO and declared > 10 * 1024 * 1024:
        return True, "archive expands more than 100:1 (decompression bomb)"
    for i in infos:
        if i.flag_bits & 0x1:
            return True, "password-protected archive"
        name = i.filename.lower()
        if name.endswith("vbaproject.bin"):
            return True, "Office document contains macros (vbaProject.bin)"
        inner = _ext(name)
        if inner in RISKY_EXTENSIONS:
            return True, f"archive contains a .{inner} file"
        if inner in ARCHIVE_EXTENSIONS:
            return True, "archive contains another archive"
    return False, None


def _pdf_has_javascript(data: bytes) -> bool:
    head = data[:4 * 1024 * 1024]
    return b"/JavaScript" in head or b"/JS " in head or b"/Launch" in head


# --------------------------------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------------------------------
def _textobj_text(raw, obj, textpage) -> str:
    n = raw.FPDFTextObj_GetText(obj, textpage, None, 0)  # byte count incl. UTF-16 terminator
    if n <= 2:
        return ""
    buf = ctypes.create_string_buffer(n)
    raw.FPDFTextObj_GetText(obj, textpage, ctypes.cast(buf, ctypes.POINTER(ctypes.c_ushort)), n)
    return buf.raw[: n - 2].decode("utf-16-le", errors="replace")


def _hidden_text_on_page(raw, page, textpage, page_no: int) -> list[dict]:
    flags: list[dict] = []
    left, bottom, right, top = page.get_mediabox()
    for obj in page.get_objects():
        if obj.type != raw.FPDF_PAGEOBJ_TEXT:
            continue
        text = _textobj_text(raw, obj, textpage).strip()
        if not text:
            continue
        snippet = text[:120]
        r, g, b, a = (ctypes.c_uint() for _ in range(4))
        if raw.FPDFPageObj_GetFillColor(obj, r, g, b, a):
            if a.value > 0 and min(r.value, g.value, b.value) >= 255 - WHITE_TOLERANCE:
                flags.append({"kind": "white_text", "page": page_no, "text": snippet})
            elif a.value == 0:
                flags.append({"kind": "transparent_text", "page": page_no, "text": snippet})
        size = ctypes.c_float()
        if raw.FPDFTextObj_GetFontSize(obj, size) and 0 < size.value < MIN_FONT_PT:
            flags.append({"kind": "tiny_text", "page": page_no, "text": snippet})
        try:
            ol, ob, orr, ot = obj.get_bounds()
            if orr < left or ol > right or ot < bottom or ob > top:
                flags.append({"kind": "offpage_text", "page": page_no, "text": snippet})
        except Exception:
            pass
    return flags


def ai_instruction_flags(text: str, page_no: int) -> list[dict]:
    out = []
    for m in _AI_RE.finditer(text or ""):
        out.append({"kind": "ai_instruction", "page": page_no, "text": m.group(0)[:120]})
    return out


def _render_png(page, dpi: int = RENDER_DPI) -> bytes:
    w, h = page.get_size()
    scale = dpi / 72.0
    if max(w, h) * scale > MAX_RENDER_PX:
        scale = MAX_RENDER_PX / max(w, h)
    pil = page.render(scale=scale).to_pil().convert("RGB")
    buf = io.BytesIO()
    pil.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def analyze_pdf(data: bytes, *, max_pages: int, render: bool = True) -> Analysis:
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw

    try:
        doc = pdfium.PdfDocument(data)
    except Exception:
        raise ParseError("The PDF could not be opened; it may be damaged or password-protected.")
    try:
        n = len(doc)
        if n == 0:
            raise ParseError("The PDF has no pages.")
        if n > max_pages:
            raise ParseError(f"The PDF has {n} pages; the limit is {max_pages}.")
        out = Analysis(page_count=n)
        try:
            out.metadata = {k: v for k, v in doc.get_metadata_dict().items() if v}
        except Exception:
            out.metadata = {}
        for i in range(n):
            page = doc[i]
            try:
                textpage = page.get_textpage()
                try:
                    text = textpage.get_text_range() or ""
                    out.hidden_text_flags.extend(_hidden_text_on_page(raw, page, textpage, i + 1))
                finally:
                    textpage.close()
                out.hidden_text_flags.extend(ai_instruction_flags(text, i + 1))
                out.page_texts.append(text)
                if render:
                    out.pages.append(_render_png(page))
            finally:
                page.close()
        out.text = "\n\f\n".join(out.page_texts)
        return out
    finally:
        doc.close()


# --------------------------------------------------------------------------------------------------
# Images
# --------------------------------------------------------------------------------------------------
def analyze_image(data: bytes) -> Analysis:
    from PIL import Image

    Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS
    try:
        img = Image.open(io.BytesIO(data))
        img.load()
    except Exception:
        raise ParseError("The image could not be opened.")
    pages: list[bytes] = []
    frames = getattr(img, "n_frames", 1)  # multi-page TIFF: one invoice page per frame
    for f in range(min(frames, 50)):
        if frames > 1:
            img.seek(f)
        frame = img.convert("RGB")
        if max(frame.size) > MAX_RENDER_PX:
            ratio = MAX_RENDER_PX / max(frame.size)
            frame = frame.resize((max(1, int(frame.width * ratio)), max(1, int(frame.height * ratio))))
        buf = io.BytesIO()
        frame.save(buf, format="PNG", optimize=True)
        pages.append(buf.getvalue())
    return Analysis(pages=pages, page_texts=[""] * len(pages), text="", page_count=len(pages),
                    metadata={"format": img.format or ""})


# --------------------------------------------------------------------------------------------------
# Entry
# --------------------------------------------------------------------------------------------------
def analyze(data: bytes, content_type: str | None, filename: str | None, *, max_pages: int, max_bytes: int,
            render: bool = True) -> Analysis:
    """One attachment in, one Analysis out. Risky files come back with `risky=True` and no pages — they
    are never opened past the header sniff. Unsupported types raise ParseError."""
    if len(data) > max_bytes:
        raise ParseError(f"The file is {len(data) // (1024 * 1024)} MB; the limit is {max_bytes // (1024 * 1024)} MB.")
    risky, reason = assess_risk(filename, content_type, data)
    if risky:
        return Analysis(risky=True, risk_reason=reason)
    ct = (content_type or "").lower().split(";")[0].strip()
    ext = _ext(filename)
    if is_pdf(data) or ct in PDF_TYPES or ext == "pdf":
        if not is_pdf(data):
            raise ParseError("The file is named .pdf but is not a PDF.")
        return analyze_pdf(data, max_pages=max_pages, render=render)
    if ct in IMAGE_TYPES or ext in {"jpg", "jpeg", "png", "tif", "tiff", "webp", "bmp", "gif"}:
        return analyze_image(data)
    raise ParseError(f"Files of type {ct or ('.' + ext if ext else 'unknown')} cannot be read as an invoice.")
