"""`parse(data, content_type, filename) -> ParseResult` — the portal's only way to open an attachment.

If the parser socket exists the bytes go to the container. If it does not and we are in test mode
the engine runs in-process, so pytest and a developer laptop need no container. Outside test mode a
missing socket is an error, not a fallback: INT-5 says attachments are opened only in the isolated
parser, and silently opening them in the worker would be exactly the control it exists to prevent.
"""
from __future__ import annotations

import base64
import os
import socket
from dataclasses import dataclass, field

from accounting.config import settings
from accounting.parse import engine
from accounting.parse.service import handle_request, read_frame, write_frame

ParseError = engine.ParseError


class ParserUnavailable(Exception):
    pass


@dataclass
class ParseResult:
    pages: list[bytes] = field(default_factory=list)
    page_texts: list[str] = field(default_factory=list)
    text: str = ""
    page_count: int = 0
    hidden_text_flags: list[dict] = field(default_factory=list)
    risky: bool = False
    risk_reason: str | None = None
    metadata: dict = field(default_factory=dict)


def _request(data: bytes, content_type: str, filename: str, render: bool) -> dict:
    return {"bytes": base64.b64encode(data).decode("ascii"), "content_type": content_type,
            "filename": filename, "render": render}


def _call(req: dict) -> dict:
    path = settings.parser_socket
    if os.path.exists(path):
        with socket.socket(socket.AF_UNIX, socket.SOCK_STREAM) as s:
            s.settimeout(settings.parser_cpu_seconds + 30)
            s.connect(path)
            write_frame(s, req)
            return read_frame(s)
    if settings.is_test:
        return handle_request(req)
    raise ParserUnavailable(f"parser socket {path} is not present")


def _result(resp: dict) -> ParseResult:
    if not resp.get("ok"):
        raise ParseError(resp.get("error") or "The file could not be processed.")
    return ParseResult(
        pages=[base64.b64decode(p) for p in resp.get("pages") or []],
        page_texts=list(resp.get("page_texts") or []),
        text=resp.get("text") or "",
        page_count=int(resp.get("page_count") or 0),
        hidden_text_flags=list(resp.get("hidden_text_flags") or []),
        risky=bool(resp.get("risky")),
        risk_reason=resp.get("risk_reason"),
        metadata=dict(resp.get("metadata") or {}),
    )


def parse(data: bytes, content_type: str, filename: str) -> ParseResult:
    return _result(_call(_request(data, content_type, filename, render=True)))


def page_texts(data: bytes, content_type: str, filename: str) -> ParseResult:
    """Text layer only (no PNGs) — what the splitter needs to find invoice-number boundaries."""
    return _result(_call(_request(data, content_type, filename, render=False)))
