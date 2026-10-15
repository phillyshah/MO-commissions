"""The parser container entrypoint:  python -m accounting.parse.service

A Unix-socket server (path = settings.parser_socket) that takes one request per connection and
answers with the engine's result. Wire format, both ways: 4-byte big-endian length, then JSON; bytes
travel base64-encoded. Request {bytes, content_type, filename, render?}; response either
{ok: true, pages:[b64], page_texts, text, page_count, hidden_text_flags, risky, risk_reason, metadata}
or {ok: false, error: <text for a person>}.

Every request is handled in a FORKED child (socketserver.ForkingMixIn), for two reasons that are
really one: a malformed file that crashes pdfium takes the child with it, not the server, and the
INT-5 CPU limit is applied with RLIMIT_CPU in the child so a pathological PDF is killed after
`settings.parser_cpu_seconds` instead of pinning the container. Memory, pids and the absence of a
network are the container's job (compose: network_mode none, read-only root, pids_limit, mem_limit).
"""
from __future__ import annotations

import base64
import json
import logging
import os
import resource
import socket
import socketserver
import struct
import sys

from accounting.config import settings
from accounting.parse import engine

log = logging.getLogger("accounting.parse.service")

MAX_FRAME = 256 * 1024 * 1024  # generous: the engine applies the real byte limit


def read_frame(sock: socket.socket) -> dict:
    header = _recv_exact(sock, 4)
    (n,) = struct.unpack("!I", header)
    if n > MAX_FRAME:
        raise ValueError("frame too large")
    return json.loads(_recv_exact(sock, n).decode("utf-8"))


def write_frame(sock: socket.socket, obj: dict) -> None:
    body = json.dumps(obj).encode("utf-8")
    sock.sendall(struct.pack("!I", len(body)) + body)


def _recv_exact(sock: socket.socket, n: int) -> bytes:
    chunks = []
    while n > 0:
        c = sock.recv(min(n, 1 << 20))
        if not c:
            raise ConnectionError("peer closed")
        chunks.append(c)
        n -= len(c)
    return b"".join(chunks)


def handle_request(req: dict) -> dict:
    """Pure: request dict → response dict. Shared with the in-process client so both paths agree."""
    try:
        data = base64.b64decode(req.get("bytes") or b"")
        a = engine.analyze(
            data, req.get("content_type"), req.get("filename"),
            max_pages=settings.parser_max_pages, max_bytes=settings.parser_max_bytes,
            render=bool(req.get("render", True)),
        )
        d = a.as_dict()
        d["pages"] = [base64.b64encode(p).decode("ascii") for p in a.pages]
        d["ok"] = True
        return d
    except engine.ParseError as e:
        return {"ok": False, "error": str(e)}
    except Exception:
        log.exception("parser failed")
        return {"ok": False, "error": "The file could not be processed."}


class _Handler(socketserver.BaseRequestHandler):
    def handle(self) -> None:
        cpu = settings.parser_cpu_seconds
        resource.setrlimit(resource.RLIMIT_CPU, (cpu, cpu + 5))
        try:
            req = read_frame(self.request)
        except Exception:
            write_frame(self.request, {"ok": False, "error": "bad request frame"})
            return
        write_frame(self.request, handle_request(req))


class _Server(socketserver.ForkingMixIn, socketserver.UnixStreamServer):
    allow_reuse_address = True
    max_children = 4


def serve(path: str | None = None) -> None:
    path = path or settings.parser_socket
    if os.path.exists(path):
        os.unlink(path)
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with _Server(path, _Handler) as srv:
        os.chmod(path, 0o660)
        log.info("parser listening on %s", path)
        srv.serve_forever()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, stream=sys.stderr)
    serve(sys.argv[1] if len(sys.argv) > 1 else None)
