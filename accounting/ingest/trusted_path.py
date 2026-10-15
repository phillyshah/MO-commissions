"""INT-4: did this message come through the Maxx Microsoft 365 redirect, or straight to accounts@?

We trust exactly one path: Maxx's tenant receives the mail, authenticates it, and a mail-flow rule
redirects it to the Hostinger mailbox with the original headers intact. So two things must both be
true, and both are read from headers the LAST hop wrote, not ones the sender could have forged:
  1. The `Received:` hop that handed the message to Hostinger came from a host whose name ends in one
     of settings.forwarder_host_suffixes (`*.outbound.protection.outlook.com`). Hostinger writes that
     line; a sender cannot.
  2. An `Authentication-Results:` (or `ARC-Authentication-Results:`) header carries the authserv-id
     settings.forwarder_authserv_id — the stamp Maxx's tenant puts on what it authenticated. A sender
     can add a header with that name, but condition 1 means it did not reach us unless it also passed
     through the tenant, and the tenant's own stamp is the one we parse (first matching, top-most).
Anything else is `direct` and is quarantined. Our own notifications (Message-ID recorded in
`notification.message_id_header`, or From = our address), bounces and auto-replies are recorded and
dropped so the portal never answers itself.

The `Received:` walk: headers are prepended, so index 0 is the Hostinger hop. Hostinger may add more
than one internal hop (MX → spool → mailbox); we skip hops whose `from` is local (localhost, 127/::1)
or inside the same domain as its `by`, and judge the first external one. The judgement is recorded in
`message.headers` so a wrong call can be explained from the row.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from email.message import Message

from accounting.config import settings

_RECEIVED_FROM = re.compile(r"^\s*from\s+(\S+)(?:\s*\(([^)]*)\))?", re.I | re.S)
_RECEIVED_BY = re.compile(r"\bby\s+(\S+)", re.I)
_HOSTNAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9.-]*\.[A-Za-z]{2,}")
_LOCAL = ("localhost", "127.0.0.1", "::1", "[127.0.0.1]", "[::1]")


@dataclass
class TrustResult:
    trust: str                       # forwarded_by_maxx | direct | portal_own | bounce
    detail: dict = field(default_factory=dict)
    auth_results: dict = field(default_factory=dict)


def _domain_of(host: str) -> str:
    host = host.strip("[]<>").lower().rstrip(".")
    parts = host.split(".")
    return ".".join(parts[-2:]) if len(parts) >= 2 else host


def parse_received(value: str) -> dict:
    """{from, from_rdns, by} from one Received header; hostnames lower-cased."""
    v = " ".join(value.split())
    out = {"from": "", "from_rdns": "", "by": ""}
    m = _RECEIVED_FROM.match(v)
    if m:
        out["from"] = m.group(1).strip("[]<>;").lower().rstrip(".")
        rd = m.group(2) or ""
        hm = _HOSTNAME.search(rd)
        out["from_rdns"] = hm.group(0).lower() if hm else ""
    b = _RECEIVED_BY.search(v)
    if b:
        out["by"] = b.group(1).strip("[]<>;").lower().rstrip(".")
    return out


def handoff_hop(received_headers: list[str]) -> dict | None:
    """The first hop (top-most) whose `from` is outside the receiving infrastructure."""
    for raw in received_headers:
        hop = parse_received(raw)
        frm = hop["from"] or hop["from_rdns"]
        if not frm or frm in _LOCAL:
            continue
        if hop["by"] and _domain_of(frm) == _domain_of(hop["by"]) and not _matches_forwarder(frm) and not _matches_forwarder(hop["from_rdns"]):
            continue  # internal relay of the receiving side
        return hop
    return None


def _matches_forwarder(host: str) -> bool:
    host = (host or "").lower().rstrip(".")
    return bool(host) and any(host.endswith(sfx.lower()) for sfx in settings.forwarder_host_suffixes)


def parse_auth_results(value: str) -> dict:
    """RFC 8601: `authserv-id [version]; method=result ...`. Microsoft's ARC form `i=1; authserv-id; ...`
    is handled too. Returns {authserv_id, spf, dkim, dmarc:{result, action, header_from}, compauth}."""
    v = " ".join(value.split())
    parts = [p.strip() for p in v.split(";") if p.strip()]
    out: dict = {"authserv_id": "", "raw": v[:1000]}
    if parts and re.match(r"^i=\d+$", parts[0]):
        parts = parts[1:]
    if parts and "=" not in parts[0].split()[0]:
        out["authserv_id"] = parts[0].split()[0].lower()
        parts = parts[1:]
    for p in parts:
        m = re.match(r"^(spf|dkim|dmarc|arc|compauth)=(\w+)(.*)$", p, re.I)
        if not m:
            continue
        method, result, rest = m.group(1).lower(), m.group(2).lower(), m.group(3)
        entry: dict = {"result": result}
        for k, val in re.findall(r"\b([\w.]+)=([^\s;]+)", rest):
            entry[k.lower().replace(".", "_")] = val
        if method == "dmarc" and "action" not in entry:
            entry["action"] = "none"
        out[method] = entry
    return out


def _authserv_matches(auth: dict) -> bool:
    want = settings.forwarder_authserv_id.lower()
    sid = auth.get("authserv_id", "")
    return bool(sid) and (sid == want or sid.endswith("." + want))


def is_auto_reply_or_bounce(msg: Message) -> str | None:
    frm = (msg.get("From") or "").lower()
    if any(k in frm for k in ("mailer-daemon", "postmaster@", "noreply-dmarc", "no-reply@bounce")):
        return "mailer-daemon sender"
    ct = (msg.get_content_type() or "").lower()
    if ct == "multipart/report":
        return "delivery status report"
    auto = (msg.get("Auto-Submitted") or "").strip().lower()
    if auto and auto != "no":
        return f"Auto-Submitted: {auto}"
    if msg.get("X-Autoreply") or msg.get("X-Autorespond") or msg.get("X-Auto-Response-Suppress") == "All":
        return "auto-reply header"
    subj = (msg.get("Subject") or "").lower()
    if subj.startswith(("automatic reply", "auto reply", "auto-reply", "out of office", "undeliverable", "delivery status notification")):
        return "auto-reply subject"
    return None


def decide(msg: Message, *, is_portal_message_id) -> TrustResult:
    """`is_portal_message_id(mid) -> bool` looks up notification.message_id_header (a DB call the caller
    supplies, so this module stays free of connections)."""
    mid = (msg.get("Message-ID") or "").strip()
    frm = (msg.get("From") or "").lower()
    refs = " ".join(filter(None, [msg.get("In-Reply-To") or "", msg.get("References") or ""]))
    if (mid and is_portal_message_id(mid)) or settings.email_from.lower() in frm:
        return TrustResult("portal_own", {"reason": "sent by the portal"})
    bounce_reason = is_auto_reply_or_bounce(msg)
    if bounce_reason:
        return TrustResult("bounce", {"reason": bounce_reason})
    for ref in re.findall(r"<[^>]+>", refs):
        if is_portal_message_id(ref):
            return TrustResult("bounce", {"reason": "reply to a portal notification"})

    received = msg.get_all("Received") or []
    hop = handoff_hop(received)
    hop_ok = bool(hop) and (_matches_forwarder(hop["from"]) or _matches_forwarder(hop["from_rdns"]))

    auth_headers = (msg.get_all("Authentication-Results") or []) + (msg.get_all("ARC-Authentication-Results") or [])
    parsed = [parse_auth_results(h) for h in auth_headers]
    tenant = next((a for a in parsed if _authserv_matches(a)), None)
    auth_results = tenant or (parsed[0] if parsed else {})

    detail = {"handoff_hop": hop, "hop_ok": hop_ok, "tenant_auth_found": tenant is not None,
              "received_count": len(received)}
    if hop_ok and tenant is not None:
        return TrustResult("forwarded_by_maxx", detail, auth_results)
    return TrustResult("direct", detail, auth_results)
