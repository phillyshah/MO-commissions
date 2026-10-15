"""Normalisation shared by matching, screening and duplicate checks. One definition each, because a
name_key computed two ways is two vendors."""
from __future__ import annotations

import re
import unicodedata

_BUSINESS_SUFFIXES = {
    "inc", "incorporated", "llc", "l.l.c", "ltd", "limited", "corp", "corporation", "co", "company", "plc",
    "gmbh", "sa", "s.a", "ag", "bv", "b.v", "pty", "llp", "lp", "pc", "pllc", "the", "and", "&",
}
BUSINESS_TOKENS = _BUSINESS_SUFFIXES | {
    "services", "service", "supply", "supplies", "medical", "surgical", "group", "holdings", "partners",
    "associates", "solutions", "systems", "technologies", "technology", "industries", "international",
    "logistics", "consulting", "enterprises", "distribution", "manufacturing", "labs", "laboratories",
    "hospital", "clinic", "center", "centre", "university", "bank", "trust", "insurance", "studio", "agency",
    "orthopedics", "orthopaedics", "health", "healthcare", "pharma", "biotech", "devices", "instruments",
}


def _ascii(s: str) -> str:
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")


def name_key(name: str | None) -> str:
    """Lower-case ASCII, punctuation removed, legal suffixes dropped, single spaces.
    "Acme Medical Supply, Inc." and "ACME MEDICAL SUPPLY INC" share one key."""
    if not name:
        return ""
    s = _ascii(name).lower()
    s = re.sub(r"[^a-z0-9&\s]", " ", s)
    tokens = [t for t in s.split() if t not in _BUSINESS_SUFFIXES]
    return " ".join(tokens)


def digits(s: str | None) -> str:
    return re.sub(r"\D", "", s or "")


_STREET_WORDS = {"street": "st", "avenue": "ave", "road": "rd", "drive": "dr", "boulevard": "blvd", "lane": "ln", "court": "ct",
                 "place": "pl", "parkway": "pkwy", "highway": "hwy", "circle": "cir", "terrace": "ter", "north": "n", "south": "s",
                 "east": "e", "west": "w", "suite": "", "ste": "", "unit": "", "apt": "", "floor": "", "fl": "", "#": ""}


def remit_key(remit: dict | None) -> str:
    """Address identity: the house/box number from line 1 plus ZIP5. Wording changes ("100 Main Street,
    Suite 4" vs "100 MAIN ST STE 4") do not move it; a different street number or ZIP does. A line with
    no digits falls back to its words with street suffixes normalised."""
    if not remit:
        return ""
    raw = _ascii(str(remit.get("line1") or "")).lower()
    postal = digits(str(remit.get("postal") or ""))[:5]
    numbers = re.findall(r"\d+", raw)
    if numbers:
        line1 = numbers[0]
    else:
        words = [_STREET_WORDS.get(w, w) for w in re.sub(r"[^a-z0-9#\s]", " ", raw).split()]
        line1 = "".join(w for w in words if w)
    return f"{line1}|{postal}" if (line1 or postal) else ""


def remit_wording_key(remit: dict | None) -> str:
    """Every field, normalised but not reduced — what "new remit-to wording" compares."""
    if not remit:
        return ""
    return "|".join(re.sub(r"\s+", " ", _ascii(str(remit.get(k) or "")).lower().strip())
                    for k in ("name", "line1", "line2", "city", "state", "postal", "country"))


def invoice_number_norm(number: str | None) -> str:
    """DUP-1: strip spaces and punctuation, drop leading zeros, upper-case."""
    if not number:
        return ""
    s = re.sub(r"[\s\-_./#:]", "", _ascii(number)).upper()
    # DUP-1 reads literally: zeros at the START of the number ("000123" → "123"); zeros after a prefix
    # ("INV-00123") stay, because "INV-123" may be a different invoice at that vendor.
    return s.lstrip("0")


def email_domain(addr: str | None) -> str:
    if not addr or "@" not in addr:
        return ""
    return addr.rsplit("@", 1)[1].strip().lower().strip(">").rstrip(".")


def registrable(domain: str) -> str:
    """The part a sender controls: last two labels, or three for the usual two-part public suffixes
    (co.uk, com.au, ...). Good enough for lookalike comparison without a public-suffix list."""
    parts = [p for p in domain.lower().split(".") if p]
    if len(parts) <= 2:
        return ".".join(parts)
    if parts[-1] in {"uk", "au", "nz", "za", "br", "jp", "in", "mx"} and parts[-2] in {"co", "com", "org", "net", "gov", "ac", "edu"}:
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


def looks_like_person(name: str | None) -> bool:
    """Two or three alphabetic words, none a business token, no digits: "Jane A Smith". Used for the
    payee-is-a-person hard trigger; a business name with any trade word escapes it."""
    if not name:
        return False
    s = _ascii(name).strip()
    if any(ch.isdigit() for ch in s):
        return False
    words = [w.strip(".,'") for w in s.split()]
    words = [w for w in words if w]
    if not 2 <= len(words) <= 3:
        return False
    if any(w.lower() in BUSINESS_TOKENS for w in words):
        return False
    return all(w.replace("-", "").isalpha() for w in words)
