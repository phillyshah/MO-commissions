"""Is this sender domain pretending to be one we know? `is_lookalike(sender_domain, known_domains)`.

Three families, in the order an attacker reaches for them:
  homoglyph   — characters that render alike: rn/m, vv/w, 0/o, 1/l, Cyrillic а/е/о/р/с/х, punycode.
                Folding both sides and comparing catches "paypa1.com" and "rnicrosoft.com".
  edit        — Levenshtein ≤ 2 on the registrable domain, with the second-level label at least 5
                characters so "abc.com" vs "abd.com" (both short, both legitimate) stays quiet. A
                distance of exactly 2 on a short label is too easy to hit by accident.
  added_words — the known label embedded in a longer one ("paypal-billing.com", "securepaypal.com"),
                or the whole known domain used as a subdomain of something else ("paypal.com.evil.net").
A real subdomain of a known domain (billing.paypal.com) is NOT a lookalike. Exact matches are not
lookalikes either — the caller already knows those are allowed.
"""
from __future__ import annotations

from accounting.vendors.normalize import registrable

_HOMOGLYPH_SEQ = (("rn", "m"), ("vv", "w"), ("cl", "d"), ("1", "l"), ("0", "o"), ("3", "e"), ("5", "s"),
                  ("7", "t"), ("8", "b"), ("|", "l"), ("!", "i"), ("9", "g"), ("2", "z"), ("4", "a"))
_HOMOGLYPH_CHARS = str.maketrans({
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "х": "x", "у": "y", "і": "i", "ј": "j", "ѕ": "s",
    "ԁ": "d", "ɡ": "g", "ⅰ": "i", "ı": "i", "ł": "l", "ö": "o", "ü": "u", "ä": "a", "é": "e", "è": "e",
    "ñ": "n", "ç": "c",
})


def _decode_idna(domain: str) -> str:
    try:
        return domain.encode("ascii").decode("idna") if "xn--" in domain else domain
    except Exception:
        return domain


def fold(domain: str) -> str:
    s = _decode_idna(domain.strip().lower().rstrip(".")).translate(_HOMOGLYPH_CHARS)
    for a, b in _HOMOGLYPH_SEQ:
        s = s.replace(a, b)
    return s


def levenshtein(a: str, b: str) -> int:
    if a == b:
        return 0
    if not a or not b:
        return max(len(a), len(b))
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _sld(domain: str) -> str:
    return registrable(domain).split(".")[0]


def is_lookalike(sender_domain: str | None, known_domains) -> tuple[bool, str]:
    sender = (sender_domain or "").strip().lower().rstrip(".")
    if not sender:
        return False, ""
    sender_reg = registrable(sender)
    folded_sender = fold(sender_reg)
    for raw in known_domains or []:
        known = (raw or "").strip().lower()
        if not known or "@" in known:
            continue
        known_reg = registrable(known)
        if sender == known or sender_reg == known_reg or sender.endswith("." + known_reg):
            continue  # same vendor, possibly a real subdomain — not a lookalike
        if folded_sender == fold(known_reg):
            return True, f"homoglyph of {known_reg}"
        ks, ss = _sld(known_reg), _sld(sender_reg)
        if ks and ss and len(ks) >= 5 and levenshtein(ss, ks) <= 2:
            return True, f"within two edits of {known_reg}"
        if ks and ss and ks != ss and len(ks) >= 4 and ks in ss:
            return True, f"{known_reg} with words added"
        if len(ks) >= 4 and ks in fold(ss) and fold(ss) != ks:
            return True, f"{known_reg} with words added (homoglyph)"
        # "paypal.com.evil.net": the whole known domain appears as a subdomain of something else
        if sender != known_reg and ("." + known_reg + ".") in ("." + sender + "."):
            return True, f"{known_reg} used as a subdomain of {sender_reg}"
    return False, ""
