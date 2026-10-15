"""Field encryption (AES-256-GCM) for the few columns the DB alone must not be able to read:
Supabase refresh tokens in portal_session, TINs, bank fingerprints (DATA-1).

The key is ACCOUNTING_FIELD_KEY (32 bytes, base64). In test mode without a key a fixed dev key is
used so tests run; production refuses to start without one (see main.py).
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import os

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

from accounting.config import settings

_DEV_KEY = hashlib.sha256(b"maxx-accounting-dev-field-key").digest()


def field_key() -> bytes:
    if settings.field_key_b64:
        key = base64.b64decode(settings.field_key_b64)
        if len(key) != 32:
            raise RuntimeError("ACCOUNTING_FIELD_KEY must decode to 32 bytes")
        return key
    if settings.is_test:
        return _DEV_KEY
    raise RuntimeError("ACCOUNTING_FIELD_KEY is required outside test mode")


def encrypt(plaintext: str | bytes, aad: bytes = b"") -> bytes:
    data = plaintext.encode() if isinstance(plaintext, str) else plaintext
    nonce = os.urandom(12)
    return nonce + AESGCM(field_key()).encrypt(nonce, data, aad)


def decrypt(blob: bytes | memoryview, aad: bytes = b"") -> bytes:
    blob = bytes(blob)
    return AESGCM(field_key()).decrypt(blob[:12], blob[12:], aad)


def hmac_fingerprint(*parts: str) -> bytes:
    """Keyed HMAC used for the bank fingerprint (VEN-2): same key, same input, same fingerprint;
    unreadable without the key. Inputs are normalised (digits only) by the caller."""
    return hmac.new(field_key(), "|".join(parts).encode(), hashlib.sha256).digest()


def sha256(data: bytes | str) -> bytes:
    return hashlib.sha256(data.encode() if isinstance(data, str) else data).digest()
