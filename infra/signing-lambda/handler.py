"""Signing service (SIG-0): verifies a passkey assertion and seals the payload hash with KMS.

Runs as AWS Lambda (Python 3.12) behind a function URL with AWS_IAM auth in Lenny's account. The
portal's IAM user may only invoke this URL. The function never trusts the portal's word for anything
that matters:
  - the passkey public key comes from DynamoDB `ap_passkeys` (written at enrollment through this
    same function, action 'enroll'), not from the request body;
  - the challenge inside clientDataJSON must equal the payload hash;
  - the payload hash is recomputed here from the canonical payload;
  - a click-under-step-up (no assertion) is accepted only when a recent session assertion for the
    same user was verified by this function (table `ap_signing_sessions`).
Only then does KMS sign. The KMS key policy allows kms:Sign to this function's role alone, so the
portal cannot produce a seal without going through these checks.

Request (JSON):
  {"action": "seal", "payload_hash_hex": ..., "assertion": {...}|null, "passkey_public_key_b64": ...|null, "payload": {...}}
  {"action": "session_start", "user_id": ..., "assertion": {...}, "challenge_b64url": ...}
  {"action": "enroll", "user_id": ..., "credential_id_b64url": ..., "public_key_b64": ..., "label": ...}
Response for seal: {"signature_b64": ..., "key_id": ...}
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import time
from typing import Any

import boto3
from webauthn import verify_authentication_response
from webauthn.helpers import base64url_to_bytes, parse_client_data_json
from webauthn.helpers.exceptions import InvalidAuthenticationResponse

KMS_KEY_ID = os.environ["KMS_KEY_ID"]
RP_ID = os.environ.get("RP_ID", "accounting.90ten.life")
EXPECTED_ORIGIN = os.environ.get("EXPECTED_ORIGIN", "https://accounting.90ten.life")
PASSKEYS_TABLE = os.environ.get("PASSKEYS_TABLE", "ap_passkeys")
SESSIONS_TABLE = os.environ.get("SESSIONS_TABLE", "ap_signing_sessions")
SESSION_MAX_SECONDS = int(os.environ.get("SESSION_MAX_SECONDS", "3600"))

_kms = boto3.client("kms")
_ddb = boto3.resource("dynamodb")


class Refused(Exception):
    pass


def canonical(obj: Any) -> bytes:
    """Same rule as accounting.lifecycle.canonical: sorted keys, no whitespace, UTF-8."""
    return json.dumps(obj, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")


def _b64url(s: str) -> bytes:
    return base64url_to_bytes(s)


def _lookup_passkey(credential_id_b64url: str) -> dict:
    item = _ddb.Table(PASSKEYS_TABLE).get_item(Key={"credential_id": credential_id_b64url}).get("Item")
    if not item or item.get("status", "active") != "active":
        raise Refused("unknown or revoked passkey")
    return item


def _verify_assertion(assertion: dict, expected_challenge: bytes) -> dict:
    """Verify with the registry's public key; returns the passkey item. The sign count is updated so a
    cloned authenticator shows up as a count going backwards."""
    item = _lookup_passkey(assertion["rawId"])
    client_data = parse_client_data_json(_b64url(assertion["response"]["clientDataJSON"]))
    if client_data.challenge != expected_challenge:
        raise Refused("challenge does not match payload hash")
    try:
        v = verify_authentication_response(
            credential=assertion, expected_challenge=expected_challenge, expected_rp_id=RP_ID, expected_origin=EXPECTED_ORIGIN,
            credential_public_key=base64.b64decode(item["public_key_b64"]), credential_current_sign_count=int(item.get("sign_count", 0)),
            require_user_verification=True,
        )
    except InvalidAuthenticationResponse as e:
        raise Refused(f"assertion invalid: {e}") from e
    _ddb.Table(PASSKEYS_TABLE).update_item(
        Key={"credential_id": assertion["rawId"]},
        UpdateExpression="set sign_count = :c, last_used_at = :t",
        ExpressionAttributeValues={":c": int(v.new_sign_count), ":t": int(time.time())},
    )
    return item


def _recent_session(user_id: str) -> dict:
    item = _ddb.Table(SESSIONS_TABLE).get_item(Key={"user_id": user_id}).get("Item")
    if not item:
        raise Refused("no signing session on record for this user")
    if int(time.time()) - int(item["started_at"]) > SESSION_MAX_SECONDS:
        raise Refused("signing session on record is too old")
    return item


def _sign(payload_hash: bytes) -> dict:
    r = _kms.sign(KeyId=KMS_KEY_ID, Message=payload_hash, MessageType="DIGEST", SigningAlgorithm="ECDSA_SHA_256")
    return {"signature_b64": base64.b64encode(r["Signature"]).decode(), "key_id": r["KeyId"]}


def seal(body: dict) -> dict:
    payload = body.get("payload")
    payload_hash = bytes.fromhex(body["payload_hash_hex"])
    if len(payload_hash) != 32:
        raise Refused("payload hash must be 32 bytes")
    if payload is None or hashlib.sha256(canonical(payload)).digest() != payload_hash:
        raise Refused("payload does not hash to payload_hash")
    signer = str(payload.get("signer_user_id") or "")
    assertion = body.get("assertion")
    if assertion:
        item = _verify_assertion(assertion, payload_hash)
        if str(item.get("user_id")) != signer:
            raise Refused("passkey belongs to a different user than the payload signer")
        # The portal may send the key it holds; it is checked against the registry, never used instead of it.
        sent = body.get("passkey_public_key_b64")
        if sent and sent != item["public_key_b64"]:
            raise Refused("portal-supplied public key differs from the registry")
    else:
        _recent_session(signer)
    return _sign(payload_hash)


def session_start(body: dict) -> dict:
    challenge = _b64url(body["challenge_b64url"])
    item = _verify_assertion(body["assertion"], challenge)
    if str(item.get("user_id")) != str(body["user_id"]):
        raise Refused("passkey belongs to a different user")
    _ddb.Table(SESSIONS_TABLE).put_item(Item={"user_id": str(body["user_id"]), "credential_id": body["assertion"]["rawId"],
                                             "started_at": int(time.time())})
    return {"ok": True}


def enroll(body: dict) -> dict:
    """Registry write at enrollment. Logged in DynamoDB (enrolled_at) so a silent key swap is visible."""
    table = _ddb.Table(PASSKEYS_TABLE)
    existing = table.get_item(Key={"credential_id": body["credential_id_b64url"]}).get("Item")
    if existing and existing.get("public_key_b64") != body["public_key_b64"]:
        raise Refused("credential already registered with a different public key")
    table.put_item(Item={"credential_id": body["credential_id_b64url"], "user_id": str(body["user_id"]),
                         "public_key_b64": body["public_key_b64"], "sign_count": 0, "label": body.get("label") or "",
                         "status": "active", "enrolled_at": int(time.time())})
    return {"ok": True}


def public_key() -> dict:
    r = _kms.get_public_key(KeyId=KMS_KEY_ID)
    return {"key_id": r["KeyId"], "public_key_der_b64": base64.b64encode(r["PublicKey"]).decode(), "algorithms": r["SigningAlgorithms"]}


def lambda_handler(event, context):
    try:
        raw = event.get("body") or "{}"
        if event.get("isBase64Encoded"):
            raw = base64.b64decode(raw).decode()
        body = json.loads(raw)
        action = body.get("action", "seal")
        if action == "seal":
            out = seal(body)
        elif action == "session_start":
            out = session_start(body)
        elif action == "enroll":
            out = enroll(body)
        elif action == "public_key":
            out = public_key()
        else:
            raise Refused("unknown action")
        return {"statusCode": 200, "headers": {"content-type": "application/json"}, "body": json.dumps(out)}
    except Refused as e:
        return {"statusCode": 403, "headers": {"content-type": "application/json"}, "body": json.dumps({"error": str(e)})}
    except (KeyError, ValueError, TypeError) as e:
        return {"statusCode": 400, "headers": {"content-type": "application/json"}, "body": json.dumps({"error": f"bad request: {e}"})}
