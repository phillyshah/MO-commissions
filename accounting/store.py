"""Object store for originals and renders (INT-1, DATA-2).

A ref is `"<bucket>/<key>"` and is what the database rows carry (`message.raw_ref`, `attachment.ref`,
`invoice_render.ref`). Keys are content-addressed where it matters (sha256 of the bytes), so writing the
same attachment twice is harmless and a row can always be checked against what it points at.

Two backends. `local` writes under `settings.store_local_dir` (tests, the shadow run on the VPS);
`s3` talks to Supabase Storage's S3-compatible endpoint through boto3, imported lazily so the web
process in test mode needs no AWS libraries. Nothing here deletes: RET-1 says originals live seven
years and no code path removes them, so there is no `delete()` to call by accident.
"""
from __future__ import annotations

import os
import re
from pathlib import Path

from accounting.config import settings

BUCKETS = ("originals", "renders")

_SAFE_KEY = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._/-]*$")


class StoreError(Exception):
    pass


def _check(bucket: str, key: str) -> None:
    if bucket not in BUCKETS:
        raise StoreError(f"unknown bucket {bucket!r}")
    # A key is never derived from an attachment's filename (that is attacker-chosen), but the check
    # stays so a future caller cannot walk out of the bucket directory with "../".
    if not _SAFE_KEY.match(key) or ".." in key.split("/"):
        raise StoreError(f"unsafe object key {key!r}")


def _split(ref: str) -> tuple[str, str]:
    bucket, _, key = ref.partition("/")
    _check(bucket, key)
    return bucket, key


def _bucket_name(bucket: str) -> str:
    return settings.store_s3_bucket_originals if bucket == "originals" else settings.store_s3_bucket_renders


def _s3():
    import boto3  # lazy: only the s3 backend needs it

    return boto3.client(
        "s3",
        endpoint_url=settings.store_s3_endpoint,
        aws_access_key_id=settings.store_s3_access_key,
        aws_secret_access_key=settings.store_s3_secret_key,
        region_name=settings.aws_region,
    )


def _local_path(bucket: str, key: str) -> Path:
    return Path(settings.store_local_dir) / bucket / key


def put(bucket: str, key: str, data: bytes, content_type: str) -> str:
    """Write once. Returns the ref. Writing identical bytes to an existing key is a no-op; different
    bytes to an existing key is refused — originals are immutable."""
    _check(bucket, key)
    if settings.store_backend == "s3":
        _s3().put_object(Bucket=_bucket_name(bucket), Key=key, Body=data, ContentType=content_type)
        return f"{bucket}/{key}"
    path = _local_path(bucket, key)
    if path.exists():
        if path.read_bytes() != data:
            raise StoreError(f"{bucket}/{key} already exists with different content")
        return f"{bucket}/{key}"
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + f".tmp{os.getpid()}")
    tmp.write_bytes(data)
    os.replace(tmp, path)  # atomic: a reader never sees a half-written file
    return f"{bucket}/{key}"


def get(ref: str) -> bytes:
    bucket, key = _split(ref)
    if settings.store_backend == "s3":
        return _s3().get_object(Bucket=_bucket_name(bucket), Key=key)["Body"].read()
    path = _local_path(bucket, key)
    if not path.exists():
        raise StoreError(f"{ref} not found")
    return path.read_bytes()


def exists(ref: str) -> bool:
    bucket, key = _split(ref)
    if settings.store_backend == "s3":
        try:
            _s3().head_object(Bucket=_bucket_name(bucket), Key=key)
            return True
        except Exception:
            return False
    return _local_path(bucket, key).exists()
