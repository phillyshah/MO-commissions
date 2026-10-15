"""Is this exception worth retrying? Ported from usage/app/pipeline/transient.py.

Lives on its own so the extractor and the pipeline agree on the answer. In the usage app the vision
call once swallowed every exception — including the rate-limit and overload errors the retry loop was
written for — so the loop was unreachable. Here the extractor re-raises what this says is transient
and the pipeline leaves the invoice RECEIVED for the next pass.
"""
from __future__ import annotations

import logging
import random
import time

log = logging.getLogger("accounting.capture.transient")

TRANSIENT_TYPES = {
    "ConnectionTerminated", "RemoteProtocolError", "ConnectError", "ConnectTimeout",
    "ReadError", "ReadTimeout", "WriteError", "PoolTimeout", "ConnectionError",
    "APIConnectionError", "APITimeoutError", "RateLimitError", "InternalServerError",
    "ServerDisconnectedError", "OverloadedError",
}

_TRANSIENT_TEXT = (
    "connectionterminated", "goaway", "server disconnected", "connection reset",
    "connection aborted", "overloaded", "timed out", "timeout", "502", "503", "504", "529",
)


def is_transient(exc: Exception) -> bool:
    names = {type(exc).__name__}
    for ctx in (exc.__cause__, exc.__context__):
        if ctx is not None:
            names.add(type(ctx).__name__)
    if names & TRANSIENT_TYPES:
        return True
    blob = f"{type(exc).__name__}: {exc}".lower()
    return any(k in blob for k in _TRANSIENT_TEXT)


def retry(fn, *args, attempts: int = 3, label: str = "", **kwargs):
    """Call fn, retrying transient failures with exponential backoff; re-raise the last one."""
    for attempt in range(attempts):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            if attempt < attempts - 1 and is_transient(e):
                delay = 0.5 * (2 ** attempt) + random.random() * 0.3
                log.warning("transient failure%s (attempt %d/%d), retrying in %.1fs: %s",
                            f" on {label}" if label else "", attempt + 1, attempts, delay, e)
                time.sleep(delay)
                continue
            raise
