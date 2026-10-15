"""In-memory sign-in limiter, keyed on (client address, email) and on the address alone — the same shape
as MaxxDash's lib/rate-limit.ts. Charged before GoTrue is asked; cleared on success."""
from __future__ import annotations

import threading
import time
from collections import deque

from accounting.config import settings

WINDOW = 15 * 60
TOO_MANY = "Too many sign-in attempts. Wait a few minutes and try again."


class _Limiter:
    def __init__(self, limit: int):
        self.limit = limit
        self._hits: dict[str, deque[float]] = {}
        self._lock = threading.Lock()

    def hit(self, key: str) -> tuple[bool, int]:
        now = time.time()
        with self._lock:
            q = self._hits.setdefault(key, deque())
            while q and now - q[0] > WINDOW:
                q.popleft()
            if len(q) >= self.limit:
                return False, int(WINDOW - (now - q[0])) + 1
            q.append(now)
            return True, 0

    def clear(self, key: str) -> None:
        with self._lock:
            self._hits.pop(key, None)


per_account = _Limiter(settings.signin_rate_limit_per_15min)
per_address = _Limiter(settings.signin_rate_limit_per_15min * 5)
