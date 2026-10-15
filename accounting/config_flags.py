"""Read a business switch from `config_current` (ADM-1), with the default the spec names.

Defaults live here and nowhere else so a missing row means the same thing to every package:
  dual_control_enabled  false — Andy acts alone before cutover (VEN-3, recorded decision)
  soft_signals_enabled  true  — soft signals banner from day one; the switch exists to quiet them
"""
from __future__ import annotations

import psycopg

DEFAULTS: dict[str, object] = {
    "dual_control_enabled": False,
    "soft_signals_enabled": True,
}


def config_value(conn: psycopg.Connection, key: str, default=None):
    row = conn.execute("select value from config_current where key = %s", (key,)).fetchone()
    if row is None:
        return DEFAULTS.get(key, default) if default is None else default
    return row["value"]


def config_flag(conn: psycopg.Connection, key: str) -> bool:
    v = config_value(conn, key)
    if isinstance(v, str):
        return v.strip().lower() in ("1", "true", "yes", "on")
    return bool(v)
