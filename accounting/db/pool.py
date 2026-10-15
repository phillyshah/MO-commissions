"""psycopg3 connection pool. Rows come back as dicts.

Money crosses this boundary as integer cents in Postgres (`bigint`) and as `int` in Python —
never float. Anything shown to a person is formatted at the edge (web/filters.py).
"""
from __future__ import annotations

import contextlib
from typing import Any, Iterator

import psycopg
from psycopg.rows import dict_row
from psycopg_pool import ConnectionPool

from accounting.config import settings

_pool: ConnectionPool | None = None


def get_pool() -> ConnectionPool:
    global _pool
    if _pool is None:
        _pool = ConnectionPool(
            settings.database_url,
            min_size=settings.db_pool_min,
            max_size=settings.db_pool_max,
            kwargs={"row_factory": dict_row, "autocommit": False},
            open=True,
        )
    return _pool


def close_pool() -> None:
    global _pool
    if _pool is not None:
        _pool.close()
        _pool = None


@contextlib.contextmanager
def connection() -> Iterator[psycopg.Connection]:
    """One transaction. Commits on clean exit, rolls back on exception."""
    with get_pool().connection() as conn:
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise


def fetch_one(sql: str, params: Any = None, conn: psycopg.Connection | None = None) -> dict | None:
    if conn is not None:
        return conn.execute(sql, params).fetchone()
    with connection() as c:
        return c.execute(sql, params).fetchone()


def fetch_all(sql: str, params: Any = None, conn: psycopg.Connection | None = None) -> list[dict]:
    if conn is not None:
        return conn.execute(sql, params).fetchall()
    with connection() as c:
        return c.execute(sql, params).fetchall()


def execute(sql: str, params: Any = None, conn: psycopg.Connection | None = None) -> int:
    if conn is not None:
        return conn.execute(sql, params).rowcount
    with connection() as c:
        return c.execute(sql, params).rowcount
