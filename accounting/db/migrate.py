"""Apply `accounting/db/migrations/*.sql` in filename order, once each.

    python -m accounting.db.migrate            # apply pending
    python -m accounting.db.migrate --status   # list
"""
from __future__ import annotations

import sys
from pathlib import Path

from accounting.db.pool import close_pool, connection

MIGRATIONS_DIR = Path(__file__).parent / "migrations"


def pending(conn) -> list[Path]:
    conn.execute(
        """create table if not exists schema_migrations (
             filename text primary key, applied_at timestamptz not null default now())"""
    )
    done = {r["filename"] for r in conn.execute("select filename from schema_migrations").fetchall()}
    return [p for p in sorted(MIGRATIONS_DIR.glob("*.sql")) if p.name not in done]


def apply_all() -> list[str]:
    applied: list[str] = []
    with connection() as conn:
        for path in pending(conn):
            conn.execute(path.read_text())
            conn.execute("insert into schema_migrations(filename) values (%s)", (path.name,))
            applied.append(path.name)
    return applied


def main(argv: list[str]) -> int:
    if "--status" in argv:
        with connection() as conn:
            for p in pending(conn):
                print("pending ", p.name)
        return 0
    for name in apply_all():
        print("applied ", name)
    return 0


if __name__ == "__main__":
    try:
        code = main(sys.argv[1:])
    finally:
        close_pool()
    sys.exit(code)
