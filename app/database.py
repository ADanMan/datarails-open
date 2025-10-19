"""Database helpers for the datarails-open MVP."""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Iterable

SCHEMA = """
CREATE TABLE IF NOT EXISTS financial_facts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source TEXT NOT NULL,
    scenario TEXT NOT NULL,
    period TEXT NOT NULL,
    department TEXT NOT NULL,
    account TEXT NOT NULL,
    value REAL NOT NULL,
    currency TEXT NOT NULL DEFAULT 'USD',
    metadata TEXT
);

CREATE INDEX IF NOT EXISTS idx_financial_facts_period
    ON financial_facts(period);
CREATE INDEX IF NOT EXISTS idx_financial_facts_department
    ON financial_facts(department);
"""

SCHEMA_MIGRATIONS = """
CREATE TABLE IF NOT EXISTS schema_migrations (
    name TEXT PRIMARY KEY,
    applied_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
"""

MIGRATIONS: tuple[tuple[str, str], ...] = (
    (
        "001_create_ai_insights",
        """
        CREATE TABLE IF NOT EXISTS ai_insights (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            actual TEXT NOT NULL,
            budget TEXT NOT NULL,
            prompt TEXT,
            insights TEXT NOT NULL,
            row_count INTEGER NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_ai_insights_created_at
            ON ai_insights(created_at DESC);
        """,
    ),
)


def get_connection(db_path: Path | str) -> sqlite3.Connection:
    """Return a SQLite connection with row factory enabled."""
    path = Path(db_path)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path: Path | str) -> None:
    """Initialise the SQLite database with the required tables."""
    conn = get_connection(db_path)
    try:
        conn.executescript(SCHEMA)
        conn.executescript(SCHEMA_MIGRATIONS)
        run_migrations(conn)
        conn.commit()
    finally:
        conn.close()


def insert_rows(
    conn: sqlite3.Connection,
    rows: Iterable[tuple[str, str, str, str, str, float, str, str | None]],
) -> int:
    """Bulk insert rows into the financial_facts table.

    Returns the number of inserted records.
    """
    cursor = conn.executemany(
        """
        INSERT INTO financial_facts (
            source, scenario, period, department, account, value, currency, metadata
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        list(rows),
    )
    conn.commit()
    return cursor.rowcount


def run_migrations(conn: sqlite3.Connection) -> None:
    """Apply outstanding migrations to the database."""

    conn.executescript(SCHEMA_MIGRATIONS)
    for name, script in MIGRATIONS:
        row = conn.execute(
            "SELECT 1 FROM schema_migrations WHERE name = ?", (name,)
        ).fetchone()
        if row:
            continue
        conn.executescript(script)
        conn.execute(
            "INSERT INTO schema_migrations (name) VALUES (?)",
            (name,),
        )
    conn.commit()
