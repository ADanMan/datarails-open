"""Repository utilities for persisting and fetching AI insights."""
from __future__ import annotations

from dataclasses import dataclass
from sqlite3 import Connection
from typing import List, Optional, Tuple


@dataclass(slots=True)
class InsightRecord:
    """Simple representation of an insight stored in the database."""

    id: int
    actual: str
    budget: str
    prompt: str | None
    insights: str
    row_count: int
    created_at: str


class InsightsRepository:
    """Data access layer for the ``ai_insights`` table."""

    def __init__(self, conn: Connection) -> None:
        self._conn = conn

    def create(
        self,
        *,
        actual: str,
        budget: str,
        prompt: Optional[str],
        insights: str,
        row_count: int,
    ) -> int:
        """Persist a new insight and return its identifier."""

        cursor = self._conn.execute(
            """
            INSERT INTO ai_insights (actual, budget, prompt, insights, row_count)
            VALUES (?, ?, ?, ?, ?)
            """,
            (actual, budget, prompt, insights, row_count),
        )
        self._conn.commit()
        return int(cursor.lastrowid)

    def list(
        self,
        *,
        actual: Optional[str] = None,
        budget: Optional[str] = None,
        prompt: Optional[str] = None,
        limit: int,
        offset: int,
    ) -> Tuple[List[InsightRecord], int]:
        """Return paginated insight records and the total count."""

        where_clauses: List[str] = []
        params: List[object] = []
        if actual:
            where_clauses.append("actual = ?")
            params.append(actual)
        if budget:
            where_clauses.append("budget = ?")
            params.append(budget)
        if prompt:
            where_clauses.append("COALESCE(prompt, '') LIKE ?")
            params.append(f"%{prompt}%")

        where_sql = ""
        if where_clauses:
            where_sql = " WHERE " + " AND ".join(where_clauses)

        total_row = self._conn.execute(
            f"SELECT COUNT(*) AS count FROM ai_insights{where_sql}",
            tuple(params),
        ).fetchone()
        total = int(total_row["count"] if total_row else 0)

        rows = self._conn.execute(
            f"""
            SELECT id, actual, budget, prompt, insights, row_count, created_at
            FROM ai_insights{where_sql}
            ORDER BY datetime(created_at) DESC, id DESC
            LIMIT ? OFFSET ?
            """,
            (*params, limit, offset),
        ).fetchall()

        records = [
            InsightRecord(
                id=int(row["id"]),
                actual=str(row["actual"]),
                budget=str(row["budget"]),
                prompt=row["prompt"],
                insights=str(row["insights"]),
                row_count=int(row["row_count"]),
                created_at=str(row["created_at"]),
            )
            for row in rows
        ]
        return records, total
