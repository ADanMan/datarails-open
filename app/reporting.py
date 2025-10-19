"""Reporting utilities that leverage SQL for aggregation."""
from __future__ import annotations

from sqlite3 import Connection
from typing import List, Tuple


def summarise_by_department(
    conn: Connection,
    *,
    scenario: str | None = None,
) -> List[Tuple[str, str, float]]:
    """Return aggregated totals per department and period."""
    if scenario:
        query = (
            "SELECT period, department, SUM(value) as total "
            "FROM financial_facts WHERE scenario = ? GROUP BY period, department ORDER BY period, department"
        )
        rows = conn.execute(query, (scenario,)).fetchall()
    else:
        query = (
            "SELECT period, department, SUM(value) as total "
            "FROM financial_facts GROUP BY period, department ORDER BY period, department"
        )
        rows = conn.execute(query).fetchall()
    return [(row["period"], row["department"], float(row["total"])) for row in rows]


def variance_report(
    conn: Connection,
    *,
    actual_scenario: str,
    budget_scenario: str,
) -> List[Tuple[str, str, str, float, float, float]]:
    """Produce a variance report between two scenarios."""
    query = """
    SELECT
        period,
        department,
        account,
        SUM(CASE WHEN scenario = ? THEN value ELSE 0 END) AS actual,
        SUM(CASE WHEN scenario = ? THEN value ELSE 0 END) AS budget,
        SUM(CASE WHEN scenario = ? THEN value ELSE 0 END) -
        SUM(CASE WHEN scenario = ? THEN value ELSE 0 END) AS variance
    FROM financial_facts
    WHERE scenario IN (?, ?)
    GROUP BY period, department, account
    ORDER BY period, department, account
    """
    rows = conn.execute(
        query,
        (
            actual_scenario,
            budget_scenario,
            actual_scenario,
            budget_scenario,
            actual_scenario,
            budget_scenario,
        ),
    ).fetchall()
    return [
        (row["period"], row["department"], row["account"], float(row["actual"]), float(row["budget"]), float(row["variance"]))
        for row in rows
    ]
