"""Scenario modelling utilities built on top of SQLite."""
from __future__ import annotations

from dataclasses import dataclass
from sqlite3 import Connection
from typing import Iterable, List, Tuple

Row = Tuple[str, str, str, float, str, str]


def fetch_dataset(conn: Connection, scenario: str) -> List[Row]:
    rows = conn.execute(
        "SELECT period, department, account, value, currency, IFNULL(metadata, '') as metadata "
        "FROM financial_facts WHERE scenario = ?",
        (scenario,),
    ).fetchall()
    return [
        (
            row["period"],
            row["department"],
            row["account"],
            float(row["value"]),
            row["currency"],
            row["metadata"],
        )
        for row in rows
    ]


@dataclass
class ScenarioAdjustment:
    department: str | None = None
    account: str | None = None
    percentage_change: float = 0.0

    def matches(self, row: Row) -> bool:
        _, department, account, _, _, _ = row
        if self.department and department.lower() != self.department.lower():
            return False
        if self.account and account.lower() != self.account.lower():
            return False
        return True


def apply_adjustments(
    rows: Iterable[Row],
    adjustments: Iterable[ScenarioAdjustment],
) -> List[Row]:
    adjusted: List[Row] = []
    adjustments = list(adjustments)
    for row in rows:
        period, department, account, value, currency, metadata = row
        adjusted_value = value
        for adj in adjustments:
            if adj.matches(row):
                adjusted_value = adjusted_value * (1 + adj.percentage_change)
        adjusted.append((period, department, account, adjusted_value, currency, metadata))
    return adjusted


def build_scenario(
    conn: Connection,
    *,
    source_scenario: str,
    adjustments: Iterable[ScenarioAdjustment],
) -> List[Row]:
    base_rows = fetch_dataset(conn, source_scenario)
    if not base_rows:
        return []
    return apply_adjustments(base_rows, adjustments)
