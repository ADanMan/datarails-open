"""Data loading utilities using only the Python standard library."""
from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Tuple

REQUIRED_COLUMNS = ["period", "department", "account", "value"]


@dataclass
class LoadSummary:
    rows_loaded: int
    source: str
    scenario: str

    def __str__(self) -> str:  # pragma: no cover - convenience only
        return f"Loaded {self.rows_loaded} rows from {self.source} into scenario {self.scenario}"


def _normalise_row(row: dict[str, str]) -> Tuple[str, str, str, float, str, str]:
    try:
        period = row["period"].strip()
        department = row["department"].strip()
        account = row["account"].strip()
        raw_value = row["value"].replace(",", "").strip()
        value = float(raw_value)
    except KeyError as exc:  # pragma: no cover - defensive branch
        raise ValueError(f"Missing column: {exc.args[0]}") from exc

    currency = (row.get("currency") or "USD").strip()
    metadata = (row.get("metadata") or "").strip()
    return period, department, account, value, currency, metadata


def read_dataset(path: Path | str) -> List[Tuple[str, str, str, float, str, str]]:
    """Read a CSV file and return normalised tuples ready for insertion."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)

    if path.suffix.lower() not in {".csv"}:
        raise ValueError("Only CSV files are supported in the open MVP")

    with path.open(newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        headers = [h.strip().lower() for h in reader.fieldnames or []]
        if not headers:
            return []
        missing = set(REQUIRED_COLUMNS) - set(headers)
        if missing:
            raise ValueError(f"Missing required columns: {sorted(missing)}")

        normalised_rows = []
        for raw in reader:
            lowered = {k.strip().lower(): v for k, v in raw.items() if k}
            normalised_rows.append(_normalise_row(lowered))
    return normalised_rows


def load_file(
    conn,
    path: Path | str,
    *,
    source: str,
    scenario: str,
) -> LoadSummary:
    rows = read_dataset(path)
    payload: Iterable[tuple[str, str, str, str, str, float, str, str]] = (
        (
            source,
            scenario,
            period,
            department,
            account,
            value,
            currency,
            metadata,
        )
        for period, department, account, value, currency, metadata in rows
    )

    cursor = conn.executemany(
        """
        INSERT INTO financial_facts (
            source, scenario, period, department, account, value, currency, metadata
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        list(payload),
    )
    conn.commit()
    return LoadSummary(rows_loaded=cursor.rowcount, source=source, scenario=scenario)
