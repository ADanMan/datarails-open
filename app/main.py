"""Command line interface for the datarails-open MVP."""
from __future__ import annotations

import argparse
import csv
from pathlib import Path
from typing import Iterable, Sequence

from . import database, loader, reporting, scenario

DEFAULT_DB = Path("financials.db")


def _ensure_db(db_path: Path) -> None:
    if not db_path.exists():
        database.init_db(db_path)


def _open_connection(db_path: Path):
    _ensure_db(db_path)
    return database.get_connection(db_path)


def _print_table(headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
    if not rows:
        print("(no data)")
        return
    widths = [len(header) for header in headers]
    for row in rows:
        for idx, value in enumerate(row):
            widths[idx] = max(widths[idx], len(f"{value}"))
    header_line = " | ".join(header.ljust(widths[idx]) for idx, header in enumerate(headers))
    separator = "-+-".join("-" * widths[idx] for idx in range(len(headers)))
    print(header_line)
    print(separator)
    for row in rows:
        print(" | ".join(str(value).ljust(widths[idx]) for idx, value in enumerate(row)))


def init_db_command(args: argparse.Namespace) -> None:
    database.init_db(args.db)
    print(f"Database ready at {args.db}")


def load_data_command(args: argparse.Namespace) -> None:
    conn = _open_connection(args.db)
    try:
        summary = loader.load_file(
            conn,
            args.path,
            source=args.source,
            scenario=args.scenario,
        )
    finally:
        conn.close()
    print(summary)


def report_command(args: argparse.Namespace) -> None:
    conn = _open_connection(args.db)
    try:
        rows = reporting.summarise_by_department(conn, scenario=args.scenario)
    finally:
        conn.close()

    if args.output:
        with Path(args.output).open("w", newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh)
            writer.writerow(["period", "department", "total"])
            writer.writerows(rows)
        print(f"Report written to {args.output}")
    else:
        _print_table(["period", "department", "total"], rows)


def variance_command(args: argparse.Namespace) -> None:
    conn = _open_connection(args.db)
    try:
        rows = reporting.variance_report(
            conn,
            actual_scenario=args.actual,
            budget_scenario=args.budget,
        )
    finally:
        conn.close()

    if args.output:
        with Path(args.output).open("w", newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh)
            writer.writerow(["period", "department", "account", "actual", "budget", "variance"])
            writer.writerows(rows)
        print(f"Variance report written to {args.output}")
    else:
        _print_table(["period", "department", "account", "actual", "budget", "variance"], rows)


def build_scenario_command(args: argparse.Namespace) -> None:
    conn = _open_connection(args.db)
    records: list[tuple[str, str, str, str, str, float, str, str]] = []
    try:
        adjustments = [
            scenario.ScenarioAdjustment(
                department=args.department,
                account=args.account,
                percentage_change=args.adjustment,
            )
        ]
        rows = scenario.build_scenario(
            conn,
            source_scenario=args.source,
            adjustments=adjustments,
        )
        if not rows:
            print("Source scenario has no data")
            return

        records = [
            (
                f"scenario:{args.source}",
                args.target,
                period,
                department,
                account,
                value,
                currency,
                metadata,
            )
            for period, department, account, value, currency, metadata in rows
        ]

        if args.persist:
            inserted = database.insert_rows(conn, records)
            print(f"Scenario '{args.target}' stored in the database ({inserted} rows)")
    finally:
        conn.close()

    display_rows = [
        (period, department, account, round(value, 2), currency)
        for (_, _, period, department, account, value, currency, _) in records
    ]
    if args.output:
        with Path(args.output).open("w", newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh)
            writer.writerow(["period", "department", "account", "value", "currency", "metadata"])
            writer.writerows(
                (period, department, account, value, currency, metadata)
                for (_, _, period, department, account, value, currency, metadata) in records
            )
        print(f"Scenario exported to {args.output}")
    else:
        _print_table(["period", "department", "account", "value", "currency"], display_rows)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Open-source FP&A console inspired by Datarails")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB, help="Location of the SQLite database")

    subparsers = parser.add_subparsers(dest="command", required=True)

    init_parser = subparsers.add_parser("init-db", help="Initialise the database")
    init_parser.set_defaults(func=init_db_command)

    load_parser = subparsers.add_parser("load-data", help="Load a CSV file into the warehouse")
    load_parser.add_argument("path", type=Path, help="Path to a CSV file")
    load_parser.add_argument("--source", default="manual-upload", help="Identifier for the data source")
    load_parser.add_argument("--scenario", default="actual", help="Scenario label (e.g. actual, budget)")
    load_parser.set_defaults(func=load_data_command)

    report_parser = subparsers.add_parser("report", help="Generate a consolidated report by department")
    report_parser.add_argument("--scenario", help="Filter report by scenario")
    report_parser.add_argument("--output", help="Optional path to write CSV output")
    report_parser.set_defaults(func=report_command)

    variance_parser = subparsers.add_parser("variance", help="Generate a variance report")
    variance_parser.add_argument("--actual", required=True, help="Scenario representing actuals")
    variance_parser.add_argument("--budget", required=True, help="Scenario representing budget")
    variance_parser.add_argument("--output", help="Optional path to write CSV output")
    variance_parser.set_defaults(func=variance_command)

    scenario_parser = subparsers.add_parser("build-scenario", help="Create a new scenario based on adjustments")
    scenario_parser.add_argument("--source", required=True, help="Scenario to use as a base")
    scenario_parser.add_argument("--target", required=True, help="Name of the scenario to create")
    scenario_parser.add_argument("--adjustment", type=float, default=0.0, help="Percentage adjustment as decimal (e.g. 0.1)")
    scenario_parser.add_argument("--department", help="Optional department filter")
    scenario_parser.add_argument("--account", help="Optional account filter")
    scenario_parser.add_argument(
        "--persist",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Persist the generated scenario to the database (default: True)",
    )
    scenario_parser.add_argument("--output", help="Optional path to export the scenario as CSV")
    scenario_parser.set_defaults(func=build_scenario_command)

    return parser


def main(argv: Iterable[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    main()
