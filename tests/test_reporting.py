from pathlib import Path

from app import database, loader, reporting


def prepare_db(tmp_path: Path) -> Path:
    db_path = tmp_path / "test.db"
    database.init_db(db_path)
    conn = database.get_connection(db_path)
    try:
        loader.load_file(
            conn,
            tmp_path / "actuals.csv",
            source="seed",
            scenario="actual",
        )
        loader.load_file(
            conn,
            tmp_path / "budget.csv",
            source="seed",
            scenario="budget",
        )
    finally:
        conn.close()
    return db_path


def write_sample_files(tmp_path: Path) -> None:
    (tmp_path / "actuals.csv").write_text(
        """period,department,account,value\n2024-01,Sales,Revenue,120\n2024-01,Sales,Expenses,-40"""
    )
    (tmp_path / "budget.csv").write_text(
        """period,department,account,value\n2024-01,Sales,Revenue,100\n2024-01,Sales,Expenses,-30"""
    )


def test_summarise_by_department(tmp_path: Path):
    write_sample_files(tmp_path)
    db_path = prepare_db(tmp_path)

    conn = database.get_connection(db_path)
    try:
        rows = reporting.summarise_by_department(conn)
    finally:
        conn.close()

    assert rows == [("2024-01", "Sales", 150.0)]

    conn = database.get_connection(db_path)
    try:
        filtered = reporting.summarise_by_department(conn, scenario="actual")
    finally:
        conn.close()

    assert filtered == [("2024-01", "Sales", 80.0)]


def test_variance_report(tmp_path: Path):
    write_sample_files(tmp_path)
    db_path = prepare_db(tmp_path)

    conn = database.get_connection(db_path)
    try:
        rows = reporting.variance_report(conn, actual_scenario="actual", budget_scenario="budget")
    finally:
        conn.close()

    revenue_row = next(row for row in rows if row[2] == "Revenue")
    assert revenue_row[3] == 120.0
    assert revenue_row[4] == 100.0
    assert revenue_row[5] == 20.0
