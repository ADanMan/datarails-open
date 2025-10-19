from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.office_bridge import create_app


def _write_sample_csv(path: Path) -> None:
    path.write_text(
        "period,department,account,value\n"
        "2024-Q1,Sales,Revenue,1000\n",
        encoding="utf-8",
    )


def _write_sample_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["period", "department", "account", "value"])
    sheet.append(["2024-Q1", "Sales", "Revenue", 1000])
    workbook.save(path)
    workbook.close()


@pytest.fixture()
def client(tmp_path: Path) -> TestClient:
    db_path = tmp_path / "test.db"
    app = create_app(database_path=db_path)
    return TestClient(app)


def test_load_csv_and_refresh_report(client: TestClient, tmp_path: Path) -> None:
    source = tmp_path / "actuals.csv"
    _write_sample_csv(source)

    response = client.post(
        "/load-data",
        json={
            "path": str(source),
            "source": "csv", 
            "scenario": "Actuals",
        },
    )
    payload = response.json()
    assert response.status_code == 200
    assert payload["rowsLoaded"] == 1
    assert payload["scenario"] == "Actuals"

    report = client.get("/reports/summary", params={"scenario": "Actuals"}).json()
    assert report["scenario"] == "Actuals"
    assert len(report["rows"]) == 1
    assert report["rows"][0]["department"] == "Sales"
    assert report["rows"][0]["total"] == 1000.0


def test_export_scenario_persists_rows(client: TestClient, tmp_path: Path) -> None:
    source = tmp_path / "base.csv"
    _write_sample_csv(source)

    client.post(
        "/load-data",
        json={"path": str(source), "source": "csv", "scenario": "Actuals"},
    )

    response = client.post(
        "/scenarios/export",
        json={
            "sourceScenario": "Actuals",
            "targetScenario": "Forecast",
            "percentageChange": 0.1,
            "persist": True,
        },
    )
    payload = response.json()
    assert response.status_code == 200
    assert pytest.approx(payload["rows"][0]["value"], rel=1e-5) == 1100.0
    assert "Persisted 1 rows" in payload["message"]

    forecast = client.get("/reports/summary", params={"scenario": "Forecast"}).json()
    assert len(forecast["rows"]) == 1
    assert pytest.approx(forecast["rows"][0]["total"], rel=1e-5) == 1100.0


def test_export_missing_source_returns_404(client: TestClient) -> None:
    response = client.post(
        "/scenarios/export",
        json={
            "sourceScenario": "Missing",
            "targetScenario": "Forecast",
            "percentageChange": 0.2,
        },
    )
    assert response.status_code == 404
    assert "Source scenario 'Missing'" in response.json()["detail"]


def test_load_invalid_path_returns_404(client: TestClient) -> None:
    response = client.post(
        "/load-data",
        json={
            "path": "/tmp/does-not-exist.csv",
            "source": "csv",
            "scenario": "Actuals",
        },
    )
    assert response.status_code == 404
    assert "File not found" in response.json()["detail"]


def test_load_csv_missing_required_columns_returns_400(
    client: TestClient, tmp_path: Path
) -> None:
    source = tmp_path / "invalid.csv"
    source.write_text(
        "period,department,value\n2024-Q1,Sales,1000\n",
        encoding="utf-8",
    )

    response = client.post(
        "/load-data",
        json={
            "path": str(source),
            "source": "csv",
            "scenario": "Actuals",
        },
    )

    assert response.status_code == 400
    assert "Missing required columns" in response.json()["detail"]


def test_load_xlsx_missing_sheet_returns_400(
    client: TestClient, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "data.xlsx"
    _write_sample_workbook(workbook_path)

    response = client.post(
        "/load-data",
        json={
            "path": str(workbook_path),
            "source": "excel",
            "scenario": "Actuals",
            "sheets": ["NotPresent"],
        },
    )

    assert response.status_code == 400
    assert "Worksheet 'NotPresent' not found" in response.json()["detail"]


def test_load_xlsx_missing_table_returns_400(
    client: TestClient, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "data.xlsx"
    _write_sample_workbook(workbook_path)

    response = client.post(
        "/load-data",
        json={
            "path": str(workbook_path),
            "source": "excel",
            "scenario": "Actuals",
            "tables": ["MissingTable"],
        },
    )

    assert response.status_code == 400
    assert "Table 'MissingTable' not found" in response.json()["detail"]
