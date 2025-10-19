from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.office_bridge import (
    BRIDGE_TOKEN_ENV,
    ENCRYPTED_API_KEY_PATH,
    SECRET_KEY_PATH,
    create_app,
)


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
    ENCRYPTED_API_KEY_PATH.unlink(missing_ok=True)
    SECRET_KEY_PATH.unlink(missing_ok=True)
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


def test_generate_insights_route_returns_summary(
    monkeypatch: pytest.MonkeyPatch, client: TestClient, tmp_path: Path
) -> None:
    source = tmp_path / "data.csv"
    _write_sample_csv(source)

    client.post(
        "/load-data",
        json={"path": str(source), "source": "csv", "scenario": "Actuals"},
    )
    client.post(
        "/load-data",
        json={"path": str(source), "source": "csv", "scenario": "Budget"},
    )

    captured: dict[str, object] = {}

    def fake_generate_insights(rows, config, prompt=None):  # type: ignore[no-untyped-def]
        captured["rows"] = rows
        captured["config"] = config
        captured["prompt"] = prompt
        return "Variance summary"

    monkeypatch.setattr("app.office_bridge.ai.generate_insights", fake_generate_insights)

    response = client.post(
        "/insights/variance",
        json={
            "actualScenario": "Actuals",
            "budgetScenario": "Budget",
            "prompt": "Explain the movement",
            "includeRows": True,
            "api": {
                "apiKey": "test-key",
                "apiBase": "https://example.test/v1",
                "model": "gpt-test",
                "mode": "responses",
            },
        },
    )

    payload = response.json()
    assert response.status_code == 200
    assert payload["insights"] == "Variance summary"
    assert payload["rowCount"] == 1
    assert payload["rows"][0]["account"] == "Revenue"
    assert captured["prompt"] == "Explain the movement"
    config = captured["config"]
    assert config.api_key == "test-key"
    assert config.api_base == "https://example.test/v1"
    assert config.model == "gpt-test"
    assert config.mode == "responses"

    history = client.get("/insights/history").json()
    assert history["total"] == 1
    assert history["items"][0]["prompt"] == "Explain the movement"
    assert history["items"][0]["actual"] == "Actuals"
    assert history["items"][0]["budget"] == "Budget"
    assert history["items"][0]["rowCount"] == 1


def test_insights_history_filters_and_pagination(
    monkeypatch: pytest.MonkeyPatch, client: TestClient, tmp_path: Path
) -> None:
    source = tmp_path / "data.csv"
    _write_sample_csv(source)

    client.post(
        "/load-data",
        json={"path": str(source), "source": "csv", "scenario": "Actuals"},
    )
    client.post(
        "/load-data",
        json={"path": str(source), "source": "csv", "scenario": "Budget"},
    )

    def fake_generate(rows, config, prompt=None):  # type: ignore[no-untyped-def]
        return f"Summary for {prompt or 'default'}"

    monkeypatch.setattr("app.office_bridge.ai.generate_insights", fake_generate)

    prompts = ["Alpha", "Beta", "Gamma"]
    for prompt in prompts:
        response = client.post(
            "/insights/variance",
            json={
                "actualScenario": "Actuals",
                "budgetScenario": "Budget",
                "prompt": prompt,
                "api": {"apiKey": "test-key"},
            },
        )
        assert response.status_code == 200

    first_page = client.get(
        "/insights/history",
        params={"page": 1, "pageSize": 2},
    ).json()
    assert first_page["page"] == 1
    assert first_page["pageSize"] == 2
    assert first_page["total"] == 3
    assert len(first_page["items"]) == 2

    second_page = client.get(
        "/insights/history",
        params={"page": 2, "pageSize": 2},
    ).json()
    assert second_page["page"] == 2
    assert len(second_page["items"]) == 1

    beta_only = client.get(
        "/insights/history",
        params={"actual": "Actuals", "prompt": "Beta"},
    ).json()
    assert beta_only["total"] == 1
    assert beta_only["items"][0]["prompt"] == "Beta"


def test_store_api_key_and_generate_without_payload(
    monkeypatch: pytest.MonkeyPatch, client: TestClient, tmp_path: Path
) -> None:
    source = tmp_path / "data.csv"
    _write_sample_csv(source)

    client.post(
        "/load-data",
        json={"path": str(source), "source": "csv", "scenario": "Actuals"},
    )
    client.post(
        "/load-data",
        json={"path": str(source), "source": "csv", "scenario": "Budget"},
    )

    captured: dict[str, object] = {}

    def fake_generate_insights(rows, config, prompt=None):  # type: ignore[no-untyped-def]
        captured["rows"] = rows
        captured["config"] = config
        captured["prompt"] = prompt
        return "Variance summary"

    monkeypatch.setattr("app.office_bridge.ai.generate_insights", fake_generate_insights)
    monkeypatch.setenv(BRIDGE_TOKEN_ENV, "bridge-secret")

    response = client.post(
        "/settings/api-key",
        json={"apiKey": "stored-key"},
        headers={"Authorization": "Bearer bridge-secret"},
    )
    assert response.status_code == 204

    response = client.post(
        "/insights/variance",
        json={
            "actualScenario": "Actuals",
            "budgetScenario": "Budget",
            "includeRows": False,
        },
    )

    assert response.status_code == 200
    config = captured["config"]
    assert config.api_key == "stored-key"
    assert response.json()["insights"] == "Variance summary"


def test_generate_insights_missing_key_returns_400(
    monkeypatch: pytest.MonkeyPatch, client: TestClient
) -> None:
    monkeypatch.delenv("DATARAILS_OPEN_API_KEY", raising=False)

    response = client.post(
        "/insights/variance",
        json={"actualScenario": "Actuals", "budgetScenario": "Budget"},
    )

    assert response.status_code == 400
    assert "/settings/api-key" in response.json()["detail"]
