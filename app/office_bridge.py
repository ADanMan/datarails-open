"""HTTP bridge exposing key datarails-open operations for the Excel add-in."""
from __future__ import annotations

import os
from contextlib import closing
from pathlib import Path
from typing import Iterable, List, Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, ConfigDict, Field

from . import ai, database, excel_loader, loader, reporting, scenario
from .settings import (
    API_BASE_ENV,
    API_KEY_ENV,
    API_MODE_ENV,
    MODEL_ENV,
    DEFAULT_API_BASE,
    DEFAULT_API_MODE,
    DEFAULT_MODEL,
)

DEFAULT_DB_PATH = Path(os.environ.get("DATARAILS_DB", "financials.db"))


class LoadDataRequest(BaseModel):
    path: str = Field(..., description="Absolute path to the source file")
    source: str = Field("imports", description="Logical source identifier")
    scenario: str = Field("Actuals", description="Scenario name")
    sheets: Optional[List[str]] = Field(
        default=None,
        description="List of worksheet names to import when loading from Excel",
    )
    tables: Optional[List[str]] = Field(
        default=None,
        description="List of table names to import when loading from Excel",
    )


class ScenarioExportRequest(BaseModel):
    source_scenario: str = Field(..., alias="sourceScenario")
    target_scenario: str = Field(..., alias="targetScenario")
    department: Optional[str] = None
    account: Optional[str] = None
    percentage_change: float = Field(..., alias="percentageChange")
    persist: bool = True

    model_config = ConfigDict(populate_by_name=True)


class AISettings(BaseModel):
    api_key: Optional[str] = Field(default=None, alias="apiKey")
    api_base: Optional[str] = Field(default=None, alias="apiBase")
    model: Optional[str] = None
    mode: Optional[str] = None

    model_config = ConfigDict(populate_by_name=True)


class InsightsRequest(BaseModel):
    actual_scenario: str = Field(..., alias="actualScenario")
    budget_scenario: str = Field(..., alias="budgetScenario")
    prompt: Optional[str] = None
    include_rows: bool = Field(False, alias="includeRows")
    api: Optional[AISettings] = None

    model_config = ConfigDict(populate_by_name=True)


def _normalise_path(path_str: str) -> Path:
    path = Path(path_str).expanduser()
    if not path.is_absolute():
        # Allow paths relative to the repository root for convenience.
        path = Path.cwd() / path
    return path


def _ensure_database(db_path: Path) -> None:
    if not db_path.exists():
        db_path.parent.mkdir(parents=True, exist_ok=True)
        database.init_db(db_path)


class BridgeService:
    """Wrapper around the core modules that enforces consistent database usage."""

    def __init__(self, database_path: Path | str | None = None) -> None:
        self.database_path = Path(database_path or DEFAULT_DB_PATH)

    def _connection(self):
        _ensure_database(self.database_path)
        return database.get_connection(self.database_path)

    def load_data(self, request: LoadDataRequest) -> dict:
        source_path = _normalise_path(request.path)
        if not source_path.exists():
            raise HTTPException(status_code=404, detail=f"File not found: {source_path}")

        suffix = source_path.suffix.lower()
        with closing(self._connection()) as conn:
            try:
                if suffix == ".csv":
                    summary = loader.load_file(
                        conn,
                        source_path,
                        source=request.source,
                        scenario=request.scenario,
                    )
                elif suffix == ".xlsx":
                    summary = excel_loader.load_workbook_file(
                        conn,
                        source_path,
                        source=request.source,
                        scenario=request.scenario,
                        sheets=request.sheets,
                        tables=request.tables,
                    )
                else:
                    raise HTTPException(
                        status_code=400,
                        detail="Only CSV and XLSX files are supported",
                    )
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {
            "rowsLoaded": summary.rows_loaded,
            "source": summary.source,
            "scenario": summary.scenario,
            "message": str(summary),
        }

    def refresh_report(self, scenario_name: Optional[str]) -> dict:
        with closing(self._connection()) as conn:
            rows = reporting.summarise_by_department(conn, scenario=scenario_name)
        serialised = [
            {"period": period, "department": department, "total": total}
            for period, department, total in rows
        ]
        return {"scenario": scenario_name, "rows": serialised}

    def export_scenario(self, request: ScenarioExportRequest) -> dict:
        adjustment = scenario.ScenarioAdjustment(
            department=request.department,
            account=request.account,
            percentage_change=request.percentage_change,
        )
        with closing(self._connection()) as conn:
            rows = scenario.build_scenario(
                conn,
                source_scenario=request.source_scenario,
                adjustments=[adjustment],
            )
            if not rows:
                raise HTTPException(
                    status_code=404,
                    detail=(
                        f"Source scenario '{request.source_scenario}' has no data. "
                        "Ensure it has been loaded before exporting."
                    ),
                )
            payload: Iterable[tuple[str, str, str, str, str, float, str, str]] = (
                (
                    f"scenario:{request.source_scenario}",
                    request.target_scenario,
                    period,
                    department,
                    account,
                    value,
                    currency,
                    metadata,
                )
                for period, department, account, value, currency, metadata in rows
            )
            inserted = 0
            if request.persist:
                inserted = database.insert_rows(conn, payload)
        serialised = [
            {
                "period": period,
                "department": department,
                "account": account,
                "value": value,
                "currency": currency,
                "metadata": metadata,
            }
            for period, department, account, value, currency, metadata in rows
        ]
        message = (
            f"Scenario '{request.target_scenario}' contains {len(serialised)} rows."
        )
        if request.persist:
            message += f" Persisted {inserted} rows to the database."
        return {
            "rows": serialised,
            "message": message,
            "targetScenario": request.target_scenario,
        }

    def generate_insights(self, request: InsightsRequest) -> dict:
        api_settings = request.api
        api_key = (api_settings.api_key if api_settings else None) or os.environ.get(API_KEY_ENV)
        if not api_key:
            raise HTTPException(
                status_code=400,
                detail=(
                    "An API key is required to generate insights. Provide one via "
                    "api.apiKey or the "
                    f"{API_KEY_ENV} environment variable."
                ),
            )

        api_base = (api_settings.api_base if api_settings else None) or os.environ.get(API_BASE_ENV) or DEFAULT_API_BASE
        model = (api_settings.model if api_settings else None) or os.environ.get(MODEL_ENV) or DEFAULT_MODEL
        mode_value = (api_settings.mode if api_settings else None) or os.environ.get(API_MODE_ENV) or DEFAULT_API_MODE
        if mode_value not in {"chat-completions", "responses"}:
            raise HTTPException(
                status_code=400,
                detail="API mode must be either 'chat-completions' or 'responses'.",
            )
        mode = "responses" if mode_value == "responses" else "chat_completions"

        with closing(self._connection()) as conn:
            rows = reporting.variance_report(
                conn,
                actual_scenario=request.actual_scenario,
                budget_scenario=request.budget_scenario,
            )
        structured_rows = reporting.serialise_variance_rows(rows)

        config = ai.AIConfig(api_key=api_key, api_base=api_base, model=model, mode=mode)
        try:
            insights_text = ai.generate_insights(
                structured_rows,
                config,
                prompt=request.prompt,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except RuntimeError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc
        except Exception as exc:  # pragma: no cover - safeguard against unexpected errors
            raise HTTPException(status_code=502, detail=f"AI request failed: {exc}") from exc

        payload: dict[str, object] = {
            "actualScenario": request.actual_scenario,
            "budgetScenario": request.budget_scenario,
            "insights": insights_text,
            "rowCount": len(structured_rows),
        }
        if request.include_rows:
            payload["rows"] = structured_rows
        return payload


def create_app(database_path: Path | str | None = None) -> FastAPI:
    """Create and configure the FastAPI application."""

    service = BridgeService(database_path=database_path)
    app = FastAPI(title="datarails-open office bridge", version="1.0.0")

    allowed_origins = {
        "https://localhost:3000",
        "https://127.0.0.1:3000",
    }
    app.add_middleware(
        CORSMiddleware,
        allow_origins=list(allowed_origins),
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.post("/load-data")
    def load_data(request: LoadDataRequest):
        return service.load_data(request)

    @app.get("/reports/summary")
    def reports_summary(scenario: Optional[str] = Query(default=None)):
        return service.refresh_report(scenario)

    @app.post("/scenarios/export")
    def scenarios_export(request: ScenarioExportRequest):
        return service.export_scenario(request)

    @app.post("/insights/variance")
    def insights_variance(request: InsightsRequest):
        return service.generate_insights(request)

    return app


def app_factory() -> FastAPI:
    """Entry point used by ASGI servers such as uvicorn."""

    return create_app()


app = app_factory()


if __name__ == "__main__":  # pragma: no cover - convenience for manual testing
    import uvicorn

    uvicorn.run("app.office_bridge:app", host="0.0.0.0", port=8000, reload=True)
