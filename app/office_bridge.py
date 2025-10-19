"""HTTP bridge exposing key datarails-open operations for the Excel add-in."""
from __future__ import annotations

import os
from contextlib import closing
from pathlib import Path
from typing import Iterable, List, Optional

from cryptography.fernet import Fernet, InvalidToken
from fastapi import FastAPI, HTTPException, Query, Security
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from pydantic import BaseModel, ConfigDict, Field

from . import ai, database, excel_loader, loader, reporting, scenario
from . import insights_repository
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
BRIDGE_TOKEN_ENV = "DATARAILS_OPEN_BRIDGE_TOKEN"
SECRET_STORAGE_DIR = Path(__file__).resolve().parent
SECRET_KEY_PATH = SECRET_STORAGE_DIR / ".bridge_api_secret"
ENCRYPTED_API_KEY_PATH = SECRET_STORAGE_DIR / ".bridge_api_key"


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


class StoreApiKeyRequest(BaseModel):
    api_key: Optional[str] = Field(default=None, alias="apiKey")

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


def _write_secure_file(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    fd = os.open(temp_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "wb") as handle:
        handle.write(data)
    os.replace(temp_path, path)
    os.chmod(path, 0o600)


def _load_or_create_secret_key() -> bytes:
    if SECRET_KEY_PATH.exists():
        return SECRET_KEY_PATH.read_bytes()
    key = Fernet.generate_key()
    _write_secure_file(SECRET_KEY_PATH, key)
    return key


class BridgeService:
    """Wrapper around the core modules that enforces consistent database usage."""

    def __init__(self, database_path: Path | str | None = None) -> None:
        self.database_path = Path(database_path or DEFAULT_DB_PATH)
        self._cipher: Fernet | None = None

    def _connection(self):
        _ensure_database(self.database_path)
        conn = database.get_connection(self.database_path)
        database.run_migrations(conn)
        return conn

    def _get_cipher(self) -> Fernet:
        if self._cipher is None:
            key_bytes = _load_or_create_secret_key()
            self._cipher = Fernet(key_bytes)
        return self._cipher

    def store_api_key(self, api_key: Optional[str]) -> None:
        if not api_key:
            ENCRYPTED_API_KEY_PATH.unlink(missing_ok=True)
            return

        cipher = self._get_cipher()
        encrypted = cipher.encrypt(api_key.encode("utf-8"))
        _write_secure_file(ENCRYPTED_API_KEY_PATH, encrypted)

    def get_api_key(self) -> Optional[str]:
        env_key = os.environ.get(API_KEY_ENV)
        if env_key:
            return env_key
        if not ENCRYPTED_API_KEY_PATH.exists():
            return None

        cipher = self._get_cipher()
        try:
            encrypted = ENCRYPTED_API_KEY_PATH.read_bytes()
            decrypted = cipher.decrypt(encrypted)
        except InvalidToken as exc:  # pragma: no cover - unexpected corruption
            raise HTTPException(
                status_code=500,
                detail="Stored API key could not be decrypted. Reconfigure the key.",
            ) from exc
        return decrypted.decode("utf-8")

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

    def list_scenarios(self) -> list[str]:
        with closing(self._connection()) as conn:
            cursor = conn.execute(
                "SELECT DISTINCT scenario FROM financial_facts "
                "WHERE scenario IS NOT NULL ORDER BY scenario"
            )
            return [value for (value,) in cursor.fetchall()]

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
        api_key = (api_settings.api_key if api_settings else None) or self.get_api_key()
        if not api_key:
            raise HTTPException(
                status_code=400,
                detail=(
                    "An API key is required to generate insights. Configure one via "
                    "/settings/api-key or set the "
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

        row_count = len(structured_rows)

        with closing(self._connection()) as conn:
            repository = insights_repository.InsightsRepository(conn)
            repository.create(
                actual=request.actual_scenario,
                budget=request.budget_scenario,
                prompt=request.prompt,
                insights=insights_text,
                row_count=row_count,
            )

        payload: dict[str, object] = {
            "actualScenario": request.actual_scenario,
            "budgetScenario": request.budget_scenario,
            "insights": insights_text,
            "rowCount": row_count,
        }
        if request.include_rows:
            payload["rows"] = structured_rows
        return payload

    def get_insights_history(
        self,
        *,
        actual: str | None,
        budget: str | None,
        prompt: str | None,
        page: int,
        page_size: int,
    ) -> dict[str, object]:
        page = max(page, 1)
        page_size = max(1, min(page_size, 100))
        offset = (page - 1) * page_size

        with closing(self._connection()) as conn:
            repository = insights_repository.InsightsRepository(conn)
            records, total = repository.list(
                actual=actual,
                budget=budget,
                prompt=prompt,
                limit=page_size,
                offset=offset,
            )

        items = [
            {
                "id": record.id,
                "actual": record.actual,
                "budget": record.budget,
                "prompt": record.prompt,
                "insights": record.insights,
                "rowCount": record.row_count,
                "createdAt": record.created_at,
            }
            for record in records
        ]
        return {
            "items": items,
            "page": page,
            "pageSize": page_size,
            "total": total,
        }


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

    auth_scheme = HTTPBearer(auto_error=False)

    def require_bridge_token(
        credentials: HTTPAuthorizationCredentials | None = Security(auth_scheme),
    ) -> bool:
        expected = os.environ.get(BRIDGE_TOKEN_ENV)
        if not expected:
            raise HTTPException(
                status_code=503,
                detail=(
                    "Bridge token is not configured on the server. Set "
                    f"{BRIDGE_TOKEN_ENV} to enable secure credential storage."
                ),
            )
        if (
            credentials is None
            or credentials.scheme.lower() != "bearer"
            or credentials.credentials != expected
        ):
            raise HTTPException(
                status_code=401,
                detail="Invalid or missing authentication token.",
            )
        return True

    @app.post("/load-data")
    def load_data(request: LoadDataRequest):
        return service.load_data(request)

    @app.get("/reports/summary")
    def reports_summary(scenario: Optional[str] = Query(default=None)):
        return service.refresh_report(scenario)

    @app.post("/scenarios/export")
    def scenarios_export(request: ScenarioExportRequest):
        return service.export_scenario(request)

    @app.get("/scenarios/list")
    def scenarios_list():
        return {"items": service.list_scenarios()}

    @app.post("/insights/variance")
    def insights_variance(request: InsightsRequest):
        return service.generate_insights(request)

    @app.get("/insights/history")
    def insights_history(
        page: int = Query(default=1, ge=1),
        page_size: int = Query(default=20, ge=1, le=100, alias="pageSize"),
        actual: str | None = Query(default=None),
        budget: str | None = Query(default=None),
        prompt: str | None = Query(default=None),
    ):
        return service.get_insights_history(
            actual=actual,
            budget=budget,
            prompt=prompt,
            page=page,
            page_size=page_size,
        )

    @app.post("/settings/api-key", status_code=204)
    def settings_api_key(
        request: StoreApiKeyRequest,
        _: bool = Security(require_bridge_token),
    ) -> None:
        service.store_api_key(request.api_key)

    return app


def app_factory() -> FastAPI:
    """Entry point used by ASGI servers such as uvicorn."""

    return create_app()


app = app_factory()


if __name__ == "__main__":  # pragma: no cover - convenience for manual testing
    import uvicorn

    uvicorn.run("app.office_bridge:app", host="0.0.0.0", port=8000, reload=True)
