# datarails-open

Open-source MVP inspired by the Datarails FP&A platform. The goal is to provide a
lightweight, spreadsheet-friendly toolchain for consolidating financial data,
producing management reports, and experimenting with scenarios without leaving
the command line.

## Features

- **SQLite financial warehouse** with a single table of financial facts.
- **Spreadsheet ingestion** from CSV or Excel workbooks via the `load-data` CLI.
- **Department summaries** with totals per period (`report`).
- **Variance analysis** between budget and actual scenarios (`variance`).
- **Scenario modelling** that applies percentage adjustments and stores or
  exports the resulting dataset (`build-scenario`).

## Quick start

```bash
# Create a virtual environment and install the package locally
python -m venv .venv
source .venv/bin/activate
pip install -e .

# Initialise the database
python -m app.main --db financials.db init-db

# Load sample data from CSV
python -m app.main --db financials.db load-data data/actuals.csv --scenario actual
python -m app.main --db financials.db load-data data/budget.csv --scenario budget

# Load specific worksheets from an Excel workbook
python -m app.main --db financials.db load-data data/plan.xlsx --scenario plan \
  --sheet "Consolidated" --sheet "Adjustments"

# Load a named table from an Excel workbook
python -m app.main --db financials.db load-data data/plan.xlsx --scenario plan \
  --table "SalesTable"

# Generate a consolidated report
python -m app.main --db financials.db report --scenario actual

# Compare actuals vs budget
python -m app.main --db financials.db variance --actual actual --budget budget

# Build a scenario with a 5% increase for Sales department and persist it
python -m app.main --db financials.db build-scenario --source budget --target plan-plus-5 \
  --adjustment 0.05 --department Sales --output data/plan-plus-5.csv
```

> Tip: after `pip install -e .` you can also use the console script `datarails-open`
> instead of `python -m app.main` for shorter commands.

## Data format

Files can be CSV (UTF-8) or `.xlsx` Excel workbooks. In both cases the data must
include the following columns (case-insensitive):

- `period` – typically YYYY-MM (e.g. `2024-01`).
- `department` – e.g. `Sales`.
- `account` – e.g. `Revenue`.
- `value` – numeric amount (positive for revenue, negative for costs).

Optional columns:

- `currency` – defaults to `USD` if omitted.
- `metadata` – free-form text, JSON, or notes associated with the record.

## Development

Install dev dependencies and run tests with:

```bash
pip install -e .[dev]
pytest
```

The runtime depends on `openpyxl` for Excel support and `pytest` is used for
tests, keeping the remainder of the MVP within the Python standard library.

## Excel add-in bridge

The repository ships with a companion Excel add-in under `excel_addin/` that
talks to a lightweight FastAPI bridge. The add-in surfaces three ribbon
commands on macOS Excel:

- **Load data** – call the `/load-data` endpoint to ingest CSV/Excel files.
- **Refresh reports** – call `/reports/summary` and push the aggregated rows
  into a worksheet table named `ReportsTable`.
- **Export scenario** – call `/scenarios/export`, optionally persist the result
  to SQLite, and push the adjusted rows to `ScenariosTable`.

### Start the HTTPS services

Office on macOS requires HTTPS for both the task pane and backend. The
instructions below assume you are using the default local ports defined in the
manifest (3000 for the task pane, 8000 for the Python bridge).

1. Create and trust a development certificate (see the next section for
   platform-specific guidance) and copy the certificate (`devcert.pem`) and key
   (`devcert-key.pem`) into `excel_addin/certs/`.
2. Start the FastAPI bridge with TLS enabled:

   ```bash
   uvicorn app.office_bridge:app \
       --host 0.0.0.0 --port 8000 \
       --ssl-certfile excel_addin/certs/devcert.pem \
       --ssl-keyfile excel_addin/certs/devcert-key.pem
   ```

3. In a separate terminal, serve the task pane assets:

   ```bash
   cd excel_addin
   npm install
   npm run dev
   ```

### macOS sideloading notes

The manifest file `excel_addin/manifest.xml` registers a custom tab and ribbon
group so you can sideload the add-in on macOS Excel (v16.67 or newer):

1. Open Excel and navigate to **Insert → Add-ins → My Add-ins → Upload My
   Add-in…**.
2. Browse to `excel_addin/manifest.xml` and confirm the warning about private
   add-ins.
3. If macOS reports that the HTTPS endpoint is untrusted, open **Keychain
   Access**, locate the certificate used above, and set **Trust → When using
   this certificate** to **Always Trust**. Re-launch Excel afterwards.
4. If you are using a self-signed certificate, ensure the Common Name matches
   `localhost`. Tools like [`mkcert`](https://github.com/FiloSottile/mkcert) can
   generate trusted certificates tied to your local keychain.
5. After sideloading, Excel displays a **Datarails** tab with buttons for each
   supported workflow.

### Manual QA checklist

Use the following steps to validate end-to-end behaviour with a connected
workbook:

1. **Load data** – In the task pane, set the bridge URL to
   `https://localhost:8000`, provide a CSV or XLSX path, and run **Load data**.
   Confirm the success message and inspect the SQLite database if needed.
2. **Refresh reports** – Create a blank worksheet named `Reports` (or let the
   add-in create it). Press **Refresh report table**. A table named
   `ReportsTable` should be created/updated with the aggregated rows and the
   sheet activated.
3. **Export scenario** – Configure the source/target scenarios, adjustment, and
   persistence options, then click **Export scenario to worksheet**. The
   `Scenarios` worksheet should contain a refreshed `ScenariosTable` with the
   adjusted data.
4. Repeat the workflow after modifying the underlying data to ensure the tables
   update in-place.

Automated integration coverage is provided by `tests/test_office_bridge.py`,
which exercises the FastAPI endpoints and verifies persistence/aggregation
behaviour.
