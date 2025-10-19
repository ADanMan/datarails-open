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
- **Narrative insights** that summarise scenario variance using OpenAI-compatible
  APIs (`insights`).
- **Scenario modelling** that applies percentage adjustments and stores or
  exports the resulting dataset (`build-scenario`).
- **Excel ribbon add-in** (VBA/.xlam) that talks to a FastAPI bridge for
  loading data, refreshing reports, exporting scenarios, and generating AI
  insights without relying on Office.js.

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

# Generate AI insights using environment-configured credentials
python -m app.main --db financials.db insights --actual actual --budget budget
```

> Tip: after `pip install -e .` you can also use the console script `datarails-open`
> instead of `python -m app.main` for shorter commands.

## AI-powered insights

The `insights` sub-command sends a variance report to an OpenAI-compatible text
generation endpoint and prints or stores the returned narrative summary.

- Provide credentials via CLI flags (`--api-key`, `--api-base`, `--model`) or
  environment variables (`DATARAILS_OPEN_API_KEY`, `DATARAILS_OPEN_API_BASE`,
  `DATARAILS_OPEN_MODEL`).
- Choose between the legacy chat completions endpoint and the modern responses
  endpoint using `--api-mode` (or `DATARAILS_OPEN_API_MODE`). The helper sends
  compatible payloads for both so you can migrate gradually.
- When `--output` is omitted the insights are written to stdout. Combine with
  `--output` and `--format json` to persist both the narrative and the
  underlying rows in a machine-readable file.

Example:

```bash
export DATARAILS_OPEN_API_KEY="sk-..."
python -m app.main --db financials.db insights --actual actual --budget budget \
  --model gpt-4o-mini --api-mode responses --output insights.txt
```

### Security guidance

- Never commit API keys or service endpoints to source control. Prefer
  ephemeral environment variables or a secrets manager provided by your shell,
  container runtime, or CI system.
- If you need to use a `.env` file for local development, limit file
  permissions (`chmod 600 .env`) and add the file to `.gitignore`.
- Rotate credentials regularly and revoke keys that are no longer needed.
- Review your provider's data retention policy before sending sensitive
  financial information and consider masking values where possible.

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

The runtime depends on `openpyxl` for Excel support and `httpx` for the
OpenAI-compatible client. `pytest` powers the test suite.

## Excel add-in (VBA)

The repository now ships with a VBA implementation of the Excel add-in under
`excel_vba/`. It replaces the legacy Office.js task pane with a ribbon tab that
works on any desktop build of Excel, including perpetual editions such as
Office LTSC 2021. The add-in reuses the existing FastAPI bridge
(`app/office_bridge.py`) for all data operations.

### Start the FastAPI bridge

Run the bridge in a terminal before launching Excel. HTTPS is no longer
required because the VBA add-in talks to the backend over HTTP on localhost.

```bash
uvicorn app.office_bridge:app --host 0.0.0.0 --port 8000
```

You can customise the database location and admin token using the environment
variables described in `app/settings.py` (for example
`DATARAILS_OPEN_BRIDGE_TOKEN` to secure the API-key storage endpoint).

### Build the ribbon add-in

Follow the detailed walkthrough in [`excel_vba/README.md`](excel_vba/README.md):

1. Import the modules and `frmSettings` form into a blank workbook.
2. Attach `customUI/customUI14.xml` using the Office RibbonX editor.
3. Save the workbook as `Datarails.xlam` and load it via Excel's Add-ins manager.

The ribbon exposes buttons for loading data, refreshing reports, exporting
scenarios, generating AI insights, fetching history, storing API keys, and
opening the connection-settings dialog. Preferences are persisted in a hidden
`_DatarailsConfig` worksheet inside the add-in workbook.

### Configure AI credentials

The FastAPI bridge reads the same environment variables as the CLI when calling
an OpenAI-compatible endpoint:

- `DATARAILS_OPEN_API_KEY` – optional default key used when no stored or
  request-level key is available.
- `DATARAILS_OPEN_API_BASE` – optional override for the base URL (defaults to
  `https://api.openai.com/v1`).
- `DATARAILS_OPEN_MODEL` – optional model identifier (defaults to
  `gpt-4o-mini`).
- `DATARAILS_OPEN_API_MODE` – optional endpoint selection
  (`chat-completions` or `responses`, default `chat-completions`).

To store a key on the bridge, set `DATARAILS_OPEN_BRIDGE_TOKEN` before starting
uvicorn. The **Store API Key** button on the ribbon will send the key to
`POST /settings/api-key`, and you can clear it by submitting an empty value.

### Legacy Office.js add-in

The previous React/Office.js task pane has been archived under `legacy_web_ui/`.
It remains in the repository for reference but is no longer maintained. If you
still need the web UI, serve it with `npm run dev` inside that directory and
update the manifest as required.

