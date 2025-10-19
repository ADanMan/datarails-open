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
- **Excel task pane** that talks to a FastAPI bridge for loading data,
  refreshing reports, exporting scenarios, and generating AI insights directly
  from Excel.

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

## Excel add-in bridge

The repository ships with a companion Excel add-in under `excel_addin/` that
talks to a lightweight FastAPI bridge. The add-in surfaces three ribbon
commands on macOS Excel:

- **Load data** – call the `/load-data` endpoint to ingest CSV/Excel files.
- **Refresh reports** – call `/reports/summary` and push the aggregated rows
  into a worksheet table named `ReportsTable`.
- **Export scenario** – call `/scenarios/export`, optionally persist the result
  to SQLite, and push the adjusted rows to `ScenariosTable`.
- **AI insights** – call `/insights/variance` to request a narrative summary of
  variance data and, optionally, populate an `Insights` worksheet.

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

In addition to environment variables, the bridge can persist an API key in an
encrypted file located under `app/`. To enable secure storage, configure an
administrative bearer token on the backend:

```bash
export DATARAILS_OPEN_BRIDGE_TOKEN="bridge-admin-token"
```

Clients can then call the protected endpoint to update the stored key:

```bash
curl -X POST "https://localhost:8000/settings/api-key" \
  -H "Authorization: Bearer ${DATARAILS_OPEN_BRIDGE_TOKEN}" \
  -H "Content-Type: application/json" \
  -d '{"apiKey": "sk-..."}'
```

The Excel task pane surfaces a **Store personal API key on the bridge** checkbox
that uses this endpoint automatically. Provide the same bearer token in the
task pane's connection settings so the add-in can authenticate. The add-in only
stores the "use personal key" flag locally; the key itself is encrypted on disk
by the backend and reused for subsequent `/insights/variance` requests.

### Generate insights from Excel

Once the backend and task pane are running:

1. Open the **Datarails** tab in Excel and verify the bridge URL matches your
   FastAPI instance (for example `https://localhost:8000`).
2. In the **AI Insights** section select the actual and budget scenarios that
   exist in your SQLite database, optionally tweak the prompt, and provide an
   API key (unless the backend already has `DATARAILS_OPEN_API_KEY`).
3. Choose whether to store the response in the `Insights` worksheet and adjust
   the advanced options (API base, model, endpoint) if required.
4. Click **Generate insights**. The task pane displays the narrative response
   and, when enabled, populates the `Insights` worksheet with both the
   narrative and the underlying variance rows.

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
4. **Generate AI insights** – Populate the **AI Insights** form with actual and
   budget scenarios, provide an API key if needed, and press **Generate
   insights**. The task pane should render the narrative response, and the
   `Insights` worksheet should contain the narrative plus an
   `InsightsVarianceTable` when the checkbox is enabled.
5. Repeat the workflow after modifying the underlying data to ensure the tables
   update in-place.

Automated integration coverage is provided by `tests/test_office_bridge.py`,
which exercises the FastAPI endpoints and verifies persistence/aggregation
behaviour.
