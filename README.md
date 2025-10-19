# datarails-open

Open-source MVP inspired by the Datarails FP&A platform. The goal is to provide a
lightweight, spreadsheet-friendly toolchain for consolidating financial data,
producing management reports, and experimenting with scenarios without leaving
the command line.

## Features

- **SQLite financial warehouse** with a single table of financial facts.
- **CSV ingestion** via a lightweight CLI (`load-data`).
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

# Load sample data
python -m app.main --db financials.db load-data data/actuals.csv --scenario actual
python -m app.main --db financials.db load-data data/budget.csv --scenario budget

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

Files must be CSV (UTF-8) and include the following columns (case-insensitive):

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

The project intentionally sticks to the Python standard library (plus `pytest`
for tests) to keep the MVP easy to run in restricted environments.
