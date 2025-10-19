/* global Office, Excel */

const SETTINGS_KEY = "datarails-open-backend";

function $(id) {
  return document.getElementById(id);
}

function log(message) {
  const logEl = $("status-log");
  const time = new Date().toLocaleTimeString();
  logEl.textContent = `[${time}] ${message}\n${logEl.textContent}`;
}

function getBackendUrl() {
  return window.localStorage.getItem(SETTINGS_KEY) || "https://localhost:8000";
}

function saveBackendUrl(url) {
  window.localStorage.setItem(SETTINGS_KEY, url);
}

async function callBackend(path, options = {}) {
  const base = getBackendUrl().replace(/\/$/, "");
  const url = `${base}${path}`;
  const response = await fetch(url, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Request failed (${response.status}): ${text}`);
  }
  return response.json();
}

function parseList(value) {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function columnAddress(index) {
  let column = "";
  let dividend = index + 1;
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    column = String.fromCharCode(65 + modulo) + column;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return column;
}

async function ensureWorksheet(context, name) {
  let sheet = context.workbook.worksheets.getItemOrNullObject(name);
  sheet.load("name");
  await context.sync();
  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(name);
  }
  return sheet;
}

async function upsertTable(context, sheet, tableName, headers, rows) {
  let table = sheet.tables.getItemOrNullObject(tableName);
  await context.sync();
  if (table.isNullObject) {
    const lastColumn = columnAddress(headers.length - 1);
    const lastRow = Math.max(rows.length, 1) + 1;
    const address = `${sheet.name}!A1:${lastColumn}${lastRow}`;
    table = sheet.tables.add(address, true /* hasHeaders */);
    table.name = tableName;
  }
  table.getHeaderRowRange().values = [headers];
  try {
    const bodyRange = table.getDataBodyRange();
    bodyRange.clear();
  } catch (error) {
    // Ignore when the table has no body rows yet.
  }
  if (rows.length) {
    table.rows.add(null, rows);
  }
}

async function loadData() {
  const path = $("load-path").value.trim();
  const source = $("load-source").value.trim();
  const scenario = $("load-scenario").value.trim();
  const sheets = $("load-sheets").value.trim();
  const tables = $("load-tables").value.trim();
  if (!path) {
    log("Please provide a file path to load.");
    return;
  }
  const payload = {
    path,
    source: source || "imports",
    scenario: scenario || "Actuals",
    sheets: sheets ? parseList(sheets) : undefined,
    tables: tables ? parseList(tables) : undefined,
  };
  log(`Loading data from ${payload.path}...`);
  const result = await callBackend("/load-data", {
    method: "POST",
    body: JSON.stringify(payload),
  });
  log(result.message || `Loaded ${result.rowsLoaded} rows.`);
}

async function refreshReport() {
  const scenario = $("report-scenario").value.trim();
  log(`Refreshing report for scenario '${scenario || "(all)"}'...`);
  const query = scenario ? `?scenario=${encodeURIComponent(scenario)}` : "";
  const result = await callBackend(`/reports/summary${query}`, {
    method: "GET",
  });
  const headers = ["Period", "Department", "Total"];
  const rows = result.rows.map((row) => [row.period, row.department, row.total]);
  await Excel.run(async (context) => {
    const sheet = await ensureWorksheet(context, "Reports");
    await upsertTable(context, sheet, "ReportsTable", headers, rows);
    sheet.activate();
  });
  log(`Report refreshed with ${rows.length} rows.`);
}

async function exportScenario() {
  const source = $("scenario-source").value.trim();
  const target = $("scenario-target").value.trim();
  const department = $("scenario-department").value.trim();
  const account = $("scenario-account").value.trim();
  const adjustment = parseFloat($("scenario-adjustment").value);
  const persist = $("scenario-persist").checked;
  if (!source || !target || Number.isNaN(adjustment)) {
    log("Source, target, and adjustment are required.");
    return;
  }
  const payload = {
    sourceScenario: source,
    targetScenario: target,
    department: department || undefined,
    account: account || undefined,
    percentageChange: adjustment,
    persist,
  };
  log(`Exporting scenario '${target}' from '${source}'...`);
  const result = await callBackend("/scenarios/export", {
    method: "POST",
    body: JSON.stringify(payload),
  });
  const headers = ["Period", "Department", "Account", "Value", "Currency", "Metadata"];
  const rows = result.rows.map((row) => [
    row.period,
    row.department,
    row.account,
    row.value,
    row.currency,
    row.metadata,
  ]);
  await Excel.run(async (context) => {
    const sheet = await ensureWorksheet(context, "Scenarios");
    await upsertTable(context, sheet, "ScenariosTable", headers, rows);
    sheet.activate();
  });
  log(result.message || `Scenario exported with ${rows.length} rows.`);
}

async function saveSettings() {
  const url = $("backend-url").value.trim();
  if (!url) {
    log("Please provide a backend base URL.");
    return;
  }
  saveBackendUrl(url);
  $("connection-status").textContent = `Using backend at ${url}`;
  log(`Saved backend URL: ${url}`);
}

Office.onReady(() => {
  $("backend-url").value = getBackendUrl();
  $("connection-status").textContent = `Using backend at ${getBackendUrl()}`;
  $("save-settings").addEventListener("click", () => {
    saveSettings().catch((error) => log(error.message));
  });
  $("load-data").addEventListener("click", () => {
    loadData().catch((error) => log(error.message));
  });
  $("refresh-report").addEventListener("click", () => {
    refreshReport().catch((error) => log(error.message));
  });
  $("export-scenario").addEventListener("click", () => {
    exportScenario().catch((error) => log(error.message));
  });
});
