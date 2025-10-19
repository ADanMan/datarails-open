/* global Office, Excel */

const SETTINGS_KEY = "datarails-open-backend";
const AI_PREFERENCES_KEY = "datarails-open-ai-preferences";
const BRIDGE_TOKEN_KEY = "datarails-open-bridge-token";

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

function getBridgeToken() {
  return window.localStorage.getItem(BRIDGE_TOKEN_KEY) || "";
}

function saveBridgeToken(token) {
  if (token) {
    window.localStorage.setItem(BRIDGE_TOKEN_KEY, token);
  } else {
    window.localStorage.removeItem(BRIDGE_TOKEN_KEY);
  }
}

function getAiPreferences() {
  const raw = window.localStorage.getItem(AI_PREFERENCES_KEY);
  if (!raw) {
    return {};
  }
  try {
    return JSON.parse(raw);
  } catch (error) {
    return {};
  }
}

function saveAiPreferences(preferences) {
  const payload = {};
  Object.entries(preferences).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      payload[key] = value;
    }
  });
  if (Object.keys(payload).length) {
    window.localStorage.setItem(AI_PREFERENCES_KEY, JSON.stringify(payload));
  } else {
    window.localStorage.removeItem(AI_PREFERENCES_KEY);
  }
}

async function callBackend(path, options = {}) {
  const base = getBackendUrl().replace(/\/$/, "");
  const url = `${base}${path}`;
  const headers = {
    "Content-Type": "application/json",
    ...(options.headers || {}),
  };
  const response = await fetch(url, {
    ...options,
    headers,
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Request failed (${response.status}): ${text}`);
  }
  if (response.status === 204) {
    return null;
  }
  const contentType = response.headers.get("content-type") || "";
  if (!contentType.includes("application/json")) {
    return null;
  }
  return response.json();
}

async function storeApiKeyOnBridge(apiKey) {
  const token = getBridgeToken();
  if (!token) {
    throw new Error(
      "Set the bridge admin token in connection settings before storing API keys.",
    );
  }
  const headers = {};
  if (token) {
    headers.Authorization = `Bearer ${token}`;
  }
  await callBackend("/settings/api-key", {
    method: "POST",
    headers,
    body: JSON.stringify({ apiKey: apiKey || null }),
  });
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

async function upsertTable(context, sheet, tableName, headers, rows, startRow = 1) {
  let table = sheet.tables.getItemOrNullObject(tableName);
  await context.sync();
  if (table.isNullObject) {
    const lastColumn = columnAddress(headers.length - 1);
    const topRow = startRow;
    const lastRow = startRow + Math.max(rows.length, 1);
    const address = `${sheet.name}!A${topRow}:${lastColumn}${lastRow}`;
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

async function writeInsightsToWorksheet(result) {
  const rows = Array.isArray(result.rows) ? result.rows : [];
  const headers = ["Period", "Department", "Account", "Actual", "Budget", "Variance"];
  const tableRows = rows.map((row) => [
    row.period,
    row.department,
    row.account,
    row.actual,
    row.budget,
    row.variance,
  ]);

  await Excel.run(async (context) => {
    const sheet = await ensureWorksheet(context, "Insights");
    sheet.getRange("A1").values = [["Narrative insights"]];
    const narrativeRange = sheet.getRange("A2");
    narrativeRange.values = [[result.insights || ""]];
    narrativeRange.format.wrapText = true;
    sheet.getRange("A3").values = [[
      `Actual: ${result.actualScenario || ""} vs Budget: ${result.budgetScenario || ""}`,
    ]];
    await upsertTable(context, sheet, "InsightsVarianceTable", headers, tableRows, 5);
    sheet.activate();
  });
}

async function generateInsights() {
  const existingPreferences = getAiPreferences();
  const actual = $("insights-actual").value.trim();
  const budget = $("insights-budget").value.trim();
  if (!actual || !budget) {
    log("Both actual and budget scenarios are required.");
    return;
  }

  const prompt = $("insights-prompt").value.trim();
  const apiKey = $("insights-api-key").value.trim();
  const apiBase = $("insights-api-base").value.trim();
  const model = $("insights-model").value.trim();
  const mode = $("insights-mode").value.trim();
  const saveToSheet = $("insights-save-sheet").checked;
  const usePersonalKey = $("insights-use-personal-key").checked;

  const payload = {
    actualScenario: actual,
    budgetScenario: budget,
    includeRows: saveToSheet,
  };
  if (prompt) {
    payload.prompt = prompt;
  }

  const apiConfig = {};
  if (apiBase) apiConfig.apiBase = apiBase;
  if (model) apiConfig.model = model;
  if (mode) apiConfig.mode = mode;
  if (Object.keys(apiConfig).length) {
    payload.api = apiConfig;
  }

  saveAiPreferences({
    apiBase,
    model,
    mode,
    usePersonalKey,
  });

  if (usePersonalKey) {
    if (apiKey) {
      log("Storing API key on the bridge...");
      await storeApiKeyOnBridge(apiKey);
    }
  } else if (existingPreferences.usePersonalKey) {
    log("Clearing stored API key on the bridge...");
    await storeApiKeyOnBridge(null);
  }

  log(`Requesting insights for ${actual} vs ${budget}...`);
  const resultEl = $("insights-result");
  resultEl.textContent = "Loading insights...";

  try {
    const result = await callBackend("/insights/variance", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    const narrative = result.insights || "No insights returned.";
    resultEl.textContent = narrative;

    const rowCount = typeof result.rowCount === "number" ? result.rowCount : (Array.isArray(result.rows) ? result.rows.length : 0);
    log(`Received insights for ${rowCount} rows.`);

    if (saveToSheet) {
      await writeInsightsToWorksheet(result);
      log("Insights worksheet updated.");
    }
  } catch (error) {
    resultEl.textContent = "";
    throw error;
  }
}

function initialiseAiForm() {
  const preferences = getAiPreferences();
  if (preferences.apiBase) {
    $("insights-api-base").value = preferences.apiBase;
  }
  if (preferences.model) {
    $("insights-model").value = preferences.model;
  }
  if (preferences.mode) {
    $("insights-mode").value = preferences.mode;
  }
  if (typeof preferences.usePersonalKey === "boolean") {
    $("insights-use-personal-key").checked = preferences.usePersonalKey;
  }
}

async function saveSettings() {
  const url = $("backend-url").value.trim();
  const token = $("bridge-token").value.trim();
  if (!url) {
    log("Please provide a backend base URL.");
    return;
  }
  saveBackendUrl(url);
  saveBridgeToken(token);
  $("connection-status").textContent = `Using backend at ${url}`;
  log(`Saved backend URL: ${url}`);
}

Office.onReady(() => {
  $("backend-url").value = getBackendUrl();
  $("bridge-token").value = getBridgeToken();
  $("connection-status").textContent = `Using backend at ${getBackendUrl()}`;
  initialiseAiForm();
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
  $("generate-insights").addEventListener("click", () => {
    generateInsights().catch((error) => log(error.message));
  });
});
