# Datarails VBA Add-in

This directory contains the source files for the re-packaged Datarails Excel add-in.  The goal is to ship a single `Datarails.xlam` file that exposes the same FastAPI bridge used by the legacy task pane UI, but with a ribbon-based, VBA-driven experience that runs on any supported desktop edition of Excel.

## Contents

```
excel_vba/
├── customUI/
│   └── customUI14.xml       # Ribbon definition for the Datarails tab
├── src/
│   ├── Forms/
│   │   └── frmSettings.frm  # VBA user form for configuring the backend connection
│   └── Modules/             # Exported standard modules
│       ├── AIInsights.bas
│       ├── API_Client.bas
│       ├── Config.bas
│       ├── DataLoader.bas
│       ├── JsonConverter.bas
│       ├── ReportLoader.bas
│       ├── RibbonCallbacks.bas
│       ├── ScenarioExporter.bas
│       ├── Settings.bas
│       └── Utils.bas
└── README.md                # This file
```

> **Note**
> The add-in depends on [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) for JSON parsing.  The module is included verbatim as `JsonConverter.bas` under its MIT licence.

## Building `Datarails.xlam`

1. Open a blank workbook in Excel and press `Alt+F11` to launch the VBA editor.
2. Import all modules and the `frmSettings` userform from `excel_vba/src` (`File → Import File…`).
3. Add a new module named `ThisWorkbook` (or reuse the existing one) and mark the VBA project as trusted if prompted.
4. Use the **Custom UI Editor for Microsoft Office** (or the modern Office RibbonX Editor) to attach `customUI/customUI14.xml` to the workbook.  The callbacks declared in `RibbonCallbacks.bas` map 1:1 with the controls in the XML.
5. Save the workbook as an add-in (`File → Save As…`, choose `Excel Add-In (*.xlam)`) and name it `Datarails.xlam`.
6. Copy the resulting file to a trusted location and load it into Excel (`File → Options → Add-ins → Manage: Excel Add-ins → Go…`).

## Runtime configuration

The add-in stores connection details and user preferences in a hidden worksheet named `_DatarailsConfig`.  Use the **Connection Settings** button on the Datarails ribbon tab to provide:

- the FastAPI backend URL (defaults to `http://localhost:8000`), and
- the bridge administrator token (required for securely storing OpenAI-style API keys on the server).

Other settings – such as default scenarios, prompts, and load parameters – are persisted automatically after each command is executed.  The `_DatarailsConfig` sheet is marked as `xlSheetVeryHidden` so that end users do not edit it accidentally.

## Mapping of ribbon commands

| Ribbon button | VBA entry point | Backend endpoint | Description |
| --- | --- | --- | --- |
| **Load Data** | `DataLoader.LoadDataCommand` | `POST /load-data` | Imports CSV/XLSX files into the database using the FastAPI bridge. |
| **Refresh Reports** | `ReportLoader.RefreshReportsCommand` | `GET /reports/summary` | Builds the departmental summary report on the `Reports` worksheet. |
| **Export Scenario** | `ScenarioExporter.ExportScenarioCommand` | `POST /scenarios/export` | Creates a derived scenario, optionally persisting it to the database. |
| **Generate AI Insights** | `AIInsights.GenerateInsightsCommand` | `POST /insights/variance` | Requests AI-generated commentary and (optionally) writes results to the `Insights` sheet. |
| **Load Insights History** | `AIInsights.LoadInsightsHistoryCommand` | `GET /insights/history` | Fetches saved insight runs into `InsightsHistory`. |
| **Store API Key** | `AIInsights.StoreApiKeyCommand` | `POST /settings/api-key` | Stores or clears the encrypted API key on the bridge (requires admin token). |
| **Connection Settings** | `Settings.ShowSettingsDialog` | – | Presents the `frmSettings` dialog for backend configuration. |

## Suggested workbook scaffolding

When the macros run for the first time they automatically create the necessary worksheets (`Reports`, `Scenarios`, `Insights`, `InsightsHistory`) and structured tables (`ReportsTable`, `ScenariosTable`, `InsightsVarianceTable`, `InsightsHistoryTable`).  Existing tables are cleared and resized before new data is written, which keeps the workbook tidy between refreshes.

If you want to customise formatting further, adjust the sheets after the first refresh and Excel will retain your styling.
