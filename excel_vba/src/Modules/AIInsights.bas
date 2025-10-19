Attribute VB_Name = "AIInsights"
Option Explicit

Private Const INSIGHTS_SHEET As String = "Insights"
Private Const INSIGHTS_TABLE As String = "InsightsVarianceTable"
Private Const HISTORY_SHEET As String = "InsightsHistory"
Private Const HISTORY_TABLE As String = "InsightsHistoryTable"

Public Sub GenerateInsightsCommand()
    On Error GoTo HandleError

    Dim actualScenario As String
    actualScenario = InputBox("Actual scenario", "Generate AI Insights", GetConfigValue("InsightsActual", "Actuals"))
    If Trim$(actualScenario) = "" Then Exit Sub

    Dim budgetScenario As String
    budgetScenario = InputBox("Budget scenario", "Generate AI Insights", GetConfigValue("InsightsBudget", "Budget"))
    If Trim$(budgetScenario) = "" Then Exit Sub

    Dim prompt As String
    prompt = InputBox("Prompt (optional)", "Generate AI Insights", GetConfigValue("InsightsPrompt"))

    Dim includeRows As Boolean
    includeRows = (MsgBox("Include variance rows and narrative on the Insights worksheet?", vbYesNo + vbQuestion, "Generate AI Insights") = vbYes)

    Dim apiBase As String
    apiBase = InputBox("Custom API base (optional)", "Generate AI Insights", GetConfigValue("InsightsApiBase"))

    Dim model As String
    model = InputBox("Model (optional)", "Generate AI Insights", GetConfigValue("InsightsModel"))

    Dim mode As String
    mode = InputBox("Mode (chat-completions/responses, optional)", "Generate AI Insights", GetConfigValue("InsightsMode", "chat-completions"))

    Dim response As Object
    Set response = API_Client.GenerateInsights(actualScenario, budgetScenario, prompt, includeRows, apiBase, model, mode)

    SetConfigValue "InsightsActual", actualScenario
    SetConfigValue "InsightsBudget", budgetScenario
    SetConfigValue "InsightsPrompt", prompt
    SetConfigValue "InsightsApiBase", apiBase
    SetConfigValue "InsightsModel", model
    SetConfigValue "InsightsMode", mode

    Dim narrative As String
    If response.Exists("insights") Then
        narrative = CStr(response("insights"))
    End If

    WriteInsightsToWorksheet response, narrative, actualScenario, budgetScenario

    If narrative <> "" Then
        ShowInfo "Insights generated:" & vbCrLf & Left$(narrative, 400)
    Else
        ShowInfo "Insights request completed."
    End If
    Exit Sub

HandleError:
    ShowError "Generate AI Insights", Err.Description
End Sub

Public Sub StoreApiKeyCommand()
    On Error GoTo HandleError

    Dim apiKey As String
    apiKey = InputBox("Enter your OpenAI-compatible API key (leave blank to clear).", "Datarails AI Credentials")

    API_Client.StoreApiKey IIf(Trim$(apiKey) = "", vbNullString, apiKey)

    If Trim$(apiKey) = "" Then
        ShowInfo "Stored API key cleared on the backend."
    Else
        ShowInfo "API key securely stored on the backend."
    End If
    Exit Sub

HandleError:
    ShowError "Store AI API Key", Err.Description
End Sub

Public Sub LoadInsightsHistoryCommand()
    On Error GoTo HandleError

    Dim response As Object
    Set response = API_Client.FetchInsightsHistory()

    Dim items As Collection
    If response.Exists("items") Then
        Set items = response("items")
    Else
        Set items = New Collection
    End If

    Dim headers As Variant
    headers = Array("ID", "Actual", "Budget", "Prompt", "Row Count", "Created At")

    Dim values As Variant
    values = JsonToArray(items, Array("id", "actual", "budget", "prompt", "rowCount", "createdAt"))

    Dim table As ListObject
    Set table = GetOrCreateTable(HISTORY_SHEET, HISTORY_TABLE, headers)
    WriteTable table, headers, values

    Dim count As Long
    count = ArrayRowCount(values)
    Dim message As String
    message = "Loaded " & count & " insights history entries."
    If response.Exists("total") Then
        message = message & " Total stored: " & CStr(response("total"))
    End If
    ShowInfo message
    Exit Sub

HandleError:
    ShowError "Load Insights History", Err.Description
End Sub

Private Sub WriteInsightsToWorksheet(ByVal response As Object, ByVal narrative As String, ByVal actualScenario As String, ByVal budgetScenario As String)
    EnsureWorksheetExists INSIGHTS_SHEET
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(INSIGHTS_SHEET)

    ws.Range("A1").Value = "Narrative insights"
    ws.Range("A2").Value = narrative
    ws.Range("A2").WrapText = True
    ws.Range("A3").Value = "Actual: " & actualScenario & " vs Budget: " & budgetScenario

    Dim headers As Variant
    headers = Array("Period", "Department", "Account", "Actual", "Budget", "Variance")

    Dim rows As Variant
    If response.Exists("rows") Then
        Dim rowItems As Collection
        Set rowItems = response("rows")
        rows = JsonToArray(rowItems, Array("period", "department", "account", "actual", "budget", "variance"))
    End If

    Dim table As ListObject
    Set table = GetOrCreateTable(INSIGHTS_SHEET, INSIGHTS_TABLE, headers)
    WriteTable table, headers, rows
End Sub
