Attribute VB_Name = "ScenarioExporter"
Option Explicit

Private Const SCENARIO_SHEET As String = "Scenarios"
Private Const SCENARIO_TABLE As String = "ScenariosTable"

Public Sub ExportScenarioCommand()
    On Error GoTo HandleError

    Dim sourceScenario As String
    sourceScenario = InputBox("Source scenario", "Export Scenario", GetConfigValue("SourceScenario", "Actuals"))
    If sourceScenario = "" Then Exit Sub

    Dim targetScenario As String
    targetScenario = InputBox("Target scenario", "Export Scenario", GetConfigValue("TargetScenario", "Working"))
    If targetScenario = "" Then Exit Sub

    Dim percentageChangeText As String
    percentageChangeText = InputBox("Percentage change (e.g. 5 for 5%)", "Export Scenario", "0")
    If percentageChangeText = "" Then Exit Sub

    Dim percentageChange As Double
    percentageChange = CDbl(percentageChangeText)

    Dim department As String
    department = InputBox("Filter by department (optional)", "Export Scenario", GetConfigValue("ScenarioDepartment"))

    Dim account As String
    account = InputBox("Filter by account (optional)", "Export Scenario", GetConfigValue("ScenarioAccount"))

    Dim persistResult As VbMsgBoxResult
    persistResult = MsgBox("Persist scenario rows to the database?", vbYesNo + vbQuestion, "Export Scenario")
    Dim persist As Boolean
    persist = (persistResult = vbYes)

    Dim payload As Object
    Set payload = JsonConverter.ParseJson("{}")
    payload.Add "sourceScenario", sourceScenario
    payload.Add "targetScenario", targetScenario
    payload.Add "percentageChange", percentageChange
    If Trim(department) <> "" Then
        payload.Add "department", department
    End If
    If Trim(account) <> "" Then
        payload.Add "account", account
    End If
    payload.Add "persist", persist

    Dim response As Object
    Set response = API_Client.HttpPost("/scenarios/export", payload)

    Dim rows As Collection
    Set rows = response("rows")

    Dim headers As Variant
    headers = Array("Period", "Department", "Account", "Value", "Currency", "Metadata")

    Dim values As Variant
    values = JsonToArray(rows, Array("period", "department", "account", "value", "currency", "metadata"))

    Dim table As ListObject
    Set table = GetOrCreateTable(SCENARIO_SHEET, SCENARIO_TABLE, headers)
    WriteTable table, headers, values

    SetConfigValue "SourceScenario", sourceScenario
    SetConfigValue "TargetScenario", targetScenario
    SetConfigValue "ScenarioDepartment", department
    SetConfigValue "ScenarioAccount", account

    If response.Exists("message") Then
        ShowInfo CStr(response("message"))
    Else
        ShowInfo "Scenario exported to sheet '" & SCENARIO_SHEET & "'."
    End If
    Exit Sub

HandleError:
    ShowError "Export Scenario", Err.Description
End Sub
