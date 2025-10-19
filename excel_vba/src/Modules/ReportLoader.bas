Attribute VB_Name = "ReportLoader"
Option Explicit

Private Const REPORT_SHEET As String = "Reports"
Private Const REPORT_TABLE As String = "ReportsTable"

Public Sub RefreshReportsCommand()
    On Error GoTo HandleError

    Dim response As Object
    Set response = API_Client.RefreshReports()

    Dim scenarioName As String
    If response.Exists("scenario") Then
        scenarioName = CStr(response("scenario"))
    Else
        scenarioName = "(all scenarios)"
    End If

    Dim rows As Collection
    Set rows = response("rows")

    Dim headers As Variant
    headers = Array("Period", "Department", "Total")

    Dim values As Variant
    values = JsonToArray(rows, Array("period", "department", "total"))

    Dim table As ListObject
    Set table = GetOrCreateTable(REPORT_SHEET, REPORT_TABLE, headers)
    WriteTable table, headers, values

    ShowInfo "Report refreshed for " & scenarioName
    Exit Sub

HandleError:
    ShowError "Refresh Reports", Err.Description
End Sub
