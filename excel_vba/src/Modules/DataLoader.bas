Attribute VB_Name = "DataLoader"
Option Explicit

Public Sub LoadDataCommand()
    On Error GoTo HandleError

    Dim filePath As Variant
    filePath = Application.GetOpenFilename("CSV or Excel (*.csv;*.xlsx),*.csv;*.xlsx", , "Select data file")
    If VarType(filePath) = vbBoolean And filePath = False Then Exit Sub

    Dim source As String
    source = InputBox("Logical source name", "Load Data", GetConfigValue("LoadSource", "imports"))
    If source = "" Then Exit Sub

    Dim scenario As String
    scenario = InputBox("Scenario", "Load Data", GetConfigValue("LoadScenario", "Actuals"))
    If scenario = "" Then Exit Sub

    Dim sheetsInput As String
    sheetsInput = InputBox("Worksheet names (comma-separated, optional)", "Load Data", GetConfigValue("LoadSheets"))

    Dim tablesInput As String
    tablesInput = InputBox("Table names (comma-separated, optional)", "Load Data", GetConfigValue("LoadTables"))

    Dim sheets As Variant
    sheets = SplitList(sheetsInput)

    Dim tables As Variant
    tables = SplitList(tablesInput)

    Dim response As Object
    If IsEmpty(sheets) And IsEmpty(tables) Then
        Set response = API_Client.LoadData(CStr(filePath), source, scenario)
    ElseIf IsEmpty(tables) Then
        Set response = API_Client.LoadData(CStr(filePath), source, scenario, sheets)
    Else
        Set response = API_Client.LoadData(CStr(filePath), source, scenario, sheets, tables)
    End If

    SetConfigValue "LoadSource", source
    SetConfigValue "LoadScenario", scenario
    SetConfigValue "LoadSheets", sheetsInput
    SetConfigValue "LoadTables", tablesInput

    Dim message As String
    If response.Exists("message") Then
        message = CStr(response("message"))
    Else
        message = "Loaded data into scenario '" & scenario & "'."
    End If
    ShowInfo message
    Exit Sub

HandleError:
    ShowError "Load Data", Err.Description
End Sub

Private Function SplitList(ByVal text As String) As Variant
    If Trim$(text) = "" Then
        SplitList = Empty
        Exit Function
    End If

    Dim parts() As String
    parts = Split(text, ",")

    Dim cleaned() As String
    ReDim cleaned(LBound(parts) To UBound(parts))

    Dim index As Long
    Dim count As Long
    count = -1
    For index = LBound(parts) To UBound(parts)
        Dim value As String
        value = Trim$(parts(index))
        If value <> "" Then
            count = count + 1
            cleaned(count) = value
        End If
    Next index

    If count = -1 Then
        SplitList = Empty
    Else
        ReDim Preserve cleaned(0 To count)
        SplitList = cleaned
    End If
End Function
