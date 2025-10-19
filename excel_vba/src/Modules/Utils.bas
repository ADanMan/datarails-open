Attribute VB_Name = "Utils"
Option Explicit

Public Function RequireBackendUrl() As String
    Dim url As String
    url = Trim(GetConfigValue("BackendUrl"))

    If url = "" Then
        Err.Raise vbObjectError + 601, "Utils.RequireBackendUrl", "Backend URL is not configured. Use the Settings button to provide it."
    End If

    RequireBackendUrl = url
End Function

Public Function RequireBridgeToken() As String
    Dim token As String
    token = Trim(GetConfigValue("BridgeToken"))

    If token = "" Then
        Err.Raise vbObjectError + 602, "Utils.RequireBridgeToken", "Bridge token is not configured. Use the Settings button to provide it."
    End If

    RequireBridgeToken = token
End Function

Public Sub ShowError(ByVal source As String, ByVal message As String)
    MsgBox message, vbCritical + vbOKOnly, "Datarails • " & source
End Sub

Public Sub ShowInfo(ByVal message As String)
    MsgBox message, vbInformation + vbOKOnly, "Datarails"
End Sub

Public Sub WriteTable(ByVal target As ListObject, ByVal headers As Variant, ByVal rows As Variant)
    Dim ws As Worksheet
    Set ws = target.Parent

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanUp

    If Not target.DataBodyRange Is Nothing Then
        target.DataBodyRange.ClearContents
    End If

    target.HeaderRowRange.Value = ToRow(headers)

    Dim columnCount As Long
    columnCount = UBound(headers) - LBound(headers) + 1

    If IsArray(rows) Then
        Dim rowCount As Long
        On Error Resume Next
        rowCount = UBound(rows, 1) - LBound(rows, 1) + 1
        On Error GoTo 0

        Dim newRange As Range
        Set newRange = target.Range.Resize(Application.Max(rowCount, 1) + 1, columnCount)
        target.Resize newRange

        If rowCount > 0 Then
            target.DataBodyRange.Value = rows
        End If
    Else
        Dim defaultRange As Range
        Set defaultRange = target.Range.Resize(2, columnCount)
        target.Resize defaultRange
    End If

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Function JsonToArray(ByVal jsonCollection As Collection, ByVal fieldNames As Variant) As Variant
    Dim rowCount As Long
    rowCount = jsonCollection.Count

    If rowCount = 0 Then
        JsonToArray = Array()
        Exit Function
    End If

    Dim columnCount As Long
    columnCount = UBound(fieldNames) - LBound(fieldNames) + 1

    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To columnCount)

    Dim rowIndex As Long
    For rowIndex = 1 To rowCount
        Dim item As Object
        Set item = jsonCollection(rowIndex)

        Dim columnIndex As Long
        For columnIndex = 1 To columnCount
            Dim fieldName As String
            fieldName = fieldNames(columnIndex - 1)

            If item.Exists(fieldName) Then
                result(rowIndex, columnIndex) = item(fieldName)
            Else
                result(rowIndex, columnIndex) = ""
            End If
        Next columnIndex
    Next rowIndex

    JsonToArray = result
End Function

Public Sub EnsureWorksheetExists(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
End Sub

Public Function GetOrCreateTable(ByVal sheetName As String, ByVal tableName As String, ByVal headers As Variant) As ListObject
    EnsureWorksheetExists sheetName
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        Dim headerRange As Range
        Dim columnCount As Long
        columnCount = UBound(headers) - LBound(headers) + 1

        Set headerRange = ws.Range("A1").Resize(1, columnCount)
        headerRange.Value = ToRow(headers)
        Set tbl = ws.ListObjects.Add(xlSrcRange, headerRange, , xlYes)
        tbl.Name = tableName
    End If

    Set GetOrCreateTable = tbl
End Function

Public Function JsonToString(ByVal jsonObject As Object) As String
    JsonToString = JsonConverter.ConvertToJson(jsonObject)
End Function

Private Function ToRow(ByVal headers As Variant) As Variant
    Dim columnCount As Long
    columnCount = UBound(headers) - LBound(headers) + 1

    Dim result() As Variant
    ReDim result(1 To 1, 1 To columnCount)

    Dim index As Long
    For index = 1 To columnCount
        result(1, index) = headers(LBound(headers) + index - 1)
    Next index

    ToRow = result
End Function

Public Function ArrayRowCount(ByVal rows As Variant) As Long
    If IsArray(rows) Then
        On Error Resume Next
        ArrayRowCount = UBound(rows, 1) - LBound(rows, 1) + 1
        If Err.Number <> 0 Then ArrayRowCount = 0
        On Error GoTo 0
    Else
        ArrayRowCount = 0
    End If
End Function
