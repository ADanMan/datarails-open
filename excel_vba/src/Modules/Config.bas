Attribute VB_Name = "Config"
Option Explicit

Private Const CONFIG_SHEET_NAME As String = "_DatarailsConfig"

Public Sub EnsureConfigSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET_NAME
        ws.Visible = xlSheetVeryHidden
        ws.Cells(1, 1).Value = "Key"
        ws.Cells(1, 2).Value = "Value"
    End If
End Sub

Public Function GetConfigValue(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    EnsureConfigSheet

    Dim rng As Range
    Dim lastRow As Long
    lastRow = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME).Cells(ThisWorkbook.Worksheets(CONFIG_SHEET_NAME).Rows.Count, 1).End(xlUp).Row

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If ThisWorkbook.Worksheets(CONFIG_SHEET_NAME).Cells(rowIndex, 1).Value = key Then
            GetConfigValue = NzString(ThisWorkbook.Worksheets(CONFIG_SHEET_NAME).Cells(rowIndex, 2).Value, defaultValue)
            Exit Function
        End If
    Next rowIndex

    GetConfigValue = defaultValue
End Function

Public Sub SetConfigValue(ByVal key As String, ByVal value As String)
    EnsureConfigSheet

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If ws.Cells(rowIndex, 1).Value = key Then
            ws.Cells(rowIndex, 2).Value = value
            Exit Sub
        End If
    Next rowIndex

    ws.Cells(lastRow + 1, 1).Value = key
    ws.Cells(lastRow + 1, 2).Value = value
End Sub

Public Sub RemoveConfigValue(ByVal key As String)
    EnsureConfigSheet

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If ws.Cells(rowIndex, 1).Value = key Then
            ws.Rows(rowIndex).Delete
            Exit Sub
        End If
    Next rowIndex
End Sub

Private Function NzString(ByVal candidate As Variant, ByVal defaultValue As String) As String
    If IsError(candidate) Then
        NzString = defaultValue
    ElseIf IsNull(candidate) Then
        NzString = defaultValue
    ElseIf Trim(CStr(candidate)) = "" Then
        NzString = defaultValue
    Else
        NzString = CStr(candidate)
    End If
End Function
