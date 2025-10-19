Attribute VB_Name = "API_Client"
Option Explicit

Private Const DEFAULT_TIMEOUT As Long = 30000

Public Function BuildUrl(ByVal endpoint As String) As String
    Dim baseUrl As String
    baseUrl = RequireBackendUrl()

    If Right$(baseUrl, 1) = "/" And Left$(endpoint, 1) = "/" Then
        BuildUrl = baseUrl & Mid$(endpoint, 2)
    ElseIf Right$(baseUrl, 1) <> "/" And Left$(endpoint, 1) <> "/" Then
        BuildUrl = baseUrl & "/" & endpoint
    Else
        BuildUrl = baseUrl & endpoint
    End If
End Function

Public Function HttpGet(ByVal endpoint As String) As Object
    Dim responseText As String
    responseText = ExecuteRequest("GET", endpoint, "")
    Set HttpGet = JsonConverter.ParseJson(responseText)
End Function

Public Function HttpPost(ByVal endpoint As String, ByVal payload As Object) As Object
    Dim responseText As String
    Dim body As String
    body = JsonConverter.ConvertToJson(payload)

    responseText = ExecuteRequest("POST", endpoint, body)
    Set HttpPost = JsonConverter.ParseJson(responseText)
End Function

Public Function ExecuteRequest(ByVal method As String, ByVal endpoint As String, ByVal body As String) As String
    Dim http As Object
    Set http = CreateObject("WinHTTP.WinHTTPRequest.5.1")

    Dim url As String
    url = BuildUrl(endpoint)

    http.SetTimeouts DEFAULT_TIMEOUT, DEFAULT_TIMEOUT, DEFAULT_TIMEOUT, DEFAULT_TIMEOUT
    http.Open method, url, False
    http.SetRequestHeader "Content-Type", "application/json"

    Dim token As String
    token = Trim(GetConfigValue("BridgeToken"))
    If token <> "" Then
        http.SetRequestHeader "Authorization", "Bearer " & token
    End If

    On Error GoTo RequestError
    If method = "GET" Then
        http.Send
    Else
        http.Send body
    End If
    On Error GoTo 0

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 701, "API_Client.ExecuteRequest", "Request to " & url & " failed with status " & http.Status & ": " & http.ResponseText
    End If

    ExecuteRequest = http.ResponseText
    Exit Function

RequestError:
    Err.Raise vbObjectError + 702, "API_Client.ExecuteRequest", "Unable to reach backend: " & Err.Description
End Function

Public Function LoadData(
    ByVal filePath As String,
    ByVal source As String,
    ByVal scenario As String,
    Optional ByVal sheets As Variant,
    Optional ByVal tables As Variant
) As Object
    Dim payload As Object
    Set payload = JsonConverter.ParseJson("{}")
    payload.Add "path", filePath
    payload.Add "source", source
    payload.Add "scenario", scenario

    If Not IsMissing(sheets) Then
        Dim sheetList As Collection
        Set sheetList = ToCollection(sheets)
        If sheetList.Count > 0 Then
            payload.Add "sheets", sheetList
        End If
    End If

    If Not IsMissing(tables) Then
        Dim tableList As Collection
        Set tableList = ToCollection(tables)
        If tableList.Count > 0 Then
            payload.Add "tables", tableList
        End If
    End If

    Set LoadData = HttpPost("/load-data", payload)
End Function

Public Function RefreshReports() As Object
    Set RefreshReports = HttpGet("/reports/summary")
End Function

Public Function ExportScenario(ByVal scenarioId As String, ByVal persist As Boolean, ByVal worksheetName As String) As Object
    Dim payload As Object
    Set payload = JsonConverter.ParseJson("{}")
    payload.Add "scenarioId", scenarioId
    payload.Add "persist", persist
    payload.Add "worksheet", worksheetName

    Set ExportScenario = HttpPost("/scenarios/export", payload)
End Function

Public Function GenerateInsights(
    ByVal actualScenario As String,
    ByVal budgetScenario As String,
    ByVal prompt As String,
    ByVal includeRows As Boolean,
    Optional ByVal apiBase As String = "",
    Optional ByVal model As String = "",
    Optional ByVal mode As String = ""
) As Object
    Dim payload As Object
    Set payload = JsonConverter.ParseJson("{}")
    payload.Add "actualScenario", actualScenario
    payload.Add "budgetScenario", budgetScenario
    payload.Add "includeRows", includeRows

    If Trim(prompt) <> "" Then
        payload.Add "prompt", prompt
    End If

    Dim apiConfig As Object
    Set apiConfig = Nothing

    If Trim(apiBase) <> "" Or Trim(model) <> "" Or Trim(mode) <> "" Then
        Set apiConfig = JsonConverter.ParseJson("{}")
        If Trim(apiBase) <> "" Then apiConfig.Add "apiBase", apiBase
        If Trim(model) <> "" Then apiConfig.Add "model", model
        If Trim(mode) <> "" Then apiConfig.Add "mode", mode
    End If

    If Not apiConfig Is Nothing Then
        payload.Add "api", apiConfig
    End If

    Set GenerateInsights = HttpPost("/insights/variance", payload)
End Function

Public Function FetchInsightsHistory() As Object
    Set FetchInsightsHistory = HttpGet("/insights/history")
End Function

Public Sub StoreApiKey(ByVal apiKey As String)
    Dim payload As Object
    Set payload = JsonConverter.ParseJson("{}")
    If Trim$(apiKey) = "" Then
        payload.Add "apiKey", Null
    Else
        payload.Add "apiKey", apiKey
    End If

    Call HttpPost("/settings/api-key", payload)
End Sub

Public Function ListScenarios() As Object
    Set ListScenarios = HttpGet("/scenarios/list")
End Function

Private Function ToCollection(ByVal items As Variant) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim index As Long
    If IsArray(items) Then
        For index = LBound(items) To UBound(items)
            If Trim(CStr(items(index))) <> "" Then
                result.Add CStr(items(index))
            End If
        Next index
    ElseIf TypeName(items) = "Collection" Then
        Dim item As Variant
        For Each item In items
            If Trim(CStr(item)) <> "" Then
                result.Add CStr(item)
            End If
        Next item
    Else
        If Trim(CStr(items)) <> "" Then
            result.Add CStr(items)
        End If
    End If

    Set ToCollection = result
End Function
