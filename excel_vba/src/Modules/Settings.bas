Attribute VB_Name = "Settings"
Option Explicit

Public Sub ShowSettingsDialog()
    EnsureConfigSheet

    With frmSettings
        .txtBackendUrl.Text = GetConfigValue("BackendUrl", "http://localhost:8000")
        .txtBridgeToken.Text = GetConfigValue("BridgeToken")
        .Show
    End With
End Sub

Public Sub PersistSettings(ByVal backendUrl As String, ByVal bridgeToken As String)
    If Trim$(backendUrl) = "" Then
        Err.Raise vbObjectError + 801, "Settings.PersistSettings", "Backend URL cannot be empty."
    End If

    SetConfigValue "BackendUrl", backendUrl

    If Trim$(bridgeToken) = "" Then
        RemoveConfigValue "BridgeToken"
    Else
        SetConfigValue "BridgeToken", bridgeToken
    End If

    ShowInfo "Settings saved."
End Sub
