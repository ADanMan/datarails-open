Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Public Sub Ribbon_LoadData(control As IRibbonControl)
    DataLoader.LoadDataCommand
End Sub

Public Sub Ribbon_RefreshReports(control As IRibbonControl)
    ReportLoader.RefreshReportsCommand
End Sub

Public Sub Ribbon_ExportScenario(control As IRibbonControl)
    ScenarioExporter.ExportScenarioCommand
End Sub

Public Sub Ribbon_GenerateInsights(control As IRibbonControl)
    AIInsights.GenerateInsightsCommand
End Sub

Public Sub Ribbon_StoreApiKey(control As IRibbonControl)
    AIInsights.StoreApiKeyCommand
End Sub

Public Sub Ribbon_LoadInsightsHistory(control As IRibbonControl)
    AIInsights.LoadInsightsHistoryCommand
End Sub

Public Sub Ribbon_OpenSettings(control As IRibbonControl)
    Settings.ShowSettingsDialog
End Sub
