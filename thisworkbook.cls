VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Option Explicit

Private gSettingsPane As Office.CustomTaskPane

Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    InitializeSettingsPane
CleanExit:
    Exit Sub
ErrHandler:
    LoggingModule.LogError "ThisWorkbook.Workbook_Open", Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrHandler
    If Not gSettingsPane Is Nothing Then
        Application.CustomTaskPanes.Remove gSettingsPane
        Set gSettingsPane = Nothing
    End If
CleanExit:
    Exit Sub
ErrHandler:
    LoggingModule.LogError "ThisWorkbook.Workbook_BeforeClose", Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub OnAnalyzeButton(control As IRibbonControl)
    On Error GoTo ErrHandler
    InitializeSettingsPane
    AnalyzeWorkbook
CleanExit:
    Exit Sub
ErrHandler:
    LoggingModule.LogError "ThisWorkbook.OnAnalyzeButton", Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Sub InitializeSettingsPane()
    On Error GoTo ErrHandler
    If gSettingsPane Is Nothing Then
        Dim paneControl As SettingsPane
        Set paneControl = New SettingsPane
        Set gSettingsPane = Application.CustomTaskPanes.Add(paneControl, "FlotaMaster Settings")
    End If
    gSettingsPane.Visible = True
CleanExit:
    Exit Sub
ErrHandler:
    LoggingModule.LogError "ThisWorkbook.InitializeSettingsPane", Err.Number, Err.Description
    Resume CleanExit
End Sub