Option Explicit

Public Settings As Scripting.Dictionary

Public Sub LoadSettings()
    Dim ws As Worksheet
    If Settings Is Nothing Then Set Settings = New Scripting.Dictionary

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0

    If ws Is Nothing Then
        Dim newWS As Worksheet
        Set newWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error GoTo NameError
        newWS.Name = "Settings"
        On Error GoTo 0
        newWS.Visible = xlSheetVeryHidden
        Set ws = newWS
    End If

    Settings.RemoveAll
    Dim lastRow As Long, i As Long, key As String
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 1 To lastRow
        key = CStr(ws.Cells(i, "A").Value)
        If Len(key) > 0 Then
            Settings(key) = ws.Cells(i, "B").Value
        End If
    Next i
    Exit Sub

NameError:
    Err.Clear
    newWS.Delete
    MsgBox "Error: A worksheet named 'Settings' already exists or could not be renamed.", vbCritical
    End
End Sub

Public Sub SaveSettings()
    Dim ws As Worksheet
    If Settings Is Nothing Then Set Settings = New Scripting.Dictionary

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Settings"
        ws.Visible = xlSheetVeryHidden
    End If

    ws.Cells.ClearContents

    Dim i As Long: i = 1
    Dim key As Variant
    For Each key In Settings.Keys
        ws.Cells(i, "A").Value = key
        ws.Cells(i, "B").Value = Settings(key)
        i = i + 1
    Next key

    ws.Visible = xlSheetVeryHidden
End Sub

Public Sub ShowSettingsPane()
    LoadSettings
    frmSettings.LoadSettings Settings
    frmSettings.Show vbModal
End Sub

Public Sub OnSettingChanged(key As String, value As Variant)
    If Settings Is Nothing Then LoadSettings
    If Settings.Exists(key) Then
        Settings(key) = value
    Else
        Settings.Add key, value
    End If
    SaveSettings
End Sub

Public Function GetSetting(key As String, Optional defaultValue As Variant) As Variant
    If Settings Is Nothing Then LoadSettings
    If Settings.Exists(key) Then
        GetSetting = Settings(key)
    Else
        GetSetting = defaultValue
    End If
End Function