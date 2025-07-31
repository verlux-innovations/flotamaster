Option Explicit

Public Sub ClearDashboard()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Dashboard").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Public Sub CreateDashboardSheet()
    Dim ws As Worksheet, cd As Worksheet, btn As Button
    ' Delete existing Dashboard
    ClearDashboard
    ' Add Dashboard sheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Dashboard"
    ' Add Generate Comments button dynamically below existing content
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    Set btn = ws.Buttons.Add(ws.Cells(lastRow + 2, "B").Left, ws.Cells(lastRow + 2, "B").Top, 100, 20)
    btn.Caption = "Generate Comments"
    btn.OnAction = "'" & ThisWorkbook.Name & "'!GenerateComments"
    ' Delete existing ChartData sheet if any
    On Error Resume Next
    ThisWorkbook.Sheets("ChartData").Delete
    On Error GoTo 0
    ' Add hidden ChartData sheet
    Set cd = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    cd.Name = "ChartData"
    cd.Visible = xlSheetHidden
End Sub

Private Sub GenerateChart(data As Variant, startCell As String, chtType As XlChartType, chartTitle As String)
    Dim cd As Worksheet, ws As Worksheet
    Set cd = ThisWorkbook.Sheets("ChartData")
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Dim rows As Long, cols As Long
    rows = UBound(data, 1)
    cols = UBound(data, 2)
    ' Clear previous data
    cd.Cells.Clear
    ' Write data to hidden sheet
    cd.Range("A1").Resize(rows, cols).Value = data
    ' Add chart
    Dim chartObj As ChartObject, ch As Chart
    Set chartObj = ws.ChartObjects.Add(ws.Range(startCell).Left, ws.Range(startCell).Top, 400, 300)
    Set ch = chartObj.Chart
    ch.ChartType = chtType
    ' X values
    Dim rngX As Range
    Set rngX = cd.Range(cd.Cells(2, 1), cd.Cells(rows, 1))
    ' Add series for each data column
    Dim i As Long
    For i = 2 To cols
        With ch.SeriesCollection.NewSeries
            .Name = cd.Cells(1, i).Value
            .XValues = rngX
            .Values = cd.Range(cd.Cells(2, i), cd.Cells(rows, i))
        End With
    Next i
    ' Format chart
    ch.HasTitle = True
    ch.ChartTitle.Text = chartTitle
    ch.Legend.Position = xlLegendPositionBottom
End Sub

Public Sub GeneratePerformanceChart(data As Variant, startCell As String)
    GenerateChart data, startCell, xlLineMarkers, "Performance"
End Sub

Public Sub GenerateKineticsChart(data As Variant, startCell As String)
    GenerateChart data, startCell, xlXYScatterSmooth, "Kinetics"
End Sub

Public Sub GenerateComments()
    On Error GoTo ErrHandler
    ' Ensure Dashboard exists
    If Not SheetExists("Dashboard") Then CreateDashboardSheet
    ' Call the report generator's comment routine
    Application.Run "'" & ThisWorkbook.Name & "'!ReportGenerator.GenerateComments"
    Exit Sub
ErrHandler:
    MsgBox "Error generating comments: " & Err.Description, vbExclamation
End Sub

Private Function SheetExists(sName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ThisWorkbook.Sheets(sName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function