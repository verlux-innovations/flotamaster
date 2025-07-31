Option Explicit

Private Const DATA_SHEET_NAME As String = "DataEntry"

Private Function GetCollectorCells() As Variant
    GetCollectorCells = Array("C14", "C26", "C38", "C50")
End Function

Private Function GetCollectorNames() As Variant
    Dim cells As Variant
    cells = GetCollectorCells()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    Dim names() As String
    ReDim names(LBound(cells) To UBound(cells))
    Dim i As Long
    For i = LBound(cells) To UBound(cells)
        names(i) = CStr(ws.Range(cells(i)).Value2)
    Next i
    GetCollectorNames = names
End Function

Private Function ValidateCollectorNames(ByRef missingCell As String) As Boolean
    Dim names As Variant
    names = GetCollectorNames()
    Dim cells As Variant
    cells = GetCollectorCells()
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Trim(names(i)) = "" Then
            missingCell = cells(i)
            ValidateCollectorNames = False
            Exit Function
        End If
    Next i
    ValidateCollectorNames = True
End Function

Public Function ValidateInputs() As Boolean
    Dim missing As String
    If Not ValidateCollectorNames(missing) Then
        LogError "Input validation failed: missing collector name in cell " & missing
        MsgBox "Collector name missing in cell " & missing & ". Please enter all collector names.", vbExclamation, "Input Validation"
        ValidateInputs = False
    Else
        ValidateInputs = True
    End If
End Function

Public Function PreprocessData() As Variant
    If Not ValidateInputs Then
        PreprocessData = CVErr(xlErrValue)
        Exit Function
    End If
    PreprocessData = GetCollectorData(GetCollectorNames())
End Function

Private Function GetStartRows() As Variant
    GetStartRows = Array(18, 30, 42, 54)
End Function

Private Function GetCategories() As Variant
    GetCategories = Array("V", "W", "AH", "Z", "AB", "AJ", "AE")
End Function

Public Function GetCollectorData(collectorNames As Variant) As Variant
    Dim startRows As Variant
    startRows = GetStartRows()
    Dim categories As Variant
    categories = GetCategories()
    Dim numCollectors As Long, numTypes As Long, numPoints As Long
    numCollectors = UBound(collectorNames) - LBound(collectorNames) + 1
    numTypes = UBound(categories) - LBound(categories) + 1
    numPoints = 4
    Dim output As Variant
    ReDim output(1 To numCollectors, 1 To numTypes, 1 To numPoints)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    Dim i As Long, j As Long, k As Long
    For i = 1 To numCollectors
        Dim rowStart As Long
        rowStart = startRows(i - 1)
        For j = 1 To numTypes
            Dim col As String
            col = categories(j - 1)
            Dim rng As Range
            Set rng = ws.Range(col & rowStart & ":" & col & (rowStart + numPoints - 1))
            Dim block As Variant
            block = rng.Value2
            For k = 1 To numPoints
                output(i, j, k) = block(k, 1)
            Next k
        Next j
    Next i
    GetCollectorData = output
End Function