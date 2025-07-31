Attribute VB_Name = "kineticsmodule"
Option Explicit

Private g_k As Double
Private g_intercept As Double
Private g_RSq As Double
Private g_modelFitted As Boolean

Public Function AnalyzeKinetics(timeData As Variant, concData As Variant) As Variant
    Dim tArr As Variant, cArr As Variant
    Dim dataArr As Variant
    Dim n As Long, i As Long

    Call ResetModel

    On Error GoTo ErrFlatten
    tArr = FlattenTo1D(timeData)
    cArr = FlattenTo1D(concData)
    On Error GoTo 0

    n = UBound(tArr) - LBound(tArr) + 1
    If n <> UBound(cArr) - LBound(cArr) + 1 Then
        AnalyzeKinetics = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim dataArr(1 To n, 1 To 2)
    For i = 1 To n
        dataArr(i, 1) = CDbl(tArr(LBound(tArr) + i - 1))
        dataArr(i, 2) = CDbl(cArr(LBound(cArr) + i - 1))
    Next i

    AnalyzeKinetics = FitKineticsModel(dataArr)
    Exit Function

ErrFlatten:
    AnalyzeKinetics = CVErr(xlErrValue)
End Function

Public Function FitKineticsModel(data As Variant) As Variant
    Dim n As Long, i As Long
    Dim X() As Double, Y() As Double
    Dim sumX As Double, sumY As Double, sumXY As Double, sumXX As Double
    Dim slope As Double, intercept As Double, denom As Double
    Dim meanY As Double, ssTot As Double, ssRes As Double

    n = UBound(data, 1)
    If UBound(data, 2) < 2 Then
        FitKineticsModel = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim X(1 To n)
    ReDim Y(1 To n)
    For i = 1 To n
        If data(i, 2) <= 0 Then
            FitKineticsModel = CVErr(xlErrNum)
            Exit Function
        End If
        X(i) = data(i, 1)
        Y(i) = Log(data(i, 2))
        sumX = sumX + X(i)
        sumY = sumY + Y(i)
    Next i

    meanY = sumY / n
    For i = 1 To n
        sumXY = sumXY + X(i) * Y(i)
        sumXX = sumXX + X(i) * X(i)
    Next i

    denom = n * sumXX - sumX * sumX
    If denom = 0 Then
        FitKineticsModel = CVErr(xlErrDiv0)
        Exit Function
    End If

    slope = (n * sumXY - sumX * sumY) / denom
    intercept = (sumY - slope * sumX) / n

    g_k = -slope
    g_intercept = intercept

    For i = 1 To n
        ssTot = ssTot + (Y(i) - meanY) ^ 2
        ssRes = ssRes + (Y(i) - (slope * X(i) + intercept)) ^ 2
    Next i

    If ssTot <> 0 Then
        g_RSq = 1 - ssRes / ssTot
    Else
        g_RSq = 0
    End If

    g_modelFitted = True

    Dim result(1 To 3) As Double
    result(1) = g_k
    result(2) = g_intercept
    result(3) = g_RSq
    FitKineticsModel = result
End Function

Public Function CalculateRateConstant(time As Double, conc As Double) As Variant
    If time <= 0 Or conc <= 0 Or Not g_modelFitted Then
        CalculateRateConstant = CVErr(xlErrNum)
        Exit Function
    End If
    CalculateRateConstant = (1 / time) * Log(Exp(g_intercept) / conc)
End Function

'-------------------- Private helpers --------------------

Private Sub ResetModel()
    g_k = 0
    g_intercept = 0
    g_RSq = 0
    g_modelFitted = False
End Sub

Private Function GetArrayDimCount(arr As Variant) As Integer
    Dim dimCount As Integer, tmp As Long
    On Error GoTo Done
    Do
        dimCount = dimCount + 1
        tmp = UBound(arr, dimCount)
    Loop
Done:
    GetArrayDimCount = dimCount - 1
    On Error GoTo 0
End Function

Private Function FlattenTo1D(arr As Variant) As Variant
    Dim d As Integer
    d = GetArrayDimCount(arr)

    If d = 1 Then
        FlattenTo1D = arr
    ElseIf d = 2 Then
        Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
        lb1 = LBound(arr, 1): ub1 = UBound(arr, 1)
        lb2 = LBound(arr, 2): ub2 = UBound(arr, 2)

        If lb1 = ub1 Then
            Dim tmp1() As Variant, j As Long
            ReDim tmp1(lb2 To ub2)
            For j = lb2 To ub2
                tmp1(j) = arr(lb1, j)
            Next j
            FlattenTo1D = tmp1

        ElseIf lb2 = ub2 Then
            Dim tmp2() As Variant, i As Long
            ReDim tmp2(lb1 To ub1)
            For i = lb1 To ub1
                tmp2(i) = arr(i, lb2)
            Next i
            FlattenTo1D = tmp2

        Else
            Err.Raise vbObjectError + 1, , "Cannot flatten 2D array with multiple rows and columns"
        End If

    Else
        Err.Raise vbObjectError + 1, , "Unsupported array dimensions"
    End If
End Function