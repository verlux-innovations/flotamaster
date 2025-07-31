Attribute VB_Name = "scoringengine"
Option Explicit

Public Function ComputeWeightedScores(data As Variant, weights As Variant) As Variant
    ' Computes weighted scores for each row in a two-dimensional data array.
    ' data: 2D array where column 1 = massPull, column 2 = cuGrade
    ' weights: 1D array where element 0 = weightMass, element 1 = weightCu
    Dim rowLow As Long, rowHigh As Long
    Dim colLow As Long, colHigh As Long
    Dim wLow As Long, wHigh As Long
    Dim i As Long
    Dim result() As Double
    
    ' Validate inputs
    If Not IsArray(data) Then Err.Raise vbObjectError + 9501, "ComputeWeightedScores", "Input 'data' is not an array."
    If Not IsArray(weights) Then Err.Raise vbObjectError + 9502, "ComputeWeightedScores", "Input 'weights' is not an array."
    
    ' Validate data dimensions
    On Error GoTo DimError
        rowLow = LBound(data, 1)
        rowHigh = UBound(data, 1)
        colLow = LBound(data, 2)
        colHigh = UBound(data, 2)
    On Error GoTo 0
    If colHigh < colLow + 1 Then Err.Raise vbObjectError + 9503, "ComputeWeightedScores", "Data array must have at least two columns."
    
    ' Validate weights length
    wLow = LBound(weights)
    wHigh = UBound(weights)
    If wHigh < wLow + 1 Then Err.Raise vbObjectError + 9504, "ComputeWeightedScores", "Weights array must contain at least two elements."
    
    ' Prepare result array
    ReDim result(rowLow To rowHigh)
    
    ' Compute weighted scores
    For i = rowLow To rowHigh
        If IsError(data(i, 1)) Or IsError(data(i, 2)) Then _
            Err.Raise vbObjectError + 9505, "ComputeWeightedScores", "Data contains an error at row " & i & "."
        result(i) = CalculateScore( _
            CDbl(data(i, 1)), _
            CDbl(data(i, 2)), _
            CDbl(weights(wLow)), _
            CDbl(weights(wLow + 1)) _
        )
    Next i
    
    ComputeWeightedScores = result
    Exit Function

DimError:
    Err.Raise vbObjectError + 9506, "ComputeWeightedScores", "Data must be a two-dimensional array."
End Function

Public Function CalculateScore( _
    massPull As Double, _
    cuGrade As Double, _
    weightMass As Double, _
    weightCu As Double _
) As Double
    ' Calculates a linear weighted score: massPull*weightMass + cuGrade*weightCu
    CalculateScore = massPull * weightMass + cuGrade * weightCu
End Function

Public Function NormalizeScores(scores As Variant) As Variant
    ' Normalizes an array of numeric scores to the range [0, 1].
    Dim low As Long, high As Long
    Dim minVal As Double, maxVal As Double
    Dim norm() As Double
    Dim i As Long
    
    ' Validate input
    If Not IsArray(scores) Then Err.Raise vbObjectError + 9511, "NormalizeScores", "Input 'scores' is not an array."
    low = LBound(scores)
    high = UBound(scores)
    If high < low Then Err.Raise vbObjectError + 9512, "NormalizeScores", "Scores array is empty or has invalid bounds."
    
    ' Initialize
    ReDim norm(low To high)
    minVal = scores(low)
    maxVal = scores(low)
    
    ' Find min and max
    For i = low + 1 To high
        If scores(i) < minVal Then minVal = scores(i)
        If scores(i) > maxVal Then maxVal = scores(i)
    Next i
    
    ' Normalize
    If maxVal = minVal Then
        For i = low To high
            norm(i) = 0
        Next i
    Else
        For i = low To high
            norm(i) = (scores(i) - minVal) / (maxVal - minVal)
        Next i
    End If
    
    NormalizeScores = norm
End Function