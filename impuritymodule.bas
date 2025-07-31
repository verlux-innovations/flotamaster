Attribute VB_Name = "ImpurityModule"
Option Explicit

Public Function EvaluatePyriteLevel(pyriteValue As Double, threshold As Double) As String
    If pyriteValue >= threshold Then
        EvaluatePyriteLevel = "High Pyrite"
    Else
        EvaluatePyriteLevel = "Low Pyrite"
    End If
End Function

Public Function EvaluateCarbonLevel(carbonValue As Double, threshold As Double) As String
    If carbonValue >= threshold Then
        EvaluateCarbonLevel = "High Carbon"
    Else
        EvaluateCarbonLevel = "Low Carbon"
    End If
End Function

Public Function AssessImpurities( _
    massPullData As Variant, _
    pyriteData As Variant, _
    carbonData As Variant, _
    Optional pyriteThreshold As Double = 1#, _
    Optional carbonThreshold As Double = 0.5 _
) As Variant

    Dim i As Long
    Dim totalMassPull As Double: totalMassPull = 0#
    Dim totalWeightedPyrite As Double: totalWeightedPyrite = 0#
    Dim totalWeightedCarbon As Double: totalWeightedCarbon = 0#
    Dim lb1 As Long, ub1 As Long
    Dim lb2 As Long, ub2 As Long
    Dim lb3 As Long, ub3 As Long
    Dim lb As Long, ub As Long
    Dim avgPyrite As Double, avgCarbon As Double
    Dim pyriteAssessment As String, carbonAssessment As String
    Dim result(1 To 2) As Variant
    
    ' Validate inputs
    If Not IsArray(massPullData) Then Err.Raise vbObjectError + 1000, "AssessImpurities", "massPullData must be an array"
    If Not IsArray(pyriteData) Then Err.Raise vbObjectError + 1001, "AssessImpurities", "pyriteData must be an array"
    If Not IsArray(carbonData) Then Err.Raise vbObjectError + 1002, "AssessImpurities", "carbonData must be an array"
    
    lb1 = LBound(massPullData): ub1 = UBound(massPullData)
    lb2 = LBound(pyriteData):   ub2 = UBound(pyriteData)
    lb3 = LBound(carbonData):   ub3 = UBound(carbonData)
    
    If ub1 < lb1 Then Err.Raise vbObjectError + 1003, "AssessImpurities", "massPullData is empty"
    If ub2 < lb2 Then Err.Raise vbObjectError + 1004, "AssessImpurities", "pyriteData is empty"
    If ub3 < lb3 Then Err.Raise vbObjectError + 1005, "AssessImpurities", "carbonData is empty"
    
    ' Determine common bounds to avoid out-of-bounds
    lb = lb1
    If lb2 > lb Then lb = lb2
    If lb3 > lb Then lb = lb3
    
    ub = ub1
    If ub2 < ub Then ub = ub2
    If ub3 < ub Then ub = ub3
    
    If ub < lb Then Err.Raise vbObjectError + 1006, "AssessImpurities", "No overlapping elements among inputs"
    
    ' Compute totals
    For i = lb To ub
        If IsNumeric(massPullData(i)) Then
            totalMassPull = totalMassPull + CDbl(massPullData(i))
            If IsNumeric(pyriteData(i)) Then
                totalWeightedPyrite = totalWeightedPyrite + CDbl(massPullData(i)) * CDbl(pyriteData(i))
            End If
            If IsNumeric(carbonData(i)) Then
                totalWeightedCarbon = totalWeightedCarbon + CDbl(massPullData(i)) * CDbl(carbonData(i))
            End If
        End If
    Next i
    
    If totalMassPull > 0# Then
        avgPyrite = totalWeightedPyrite / totalMassPull
        avgCarbon = totalWeightedCarbon / totalMassPull
    Else
        avgPyrite = 0#
        avgCarbon = 0#
    End If
    
    pyriteAssessment = EvaluatePyriteLevel(avgPyrite, pyriteThreshold)
    carbonAssessment = EvaluateCarbonLevel(avgCarbon, carbonThreshold)
    
    result(1) = pyriteAssessment
    result(2) = carbonAssessment
    AssessImpurities = result

End Function