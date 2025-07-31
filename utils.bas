Attribute VB_Name = "utils"
Option Explicit

Public Sub LogMessage(level As String, message As String)
    Dim ts As String
    ts = FormatDate(Now)
    Debug.Print "[" & ts & "] [" & UCase$(level) & "] " & message
End Sub

Public Function FormatDate(dateValue As Date) As String
    FormatDate = Format$(dateValue, "yyyy-mm-dd hh:nn:ss")
End Function

Public Function SanitizeString(input As String) As String
    Dim result As String
    Dim re As Object
    result = input
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    ' Remove non-printable ASCII characters (0-31 and 127)
    re.Pattern = "[\x00-\x1F\x7F]"
    result = re.Replace(result, "")
    ' Normalize whitespace: replace one or more whitespace chars with a single space
    re.Pattern = "\s+"
    result = Trim$(re.Replace(result, " "))
    Set re = Nothing
    SanitizeString = result
End Function

Public Function CalculateAverage(values As Variant) As Variant
    ' Returns Null if no numeric entries are found; otherwise returns the numeric average
    Dim sum As Double
    Dim cnt As Long
    Dim v As Variant
    If TypeName(values) = "Range" Then
        For Each v In values.Cells
            If IsNumeric(v.Value) Then
                sum = sum + CDbl(v.Value)
                cnt = cnt + 1
            End If
        Next v
    ElseIf IsArray(values) Then
        For Each v In values
            If IsNumeric(v) Then
                sum = sum + CDbl(v)
                cnt = cnt + 1
            End If
        Next v
    Else
        If IsNumeric(values) Then
            CalculateAverage = CDbl(values)
            Exit Function
        End If
    End If
    If cnt > 0 Then
        CalculateAverage = sum / cnt
    Else
        CalculateAverage = Null
    End If
End Function