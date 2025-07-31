Option Explicit

Public Sub GenerateReport(results As Variant, templates As Variant, Optional ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim comments As Variant, suggestions As Variant
    comments = BuildText(results, templates, "Comments")
    suggestions = BuildText(results, templates, "Suggestions")
    ClearPreviousReport ws
    WriteText ws, comments, 0
    WriteText ws, suggestions, 1
End Sub

Private Sub ClearPreviousReport(ws As Worksheet)
    ws.Range("B105:C200").ClearContents
End Sub

Private Sub WriteText(ws As Worksheet, lines As Variant, colOffset As Long)
    Dim i As Long
    If IsEmpty(lines) Then Exit Sub
    For i = LBound(lines) To UBound(lines)
        ws.Cells(106 + i, 2 + colOffset).Value = lines(i)
    Next i
End Sub

Private Function BuildText(results As Variant, templates As Variant, keyName As String) As Variant
    Dim tpl As Variant, out() As String
    Dim i As Long, key As Variant
    If Not IsObject(templates) Then
        BuildText = Array(): Exit Function
    End If
    If Not templates.Exists(keyName) Then
        BuildText = Array(): Exit Function
    End If
    tpl = templates(keyName)
    If Not IsArray(tpl) Then
        BuildText = Array(): Exit Function
    End If
    ReDim out(LBound(tpl) To UBound(tpl))
    For i = LBound(tpl) To UBound(tpl)
        out(i) = tpl(i)
        If IsObject(results) Then
            For Each key In results.Keys
                out(i) = Replace(out(i), "{" & key & "}", CStr(results(key)))
            Next key
        End If
    Next i
    BuildText = out
End Function