Attribute VB_Name = "loggingmodule"
Option Explicit

Private logFilePath As String
Private logBuffer As Collection
Private isInitialized As Boolean
Private bufferLimit As Long
Private Const DEFAULT_BUFFER_LIMIT As Long = 100

Public Sub SetBufferLimit(ByVal limit As Long)
    If limit > 0 Then bufferLimit = limit
End Sub

Public Sub InitializeLogger(Optional ByVal logFileName As String = "", Optional ByVal customBufferLimit As Variant)
    Dim dirSeparator As String: dirSeparator = Application.PathSeparator
    Dim logDir As String
    Dim testFilePath As String
    Dim fn As Integer
    
    If isInitialized Then Exit Sub
    If logBuffer Is Nothing Then Set logBuffer = New Collection
    If Not IsMissing(customBufferLimit) And IsNumeric(customBufferLimit) Then
        If CLng(customBufferLimit) > 0 Then bufferLimit = CLng(customBufferLimit)
    End If
    If bufferLimit = 0 Then bufferLimit = DEFAULT_BUFFER_LIMIT
    If logFileName = "" Then
        If ThisWorkbook.Path <> "" Then
            logFileName = ThisWorkbook.Path & dirSeparator & "FlotaMasterAnalyzer.log"
        Else
            logFileName = CurDir & dirSeparator & "FlotaMasterAnalyzer.log"
        End If
    End If
    logFilePath = logFileName
    logDir = Left(logFilePath, InStrRev(logFilePath, dirSeparator) - 1)
    If Dir(logDir, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1000, "InitializeLogger", "Log directory does not exist: " & logDir
    End If
    testFilePath = logDir & dirSeparator & "~logtest.tmp"
    On Error GoTo ErrTestWrite
    fn = FreeFile
    Open testFilePath For Binary Access Write As #fn
    Close #fn
    Kill testFilePath
    On Error GoTo 0
    If Dir(logFilePath) = "" Then
        fn = FreeFile
        Open logFilePath For Output As #fn
        Close #fn
    End If
    isInitialized = True
    Exit Sub
ErrTestWrite:
    Err.Raise vbObjectError + 1001, "InitializeLogger", "Cannot write to log directory: " & logDir
End Sub

Private Sub EnsureInitialized()
    If Not isInitialized Then InitializeLogger
    If logBuffer Is Nothing Then Set logBuffer = New Collection
End Sub

Private Sub AddLogEntry(ByVal level As String, ByVal message As String)
    Dim timeStamp As String
    timeStamp = Format(Now, "yyyy-mm-dd HH:nn:ss")
    logBuffer.Add timeStamp & " [" & level & "] " & message
    If logBuffer.Count >= bufferLimit Then
        Dim errMsg As String
        FlushLogs errMsg
    End If
End Sub

Public Sub LogInfo(ByVal message As String)
    EnsureInitialized
    AddLogEntry "INFO", message
End Sub

Public Sub LogWarning(ByVal message As String)
    EnsureInitialized
    AddLogEntry "WARNING", message
End Sub

Public Sub LogError(ByVal message As String)
    EnsureInitialized
    AddLogEntry "ERROR", message
End Sub

Public Function FlushLogs(Optional ByRef errorMsg As String) As Boolean
    Dim fnum As Integer
    Dim i As Long
    FlushLogs = False
    errorMsg = ""
    If Not isInitialized Then
        errorMsg = "Logger not initialized."
        Exit Function
    End If
    If logBuffer Is Nothing Or logBuffer.Count = 0 Then
        FlushLogs = True
        Exit Function
    End If
    On Error GoTo ErrHandler
    fnum = FreeFile
    Open logFilePath For Append As #fnum
    For i = 1 To logBuffer.Count
        Print #fnum, logBuffer(i)
    Next i
    Close #fnum
    Set logBuffer = New Collection
    FlushLogs = True
    Exit Function
ErrHandler:
    errorMsg = Err.Description
    On Error Resume Next
    Close #fnum
End Function