Attribute VB_Name = "configmodule"
Option Explicit

Const SECTION_NAME As String = "Config"

Private pConfig As Object
Private pIniPath As String

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String _
) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Private Function GetIniPath() As String
    If pIniPath = "" Then
        If ThisWorkbook.Path <> "" Then
            pIniPath = ThisWorkbook.Path & "\config.ini"
        Else
            pIniPath = Application.DefaultFilePath & "\config.ini"
        End If
    End If
    GetIniPath = pIniPath
End Function

Public Sub LoadConfig()
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim path As String: path = GetIniPath()
    Dim raw As String
    Dim size As Long: size = 1024
    Dim ret As Long

    Do
        raw = String$(size, vbNullChar)
        ret = GetPrivateProfileString(SECTION_NAME, vbNullString, "", raw, size, path)
        If ret < size - 1 Then Exit Do
        size = size * 2
    Loop

    If ret > 0 Then
        Dim arrKeys() As String
        arrKeys = Split(Left$(raw, ret), vbNullChar)
        Dim k As Variant
        For Each k In arrKeys
            If Len(k) > 0 Then
                Dim valBuf As String
                Dim vLen As Long
                Dim vSize As Long: vSize = 255
                Do
                    valBuf = String$(vSize, vbNullChar)
                    vLen = GetPrivateProfileString(SECTION_NAME, k, "", valBuf, vSize, path)
                    If vLen < vSize - 1 Then Exit Do
                    vSize = vSize * 2
                Loop
                dict(k) = Left$(valBuf, vLen)
            End If
        Next k
    End If

    Set pConfig = dict
    SetDefaults
    SaveConfig
End Sub

Public Sub SetDefaults()
    Dim defaults As Variant
    defaults = Array( _
        Array("DataFolder", ThisWorkbook.Path & "\Data"), _
        Array("OutputFolder", ThisWorkbook.Path & "\Output"), _
        Array("LogLevel", "INFO"), _
        Array("TimeoutSeconds", "60"), _
        Array("CommentsTemplate", "Standard"), _
        Array("BestCollectorCriteria", "MaxCuRecovery") _
    )
    Dim i As Long
    For i = LBound(defaults) To UBound(defaults)
        Dim key As String: key = defaults(i)(0)
        Dim value As String: value = defaults(i)(1)
        If Not pConfig.Exists(key) Then pConfig(key) = value
    Next i
End Sub

Public Function GetConfigValue(key As String) As Variant
    If pConfig Is Nothing Then LoadConfig
    If pConfig.Exists(key) Then
        GetConfigValue = pConfig(key)
    Else
        GetConfigValue = Empty
    End If
End Function

Public Sub SaveConfig()
    Dim path As String: path = GetIniPath()
    On Error Resume Next: Kill path: On Error GoTo 0
    Dim k As Variant
    For Each k In pConfig.Keys
        Dim res As Long
        res = WritePrivateProfileString(SECTION_NAME, CStr(k), CStr(pConfig(k)), path)
        If res = 0 Then Debug.Print "Failed to write key " & CStr(k) & " to INI file: " & path
    Next k
End Sub