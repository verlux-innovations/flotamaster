Attribute VB_Name = "digitalsignature"
Option Explicit

Private Const CAPICOM_CURRENT_USER_STORE As Long = 2
Private Const CAPICOM_STORE_OPEN_READ_ONLY As Long = 0

Private Function IsVBAModelAccessible() As Boolean
    On Error Resume Next
    Dim prj As VBIDE.VBProject
    Set prj = ThisWorkbook.VBProject
    IsVBAModelAccessible = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function VerifySignature() As Boolean
    If Not IsVBAModelAccessible Then
        MsgBox "Programmatic access to the VBA project model is disabled." & vbCrLf & _
               "Enable 'Trust access to the VBA project object model' in the Trust Center.", vbCritical
        VerifySignature = False
        Exit Function
    End If
    On Error GoTo ErrHandler
    Dim vbProj As VBIDE.VBProject
    Dim sig As VBIDE.Signature
    Set vbProj = ThisWorkbook.VBProject
    Set sig = vbProj.Signature
    If Not sig.Signed Then
        Debug.Print "VerifySignature: Project is unsigned."
        VerifySignature = False
    ElseIf Not sig.IsSignatureValid Then
        Debug.Print "VerifySignature: Signature is invalid."
        VerifySignature = False
    ElseIf sig.IsCertificateExpired Then
        Debug.Print "VerifySignature: Certificate is expired."
        VerifySignature = False
    Else
        VerifySignature = True
    End If
    Exit Function
ErrHandler:
    Debug.Print "VerifySignature Error " & Err.Number & ": " & Err.Description & " Source: " & Err.Source
    VerifySignature = False
End Function

Public Sub SignProject()
    If Not IsVBAModelAccessible Then
        MsgBox "Programmatic access to the VBA project model is disabled." & vbCrLf & _
               "Enable 'Trust access to the VBA project object model' in the Trust Center.", vbCritical
        Exit Sub
    End If
    On Error GoTo ErrHandler
    Dim vbProj As VBIDE.VBProject
    Dim certInput As String
    Dim fso As Object
    Dim isFile As Boolean
    Set vbProj = ThisWorkbook.VBProject
    certInput = InputBox("Enter certificate subject name or PFX file path to sign the project:", "Sign VBA Project")
    If Len(Trim(certInput)) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    isFile = (InStr(certInput, "\") > 0 Or InStr(certInput, "/") > 0)
    If isFile Then
        If Not fso.FileExists(certInput) Then
            MsgBox "The specified PFX file does not exist: " & certInput, vbExclamation
            Exit Sub
        End If
    Else
        If Not CertificateExists(certInput) Then
            MsgBox "No certificate found with subject: " & certInput, vbExclamation
            Exit Sub
        End If
    End If
    vbProj.Sign certInput
    MsgBox "VBA project signed successfully with: " & certInput, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Failed to sign VBA project." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical
End Sub

Private Function CertificateExists(subject As String) As Boolean
    On Error GoTo ErrHandler
    Dim store As Object
    Dim cert As Object
    Set store = CreateObject("CAPICOM.Store")
    store.Open CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_READ_ONLY
    For Each cert In store.Certificates
        If InStr(1, cert.SubjectName, subject, vbTextCompare) > 0 Then
            CertificateExists = True
            Exit Function
        End If
    Next
    CertificateExists = False
    Exit Function
ErrHandler:
    CertificateExists = False
End Function