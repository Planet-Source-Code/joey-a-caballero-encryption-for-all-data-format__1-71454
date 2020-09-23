Attribute VB_Name = "Mod_Global"
Public strFile As String
Public intIndex As Integer
Public strSourceFile As String
Public strByte As String
Public LookUp As Boolean
Public Const RATE_ENCRYPT = 520000
Public Const RATE_DECRYPT = 520000
Public Sub main()
    ChkData
    
If Ekey = "" Then
    MsgBox "Please set up your Encryption Key before proceeding.", vbInformation + vbOKOnly, "Information"
    frm_Key.Show 1
ElseIf SysPass <> "" Then
    frm_Login.Show 1
Else
    MsgBox "Please set up your Administrator Password now.", vbInformation + vbOKOnly, "Information"
    frm_SetPassword.Show 1
End If

End Sub

Private Sub ChkData()

On Error Resume Next
    Open App.Path & "\Data\Enc501.SCC.CCS.Pak" For Input As #1
        Input #1, SysPass
    Close #1
    
On Error Resume Next
    Open App.Path & "\Data\Enc500.SCC.CCS.Pak" For Input As #1
        Input #1, Ekey
    Close #1
    
End Sub
