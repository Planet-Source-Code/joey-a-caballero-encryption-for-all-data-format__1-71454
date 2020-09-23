Attribute VB_Name = "Mod_Encryption"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''' Variables Dec '''''''''''''''''''''''''''''''
Dim OutData$
Dim InData$
Dim SCC%

Private Const FileLim = 10240


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''Encryption Here'''''''''''''''''''''''''''''''
Private Sub Encrypt(strEncData$, strEncKey$)

    If Len(strEncData$) = 0 Then Exit Sub
    'frm_Thesis.xp_Prog.Max = Len(WorkingSTring) ' Progress bar Max value
    For CCS = 1 To Len(strEncData$)
        DoEvents
        SCC% = SCC% + 1
        
        If SCC% > Len(strEncKey$) Then SCC% = 1
        strParseData$ = Mid$(strEncData$, CCS, 1)
        strParseKey$ = Mid$(strEncKey$, SCC%, 1)
        
        intAsciiVal% = Asc(strParseData$) + (Asc(strParseKey$) - 1)
        If intAsciiVal% > 255 Then intAsciiVal% = intAsciiVal% - 256
        
        strTempData$ = strTempData$ + Chr$(intAsciiVal%)
        
        'frm_Thesis.xp_Prog.Value = CCS 'Progress bar increment
    Next CCS
    
    OutData$ = OutData$ + strTempData$
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''Decryption Here'''''''''''''''''''''''''''''''

Private Sub Decrypt(strEncData$, strEncKey$)
    If Len(strEncData$) = 0 Then Exit Sub
    'frm_Thesis.xp_Prog.Max = Len(WorkingSTring) ' Progress bar Max value
    For CCS = 1 To Len(strEncData$)
        DoEvents
        SCC% = SCC% + 1
        
        If SCC% > Len(strEncKey$) Then SCC% = 1
        strParseData$ = Mid$(strEncData$, CCS, 1)
        strParseKey$ = Mid$(strEncKey$, SCC%, 1)
        
        intAsciiVal% = Asc(strParseData$) - (Asc(strParseKey$) - 1)
        If intAsciiVal% < 0 Then intAsciiVal% = intAsciiVal% + 256
        
        strTempData$ = strTempData$ + Chr$(intAsciiVal%)
        'frm_Thesis.xp_Prog.Value = CCS 'Progress bar increment
    Next CCS
    
    OutData$ = OutData$ + strTempData$

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Extract Data for Encryption/Decryption Here'''''''''''''''
Public Function EncFile(ByVal sSource As String, ByVal sDestination As String) As Boolean

On Error GoTo err:

    Dim FileO As Integer
    Dim strData As String
    Dim X As Double
    Dim z As Double
    
    FileO = FreeFile
    
    SCC% = 0
    
    frm_Thesis.Enabled = False
    frm_Thesis.MousePointer = 11
    
    Open sSource For Binary As FileO
    
        If LOF(FileO) > FileLim Then
                strData = Space(FileLim)
                z = Int(LOF(FileO) / FileLim)
                
                If z >= 10000 Then
                    z = z / 100
                End If
                
                    frm_Thesis.xp_Prog.Max = z  ' Progress bar Max value
                    For X = 1 To z
                        DoEvents
                        Get FileO, , strData
                        
                        If LookUp = True Then
                            Call Encrypt(strData, Ekey)
                        Else
                            Call Decrypt(strData, Ekey)
                        End If
                        
                        frm_Thesis.xp_Prog.Value = X 'Progress bar increment
                        
                    Next X
            
        ElseIf LOF(FileO) <= FileLim And LOF(FileO) > 0 Then
            strData = Space(LOF(FileO))
            Get FileO, , strData
            frm_Thesis.xp_Prog.Max = 2 ' Progress bar Max value
                    
            If LookUp = True Then
                Call Encrypt(strData, Ekey)
            Else
                Call Decrypt(strData, Ekey)
            End If
            
            frm_Thesis.xp_Prog.Value = 1 'Progress bar increment
            
        End If
    
    Close FileO

    
    Open sSource For Binary As FileO
       Put FileO, , OutData$
    Close FileO
    
        frm_Thesis.xp_Prog.Value = frm_Thesis.xp_Prog.Value + 1 'Progress bar increment
        If LookUp = True Then
            MsgBox "Encryption Successful"
            Call SaveEnc(sSource, frm_Thesis.lblFileName)
        Else
            MsgBox "Decryption Successful"
        End If
        
        strData = Empty
        OutData$ = Empty
        
        frm_Thesis.Enabled = True
        frm_Thesis.MousePointer = 0
        
        SetIcon
        Init_List
        Exit Function
err:
    MsgBox err.Description, vbCritical, "System information Error"
    frm_Thesis.Enabled = True
    frm_Thesis.MousePointer = 0
End Function




