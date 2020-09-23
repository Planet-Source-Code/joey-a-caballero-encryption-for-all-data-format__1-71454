Attribute VB_Name = "Mod_SEncrypted"
Public Function SaveEnc(PFile As String, FileTitle As String) As String
On Error GoTo err:

    Dim SVEnc As String
    Dim sPath As String
    Dim FileO As Integer

    sPath = App.Path & "\Data\EncList.SCC.CCS.Pak"
    SVEnc = PFile & "//" & FileTitle
        
    FileO = FreeFile
    
    Open sPath For Append As #FileO
       Print #FileO, "" & SVEnc & ""
    Close #FileO
    
    Exit Function
err:
    MsgBox err.Description, vbCritical, "Save Encrypted Error"
End Function

Public Sub Init_List()
    Dim sPath As String
    Dim FileO As Integer
    Dim Dsource, Dfile As String
    
    sPath = App.Path & "\Data\EncList.SCC.CCS.Pak"
    frm_Thesis.lvwList.ListItems.Clear
    FileO = FreeFile
    If FileExists(sPath) = True Then
        Open sPath For Input As #FileO
            Do Until EOF(1)
            'Get FileO, , strData
            Input #FileO, ListSV
            Dsource = Split(ListSV, "//")(0)
            Dfile = Split(ListSV, "//")(1)
            'Dbytes = Split(ListSV, "//")(2)
            'Dcon = Split(Source, "//")(3)
            frm_Thesis.lvwList.ListItems.Add , , Dsource
            frm_Thesis.lvwList.ListItems(frm_Thesis.lvwList.ListItems.Count).SubItems(1) = Dfile
            'frm_Thesis.lvwList.ListItems(frm_Thesis.lvwList.ListItems.Count).SubItems(2) = Dbytes
            'frm_Thesis.lvwList.ListItems(frm_Thesis.lvwList.ListItems.Count).SubItems(3) = Dcon
            Loop
        Close #FileO
    Else
        Exit Sub
    End If
End Sub

Public Sub RMList()
On Error GoTo err:
    Dim sPath As String
    
    Dim sTemp1 As String
    Dim sTemp2 As String
    
    If frm_Thesis.lvwList.ListItems.Count <> 0 Then
        frm_Thesis.lvwList.ListItems.Remove (intIndex)
    
    
        sPath = App.Path & "\Data\EncList.SCC.CCS.Pak"
        
        If FileExists(sPath) = True Then
            Kill sPath
        End If

    
         Open sPath For Append As #1
         For i = 1 To frm_Thesis.lvwList.ListItems.Count
             sTemp1 = frm_Thesis.lvwList.ListItems.Item(i)
             sTemp2 = frm_Thesis.lvwList.ListItems(i).SubItems(1)
             If sTemp1 <> "" And sTemp2 <> "" Then
                Print #1, "" & frm_Thesis.lvwList.ListItems.Item(i) & "//" & frm_Thesis.lvwList.ListItems(i).SubItems(1) & ""
             End If
         Next
         Close #1
         
    End If
    
    'frm_Thesis.lvwList.SelectedItem.Selected = False
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
End Sub
