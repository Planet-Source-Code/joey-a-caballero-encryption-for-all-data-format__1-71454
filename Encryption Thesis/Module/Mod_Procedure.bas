Attribute VB_Name = "Mod_Procedure"
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSID
    id((123)) As Byte
End Type
Private Const MAX_PATH = 260
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1

Public Sub SetIcon()
Dim fName As String
Dim Index As Integer
Dim hIcon As Long
Dim item_num As Long
Dim icon_pic As IPictureDisp
Dim sh_info As SHFILEINFO
Dim cls_id As CLSID
Dim hRes As Long
Dim new_icon As TypeIcon
Dim lpUnk As IUnknown
Dim strpath As String
strpath = strSourceFile
fName = Trim(strpath)
If fName = "" Then frm_Thesis.imgIcon.Picture = Nothing: Exit Sub
SHGetFileInfo fName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_LARGEICON
hIcon = sh_info.hIcon
With new_icon
    .cbSize = Len(new_icon)
    .picType = vbPicTypeIcon
    .hIcon = hIcon
End With
With cls_id
    .id(8) = &HC0
    .id(15) = &H46
End With
hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
If hRes = 0 Then Set icon_pic = lpUnk
frm_Thesis.imgIcon = icon_pic

End Sub

Public Sub hl_text(oText As TextBox)
    With oText
        .SelStart = 0
        .SelLength = Len(oText)
        
    End With
End Sub


Public Function FileExists(FullFileName As String) As Boolean
   On Error GoTo err:
   Open FullFileName For Input As #1
   Close #1
   FileExists = True
   Exit Function
   
err:
   FileExists = False
   Exit Function
End Function

