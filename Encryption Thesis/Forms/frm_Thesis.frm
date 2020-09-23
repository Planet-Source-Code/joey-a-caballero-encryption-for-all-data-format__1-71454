VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Thesis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   Icon            =   "frm_Thesis.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin Project1.XP_ProgressBar xp_Prog 
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   8421504
      Scrolling       =   5
      ShowText        =   -1  'True
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files  (*.*)"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Thesis.frx":625A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Thesis.frx":67F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Thesis.frx":6D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Thesis.frx":1D860
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Thesis.frx":34332
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Thesis.frx":4A35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Thesis.frx":4A8F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   953
      ButtonWidth     =   1588
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Key"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Encrypt"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Decrypt"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Password"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exit"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2055
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483634
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Source File"
         Object.Width           =   8291
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Byte"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblSource 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label lblShow 
      Caption         =   "&Encrypted List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblConversionTime 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label lblBytes 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblFileName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   2
      X1              =   40
      X2              =   8160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   1080
      MouseIcon       =   "frm_Thesis.frx":4AE90
      Stretch         =   -1  'True
      Top             =   960
      Width           =   720
   End
   Begin VB.Menu mnupopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDecrypt 
         Caption         =   "&Decrypt Selected"
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
   End
End
Attribute VB_Name = "frm_Thesis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Init_List
    CheckMenu
    Call PKencrypt(Ekey)
    Ekey = strConData
End Sub

Private Sub lvwList_DblClick()
If Me.lvwList.ListItems.Count <> 0 Then
    Call PopupMenu(mnupopup)
    Me.ToolBar.Buttons(6).Enabled = False
End If
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cd.FileName = Empty
    strFile = Me.lvwList.SelectedItem
    lblSource.Caption = Me.lvwList.SelectedItem
    strSourceFile = Me.lvwList.SelectedItem.Text
    Me.lblFileName.Caption = Me.lvwList.SelectedItem.SubItems(1)
    Me.lblBytes.Caption = Format(FileLen(strFile), "###,###,###,###") & " Bytes"
    Me.lblConversionTime.Caption = "Estimated Conversion Time"
    GetConvertionTime

    intIndex = Me.lvwList.SelectedItem.Index
    SetIcon
    Me.ToolBar.Buttons(8).Enabled = True
    Me.ToolBar.Buttons(6).Enabled = False
End Sub

Private Sub mnuDecrypt_Click()
    dCrypto
End Sub


Private Sub mnuRemove_Click()
    RMList
End Sub

Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 2
        GetFile
        frm_Thesis.xp_Prog.Value = 0
    Case 4
        frm_Key.Show 1
    Case 6
        Crypto
    Case 8
        dCrypto
    Case 10
        frm_SetPassword.Show 1
    Case 12
        Unload Me
        
End Select
End Sub

Private Sub GetFile()
On Error GoTo err:
Dim cSec As Integer
Dim cMin As Integer

    cd.ShowOpen
    
    strSourceFile = cd.FileName
    strFile = cd.FileTitle
    If strFile <> "" Then
        
        Me.lblFileName.Caption = strFile
        strByte = Format(FileLen(strFile), "###,###,###,###")
        Me.lblBytes.Caption = strByte & " Bytes"
        Me.lblConversionTime.Caption = "Estimated Conversion Time"
        
        'Me.lvwList.ListItems.Clear
        'Me.lvwList.ListItems.Add , , strSourceFile
        'Me.lvwList.ListItems(Me.lvwList.ListItems.Count).SubItems(1) = strFile
        'Me.lvwList.ListItems(Me.lvwList.ListItems.Count).SubItems(2) = Format(FileLen(strFile), "###,###,###,##0") & " Bytes"
        GetConvertionTime
            
        SetIcon
        Me.ToolBar.Buttons(6).Enabled = True
    End If
    Exit Sub
err:
MsgBox err.Description, vbCritical, "Error"
End Sub
Private Sub GetConvertionTime()
    cRate = RATE_ENCRYPT
    If FileLen(strFile) > 75000000 Then
        tmplen = FileLen(strFile) / 20
        cSec = tmplen / cRate
    Else
        cSec = FileLen(strFile) / cRate
    End If
    cMin = Int(cSec / 60)
    cSec = Int(cSec - (cMin * 60))
    If cSec = 0 Then cSec = 1
    lblTime.Caption = "min: " & Trim(Str(cMin)) & "  sec: " & Trim(Str(cSec))
    'Me.lvwList.ListItems(Me.lvwList.ListItems.Count).SubItems(3) = "min: " & Trim(Str(cMin)) & "  sec: " & Trim(Str(cSec))
End Sub

Private Sub Crypto()

If MsgBox("Are you sure you want to Encrypt the data of this File?", vbQuestion + vbYesNo, "System Information") = vbYes Then
    strSourceFile = cd.FileName
    LookUp = True
    Call EncFile(strSourceFile, strSourceFile)
    strSourceFile = ""
    cd.FileName = Empty
Else
    Exit Sub
End If
Me.ToolBar.Buttons(6).Enabled = False
End Sub

Private Sub dCrypto()
If MsgBox("Are you sure you want to Decrypt the data of this File?", vbQuestion + vbYesNo, "System Information") = vbYes Then
    Me.ToolBar.Buttons(8).Enabled = False
    LookUp = False
    Call EncFile(strSourceFile, strSourceFile)
    If strSourceFile = "" Then
        Exit Sub
    End If
    'If strSourceFile = cd.FileName Then
        'MsgBox "You browse the file." & vbCrLf & "If the file is in the List of Encrypted file, please remove it manually.", vbInformation, "Information"
        'strSourceFile = ""
        'cd.FileName = Empty
        'Exit Sub
    'End If
    RMList
Else
    Exit Sub
End If
End Sub


Private Sub CheckMenu()
If Me.lvwList.ListItems.Count = 0 Then
    Me.ToolBar.Buttons(8).Enabled = False
Else
    Me.ToolBar.Buttons(8).Enabled = True
End If
    Me.ToolBar.Buttons(6).Enabled = False
End Sub


Private Sub SaveEnc()
    Open frm_Thesis.cd.FileName For Binary As #6
        Put #6, , strConData
    Close #6
End Sub

