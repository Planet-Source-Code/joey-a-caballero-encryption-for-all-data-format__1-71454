VERSION 5.00
Begin VB.Form frm_Key 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption Key"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEncKey 
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "•"
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox txtEncKeyCon 
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show Encryption Key."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Encryption Key"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Confirm Encryption Key"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frm_Key"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub checkMe()
    If chkShow.Value = 0 Then
        Me.txtEncKey.PasswordChar = "•"
        Me.txtEncKeyCon.PasswordChar = "•"
    ElseIf chkShow.Value = 1 Then
        Me.txtEncKey.PasswordChar = ""
        Me.txtEncKeyCon.PasswordChar = ""
    End If
End Sub

Private Sub chkShow_Click()
    checkMe
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub checkKey()
    If Me.txtEncKey.Text = "" Or Me.txtEncKeyCon.Text = "" Then
       MsgBox "Please enterEncryption Key.", vbInformation, "Information"
       Me.txtEncKey.SetFocus
    ElseIf Me.txtEncKeyCon.Text <> Me.txtEncKey.Text Then
        MsgBox "Encryption Key and Confirm Encryption Key did not match.", vbCritical, "Error"
        Me.txtEncKey.SetFocus
    ElseIf Me.txtEncKey.Text = Me.txtEncKeyCon.Text Then
        Ekey = txtEncKey.Text
        Call PKencrypt(Ekey)
        Ekey = strConData
        SaveKey
        
    End If
End Sub

Private Sub SaveKey()
    If FileExists(App.Path & "\Data\Enc500.SCC.CCS.Pak") = True Then
        If MsgBox("Their is an existing encryption key. Do you want to change it?" & vbCrLf & vbCrLf & "Warning" & vbCrLf & "Changing Encryption Key may lead your encrypted file not to be decrypted anymore.", vbInformation + vbYesNo, "Information") = vbYes Then
            Kill (App.Path & "\Data\Enc500.SCC.CCS.Pak")
            ReKey
        Else
            Exit Sub
        End If
    End If
    
    Open App.Path & "\Data\Enc500.SCC.CCS.Pak" For Append As #1
        Print #1, Ekey
    Close #1
    Unload Me
    MsgBox "Encryption Key Successfully setup.", vbInformation, "Information"
    Call main
End Sub

Private Sub cmdOK_Click()
    checkKey
End Sub

Private Sub txtEncKey_GotFocus()
    hl_text txtEncKey
End Sub

Private Sub txtEncKey_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            Me.txtEncKeyCon.SetFocus
    End Select
End Sub

Private Sub txtEncKeyCon_GotFocus()
    hl_text txtEncKeyCon
End Sub

Private Sub txtEncKeyCon_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            checkKey
    End Select
End Sub

Private Sub ReKey()
If FileExists(App.Path & "\Data\Enc500.SCC.CCS.Pak") = True Then
    Kill (App.Path & "\Data\Enc500.SCC.CCS.Pak")
End If
    Open App.Path & "\Data\Enc500.SCC.CCS.Pak" For Append As #1
        Print #1, Ekey
    Close #1
    MsgBox "Encryption Key Successfully setup. To Activate New Encryption Key System need to restart." & vbCrLf & vbCrLf & "System will now Terminate.", vbInformation, "Information"
    End
End Sub
