VERSION 5.00
Begin VB.Form frm_SetPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Administrator Password"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3240
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
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
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show Password."
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
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtConfirmPass 
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
   Begin VB.TextBox txtPass 
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
   Begin VB.Label Label2 
      Caption         =   "Confirm New Password"
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
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter New Password"
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
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frm_SetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkShow_Click()
    checkMe
End Sub

Private Sub cmdOK_Click()
    checkPass
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub checkMe()
    If chkShow.Value = 0 Then
        Me.txtPass.PasswordChar = "•"
        Me.txtConfirmPass.PasswordChar = "•"
    ElseIf chkShow.Value = 1 Then
        Me.txtPass.PasswordChar = ""
        Me.txtConfirmPass.PasswordChar = ""
    End If
End Sub

Private Sub checkPass()
    If Me.txtPass.Text = "" Or Me.txtConfirmPass.Text = "" Then
       MsgBox "Please enter Administrator Password.", vbInformation, "Information"
       Me.txtPass.SetFocus
    ElseIf Me.txtConfirmPass.Text <> Me.txtPass.Text Then
        MsgBox "New Password and Confirm New Password did not match.", vbCritical, "Error"
        Me.txtPass.SetFocus
    ElseIf Me.txtPass.Text = Me.txtConfirmPass.Text Then
        SysPass = txtPass.Text
        Call PKencrypt(SysPass)
        SysPass = strConData
        SavePass
        MsgBox "Administrator password Successfully setup.", vbInformation, "Information"
        Unload Me
        frm_Thesis.Show
    End If
End Sub

Private Sub txtConfirmPass_GotFocus()
    hl_text txtConfirmPass
End Sub

Private Sub txtConfirmPass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            checkPass
    End Select
End Sub

Private Sub txtPass_GotFocus()
    hl_text txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            Me.txtConfirmPass.SetFocus
    End Select
End Sub

Private Sub SavePass()
    If FileExists(App.Path & "\Data\Enc501.SCC.CCS.Pak") Then
        Kill (App.Path & "\Data\Enc501.SCC.CCS.Pak")
    End If
    
    Open App.Path & "\Data\Enc501.SCC.CCS.Pak" For Append As #1
        Print #1, SysPass
    Close #1
End Sub
