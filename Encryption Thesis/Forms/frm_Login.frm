VERSION 5.00
Begin VB.Form frm_Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator Login"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4650
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
      TabIndex        =   2
      Top             =   1200
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
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
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
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter your administrator password."
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
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    If Me.txtPassword.Text = SysPass Then
        Unload Me
        frm_Thesis.Show
    Else
        MsgBox "Access Denied!!!", vbCritical, "Error Accessing"
        hl_text txtPassword
        Me.txtPassword.SetFocus
    End If
End Sub


Private Sub Form_Load()
    Call PKencrypt(SysPass)
    SysPass = strConData
End Sub

Private Sub txtPassword_GotFocus()
    hl_text txtPassword
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            cmdOK_Click
    End Select
End Sub
