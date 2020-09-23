Attribute VB_Name = "ProgressInStatusBar"


Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Type RECT
  Left        As Long
  Top         As Long
  Right       As Long
  Bottom      As Long
End Type

Const WM_USER    As Long = &H400
Const SB_GETRECT As Long = (WM_USER + 10)


Public Sub ShowProgressInStatusBar(ByRef Progress As Control, ByRef StatusBar As StatusBar, ByVal PanelNumber As Long)

    Dim TRC As RECT
        StatusBar.Panels(PanelNumber).Width = Progress.Width + 15
        SendMessageAny StatusBar.hwnd, SB_GETRECT, PanelNumber - 1, TRC
       
         
        With Progress
            SetParent .hwnd, StatusBar.hwnd
            .Move TRC.Left + ((TRC.Right - TRC.Left) / 2) - (.Width / 2), TRC.Top + ((TRC.Bottom - TRC.Top) / 2) - (.Height / 2), .Width, .Height
            .Visible = True
            .Value = 0
        End With
        
End Sub
