Attribute VB_Name = "Mod_PKencrypt"
Public SysPass As String
Public Ekey As String
Public strConData As String


Public Function PKencrypt(cString As String) As String

    Dim X As Integer
    
    For X = 1 To Len(cString)
        Convert = Convert + Chr(255 - Asc(Mid(cString, X, 1)))
    Next X
    
    strConData = Convert

End Function
