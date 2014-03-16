Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    Dim ret As Boolean
    
    ret = konekToServer
    frmMain.Show
End Sub
