VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Postal Launchpad"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Shell "regsvr32 mswinsck.ocx /s"
    'FileCopy "the.app", "Postal.exe"
    'DoEvents: MsgBox "Installed the ocx." & vbCrLf & "You may start Postal."
    If FEx("the.ocx") = False Or FEx("the.app") = False Or FEx("iPostal.exe") = False Then
        MsgBox "Please extract the .zip file to a folder, and run this application from there." & vbCrLf & vbCrLf & _
               "Pakk ut .zip filen til en mappe, og start dette programmet derfra."
        End
    End If
    Open "script.bat" For Output As #1
        Print #1, "@echo off"
        Print #1, "copy ""the.ocx"" """ & Environ("windir") & "\system32\mswinsck.ocx"""
        Print #1, "regsvr32 mswinsck.ocx /s"
        Print #1, "copy ""the.app"" ""Postal.exe"""
        Print #1, "del the.app"
        Print #1, "del the.ocx"
        Print #1, "del iPostal.exe"
        Print #1, "START Postal.exe"
        Print #1, "del script.bat"
    Close #1
    Shell "script.bat"
    End
End Sub

Private Function FEx(ByVal file As String) As Boolean
    On Error Resume Next
    FEx = (GetAttr(file) And vbDirectory) = 0
End Function
