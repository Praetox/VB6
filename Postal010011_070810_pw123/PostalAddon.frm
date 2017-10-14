VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Postal Addon"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   1455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tMain 
      Interval        =   3
      Left            =   480
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub tMain_Timer()
    If GetAsyncKeyState(145) = -32767 Then
        tmp = Clipboard.GetText
        tmp = Split(tmp, ", ")
        For a = 0 To UBound(tmp)
            Clipboard.Clear
            Clipboard.SetText tmp(a)
            DoEvents
            SendKeys "^v", 1
            DoEvents
            SendKeys "{enter}", 1
            DoEvents
            Sleep 50
        Next
    End If
End Sub
