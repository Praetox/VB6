VERSION 5.00
Begin VB.Form frmWhitelist 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Whitelist"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmWhitelist.frx":0000
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmWhitelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Form_Unload(Cancel As Integer)
    If Compiled = True Then On Error Resume Next
    Cancel = 1
    Me.Hide
End Sub
