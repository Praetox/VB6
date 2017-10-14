VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update Daemon"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet net 
      Left            =   3000
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label a3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Done! Exiting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label a2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Overwriting old TM..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label a1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading update..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub form_load()
    Me.Show
    a1.Visible = True
    Dim nex() As Byte
    nex = net.OpenURL(ST2("®ºº¶€uuº¯¨twwv³¨t©µ³uº³t«¾«"), icByteArray)
    a2.Visible = True
    deltm
    Open ST2("º³t«¾«") For Binary Access Write As #1
        Put #1, , nex()
    Close #1
    a3.Visible = True
    DoEvents
    tmp = MsgBox("Update completed. Run new TM?", vbYesNo)
    If tmp = vbYes Then Shell ST2("º³t«¾«"), vbNormalFocus
    Unload Me: End
End Sub
Private Sub deltm()
    If FindWindow(vbNullString, ST2("š¯¨¯§f“»²º¯")) <> 0 Then
        While FindWindow(vbNullString, ST2("š¯¨¯§f“»²º¯")) <> 0
            DoEvents
            Sleep (10)
        Wend
        Sleep (250)
    End If
    On Error GoTo 10
10  DoEvents
    If FileExists(ST2("º³t«¾«")) Then
        Kill (ST2("º³t«¾«"))
    End If
End Sub
Private Function ST1(ByVal s1 As String) As String
    For a = 1 To Len(s1)
        ST1 = ST1 & Chr(Asc(Mid$(s1, a, 1)) + 70)
    Next
End Function
Private Function ST2(ByVal s1 As String) As String
    For a = 1 To Len(s1)
        ST2 = ST2 & Chr(Asc(Mid$(s1, a, 1)) - 70)
    Next
End Function
Function FileExists(Filename As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(Filename) And vbDirectory) = 0
End Function

