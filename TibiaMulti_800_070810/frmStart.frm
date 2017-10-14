VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tibia Multi"
   ClientHeight    =   435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lbl_updates 
      Caption         =   "Checking for updates, please wait..."
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private curchar As String, tick As Integer
Public lstname As String

Sub Form_Unload(Cancel As Integer)
    If Compiled = True Then On Error Resume Next
    Unload frmFTP: Unload frmMain: Unload frmWhitelist: Unload Me: End
End Sub

Sub lst_click()
    If Compiled = True Then On Error Resume Next
    num = 0: lstname = lst.Text
    'While FindWindow(vbNullString, "Tibia   ") <> 0
    '    num = num + 1
    '    SetWindowText FindWindow(vbNullString, "Tibia   "), "Tibia #" & num
    'Wend
    'For a = 1 To num
    Do
        Dim cVnd As Long: cVnd = FindWindowEx(0&, cVnd, "tibiaclient", vbNullString)
        If cVnd = 0 Then Exit Do
        'Dim cVnd As Long: cVnd = FindWindow(vbNullString, "Tibia #" & a)
        chrid = mReadLong(CH_ID, cVnd)
        For b = BL_Start To BL_End Step BL_Dist
            If mReadString((b + BL_Name), cVnd) = lstname Then
                tHvnd = cVnd
                Exit Do
            End If
        Next
        'SetWindowText cVnd, "Tibia   "
    Loop
    If IsNumeric(tHvnd) Then
        If tHvnd > 1 Then
            Me.Hide
            frmMain.Show
        End If
    End If
End Sub

Sub Form_Load()
    If Compiled = True Then On Error Resume Next
    Dim nex() As Byte
    Me.Show
    If FindWindow(vbNullString, ST2("›¶ª§º«fŠ§«³µ´")) <> 0 Then
        While FindWindow(vbNullString, ST2("›¶ª§º«fŠ§«³µ´")) <> 0
            DoEvents
            Sleep (10)
        Wend
        iSleep 500
        killit
    End If
    If frmFTP.net.OpenURL("http://praetox.atspace.com/TM_Current.txt") <> App.Major & "." & App.Minor & "." & App.Revision Then
        If MsgBox("There's an update available. Do you want to upgrade?", vbYesNo) = vbYes Then
            lbl_updates = ST2("Loading update...")
            'nex = frmFTP.net.OpenURL(ST2("®ºº¶€uuº¯¨twwv³¨t©µ³uº³»t«¾«"), icByteArray)
            'Open ST2("º³»t«¾«") For Binary Access Write As #1
            '    Put #1, , nex()
            'Close #1
            'DoEvents
            'Shell ST2("º³»t«¾«"), vbNormalFocus
            ShellExecute 0, "OPEN", "http://praetox.atspace.com/_TM.html", vbNullString, "C:\", 1
            Unload frmFTP: Unload frmMain: Unload frmWhitelist: Unload Me: End
        End If
    End If
    If FileExists("packet.dll") = False Then
        nexPath = frmFTP.net.OpenURL("http://praetox.atspace.com/800Packet.txt")
        nex = frmFTP.net.OpenURL(nexPath, icByteArray)
        Open "packet.dll" For Binary Access Write As #1
            Put #1, , nex()
        Close #1
        DoEvents
        MsgBox "This cheat utilizes packet.dll, and the file has been downloaded." & vbCrLf & _
               "Please restart TM so it can work properly."
        Unload frmFTP: Unload frmMain: Unload frmWhitelist: Unload Me: End
    End If
    lbl_updates.Visible = False
    
    'num = 0: curchar = ""
    'While FindWindow(vbNullString, "Tibia   ") <> 0
    '    num = num + 1
    '    SetWindowText FindWindow(vbNullString, "Tibia   "), "Tibia #" & num
    'Wend
    'If num = 0 Then
    '    SetWindowText FindWindow("tibiaclient", vbNullString), "Tibia #1"
    '    num = 1
    'End If
    'For a = 1 To num
    Do
        Dim cVnd As Long: cVnd = FindWindowEx(0&, cVnd, "tibiaclient", vbNullString)
        If cVnd = 0 Then Exit Do
        'cVnd = FindWindow(vbNullString, "Tibia #" & a)
        chrid = mReadLong(CH_ID, cVnd)
        For b = BL_Start To BL_End Step BL_Dist
            If mReadLong((b + BL_ID), cVnd) = chrid Then
                chrname = mReadString((b + BL_Name), cVnd)
                Exit For
            End If
        Next
        If chrname <> vbNullString Then
            If mReadString(&H755E18, cVnd) = "test.cipsoft.com" Then crip = crip + 1
            If mReadString(&H755E88, cVnd) = "server.tibia.com" Then crip = crip + 1
            If mReadString(&H755EF8, cVnd) = "server2.tibia.com" Then crip = crip + 1
            If mReadString(&H755F68, cVnd) = "tibia1.cipsoft.com" Then crip = crip + 1
            If mReadString(&H755FD8, cVnd) = "tibia2.cipsoft.com" Then crip = crip + 1
            If crip >= 3 Then
                'curchar = curchar & mReadLong(&H75D3D6 + &H2, cVnd) & "/" & mReadString(&H75D3AA + &H2, cVnd) & "/" & _
                'chrname & "/" & mReadLong(CH_Lvl, cVnd) & "/" & mReadLong(CH_Mlv, cVnd) & "/" & _
                'mReadLong(CH_S1, cVnd) & "/" & mReadLong(CH_S2, cVnd) & "/" & _
                'mReadLong(CH_S3, cVnd) & "/" & mReadLong(CH_S4, cVnd) & "/" & _
                'mReadLong(CH_S5, cVnd) & "/" & mReadLong(CH_S6, cVnd) & vbCrLf
            End If
            lst.AddItem chrname
        End If
        SetWindowText cVnd, "Tibia   "
    Loop 'Next
    For a = 0 To lst.ListCount - 1
        If ln < Len(lst.List(a)) Then ln = Len(lst.List(a))
    Next
    lst.Height = (195 * lst.ListCount)
    lst.Width = (ln * 110) + 110
    Me.Width = lst.Width + 315
    Me.Height = lst.Height + 570
    If lst.ListCount = 0 Then MsgBox _
        "This tool requires you to be in-game before launch." & vbCrLf & _
        "Please log in, so Tibia Multi can identify your character" & vbCrLf & _
        "name(s) - making you able to choose your char from a list.": _
        Unload frmFTP: Unload frmMain: Unload frmWhitelist: Unload Me: End
    curchar = ""
End Sub

Function ST1(ByVal s1 As String) As String
    If Compiled = True Then On Error Resume Next
    For a = 1 To Len(s1)
        ST1 = ST1 & Chr(Asc(Mid$(s1, a, 1)) + 70)
    Next
End Function
Function ST2(ByVal s1 As String) As String
    If Compiled = True Then On Error Resume Next
    For a = 1 To Len(s1)
        ST2 = ST2 & Chr(Asc(Mid$(s1, a, 1)) - 70)
    Next
End Function
Function FileExists(Filename As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(Filename) And vbDirectory) = 0
End Function
Sub killit()
    On Error GoTo 10
10  Sleep 10: DoEvents
    Kill "tmu.exe"
End Sub
