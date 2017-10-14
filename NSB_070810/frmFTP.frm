VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFTP 
   BackColor       =   &H00000000&
   Caption         =   "Tibia Multi Stats Uploader"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   5040
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8880
      TabIndex        =   24
      Text            =   "FFFFFF"
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Save"
      Height          =   255
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Load"
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Reset"
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox sig 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   4680
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   13
      Top             =   120
      Width           =   5415
   End
   Begin InetCtlsObjects.Inet net 
      Left            =   240
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2280
   End
   Begin VB.TextBox interval 
      Alignment       =   2  'Center
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
      Height          =   240
      Left            =   2760
      TabIndex        =   10
      Text            =   "180"
      ToolTipText     =   "Color of the light - default is 203, torch is 206. Use the up/down keys on your keyboard to change it."
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Contents 
      Alignment       =   2  'Center
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
      Height          =   2880
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmFTP.frx":0000
      ToolTipText     =   "Color of the light - default is 203, torch is 206. Use the up/down keys on your keyboard to change it."
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Filename 
      Alignment       =   2  'Center
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
      Height          =   240
      Left            =   1200
      TabIndex        =   6
      Text            =   "index.html"
      ToolTipText     =   "Color of the light - default is 203, torch is 206. Use the up/down keys on your keyboard to change it."
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Website 
      Alignment       =   2  'Center
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
      Height          =   240
      Left            =   1200
      TabIndex        =   4
      Text            =   "ftp."
      ToolTipText     =   "Color of the light - default is 203, torch is 206. Use the up/down keys on your keyboard to change it."
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Password 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Color of the light - default is 203, torch is 206. Use the up/down keys on your keyboard to change it."
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Username 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Color of the light - default is 203, torch is 206. Use the up/down keys on your keyboard to change it."
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label setText 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Maglvl"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   9240
      TabIndex        =   20
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label setText 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exp"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   8520
      TabIndex        =   19
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label setText 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7800
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label setText 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Online"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label setText 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Position"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label setText 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mana"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   15
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label setText 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HP"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   14
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "second(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Upload every"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   4695
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1575
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Website"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   855
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   495
      Width           =   855
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   855
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private delay As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private setting, ptx(6), pty(6), ImgPath As String

Sub cmdReset_Click()
    zForm_Load
    sig.ForeColor = &HFFFFFF
End Sub
Sub cmdLoad_Click()
    Open "ftp.txt" For Input As #1
    Line Input #1, vl:     UserName = vl
    Line Input #1, vl:     Password = vl
    Line Input #1, vl:     Website = vl
    Line Input #1, vl:     Filename = vl
    Line Input #1, vl:     Contents = vl
    Line Input #1, vl:     Interval = vl
    Line Input #1, vl:     txtColor = vl
    Line Input #1, vl:     sig.ForeColor = vl
    Line Input #1, vl:     ImgPath = vl
    For a = 0 To 6
        Line Input #1, vl: ptx(a) = vl
        Line Input #1, vl: pty(a) = vl
    Next
    Close #1
    DoEvents
    setTxts
End Sub
Sub cmdSave_Click()
    Open "ftp.txt" For Output As #1
    Print #1, UserName
    Print #1, Password
    Print #1, Website
    Print #1, Filename
    Print #1, Contents
    Print #1, Interval
    Print #1, txtColor
    Print #1, sig.ForeColor
    Print #1, ImgPath
    For a = 0 To 6
        Print #1, ptx(a)
        Print #1, pty(a)
    Next
    Close #1
End Sub

Sub zForm_Load()
    For a = 0 To 6
        ptx(a) = 130
    Next
    For a = 0 To 6
        pty(a) = 10 + (a * 15)
    Next a
    ImgPath = "Graphics\abstract_02m.JPG"
    DoEvents
    setTxts
End Sub

Sub Form_Unload(Cancel As Integer)
    If Compiled = True Then On Error Resume Next
    Cancel = 1
    Me.Hide
End Sub

Sub setText_Click(Index As Integer)
    setting = Index
    For a = 0 To 6
        If a = setting Then SetText(a).BackColor = &H8000& Else SetText(a).BackColor = &H80&
    Next
End Sub

Sub sig_DblClick()
    cmdlg.Filter = "Image files (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
    cmdlg.ShowOpen
    If cmdlg.CancelError = 1 Then Exit Sub
    ImgPath = cmdlg.Filename
    setTxts
End Sub

Sub sig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ptx(setting) = X
        pty(setting) = Y
        setTxts
    End If
End Sub
Sub setTxts()
    DoEvents
    sig.Picture = LoadPicture(ImgPath)
    For a = 0 To 6
        Dim r As RECT
        r.Left = ptx(a)
        r.Top = pty(a)
        r.Right = (ptx(a)) + 1000
        r.Bottom = (pty(a)) + 1000
        If a = 0 Then txt = "HP:     " & mReadLong(CH_HP)
        If a = 1 Then txt = "Mana:   " & mReadLong(CH_Ma)
        If a = 2 Then txt = "Pos:    " & mReadLong(CH_X) & "," & mReadLong(CH_Y) & "," & mReadLong(CH_Z)
        If mReadLong(CH_Con) > 0 Then con = "Yes" Else con = "No"
        If a = 3 Then txt = "Online: " & con
        If a = 4 Then txt = "Level:  " & mReadLong(CH_Lvl)
        If a = 5 Then txt = "Exp:    " & mReadLong(CH_Exp)
        If a = 6 Then txt = "Maglvl: " & mReadLong(CH_Mlv)
        DrawText sig.hdc, txt, Len(txt), r, 0
    Next
End Sub

Sub tmr_Timer()
    If Compiled = True Then On Error Resume Next
    delay = delay + 1
    Me.Caption = "Upload in " & Interval - delay & " seconds."
    If delay >= Interval Then
        delay = 0
        Me.Caption = "Uploading..."
        
        Open App.Path & "\" & Filename For Output As #1
            tmp = Contents
            online = mReadLong(CH_Con)
            If online = 8 Then online = "Yes" Else online = "No"
            tmp = Replace(tmp, "{name}", mReadString(BL_Player + BL_Name))
            tmp = Replace(tmp, "{hp}", mReadLong(CH_HP))
            tmp = Replace(tmp, "{mana}", mReadLong(CH_Ma))
            tmp = Replace(tmp, "{online}", online)
            tmp = Replace(tmp, "{x}", mReadLong(CH_X))
            tmp = Replace(tmp, "{y}", mReadLong(CH_Y))
            tmp = Replace(tmp, "{z}", mReadLong(CH_Z))
            tmp = Replace(tmp, "{lvl}", mReadLong(CH_Lvl))
            tmp = Replace(tmp, "{exp}", mReadLong(CH_Exp))
            If InStr(1, tmp, "{exp2lvl}") > 0 Then
                cLevel = mReadLong(CH_Lvl) + 1
                cExp = mReadLong(CH_Exp)
                cExpNext = (((50 / 3) * (cLevel ^ 3)) - (100 * (cLevel ^ 2)) + ((850 / 3) * cLevel) - 200) - cExp
                tmp = Replace(tmp, "{exp2lvl}", cExpNext)
            End If
            tmp = Replace(tmp, "{maglvl}", mReadLong(CH_Mlv))
            tmp = Replace(tmp, vbCrLf, "<br>" & vbCrLf)
            Print #1, tmp
        Close #1
    
        net.Protocol = icFTP
        net.RemoteHost = Website
        net.UserName = UserName
        net.Password = Password
        net.Execute , "PUT """ & App.Path & "\" & Filename & """ """ & Filename & """"
        Me.Caption = "Upload completed."
    End If
End Sub

Sub txtColor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        colr = CByte("&H" & Mid(txtColor, 1, 2))
        colg = CByte("&H" & Mid(txtColor, 3, 2))
        colb = CByte("&H" & Mid(txtColor, 5, 2))
        sig.ForeColor = (colb * 256 * 256) + (colg * 256) + colr
        setTxts
    End If
End Sub
