VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AutoNM :: Login"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Nordicmafia user / pass"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   3735
      Begin VB.TextBox nmPass 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox nmUser 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AutoNM Access Verification"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3735
      Begin VB.TextBox botUser 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox botPass 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Timer tLogin 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   -120
   End
   Begin VB.CommandButton cmdAvslutt 
      BackColor       =   &H00FFBE58&
      Caption         =   "Avslutt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdEmail 
      BackColor       =   &H00FFBE58&
      Caption         =   "FAQ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFBE58&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblUpdate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFDD88&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   3735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private proceed As Boolean, wbls As String
Private Const AuthSite = "http://authenticate.awardspace.com/"
Private Const AuthHash = "3B609821152A2E16BC535B1F9918E5EB"

Private Sub cmdStart_Click()
    If COMPILED Then On Error Resume Next
    If cmdStart.Enabled = False Then Exit Sub
    If nmUser <> "" And nmPass <> "" And botUser <> "" And botPass <> "" Then
        botUser = UCase(botUser)
        cmdStart.Enabled = False
        cmdEmail.Enabled = False
        lblUpdate = "Verifying..."
        frmMain.WbL.Navigate AuthSite: DoEvents
        While frmMain.WbL.Busy = True
            DoEvents
            Sleep (1)
        Wend
        frmMain.WbL.Document.All("anmuser").Value = botUser
        frmMain.WbL.Document.All("anmpass").Value = botPass
        frmMain.WbL.Document.All("submit").Click: DoEvents
        While frmMain.WbL.Busy = True
            DoEvents
            Sleep (1)
        Wend
        src = Split(Split(frmMain.WbL.Document.body.parentelement.innerhtml, "<BODY>")(1), "</BODY>")(0)
        If UCase(src) = UCase(md5(botUser & botPass & "verified_abcdefghijklmnop")) Then
            ResumeLaunch = True
        ElseIf src = "MULTIPLE_USERS" Then
            MsgBox "Account banned: Suspected sharing. Contact Shade.": DirectExit
        Else
            MsgBox "Fuck off, wannabe hacker.": DirectExit
        End If
        frmMain.WbL.Navigate "about:No updates available."
        If md5(AuthSite) <> AuthHash Then MsgBox "Stop hex-editing my bot, fucker.": DirectExit
        If ResumeLaunch = False Then DirectExit
        
        User = nmUser
        Pass = nmPass
        bUser = botUser
        bPass = botPass
        frmLogin.Hide
        frmMain.Show
        frmMain.Timer1.Enabled = True
        frmMain.lblNews = Replace(wbls, "<BR>", vbCrLf)
    Else
        MsgBox "Vennligst skriv inn ditt brukernavn og passord for NordicMafia." & vbCrLf & vbCrLf & _
               "Please enter your nordicmafia username and password."
    End If
End Sub

'Private Sub cmdStart_Click()
'    If COMPILED Then On Error Resume Next
'    If cmdStart.Enabled = False Then Exit Sub
'    If txtUser <> "" And txtPass <> "" Then
'        User = txtUser
'        Pass = txtPass
'        frmLogin.Hide
'        frmMain.Show
'        frmMain.Timer1.Enabled = True
'        frmMain.lblNews = Replace(wbls, "<BR>", vbCrLf)
'    Else
'        MsgBox "Vennligst skriv inn ditt brukernavn og passord for NordicMafia." & vbCrLf & vbCrLf & _
'               "Please enter your nordicmafia username and password."
'    End If
'End Sub

Private Sub cmdEmail_Click()
    ShellExecute Me.hwnd, "OPEN", SiteFAQ, vbNullString, "C:\", 1
    DirectExit
End Sub
Private Sub cmdAvslutt_Click()
    If COMPILED Then On Error Resume Next
    ExitAPP
End Sub

Private Sub Form_Load()
    If COMPILED Then On Error Resume Next
    l "Start"
    Me.Caption = "AutoNM v" & App.Major & "." & App.Minor & "." & App.Revision & " :: Login"
    Me.Show: wbLV = WpGT(SiteLVer): l "L" & wbLV: DoEvents
             wbCV = WpGT(SiteCVer): l "C" & wbCV: DoEvents
    AppVID = (App.Major * 100 * 100) + (App.Minor * 100) + (App.Revision)
    'AppNEW = Split(wbCV / 10000, ".")(0) & "." & Right(Split(wbCV / 100, ".")(0), 2) & "." & Right(wbCV, 2)
    AppNEW = Mid$(wbCV, 1, 2) & "." & Mid$(wbCV, 3, 2) & "." & Mid$(wbCV, 5, 2)
    If (Int(wbLV) > Int(AppVID)) Or (Int(wbCV) < Int(AppVID)) Then
        uinf = dec(")johfo!jogp")
        MsgBox dec("Oz!cpu;") & " " & AppNEW & vbCrLf & vbCrLf & _
               dec("Tubsufs!epxompbe///")
        ShellExecute Me.hwnd, "OPEN", SiteDL, vbNullString, vbNullString, 1
        ExitAPP
    End If
    l "xForce"
    If Int(wbCV) > Int(AppVID) Then
        If MsgBox("En ny versjon (" & AppNEW & ") er tilgjengelig. Oppdater?", vbYesNo, "Update daemon") = vbYes Then
            ShellExecute Me.hwnd, "OPEN", SiteDL, vbNullString, vbNullString, 1
            ExitAPP
        End If
    End If
    l "xAsk"
    wbls = WpGT(SiteNews)
    
    lblUpdate = "Optimerer cracker"
    aKrim = 1: aPress = 1: aFight = 1: aBil = 1: aFengsel = 0: aTTNR = 1: fTTNR = 120: aBreakout = 0: aUtfordrer = 0
    vUtfordrer = 118: aBankIt = 1
    
    If FEx("autoboot.ini") Then
        bUser = INI(True, "bUser")
        bPass = ROT13(INI(True, "bPass"))
        User = INI(True, "User")
        Pass = ROT13(INI(True, "Pass"))
        aKrim = INI(True, "aKrim")
        aPress = INI(True, "aPress")
        aFight = INI(True, "aFight")
        aBil = INI(True, "aBil")
        aFengsel = INI(True, "aFengsel")
        aTTNR = INI(True, "aTTNR")
        fTTNR = INI(True, "fTTNR")
        aBreakout = INI(True, "aBreakout")
        aUtfordrer = INI(True, "aUtfordrer")
        vUtfordrer = INI(True, "vUtfordrer")
        aBankIt = INI(True, "aBankIt")
        aBotC = INI(True, "aBotC")
        'If aBotC = 4 Then aBotC = 3
        If aBotC = 4 Then abType = vbYes Else abType = vbNo
        nmUser = User: nmPass = Pass: botUser = bUser: botPass = bPass
    Else
        abType = MsgBox("Ønsker du å bruke den helautomatiske antiboten?", vbYesNo, "Velg antibot breaker")
        'aBotC = 3: abType = vbNo
    End If
    If abType = vbYes Then
        'Unload fWB: Unload frmMain: DoEvents
        'Open path & "reg.reg" For Output As #1
        '    Print #1, "Windows Registry Editor Version 5.00"
        '    Print #1, ""
        '    Print #1, "[HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main]"
        '    Print #1, """Display Inline Images""=""no"""
        'Close #1
        'ShellandWait ("regedit /s " & path & "reg.reg"): fWB.WB.Navigate ("about:blank"): DoEvents
        'Open path & "reg.reg" For Output As #1
        '    Print #1, "Windows Registry Editor Version 5.00"
        '    Print #1, ""
        '    Print #1, "[HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main]"
        '    Print #1, """Display Inline Images""=""yes"""
        'Close #1
        'ShellandWait ("regedit /s " & path & "reg.reg")
        'Kill path & "reg.reg"
        aBotC = 4
    Else
        If aBotC = 0 Then aBotC = 3
    End If
    cmdStart.Enabled = True: cmdEmail.Enabled = True: tLogin.Enabled = True
    lblUpdate = "Alle operasjoner fullført."
    
    frmMain.SepWB.Silent = True
    frmMain.WbL.Silent = True
    fWB.WB.Silent = True
End Sub

Private Sub tLogin_Timer()
    tLogin.Enabled = False
    If User <> "" And Pass <> "" Then
        cmdStart_Click
        frmMain.cmdStart_Click
    End If
End Sub

Private Sub nmPass_KeyDown(KeyCode As Integer, Shift As Integer)
    If COMPILED Then On Error Resume Next
    If KeyCode = 13 Then cmdStart_Click
End Sub

Private Sub l(ByVal tx As String)
    lblUpdate = lblUpdate & tx & " - "
End Sub
