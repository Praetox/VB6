Attribute VB_Name = "modMain"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public User, Pass, bUser, bPass, Showing, isFree, intMain, abType, xFcAA, OldRankP, TimeSpent, aPopHide, COMPILED As Boolean
Public tKrim, tPress, tFight, tBil, tFengsel, tTTNR, tFcAA, fTTNR, tBuEnter, tBuInvite, tBump, tBankIt
Public aKrim, aPress, aFight, aBil, aFengsel, aBotC, aFcAA, aTTNR, aBuEnter, aBuInvite, aBump, aBreakout, aUtfordrer, aBankIt
Public vBuPris, vBuTimer, vBuChars, vBuEnterLDate, vBuInviteLDate, vUtfordrer, vBumpAdr, vBumpMsg, vBumped
Public svlKrim1, svlKrim2, svlPress1, svlPress2, svlFight1, svlBil1, svlBil2, svlJail1, svlJail2, svlJailLst, svlBotLaunch
Public dlyKrim, dlyPress, dlyFight, dlyBil, dlyFengsel, dlyFcAA, dlyBump
Public SiteRoot As String, SiteBSub As String, SiteMail As String, SiteDL As String, SiteName As String, _
       SiteIP As String, SiteLVer As String, SiteCVer As String, SiteNews As String, SiteForum As String, _
       SiteFAQ As String, SiteCars As String, path As String
Public sW8d As Boolean, Statuslog As String
Private Const Timeout = 60

Sub Main()
    COMPILED = True
    If COMPILED Then On Error Resume Next
    SiteRoot = "http://nordic.awardspace.com/"
    If COMPILED = False Then SiteRoot = App.path & "/botmirror/"
    SiteBSub = SiteRoot & "botinf/"
    SiteMail = SiteBSub & "phost.html"
    SiteName = SiteBSub & "mynameis.html"
    SiteIP = "http://myip.dk/CMyip.dll"
    SiteLVer = SiteBSub & "appLid.html"
    SiteCVer = SiteBSub & "appCid.html"
    SiteNews = SiteBSub & "news.html"
    SiteForum = SiteRoot & "forum"
    SiteFAQ = SiteRoot & "faq.html"
    SiteDL = SiteRoot & "dl_redir.html"
    SiteCars = SiteBSub & "abi.txt"
    path = App.path: If Right(path, 1) <> "\" Then path = path & "\"
    'SiteMail = "http://nm.rasrv.net/post.html"
    'SiteDL = "http://nm.rasrv.net/anm.exe"
    'SiteName = "http://nm.rasrv.net/mynameis.html"
    OldRankP = 10000
    dlyKrim = 183: dlyPress = 963: dlyFight = 123: dlyBil = 363: dlyFengsel = 603: dlyFcAA = 33: dlyBump = 183
    svlBotLaunch = Date$ & " :: " & Time$
    frmLogin.Show
End Sub

Function gEnd(ByVal sec As Long)
    If COMPILED Then On Error Resume Next
    gEnd = Timer + sec
    If gEnd > 86400 Then gEnd = gEnd - 86400
    gEnd = gEnd \ 1
End Function
Function gTil(ByVal endsec As Long)
    If COMPILED Then On Error Resume Next
    gTil = endsec - Timer
    If gTil < 0 Then gTil = gTil + 86400
    If gTil > 10000 Then gTil = 0
    gTil = gTil \ 1
End Function
Function gOver(ByVal endsec As Long)
    If COMPILED Then On Error Resume Next
    gOver = Timer - endsec
    If gOver < 0 Then gOver = gOver + 86400
    If gOver > 10000 Then gOver = 0
    gOver = gOver \ 1
End Function
Function GetTime(ByVal sec As Long) As String
    If COMPILED Then On Error Resume Next
    Dim tH As Long, tM As Long, tHs As String, tMs As String, tSs As String
    If sec > 3599 Then
        tH = Int(Split((sec / 60 / 60), ".")(0))
        sec = sec - (tH * 60 * 60)
        tHs = tH: If Len(tHs) <> 2 Then tHs = "0" & tHs
    Else
        tHs = "00"
    End If
    If sec > 59 Then
        tM = Int(Split((sec / 60), ".")(0))
        sec = sec - (tM * 60)
        tMs = tM: If Len(tMs) <> 2 Then tMs = "0" & tMs
    Else
        tMs = "00"
    End If
    tSs = sec: If Len(tSs) <> 2 Then tSs = "0" & tSs
    GetTime = tHs & ":" & tMs & ":" & tSs
End Function

Sub Nav(Optional URL As String, Optional vl As String = "")
    If COMPILED Then On Error Resume Next
    w8 URL: If LogIn Then w8 URL
    If vl <> "" Then w8u vl
End Sub
Sub w8(Optional URL As String)
    If COMPILED Then On Error Resume Next
    If URL <> "" Then fWB.WB.Stop: DoEvents: fWB.WB.Navigate URL
    wa1 = Timer
    DoEvents
    While fWB.WB.Busy = True
        DoEvents
        diff = Timer - wa1
        If diff < 0 Then diff = -diff
        'frmMain.Caption = diff
        If diff > Timeout Then Exit Sub
        Sleep (1)
    Wend
End Sub
Sub w8i()
    If COMPILED Then On Error Resume Next
    While fWB.WB.Busy = False
        Sleep (5)
        DoEvents
    Wend
End Sub
Sub w8u(ByVal vl As String)
    If COMPILED Then On Error Resume Next
    stt = Timer
    DoEvents
    While ((Timer - stt) < 30) And wpc(vl) = False
        While fWB.WB.Busy
            DoEvents
            Sleep (1)
        Wend
        DoEvents
        Sleep (1)
    Wend
End Sub
Function wpc(ByVal term As String, Optional ncs As Boolean = False) As Boolean
    If COMPILED Then On Error Resume Next
    If ncs = True Then
        If InStr(1, wSRC, term, vbTextCompare) > 0 Then wpc = True
    Else
        If InStr(1, wSRC, term, vbBinaryCompare) > 0 Then wpc = True
    End If
End Function
Function wSRC() As String 'ByVal term As String
    If COMPILED Then On Error Resume Next
    wSRC = fWB.WB.Document.body.parentelement.innerhtml
End Function
Function PFE(ByVal vl, Optional isBlank As Boolean = True) As Boolean
    On Error GoTo 10
    tmp = fWB.WB.Document.All(vl).Value
    If isBlank = False Then
        If fWB.WB.Document.All(vl).Value <> "" Then PFE = True
    Else
        PFE = True
    End If
10
End Function

Function enc(ByVal tx As String)
    For a = 1 To Len(tx)
        enc = enc & Chr(Asc(Mid$(tx, a, 1)) + 1)
    Next
End Function
Function dec(ByVal tx As String)
    For a = 1 To Len(tx)
        dec = dec & Chr(Asc(Mid$(tx, a, 1)) - 1)
    Next
End Function

Sub DirectExit()
    If COMPILED Then On Error Resume Next
    mStop
    Unload frmLogin
    Unload frmMain
    Unload frmConfig
    Unload fWB
    End
End Sub
Sub ExitAPP()
    If COMPILED Then On Error Resume Next
    mStop
    frmMain.WbL.Navigate "http://authenticate.awardspace.com/ANM_REG.php?logout=" & bUser: DoEvents
    While frmMain.WbL.Busy = True
        DoEvents
        Sleep (1)
    Wend
    Unload frmLogin
    Unload frmMain
    Unload frmConfig
    Unload fWB
    End
End Sub


Function LogIn() As Boolean
    If COMPILED Then On Error Resume Next
    If wpc("Du er ikke logget inn!") Or wpc("Du ble logget inn fra et annet sted.") Then
        LogIn = True
        fWB.WB.Document.All("brukernavn").Value = User
        fWB.WB.Document.All("passoord").Value = Pass
        fWB.WB.Document.All("submit").Click
        fStatus "Logger inn...": w8
        fStatus "Logget inn.": DoEvents
    End If
End Function
Function Jail() As Boolean
    If COMPILED Then On Error Resume Next
    If (wpc("Kriminalitet løser ingenting...") And wpc("BoomBoom Spillet")) Then inJail = True
    If (wpc("Du ble tatt mens du skulle bryte en ut,<BR>du er nå i fengsel!")) Then inJail = True
    If (wpc("Du kan kjøpe deg ut av fengselet for ")) Then inJail = True
    If inJail = True Then
        If aFengsel = 0 Then
            Jail = True
            Nav "http://www.nordicmafia.net/nordic/index.php?side=fengsel"
            If InStr(1, wSRC, "Tid igjen:") Then
                tij = Split(Split(wSRC, "<SPAN id=tell>")(1), "</SPAN>")(0) '"Tid igjen:"
                tFengsel = gEnd(tij + 5)
                fStatus "Fengsel i " & tij & " sekunder."
            End If
        Else
            fWB.WB.Document.write ("<form action=""http://www.nordicmafia.net/nordic/index.php?side=fengsel"" method=""POST""><input type=""submit"" name=""subbestikk"" value=""Kjøp deg ut!""></form>")
            DoEvents: fWB.WB.Document.All("subbestikk").Click: w8
        End If
    End If
    DoEvents
End Function
Function AntiBot(ByVal vl As String) As Boolean
    If COMPILED Then On Error Resume Next
    If wpc("tre bilder som inneholder") Then
        AntiBot = True
        If aBotC = 0 Then
            fStatus "Antibot :: Venter..."
            Call w8i: w8
        ElseIf aBotC = 1 Then
            fStatus "Antibot :: Alarm."
            frmMain.abAlert.Enabled = True: w8i
            frmMain.abAlert.Enabled = False: mStop: w8
        ElseIf aBotC = 2 Then
            fStatus "Antibot :: Popup.": fWB.Show: DoEvents
            SetWindowPos fWB.hwnd, -1, 0, 0, (Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY) - 20, 0: DoEvents
            SetWindowPos fWB.hwnd, -2, 0, 0, (Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY) - 20, 0: DoEvents
            Call w8i: If Showing = False Then fWB.Hide
            w8
        ElseIf aBotC = 3 Then
            fStatus "Antibot :: Alarm + popup."
            frmMain.abAlert.Enabled = True: fWB.Show: DoEvents
            SetWindowPos fWB.hwnd, -1, 0, 0, (Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY) - 20, 0: DoEvents
            SetWindowPos fWB.hwnd, -2, 0, 0, (Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY) - 20, 0: DoEvents
            Call w8i: frmMain.abAlert.Enabled = False: mStop: DoEvents
            If Showing = False Then fWB.Hide
            w8
        ElseIf aBotC = 4 Then
            fStatus "Antibot :: BF (automodus)..."
10          imgs = fWB.BreakTheAntibot(vl)
            wbs = "<body bgcolor=""#000000"">" & _
                  "<FORM id=antibot_form action=http://www.nordicmafia.net/nordic/index.php?side=" & vl & " method=post>" & _
                  "<INPUT type=checkbox CHECKED value=" & Int(Mid$(imgs, 1, 1)) - 1 & " name=bilde[" & Int(Mid$(imgs, 1, 1)) - 1 & "]>" & _
                  "<INPUT type=checkbox CHECKED value=" & Int(Mid$(imgs, 2, 1)) - 1 & " name=bilde[" & Int(Mid$(imgs, 2, 1)) - 1 & "]>" & _
                  "<INPUT type=checkbox CHECKED value=" & Int(Mid$(imgs, 3, 1)) - 1 & " name=bilde[" & Int(Mid$(imgs, 3, 1)) - 1 & "]>" & _
                  "<INPUT type=submit value=Fullfør name=antibot_valider></form>"
            fWB.WB.Stop: DoEvents
            fWB.WB.Document.write wbs: DoEvents
            fWB.WB.Document.All("antibot_valider").Click: w8
            If LogIn Then Nav "http://www.nordicmafia.net/nordic/index.php?side=" & vl: GoTo 10
            If wpc("Mislykket! Prøv igjen.") Then GoTo 10
        End If
        fStatus "Antibot fullført."
    End If
    DoEvents
End Function

Public Function ShellandWait(ExeFullPath As String, Optional TimeOutValue As Long = 0) As Boolean
    Dim lInst As Long, LStart As Long, lTimeToQuit As Long, sExeName As String, _
        lProcessId As Long, lExitCode As Long, bPastMidnight As Boolean
    
    If COMPILED Then On Error Resume Next
    LStart = CLng(Timer)
    sExeName = ExeFullPath
    
    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If LStart + TimeOutValue < 86400 Then
            lTimeToQuit = LStart + TimeOutValue
        Else
            lTimeToQuit = (LStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If
    lInst = Shell(sExeName, vbMinimizedNoFocus)
    lProcessId = OpenProcess(&H400, False, lInst)
    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < LStart Then Exit Do
            Else
                 Exit Do
            End If
        End If
        Sleep (5)
    Loop While lExitCode = &H103&
End Function

Sub INIT()
    If COMPILED Then On Error Resume Next
    fStatus "Verifiserer brukernavn og passord..."
    Nav ("about:<form method=""post"" action=""http://www.nordicmafia.net/nordic/index.php?side=loggut""><div align=""center""><input type=""submit"" value=""Logg ut!"" name=""subloggut""><input type=""hidden"" name=""luz""></form>")
    DoEvents
    fWB.WB.Document.All("subloggut").Click: w8
    fWB.WB.Navigate ("http://www.nordicmafia.net/"): w8: LogIn
    If wpc("Feil brukernavn/passord!") Then MsgBox dec("Gfjm!mphjo"""): ExitAPP
    fStatus "Informasjon godkjent. Starter."
    
    frmMain.cmdBump.Enabled = True
    frmMain.cmdSellCars.Enabled = True
    frmMain.cmdSok1.Enabled = True
    frmMain.cmdDoner.Enabled = True
End Sub

Function FEx(FileName As String) As Boolean
    On Error Resume Next
    FEx = (GetAttr(FileName) And vbDirectory) = 0
End Function

Function INI(ByVal iRead As Boolean, iID As String, Optional iVL As String) As String
    If iRead Then
        Dim vl As String: vl = String$(50, 0)
        I = GetPrivateProfileString("AutoNM", iID, "", vl, Len(vl), path & "autoboot.ini")
        If I > 0 Then vl = Left(vl, I): INI = vl
    Else
        I = WritePrivateProfileString("AutoNM", iID, iVL, path & "autoboot.ini")
    End If
End Function

Function WpGT(ByVal address As String) As String
    frmMain.SepWB.Navigate address
    DoEvents
    While frmMain.SepWB.Busy
        Sleep (1)
        DoEvents
    Wend
    'WpGT = frmMain.SepWB.Document.body.parentelement.innerhtml
    WpGT = frmMain.SepWB.Document.body.parentelement.innertext
    'WpGT = Mid$(WpGT, 22, Len(WpGT) - 22 - 6)
    'MsgBox WpGT & " (" & Len(WpGT) & ")"
End Function
Function WpGS(ByVal address As String) As String
    frmMain.SepWB.Navigate address
    DoEvents
    While frmMain.SepWB.Busy
        Sleep (1)
        DoEvents
    Wend
    WpGS = frmMain.SepWB.Document.body.parentelement.innerhtml
End Function

Sub fStatus(ByVal tx As String)
    If Left(frmMain.Status, 10) = "Fengsel i " And Left(tx, 10) = "Fengsel i " Then Exit Sub
    frmMain.Status = tx
    Statuslog = Time$ & " " & tx & vbCrLf & Statuslog
    tmp = Split(Statuslog, vbCrLf)
    If UBound(tmp) > 20 Then
        Statuslog = ""
        For a = 0 To 20
            Statuslog = Statuslog & tmp(a) & vbCrLf
        Next
    End If
End Sub

Function GenHtmlLog() As String
    On Error Resume Next: Dim Woot As String
    t = "<center>" & vbCrLf
    t = t & "<font size=5>AutoNM logg for " & User & "</size><br>" & vbCrLf
    t = t & "<font size=3>Bot åpnet " & svlBotLaunch & "</size><br>" & vbCrLf
    t = t & "<font size=3>Oppdatert " & Date$ & " :: " & Time$ & "</size><br><br>" & vbCrLf
    t = t & "" & vbCrLf
    t = t & "<table align=center>" & vbCrLf
    t = t & GenLogLine(frmMain.lblRank, frmMain.lblRankPerc) & vbCrLf
    t = t & GenLogLine(frmMain.lblPenger, "T2R " & frmMain.lbl2rank) & vbCrLf
    t = t & GenLogLine(" ", " ") & vbCrLf
    percval = 0: percval = (100 / svlKrim1) * svlKrim2
    t = t & GenLogLine("<b>" & Int(svlKrim1) & "</b> Kriminalitet", Int(percval) & "% Success rate") & vbCrLf
    percval = 0: percval = (100 / svlPress1) * svlPress2
    t = t & GenLogLine("<b>" & Int(svlPress1) & "</b> Utpressing", Int(percval) & "% Success rate") & vbCrLf
    percval = 0: percval = (100 / svlBil1) * svlBil2
    t = t & GenLogLine("<b>" & Int(svlBil1) & "</b> Biltyveri", Int(percval) & "% Success rate") & vbCrLf
    t = t & GenLogLine("<b>" & Int(svlFight1) & "</b> Fightclub", "Level <b>" & frmMain.lblFCLvl & "</b>") & vbCrLf
    t = t & GenLogLine(" ", " ") & vbCrLf
    percval = 0: percval = (100 / svlJail1) * svlJail2
    t = t & GenLogLine("<b>" & Int(svlJail1) & "</b> Utbrytninger", Int(percval) & "% Success rate") & vbCrLf
    t = t & "" & vbCrLf
    If svlJailLst <> "" Then
        tmplol = Split(svlJailLst, vbCrLf)
        For a = 0 To UBound(tmplol) - 1
            Woot = tmplol(a)
            t = t & GenLogLine(" ", Woot) & vbCrLf
        Next
    End If
    't = t & Replace(svlJailLst, vbCrLf, "<br>" & vbCrLf)
    t = t & "</td></table>"
    GenHtmlLog = t
End Function
Function GenLogLine(ByVal VL1 As String, VL2 As String) As String
    GenLogLine = "<tr><td width=150><i>" & VL1 & "</i></td>" & _
                 "<td width=150><i>" & VL2 & "</i></td></tr>"
End Function

Sub DownloadFile(ByVal file As String, path As String)
    Call URLDownloadToFile(0, file, path, 0, 0)
End Sub
Function ROT13(ByVal tIN As String) As String
    For a = 1 To Len(tIN)
        tmp = Asc(Mid$(tIN, a, 1))
        If tmp >= 65 And tmp <= 90 Then
            If tmp <= 77 Then tmp = tmp + 13 Else tmp = tmp - 13
        ElseIf tmp >= 97 And tmp <= 122 Then
            If tmp <= 109 Then tmp = tmp + 13 Else tmp = tmp - 13
        End If
        ROT13 = ROT13 & Chr(tmp)
    Next
End Function

