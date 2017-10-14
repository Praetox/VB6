Attribute VB_Name = "modMain"
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Inbox(14) As Mail, Outbox(999) As Mail, Filter As Long, SendDelay As Long, ASDelay As Long, ORDelay As Long
Public COMPILED As Boolean, Path As String, User As String, Pass As String, Website As String
Const Timeout = 30

Type Mail
    ID As Long
    Name As String
    Topic As String
    Body As String
    Unread As Boolean
    rDate As String
    rTime As String
End Type

Sub Main()
    If App.PrevInstance Then MsgBox "NPS is already running. Exiting...": ExitAPP
    COMPILED = True: Path = App.Path: Website = "uggc://cenrgbk.ngfcnpr.pbz/"
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    modLang.Lang_Init
    
    ieImages False
    frmMain.Show
    DoEvents
    ieImages True
    
    User = GetSetting("Praetox_NPS", "Nordicmafia Profile", "Username")
    Pass = GetSetting("Praetox_NPS", "Nordicmafia Profile", "Password")
    CgMR = GetSetting("Praetox_NPS", "Configuration", "cfg_MaxRead")
    CgAS = GetSetting("Praetox_NPS", "Configuration", "cfg_Autosync")
    CgAC = GetSetting("Praetox_NPS", "Configuration", "cfg_Autocheck")
    CgCR = GetSetting("Praetox_NPS", "Configuration", "cfg_CntRefresh")
    CgCW = GetSetting("Praetox_NPS", "Configuration", "cfg_CntWarn")
    CgCL = GetSetting("Praetox_NPS", "Configuration", "cfg_CntLaunch")
    CgLN = GetSetting("Praetox_NPS", "Configuration", "cfg_Language")
    'User = INI(True, "USR")
    'Pass = INI(True, "PWD")
    frmMain.LoadContacts
    If User <> "" Then frmMain.accUser = User
    If Pass <> "" Then frmMain.accPass = Pass
    If User = "" Or Pass = "" Then frmMain.accUser.SetFocus
    If CgMR <> "" Then frmMain.cfgMaxRead = CgMR
    If CgAS <> "" Then frmMain.cfgAutosync = CgAS
    If CgAC <> "" Then frmMain.cfgAutocheck = CgAC
    If CgCR <> "" Then frmMain.cfgCntRefresh = CgCR
    If CgCW <> "" Then frmMain.cfgCntWarn = CgCW
    If CgCL <> "" Then frmMain.cfgCntLaunch = CgCL
    If CgLN <> "" Then modLang.Lang_Set CgLN Else modLang.Lang_Set "ENG"
    If User <> "" And Pass <> "" And CgMR <> "" And CgAC = True Then frmMain.cmdRefresh_Click
    If User <> "" And Pass <> "" And CgCL = True Then
        frmMain.ContactList_Refresh False
    End If
    frmMain.tSender.Enabled = True
    frmMain.tAutosync.Enabled = True
End Sub
Sub ExitAPP()
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next
    End
End Sub

Sub Nav(Optional URL As String, Optional vl As String = "")
    If COMPILED Then On Error Resume Next
    w8 URL: If LogIn Then w8 URL
    If vl <> "" Then w8u vl
End Sub
Sub w8(Optional URL As String)
    If COMPILED Then On Error Resume Next
    If URL <> "" Then frmMain.wb.Stop: DoEvents: frmMain.wb.Navigate URL
    wa1 = Timer
    DoEvents
    While frmMain.wb.Busy = True
        DoEvents
        diff = Timer - wa1
        If diff < 0 Then diff = -diff
        If diff > Timeout Then
            wav "slow"
            Exit Sub
        End If
        Sleep (1)
    Wend
End Sub
Sub w8i()
    If COMPILED Then On Error Resume Next
    While frmMain.wb.Busy = False
        Sleep (5)
        DoEvents
    Wend
End Sub
Sub w8u(ByVal vl As String)
    If COMPILED Then On Error Resume Next
    stt = Timer
    DoEvents
    While ((Timer - stt) < 30) And WPC(vl) = False
        While frmMain.wb.Busy
            DoEvents
            Sleep (1)
        Wend
        DoEvents
        Sleep (1)
    Wend
End Sub
Function wSRC() As String 'ByVal term As String
    If COMPILED Then On Error Resume Next
    wSRC = frmMain.wb.Document.Body.parentelement.innerhtml
End Function
Function WPC(ByVal term As String, Optional ncs As Boolean = False) As Boolean
    If COMPILED Then On Error Resume Next
    If ncs = True Then
        If InStr(1, wSRC, term, vbTextCompare) > 0 Then WPC = True
    Else
        If InStr(1, wSRC, term, vbBinaryCompare) > 0 Then WPC = True
    End If
End Function
Function PFE(ByVal vl, Optional isBlank As Boolean = True) As Boolean
    On Error GoTo 10
    tmp = frmMain.wb.Document.All(vl).Value
    If isBlank = False Then
        If frmMain.wb.Document.All(vl).Value <> "" Then PFE = True
    Else
        PFE = True
    End If
10
End Function
Function FEx(FileName As String) As Boolean
    On Error Resume Next
    FEx = (GetAttr(FileName) And vbDirectory) = 0
End Function
Sub Dump(Optional vl As String)
    If vl = "" Then vl = wSRC
    Fn = FreeFile
    Open "c:\_dmp.html" For Output As #Fn
    Print #Fn, vl
    Close #Fn
End Sub

Function LogIn() As Boolean
    If COMPILED Then On Error Resume Next
    If WPC("Du er ikke logget inn!") Or WPC("Du ble logget inn fra et annet sted.") Then
        LogIn = True
        frmMain.wb.Document.All("brukernavn").Value = User
        frmMain.wb.Document.All("passoord").Value = Pass
        frmMain.wb.Document.All("submit").Click
        L 2, "Logging in...": w8
        L 2, "Logged in."
    End If
End Function
Sub LogOut()
    If COMPILED Then On Error Resume Next
    Nav ("about:<form method=""post"" action=""http://www.nordicmafia.net/nordic/index.php?side=loggut""><div align=""center""><input type=""submit"" value=""Logg ut!"" name=""subloggut""><input type=""hidden"" name=""luz""></form>")
    DoEvents
    frmMain.wb.Document.All("subloggut").Click
    L 2, "Logging out...": w8
    L 2, "Logged out."
End Sub

Function gEnd(ByVal vl As Long, Optional Thimer As Long = -1)
    If COMPILED Then On Error Resume Next
    If Thimer = -1 Then Thimer = Timer
    gEnd = Thimer + vl
    If gEnd > 86400 Then gEnd = gEnd - 86400
    gEnd = gEnd \ 1
End Function
Function gTil(ByVal vl As Long, Optional Thimer As Long = -1)
    If COMPILED Then On Error Resume Next
    If Thimer = -1 Then Thimer = Timer
    gTil = vl - Thimer
    If gTil < 0 Then gTil = gTil + 86400
    If gTil > 600 Then gTil = 0
    gTil = gTil \ 1
End Function
Function INI(ByVal iRead As Boolean, iID As String, Optional iVL As String) As String
    If iRead Then
        Dim vl As String: vl = String$(50, 0)
        i = GetPrivateProfileString("NPS", iID, "", vl, Len(vl), Path & "NPS.ini")
        If i > 0 Then vl = Left(vl, i): INI = vl
    Else
        i = WritePrivateProfileString("NPS", iID, iVL, Path & "NPS.ini")
    End If
End Function

Function Session() As String
    Session = Environ("userprofile") & "\cookies\"
    Session = Session & Dir$(Session & "*www.nordicmafia[*")
    Fn = FreeFile
    Open Session For Input As #Fn
    Line Input #Fn, Session
    Close #Fn
    Session = Split(Session, vbLf)(1)
End Function
Function GetMsgAry(ByVal MsgID As Long) As Mail
    For a = 0 To UBound(Inbox)
        If Inbox(a).ID = MsgID Then
            GetMsgAry = Inbox(a)
            Exit Function
        End If
    Next
End Function
Sub SetMsgAry(ByVal MsgID As Long, MsgData As Mail)
    For a = 0 To UBound(Inbox)
        If Inbox(a).ID = MsgID Then
            Inbox(a) = MsgData
            Exit Sub
        End If
    Next
End Sub

Sub L(ByVal im As Integer, vl As String)
    vl = Time$ & " :: " & vl
    If im = 1 Then frmMain.ST2 = ""
    If im = 1 Then frmMain.ST1 = vl Else frmMain.ST2 = vl
End Sub
Sub wav(ByVal vl As String, Optional Async As Integer = 1)
    PlaySound Path & vl & ".wav", 0, Async
End Sub
