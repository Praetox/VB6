VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox DirBox 
      Height          =   1440
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdShowSessions 
      Caption         =   "Sessions"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4440
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private sWSit As Boolean, WsInData As String, Userdata(100) As uInf, cntUsers As Long
Private Type uInf
    Username As String
    Password As String
    Session As String
End Type

Private Sub cmdShowSessions_Click()
    For a = 0 To cntUsers
        tmp = tmp & Userdata(a).Session & vbCrLf
    Next
    MsgBox tmp
    RemCookie
    wb.Navigate "http://www.nordicmafia.net/"
    'SetUserSession 1
End Sub
Private Sub Form_Load()
    CookiePath = Environ("temp"): cpOffset = 1: tOffset = 1
    Do While tOffset > 0
        tOffset = InStr(cpOffset + 1, CookiePath, "\")
        If tOffset > 0 Then cpOffset = tOffset
    Loop
    DirBox.Path = Left(CookiePath, cpOffset - 1) & "\Temporary Internet Files\Content.IE5\"
    
    Me.Show: DoEvents: RemCookie
    Userdata(0).Username = InputBox("Username of the player to be ranked?", "User information", "ikt_societyx")
    Userdata(0).Password = InputBox("Password of the player to be ranked?", "User information", "asdfghjkløæ")
    For a = 1 To UBound(Userdata)
        Userdata(a).Username = InputBox("Username of ranker #" & a & "." & vbCrLf & vbCrLf & "Start Doomsday with current rankers by leaving field empty and hitting enter.", "User information", "postal xd")
        If Userdata(a).Username = "" Then Exit For
        Userdata(a).Password = InputBox("Password of ranker #" & a & ".", "User information", "dv1190ea")
        cntUsers = a
    Next
    If cntUsers = 0 Then MsgBox "Please enter at least one ranker. 6 (one attack every 5. second) is recommended.": ExitAPP
    
    Logout
    MsgBox "Doomsday will now open internet explorer, and navigate to Nordicmafia." & vbCrLf & vbCrLf & _
           "After internet explorer has opened nordicmafia completely, close it."
    Shell Environ("PROGRAMFILES") & "\internet explorer\iexplore.exe www.nordicmafia.net", vbMaximizedFocus
    MsgBox "When you have closed internet explorer, click OK."
    
    SetUserSession 0
    'For a = 0 To cntUsers
    '    SetUserSession a
    'Next
    
    tmp = WSit("/nordic/index.php?side=fightclub", "http://www.nordicmafia.net/nordic/index.php?side=fightclub", _
               "www.nordicmafia.net", "fcstartbelop=100&fcstartkamp=Start%21", getSession, True)
    
End Sub

Sub SetUserSession(ByVal user As Long)
    RemCookie
    Login Userdata(user).Username, Userdata(user).Password
    Userdata(user).Session = getSession
    If Userdata(user).Session = "No session" Then MsgBox "Major problem:" & vbCrLf & vbCrLf & "No cookie was found!"
End Sub
Function getSession() As String
    getSession = CookiePath 'Environ("userprofile") & "\cookies\"
    MsgBox "Result of " & CookiePath & "*www.nordicmafia.net*" & ":" & vbCrLf & dir$(CookiePath & "*www.nordicmafia.net*")
    getSession = getSession & dir$(getSession & "*www.nordicmafia.net*")
    MsgBox getSession
    If FEx(getSession) = False Then getSession = "No session": Exit Function
    Fn = FreeFile
    Open getSession For Input As #Fn
    Line Input #Fn, getSession
    Close #Fn
    getSession = Split(getSession, vbLf)(1)
End Function
Function RemCookie() As Boolean
    'Cookie = CookiePath 'Environ("userprofile") & "\cookies\"
    'Cookie = Cookie & dir$(Cookie & "*www.nordicmafia.net*")
    'If FEx(Cookie) = False Then RemCookie = False Else RemCookie = True
    'While FEx(Cookie)
    '    Kill Cookie
    '    DoEvents
    'Wend
    RemCookie = False
    flist = GetFiles
    If InStr(1, flist, "dicmaf", vbTextCompare) > 0 Then MsgBox "aye!"
    Clipboard.Clear
    Clipboard.SetText flist
    ExitAPP
            If InStr(1, file.List(b), "www.nordicmafia.net") > 1 Then
                RemCookie = True
                MsgBox file.List(b)
            End If
    Clipboard.Clear
    Clipboard.SetText lolz
    ExitAPP
End Function
Function GetFiles() As String
    For a = DirBox.ListCount - 4 To DirBox.ListCount - 1
        DirPath = DirBox.List(a) & "\"
        GetFiles = GetFiles & DirPath & dir(DirPath & "*") & vbCrLf
        Do
            tmp = dir: If tmp = "" Then Exit Do
            GetFiles = GetFiles & DirPath & tmp & vbCrLf
        Loop
    Next
End Function
Function FEx(ByVal FileName As String) As Boolean
    On Error Resume Next
    FEx = (GetAttr(FileName) And vbDirectory) = 0
End Function

Sub ExitAPP()
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next
    End
End Sub
Sub Login(ByVal user As String, Pass As String)
    wb.Navigate "http://www.nordicmafia.net/": WBW8
    wb.Document.All("brukernavn").Value = user
    wb.Document.All("passoord").Value = Pass
    wb.Document.All("submit").Click: WBW8
End Sub
Sub Logout()
    wb.Navigate "http://www.nordicmafia.net": WBW8
    src = wb.Document.Body.parentelement.innerhtml
    If InStr(1, src, "passoord") = 0 Then
        wb.Navigate "about:<form method=""post"" action=""http://www.nordicmafia.net/nordic/index.php?side=loggut""><div align=""center""><input type=""submit"" value=""Logg ut!"" name=""subloggut""><input type=""hidden"" name=""luz""></form>": DoEvents
        wb.Document.All("subloggut").Click: WBW8
    End If
End Sub
Sub WBW8()
    DoEvents
    While wb.Busy = True
        DoEvents
        Sleep (1)
    Wend
End Sub

Function WSit(ByVal PostPath As String, Referer As String, Host As String, Content As String, Session As String, Optional ReturnStr As Boolean = False) As String
    If COMPILED Then On Error GoTo hell
    tmp = "POST " & PostPath & " " & _
          "HTTP/1.1" & vbCrLf & _
          "Accept: */*" & vbCrLf & _
          "Referer: " & Referer & vbCrLf & _
          "Accept-Language: en-gb" & vbCrLf & _
          "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
          "Accept-Encoding: gzip, deflate" & vbCrLf & _
          "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.0.3705; .NET CLR 1.1.4322; Media Center PC 4.0; .NET CLR 2.0.50727)" & vbCrLf & _
          "Host: " & Host & vbCrLf & _
          "Content-Length: " & Len(Content) & vbCrLf & _
          "Connection: Keep-Alive" & vbCrLf & _
          "Cache-Control: no-cache" & vbCrLf & _
          "Cookie: PHPSESSID=" & Session & vbCrLf & _
          "" & vbCrLf & Content
    sWSit = True: WsInData = ""
    ws.Connect Host, "80"
    While sWSit
        DoEvents
        Sleep (1)
    Wend
    ws.SendData tmp
    DoEvents
    If ReturnStr = False Then
        ws.Close
        DoEvents
    Else
        While ws.State = 7
            DoEvents
            Sleep (1)
        Wend
        WSit = WsInData
    End If
    Exit Function
hell: MsgBox "Vindsokken fløy vekk."
End Function
Private Sub ws_Connect()
    sWSit = False
End Sub
Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    ws.GetData Data, vbString
    WsInData = WsInData & Data & vbCrLf
End Sub
Sub iSleep(ByVal ms As Long)
    lol = Timer * 1000
    While lol + ms > Timer * 1000
        DoEvents
        Sleep (1)
    Wend
End Sub
