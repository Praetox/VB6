VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form fWB 
   Caption         =   "AutoNM :: Nettleser"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Left            =   120
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   17
      Top             =   120
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   8700
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   10155
      ExtentX         =   17912
      ExtentY         =   15346
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
      Location        =   ""
   End
   Begin VB.Label logs 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   16
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label LUsed 
      Height          =   255
      Left            =   3405
      TabIndex        =   15
      Top             =   900
      Width           =   735
   End
   Begin VB.Label LStart 
      Height          =   255
      Left            =   3400
      TabIndex        =   14
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LEnd 
      Height          =   255
      Left            =   3405
      TabIndex        =   13
      Top             =   540
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Used"
      Height          =   255
      Left            =   2940
      TabIndex        =   12
      Top             =   900
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   2940
      TabIndex        =   11
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Start"
      Height          =   255
      Left            =   2940
      TabIndex        =   10
      Top             =   180
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   1680
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line4 
      X1              =   2400
      X2              =   2400
      Y1              =   120
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   2040
      X2              =   2040
      Y1              =   120
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   1680
      X2              =   2760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   2760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2400
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label IsCar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "fWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Session As String, Treshold As Long, CurABot As String
Sub DownloadFile(ByVal file As String, path As String)
    Call URLDownloadToFile(0, file, path, 0, 0)
End Sub

Function BreakTheAntibot(ByVal vl As String) As String
    Dim lstColor As String
    Open "cars.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, tmp
        If Left(tmp, 1) = "A" Then lstColor = lstColor & Mid$(tmp, 2) & ","
        If Left(tmp, 1) = "B" Then lstImg1 = lstImg1 & Mid$(tmp, 2) & "#"
        If Left(tmp, 1) = "C" Then lstImg2 = lstImg2 & Mid$(tmp, 2) & "#"
    Wend
    Close #1
    lstColor = Left(lstColor, Len(lstColor) - 1)
    lstColor = Replace(lstColor, "f", "1")
    lstColor = Replace(lstColor, "d", "2")
    lstColor = Replace(lstColor, "b", "3")
    lstColor = Replace(lstColor, "e", "5")
    lstColor = Replace(lstColor, "a", "6")
    lstColor = Replace(lstColor, "c", "9")
    lstImg1 = Left(lstImg1, Len(lstImg1) - 1)
    lstImg2 = Left(lstImg2, Len(lstImg2) - 1)
    lstImg1 = Split(lstImg1, "#")
    lstImg2 = Split(lstImg2, "#")
    
    CurABot = vl
    For a = 0 To 8
        IsCar(a) = ""
    Next
    DoEvents
    LStart = Timer
    Do
        For a = 0 To 8
            If IsCar(a) = "" Then
                logs = "IMG" & a: DoEvents
                SetPic ("http://www.nordicmafia.net/nordic/hget_genbilde.php?id=" & a & "&" & Session)
                tmp = verPicM(lstColor)
                'If tmp <> "," Then FileCopy "c:\lol.jpg", "c:\" & a + 1 & "_" & tmp & "_" & Timer & ".jpg"
                For J = 0 To UBound(lstImg1)
                    If InStr(1, tmp, lstImg1(J)) > 0 Then IsCar(a) = "o" & J
                Next
                For J = 0 To UBound(lstImg2)
                    If InStr(1, tmp, lstImg2(J)) > 0 Then IsCar(a) = "x" & J
                Next
            End If
        Next
        carsend = "": blocksz = 0
        For a = 0 To 8
            If Left(IsCar(a), 1) = "o" Then carsend = carsend & (a + 1) & " "
            If Left(IsCar(a), 1) = "x" Then blocksz = blocksz + 1
            If Len(carsend) > 5 Then GoTo 10
        Next
        If blocksz > 6 Then
            For a = 0 To 8
                If Left(IsCar(a), 1) = "x" Then IsCar(a) = ""
            Next
        End If
    Loop
10  LEnd = Timer: LUsed = (Int(LEnd) - Int(LStart) \ 1) & "s": DoEvents
    logs.Caption = carsend
    BreakTheAntibot = Replace(carsend, " ", "")
End Function

Private Sub SetPic(ByVal addy As String)
    On Error GoTo gokk
    DownloadFile addy, "lol.jpg": DoEvents
    pic.Picture = LoadPicture("lol.jpg")
    DoEvents
    Exit Sub
gokk: Nav "http://www.nordicmafia.net/nordic/index.php?side=" & CurABot
      Session = Split(Split(wSRC, "genbilde.php?id=0&")(1), """")(0)
      For a = 0 To 8
          IsCar(a) = ""
      Next
End Sub

Private Sub cLong(ByVal vl As Long, cR As Integer, cG As Integer, cB As Integer)
    cR = vl Mod &H100
    vl = vl \ &H100
    cG = vl Mod &H100
    vl = vl \ &H100
    cB = vl Mod &H100
End Sub

Private Function cIdent(ByVal cl1 As Long, cl2 As Long) As Boolean
    Dim c1 As Integer, c2 As Integer, c3 As Integer, c4 As Integer, c5 As Integer, c6 As Integer
    cLong cl1, c1, c3, c5
    cLong cl2, c2, c4, c6
    cx = c2 - c1
    cy = c4 - c3
    cz = c6 - c5
    If cx < 0 Then cx = -cx
    If cy < 0 Then cy = -cy
    If cz < 0 Then cz = -cz
    If cx < Treshold Then
        If cy < Treshold Then
            If cz < Treshold Then
                cIdent = True
            End If
        End If
    End If
End Function

Private Function verPicM(cl As String) As String
    vl = Split(cl, ",")
    Dim aryColor() As Long, aryMinVal() As Long, aryMaxVal() As Long, aryThres() As Long, pix() As Long
    ReDim aryColor(UBound(vl)), aryMinVal(UBound(vl)), aryMaxVal(UBound(vl)), aryThres(UBound(vl)), pix(UBound(vl))
    For a = 0 To UBound(vl)
        aryTmp = Split(vl(a), "x")
        aryColor(a) = aryTmp(0)
        aryVals = Split(aryTmp(1), "-")
        aryMinVal(a) = aryVals(0)
        aryMaxVal(a) = aryVals(1)
        aryThres(a) = aryTmp(2)
    Next
    
    For Y = 0 To 65 Step 1
        For X = 0 To 90 Step 1
            ThisCol = pic.Point(X, Y)
            For N = 0 To UBound(vl)
                Treshold = aryThres(N)
                If cIdent(ThisCol, aryColor(N)) Then pix(N) = pix(N) + 1
            Next
        Next
    Next
    For a = 0 To UBound(vl)
        CMinCol = Split(Split(vl(a), "x")(1), "-")(0)
        CMaxCol = Split(Split(vl(a), "x")(1), "-")(1)
        If (Int(pix(a)) >= Int(CMinCol)) And (Int(pix(a)) <= Int(CMaxCol)) Then
            passed = passed & a & ","
        End If
    Next
    verPicM = "," & passed
End Function













Private Sub Form_Load()
    DownloadFile SiteCars, "cars.txt"
End Sub

Private Sub Form_Resize()
    If COMPILED Then On Error Resume Next
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        WB.Width = Me.Width - 120: WB.Height = Me.Height - 675
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If COMPILED Then On Error Resume Next
    Cancel = 1
    Me.Hide
    Showing = False
End Sub
Private Sub WB_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    sW8d = True
End Sub
Private Function EncCols(ByVal vl As String) As String
    vl = Replace(vl, "1", "f")
    vl = Replace(vl, "2", "d")
    vl = Replace(vl, "3", "b")
    vl = Replace(vl, "5", "e")
    vl = Replace(vl, "6", "a")
    vl = Replace(vl, "9", "c")
    EncCols = vl
End Function
Private Function DecCols(ByVal vl As String) As String
    vl = Replace(vl, "f", "1")
    vl = Replace(vl, "d", "2")
    vl = Replace(vl, "b", "3")
    vl = Replace(vl, "e", "5")
    vl = Replace(vl, "a", "6")
    vl = Replace(vl, "c", "9")
    DecCols = vl
End Function
