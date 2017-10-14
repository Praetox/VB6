VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GetKols!"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Add to box"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton SetLimit 
      Caption         =   "Set Limit"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LPix Decoder"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Init. Analyse"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox p2 
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txcars 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
   End
   Begin VB.PictureBox img 
      Height          =   1095
      Left            =   120
      Picture         =   "getcols.frx":0000
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label limit 
      Alignment       =   1  'Right Justify
      Caption         =   "10000"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Lastclick 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cols(6005) As Long, popularity(6005) As Long, recpop As Long
Private Type XY
    X As Long
    Y As Long
End Type
Const Treshold = 10

Private Sub Command1_Click()
    txcars = txcars & Me.Caption & ","
End Sub

Private Sub Command2_Click()
    targt = InputBox("What pixel are you looking for?", "Long pixel decoder")
    MsgBox LPIXDEC(targt).X & "x" & LPIXDEC(targt).Y
End Sub
Private Function LPIXDEC(ByVal longpix As Long) As XY
    For cy = 0 To 65
        For cx = 0 To 90
            If Int(nm) = Int(longpix) Then LPIXDEC.X = cx: LPIXDEC.Y = cy
            nm = nm + 1
        Next
    Next
End Function

Private Sub Command3_Click()
    limit = 1000: Dim RecPixCol As Long
    Me.Show: DoEvents: DoEvents
    For cy = 0 To 65
        For cx = 0 To 90
            cols(nm) = img.Point(cx, cy)
            nm = nm + 1
        Next
        Me.Caption = "Setting array " & cy & "/65"
        DoEvents
    Next
    For a = 0 To 6005
        curpop = 0
        For b = 0 To 6005
            If cIdent(cols(a), cols(b)) Then curpop = curpop + 1
        Next
        If curpop > recpop Then
            recpop = curpop: recown = a
            If recpop > Int(limit) Then
                txcars = txcars & recown & " (" & LPIXDEC(recown).X & "x" & LPIXDEC(recown).Y & ") with " & recpop & " equals" & vbCrLf
                Call P2CLS: RecPixCol = img.Point(LPIXDEC(recown).X, LPIXDEC(recown).Y)
                For cy = 0 To 65
                    For cx = 0 To 90
                        If cIdent(img.Point(cx, cy), RecPixCol) Then p2.PSet (X, Y)
                    Next
                Next
            End If
        End If
        If Right(a, 1) = "0" Then
            Me.Caption = a & " (" & (((100 / 6005) * a) \ 1) & "%) Rec: " & recown & "/" & recpop
            DoEvents
        End If
    Next
    MsgBox "Most popular pixel is " & recown & ", with " & recpop & " equals."
End Sub

Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = X & "x" & Y
End Sub

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.BackColor = img.Point(X, Y)
    P2CLS
    For cx = 0 To 90
        For cy = 0 To 65
            If cIdent(img.Point(cx, cy), Me.BackColor) Then enums = enums + 1: p2.PSet (cx, cy)
        Next
    Next
    Lastclick = Hex(Me.BackColor) & " - " & Me.BackColor & ", " & enums
    Clipboard.Clear
    Clipboard.SetText Lastclick
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
    If cx < Treshold And cy < Treshold And cz < Treshold Then cIdent = True
End Function

Private Sub SetLimit_Click()
    limit = recpop
End Sub

Private Sub P2CLS()
    'p2.Picture = img.Picture
    p2.Cls
End Sub
