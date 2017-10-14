VERSION 5.00
Begin VB.Form frmFcAA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANM FcAA"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   1800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cFcAA 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdRem 
      Caption         =   "Rem"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin VB.ListBox LFcAA 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "Time2next"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   3480
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   960
      Y1              =   4800
      Y2              =   4560
   End
   Begin VB.Label Losses 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Wins 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label LastAttacked 
      Caption         =   "Not started yet... =P"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
End
Attribute VB_Name = "frmFcAA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LFcAA_Last As Integer

Private Sub cFcAA_Click()
    aFcAA = cFcAA.Value
End Sub
Private Sub cmdAdd_Click()
    LFcAA.AddItem InputBox("Skriv fighterens navn." & vbcrl & "Caps matters!", "Add fighter")
End Sub
Private Sub cmdRem_click()
    If LFcAA_Last >= 0 Then LFcAA.RemoveItem LFcAA_Last: LFcAA_Last = -1
End Sub
Private Sub cmdClear_Click()
    While LFcAA.ListCount > 0
        LFcAA.RemoveItem (0)
    Wend
End Sub
Private Sub cmdLoad_Click()
    Open "fclist.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, tmp
            LFcAA.AddItem tmp
        Wend
    Close #1
End Sub
Private Sub cmdSave_Click()
    Open "fclist.txt" For Output As #1
        For a = 0 To LFcAA.ListCount - 1
            Print #1, LFcAA.List(a)
        Next
    Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Hide
End Sub

Private Sub LFcAA_Click()
    tmp = LFcAA.Text
    For a = 0 To LFcAA.ListCount - 1
        If tmp = LFcAA.List(a) Then LFcAA_Last = a
    Next
    Me.Caption = LFcAA_Last
End Sub

Public Sub doFcAA()
    If COMPILED Then On Error Resume Next
    Nav "http://www.nordicmafia.net/nordic/index.php?side=fightclub"
    If Jail Then Exit Sub
    If AntiBot("fightclub") Then doFcAA: Exit Sub
    webs = wSRC
    For a = 0 To LFcAA.ListCount - 1
        If InStr(1, webs, LFcAA.List(a)) > 0 Then
            Target = LFcAA.List(a): LastAttacked = Target
            tsrc = Split(Split(webs, Target)(0), "type=radio value=")
            Id = Split(tsrc(UBound(tsrc)), " name=motstandervelg")(0)
            fWB.WB.Document.write ("<FORM method=""POST"" action=""http://www.nordicmafia.net/nordic/index.php?side=fightclub""><input type=radio name=motstandervelg value=""" & Id & """ checked><input type=submit name=""subutford"" value=""Utfordre!""></form>")
            DoEvents: fWB.WB.Document.All("subutford").Click: w8
            webs = wSRC: If LogIn Or AntiBot("fightclub") Then doFcAA: Exit Sub
            If InStr(1, webs, "Du var sterkere enn motstanderen din. Du vant") Then
                tFcAA = gEnd(30)
                Wins = Wins + 1
                fStatus "Vant over " & Target & " (" & Wins & "/" & Losses & ")"
                Exit Sub
            ElseIf InStr(1, webs, "Motstanderen var sterkere enn deg. Du tapte") Then
                For b = 0 To LFcAA.ListCount - 1
                    If LFcAA.List(b) = Target Then LFcAA.RemoveItem b
                Next
                Losses = Losses + 1
                tFcAA = gEnd(30)
                fStatus "Tapte mot " & Target & " (" & Wins & "/" & Losses & ")"
                Exit Sub
            Else
                tFcAA = gEnd(0)
                fStatus "Ingen mål. Prøver igjen..." & " (" & Wins & "/" & Losses & ")"
                Exit Sub
            End If
        End If
    Next
End Sub
