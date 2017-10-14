VERSION 5.00
Begin VB.Form frmSok1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Søker"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Søk!"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox txtList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmSok1.frx":0000
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblCurrent 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1695
   End
End
Attribute VB_Name = "frmSok1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGo_Click()
    If COMPILED Then On Error Resume Next
    If intMain = True Then MsgBox "Vennligst vent til boten er ferdig med å utføre andre ting først.": Exit Sub
    frmMain.tMain.Enabled = False
    fStatus "Starter søkefunksjon..."
    soktimer = InputBox("I hvor mange timer ønsker du å søke etter personen(e)?", "Søkebot", "4")
    Dim SokL(5) As String: SokL(0) = "Oslo": SokL(1) = "Stockholm": SokL(2) = "København"
                           SokL(3) = "Helsinki": SokL(4) = "London": SokL(5) = "Moskva"
    Targets = Split(txtList, vbCrLf)
    For I = 0 To UBound(Targets)
        If Targets(I) <> "" Then
            SokSpiller = Targets(I)
            lblCurrent = SokSpiller
            For a = 0 To UBound(SokL)
                fStatus "Søker etter " & SokSpiller & " i " & SokL(a) & " (" & soktimer & " timer)..."
                Nav ("http://www.nordicmafia.net/nordic/index.php?side=drep")
                fWB.WB.Document.All("finnspiller").Value = SokSpiller
                fWB.WB.Document.All("finnsted").Value = SokL(a)
                fWB.WB.Document.All("antalltimer").Value = soktimer
                fWB.WB.Document.All("findsubmit").Click: w8
            Next
            fStatus "Søkte etter " & SokSpiller & " i " & soktimer & " timer."
        End If
    Next
    fStatus "Alle søk fullført."
    frmMain.tMain.Enabled = True
    Me.Hide
End Sub
