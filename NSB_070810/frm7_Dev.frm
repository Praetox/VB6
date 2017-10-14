VERSION 5.00
Begin VB.Form frm7 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox DEV1_VL1 
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "Item ID"
      Top             =   900
      Width           =   1455
   End
   Begin VB.TextBox DEV1_VL2 
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "Result of scan"
      Top             =   1260
      Width           =   2175
   End
   Begin VB.CommandButton DEV1_Execute 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   900
      Width           =   615
   End
   Begin VB.Label cmdClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   8880
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Shape Shape9 
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   9375
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape22 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dev :: Item Lookup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   495
      Width           =   2415
   End
End
Attribute VB_Name = "frm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MdX As Long, MdY As Long

Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Compiled = True Then On Error Resume Next
    If Button = 1 Then MdX = X: MdY = Y
End Sub
Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Compiled = True Then On Error Resume Next
    If Button = 1 Then Me.Move (Me.Left + X) - MdX, (Me.Top + Y) - MdY
    frmMain.Move Me.Left, Me.Top
End Sub
Sub cmdClose_Click()
    Me.Hide: frmMain.Show
End Sub
Sub Form_Load()
    Me.Picture = frmMain.Picture
End Sub

' ITEM LOOKUP RELATED

    Sub DEV1_VL1_click()
        If Compiled = True Then On Error Resume Next
        DEV1_VL1 = mReadLong(Look_ID)
    End Sub
    Sub DEV1_Execute_Click()
        If Compiled = True Then On Error Resume Next
        Dim Cont As Long, iPos As Long
        Cont = cWithItem(Int(DEV1_VL1))
        iPos = iPosInCont(Int(DEV1_VL1), Cont)
        DEV1_VL2 = "BP" & Cont & " SLOT" & iPos
    End Sub

