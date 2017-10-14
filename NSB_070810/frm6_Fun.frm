VERSION 5.00
Begin VB.Form frm6 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "TM Fun"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Chameleon_Value 
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   360
      ItemData        =   "frm6_Fun.frx":0000
      Left            =   3720
      List            =   "frm6_Fun.frx":028E
      TabIndex        =   13
      Text            =   "Select creature..."
      Top             =   1020
      Width           =   2175
   End
   Begin VB.CheckBox Disco_Female 
      BackColor       =   &H00517362&
      Caption         =   "Alert_Battle_Logout"
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Check this if your character is female"
      Top             =   2880
      Width           =   200
   End
   Begin VB.CheckBox Disco_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   6
      ToolTipText     =   "Enable/disable the disco feature (yay)"
      Top             =   2430
      Width           =   200
   End
   Begin VB.Timer Disco_Timer 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   2640
      Top             =   3360
   End
   Begin VB.CommandButton Disco_Plus 
      BackColor       =   &H00FFC0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Change faster"
      Top             =   3195
      Width           =   255
   End
   Begin VB.CommandButton Disco_Minus 
      BackColor       =   &H00FFC0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Change slower"
      Top             =   3195
      Width           =   255
   End
   Begin VB.TextBox Name_Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   240
      Left            =   720
      TabIndex        =   2
      Text            =   "Shade of Black"
      ToolTipText     =   "What to name your character"
      Top             =   885
      Width           =   1935
   End
   Begin VB.CommandButton Name_Execute 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Execute change"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Do the change"
      Top             =   1260
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chameleon"
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
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fun =D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   9375
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Female"
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
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
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
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change char name"
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
      Height          =   240
      Left            =   480
      TabIndex        =   3
      Top             =   495
      Width           =   2415
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   8880
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disco Hack"
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
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2415
      Width           =   2415
   End
End
Attribute VB_Name = "frm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MdX As Long, MdY As Long, colr As Long, coldir As Long

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

'CHARNAME RELATED
    
    Sub Name_Execute_Click()
        If Compiled = True Then On Error Resume Next
        tmp = BL_Player
        If tmp <> 0 Then
            mWriteString tmp + BL_Name, Name_Value
            Smsg "Changed char name to " & Name_Value & "."
        End If
    End Sub

'DISCO RELATED

    Sub Disco_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        colr = 48
        Disco_Timer.Enabled = Disco_Enabled
        Smsg "Discohack enabled: " & Disco_Timer
    End Sub
    Sub Disco_Plus_Click()
        If Compiled = True Then On Error Resume Next
        If Disco_Timer.Interval > 20 Then
            Disco_Timer.Interval = Disco_Timer.Interval - 20
        Else
            If Disco_Timer.Interval > 5 Then
                Disco_Timer.Interval = Disco_Timer.Interval - 4
            End If
        End If
    End Sub
    Sub Disco_Minus_Click()
        If Compiled = True Then On Error Resume Next
        Disco_Timer.Interval = Disco_Timer.Interval + 20
    End Sub
    Sub Disco_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Dim pck(9) As Byte
        pck(0) = &H8
        pck(1) = &H0
        pck(2) = &HD3
        If Disco_Female Then pck(3) = &H88 Else pck(3) = &H80
        pck(4) = &H0
        pck(5) = (colr - 0)
        pck(6) = (colr - 1)
        pck(7) = (colr - 2)
        pck(8) = (colr - 3)
        pck(9) = &H0
        sPck pck
        colr = colr + coldir
        If colr > 132 Then coldir = -1
        If colr < 50 Then coldir = 1
    End Sub

'CHAMELEON RELATED
    
    Private Sub Chameleon_Value_Click()
        mWriteLong BL_Player + BL_OFit, Chameleon_Value.ItemData(Chameleon_Value.ListIndex)
        Smsg "Morphed into " & Chameleon_Value.List(Chameleon_Value.ListIndex) & "!"
    End Sub
