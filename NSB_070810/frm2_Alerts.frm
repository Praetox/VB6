VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "TM Alerts"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Heal_Delay 
      Alignment       =   2  'Center
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   240
      Left            =   8400
      TabIndex        =   40
      Text            =   "700"
      ToolTipText     =   "The limit for mana"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CheckBox Heal_Spam 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   6840
      TabIndex        =   37
      ToolTipText     =   "Enable/disable the automatic healer / manawaster feature"
      Top             =   1680
      Width           =   200
   End
   Begin VB.Timer Heal_Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8880
      Top             =   1440
   End
   Begin VB.Timer Mana_Timer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5760
      Top             =   3360
   End
   Begin VB.Timer HP_Timer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5760
      Top             =   1440
   End
   Begin VB.Timer Battle_Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   3360
   End
   Begin VB.Timer Move_Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   1440
   End
   Begin VB.CheckBox Mana_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   19
      ToolTipText     =   "Enable/disable the alert when mana is greater or less than..."
      Top             =   2430
      Width           =   200
   End
   Begin VB.CheckBox Battle_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   18
      ToolTipText     =   "Enable/disable the alert at battlelist feature"
      Top             =   2430
      Width           =   200
   End
   Begin VB.CheckBox Move_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   17
      ToolTipText     =   "Enable/disable the alert at move feature"
      Top             =   510
      Width           =   200
   End
   Begin VB.CheckBox HP_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   16
      ToolTipText     =   "Enable/disable the alert when hp is greater or less than..."
      Top             =   510
      Width           =   200
   End
   Begin VB.TextBox Heal_HP 
      Alignment       =   2  'Center
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   240
      Left            =   7800
      TabIndex        =   15
      ToolTipText     =   "The limit for hp"
      Top             =   825
      Width           =   735
   End
   Begin VB.TextBox Heal_Mana 
      Alignment       =   2  'Center
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   240
      Left            =   7800
      TabIndex        =   14
      ToolTipText     =   "The limit for mana"
      Top             =   1065
      Width           =   735
   End
   Begin VB.CheckBox Heal_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   6750
      TabIndex        =   13
      ToolTipText     =   "Enable/disable the automatic healer / manawaster feature"
      Top             =   510
      Width           =   200
   End
   Begin VB.CheckBox Mana_More 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3840
      TabIndex        =   12
      ToolTipText     =   "Alerts when mana is higher than the entered value"
      Top             =   3180
      Width           =   200
   End
   Begin VB.CheckBox Mana_Less 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3840
      TabIndex        =   11
      ToolTipText     =   "Alerts when mana is lower than the entered value"
      Top             =   2880
      Width           =   200
   End
   Begin VB.CheckBox HP_More 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3840
      TabIndex        =   10
      ToolTipText     =   "Alerts when hp is higher than the entered value"
      Top             =   1260
      Width           =   200
   End
   Begin VB.CheckBox HP_Less 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3840
      TabIndex        =   9
      ToolTipText     =   "Alerts when hp is lower than the entered value"
      Top             =   960
      Width           =   200
   End
   Begin VB.CheckBox Battle_Logout 
      BackColor       =   &H00517362&
      Caption         =   "Alert_Battle_Logout"
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   "Wether to logout the char when someone appear on-screen"
      Top             =   2820
      Width           =   200
   End
   Begin VB.CheckBox Move_Logout 
      BackColor       =   &H00517362&
      Caption         =   "Alert_Battle_Logout"
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Wether to logout the char when you are moved"
      Top             =   1080
      Width           =   200
   End
   Begin VB.CommandButton Battle_Safelist 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Safelist"
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
      TabIndex        =   6
      ToolTipText     =   "Modify the whitelist (including monsters)"
      Top             =   3180
      Width           =   1695
   End
   Begin VB.ComboBox Heal_Method 
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
      Height          =   330
      ItemData        =   "frm2_Alerts.frx":0000
      Left            =   7800
      List            =   "frm2_Alerts.frx":000D
      TabIndex        =   5
      Text            =   "UH"
      ToolTipText     =   "Healing method"
      Top             =   1290
      Width           =   735
   End
   Begin VB.TextBox Mana_Value 
      Alignment       =   2  'Center
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   240
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Mana limit for alert"
      Top             =   2985
      Width           =   615
   End
   Begin VB.TextBox HP_Value 
      Alignment       =   2  'Center
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   240
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Health limit for alert"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox Safezone_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   6750
      TabIndex        =   2
      ToolTipText     =   "Enable/disable walking to safezone at failed logout"
      Top             =   2430
      Width           =   200
   End
   Begin VB.TextBox Safezone_Value 
      Alignment       =   2  'Center
      BackColor       =   &H00151500&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   240
      Left            =   6840
      TabIndex        =   1
      ToolTipText     =   "The safezone position"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Safezone_Set 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Set"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Set safezone position to current location"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Delay  "
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
      Left            =   7680
      TabIndex        =   39
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Spam"
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
      Left            =   7080
      TabIndex        =   38
      Top             =   1680
      Width           =   735
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
      TabIndex        =   36
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   8880
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape3 
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   9375
   End
   Begin VB.Shape Shape2 
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autoheal"
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
      Left            =   6720
      TabIndex        =   35
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alert when mana..."
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
      TabIndex        =   34
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alert at battlelist"
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
      TabIndex        =   33
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alert at move"
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
      TabIndex        =   32
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alert when HP..."
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
      TabIndex        =   31
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alert related features"
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
      TabIndex        =   30
      Top             =   1920
      Width           =   9375
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Method"
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
      Left            =   7080
      TabIndex        =   29
      Top             =   1335
      Width           =   615
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mana"
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
      Left            =   7200
      TabIndex        =   28
      Top             =   1065
      Width           =   495
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
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
      Left            =   7200
      TabIndex        =   27
      Top             =   825
      Width           =   495
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Higher than"
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
      Left            =   4080
      TabIndex        =   26
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Lower than"
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
      Left            =   4080
      TabIndex        =   25
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Higher than"
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
      Left            =   4080
      TabIndex        =   24
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Lower than"
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
      Left            =   4080
      TabIndex        =   23
      Top             =   960
      Width           =   975
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Logout"
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
      TabIndex        =   22
      Top             =   2820
      Width           =   735
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Logout"
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
      TabIndex        =   21
      Top             =   1080
      Width           =   735
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   9000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape19 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Walk at failed log"
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
      Left            =   6720
      TabIndex        =   20
      Top             =   2415
      Width           =   2415
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MdX As Long, MdY As Long, LastX As Long, LastY As Long, LastZ As Long, oHP As Long, oMA As Long
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

'ALERT AT MOVE RELATED

    Sub Move_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        LastX = mReadLong(CH_X)
        LastY = mReadLong(CH_Y)
        LastZ = mReadLong(CH_Z)
        Move_Timer = Move_Enabled
        Move_Timer_Timer
        Smsg "Alert when moved from " & LastX & "-" & LastY & "-" & LastZ & " enabled: " & Move_Timer
    End Sub
    Sub Move_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        nX = mReadLong(CH_X)
        nY = mReadLong(CH_Y)
        nZ = mReadLong(CH_Z)
        If nX <> LastX Or nY <> LastY Or nZ <> LastZ Then
            doAlert = True
            If chkMoveLog Then attemptlog
        End If
        LastX = mReadLong(CH_X)
        LastY = mReadLong(CH_Y)
        LastZ = mReadLong(CH_Z)
    End Sub
    
'ALERT AT BATTLELIST RELATED

    Sub Battle_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Battle_Timer = Battle_Enabled
        Battle_Timer_Timer
        Smsg "Alert at battlelist enabled: " & Battle_Timer
    End Sub
    Sub Battle_Safelist_Click()
        If Compiled = True Then On Error Resume Next
        frmWhitelist.Show
    End Sub
    Sub Battle_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        nZ = mReadLong(CH_Z)
        cID = mReadLong(CH_ID)
        For a = BL_Start To BL_End Step BL_Dist
            If mReadLong(a + BL_Vis) <> 0 Then
                If mReadLong(a + BL_Z) = nZ Then
                    If mReadLong(a + BL_ID) <> cID Then
                        If InStr(1, frmWhitelist.txt, mReadString(a + BL_Name), vbTextCompare) = 0 Then
                            doAlert = True
                            Smsg mReadString(a + BL_Name) & " (" & mReadLong(a + BL_ID) & ") entered the screen."
                            If Battle_Logout Then attemptlog
                        End If
                    End If
                End If
            End If
        Next
    End Sub

'ALERT AT HP/MANA RELATED

    Sub HP_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        HP_Timer = HP_Enabled
        HP_Timer_Timer
        Smsg "Alert at low/high health enabled: " & HP_Timer
    End Sub
    Sub HP_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        If HP_Less Then If mReadLong(a + CH_HP) < Int(HP_Value) Then doAlert = True
        If HP_More Then If mReadLong(a + CH_HP) > Int(HP_Value) Then doAlert = True
    End Sub
    Sub Mana_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Mana_Timer = Mana_Enabled
        Mana_Timer_Timer
        Smsg "Alert at low/high mana enabled: " & Mana_Timer
    End Sub
    Sub Mana_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        If Mana_Less Then If mReadLong(a + CH_Ma) < Int(Mana_Value) Then doAlert = True
        If Mana_More Then If mReadLong(a + CH_Ma) > Int(Mana_Value) Then doAlert = True
    End Sub

'AUTOHEAL RELATED

    Sub Heal_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Heal_Timer = Heal_Enabled
        Heal_Timer_Timer
        Smsg "Healer/waster enabled: " & Heal_Timer
    End Sub
    Sub Heal_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Dim HealID As Long
        cHP = mReadLong(CH_HP)
        cMA = mReadLong(CH_Ma)
        If Heal_Spam Then If ((Timer * 1000) - LastHealRuneFire) < Heal_Delay Then Exit Sub
        If Int(cHP) <> Int(oHP) Or Int(cMA) <> Int(oMA) Then
            If Int(cHP) <= Int(Heal_HP) Then
                If Int(cMA) >= Int(Heal_Mana) Then
                    If Heal_Method = "IH" Then
                        HealID = Ru_IH
                    ElseIf Heal_Method = "UH" Then
                        HealID = Ru_UH
                    ElseIf Heal_Method = "Hotkey F12" Then
                        Call SendMessage(tHvnd, WM_KD, 123, vbNull)
                        Call SendMessage(tHvnd, WM_KU, 123, vbNull)
                        Smsg "Pressed Tibia hotkey F12."
                    End If
                    If HealID <> 0 Then
                        Dim Cont As Long, iPos As Long, gX As Long, gY As Long, gZ As Long
                        gX = mReadLong(CH_X)
                        gY = mReadLong(CH_Y)
                        gZ = mReadLong(CH_Z)
                        Cont = cWithItem(HealID)
                        iPos = iPosInCont(HealID, Cont)
                        use_WithCont HealID, gX, gY, gZ, iPos, Cont
                        Smsg "Using " & Heal_Method & " rune from backpack " & Cont & " in slot " & iPos & "."
                    End If
                End If
            End If
            If Heal_Spam Then cHP = 0: cMA = 0: LastHealRuneFire = Timer * 1000
            oHP = cHP: oMA = cMA
        End If
    End Sub
    Sub Heal_Key_KeyUp(KeyCode As Integer, Shift As Integer)
        If Compiled = True Then On Error Resume Next
        Heal_Key = KeyCode
    End Sub

'ATTEMPT LOGOUT RELATED

    Sub attemptlog()
        If Compiled = True Then On Error Resume Next
        Call Logout: DoEvents: iSleep 5000
        If mReadLong(CH_Con) > 0 And Safezone_Enabled = 1 Then
            GotoSafe (Safezone_Value)
        End If
    End Sub
    Sub safezone_set_click()
        If Compiled = True Then On Error Resume Next
        Safezone_Value = mReadLong(CH_X) & "," & mReadLong(CH_Y) & "," & mReadLong(CH_Z)
    End Sub
