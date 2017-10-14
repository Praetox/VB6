VERSION 5.00
Begin VB.Form frm4 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "TM Cavebot"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Loot_Delay 
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
      Left            =   8340
      TabIndex        =   50
      Text            =   "500"
      Top             =   3220
      Width           =   495
   End
   Begin VB.CommandButton Loot_Save 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Click here to add waypoint"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Loot_Load 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Click here to add waypoint"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Loot_Add 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Click here to add waypoint"
      Top             =   3000
      Width           =   615
   End
   Begin VB.Timer Loot_Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8880
      Top             =   3360
   End
   Begin VB.TextBox Loot_Eat 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   7200
      TabIndex        =   44
      Text            =   "x"
      ToolTipText     =   "The monsters to attack"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Loot_BP 
      Alignment       =   2  'Center
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
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   42
      Text            =   "x"
      ToolTipText     =   "The monsters to attack"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Loot_BP 
      Alignment       =   2  'Center
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
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   40
      Text            =   "x"
      ToolTipText     =   "The monsters to attack"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Loot_BP 
      Alignment       =   2  'Center
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
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   38
      Text            =   "x"
      ToolTipText     =   "The monsters to attack"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Loot_BP 
      Alignment       =   2  'Center
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
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   36
      Text            =   "x"
      ToolTipText     =   "The monsters to attack"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Loot_BP 
      Alignment       =   2  'Center
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
      Height          =   255
      Index           =   0
      Left            =   7200
      TabIndex        =   34
      Text            =   "x"
      ToolTipText     =   "The monsters to attack"
      Top             =   840
      Width           =   1815
   End
   Begin VB.CheckBox Loot_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   6750
      TabIndex        =   32
      ToolTipText     =   "Enable/disable the cavebots autowalk feature"
      Top             =   510
      Width           =   200
   End
   Begin VB.TextBox Attack_MaxY 
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
      Left            =   2160
      TabIndex        =   31
      Text            =   "5"
      Top             =   1420
      Width           =   375
   End
   Begin VB.TextBox Attack_MaxX 
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
      Left            =   1680
      TabIndex        =   30
      Text            =   "7"
      Top             =   1420
      Width           =   375
   End
   Begin VB.ListBox Walk_Value 
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
      Height          =   2205
      Left            =   3840
      TabIndex        =   21
      Top             =   780
      Width           =   1935
   End
   Begin VB.CommandButton Walk_Clr 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clear"
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
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Click here to add waypoint"
      Top             =   3015
      Width           =   495
   End
   Begin VB.CheckBox Walk_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   19
      ToolTipText     =   "Enable/disable the cavebots autowalk feature"
      Top             =   510
      Width           =   200
   End
   Begin VB.Timer Walk_Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   1380
   End
   Begin VB.CommandButton Walk_Add 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Click here to add waypoint"
      Top             =   3015
      Width           =   495
   End
   Begin VB.CheckBox WalkAa_Enabled 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00151500&
      Caption         =   "Add"
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
      Left            =   5295
      TabIndex        =   17
      ToolTipText     =   "Check this to add waypoints using the Delete button on your keyboard."
      Top             =   3135
      Width           =   615
   End
   Begin VB.CheckBox Attack_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   16
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   510
      Width           =   200
   End
   Begin VB.Timer Attack_Timer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2655
      Top             =   1455
   End
   Begin VB.TextBox Attack_Value 
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
      Height          =   615
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frm4_Cavebot.frx":0000
      ToolTipText     =   "The monsters to attack"
      Top             =   780
      Width           =   2175
   End
   Begin VB.CommandButton Walk_Rem 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rem"
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
      Left            =   4215
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Click here to add waypoint"
      Top             =   3015
      Width           =   495
   End
   Begin VB.CommandButton Walk_Load 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Load"
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
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Click here to add waypoint"
      Top             =   3255
      Width           =   735
   End
   Begin VB.CommandButton Walk_Save 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click here to add waypoint"
      Top             =   3255
      Width           =   735
   End
   Begin VB.OptionButton Walk_SpType 
      BackColor       =   &H00151500&
      Caption         =   "Option1"
      Height          =   200
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      ToolTipText     =   "Use this option to go up a rope-hole."
      Top             =   2760
      Width           =   200
   End
   Begin VB.OptionButton Walk_SpType 
      BackColor       =   &H00151500&
      Caption         =   "Option1"
      Enabled         =   0   'False
      Height          =   200
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      ToolTipText     =   "Use this option to go up a ladder."
      Top             =   3030
      Width           =   200
   End
   Begin VB.OptionButton Walk_SpType 
      BackColor       =   &H00151500&
      Caption         =   "Option1"
      Height          =   200
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      ToolTipText     =   "Use this option to add a hole, a ladder going down, or stairs up/down."
      Top             =   3315
      Width           =   200
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Execute the action selected to the right on the tile to the top left"
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   1
      Left            =   975
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Execute the action selected to the right on the tile to the top center"
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   2
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Execute the action selected to the right on the tile to the top right"
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   3
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Execute the action selected to the right on the tile to the middle left"
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   4
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Execute the action selected to the right on the tile you're currently standing on"
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   5
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Execute the action selected to the right on the tile to the middle right"
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   6
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Execute the action selected to the right on the tile to the bottom left"
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   7
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Execute the action selected to the right on the tile to the bottom center"
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Walk_Special 
      BackColor       =   &H00FFC0C0&
      Caption         =   "!"
      Height          =   255
      Index           =   8
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Execute the action selected to the right on the tile to the bottom right"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PDelay"
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
      Left            =   8280
      TabIndex        =   49
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   45
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BP5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   43
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BP4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   41
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BP3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   39
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BP2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   37
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BP1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   35
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Looting"
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
      TabIndex        =   33
      Top             =   495
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   3135
      Left            =   6720
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Max dist:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   29
      ToolTipText     =   "Use this option to go up a rope-hole."
      Top             =   1440
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
      TabIndex        =   28
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
   Begin VB.Shape Shape9 
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   9375
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H00FFFFFF&
      Height          =   3135
      Left            =   3600
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autoattack"
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
      TabIndex        =   26
      Top             =   495
      Width           =   2415
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   480
      Width           =   2415
   End
   Begin VB.Shape Shape24 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label60 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Rope"
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
      Left            =   1950
      TabIndex        =   24
      ToolTipText     =   "Use this option to go up a rope-hole."
      Top             =   2745
      Width           =   735
   End
   Begin VB.Label Label61 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
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
      Left            =   1950
      TabIndex        =   23
      ToolTipText     =   "Use this option to go up a ladder."
      Top             =   3015
      Width           =   735
   End
   Begin VB.Label Label62 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Walk to"
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
      Left            =   1950
      TabIndex        =   22
      ToolTipText     =   "Use this option to add a hole, a ladder going down, or stairs up/down."
      Top             =   3300
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add special tile"
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
      TabIndex        =   27
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autowalk"
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
      TabIndex        =   25
      Top             =   495
      Width           =   2415
   End
End
Attribute VB_Name = "frm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MdX As Long, MdY As Long, CBPos As Long, AtkId As Long, wasWalking As Boolean, laddlasX As Long, _
        laddlasY As Long, laddlasZ As Long, nwtX As Long, nwtY As Long, nwtZ As Long, wtTick As Integer, _
        CBS_LoadedBefore As Boolean

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

'AUTOATTACK RELATED

    Sub Attack_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Attack_Timer.Enabled = Attack_Enabled
        Smsg "Autoattacker enabled: " & Attack_Timer
        wasWalking = Walk_Timer
    End Sub
    Sub Attack_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Dim nX As Long, nY As Long, nZ As Long, cID As Long, mX As Long, mY As Long, mZ As Long, ThisDist As Long, _
            MaxDist As Long, RelPos As Long, a As Long, b As Long, c As Long, d As Long
        nX = mReadLong(CH_X)
        nY = mReadLong(CH_Y)
        nZ = mReadLong(CH_Z)
        cID = mReadLong(CH_ID)
        AtkNew = True
        For a = BL_Start To BL_End Step BL_Dist
            If mReadLong(a + BL_ID) = AtkId And AtkId <> 0 Then
                Label35 = "Found mob"
                If mReadLong(a + BL_HP) > 0 Then
                    Label35 = "Mob is alive"
                    If mReadLong(a + BL_Vis) > 0 Then
                        Label35 = "Mob is visible"
                        tmpx = mReadLong(a + BL_X) - mReadLong(CH_X)
                        tmpy = mReadLong(a + BL_Y) - mReadLong(CH_Y)
                        tmpz = mReadLong(a + BL_Z) - mReadLong(CH_Z)
                        If tmpx < 0 Then tmpx = -tmpx
                        If tmpy < 0 Then tmpy = -tmpy
                        If tmpx <= 1 And tmpy <= 1 And tmpz = 0 Then
                            Label35 = "Mob is within reach"
                            AtkNew = False
                        End If
                    End If
                End If
                Label35 = Right(Timer, 2) & " " & Label35
                GoTo 10
            End If
        Next
10      If AtkNew Then
            AtkId = 0
            MaxDist = 9999
            For a = BL_Start To BL_End Step BL_Dist
                If mReadLong(a + BL_Vis) <> 0 Then
                    If mReadLong(a + BL_Z) = nZ Then
                        If mReadLong(a + BL_HP) > 0 Then
                            If InStr(1, Attack_Value, mReadString(a + BL_Name), vbTextCompare) > 0 Then
                                moID = mReadLong(a + BL_ID)
                                If moID <> cID Then
                                    mX = (mReadLong(a + BL_X) - nX)
                                    mY = (mReadLong(a + BL_Y) - nY)
                                    If mX < 0 Then mX = -mX
                                    If mY < 0 Then mY = -mY
                                    ThisDist = mX + mY
                                    If ThisDist < MaxDist Then
                                        If mX <= Attack_MaxX And mY <= Attack_MaxY Then
                                            MaxDist = ThisDist
                                            'If mReadLong(a + BL_HP) < 100 Then MaxDist = 0
                                            AtkId = moID
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If AtkId <> 0 Then
                Walk_Timer.Enabled = False
                mWriteLong CH_gX, 0: mWriteLong CH_gY, 0: mWriteLong CH_gZ, 0
                Attack AtkId
                iSleep 500
                aFollow
                iSleep 1000
            Else
                If Walk_Timer.Enabled <> wasWalking Then
                    Walk_Timer.Enabled = wasWalking
                    If wasWalking = True Then GotoXYZ nwtX, nwtY, nwtZ
                End If
            End If
        End If
    End Sub

'AUTOWALK RELATED

    Sub Walk_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Walk_Timer.Enabled = Walk_Enabled
        wasWalking = Walk_Timer.Enabled
        Smsg "Autowalker enabled: " & Walk_Timer
        Label34 = "Autowalk"
        CBPos = 0
    End Sub

    Sub Walk_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Label34 = "Autowalk L" & CBPos
        Dim nX As Long, nY As Long, nZ As Long
        If CBPos <> 0 Then
            mv = ResWalk
            nX = Int(Split(Walk_Value.List((CBPos - 1) - ResWalk), ",")(0)) - mReadLong(CH_X)
            nY = Int(Split(Walk_Value.List((CBPos - 1) - ResWalk), ",")(1)) - mReadLong(CH_Y)
            nZ = Int(Split(Walk_Value.List((CBPos - 1) - ResWalk), ",")(2)) - mReadLong(CH_Z)
            If nX = 0 And nY = 0 And nZ = 0 Then continue = True
        Else
            continue = True
        End If
        If continue = True Or ResWalk = True Then
10          If CBPos <= Walk_Value.ListCount - 1 Then
                nX = Int(Split(Walk_Value.List(CBPos), ",")(0))
                nY = Int(Split(Walk_Value.List(CBPos), ",")(1))
                nZ = Int(Split(Walk_Value.List(CBPos), ",")(2))
                aZ = Int(Split(Walk_Value.List(CBPos), ",")(3))
                If aZ = 0 Or aZ = 3 Then
                    GotoXYZ nX, nY, nZ
                    nwtX = nX: nwtY = nY: nwtZ = nZ
                ElseIf aZ = 1 Then
                    Dim Cont As Long, Spot As Long, TileID As Long
                    Cont = cWithItem(ITEM_ROPE)
                    Spot = iPosInCont(ITEM_ROPE, Cont)
                    TileID = Map.Map_TileInfo(Map_PlayerTileNum).TileID
                    use_WithCont ITEM_ROPE, mReadLong(CH_X), mReadLong(CH_Y), mReadLong(CH_Z), Spot, Cont, TileID
                ElseIf aZ = 2 Then
                    ' use ladder
                End If
                CBPos = CBPos + 1
            Else
                CBPos = 0: GoTo 10
            End If
        End If
        nX = mReadLong(CH_X)
        nY = mReadLong(CH_Y)
        If nX = laddlasX And nY = laddlasY Then wtTick = wtTick + 1
        laddlasX = nX: laddlasY = nY
        If wtTick > 10 Then wtTick = 0: GoTo 10
    End Sub
    
    Sub Walk_Add_Click()
        If Compiled = True Then On Error Resume Next
        Dim nX As Long, nY As Long, nZ As Long
        nX = mReadLong(CH_X)
        nY = mReadLong(CH_Y)
        nZ = mReadLong(CH_Z)
        Walk_Value.AddItem nX & "," & nY & "," & nZ & ",0"
    End Sub
    Sub Walk_Rem_Click()
        Walk_Value.RemoveItem (Walk_Value.ListIndex)
    End Sub
    Sub Walk_Clr_Click()
        Walk_Value.Clear
    End Sub
    Sub walk_save_click()
        If Compiled Then On Error Resume Next
        Filenm = InputBox("Enter a name for this waypoint log.", "Save autowalk script", "default")
        Open Filenm & ".tmw" For Output As #1
        Print #1, Replace(Attack_Value, vbCrLf, "")
        For a = 0 To Walk_Value.ListCount - 1
            Print #1, Walk_Value.List(a)
        Next
        Close #1
        MsgBox "Information saved."
    End Sub
    Sub walk_load_click()
        If Compiled Then On Error Resume Next
        If CBS_LoadedBefore = False Then
            MsgBox "TM's filetype for cavebot scripts has changed." & vbCrLf & _
                   "Please rename all .wpt files made by TM to .tmw" & vbCrLf & vbCrLf & _
                   "This is due to upcoming compability with scripts" & vbCrLf & _
                   "made by other Tibia cheats. Thanks for your time."
            CBS_LoadedBefore = True
        End If
        frmMain.cmdlg.Filename = ""
        frmMain.cmdlg.Filter = "All cavebot scripts (*.tmw, *.wpt, *.txt, *.xml)|*.tmw;*.wpt;*.txt;*.xml|Tibia Multi (*.tmw)|*.tmw|Tibiabot NG (*.wpt)|*.wpt|BlackD Proxy (*.txt)|*.txt|Tibia Auto (*.xml)|*.xml"
        frmMain.cmdlg.ShowOpen
        file = frmMain.cmdlg.Filename
        If file <> "" Then
            Walk_Clr_Click
            Attack_Value = ""
            If Right(file, 3) = "tmw" Then
                Open file For Input As #1
                Line Input #1, tmp
                Attack_Value = tmp
                While Not EOF(1)
                    Line Input #1, tmp
                    Walk_Value.AddItem tmp
                Wend
                Close #1
                MsgBox "Loaded TM native file."
            ElseIf Right(file, 3) = "xml" Then
                Open file For Input As #1
                While Not EOF(1)
                    Line Input #1, tmp
                    vl = vl & tmp & vbCrLf
                Wend
                Close #1
                tmp = Split(vl, "<waypoint value=""")
                For a = 1 To UBound(tmp)
                    tmp(a) = Split(tmp(a), """")(0)
                    Walk_Value.AddItem tmp(a) & ",0"
                Next
                tmp = Split(vl, "<monster value=""")
                For a = 1 To UBound(tmp)
                    tmp(a) = Split(tmp(a), """")(0)
                    Attack_Value = Attack_Value & tmp(a) & ", "
                Next
                If Attack_Value <> "" Then Attack_Value = Left(Attack_Value, Len(Attack_Value) - 2)
                MsgBox "Loaded Tibia Auto script." & vbCrLf & vbCrLf & _
                       "Note that TM does not currently support multifloor TA scripts."
            ElseIf Right(file, 3) = "txt" Then
                Open file For Input As #1
                While Not EOF(1)
                    Line Input #1, tmp
                    If Left(tmp, 13) = "setMeleeKill " Then
                        Attack_Value = Attack_Value & Mid$(tmp, 14) & ", "
                    ElseIf Left(tmp, 5) = "move " Then
                        Walk_Value.AddItem Mid$(tmp, 6) & ",0"
                    End If
                Wend
                Close #1
                If Attack_Value <> "" Then Attack_Value = Left(Attack_Value, Len(Attack_Value) - 2)
                MsgBox "Loaded BlackD script." & vbCrLf & vbCrLf & _
                       "Note that TM does not currently support multifloor BlackD scripts."
            ElseIf Right(file, 3) = "wpt" Then
                Open file For Input As #1
                While Not EOF(1)
                    Line Input #1, X
                    Line Input #1, Y
                    Line Input #1, z
                    Line Input #1, tmp
                    Walk_Value.AddItem X & "," & Y & "," & z & ",0"
                Wend
                Close #1
                MsgBox "Loaded NG script." & vbCrLf & vbCrLf & _
                       "TM's NG support is very incomplete, therefore the script was probably not loaded correctly." & vbCrLf & _
                       "Please go through the waypoint list to remove obviously wrong coordinates" & vbCrLf & _
                       "(for example 10,0,0,1) and add monsters to attack manually. Thank you."
            End If
        End If
    End Sub

    Sub Walk_Special_Click(Index As Integer)
        If Compiled Then On Error Resume Next
        nX = mReadLong(CH_X)
        nY = mReadLong(CH_Y)
        nZ = mReadLong(CH_Z)
        If Index = 0 Or Index = 3 Or Index = 6 Then nX = nX - 1
        If Index = 2 Or Index = 5 Or Index = 8 Then nX = nX + 1
        If Index = 0 Or Index = 1 Or Index = 2 Then nY = nY - 1
        If Index = 6 Or Index = 7 Or Index = 8 Then nY = nY + 1
        If Walk_SpType(0).Value = True Then spType = 1
        If Walk_SpType(1).Value = True Then spType = 2
        If Walk_SpType(2).Value = True Then spType = 3
        If spType = 0 Then Exit Sub
        If spType <> 3 Then Walk_Value.AddItem nX & "," & nY & "," & nZ & ",0"
        Walk_Value.AddItem nX & "," & nY & "," & nZ & "," & spType
    End Sub

'AUTOLOOT RELATED

    Private Sub Loot_Enabled_Click()
        Loot_Timer.Enabled = Loot_Enabled
    End Sub
    Private Sub Loot_Add_Click()
        If Compiled Then On Error Resume Next
        ItemID = mReadLong(Look_ID)
        vl = InputBox("Adding the item you last looked at (" & ItemID & ")." & vbCrLf & _
             "Please specify what to do with the item." & vbCrLf & vbCrLf & _
             "1-5: Place it in backpack 1 to 5" & vbCrLf & _
             "6: Eat it", "Adding " & ItemID)
        If vl < 1 Or vl > 6 Then MsgBox "Enter a number between 1 to 6!": Exit Sub
        If vl >= 1 And vl <= 5 Then
            Loot_BP(vl - 1) = Loot_BP(vl - 1) & "," & ItemID
        Else
            Loot_Eat = Loot_Eat & "," & ItemID
        End If
    End Sub
    Private Sub Loot_Load_Click()
        Open "Lootlist.txt" For Input As #1
            For a = 0 To Loot_BP.UBound
                Line Input #1, tmp
                Loot_BP(a) = tmp
            Next
            Line Input #1, tmp
            Loot_Eat = tmp
        Close #1
    End Sub
    Private Sub Loot_Save_Click()
        Open "Lootlist.txt" For Output As #1
            For a = 0 To Loot_BP.UBound
                Print #1, Loot_BP(a)
            Next
            Print #1, Loot_Eat
        Close #1
    End Sub
    Private Sub Loot_Timer_Timer()
        If Compiled Then On Error Resume Next
        Dim LootBP As Long, b As Long, c As Long, d As Long, PosFrom As Long, PosTo As Long, ItemID As Long, _
            BpFrom As Long, BpTo As Long, ItemCnt As Long, HasBag As Long, ActionTaken As Boolean
        For LootBP = 0 To 15
            If mReadLong(CT_Start + (LootBP * CTD_Container)) = 0 Then Exit For
        Next
        If LootBP < 2 Then Exit Sub
10      For b = Loot_BP.LBound To Loot_BP.UBound
            If Loot_BP(b) <> "x" Then
                LootList = Split(Loot_BP(b), ",")
                For c = 0 To UBound(LootList)
                    PosFrom = -1: PosTo = -1
                    Do
                        BpFrom = LootBP: BpTo = b + 1: ItemID = LootList(c)
                        PosFrom = iPosInCont(ItemID, LootBP)
                        PosTo = iPosInCont(ItemID, b + 1)
                        If PosFrom = -1 Then Exit Do
                        If BpFrom = BpTo Then Exit Do
                        If PosTo = -1 Then PosTo = 1
                        ItemCnt = mReadLong(CT_Start + (CTD_Container * (BpFrom - 1)) + (CTD_ContainerItem * (PosFrom - 1)) + (CTD_ContainerItemCount))
                        Smsg "Looting " & ItemCnt & "x " & LootList(c) & " from " & BpFrom & "," & PosFrom & " to " & BpTo & "," & PosTo & "."
                        sPck s2ba("F 0 78 FF FF " & Hex(63 + BpFrom) & " 0 " & Hex(PosFrom - 1) & " " & Hex(l2b(ItemID, 1)) & " " & Hex(l2b(ItemID, 2)) & " " & Hex(PosFrom - 1) & " FF FF " & Hex(63 + BpTo) & " 0 " & Hex(PosTo - 1) & " " & Hex(ItemCnt))
                        ActionTaken = True: iSleep Loot_Delay
                    Loop
                Next
            End If
        Next
        HasBag = iPosInCont(CONT_BAG, LootBP)
        If HasBag <> -1 Then useItem CONT_BAG, LootBP, HasBag: iSleep (500): GoTo 10
        If Loot_Eat <> "x" Then
            LootList = Split(Loot_Eat, ",")
            For b = 0 To UBound(LootList)
                ItemID = LootList(b)
                PosFrom = 0: PosTo = 0
                Do
                    PosFrom = iPosInCont(Int(LootList(b)), BpFrom)
                    If PosFrom = -1 Then Exit Do
                    If PosFrom <> -1 Then
                        MsgBox "Eating item " & ItemID & " from " & BpFrom & "," & PosFrom & "."
                        useItem Int(LootList(b)), BpFrom, PosFrom
                        ActionTaken = True: iSleep Loot_Delay
                    End If
                Loop
            Next
        End If
        If ActionTaken Then Smsg "Looting completed."
    End Sub
