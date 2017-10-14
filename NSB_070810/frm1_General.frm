VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "TM General"
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
   Begin VB.CommandButton BoH_dec 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Decrease by 10%"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Decrease your walking speed by 20%"
      Top             =   3180
      Width           =   1695
   End
   Begin VB.CommandButton BoH_inc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Increase by 10%"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Increase your walking speed by 20%"
      Top             =   2820
      Width           =   1695
   End
   Begin VB.CheckBox FTP_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   6750
      TabIndex        =   19
      ToolTipText     =   "Enable/disable the stats uploader"
      Top             =   510
      Width           =   200
   End
   Begin VB.CommandButton FTP_Change 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Change settings"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Change the configuration for FTP Stats Uploader"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Timer Stat_Timer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5760
      Top             =   1440
   End
   Begin VB.TextBox Stat_Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   3840
      TabIndex        =   16
      Text            =   "{name} ({exp}/{e2l}) {hp} - {mana}"
      ToolTipText     =   "What to show in the taskbar"
      Top             =   885
      Width           =   1935
   End
   Begin VB.CheckBox Stat_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   15
      ToolTipText     =   "Enable/disable showing char info in the taskbar"
      Top             =   510
      Width           =   200
   End
   Begin VB.CommandButton Stat_Map 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show map"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Show your character position using TibiaNews map viewer."
      Top             =   1260
      Width           =   1695
   End
   Begin VB.TextBox Light_C 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "207"
      ToolTipText     =   "Color of the light - default is 207, torch is 206. Use the up/down keys on your keyboard to change it."
      Top             =   1320
      Width           =   735
   End
   Begin VB.CheckBox Light_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   2
      ToolTipText     =   "Enable/disable the lighthack"
      Top             =   510
      Width           =   200
   End
   Begin VB.CheckBox INV_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   1
      ToolTipText     =   "Enable/disable the cancel invisibility feature"
      Top             =   2430
      Width           =   200
   End
   Begin VB.CheckBox AFK_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   0
      ToolTipText     =   "Enable/disable the anti-afk feature"
      Top             =   2430
      Width           =   200
   End
   Begin VB.Timer Light_Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   1440
   End
   Begin VB.Timer AFK_Timer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2640
      Top             =   3360
   End
   Begin VB.Timer INV_Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   3360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Free BoH"
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
      TabIndex        =   21
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   9000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FTP Stats Uploader"
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
      Top             =   495
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Display status"
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
      TabIndex        =   17
      Top             =   495
      Width           =   2415
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   480
      Width           =   2415
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
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape7 
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
   Begin VB.Shape Shape8 
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sends a dummy packet every 60 seconds"
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
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   2920
      Width           =   1935
   End
   Begin VB.Label Light_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FFE99A&
      Height          =   195
      Left            =   780
      TabIndex        =   11
      ToolTipText     =   "Light strength - how much currently set"
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
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
      TabIndex        =   10
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "General features"
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
      TabIndex        =   9
      Top             =   1920
      Width           =   9375
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shows invisible creatures"
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
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2920
      Width           =   1695
   End
   Begin VB.Label Light_2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "Light strength - max power"
      Top             =   900
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Light hack"
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
      TabIndex        =   6
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel invisibility"
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
      TabIndex        =   5
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anti-AFK"
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
      TabIndex        =   4
      Top             =   2415
      Width           =   2415
   End
End
Attribute VB_Name = "frm1"
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

'LIGHT RELATED
    
    Sub Light_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Compiled = True Then On Error Resume Next
        If X > 60 And X < 1875 And Button = 1 Then
            lightstart = X - 60
            Light_1.Width = lightstart
            Light_S = (14 / 1815) * lightstart
            If Light_Timer.Enabled = True Then Light_Timer_Timer
            mWriteString CH_TSt, "Light strength set to " & Light_S & "."
            mWriteLong CH_TTi, 50
        End If
    End Sub
    Sub Light_2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Compiled = True Then On Error Resume Next
        If X > 60 And X < 1875 And Button = 1 Then
            lightstart = X - 60
            Light_1.Width = lightstart
            Light_S = (14 / 1815) * lightstart
            If Light_Timer.Enabled = True Then Light_Timer_Timer
            Smsg "Light strength set to " & Light_S & "."
        End If
        End Sub
    Sub Light_1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Compiled = True Then On Error Resume Next
        If X > 0 And X < 1815 And Button = 1 Then
            Light_1.Width = X
            Light_S = (14 / 1815) * X
            If Light_Timer.Enabled = True Then Light_Timer_Timer
            Smsg "Light strength set to " & Light_S & "."
        End If
    End Sub
    Sub Light_1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Compiled = True Then On Error Resume Next
        If X > 0 And X < 1815 And Button = 1 Then
            Light_1.Width = X
            Light_S = (14 / 1815) * X
            If Light_Timer.Enabled = True Then Light_Timer_Timer
            Smsg "Light strength set to " & Light_S & "."
        End If
    End Sub
    Sub Light_C_KeyDown(KeyCode As Integer, Shift As Integer)
        If Compiled = True Then On Error Resume Next
        If KeyCode = 38 Then Light_C = Light_C + 1: Light_Timer_Timer
        If KeyCode = 40 Then Light_C = Light_C - 1: Light_Timer_Timer
    End Sub
    Sub Light_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Light_Timer = Light_Enabled
        Light_Timer_Timer
        Smsg "Lighthack enabled: " & Light_Timer
    End Sub
    Sub Light_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        tmp = BL_Player
        If tmp <> 0 Then
            mWriteLong tmp + BL_LStr, Light_S
            mWriteLong tmp + BL_LCol, Light_C
        End If
    End Sub

'AFK RELATED

    Sub AFK_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        AFK_Timer = AFK_Enabled
        afk_timer_timer
        Smsg "AFK dance enabled: " & AFK_Timer
    End Sub
    Sub afk_timer_timer()
        If Compiled = True Then On Error Resume Next
        tmp = BL_Player
        If tmp <> 0 Then
            dr = mReadLong(tmp + BL_Dir)
            Call SendMessage(tHvnd, WM_KD, 17, vbNull)
            If dr = 0 Then
                Call SendMessage(tHvnd, WM_KD, 38, vbNull)
                Call SendMessage(tHvnd, WM_KU, 38, vbNull)
            ElseIf dr = 1 Then
                Call SendMessage(tHvnd, WM_KD, 39, vbNull)
                Call SendMessage(tHvnd, WM_KU, 39, vbNull)
            ElseIf dr = 2 Then
                Call SendMessage(tHvnd, WM_KD, 40, vbNull)
                Call SendMessage(tHvnd, WM_KU, 40, vbNull)
            ElseIf dr = 3 Then
                Call SendMessage(tHvnd, WM_KD, 37, vbNull)
                Call SendMessage(tHvnd, WM_KU, 37, vbNull)
            End If
            Call SendMessage(tHvnd, WM_KU, 17, vbNull)
        End If
    End Sub

'FTP RELATED

    Sub FTP_Change_Click()
        If Compiled = True Then On Error Resume Next
        frmFTP.Show
    End Sub
    Sub ftp_enabled_click()
        If Compiled = True Then On Error Resume Next
        frmFTP.tmr = FTP_Enabled
        Smsg "FTP Uploader enabled: " & frmFTP.tmr
    End Sub

'CANCEL INVIS RELATED

    Sub INV_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        INV_Timer = INV_Enabled
        INV_Timer_Timer
        Smsg "Showing invisible creatures: " & INV_Timer
    End Sub
    Sub INV_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        For a = BL_Start To BL_End Step BL_Dist
            If mReadLong(a + BL_OFit) = OFit_Invis Then
                mWriteLong a + BL_OFit, OFit_Druid
            End If
        Next
    End Sub

'SHOW STATS RELATED

    Sub Stat_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Stat_Timer = Stat_Enabled
        Stat_Timer_Timer
        Smsg "Showing stats in taskbar enabled: " & Stat_Timer
    End Sub
    Sub Stat_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        tmp = Stat_Value
        tmp = Replace(tmp, "{name}", frmStart.lstname)
        tmp = Replace(tmp, "{hp}", mReadLong(CH_HP))
        tmp = Replace(tmp, "{mana}", mReadLong(CH_Ma))
        tmp = Replace(tmp, "{exp}", mReadLong(CH_Exp))
        If InStr(1, tmp, "{e2l}") > 0 Then
            cLevel = mReadLong(CH_Lvl) + 1
            cExp = mReadLong(CH_Exp)
            cExpNext = (((50 / 3) * (cLevel ^ 3)) - (100 * (cLevel ^ 2)) + ((850 / 3) * cLevel) - 200) - cExp
            tmp = Replace(tmp, "{e2l}", cExpNext)
        End If
        tmp = Replace(tmp, "{lvl}", mReadLong(CH_Lvl))
        If mReadLong(CH_Con) = 8 Then con = "" Else con = "OFFLINE!  "
        SetWindowText tHvnd, con & tmp
    End Sub

'TIBIANEWS MAP RELATED

    Sub Stat_Map_Click()
        If Compiled = True Then On Error Resume Next
        TN_StartX = 1125: TN_StartY = -867
        CO_StartX = 32059: CO_StartY = 32202
        nX = mReadLong(CH_X)
        nY = mReadLong(CH_Y)
        nZ = mReadLong(CH_Z)
        xcor = CO_StartX - nX
        xcor = TN_StartX + (xcor * 3)
        ycor = CO_StartY - nY
        ycor = TN_StartY + (ycor * 3)
        ShellExecute Me.hwnd, "OPEN", "http://www.tibianews.net/worldmap.asp?xcor=" & xcor & "&ycor=" & ycor, _
                     vbNull, "c:\", 1
    End Sub

'BOH RELATED
    
    Private Sub BoH_inc_Click()
        spd = mReadLong(BL_Player + BL_Spd)
        mWriteLong BL_Player + BL_Spd, (spd * 1.1) \ 1
        Smsg "Increased speed from " & spd & " to " & (spd * 1.1) \ 1 & " (10% increase)."
    End Sub
    Private Sub BoH_dec_Click()
        spd = mReadLong(BL_Player + BL_Spd)
        mWriteLong BL_Player + BL_Spd, (spd / 1.1) \ 1
        Smsg "Decreased speed from " & spd & " to " & (spd / 1.1) \ 1 & " (10% decrease)."
    End Sub
