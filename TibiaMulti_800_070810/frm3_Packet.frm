VERSION 5.00
Begin VB.Form frm3 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "TM Packet"
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
   Begin VB.TextBox Runemaker_Eat 
      Alignment       =   2  'Center
      BackColor       =   &H00000030&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5280
      TabIndex        =   14
      ToolTipText     =   "Look at your food, then click here to enable food eater."
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Runemaker_Soul 
      Alignment       =   2  'Center
      BackColor       =   &H00000030&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4500
      TabIndex        =   13
      ToolTipText     =   "How much soulpoints you need for the rune"
      Top             =   1215
      Width           =   615
   End
   Begin VB.TextBox Runemaker_Mana 
      Alignment       =   2  'Center
      BackColor       =   &H00000030&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "How much mana you need for the rune"
      Top             =   1215
      Width           =   615
   End
   Begin VB.Timer Runemaker_Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   1440
   End
   Begin VB.CheckBox Runemaker_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   11
      ToolTipText     =   $"frm3_Packet.frx":0000
      Top             =   510
      Width           =   200
   End
   Begin VB.TextBox Aimbot_Value 
      Alignment       =   2  'Center
      BackColor       =   &H00000030&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1380
      TabIndex        =   10
      ToolTipText     =   "What rune to fire. Look at rune, then click here to set."
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox Ammo_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   5280
      TabIndex        =   9
      ToolTipText     =   "Reload ammunition"
      Top             =   2820
      Width           =   200
   End
   Begin VB.TextBox Ammo_Value 
      Alignment       =   2  'Center
      BackColor       =   &H00000030&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5280
      TabIndex        =   8
      ToolTipText     =   "Look at ammunition, then click here to set."
      Top             =   3180
      Width           =   615
   End
   Begin VB.Timer Reload_Timer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5760
      Top             =   3360
   End
   Begin VB.CheckBox Reload_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   7
      ToolTipText     =   "Enable/disable the automatic reloader."
      Top             =   2430
      Width           =   200
   End
   Begin VB.CheckBox Ring_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   4500
      TabIndex        =   6
      ToolTipText     =   "Reload rings"
      Top             =   2820
      Width           =   200
   End
   Begin VB.CheckBox Amulet_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "Reload amulets"
      Top             =   2820
      Width           =   200
   End
   Begin VB.TextBox Ring_Value 
      Alignment       =   2  'Center
      BackColor       =   &H00000030&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4500
      TabIndex        =   4
      ToolTipText     =   "Look at ring, then click here to set."
      Top             =   3180
      Width           =   615
   End
   Begin VB.TextBox Amulet_Value 
      Alignment       =   2  'Center
      BackColor       =   &H00000030&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3720
      TabIndex        =   3
      ToolTipText     =   "Look at amulet, then click here to set."
      Top             =   3180
      Width           =   615
   End
   Begin VB.CheckBox Aimbot_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   2
      ToolTipText     =   "Enable/disable the aimbot. Fire it  using the Delete button on your keyboard."
      Top             =   510
      Width           =   200
   End
   Begin VB.Timer Train_Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   3360
   End
   Begin VB.CheckBox Train_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   1
      ToolTipText     =   "Enable/disable the smart trainer - trains on slimes and other monsters"
      Top             =   2430
      Width           =   200
   End
   Begin VB.CheckBox Train_Stop 
      BackColor       =   &H00517362&
      Caption         =   "Alert_Battle_Logout"
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   880
      TabIndex        =   0
      ToolTipText     =   "Wether to attack other monsters if someone kills the mother"
      Top             =   3000
      Width           =   200
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
      TabIndex        =   27
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
   Begin VB.Shape Shape9 
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
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eater"
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
      Left            =   5280
      TabIndex        =   26
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Soul"
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
      Left            =   4500
      TabIndex        =   25
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   960
      Width           =   615
   End
   Begin VB.Shape Shape18 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFFFF&
      X1              =   3660
      X2              =   5800
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Amn."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   23
      Top             =   2835
      Width           =   495
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Ring"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4740
      TabIndex        =   22
      Top             =   2835
      Width           =   495
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Neck"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   2835
      Width           =   495
   End
   Begin VB.Shape Shape17 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop attacking if mother dies"
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
      Height          =   375
      Left            =   1200
      TabIndex        =   20
      Top             =   2920
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Smart trainer"
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
      Left            =   480
      TabIndex        =   19
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Packet based features"
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
      TabIndex        =   18
      Top             =   1920
      Width           =   9375
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reloader"
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
      Left            =   3600
      TabIndex        =   17
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aimbot"
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
      Left            =   480
      TabIndex        =   16
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Runemaker / Eater"
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
      Left            =   3600
      TabIndex        =   15
      Top             =   495
      Width           =   2415
   End
End
Attribute VB_Name = "frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MdX As Long, MdY As Long, mother As Long, eatTick As Integer
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

'AUTOTRAIN RELATED

    Sub Train_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        If Train_Enabled Then
            mother = mReadLong(BOX_2)
            If mother <> 0 Then
                Smsg "Will attack all mobs, except for " & mother & "."
            Else
                Smsg "Will attack all mobs. Put a mob on follow for slimetraining."
            End If
            Train_Timer = True
            Train_Timer_Timer
        Else
            Train_Timer = False
            Smsg "Trainer enabled: " & False
        End If
    End Sub
    Sub Train_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Dim moID As Long, nX As Long, nY As Long, nZ As Long, cID As Long, Closest As Long, mX As Long, mY As Long, _
            mthrID As Boolean
        nX = mReadLong(CH_X)
        nY = mReadLong(CH_Y)
        nZ = mReadLong(CH_Z)
        cID = mReadLong(CH_ID)
        For a = BL_Start To BL_End Step BL_Dist
            If mReadLong(a + BL_Vis) <> 0 Then
                If mReadLong(a + BL_Z) = nZ Then
                    moID = mReadLong(a + BL_ID)
                    If moID <> cID Then
                        mX = (mReadLong(a + BL_X) - nX)
                        mY = (mReadLong(a + BL_Y) - nY)
                        If mX < 0 Then mX = -mX
                        If mY < 0 Then mY = -mY
                        If mX < 2 And mY < 2 Then
                            If moID = mother Then
                                mthrfound = True
                            Else
                                Closest = moID
                                If mother <> 0 Then
                                    If mthrfound Then GoTo ATK
                                Else
                                    GoTo ATK
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        If mother <> 0 Then
            If mthrfound = False Then
                enAlert
                Smsg "Mother is dead!"
                If Train_Stop Then mother = 0
            Else
                If Closest <> 0 Then Attack Closest
            End If
        Else
ATK:        If Closest <> 0 Then Attack Closest
        End If
    End Sub

'AIMBOT RELATED

    Sub FireRune()
        If Compiled = True Then On Error Resume Next
        Dim runetype As Long, Cont As Long, iPos As Long, tX As Long, tY As Long, tZ As Long
        Trgt = mReadLong(BOX_3)
        For a = BL_Start To BL_End Step BL_Dist
            If mReadLong(a + BL_ID) = Trgt Then
                tX = mReadLong(a + BL_X)
                tY = mReadLong(a + BL_Y)
                tZ = mReadLong(a + BL_Z)
                Trgtname = mReadString(a + BL_Name)
                runetype = Aimbot_Value
                Cont = cWithItem(runetype)
                iPos = iPosInCont(runetype, Cont)
                use_WithCont runetype, tX, tY, tZ, iPos, Cont
                Smsg "Attacking using rune " & runetype & " from backpack " & Cont & ", slot " & iPos & "."
            End If
        Next
    End Sub
    Sub Aimbot_Value_click()
        If Compiled = True Then On Error Resume Next
        Aimbot_Value = mReadLong(ID_Look)
    End Sub

'RELOADER RELATED
    
    Sub Reload_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Reload_Timer.Enabled = Reload_Enabled
        Smsg "Reloader enabled: " & Reload_Timer
    End Sub
    Sub Reload_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Dim Cont As Long, iCont As Long, tmpVal As Long, iCount As Integer '2=amulet, 6=left, 9=ring, a=ammo
        If Amulet_Enabled Then
            If mReadLong(CH_S7) = 0 Then
                tmpVal = Amulet_Value
                Cont = cWithItem(tmpVal)
                iCont = iPosInCont(tmpVal, Cont)
                lodItem &H2, tmpVal, Cont, iCont
            End If
        End If
        If Ring_Enabled Then
            If mReadLong(CH_S8) = 0 Then
                tmpVal = Ring_Value
                Cont = cWithItem(tmpVal)
                iCont = iPosInCont(tmpVal, Cont)
                lodItem &H9, tmpVal, Cont, iCont
            End If
        End If
        If Ammo_Enabled Then
            If mReadLong(CH_S0 + 4) < 100 Then
                tmpVal = Ammo_Value
                Cont = cWithItem(tmpVal)
                iCont = iPosInCont(tmpVal, Cont)
                lodItem &HA, tmpVal, Cont, iCont, nCont(Cont, iCont)
            End If
        End If
    End Sub
    Sub Amulet_Value_Click()
        If Compiled = True Then On Error Resume Next
        Amulet_Value = mReadLong(ID_Look)
    End Sub
    Sub Ring_Value_Click()
        If Compiled = True Then On Error Resume Next
        Ring_Value = mReadLong(ID_Look)
    End Sub
    Sub Ammo_Value_Click()
        If Compiled = True Then On Error Resume Next
        Ammo_Value = mReadLong(ID_Look)
    End Sub

'RUNEMAKER RELATED

    Sub Runemaker_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        Runemaker_Timer = Runemaker_Enabled
        Smsg "Runemaker enabled: " & Runemaker_Timer
    End Sub
    Sub Runemaker_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Dim Cont As Long, iCont As Long, tmpVal As Long
        If Runemaker_Mana <> "" And Runemaker_Soul <> "" Then
            If mReadLong(CH_Ma) >= Int(Runemaker_Mana) Then
                If mReadLong(CH_Sol) >= Int(Runemaker_Soul) Then '2=amulet, 6=left, 9=ring, a=ammo
                    Cont = cWithItem(Ru_NE)
                    If Cont <> -1 Then
                        iCont = iPosInCont(Ru_NE, Cont)
                        lodItem &H6, Ru_NE, Cont, iCont
                        iSleep 1000
                        Call SendMessage(tHvnd, WM_KD, 122, vbNull)
                        Call SendMessage(tHvnd, WM_KU, 122, vbNull)
                        iSleep 1000
                        tosItem &H6, mReadLong(CH_S4), Cont
                        iSleep 1000
                    End If
                End If
            End If
        End If
        eatTick = eatTick + 1
        If eatTick >= 30 Then
            eatTick = 0
            If Runemaker_Eat <> "" Then
                tmpVal = Runemaker_Eat
                Cont = cWithItem(tmpVal)
                iCont = iPosInCont(tmpVal, Cont)
                useItem tmpVal, Cont, iCont
            End If
        End If
    End Sub
    Sub Runemaker_Eat_Click()
        If Compiled = True Then On Error Resume Next
        Runemaker_Eat = mReadLong(ID_Look)
    End Sub
