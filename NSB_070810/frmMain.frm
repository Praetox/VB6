VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   4215
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   7200
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFun 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fun stuff"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdMap 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Map related"
      Height          =   495
      Left            =   4000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdCavebot 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cavebot"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdPacket 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Packet related"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdAlerts 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Alerts"
      Height          =   495
      Left            =   4000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdGeneral 
      BackColor       =   &H00FFC0C0&
      Caption         =   "General"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Timer Hotkeys_Timer 
      Interval        =   3
      Left            =   3720
      Top             =   1920
   End
   Begin VB.Timer Main_Timer 
      Interval        =   250
      Left            =   5400
      Top             =   1920
   End
   Begin VB.Label cmdDev 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dev"
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
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   8880
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   615
   End
   Begin VB.Label cmdHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
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
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Shape Shape1 
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   9375
   End
   Begin VB.Label Author 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shade of Black Software 2007"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Label Apptitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NoSkill Bot v??.??.??"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFE99A&
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label cmdEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private hken As Boolean, MdX As Long, MdY As Long, Fastclick As Boolean, Autoclick As Boolean

Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Compiled = True Then On Error Resume Next
    If Button = 1 Then MdX = X: MdY = Y
End Sub
Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Compiled = True Then On Error Resume Next
    If Button = 1 Then Me.Move (Me.Left + X) - MdX, (Me.Top + Y) - MdY
End Sub
Private Sub cmdHelp_Click()
    ShellExecute 0, "OPEN", "http://praetox.atspace.com/NSB_Help.html", vbNullString, "C:\", 1
End Sub
Private Sub cmdEnd_Click()
    ExitApp
End Sub
Private Sub cmdDev_Click()
    ShowForm frm7
End Sub

Private Sub cmdGeneral_Click()
    ShowForm frm1
End Sub
Private Sub cmdAlerts_Click()
    ShowForm frm2
End Sub
Private Sub cmdPacket_Click()
    ShowForm frm3
End Sub
Private Sub cmdCavebot_Click()
    ShowForm frm4
End Sub
Private Sub cmdMap_Click()
    ShowForm frm5
End Sub
Private Sub cmdFun_Click()
    ShowForm frm6
End Sub
Private Sub ShowForm(frm As Form)
    frm.Show
    frm.Move Me.Left, Me.Top
    Me.Hide
End Sub

Sub Form_Load()
    If Compiled = True Then On Error Resume Next
    Opacity Me.hwnd, 0
    Me.Caption = "TM v" & App.Major & "." & App.Minor & "." & App.Revision
    Apptitle = "NoSkill Bot v" & App.Major & "." & App.Minor & "." & App.Revision
    Me.Show
    fadeIn 0, 255
    Light_S = 14
    hken = True
End Sub
Sub fadeIn(niv1 As Integer, niv2 As Integer)
    If Compiled = True Then On Error Resume Next
    For a = niv1 To niv2 Step 20
        Opacity Me.hwnd, a
        t = Timer
        While (t + 0.05) > Timer
            DoEvents
        Wend
    Next
    Opaque Me.hwnd
End Sub
Sub fadeOut(niv1 As Integer, niv2 As Integer)
    If Compiled = True Then On Error Resume Next
    For a = niv1 To niv2 Step -20
        Opacity Me.hwnd, a
        t = Timer
        While (t + 0.05) > Timer
            DoEvents
        Wend
    Next
End Sub

Sub Hotkeys_Timer_Timer()
    If Compiled = True Then On Error Resume Next
    If GetForegroundWindow = tHvnd Then
        If GetAsyncKeyState(34) = -32767 And hken Then
            If frm1.Light_Enabled = 1 Then frm1.Light_Enabled = 0 Else frm1.Light_Enabled = 1
            frm1.Light_Enabled_Click
        End If
        If GetAsyncKeyState(45) = -32767 And hken Then
            If frm3.Train_Enabled = 1 Then frm3.Train_Enabled = 0 Else frm3.Train_Enabled = 1
            frm3.Train_Enabled_Click
        End If
        If GetAsyncKeyState(33) = -32767 And hken Then Fastclick = Not Fastclick: Smsg "Fastclick enabled: " & Fastclick
        If GetAsyncKeyState(36) = -32767 And hken Then Autoclick = Not Autoclick: Smsg "Autoclick enabled: " & Autoclick
        If GetAsyncKeyState(35) = -32767 And hken Then Experience
        DelState = GetAsyncKeyState(46): If DelState = -32767 Then delpressed = True
        If GetAsyncKeyState(17) And delpressed Then hken = Not hken: Smsg "Hotkeys enabled: " & hken
        If delpressed Then doAlert = False
        If DelState <> 0 Then
            If frm3.Aimbot_Spam Then
                If hken And frm3.Aimbot_Enabled.Value Then
                    If ((Timer * 1000) - LastAttackRuneFire) > frm3.Aimbot_Delay Then
                        frm3.FireRune
                        LastAttackRuneFire = Timer * 1000
                    End If
                End If
            Else
                If delpressed And hken And frm3.Aimbot_Enabled.Value Then frm3.FireRune
            End If
        End If
        If delpressed And hken And frm4.WalkAa_Enabled.Value Then frm4.Walk_Add_Click
    End If
End Sub
Sub Experience()
    If Compiled = True Then On Error Resume Next
    cLevel = mReadLong(CH_Lvl) + 1
    cExp = mReadLong(CH_Exp)
    cExpNext = (((50 / 3) * (cLevel ^ 3)) - (100 * (cLevel ^ 2)) + ((850 / 3) * cLevel) - 200) - cExp
    cExpNext = cExpNext \ 1
    For a = 0 To Len(cExpNext) - 1
        If num = 3 Then cEN = "'" & cEN: num = 0
        num = num + 1
        cEN = Mid$(cExpNext, Len(cExpNext) - a, 1) & cEN
    Next
    mobs = Split("Skunk%3#Badger%5#Rat%5#Cave Rat%10#Snake%10#Bat%10#Spider%12#Bug%18#Wolf%18#Troll%20#Winter Wolf%20#Hyaena%20#Spit Nettle%20#Island troll%20#Poison Spider%22#Bear%23#" & _
                 "Frost Troll%23#Panda%23#Wasp%24#Orc%25#Goblin%25#Swamp Troll%25#Polar Bear%28#Lion%30#Cobra%30#Dworc Venomsniper%30#Crab%30#Centipede%30#Skeleton%35#Dworc Fleshhunter%35#" & _
                 "Dworc Voodoomaster%35#Orc Spearman%38#Rotworm%40#Crocodile%40#Tiger%40#Elf%42#Larva%44#Dwarf%45#Scorpion%45#Smuggler%48#Orc Warrior%50#Minotaur%50#War Wolf%55#Amazon%60#" & _
                 "Wild Warrior%60#Minotaur Archer%65#Bandit%65#Dwarf Soldier%70#Elf Scout%75#Ghoul%85#Valkyrie%85#Stalker%90#Gazer%90#Tortoise%90#Lizard Sentinel%100#Sibang%100#Novice Of The Cult%100#" & _
                 "Assassin%105#Orc Rider%110#Orc Shaman%110#Fire Devil%110#Kongra%110#Ghost%120#Witch%120#Scarab%120#Tarantula%120#Pirate Marauder%125#Merlkin%135#Dark Monk%145#Lizard Templar%145#" & _
                 "Cyclops%150#Hunter%150#Minotaur Mage%150#Mummy%150#Gargoyle%150#Terror Bird%150#Carniphila%150#Minotaur Guard%160#Slime%160#Stone Golem%160#Elephant%160#Dwarf Guard%165#" & _
                 "Beholder%170#Elf Arcanist%175#Pirate Cutthroat%175#Blue Djinn%190#Green Djinn%190#Crypt Shambler%195#Orc Berserker%195#Monk%200#Lizard Snakecharmer%200#The Horned Fox%200#" & _
                 "Fire Elemental%220#Demon Skeleton%240#Dwarf Geomancer%245#Pirate Ghost%250#Quara Constrictor%250#Orc Leader%270#Elder Beholder%280#Vampire%290#Efreet%300#Marid%300#Pirate Corsair%350#" & _
                 "Dharalion%380#Fernfang%400#Quara Mantassin%400#Priestess%420#Yeti%460#The Evil Eye%500#Enlightened of the Cult%500#General Murius%550#Bonebeast%580#Necromancer%580#Orc Warlord%670#" & _
                 "Dragon%700#Necropharus%700#Ancient Scarab%720#Quara Hydromancer%800#Banshee%900#Giant Spider%900#Lich%900#Hero%1200#Quara Pincher%1200#Black Knight%1600#Grorlam%1600#" & _
                 "Quara Predator%1600#Serpent Spawn%2000#Dragon Lord%2100#Hydra%2100#Behemoth%2500#The Old Widow%2800#Dipthrah%2900#Omruc%2950#Thalas%2950#Vashresamun%2950#Morguthis%3000#" & _
                 "Mahrdis%3050#Ashmunrah%3100#Rahemos%3100#Warlock%4000#Demodras%4000#Demon%6000", "#")
    MobID = mReadLong(BOX_3)
    If MobID <> 0 Then
        For a = BL_Start To BL_End Step BL_Dist
            If mReadLong(a + BL_ID) = MobID Then
                For b = 0 To UBound(mobs)
                    MobName = Split(mobs(b), "%")(0)
                    If MobName = mReadString(a + BL_Name) Then
                        MobExp = Split(mobs(b), "%")(1)
                        GoTo 10
                    End If
                Next
            End If
        Next
10      If Int(MobExp) > 0 Then
            MobsNeeded = Int(cExpNext) / Int(MobExp)
            If MobsNeeded <> MobsNeeded \ 1 Then MobsNeeded = (MobsNeeded \ 1) + 1
            MobInfo = "(" & MobsNeeded & " " & LCase(MobName) & "s) "
        End If
    End If
    Smsg "You need " & cEN & " exp " & MobInfo & "to get level " & cLevel \ 1 & "."
End Sub

'MAIN FUNCTIONS

    Sub Main_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        If mReadLong(CH_Con) = 0 Then doAlert = True
        If Fastclick Then mWriteLong CH_Clk, 7
        If Autoclick Then mWriteLong CH_Clk, 7: mouse_event &H2, 0, 0, 0, 0: mouse_event &H4, 0, 0, 0, 0
        If doAlert Then Beep: FlashWindow tHvnd, 0
    End Sub
