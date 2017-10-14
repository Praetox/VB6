VERSION 5.00
Begin VB.Form frm5 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "TM Map"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Spy_chkTrack 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3960
      TabIndex        =   286
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   3240
      Width           =   200
   End
   Begin VB.ComboBox Spy_List 
      Height          =   315
      Left            =   3720
      TabIndex        =   285
      Text            =   "Mob list... Select to track"
      Top             =   2820
      Width           =   2165
   End
   Begin VB.CheckBox Spy_Lister 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   283
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   2430
      Width           =   200
   End
   Begin VB.Timer Spy_Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8880
      Top             =   3360
   End
   Begin VB.CheckBox Spy_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   6750
      TabIndex        =   16
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   510
      Width           =   200
   End
   Begin VB.Timer FishDisp_Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2640
      Top             =   1440
   End
   Begin VB.CheckBox FishDisp_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   6
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   510
      Width           =   200
   End
   Begin VB.PictureBox FishDisp_Display 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   5
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   1110
      Picture         =   "frm5_Map.frx":0000
      ScaleHeight     =   840
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   780
      Width           =   1140
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   85
         Left            =   530
         ScaleHeight     =   90
         ScaleWidth      =   90
         TabIndex        =   5
         Top             =   370
         Width           =   85
      End
   End
   Begin VB.CheckBox FishHack_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   510
      TabIndex        =   3
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   2430
      Width           =   200
   End
   Begin VB.Timer FishHack_Timer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2640
      Top             =   3360
   End
   Begin VB.CheckBox FishHack_Randomize 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   795
      TabIndex        =   2
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   2835
      Width           =   200
   End
   Begin VB.CheckBox FishHack_StopAtLowCap 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   795
      TabIndex        =   1
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox Spear_Enabled 
      BackColor       =   &H00517362&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   3630
      TabIndex        =   0
      ToolTipText     =   "Enable/disable the cavebot's autoattack feature"
      Top             =   510
      Width           =   200
   End
   Begin VB.Timer Spear_Timer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5775
      Top             =   1455
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   269
      Left            =   8880
      TabIndex        =   305
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   268
      Left            =   6840
      TabIndex        =   304
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   267
      Left            =   6960
      TabIndex        =   303
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   266
      Left            =   7080
      TabIndex        =   302
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   265
      Left            =   7200
      TabIndex        =   301
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   264
      Left            =   7320
      TabIndex        =   300
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   263
      Left            =   7440
      TabIndex        =   299
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   262
      Left            =   7560
      TabIndex        =   298
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   261
      Left            =   7680
      TabIndex        =   297
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   260
      Left            =   7800
      TabIndex        =   296
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   259
      Left            =   7920
      TabIndex        =   295
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   258
      Left            =   8040
      TabIndex        =   294
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   257
      Left            =   8160
      TabIndex        =   293
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   256
      Left            =   8280
      TabIndex        =   292
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   255
      Left            =   8400
      TabIndex        =   291
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   254
      Left            =   8520
      TabIndex        =   290
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   253
      Left            =   8640
      TabIndex        =   289
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   8760
      TabIndex        =   288
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Track (else, scan)"
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
      Height          =   225
      Left            =   4260
      TabIndex        =   287
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mob Spy Lister (all)"
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
      TabIndex        =   284
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   251
      Left            =   8760
      TabIndex        =   282
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   250
      Left            =   8640
      TabIndex        =   281
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   249
      Left            =   8520
      TabIndex        =   280
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   248
      Left            =   8400
      TabIndex        =   279
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   247
      Left            =   8280
      TabIndex        =   278
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   246
      Left            =   8160
      TabIndex        =   277
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   245
      Left            =   8040
      TabIndex        =   276
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   244
      Left            =   7920
      TabIndex        =   275
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   243
      Left            =   7800
      TabIndex        =   274
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   242
      Left            =   7680
      TabIndex        =   273
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   241
      Left            =   7560
      TabIndex        =   272
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   240
      Left            =   7440
      TabIndex        =   271
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   239
      Left            =   7320
      TabIndex        =   270
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   238
      Left            =   7200
      TabIndex        =   269
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   237
      Left            =   7080
      TabIndex        =   268
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   236
      Left            =   6960
      TabIndex        =   267
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   235
      Left            =   6840
      TabIndex        =   266
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   233
      Left            =   8760
      TabIndex        =   265
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   232
      Left            =   8640
      TabIndex        =   264
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   231
      Left            =   8520
      TabIndex        =   263
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   230
      Left            =   8400
      TabIndex        =   262
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   229
      Left            =   8280
      TabIndex        =   261
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   228
      Left            =   8160
      TabIndex        =   260
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   227
      Left            =   8040
      TabIndex        =   259
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   226
      Left            =   7920
      TabIndex        =   258
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   225
      Left            =   7800
      TabIndex        =   257
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   224
      Left            =   7680
      TabIndex        =   256
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   223
      Left            =   7560
      TabIndex        =   255
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   222
      Left            =   7440
      TabIndex        =   254
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   221
      Left            =   7320
      TabIndex        =   253
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   220
      Left            =   7200
      TabIndex        =   252
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   219
      Left            =   7080
      TabIndex        =   251
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   218
      Left            =   6960
      TabIndex        =   250
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   217
      Left            =   6840
      TabIndex        =   249
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   215
      Left            =   8760
      TabIndex        =   248
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   214
      Left            =   8640
      TabIndex        =   247
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   213
      Left            =   8520
      TabIndex        =   246
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   212
      Left            =   8400
      TabIndex        =   245
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   211
      Left            =   8280
      TabIndex        =   244
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   210
      Left            =   8160
      TabIndex        =   243
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   209
      Left            =   8040
      TabIndex        =   242
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   208
      Left            =   7920
      TabIndex        =   241
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   207
      Left            =   7800
      TabIndex        =   240
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   206
      Left            =   7680
      TabIndex        =   239
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   205
      Left            =   7560
      TabIndex        =   238
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   204
      Left            =   7440
      TabIndex        =   237
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   203
      Left            =   7320
      TabIndex        =   236
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   202
      Left            =   7200
      TabIndex        =   235
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   201
      Left            =   7080
      TabIndex        =   234
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   200
      Left            =   6960
      TabIndex        =   233
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   199
      Left            =   6840
      TabIndex        =   232
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   197
      Left            =   8760
      TabIndex        =   231
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   196
      Left            =   8640
      TabIndex        =   230
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   195
      Left            =   8520
      TabIndex        =   229
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   194
      Left            =   8400
      TabIndex        =   228
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   193
      Left            =   8280
      TabIndex        =   227
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   192
      Left            =   8160
      TabIndex        =   226
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   191
      Left            =   8040
      TabIndex        =   225
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   190
      Left            =   7920
      TabIndex        =   224
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   189
      Left            =   7800
      TabIndex        =   223
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   188
      Left            =   7680
      TabIndex        =   222
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   187
      Left            =   7560
      TabIndex        =   221
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   186
      Left            =   7440
      TabIndex        =   220
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   185
      Left            =   7320
      TabIndex        =   219
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   184
      Left            =   7200
      TabIndex        =   218
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   183
      Left            =   7080
      TabIndex        =   217
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   182
      Left            =   6960
      TabIndex        =   216
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   181
      Left            =   6840
      TabIndex        =   215
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   179
      Left            =   8760
      TabIndex        =   214
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   178
      Left            =   8640
      TabIndex        =   213
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   177
      Left            =   8520
      TabIndex        =   212
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   176
      Left            =   8400
      TabIndex        =   211
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   175
      Left            =   8280
      TabIndex        =   210
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   174
      Left            =   8160
      TabIndex        =   209
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   173
      Left            =   8040
      TabIndex        =   208
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   172
      Left            =   7920
      TabIndex        =   207
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   171
      Left            =   7800
      TabIndex        =   206
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   170
      Left            =   7680
      TabIndex        =   205
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   169
      Left            =   7560
      TabIndex        =   204
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   168
      Left            =   7440
      TabIndex        =   203
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   167
      Left            =   7320
      TabIndex        =   202
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   166
      Left            =   7200
      TabIndex        =   201
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   165
      Left            =   7080
      TabIndex        =   200
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   164
      Left            =   6960
      TabIndex        =   199
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   163
      Left            =   6840
      TabIndex        =   198
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   161
      Left            =   8760
      TabIndex        =   197
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   160
      Left            =   8640
      TabIndex        =   196
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   159
      Left            =   8520
      TabIndex        =   195
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   158
      Left            =   8400
      TabIndex        =   194
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   157
      Left            =   8280
      TabIndex        =   193
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   156
      Left            =   8160
      TabIndex        =   192
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   155
      Left            =   8040
      TabIndex        =   191
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   154
      Left            =   7920
      TabIndex        =   190
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   153
      Left            =   7800
      TabIndex        =   189
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   152
      Left            =   7680
      TabIndex        =   188
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   151
      Left            =   7560
      TabIndex        =   187
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   150
      Left            =   7440
      TabIndex        =   186
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   149
      Left            =   7320
      TabIndex        =   185
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   148
      Left            =   7200
      TabIndex        =   184
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   147
      Left            =   7080
      TabIndex        =   183
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   146
      Left            =   6960
      TabIndex        =   182
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   145
      Left            =   6840
      TabIndex        =   181
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   143
      Left            =   8760
      TabIndex        =   180
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   142
      Left            =   8640
      TabIndex        =   179
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   141
      Left            =   8520
      TabIndex        =   178
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   140
      Left            =   8400
      TabIndex        =   177
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   139
      Left            =   8280
      TabIndex        =   176
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   138
      Left            =   8160
      TabIndex        =   175
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   137
      Left            =   8040
      TabIndex        =   174
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   136
      Left            =   7920
      TabIndex        =   173
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   135
      Left            =   7800
      TabIndex        =   172
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   134
      Left            =   7680
      TabIndex        =   171
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   133
      Left            =   7560
      TabIndex        =   170
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   132
      Left            =   7440
      TabIndex        =   169
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   131
      Left            =   7320
      TabIndex        =   168
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   130
      Left            =   7200
      TabIndex        =   167
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   129
      Left            =   7080
      TabIndex        =   166
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   128
      Left            =   6960
      TabIndex        =   165
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   127
      Left            =   6840
      TabIndex        =   164
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   125
      Left            =   8760
      TabIndex        =   163
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   124
      Left            =   8640
      TabIndex        =   162
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   123
      Left            =   8520
      TabIndex        =   161
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   122
      Left            =   8400
      TabIndex        =   160
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   121
      Left            =   8280
      TabIndex        =   159
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   120
      Left            =   8160
      TabIndex        =   158
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   119
      Left            =   8040
      TabIndex        =   157
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   118
      Left            =   7920
      TabIndex        =   156
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   117
      Left            =   7800
      TabIndex        =   155
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   116
      Left            =   7680
      TabIndex        =   154
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   115
      Left            =   7560
      TabIndex        =   153
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   114
      Left            =   7440
      TabIndex        =   152
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   113
      Left            =   7320
      TabIndex        =   151
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   112
      Left            =   7200
      TabIndex        =   150
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   111
      Left            =   7080
      TabIndex        =   149
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   110
      Left            =   6960
      TabIndex        =   148
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   109
      Left            =   6840
      TabIndex        =   147
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   107
      Left            =   8760
      TabIndex        =   146
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   106
      Left            =   8640
      TabIndex        =   145
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   105
      Left            =   8520
      TabIndex        =   144
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   104
      Left            =   8400
      TabIndex        =   143
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   103
      Left            =   8280
      TabIndex        =   142
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   102
      Left            =   8160
      TabIndex        =   141
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   101
      Left            =   8040
      TabIndex        =   140
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   100
      Left            =   7920
      TabIndex        =   139
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   99
      Left            =   7800
      TabIndex        =   138
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   98
      Left            =   7680
      TabIndex        =   137
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   97
      Left            =   7560
      TabIndex        =   136
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   96
      Left            =   7440
      TabIndex        =   135
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   95
      Left            =   7320
      TabIndex        =   134
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   94
      Left            =   7200
      TabIndex        =   133
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   93
      Left            =   7080
      TabIndex        =   132
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   92
      Left            =   6960
      TabIndex        =   131
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   91
      Left            =   6840
      TabIndex        =   130
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   89
      Left            =   8760
      TabIndex        =   129
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   88
      Left            =   8640
      TabIndex        =   128
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   87
      Left            =   8520
      TabIndex        =   127
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   86
      Left            =   8400
      TabIndex        =   126
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   85
      Left            =   8280
      TabIndex        =   125
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   84
      Left            =   8160
      TabIndex        =   124
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   83
      Left            =   8040
      TabIndex        =   123
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   82
      Left            =   7920
      TabIndex        =   122
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   81
      Left            =   7800
      TabIndex        =   121
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   80
      Left            =   7680
      TabIndex        =   120
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   79
      Left            =   7560
      TabIndex        =   119
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   78
      Left            =   7440
      TabIndex        =   118
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   77
      Left            =   7320
      TabIndex        =   117
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   76
      Left            =   7200
      TabIndex        =   116
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   75
      Left            =   7080
      TabIndex        =   115
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   74
      Left            =   6960
      TabIndex        =   114
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   73
      Left            =   6840
      TabIndex        =   113
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   71
      Left            =   8760
      TabIndex        =   112
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   70
      Left            =   8640
      TabIndex        =   111
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   69
      Left            =   8520
      TabIndex        =   110
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   68
      Left            =   8400
      TabIndex        =   109
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   67
      Left            =   8280
      TabIndex        =   108
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   66
      Left            =   8160
      TabIndex        =   107
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   65
      Left            =   8040
      TabIndex        =   106
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   64
      Left            =   7920
      TabIndex        =   105
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   63
      Left            =   7800
      TabIndex        =   104
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   62
      Left            =   7680
      TabIndex        =   103
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   61
      Left            =   7560
      TabIndex        =   102
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   60
      Left            =   7440
      TabIndex        =   101
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   59
      Left            =   7320
      TabIndex        =   100
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   58
      Left            =   7200
      TabIndex        =   99
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   57
      Left            =   7080
      TabIndex        =   98
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   56
      Left            =   6960
      TabIndex        =   97
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   55
      Left            =   6840
      TabIndex        =   96
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   53
      Left            =   8760
      TabIndex        =   95
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   52
      Left            =   8640
      TabIndex        =   94
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   51
      Left            =   8520
      TabIndex        =   93
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   50
      Left            =   8400
      TabIndex        =   92
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   49
      Left            =   8280
      TabIndex        =   91
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   48
      Left            =   8160
      TabIndex        =   90
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   47
      Left            =   8040
      TabIndex        =   89
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   46
      Left            =   7920
      TabIndex        =   88
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   45
      Left            =   7800
      TabIndex        =   87
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   44
      Left            =   7680
      TabIndex        =   86
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   43
      Left            =   7560
      TabIndex        =   85
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   42
      Left            =   7440
      TabIndex        =   84
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   41
      Left            =   7320
      TabIndex        =   83
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   40
      Left            =   7200
      TabIndex        =   82
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   39
      Left            =   7080
      TabIndex        =   81
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   38
      Left            =   6960
      TabIndex        =   80
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   37
      Left            =   6840
      TabIndex        =   79
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   35
      Left            =   8760
      TabIndex        =   78
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   34
      Left            =   8640
      TabIndex        =   77
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   33
      Left            =   8520
      TabIndex        =   76
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   32
      Left            =   8400
      TabIndex        =   75
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   31
      Left            =   8280
      TabIndex        =   74
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   30
      Left            =   8160
      TabIndex        =   73
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   29
      Left            =   8040
      TabIndex        =   72
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   28
      Left            =   7920
      TabIndex        =   71
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   27
      Left            =   7800
      TabIndex        =   70
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   26
      Left            =   7680
      TabIndex        =   69
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   25
      Left            =   7560
      TabIndex        =   68
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   24
      Left            =   7440
      TabIndex        =   67
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   23
      Left            =   7320
      TabIndex        =   66
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   22
      Left            =   7200
      TabIndex        =   65
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   21
      Left            =   7080
      TabIndex        =   64
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   20
      Left            =   6960
      TabIndex        =   63
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   19
      Left            =   6840
      TabIndex        =   62
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   17
      Left            =   8760
      TabIndex        =   61
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   16
      Left            =   8640
      TabIndex        =   60
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   15
      Left            =   8520
      TabIndex        =   59
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   14
      Left            =   8400
      TabIndex        =   58
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   13
      Left            =   8280
      TabIndex        =   57
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   12
      Left            =   8160
      TabIndex        =   56
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   11
      Left            =   8040
      TabIndex        =   55
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   10
      Left            =   7920
      TabIndex        =   54
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   9
      Left            =   7800
      TabIndex        =   53
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   8
      Left            =   7680
      TabIndex        =   52
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   7
      Left            =   7560
      TabIndex        =   51
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   6
      Left            =   7440
      TabIndex        =   50
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   5
      Left            =   7320
      TabIndex        =   49
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   7200
      TabIndex        =   48
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   7080
      TabIndex        =   47
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   6960
      TabIndex        =   46
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   6840
      TabIndex        =   45
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   252
      Left            =   8880
      TabIndex        =   44
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   234
      Left            =   8880
      TabIndex        =   43
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   216
      Left            =   8880
      TabIndex        =   42
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   198
      Left            =   8880
      TabIndex        =   41
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   180
      Left            =   8880
      TabIndex        =   40
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   162
      Left            =   8880
      TabIndex        =   39
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   144
      Left            =   8880
      TabIndex        =   38
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   126
      Left            =   8880
      TabIndex        =   37
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   108
      Left            =   8880
      TabIndex        =   36
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   90
      Left            =   8880
      TabIndex        =   35
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   72
      Left            =   8880
      TabIndex        =   34
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   54
      Left            =   8880
      TabIndex        =   33
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   36
      Left            =   8880
      TabIndex        =   32
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Spy_Tile 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   18
      Left            =   8880
      TabIndex        =   31
      Top             =   840
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Spy_Filter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+2"
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
      Index           =   5
      Left            =   8640
      TabIndex        =   30
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   3135
      Left            =   6720
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mob Spy Tracker"
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
      Left            =   6720
      TabIndex        =   29
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      TabIndex        =   28
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
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
      TabIndex        =   27
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Spy_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Swamp Troll"
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
      Left            =   7440
      TabIndex        =   26
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Spy_Health 
      BackStyle       =   0  'Transparent
      Caption         =   "78%"
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
      Left            =   7440
      TabIndex        =   25
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Floor:"
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
      Left            =   8040
      TabIndex        =   24
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Spy_Floor 
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
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
      Left            =   8520
      TabIndex        =   23
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter:"
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
      TabIndex        =   22
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Spy_Filter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All"
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
      Index           =   0
      Left            =   7440
      TabIndex        =   21
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Spy_Filter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-2"
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
      Index           =   1
      Left            =   7680
      TabIndex        =   20
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Spy_Filter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-1"
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
      Index           =   2
      Left            =   7920
      TabIndex        =   19
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Spy_Filter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Index           =   3
      Left            =   8160
      TabIndex        =   18
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Spy_Filter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+1"
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
      Index           =   4
      Left            =   8400
      TabIndex        =   17
      Top             =   3240
      Width           =   255
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
      TabIndex        =   15
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
   Begin VB.Shape Shape20 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label FishDisp_Enum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Not activated."
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
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape21 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Randomize casts"
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
      Height          =   225
      Left            =   1095
      TabIndex        =   10
      Top             =   2835
      Width           =   1455
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop at low cap"
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
      Height          =   225
      Left            =   1095
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Map-related features"
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
      TabIndex        =   8
      Top             =   1920
      Width           =   9375
   End
   Begin VB.Shape Shape23 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   3600
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   5880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Picks up any spears in reach of your char"
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
      Height          =   495
      Left            =   3705
      TabIndex        =   7
      Top             =   1020
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spear pickup"
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
      TabIndex        =   14
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autofishing"
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
      TabIndex        =   13
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Show fishy tiles"
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
      TabIndex        =   12
      Top             =   495
      Width           =   2415
   End
End
Attribute VB_Name = "frm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MdX As Long, MdY As Long, TrackID As Long
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

'SHOW FISHY TILES RELATED

    Sub FishDisp_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        FishDisp_Timer = FishDisp_Enabled
        If FishDisp_Timer = False Then FishDisp_Enum = "Not activated.": FishDisp_Display.Cls
        Smsg "Water tile analyzer enabled: " & FishDisp_Timer
    End Sub
    Sub FishDisp_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Fish_ShowFishyWater
        FishDisp_Enum = Fish_Map & " fish available." & vbCrLf & "Can carry " & (mReadLong(CH_Cap) \ 5) & "."
    End Sub

'AUTOMATIC FISHING RELATED

    Sub FishHack_Enabled_Click()
        If Compiled = True Then On Error Resume Next
        FishHack_Timer = FishHack_Enabled
        Smsg "Automatic fishing enabled: " & FishHack_Timer
    End Sub
    Sub FishHack_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Fish_CastRod FishHack_Randomize, FishHack_StopAtLowCap
        iSleep 800
        Stack_Items FOOD_FISH
    End Sub

'PICKUP RELATED
    
    Sub Spear_Enabled_Click()
        If Compiled Then On Error Resume Next
        Spear_Timer = Spear_Enabled
        Smsg "Spear pickup enabled: " & Spear_Timer
    End Sub
    Sub Spear_Timer_Timer()
        If Compiled = True Then On Error Resume Next
        Dim PlayerTile As TileData, a As Long, CurTile As Long, aX As Long, aY As Long, aZ As Long, ItemID As Long
        Map_Start = mReadLong(MAP_POINTER)
        PlayerTile = Map_TileInfo(Map_PlayerTileNum)
        'MsgBox PlayerTile.TileNum & " (" & PlayerTile.posX & "x" & PlayerTile.posY & "x" & PlayerTile.posZ & ")"
        For a = 0 To 2015
            If Map_TilePos(a, "z") = PlayerTile.posZ Then
                aX = Map_TilePos(a, "x")
                aY = Map_TilePos(a, "y")
                aX = aX - PlayerTile.posX
                aY = aY - PlayerTile.posY
                If aX = 17 Then aX = -1
                If aX = -17 Then aX = 1
                If aY = 13 Then aY = -1
                If aY = -13 Then aY = 1
                If aX > -2 And aX < 2 And aY > -2 And aY < 2 Then
                    CurTile = Map_Start + (a * Map_TileDist) + 4
                    For b = 0 To mReadLong(CurTile - 4) - 1
                        ItemID = mReadLong(CurTile + (Map_ObjectDist * b) + Map_ObjectIdDist)
                        If ItemID = 3277 Then
                            Label58 = "Spear @ "
                                If aY = -1 Then Label58 = Label58 & "Top "
                                If aY = 0 Then Label58 = Label58 & "Mid "
                                If aY = 1 Then Label58 = Label58 & "Bottom "
                                If aX = -1 Then Label58 = Label58 & "left"
                                If aX = 0 Then Label58 = Label58 & "mid"
                                If aX = 1 Then Label58 = Label58 & "right"
                                Label58 = Label58 & vbCrLf & "(" & Map_TilePos(a, "x") & "x" & Map_TilePos(a, "y") & ")"
                            aX = mReadLong(CH_X) + aX
                            aY = mReadLong(CH_Y) + aY
                            aZ = mReadLong(CH_Z)
                            sPck s2ba("0F 00 78 " & Hex(lbol(aX)) & " " & Hex(hbol(aX)) _
                                        & " " & Hex(lbol(aY)) & " " & Hex(hbol(aY)) & " " _
                                        & Hex(aZ) & " " & Hex(lbol(ItemID)) & " " & _
                                        Hex(hbol(ItemID)) & " " & b & " FF FF 05 00 00 01")
                            DoEvents
                        End If
                    Next
                End If
            End If
        Next
    End Sub

'SPY RELATED

    Sub Spy_Enabled_Click()
        If Compiled Then On Error Resume Next
        Spy_Timer = Spy_Enabled
        Smsg "Mob spy enabled: " & Spy_Timer
    End Sub
    Private Sub Spy_Filter_Click(Index As Integer)
        For a = Spy_Filter.LBound To Spy_Filter.UBound
            Spy_Filter(a).BackColor = &H80&
        Next
        Spy_Filter(Index).BackColor = &H8000&
        Label6 = "Mob Spy Lister ("
        If Index = 0 Then Label6 = Label6 & "all)" Else Label6 = Label6 & Index - 3
    End Sub
    Private Sub Spy_Timer_Timer()
        On Error GoTo 10
        If Spy_Lister.Enabled Then Spy_List.Clear
        For a = Spy_Tile.LBound To Spy_Tile.UBound
            Spy_Tile(a) = ""
        Next
        nZ = mReadLong(CH_Z)
        cID = mReadLong(CH_ID)
        For a = BL_Start To BL_End Step BL_Dist
            If mReadLong(a + BL_Vis) <> 0 And mReadLong(a + BL_HP) > 0 Then
                monID = mReadLong(a + BL_ID)
                If monID <> cID Then
                    MobZ = nZ - mReadLong(a + BL_Z)
                    If Spy_Filter(0).BackColor = &H8000& Then CorZ = 1
                    For b = 1 To 5
                        If Spy_Filter(b).BackColor = &H8000& Then If MobZ = b - 3 Then CorZ = 1
                    Next
                    If CorZ = 1 Or Spy_chkTrack Then
                        CorZ = 0
                        X = (mReadLong(a + BL_X) - mReadLong(CH_X)) + 8
                        Y = (mReadLong(a + BL_Y) - mReadLong(CH_Y)) + 6
                        If X >= 0 And X <= 18 And Y >= 0 And Y <= 14 Then
                            If Spy_Lister Then
                                Spy_List.AddItem mReadString(a + BL_Name)
                                Spy_List.ItemData(Spy_List.ListCount - 1) = mReadLong(a + BL_ID)
                            End If
                            If Spy_chkTrack Then If TrackID <> monID Then iDisp = 1
                            If iDisp = 0 Then Spy_Tile((18 * Y) + X + 1) = "g " & mReadLong(a + BL_ID)
                            iDisp = 0
                        End If
                    End If
                End If
            End If
        Next
        If Spy_chkTrack Then Spy_Stat TrackID
        Exit Sub
10      MsgBox "Attempted to access control array " & (18 * Y) + X + 1 & ", coordinates being " & X & "x" & Y & "."
    End Sub
    Private Sub Spy_List_Click()
        TrackID = Spy_List.ItemData(Spy_List.ListIndex)
        Spy_Stat TrackID
    End Sub
    Private Sub Spy_Tile_Click(Index As Integer)
        If Spy_Tile(Index) <> "" Then
            TrackID = Int(Mid$(Spy_Tile(Index), 3))
            Spy_Stat TrackID
        End If
    End Sub
    Private Sub Spy_Stat(ByVal vl As Long)
        For a = BL_Start To BL_End Step BL_Dist
            If mReadLong(a + BL_ID) = vl Then
                Spy_Name = mReadString(a + BL_Name)
                Spy_Health = mReadLong(a + BL_HP)
                Spy_Floor = mReadLong(a + BL_Z) - mReadLong(CH_Z)
            End If
        Next
    End Sub
