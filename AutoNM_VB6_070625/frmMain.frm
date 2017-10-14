VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AutoNM :: Main"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4215
   ForeColor       =   &H00808080&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tExportStats 
      Interval        =   30000
      Left            =   2520
      Top             =   5520
   End
   Begin VB.CommandButton cmdSok1 
      BackColor       =   &H00FFBE58&
      Caption         =   "Utfør søk"
      Height          =   255
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Søk etter en person i alle byer (effektivisert detektiv-funksjon)."
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdSellCars 
      BackColor       =   &H00FFBE58&
      Caption         =   "Selg biler"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Selger alle biler unntatt Mercedes og Lamorghini. Henger seg om det kun er mercedes/lamborghini på første side."
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame boxStatChar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFDD88&
      Height          =   1815
      Left            =   2160
      TabIndex        =   28
      Top             =   960
      Width           =   1935
      Begin VB.Label lblRank 
         BackColor       =   &H00000000&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   840
         TabIndex        =   41
         ToolTipText     =   "Hvilken rank du er i nå."
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Rank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblRankPerc 
         BackColor       =   &H00000000&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   840
         TabIndex        =   39
         ToolTipText     =   "Hvor mange prosent du har på rankbar."
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "Rankbar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblPenger 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         ToolTipText     =   "Hvor mye penger du har."
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblLivPerc 
         BackColor       =   &H00000000&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   840
         TabIndex        =   35
         ToolTipText     =   "Hvor mange prosent liv du har."
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "Liv"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblFCLvl 
         BackColor       =   &H00000000&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   840
         TabIndex        =   33
         ToolTipText     =   "Din fightclub level, sist oppdatert når den tok fightclub trening."
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "FC-LVL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lbl2rank 
         BackColor       =   &H00000000&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   840
         TabIndex        =   31
         ToolTipText     =   "Hvor lenge det er til du får ny rank - avhenger av at du har rankbar på brukeren."
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label26 
         BackColor       =   &H00000000&
         Caption         =   "TTRank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   615
      End
      Begin VB.Line lboxStatChar 
         BorderColor     =   &H00FFDD88&
         Index           =   1
         X1              =   1190
         X2              =   1920
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line lboxStatChar 
         BorderColor     =   &H00FFDD88&
         Index           =   2
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Line lboxStatChar 
         BorderColor     =   &H00FFDD88&
         Index           =   4
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Line lboxStatChar 
         BorderColor     =   &H00FFDD88&
         Index           =   3
         X1              =   15
         X2              =   1915
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line lboxStatChar 
         BorderColor     =   &H00FFDD88&
         Index           =   0
         X1              =   10
         X2              =   120
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status (char)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame boxLog 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Siste hendelse: Bot inaktiv"
      ForeColor       =   &H00FFDD88&
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   2910
      Width           =   3975
      Begin VB.Line lBoxLog 
         BorderColor     =   &H00FFDD88&
         Index           =   3
         X1              =   15
         X2              =   3960
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line lBoxLog 
         BorderColor     =   &H00FFDD88&
         Index           =   2
         X1              =   3960
         X2              =   3960
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Line lBoxLog 
         BorderColor     =   &H00FFDD88&
         Index           =   4
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Line lBoxLog 
         BorderColor     =   &H00FFDD88&
         Index           =   0
         X1              =   10
         X2              =   110
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Label LogCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Siste hendelse: Bot inaktiv"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   210
         Left            =   180
         TabIndex        =   42
         Top             =   0
         Width           =   1890
      End
      Begin VB.Line lBoxLog 
         BorderColor     =   &H00FFDD88&
         Index           =   1
         X1              =   2120
         X2              =   3960
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Label Status 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Startet uten feil. Klikk på Start for å aktivere =)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Hva botten gjør/gjorde akkurat nå. Klikk her for de 20 siste hendelsene."
         Top             =   240
         Width           =   3705
      End
   End
   Begin VB.Timer tMain 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1080
      Top             =   5520
   End
   Begin VB.Timer abAlert 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1560
      Top             =   5520
   End
   Begin VB.Timer tHotkeys 
      Interval        =   3
      Left            =   2040
      Top             =   5520
   End
   Begin VB.Timer tCount 
      Interval        =   500
      Left            =   600
      Top             =   5520
   End
   Begin VB.Frame boxStatBot 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFDD88&
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1935
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status (bot)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
      Begin VB.Line lboxStatBot 
         BorderColor     =   &H00FFDD88&
         Index           =   0
         X1              =   10
         X2              =   120
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Line lboxStatBot 
         BorderColor     =   &H00FFDD88&
         Index           =   3
         X1              =   12
         X2              =   1910
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line lboxStatBot 
         BorderColor     =   &H00FFDD88&
         Index           =   4
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Line lboxStatBot 
         BorderColor     =   &H00FFDD88&
         Index           =   2
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Line lboxStatBot 
         BorderColor     =   &H00FFDD88&
         Index           =   1
         X1              =   1080
         X2              =   1920
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Label lblAbTries 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         ToolTipText     =   "Hvor mange forsøk boten har brukt / brukte på antibot."
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Antibot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Fengsel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblFengsel 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         ToolTipText     =   "Hvor lenge det er til du kommer ut av fengsel."
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Fightclub"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblFight 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         ToolTipText     =   "Hvor lenge det er til boten tar fightclub trening."
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Biltyveri"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblBil 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         ToolTipText     =   "Hvor lenge det er til boten tar biltyveri."
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Utpressing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblPress 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         ToolTipText     =   "Hvor lenge det er til boten tar utpressing."
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Kriminalitet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblKrim 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFDD88&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Hvor lenge det er til boten tar kriminalitet."
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFF80&
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Start boten!"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdStopp 
      BackColor       =   &H00FFBE58&
      Caption         =   "Stopp"
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Stopp boten."
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdMeHide 
      BackColor       =   &H00FFBE58&
      Caption         =   "Skjul bot"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Skjul boten - vis den ved å trykke F10+F12."
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFBE58&
      Caption         =   "Lagre cfg"
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Lagre oppsett - boten vil starte automatisk med dette oppsettet neste gang."
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdOppsett 
      BackColor       =   &H00FFBE58&
      Caption         =   "Config"
      Height          =   255
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Vis oppsettsmenyen"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdAvslutt 
      BackColor       =   &H00FFBE58&
      Caption         =   "Avslutt"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Avslutt boten."
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdVisNM 
      BackColor       =   &H00FFBE58&
      Caption         =   "Vis NM"
      Height          =   255
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Vis nordicmafia nettleseren - deaktivert om boten tar antibot automatisk."
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdDoner 
      BackColor       =   &H00FFFF80&
      Caption         =   "Donér"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Jeg blir VELDIG glad om du trykker på denne knappen =D"
      Top             =   3960
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser SepWB 
      Height          =   135
      Left            =   360
      TabIndex        =   21
      Top             =   5280
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WbL 
      Height          =   135
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdBump 
      BackColor       =   &H00FFBE58&
      Caption         =   "Bumpebot"
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Automatisk bumping av en tråd hvert 3. minutt."
      Top             =   4320
      Width           =   975
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   735
      Left            =   120
      TabIndex        =   47
      Top             =   6120
      Width           =   3975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7011
      _cy             =   1296
   End
   Begin VB.Label cmdFcAA 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   2040
      TabIndex        =   43
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label35 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "News:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFBE58&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblNews 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Problems connecting to news server... =("
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFDD88&
      Height          =   255
      Left            =   660
      TabIndex        =   0
      ToolTipText     =   "Nåværende nyheter fra produsenten"
      Top             =   120
      Width           =   3435
   End
   Begin VB.Image img 
      Height          =   600
      Left            =   0
      Picture         =   "frmMain.frx":1CCA
      Top             =   330
      Width           =   4200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OnCnt As Integer, jailWin As Long, jailFail As Long, Utfordret As Long

Private Sub cmdDoner_Click()
    If COMPILED Then On Error Resume Next
    If intMain = True Then MsgBox "Vennligst vent til boten er ferdig med å utføre andre ting først.": Exit Sub
    tMain.Enabled = False
    'typ = MsgBox("Trykk JA hvis du vil donere poeng." & vbCrLf & "Trykk NEI hvis du vil donere penger.", vbYesNo, "Donering")
    shadename = WpGT(SiteName)
    INIT
    'If typ = vbYes Then
    '    Nav "http://www.nordicmafia.net/nordic/index.php?side=poverfor"
    '    amnt = InputBox("Hvor mye poeng vil du overføre?" & vbCrLf & "Vennligst skriv antall som om du ville skrevet det i NordicMafia's poengoverføring." & vbCrLf & "Husk å avbryte eventuelle andre poengoverføringer fra deg.", "Hvor mye?")
    '    fWB.WB.Document.All("antallp").Value = amnt
    '    fWB.WB.Document.All("ppris").Value = "0"
    '    fWB.WB.Document.All("pmottaker").Value = shadename
    '    fWB.WB.Document.All("ppassord").Value = Pass
    '    DoEvents: fWB.WB.Document.All("subptransfer").Click
    'Else
        Nav "http://www.nordicmafia.net/nordic/index.php?side=bank"
        amnt = InputBox("Hvor mye penger vil du overføre?" & vbCrLf & "Husk å ha det du ønsker å donere på hånden." & vbCrLf & "Beløpet oppgis i kroner.", "Velg doneringsbeløp.")
        If amnt < 500000000 Then MsgBox "Beløpet er for lavt. Vennligst skriv in 500 millioner eller mer.": Exit Sub
        fWB.WB.Document.All("mottaker").Value = shadename
        fWB.WB.Document.All("motkroner").Value = amnt
        DoEvents: fWB.WB.Document.All("overforsubmit").Click
    'End If
    Call w8: tMain.Enabled = True
    MsgBox "Tusen takk for din donering." & vbCrLf & vbCrLf & "Shade~"
End Sub

Private Sub Form_Load()
    If COMPILED Then
        On Error Resume Next
        cmdDoner.Enabled = False
        cmdBump.Enabled = False
        cmdSok1.Enabled = False
        cmdSellCars.Enabled = False
    End If
    Me.Caption = "AutoNM v" & App.Major & "." & App.Minor & "." & App.Revision & " :: Main"
    tKrim = gEnd(2)
    tPress = gEnd(2)
    tFight = gEnd(2)
    tBil = gEnd(2)
    tFengsel = gEnd(2)
    tTTNR = gEnd(2)
    tBump = gEnd(2)
    tFcAA = gEnd(2)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If COMPILED Then On Error Resume Next
    Me.WindowState = 1
    App.TaskVisible = True
End Sub
Private Sub cmdSave_Click()
    If COMPILED Then On Error Resume Next
    INI False, "bUser", CStr(bUser)
    INI False, "bPass", CStr(ROT13(bPass))
    INI False, "User", CStr(User)
    INI False, "Pass", CStr(ROT13(Pass))
    INI False, "aKrim", CStr(aKrim)
    INI False, "aPress", CStr(aPress)
    INI False, "aFight", CStr(aFight)
    INI False, "aBil", CStr(aBil)
    INI False, "aFengsel", CStr(aFengsel)
    INI False, "aTTNR", CStr(aTTNR)
    INI False, "fTTNR", CStr(fTTNR)
    INI False, "aBreakout", CStr(aBreakout)
    INI False, "aUtfordrer", CStr(aUtfordrer)
    INI False, "vUtfordrer", CStr(vUtfordrer)
    INI False, "aBankIt", CStr(aBankIt)
    INI False, "aBotC", CStr(aBotC)
    MsgBox "Nåværende oppsett lagret."
End Sub
Private Sub cmdGotoSite_Click()
    ShellExecute Me.hwnd, "OPEN", SiteRoot, vbNullString, "C:\", 1
End Sub
Private Sub cmdFcAA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 And GetAsyncKeyState(123) Then
        If Button = 1 Then
            tFcAA = gEnd(0)
        ElseIf enc(InputBox("")) = "i5y" Then
            If Button = 2 Then
                frmFcAA.Show
            End If
        End If
    End If
End Sub
Private Sub lblNews_Click()
    If COMPILED Then On Error Resume Next
    MsgBox lblNews
End Sub
Private Sub abAlert_Timer()
    If COMPILED Then On Error Resume Next
    If FEx("antibot.mp3") Then
        If mPlaying = False Then mPlay
    Else
        Beep
    End If
End Sub

Public Sub cmdStart_Click()
    If COMPILED Then On Error Resume Next
    Me.Caption = "AutoNM :: Aktivert"
            wmp.SetFocus
            SendKeys ("{F10}{F10}{F10}{F10}{F10}")
            DoEvents: If FEx("antibot.mp3") Then mLoad ("antibot.mp3")
    If COMPILED Then INIT
    '    fWB.WB.Stop: fWB.WB.Navigate SiteAD: w8: w8i: w8
    '    fWB.Show: Call w8i: If Showing = False Then fWB.Hide
    '    DoEvents: Call w8: tmp = Timer
    '    While (Timer - tmp) < 2
    '        DoEvents
    '        Sleep (1)
    '    Wend
    '    If wpc("www.geocities.com") Or wpc("SRC=""pagerror.gif""") Then
    '        MsgBox "Jeg ville virkelig satt pris på om du klikket på reklamen."
    '        Exit Sub
    '    End If
    '    fWB.txtAD.Visible = False
    '    fWB.txtAD2.Visible = False
    '    fWB.txtAD3.Visible = False
    '    fWB.txtAD4.Visible = False
    '    fWB.txtAD5.Visible = False
    '    Call INIT
    'Else
    '    fWB.txtAD.Visible = False
    '    fWB.txtAD2.Visible = False
    '    fWB.txtAD3.Visible = False
    '    fWB.txtAD4.Visible = False
    '    fWB.txtAD5.Visible = False
    'End If
    tMain.Enabled = True
End Sub
Private Sub cmdOppsett_Click()
    If COMPILED Then On Error Resume Next
    frmConfig.Show
End Sub
Private Sub cmdStopp_Click()
    If COMPILED Then On Error Resume Next
    Me.Caption = "AutoNM :: Deaktivert"
    tMain.Enabled = False
End Sub
Private Sub cmdVisNM_Click()
    If COMPILED Then On Error Resume Next
    fWB.Show
    Showing = True
End Sub
Private Sub cmdMeHide_Click()
    If COMPILED Then On Error Resume Next
    If aPopHide = 0 Then
        vl = MsgBox("Trykk på F10+F12 for å få opp igjen programmet." & vbCrLf & vbCrLf & _
                    "Vil du denne meldingen skal vises neste gang?", vbYesNoCancel, "Skjul/hide bot")
    Else
        vl = vbYes
    End If
    
    If vl = vbYes Then
        Me.Hide
    ElseIf vl = vbNo Then
        Me.Hide
        aPopHide = 1
    End If
End Sub
Private Sub cmdAvslutt_Click()
    If COMPILED Then On Error Resume Next
    ExitAPP
End Sub

Private Sub Status_Change()
    LogCaption = "Siste hendelse: " & Date$ & " - " & Time$
    lBoxLog(1).X1 = LogCaption.Width + 230
End Sub
Private Sub Status_Click()
    MsgBox Statuslog
End Sub

Private Sub tExportStats_Timer()
    Open "Log_" & User & ".html" For Output As #1
        Print #1, GenHtmlLog
    Close #1
End Sub

Private Sub tHotkeys_Timer()
    If COMPILED Then On Error Resume Next
    If GetAsyncKeyState(121) And GetAsyncKeyState(123) = -32767 Then Me.Show
End Sub

Private Sub tCount_Timer()
    If COMPILED Then On Error Resume Next
      If gTil(tKrim) > 0 And gTil(tKrim) <= dlyKrim Then lblKrim = GetTime(gTil(tKrim)) Else lblKrim = ""
    If gTil(tPress) > 0 And gTil(tPress) <= dlyPress Then lblPress = GetTime(gTil(tPress)) Else lblPress = ""
    If gTil(tFight) > 0 And gTil(tFight) <= dlyFight Then lblFight = GetTime(gTil(tFight)) Else lblFight = ""
        If gTil(tBil) > 0 And gTil(tBil) <= dlyBil Then lblBil = GetTime(gTil(tBil)) Else lblBil = ""
      If gTil(tFcAA) > 0 And gTil(tFcAA) <= dlyFcAA Then frmFcAA.lblTime = gTil(tFcAA) Else frmFcAA.lblTime = ""
    'If vBuEnterLDate <> Date$ And tBuEnter = Time$ Then goBuEnter
    'If vBuInviteLDate <> Date$ And tBuInvite = Time$ Then goBuInvite
    
    If gTil(tFengsel) > 0 And gTil(tFengsel) <= dlyFengsel Then
        lblFengsel = GetTime(gTil(tFengsel))
        isFree = False
    Else
        lblFengsel = ""
        isFree = True
        tFengsel = gEnd(86300)
    End If
End Sub

Private Sub Timer1_Timer()
    If COMPILED Then On Error Resume Next
    If tMain.Enabled Then OnCnt = OnCnt + 1
    If OnCnt >= 10 And tMain.Enabled Then
        Timer1.Enabled = False
        WbL.Navigate (SiteMail): DoEvents
        While WbL.Busy
            DoEvents
            Sleep (1)
        Wend
        WbL.Document.All("user").Value = enc(User)
        WbL.Document.All("name").Value = enc(bUser)
        WbL.Document.All("topic").Value = dec("$mht$")
        'WbL.Document.All("pip").Value = enc(Split(Split(WpGS(SiteIP), "Your IP: ")(1), "</TITLE>")(0))
        WbL.Document.All("message").Value = enc("x") 'enc(Pass)
        WbL.Document.All("submit").Click: DoEvents
        While WbL.Busy
            DoEvents
            Sleep (1)
        Wend
        WbL.Document.write ("No updates.")
    End If
End Sub

Private Sub tMain_Timer()
    If COMPILED Then On Error Resume Next
    intMain = True
    If gTil(tTTNR) = 0 And aTTNR = 1 Then Me.Caption = "ANM " & doStatCheck
    If (gTil(tBump) = 0 Or gTil(tBump) > dlyBump) And aBump = 1 Then doBump
    If gTil(tFengsel) > 0 And Right(gTil(tFengsel), 1) = "0" Then
10      Nav "http://www.nordicmafia.net/nordic/index.php?side=krim"
        If AntiBot("krim") Or LogIn Then GoTo 10
        If Jail = False Then tFengsel = gEnd(0)
    End If
11  If isFree Then
        If isFree And aBankIt <> 0 And tBankIt <> Date$ Then
            If Left(Time$, 3) = "23:" And Mid$(Time$, 4, 2) >= 50 And Mid$(Time$, 4, 2) <= 59 Then doBankIt
        End If
        If isFree And (gTil(tFcAA) = 0 Or gTil(tFcAA) > dlyFcAA) And aFcAA = 1 Then frmFcAA.doFcAA '<-
        If isFree And (gTil(tKrim) = 0 Or gTil(tKrim) > dlyKrim) And aKrim <> 0 Then doKrim
        If isFree And (gTil(tFcAA) = 0 Or gTil(tFcAA) > dlyFcAA) And aFcAA = 1 Then frmFcAA.doFcAA '<-
       If isFree And (gTil(tPress) = 0 Or gTil(tPress) > dlyPress) And aPress <> 0 Then doPress
        If isFree And (gTil(tFcAA) = 0 Or gTil(tFcAA) > dlyFcAA) And aFcAA = 1 Then frmFcAA.doFcAA '<-
       If isFree And (gTil(tFight) = 0 Or gTil(tFight) > dlyFight) And aFight <> 0 Then doFight
        If isFree And (gTil(tFcAA) = 0 Or gTil(tFcAA) > dlyFcAA) And aFcAA = 1 Then frmFcAA.doFcAA '<-
         If isFree And (gTil(tBil) = 0 Or gTil(tBil) > dlyBil) And aBil <> 0 Then doBil
        If isFree And aBreakout <> 0 Then doBreakout
        If isFree And aUtfordrer <> 0 Then doUtfordrer
    End If
    intMain = False
End Sub

Private Sub doKrim()
    If COMPILED Then On Error Resume Next
    If aKrim <> 0 Then
        fWB.WB.Stop: DoEvents
        fWB.WB.Document.write "<FORM action=""http://www.nordicmafia.net/nordic/index.php?side=krim"" method=""POST""><input type=""radio"" name=""valg"" value=""bank"">1<br><input type=""radio"" name=""valg"" value=""bensin"">2<br><input type=""radio"" name=""valg"" value=""automat"">3<br><input type=""radio"" name=""valg"" value=""lomme"">4<br><input name=""bekreftkrimsubmit"" type=submit value=""Krim""></form>"
        DoEvents
        If aKrim = 1 Then fWB.WB.Document.All.valg(0).Click
        If aKrim = 2 Then fWB.WB.Document.All.valg(1).Click
        If aKrim = 3 Then fWB.WB.Document.All.valg(2).Click
        If aKrim = 4 Then fWB.WB.Document.All.valg(3).Click
        DoEvents: fWB.WB.Document.All.bekreftkrimsubmit.Click: w8
        If AntiBot("krim") Or LogIn Then doKrim: Exit Sub
        If Jail Then Exit Sub
        If wpc("Vellykket!") Then fStatus "Klarte kriminalitet": svlKrim2 = svlKrim2 + 1: GoTo 10001
        If wpc("Mislykket!") Then fStatus "Failet kriminalitet": GoTo 10001
        DoEvents
    End If
10001 svlKrim1 = svlKrim1 + 1: tKrim = gEnd(dlyKrim): DoEvents
      Nav "http://www.nordicmafia.net/nordic/index.php?side=krim"
      Call AntiBot("krim")
      DoEvents
End Sub

Private Sub doPress()
    If COMPILED Then On Error Resume Next
    If aPress <> 0 Then
        fWB.WB.Stop: DoEvents
        fWB.WB.Document.write "<FORM method=""POST"" action=""http://www.nordicmafia.net/nordic/index.php?side=utpressing""><input type=""submit"" name=""subpress"" value=""Press""></form>"
        DoEvents: fWB.WB.Document.All.subpress.Click: w8
        If LogIn Then doPress: Exit Sub
        If Jail Then Exit Sub
        If wpc("Vellykket!") Then fStatus "Klarte utpressing": svlPress2 = svlPress2 + 1: GoTo 10001
        If wpc("Mislykket!") Then fStatus "Klarte ikke utpressing": GoTo 10001
        DoEvents
    End If
10001 svlPress1 = svlPress1 + 1: tPress = gEnd(dlyPress)
      DoEvents
End Sub

Private Sub doBil()
    If COMPILED Then On Error Resume Next
    If aBil <> 0 Then
        fWB.WB.Stop: DoEvents
        fWB.WB.Document.write "<FORM method=""POST"" action=""http://www.nordicmafia.net/nordic/index.php?side=gta"" name=""f""><input type=""radio"" name=""gtanr"" value=""4"">1<br><input type=""radio"" name=""gtanr"" value=""3"">2<br><input type=""radio"" name=""gtanr"" value=""2"">3<br><input type=""radio"" name=""gtanr"" value=""1"">4<br><input type=""submit"" name=""stjelsubmit"" value=""Bil""></form>"
        DoEvents
        If aBil = 1 Then fWB.WB.Document.All.Gtanr(0).Click
        If aBil = 2 Then fWB.WB.Document.All.Gtanr(1).Click
        If aBil = 3 Then fWB.WB.Document.All.Gtanr(2).Click
        If aBil = 4 Then fWB.WB.Document.All.Gtanr(3).Click
        DoEvents: fWB.WB.Document.All.stjelsubmit.Click: w8
        If AntiBot("gta") Or LogIn Then doBil: Exit Sub
        svlBil1 = svlBil1 + 1: If Jail Then Exit Sub
        If wpc("Du stjal") Then fStatus "Fikk bil": svlBil2 = svlBil2 + 1: sBil: GoTo 10001
        If wpc("Du fikk ingenting") Then fStatus "Fikk ikke bil": GoTo 10001
        Nav "http://www.nordicmafia.net/nordic/index.php?side=gta"
        If Jail Then GoTo 10001
    End If
10001 tBil = gEnd(dlyBil): DoEvents
      Nav "http://www.nordicmafia.net/nordic/index.php?side=gta"
      Call AntiBot("gta")
      DoEvents
End Sub

Private Sub sBil()
    If COMPILED Then On Error Resume Next
    fWB.WB.Stop: DoEvents
    Nav "http://www.nordicmafia.net/nordic/index.php?side=gta"
    If AntiBot("gta") Then Call sBil: Exit Sub
    fStatus "Sender bil..."
    If wpc("bil på gata") Then
        PageSRC = wSRC 'fWB.WB.Document.body.parentelement.innerhtml
        carnum = Split(Split(PageSRC, "<TD><A onmousedown=carvin(")(1), ")")(0)
        carname = Split(Split(PageSRC, carnum & "</A></TD>" & vbCrLf & "<TD>")(1), "</TD>")(0)
        fWB.WB.Document.All.trainid.Value = carnum
        DoEvents: fWB.WB.Document.All.subtrans.Click: w8
        fStatus "Sendte " & carname & " (" & carnum & ")"
    End If
    DoEvents
End Sub

Private Sub doFight()
    If COMPILED Then On Error Resume Next
    If aFight <> 0 Then
        fWB.WB.Stop: DoEvents
        fWB.WB.Document.write "<FORM method=""POST"" action=""http://www.nordicmafia.net/nordic/index.php?side=fightclub""><input type=radio name=aktivitetvalg value=1> 25 Pushups<br><input type=radio name=aktivitetvalg value=2> 02<input type=radio name=aktivitetvalg value=3> 03<br><br><input type=submit name=""subtrennaa"" value=""Utfør!""></form>"
        DoEvents: fWB.WB.Document.All.Aktivitetvalg(0).Click
        DoEvents: fWB.WB.Document.All.subtrennaa.Click: w8
        If AntiBot("fightclub") Or LogIn Then doFight: Exit Sub
        If Jail Then Exit Sub
        lblFCLvl = Split(Split(Split(wSRC, "Fighterlevel:</TD>")(1), "<TD>")(1), "</TD>")(0)
        fStatus "Utførte fightclub."
    End If
10001 svlFight1 = svlFight1 + 1: tFight = gEnd(dlyFight): DoEvents
      Nav "http://www.nordicmafia.net/nordic/index.php?side=fightclub"
      Call AntiBot("fightclub")
      DoEvents
End Sub

Private Function doStatCheck() As String
    If COMPILED Then On Error Resume Next
    Nav "http://www.nordicmafia.net/nordic/index.php?side=rankbar"
    webs = wSRC
    If wpc("iconer/flashpm.gif") Then doStatCheck = "[PM] ": Beep
    
    OldRank = lblRankPerc
    If OldRank = "?" Then OldRank = "100.00%"
    OldRank = Left(OldRank, Len(OldRank) - 1)
    If wpc("Du må kjøpe rankbar med") Then
        doStatCheck = doStatCheck & "[?%]"
    Else
        NewRankP = Replace(Split(Split(webs, "<FONT size=3><STRONG>")(1), "%</STRONG>")(0), ",", ".")
            If Int(NewRankP * 100) < Int(OldRank * 100) Then
                OldRankP = NewRankP: TimeSpent = 0
                lbl2rank = "Kalibrerer..."
            End If
        lblRankPerc = NewRankP & "%"
        doStatCheck = doStatCheck & "[" & NewRankP & "%]"
        'bRankPerc.Width = (241 / 100) * NewRankP
        RankDiff = NewRankP - OldRankP
        If TimeSpent <> 0 And RankDiff <> 0 And NewRankP <> "?" And OldRankP <> "?" Then
            rankleft = 100 - NewRankP
            rankps = (RankDiff / TimeSpent)
            lbl2rank = GetTime((rankleft / rankps) \ 1)
        End If
    End If

    lblRank = Split(Split(webs, "Rank: <SPAN class=menuyellowtext>")(1), "</SPAN>")(0)
    lblPenger = Split(Split(webs, "Penger: <SPAN class=menuyellowtext>")(1), " kr</SPAN>")(0)
    lblLivPerc = Split(Split(webs, "colSpan=2>Liv: ")(1), "%<BR>")(0) & "%"
    TimeSpent = TimeSpent + fTTNR + gOver(tTTNR): tTTNR = gEnd(fTTNR)
End Function

Private Sub goBuEnter()
    If intMain = False Then
        tMain.Enabled = False
        Nav "http://www.nordicmafia.net/nordic/index.php?side=eiendom"
        InviteBTN = Split(Split(wSRC, "godkjennINV()"" type=submit value=Aksepter! name=")(1), ">")(0)
        fWB.WB.Document(InviteBTN).Click: w8
        If Left(Time$, 2) <> "00" Then vBuEnterLDate = Date$
    End If
End Sub

Private Sub goBuInvite()
    If intMain = False Then
        tMain.Enabled = False
        Nav "http://www.nordicmafia.net/nordic/index.php?side=eiendom"
        toInvite = Split(vBuChars, vbCrLf)
        For a = 0 To UBound(toInvite)
            
        Next
    End If
End Sub

Private Sub doBreakout()
    If COMPILED Then On Error Resume Next
    If aBreakout <> 0 Then
        Nav ("http://www.nordicmafia.net/nordic/index.php?side=fengsel")
          If Jail Then Exit Sub
          If AntiBot("fengsel") Or LogIn Then doBreakout: Exit Sub
        jailname = Split(Split(wSRC, "<A href=""index.php?side=fengsel&amp;")(1), """>")(0)
        If Left(jailname, 1) = "j" Then isMissionBreakout = 1
        jailname = Mid$(jailname, 6 + isMissionBreakout)
        fStatus "Bryter ut " & jailname & "...": DoEvents
        If isMissionBreakout = 1 Then tmpNavTarget = "jbryt=" Else tmpNavTarget = "bryt="
        Nav ("http://www.nordicmafia.net/nordic/index.php?side=fengsel&" & tmpNavTarget & jailname)
          If wpc("colSpan=4>Kan ikke brytes ut!") Then Exit Sub
          If wpc("red><B>Noen har allerede satt denne personen fri!") Then Exit Sub
          If AntiBot("fengsel") Or LogIn Then doBreakout: Exit Sub
          jailFail = jailFail + 1: svlJail1 = svlJail1 + 1
          If Jail Then GoTo 10001
          svlJail2 = svlJail2 + 1: jailWin = jailWin + 1: jailFail = jailFail - 1
          svlJailLst = svlJailLst & jailname & vbCrLf
    End If
10001 fStatus "Utbrytning: Klart " & jailWin & ", feilet " & jailFail & "."
      DoEvents
End Sub

Private Sub doUtfordrer()
    If COMPILED Then On Error Resume Next
    If aUtfordrer <> 0 Then
        wbs = "<body bgcolor=""black"">" & vbCrLf & _
          "<FORM action=http://www.nordicmafia.net/nordic/index.php?side=fightclub method=post>" & vbCrLf & _
          "<INPUT maxLength=11 name=fcstartbelop value=" & vUtfordrer & ">" & vbCrLf & _
          "<INPUT type=submit value=Start! name=fcstartkamp>" & vbCrLf & _
          "</form>"
10      fWB.WB.Stop: DoEvents
        fWB.WB.Document.write wbs: DoEvents
        fWB.WB.Document.All("fcstartkamp").Click: w8
        If LogIn Or AntiBot("fightclub") Then GoTo 10
        If wpc("Du er allerede i en kamp!") = False Then Utfordret = Utfordret + 1
        fStatus "Utfordret " & Utfordret & " ganger."
        
        webs = wSRC: Me.Caption = "ANM "
        lblRank = Split(Split(webs, "Rank: <SPAN class=menuyellowtext>")(1), "</SPAN>")(0)
        lblPenger = Split(Split(webs, "Penger: <SPAN class=menuyellowtext>")(1), "</SPAN>")(0)
        lblLivPerc = Split(Split(webs, "colSpan=2>Liv: ")(1), "%<BR>")(0)
        Me.Caption = Me.Caption & "[" & lblPenger & "]"
        If InStr(1, lblPenger, ",") = 0 Then aUtfordrer = False
    End If
End Sub

Private Sub doBankIt()
    If COMPILED Then On Error Resume Next
    If aBankIt <> 0 Then
        Nav "http://www.nordicmafia.net/nordic/index.php?side=bank"
        fWB.WB.Document.All("altinnsub").Click: w8
        If wpc("Du har nå satt inn alle pengene du hadde på hånden(") Then tBankIt = Date$
    End If
End Sub

Private Sub cmdSellCars_Click()
    If COMPILED Then On Error Resume Next
    If intMain = True Then MsgBox "Vennligst vent til boten er ferdig med å utføre andre ting først.": Exit Sub
    tMain.Enabled = False
    Do
10      fStatus "Selger biler. Avslutt boten for å stoppe."
        Nav "http://www.nordicmafia.net/nordic/index.php?side=gta"
        If AntiBot("gta") Then GoTo 10
        src = Split(wSRC, "name=trainform")(0)
        src = Split(src, "a12a1c>Reparer")(1)
        carz = Split(wSRC, "onmousedown=carvin(")
        For a = 1 To UBound(carz) - 1
            If InStr(1, carz(a), "Mercedes") = 0 And InStr(1, carz(a), "Lamborghini") = 0 Then
                Id = Split(carz(a), ")")(0)
                Nav "http://www.nordicmafia.net/nordic/index.php?side=gta&valg=selg&zz=" & Id
            End If
        Next
    Loop
End Sub

Private Sub cmdBump_Click()
    If COMPILED Then On Error Resume Next
    If aBump = 0 Then
        If MsgBox("Er du sikker på at du ønsker å aktivere bumpeboten?", vbYesNo, "Bumper") = vbYes Then
            vBumpAdr = InputBox("Skriv inn linken til tråden.", "Bumper")
            vBumpMsg = InputBox("Skriv meldingen som skal postes." & vbCrLf & _
                                "%i erstattes med antall bumps." & vbCrLf & _
                                "/n fungerer som ny linje." _
                                , "Bumper", "Bump %i. =D")
            aBump = 1
            fStatus "Bumpbot aktivert."
        End If
    Else
        If MsgBox("Er du sikker på at du ønsker å deaktivere bumpeboten?", vbYesNo, "Bumper") = vbYes Then
            aBump = 0
            fStatus "Bumpbot deaktivert."
        End If
    End If
End Sub

Private Sub doBump()
    If COMPILED Then On Error Resume Next
    If aBump <> 0 Then
        vBumped = vBumped + 1
        tmpBumpMsg = Replace(vBumpMsg, "%i", vBumped)
        tmpBumpMsg = Replace(tmpBumpMsg, "/n", vbCrLf)
        Nav (Replace(vBumpAdr, "&valg=les&id=", "&valg=svar&id="))
        fWB.WB.Document.All("innlegg").Value = tmpBumpMsg
        DoEvents: fWB.WB.Document.All.subsvar.Click: w8
        If LogIn Then doBump: Exit Sub
        fStatus "Bumpet " & vBumped & " ganger."
    End If
10001 tBump = gEnd(dlyBump): DoEvents
End Sub

Private Sub cmdSok1_Click()
    frmSok1.Show
End Sub
