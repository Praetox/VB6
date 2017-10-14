VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00000000&
   Caption         =   "AutoNM :: Oppsett"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form2"
   ScaleHeight     =   4470
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFDD88&
      Caption         =   "Bruk valgt oppsett"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3050
      Width           =   3615
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      Caption         =   "Spillerstatus"
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
      Height          =   1095
      Left            =   5880
      TabIndex        =   52
      Top             =   3240
      Width           =   1695
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00000000&
         Caption         =   "Deaktivert"
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
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00000000&
         Caption         =   "Aktivert"
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
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox optValStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   240
         Left            =   960
         TabIndex        =   53
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Refresh"
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
         TabIndex        =   56
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00000000&
      Caption         =   "Bunker - ikke ferdig!"
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
      Height          =   2895
      Left            =   5880
      TabIndex        =   40
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Utfør:"
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
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   240
         Left            =   960
         TabIndex        =   48
         Text            =   "23:59"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   240
         Left            =   960
         TabIndex        =   45
         Text            =   "23:50"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Utfør:"
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
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   240
         Left            =   960
         TabIndex        =   43
         Text            =   "1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   240
         Left            =   960
         TabIndex        =   42
         Text            =   "1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text4 
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
         ForeColor       =   &H00FFBE58&
         Height          =   720
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Text            =   "frmConfig.frx":0000
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFDD88&
         X1              =   120
         X2              =   1560
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Inviter til bunker"
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
         TabIndex        =   51
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Aksepter bunker"
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
         TabIndex        =   50
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Caption         =   "Millioner:"
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
         TabIndex        =   47
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label33 
         BackColor       =   &H00000000&
         Caption         =   "Timer:"
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
         TabIndex        =   46
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00000000&
      Caption         =   "Legg ut på fightclub"
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
      Height          =   1095
      Left            =   3960
      TabIndex        =   35
      Top             =   3240
      Width           =   1695
      Begin VB.TextBox optValUtfordrer 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFBE58&
         Height          =   240
         Left            =   840
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optUtfordrer 
         BackColor       =   &H00000000&
         Caption         =   "Aktivert"
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
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optUtfordrer 
         BackColor       =   &H00000000&
         Caption         =   "Deaktivert"
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
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackColor       =   &H00000000&
         Caption         =   "Beløp:"
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
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00FFDD88&
      Height          =   1575
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optKrim 
         BackColor       =   &H00000000&
         Caption         =   "Deaktivert"
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
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optKrim 
         BackColor       =   &H00000000&
         Caption         =   "Gammel dame"
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
         Index           =   4
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optKrim 
         BackColor       =   &H00000000&
         Caption         =   "Spilleautomat"
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
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optKrim 
         BackColor       =   &H00000000&
         Caption         =   "Bensinstasjon"
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
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optKrim 
         BackColor       =   &H00000000&
         Caption         =   "Banken"
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
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00FFDD88&
      Height          =   1575
      Left            =   2040
      TabIndex        =   23
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optBil 
         BackColor       =   &H00000000&
         Caption         =   "Deaktivert"
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
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optBil 
         BackColor       =   &H00000000&
         Caption         =   "Bil på gata"
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
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optBil 
         BackColor       =   &H00000000&
         Caption         =   "Priv. parkering"
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
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optBil 
         BackColor       =   &H00000000&
         Caption         =   "Se etter nøkler"
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
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optBil 
         BackColor       =   &H00000000&
         Caption         =   "Off. parkering"
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
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
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
      ForeColor       =   &H00FFDD88&
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   1695
      Begin VB.OptionButton optPress 
         BackColor       =   &H00000000&
         Caption         =   "Aktivert"
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optPress 
         BackColor       =   &H00000000&
         Caption         =   "Deaktivert"
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
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Fightclub Trening"
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
      Height          =   855
      Left            =   2040
      TabIndex        =   17
      Top             =   1920
      Width           =   1695
      Begin VB.OptionButton optFC 
         BackColor       =   &H00000000&
         Caption         =   "Deaktivert"
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
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optFC 
         BackColor       =   &H00000000&
         Caption         =   "Aktivert"
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
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
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
      ForeColor       =   &H00FFDD88&
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1695
      Begin VB.OptionButton optJail 
         BackColor       =   &H00000000&
         Caption         =   "Vent til t. går ut"
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optJail 
         BackColor       =   &H00000000&
         Caption         =   "Kjøp ut (5 p.)"
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
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      Caption         =   "Utbrytning"
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
      Height          =   855
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Width           =   1695
      Begin VB.OptionButton optBreakout 
         BackColor       =   &H00000000&
         Caption         =   "Aktivert"
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
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optBreakout 
         BackColor       =   &H00000000&
         Caption         =   "Deaktivert"
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
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
      ForeColor       =   &H00FFDD88&
      Height          =   1575
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optBotC 
         BackColor       =   &H00000000&
         Caption         =   "Alarm"
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
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optBotC 
         BackColor       =   &H00000000&
         Caption         =   "Popup"
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optBotC 
         BackColor       =   &H00000000&
         Caption         =   "Popup + alarm"
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
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optBotC 
         BackColor       =   &H00000000&
         Caption         =   "Automatisk"
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
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optBotC 
         BackColor       =   &H00000000&
         Caption         =   "Ingen advarsel"
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
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Sett penger i bank"
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
      Height          =   1095
      Left            =   3960
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
      Begin VB.OptionButton optBankit 
         BackColor       =   &H00000000&
         Caption         =   "Deaktivert"
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optBankit 
         BackColor       =   &H00000000&
         Caption         =   "Aktivert"
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
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label28 
         BackColor       =   &H00000000&
         Caption         =   "Utføres:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label29 
         BackColor       =   &H00000000&
         Caption         =   "23:50:00"
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
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    If COMPILED Then On Error Resume Next
    Call Saves
End Sub
Public Sub Saves()
    If COMPILED Then On Error Resume Next
                                                            'doevents: frmmain.dbg = frmMain.dbg & " iconSaves": DoEvents
    For a = 0 To 4
        If optKrim(a).Value = True Then aKrim = a
        If optBil(a).Value = True Then aBil = a
        If optBotC(a).Value = True Then aBotC = a
    Next
    If optPress(1).Value = True Then aPress = 1 Else aPress = 0
    If optFC(1).Value = True Then aFight = 1 Else aFight = 0
    If optJail(1).Value = True Then aFengsel = 1 Else aFengsel = 0
    'aPress = optPress.Value
    'aFight = optFC.Value
    'aFengsel = optJail.Value
    fTTNR = optValStatus.Text
    If optStatus(1).Value = True Then aTTNR = 1 Else aTTNR = 0
    'aTTNR = optStatus.Value
    If optBreakout(1).Value = True Then aBreakout = 1 Else aBreakout = 0
    If optUtfordrer(1).Value = True Then aUtfordrer = 1 Else aUtfordrer = 0
    vUtfordrer = optValUtfordrer
    If optBankit(1).Value = True Then aBankIt = 1 Else aBankIt = 0
    Unload Me
                                                            'doevents: frmmain.dbg = frmMain.dbg & " oconSaves": DoEvents
End Sub

Private Sub Form_Load()
    If COMPILED Then On Error Resume Next
                                                            'doevents: frmmain.dbg = frmMain.dbg & " iconLoad": DoEvents
    Me.Caption = "AutoNM v" & App.Major & "." & App.Minor & "." & App.Revision & " :: Config"
    optKrim(aKrim).Value = True
    optPress(aPress).Value = True
    optFC(aFight).Value = True
    optBil(aBil).Value = True
    optJail(aFengsel).Value = True
    optBotC(aBotC).Value = True
    optStatus(aTTNR).Value = True
    optValStatus.Text = fTTNR
    optBreakout(aBreakout).Value = True
    optUtfordrer(aUtfordrer).Value = True
    optValUtfordrer = vUtfordrer
    optBankit(aBankIt).Value = True
                                                            'doevents: frmmain.dbg = frmMain.dbg & " oconLoad": DoEvents
End Sub

