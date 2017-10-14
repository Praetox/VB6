VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H007E543C&
   Caption         =   "Postal v"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdvertise 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Become a postal advertiser!"
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Advertise for Postal by sending PMs to 1000 online players."
      Top             =   240
      Width           =   2480
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   9600
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tAutosync 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10560
      Top             =   6600
   End
   Begin VB.Timer tSender 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10080
      Top             =   6600
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Outbox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   10440
      TabIndex        =   29
      Top             =   1440
      Width           =   2535
      Begin VB.ListBox out_List 
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4680
         ItemData        =   "Postal_frmMain.frx":0000
         Left            =   120
         List            =   "Postal_frmMain.frx":0002
         TabIndex        =   30
         ToolTipText     =   "The list of pending messages. It will send messages from top down, unless there are prioritated messages waiting."
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lbOutbox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ".:: Outbox ::."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label out_Stats 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 queued. Next: Now."
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
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00D6C2AC&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   0
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Application status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   6960
      Width           =   12855
      Begin VB.Label lbAppStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Application status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   12855
      End
      Begin VB.Label ST2 
         BackColor       =   &H00FFEFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Minor happening"
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
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Minor happenings."
         Top             =   480
         Width           =   12615
      End
      Begin VB.Label ST1 
         BackColor       =   &H00FFEFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Major happening"
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
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Major happenings."
         Top             =   220
         Width           =   12615
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00D6C2AC&
         FillStyle       =   0  'Solid
         Height          =   860
         Left            =   0
         Top             =   0
         Width           =   12855
      End
   End
   Begin VB.Frame frmOptions 
      BackColor       =   &H007E543C&
      BorderStyle     =   0  'None
      Caption         =   "Inbox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00EBD8C3&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Show the message composer, prepared for a new message."
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDelBy 
         BackColor       =   &H00EBD8C3&
         Caption         =   "Delete by..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         MaskColor       =   &H00FFBFA0&
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Delete messages by topic or author name, for example all fight club messages."
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00EBD8C3&
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
         Height          =   735
         Left            =   120
         MaskColor       =   &H00FFBFA0&
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Refresh the list over messages, and download any new messages to Postal."
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00EBD8C3&
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         MaskColor       =   &H00FFBFA0&
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Filter your messages. Same as the ""Show ..."" buttons in the inbox field."
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteRead 
         BackColor       =   &H00EBD8C3&
         Caption         =   "Delete read"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         MaskColor       =   &H00FFBFA0&
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Delete all the PMs that you have read."
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Outbox opt."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   47
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inbox options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   3975
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00D6C2AC&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   3975
      End
      Begin VB.Shape Shape9 
         FillColor       =   &H00D6C2AC&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   4200
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame frmInbox 
      BorderStyle     =   0  'None
      Caption         =   "Inbox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   2535
      Begin VB.CommandButton in_Filter 
         BackColor       =   &H00EBD8C3&
         Caption         =   "Show read"
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
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Show only messages that you have already opened."
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CommandButton in_Filter 
         BackColor       =   &H00EBD8C3&
         Caption         =   "Show unread"
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
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Show only messages that has yet to be read."
         Top             =   4080
         Width           =   2295
      End
      Begin VB.CommandButton in_Filter 
         BackColor       =   &H00EBD8C3&
         Caption         =   "Show all"
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
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Show all messages."
         Top             =   3840
         Width           =   2295
      End
      Begin VB.CommandButton in_DelAll 
         BackColor       =   &H00EBD8C3&
         Caption         =   "Delete ALL messages"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Delete all messages from your inbox."
         Top             =   4920
         Width           =   2295
      End
      Begin VB.CommandButton in_DelThese 
         BackColor       =   &H00EBD8C3&
         Caption         =   "Delete these messages"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Delete the 15 newest messages in your inbox (the first page on nordicmafia)."
         Top             =   4680
         Width           =   2295
      End
      Begin VB.ListBox in_List 
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         ItemData        =   "Postal_frmMain.frx":0004
         Left            =   120
         List            =   "Postal_frmMain.frx":0006
         TabIndex        =   14
         ToolTipText     =   "Your message inbox. Click a message to show it. Click a message, then delete, to delete it both from nordicmafia and Postal."
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label in_Stats 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 total, 0 unread"
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
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lbInbox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ".:: Inbox ::."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00D6C2AC&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   0
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Frame frmWelcome 
      BackColor       =   &H007E543C&
      BorderStyle     =   0  'None
      Caption         =   "Nordicmafia Postal System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2880
      TabIndex        =   27
      Top             =   1440
      Width           =   7335
      Begin VB.Frame frmContList 
         BackColor       =   &H00D6C2AC&
         Caption         =   "Contacts list"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   3360
         TabIndex        =   62
         Top             =   480
         Width           =   3615
         Begin VB.ListBox cntList 
            BackColor       =   &H00FFEFE0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3210
            ItemData        =   "Postal_frmMain.frx":0008
            Left            =   120
            List            =   "Postal_frmMain.frx":000A
            TabIndex        =   69
            ToolTipText     =   "The list of pending messages. It will send messages from top down, unless there are prioritated messages waiting."
            Top             =   240
            Width           =   3375
         End
         Begin VB.CheckBox cntSendCurr 
            Caption         =   "Send curr. PM"
            Height          =   200
            Left            =   240
            TabIndex        =   67
            Top             =   3975
            Width           =   200
         End
         Begin VB.CommandButton cntRefresh 
            BackColor       =   &H00EBD8C3&
            Caption         =   "Refresh list"
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
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Make Postal remember your username and password. Required for ""Check for PMs at launch"" to work."
            Top             =   3960
            Width           =   1455
         End
         Begin VB.CommandButton cntSendPM 
            BackColor       =   &H00EBD8C3&
            Caption         =   "Send PM"
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
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Make Postal remember your username and password. Required for ""Check for PMs at launch"" to work."
            Top             =   3600
            Width           =   975
         End
         Begin VB.CommandButton cntRem 
            BackColor       =   &H00EBD8C3&
            Caption         =   "Remove"
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
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Make Postal remember your username and password. Required for ""Check for PMs at launch"" to work."
            Top             =   3600
            Width           =   975
         End
         Begin VB.CommandButton cntAdd 
            BackColor       =   &H00EBD8C3&
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Make Postal remember your username and password. Required for ""Check for PMs at launch"" to work."
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label lbSendCurr 
            BackColor       =   &H00FFEFE0&
            BackStyle       =   0  'Transparent
            Caption         =   "Send curr. PM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   525
            TabIndex        =   68
            ToolTipText     =   "Click a contact, and the last written PM will be sent to him/her also."
            Top             =   3960
            Width           =   1335
         End
      End
      Begin VB.Frame frmGenConf 
         BackColor       =   &H00D6C2AC&
         Caption         =   "General configuration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   360
         TabIndex        =   48
         Top             =   2160
         Width           =   2655
         Begin VB.ComboBox cmbLanguage 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "Postal_frmMain.frx":000C
            Left            =   120
            List            =   "Postal_frmMain.frx":0016
            TabIndex        =   76
            Text            =   "Language"
            Top             =   1800
            Width           =   2415
         End
         Begin VB.CheckBox cfgCntLaunch 
            Caption         =   "Check1"
            Height          =   200
            Left            =   255
            TabIndex        =   74
            ToolTipText     =   "Automatically fetch your contacts' online statuses when you launch Postal."
            Top             =   1460
            Width           =   200
         End
         Begin VB.CheckBox cfgCntWarn 
            Caption         =   "Check1"
            Height          =   200
            Left            =   260
            TabIndex        =   73
            ToolTipText     =   "Warns you if a contact logs on or off."
            Top             =   1220
            Width           =   200
         End
         Begin VB.TextBox cfgCntRefresh 
            Alignment       =   2  'Center
            BackColor       =   &H00FFEFE0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   70
            Text            =   "30"
            ToolTipText     =   "Update the contact list every x second. Put 0 to disable the feature."
            Top             =   960
            Width           =   495
         End
         Begin VB.CheckBox cfgAutocheck 
            Caption         =   "Check1"
            Height          =   200
            Left            =   255
            TabIndex        =   54
            ToolTipText     =   "Automatically fetch all PMs when you start Postal. Recommended."
            Top             =   750
            Width           =   200
         End
         Begin VB.TextBox cfgMaxRead 
            Alignment       =   2  'Center
            BackColor       =   &H00FFEFE0&
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
            Left            =   120
            TabIndex        =   51
            Text            =   "0"
            ToolTipText     =   "The max amount of messages that Postal should read. Put 0 to read all 15."
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox cfgAutosync 
            Alignment       =   2  'Center
            BackColor       =   &H00FFEFE0&
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
            Left            =   120
            TabIndex        =   50
            Text            =   "30"
            ToolTipText     =   "Check for new PMs every x second. Put 0 to disable the feature."
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cfgSave 
            BackColor       =   &H00EBD8C3&
            Caption         =   "Save configuration"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Make Postal remember this configuration."
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Label lbCntLaunch 
            BackStyle       =   0  'Transparent
            Caption         =   "Get contacts at launch"
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
            Left            =   720
            TabIndex        =   75
            ToolTipText     =   "Automatically fetch your contacts' online statuses when you launch Postal."
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lbCntWarn 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact state warner"
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
            Left            =   720
            TabIndex        =   72
            ToolTipText     =   "Warns you if a contact logs on or off."
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lbCntRefresh 
            BackStyle       =   0  'Transparent
            Caption         =   "Online list refresher"
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
            Left            =   720
            TabIndex        =   71
            ToolTipText     =   "Update the contact list every x second. Put 0 to disable the feature."
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lbAutocheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Check for PMs at launch"
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
            Left            =   720
            TabIndex        =   55
            ToolTipText     =   "Automatically fetch all PMs when you start Postal. Recommended."
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lbMaxRead 
            BackStyle       =   0  'Transparent
            Caption         =   "Max messages to read"
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
            Left            =   720
            TabIndex        =   53
            ToolTipText     =   "The max amount of messages that Postal should read. Put 0 to read all 15."
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lbAutosync 
            BackStyle       =   0  'Transparent
            Caption         =   "Automatic PM checking"
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
            Left            =   720
            TabIndex        =   52
            ToolTipText     =   "Check for new PMs every x second. Put 0 to disable the feature."
            Top             =   480
            Width           =   1815
         End
      End
      Begin SHDocVwCtl.WebBrowser wb 
         Height          =   4215
         Left            =   8000
         TabIndex        =   36
         Top             =   600
         Width           =   3615
         ExtentX         =   6376
         ExtentY         =   7435
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
         Location        =   ""
      End
      Begin VB.Frame frmAccessInf 
         BackColor       =   &H00D6C2AC&
         Caption         =   "Nordicmafia Access Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         TabIndex        =   33
         Top             =   480
         Width           =   2655
         Begin VB.CommandButton accSave 
            BackColor       =   &H00EBD8C3&
            Caption         =   "Save this information"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Make Postal remember your username and password. Required for ""Check for PMs at launch"" to work."
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox accPass 
            BackColor       =   &H00FFEFE0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   960
            PasswordChar    =   "*"
            TabIndex        =   11
            ToolTipText     =   "Your nordicmafia password."
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox accUser 
            BackColor       =   &H00FFEFE0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   10
            ToolTipText     =   "Your nordicmafia username."
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lbPassword 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   120
            TabIndex        =   35
            Top             =   640
            Width           =   855
         End
         Begin VB.Label lbUsername 
            BackStyle       =   0  'Transparent
            Caption         =   "Username"
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
            Left            =   120
            TabIndex        =   34
            Top             =   280
            Width           =   855
         End
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nordicmafia Postal System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   7335
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00D6C2AC&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   0
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.Frame frmComp 
      BackColor       =   &H007E543C&
      BorderStyle     =   0  'None
      Caption         =   "Compose a message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2880
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton Comp_Send 
         BackColor       =   &H00FFEFE0&
         Caption         =   "Send ASAP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Send the message with priority level one."
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton Comp_Send 
         BackColor       =   &H00FFEFE0&
         Caption         =   "Add to queue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Send the message with priority level two."
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox Comp_Body 
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "Postal_frmMain.frx":0024
         ToolTipText     =   "The message to send."
         Top             =   600
         Width           =   7095
      End
      Begin VB.TextBox Comp_Topic 
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   1
         Text            =   "Topic"
         ToolTipText     =   "The topic of the message."
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox Comp_Name 
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Text            =   "Target user"
         ToolTipText     =   "The user you want to send the message to. Try this once: After writing a message, click this field and hit enter."
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton Comp_Close 
         BackColor       =   &H00FFEFE0&
         Caption         =   "X"
         Height          =   255
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Close the message composer."
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lbMsgComp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Message composer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   7335
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00FFBFA0&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   0
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.Frame frmView 
      BackColor       =   &H007E543C&
      BorderStyle     =   0  'None
      Caption         =   "Message viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2880
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox View_User 
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "The name of the person of which sent you the message."
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox View_When 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   44
         ToolTipText     =   "When the message was sent."
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox View_Topic 
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "The topic of the received message."
         Top             =   600
         Width           =   7095
      End
      Begin VB.CommandButton View_Delete 
         BackColor       =   &H00FFEFE0&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   $"Postal_frmMain.frx":002E
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton View_Reply 
         BackColor       =   &H00FFEFE0&
         Caption         =   "Reply"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click here to open the message composer ready for replying."
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox View_Message 
         BackColor       =   &H00FFEFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "The body of the message."
         Top             =   960
         Width           =   7095
      End
      Begin VB.CommandButton View_Save 
         BackColor       =   &H00FFEFE0&
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
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Save this message to the application's folder."
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton View_Close 
         BackColor       =   &H00FFEFE0&
         Caption         =   "X"
         Height          =   255
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Close the message viewer."
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lbMsgView 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Message viewer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   7335
      End
      Begin VB.Label View_Unread 
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown message"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         ToolTipText     =   "Whether it's a read or unread message."
         Top             =   4860
         Width           =   3135
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00FFBFA0&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   0
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   8400
      Picture         =   "Postal_frmMain.frx":00BF
      ToolTipText     =   "My logo and website, yo."
      Top             =   120
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   5295
      Left            =   10440
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "Postal_frmMain.frx":27E4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const FSx = 13200
Const FSy = 8430
Private sWSit As Boolean, LastReadMSG As Long

Private Sub cmbLanguage_Click()
    modLang.Lang_Set cmbLanguage.Text
End Sub
Sub Form_Load()
    L 1, "Welcome to Nordicmafia Postal System - by Praetox Technologies."
    Me.Caption = Me.Caption & App.Major & "." & App.Minor & "." & App.Revision
    'Me.Move 0, 0
End Sub
Sub Form_Resize()
    If Me.WindowState = 0 Then Me.Width = FSx: Me.Height = FSy
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ExitAPP
End Sub
Sub accUser_Change()
    If accUser <> "" Then User = accUser
End Sub
Sub accPass_Change()
    If accPass <> "" Then Pass = accPass
End Sub
Sub accSave_Click()
    SaveSetting "Praetox_NPS", "Nordicmafia Profile", "Username", accUser
    SaveSetting "Praetox_NPS", "Nordicmafia Profile", "Password", accPass
    'INI False, "USR", accUser
    'INI False, "PWD", accPass
End Sub
Sub cfgSave_Click()
    SaveSetting "Praetox_NPS", "Configuration", "cfg_MaxRead", cfgMaxRead
    SaveSetting "Praetox_NPS", "Configuration", "cfg_Autosync", cfgAutosync
    SaveSetting "Praetox_NPS", "Configuration", "cfg_Autocheck", cfgAutocheck
    SaveSetting "Praetox_NPS", "Configuration", "cfg_CntRefresh", cfgCntRefresh
    SaveSetting "Praetox_NPS", "Configuration", "cfg_CntWarn", cfgCntWarn
    SaveSetting "Praetox_NPS", "Configuration", "cfg_CntLaunch", cfgCntLaunch
    SaveSetting "Praetox_NPS", "Configuration", "cfg_Language", modLang.Lng.val
End Sub

Sub in_List_Click()
    If COMPILED Then On Error GoTo hell
    Dim CurMSG As Mail
    vl = in_List.List(in_List.ListIndex)
    LastReadMSG = in_List.ItemData(in_List.ListIndex)
    If Left(vl, 2) = "  " Then vl = Right(vl, Len(vl) - 2)
    in_List.List(in_List.ListIndex) = vl
    CurMSG = GetMsgAry(LastReadMSG)
    View_User = CurMSG.Name
    View_Topic = CurMSG.Topic
    View_Message = CurMSG.Body
    View_When = CurMSG.rDate & " :: " & CurMSG.rTime
    If CurMSG.Unread Then
        View_Unread = "New message"
        View_Unread.ForeColor = &H9000&
        CurMSG.Unread = False
    Else
        View_Unread = "Old message"
        View_Unread.ForeColor = &H90&
    End If
    SetMsgAry in_List.ItemData(in_List.ListIndex), CurMSG
    fShow frmView
    L 2, "Showing message " & CurMSG.ID & "."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while loading message from listclick!"
End Sub
Private Sub in_List_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then DeleteMessage
End Sub

Sub tAutosync_Timer()
    If COMPILED Then On Error GoTo hell
    If IsNumeric(cfgAutosync) = False Or cfgAutosync = 0 Then GoTo 10
    Dim AS_Delay As Long: AS_Delay = Int(cfgAutosync)
    If Timer - ASDelay > AS_Delay Then
        If GetForegroundWindow <> Me.hWnd Then
            cmdRefresh_Click
            If in_List.ListCount > 0 Then
                If Left(in_List.List(0), 2) = "  " Then
                    FlashWindow Me.hWnd, 0
                    wav "new"
                End If
            End If
            ASDelay = Timer
        End If
    End If
10  If IsNumeric(cfgCntRefresh) = False Or cfgCntRefresh = 0 Then Exit Sub
    Dim OR_Delay As Long: OR_Delay = Int(cfgCntRefresh)
    If Timer - ORDelay > OR_Delay Then
        If GetForegroundWindow <> Me.hWnd Then
            ContactList_Refresh
        End If
        ORDelay = Timer
    End If
    Exit Sub
hell: wav "ohno": L 1, "Error occured while autosyncing!"
End Sub

Sub tSender_Timer()
    If COMPILED Then On Error GoTo hell
    tSender.Enabled = False
    Dim iASAP As Integer, sASAP As String, NeedRefresh As Boolean
    out_Stats = Replace(Replace(Lng.soutStats, "%1", out_List.ListCount), "%2", gTil(SendDelay))
    For a = 0 To out_List.ListCount - 1
        tmpList = tmpList & out_List.List(a) & vbCrLf
    Next
    For a = 0 To UBound(Outbox)
        If Outbox(a).ID <> 0 Then
            If Outbox(a).ID = 2 Then
                iASAP = 1
                sASAP = "  "
            Else
                sASAP = ""
            End If
            tmpAry = tmpAry & sASAP & Outbox(a).Name & "> " & Outbox(a).Topic & vbCrLf
        End If
    Next
    If tmpList <> tmpAry Then
        out_List.Clear
        For a = 0 To UBound(Outbox)
            If Outbox(a).ID <> 0 Then
                If Outbox(a).ID = 2 Then
                    iASAP = 1
                    sASAP = "  "
                Else
                    sASAP = ""
                End If
                out_List.AddItem sASAP & Outbox(a).Name & "> " & Outbox(a).Topic
                out_List.ItemData(out_List.ListCount - 1) = Outbox(a).ID
            End If
        Next
    End If
    If gTil(SendDelay) > 0 Then GoTo 10
    If out_List.ListCount = 0 Then GoTo 10
    For a = 0 To UBound(Outbox)
        If Outbox(a).ID > iASAP Then
            If IsNumeric(Me.Caption) = False Then Me.Caption = 0
            Me.Caption = Me.Caption + 1
            SendMessage (a)
            out_List.AddItem "."
            GoTo 10
        End If
    Next
10  tSender.Enabled = True
    Exit Sub
hell: wav "ohno": L 1, "Error occured while scanning outbox!"
    tSender.Enabled = True
End Sub
Sub SendMessage(ByVal OutboxPos As Integer)
    If COMPILED Then On Error GoTo hell
    Dim CurMSG As Mail: CurMSG = Outbox(OutboxPos)
    L 2, "Sending """ & CurMSG.Topic & """ to " & CurMSG.Name
    Nav "http://www.nordicmafia.net/nordic/index.php?side=drep"
    WSit "/nordic/index.php?side=pm_ny2", _
         "http://www.nordicmafia.net/nordic/index.php?side=pm_ny", _
         "www.nordicmafia.net", _
         "til=" & CurMSG.Name & "&tittel=" & CurMSG.Topic & "&melding=" & CurMSG.Body & "&Submit5222=Send"
    Outbox(OutboxPos).ID = 0
    SendDelay = gEnd(20)
    L 2, "Message """ & CurMSG.Topic & """ to " & CurMSG.Name & " was sent."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while sending message!"
End Sub
Sub WSit(ByVal PostPath As String, Referer As String, Host As String, Content As String)
    If COMPILED Then On Error GoTo hell
    tmp = "POST " & PostPath & " " & _
          "HTTP/1.1" & vbCrLf & _
          "Accept: */*" & vbCrLf & _
          "Referer: " & Referer & vbCrLf & _
          "Accept-Language: en-gb" & vbCrLf & _
          "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
          "Accept-Encoding: gzip, deflate" & vbCrLf & _
          "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.0.3705; .NET CLR 1.1.4322; Media Center PC 4.0; .NET CLR 2.0.50727)" & vbCrLf & _
          "Host: " & Host & vbCrLf & _
          "Content-Length: " & Len(Content) & vbCrLf & _
          "Connection: Keep-Alive" & vbCrLf & _
          "Cache-Control: no-cache" & vbCrLf & _
          "Cookie: PHPSESSID=" & Session & vbCrLf & _
          "" & vbCrLf & Content
    sWSit = True
    ws.Connect Host, "80"
    While sWSit
        DoEvents
    Wend
    ws.SendData tmp
    DoEvents
    ws.Close
    DoEvents
    Exit Sub
hell: wav "ohno": L 1, "Error occured while putting on shoes."
End Sub
Sub ws_Connect()
    sWSit = False
End Sub

Sub View_Close_Click()
    fHide frmView
    L 2, "Closed viewer."
End Sub
Sub Comp_Close_Click()
    fHide frmComp
    L 2, "Closed composer."
End Sub
Sub cmdRefresh_Click()
    L 1, "Refreshing inbox..."
    RefreshMemoryFromSite
    ReloadFromMemory
    L 1, "Refreshing completed."
End Sub
Sub cmdFilter_Click()
    If COMPILED Then On Error GoTo hell
    Filter = InputBox( _
        "0: Show all mails" & vbCrLf & _
        "1: Show only unread mails" & vbCrLf & _
        "2: Show only read mails", _
        "Choose filter", "1")
    ReloadFromMemory
    L 2, "Filter applied."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while setting filter. How the heck did you manage that?!"
End Sub
Sub in_Filter_Click(index As Integer)
    Filter = index
    ReloadFromMemory
    L 2, "Filter applied."
End Sub
Sub in_DelThese_Click()
    If COMPILED Then On Error GoTo hell
    L 1, "Preparng to delete messages currently shown in inbox..."
    If MsgBox(Lng.sMsgDelThese, vbYesNo) = vbNo Then GoTo 10
    L 2, "Resuming deletion."
    For a = 0 To in_List.ListCount - 1
        MsgID = in_List.ItemData(a)
        L 2, "Deleting message " & MsgID & "..."
        Nav "http://www.nordicmafia.net/nordic/index.php?side=pm_slett&id=" & MsgID
    Next
    L 1, "Deletion completed."
    RefreshMemoryFromSite
    ReloadFromMemory
    Exit Sub
10  L 2, "Deletion cancelled."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while deleting first page of inbox!"
End Sub
Private Sub cmdDelBy_Click()
    'If COMPILED Then On Error GoTo hell
    Dim DelBy1 As String, DelBy2 As String, DelThis As Boolean, HasDeleted As Boolean, cDel As Long
    DelBy1 = InputBox("What do you want to delete by?" & vbCrLf & vbCrLf & "1: Sender" & vbCrLf & "2: Topic", , "2")
    DelBy2 = InputBox("Enter the name/topic you want to annihilate.", , "Fight Club Info")
    Do
        L 1, "Loading PM list..."
        Nav "http://www.nordicmafia.net/nordic/index.php?side=pm_inn"
        aMails = Split(Split(wSRC, "size=1>Emne:</FONT>")(1), "lest.jpg""> = Lest pm")(0)
        aMails = Split(aMails, "<TD width=28 bgColor=#000000>")
        HasDeleted = False
        For a = 1 To UBound(aMails)
            ThisID = Split(Split(aMails(a), "side=pm_les&amp;id=")(1), """>")(0)
            ThisName = Split(Split(Split(aMails(a), "side=bruker&amp;brukernavn=")(1), """>")(1), "</")(0)
            ThisTopic = Split(Split(Split(aMails(a), "side=pm_les&amp;id=")(2), """>")(1), "</")(0)
            'MsgBox ">" & ThisID & "<" & vbCrLf & ">" & ThisName & "<" & vbCrLf & ">" & ThisTopic & "<"
            DelThis = False
            If DelBy1 = "1" And LCase(ThisName) = LCase(DelBy2) Then DelThis = True
            If DelBy1 = "2" And LCase(ThisTopic) = LCase(DelBy2) Then DelThis = True
            If DelThis Then
                Nav "http://www.nordicmafia.net/nordic/index.php?side=pm_slett&id=" & ThisID
                HasDeleted = True: cDel = cDel + 1
                L 2, "Deleting " & ThisName & " :: " & ThisTopic & " (" & ThisID & ")..."
            End If
        Next
        If HasDeleted = False Then Exit Do
    Loop
    L 1, "Deleted " & cDel & " messages."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while deleting specified name/topic!"
End Sub
Private Sub cmdDeleteRead_Click()
    If COMPILED Then On Error GoTo hell
    Dim IDs(14) As Long, ReadMessages As Boolean, iCnt As Integer
    L 1, "Deleting all read messages. Please hold."
    If MsgBox(Lng.sMsgDelRead, vbYesNo) = vbNo Then GoTo 10
    ReadMessages = True
    While ReadMessages = True
        Nav "http://www.nordicmafia.net/nordic/index.php?side=pm_inn"
        aMails = Split(Split(wSRC, "size=1>Emne:</FONT>")(1), "lest.jpg""> = Lest pm")(0)
        If InStr(1, aMails, "iconer/pm_lest.jpg") = 0 Then ReadMessages = False
        aMails = Split(aMails, "<TD width=28 bgColor=#000000>")
        etSession = 0
        For a = 1 To UBound(aMails)
            IDs(a - 1) = Split(Split(aMails(a), "side=pm_les&amp;id=")(1), """>")(0)
            For b = 0 To UBound(Inbox)
                If Inbox(b).ID = IDs(a - 1) Then
                    If Inbox(b).Unread Then aMails(a) = aMails(a) & "iconer/pm_konf.jpg"
                End If
            Next
            If InStr(1, aMails(a), "iconer/pm_konf.jpg") = 0 Then
                L 2, "Deleting message " & IDs(a - 1) & "..."
                iCnt = iCnt + 1: etSession = etSession + 1
                Nav "http://www.nordicmafia.net/nordic/index.php?side=pm_slett&id=" & IDs(a - 1)
            End If
        Next
        If etSession = 0 Then ReadMessages = False
    Wend
    L 1, "Message deleting completed. Removed " & iCnt & " messages."
    RefreshMemoryFromSite
    ReloadFromMemory
    Exit Sub
10  L 2, "Deletion cancelled."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while deleting read messages!"
End Sub
Sub in_DelAll_Click()
    cmdEmpty_Click
End Sub
Sub cmdEmpty_Click()
    If COMPILED Then On Error GoTo hell
    L 1, "About to completely wipe your inbox..."
    If MsgBox(Lng.sMsgDelAll, vbYesNo) = vbNo Then GoTo 10
    L 1, "Deleting all messages!"
    Nav "about:<FORM name=pminn action=http://www.nordicmafia.net/nordic/index.php?side=pm_inn method=post><INPUT type=submit value='clr' name=sub_delall></form>"
    wb.Document.All("sub_delall").Click: w8
    For a = 0 To UBound(Inbox)
        Inbox(a).ID = 0
    Next
    ReloadFromMemory
    L 1, "Inbox completely and utterly annihilated."
    Exit Sub
10  L 2, "Inbox wipeage cancelled."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while wiping inbox!"
End Sub
Sub cmdNew_Click()
    If COMPILED Then On Error GoTo hell
    Comp_Name = ""
    Comp_Topic = ""
    Comp_Body = ""
    fShow frmComp
    L 2, "Composing a message..."
    Comp_Name.SetFocus
    Exit Sub
hell: wav "ohno": L 1, "Error occured while loading composer for reply!"
End Sub
Sub View_Reply_Click()
    If COMPILED Then On Error GoTo hell
    Dim CurMSG As Mail
    CurMSG = GetMsgAry(LastReadMSG)
    Comp_Name = CurMSG.Name
    Comp_Topic = "R " & CurMSG.Topic
    Comp_Body = vbCrLf & vbCrLf & "Forrige melding, sendt " & CurMSG.rDate & ", klokken " & CurMSG.rTime & ":" & vbCrLf & CurMSG.Body
    fShow frmComp
    L 2, "Replying to message " & CurMSG.ID & "..."
    Comp_Body.SetFocus
    Exit Sub
hell: wav "ohno": L 1, "Error occured while loading composer for reply!"
End Sub
Sub View_Save_Click()
    If COMPILED Then On Error GoTo hell
    Dim CurMSG As Mail
    CurMSG = GetMsgAry(LastReadMSG)
    Fn = FreeFile
    Open Path & "_" & CurMSG.ID & ".txt" For Output As #Fn
    Print #Fn, "Recieved: " & CurMSG.rDate & " :: " & CurMSG.rTime
    Print #Fn, "Name: " & CurMSG.Name
    Print #Fn, "Topic: " & CurMSG.Topic
    Print #Fn, vbCrLf & CurMSG.Body
    Close #Fn
    L 2, "Stored message " & CurMSG.ID & "."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while saving message!"
End Sub
Sub View_Delete_Click()
    DeleteMessage
End Sub
Private Sub Comp_Topic_Change()
    If Len(Comp_Topic) > 30 Then
        Comp_Topic = Left(Comp_Topic, 30)
        Comp_Topic.SelStart = Len(Comp_Topic)
        Comp_Topic.SelLength = 0
    End If
End Sub
Sub Comp_Send_Click(index As Integer)
    If COMPILED Then On Error GoTo hell
    If Comp_Name = "" Then MsgBox "No target entered!", , "Oops...?": Comp_Name.SetFocus: Exit Sub
    If Comp_Topic = "" Then MsgBox "No topic entered!", , "Oops...?": Comp_Topic.SetFocus: Exit Sub
    If Comp_Body = "" Then MsgBox "No message entered!", , "Oops...!?": Comp_Body.SetFocus: Exit Sub
    Dim FoundSlot As Boolean
    For a = 0 To UBound(Outbox)
        If Outbox(a).ID = 0 Then
            Outbox(a).ID = index + 1
            Outbox(a).Name = Comp_Name
            Outbox(a).Topic = Comp_Topic
            Outbox(a).Body = Comp_Body
            FoundSlot = True
            Exit For
        End If
    Next
    If FoundSlot = False Then
        MsgBox "Sorry, but the message could not be queued." & vbCrLf & vbCrLf & _
               "Reason:" & vbCrLf & "Limit for queued messages is " & UBound(Outbox) + 1 & "."
    End If
    Exit Sub
hell: wav "ohno": L 1, "Error occured while pre-processing message sending!"
End Sub
Sub Comp_Name_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Comp_Send_Click (0)
        Comp_Name.Text = ""
    End If
End Sub

Sub fHide(ByVal frm As Frame, Optional Conceal As Boolean = False)
    If Conceal Then
        frm.ZOrder (1)
    Else
        frm.Visible = False
        'For Each Frame In frmMain
        '    If Frame.Visible = True Then vis = vis + 1
        'Next
        'If vis = 0 Then frmWelcome.Visible = True
    End If
End Sub
Sub fShow(ByVal frm As Frame)
    frm.Visible = True
    frm.ZOrder (0)
End Sub

Sub ReloadFromMemory()
    If COMPILED Then On Error GoTo hell
    Dim Rsm As Boolean, cTotal As Long, cUnread As Long, sNew As String
    in_List.Clear
    L 1, "Refreshing application displays..."
    For a = 0 To UBound(Inbox)
        If Inbox(a).ID <> 0 Then
            If Filter = 0 Then Rsm = True
            If Filter = 1 And Inbox(a).Unread = True Then Rsm = True
            If Filter = 2 And Inbox(a).Unread = False Then Rsm = True
            If Rsm = True Then
                Rsm = False
                If Inbox(a).Unread Then sNew = "  " Else sNew = ""
                in_List.AddItem sNew & Inbox(a).Name & "> " & Inbox(a).Topic
                in_List.ItemData(in_List.ListCount - 1) = Inbox(a).ID
            End If
            If Inbox(a).Unread Then cUnread = cUnread + 1
            cTotal = cTotal + 1
        End If
    Next
    in_Stats = Replace(Replace(Lng.sinStats, "%1", cTotal), "%2", cUnread)
    L 1, "Application inbox is now up to date."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while redrawing displays!"
End Sub
Sub RefreshMemoryFromSite()
    If COMPILED Then On Error GoTo hell
    Dim Unread As Boolean, lMail As Long, lMail2 As Long, NeedsReload() As Boolean, OldIDs() As Long, _
        newBodies() As String
    L 1, "Caching message information..."
    ReDim OldIDs(UBound(Inbox))
    ReDim NeedsReload(UBound(Inbox))
    ReDim newBodies(UBound(Inbox))
    Nav "http://www.nordicmafia.net/nordic/index.php?side=pm_inn"
    aMails = Split(Split(wSRC, "size=1>Emne:</FONT>")(1), "lest.jpg""> = Lest pm")(0)
    HasMailList = vbCrLf
    For a = 0 To UBound(Inbox)
        If Inbox(a).ID <> 0 Then
            If InStr(1, aMails, "id=" & Inbox(a).ID) = False Then Inbox(a).ID = 0
            If InStr(1, HasMailList, vbCrLf & Inbox(a).ID & vbCrLf) > 0 Then Inbox(a).ID = 0
            HasMailList = HasMailList & Inbox(a).ID & vbCrLf
        End If
    Next
    aMails = Split(aMails, "<TD width=28 bgColor=#000000>")
    If UBound(aMails) = 0 Then Exit Sub
    lMail = -1
    For a = 0 To UBound(Inbox)
        OldIDs(a) = Inbox(a).ID
        NeedsReload(a) = True
    Next
    For a = 1 To UBound(aMails)
        lMail = lMail + 1: If Int(cfgMaxRead) <> 0 Then If lMail > Int(cfgMaxRead) - 1 Then GoTo 10
        L 2, "Caching message " & lMail + 1 & "..."
        TmpID = Split(Split(aMails(a), "side=pm_les&amp;id=")(1), """>")(0)
        Inbox(lMail).ID = TmpID
        Inbox(lMail).Name = Split(Split(Split(aMails(a), "side=bruker&amp;brukernavn=")(1), """>")(1), "</A>")(0)
        Inbox(lMail).Topic = Split(Split(Split(aMails(a), "side=pm_les&amp;id=")(2), """>")(1), "</A>")(0)
        tmp = Split(Split(aMails(a), " || ")(0), ">")
        Inbox(lMail).rDate = tmp(UBound(tmp))
        Inbox(lMail).rTime = Split(Split(aMails(a), " || ")(1), "<")(0)
        If InStr(1, aMails(a), "iconer/pm_konf.jpg") > 0 Then Unread = True Else Unread = False
        Inbox(lMail).Unread = Unread
    Next
    For a = 0 To UBound(Inbox)
        'If Inbox(a).ID <> 0 Then
            If Inbox(a).ID <> OldIDs(a) Then
                NeedsReload(a) = True
                For b = 0 To UBound(Inbox)
                    If Inbox(a).ID = OldIDs(b) Then
                        newBodies(a) = Inbox(b).Body
                        OldIDs(b) = 0
                        If Len(newBodies(a)) > 0 Then NeedsReload(a) = False
                    End If
                Next
            Else
                newBodies(a) = Inbox(a).Body
                NeedsReload(a) = False
            End If
        'Else
        '    NeedsReload(a) = True
        'End If
    Next
    For a = 0 To UBound(Inbox)
        Inbox(a).Body = newBodies(a)
    Next
    
10  L 1, "Reading message contents..."
    If FEx(Path & "tmpInbox.dat") Then Kill (Path & "tmpInbox.dat"): DoEvents
    For a = 0 To UBound(Inbox)
        If NeedsReload(a) = True And Inbox(a).ID <> 0 Then
            lMail2 = lMail2 + 1
            L 2, "Reading message body " & lMail2 & " of " & lMail + 1 & "..."
            Nav "http://www.nordicmafia.net/nordic/index.php?side=pm_les&id=" & Inbox(a).ID
            
            DateToday = Split(Split(wSRC, "<SPAN class=nicktext>")(1), "</SPAN>")(0)
            tmpDateToday = Split(DateToday, " ")
            DateToday = tmpDateToday(1) & " " & _
                        Month_NorToEng(tmpDateToday(2)) & " " & _
                        tmpDateToday(3)
            
            Inbox(a).Body = Split(Split(Split(wSRC, "bgColor=#ebebeb")(1), "size=1>")(1), "</FONT>")(0)
            Inbox(a).Body = Replace(Inbox(a).Body, "<BR>", vbCrLf)
            Inbox(a).Body = Replace(Inbox(a).Body, "<STRONG>", "")
            Inbox(a).Body = Replace(Inbox(a).Body, "</STRONG>", "")
            Inbox(a).Body = Replace(Inbox(a).Body, "<EM>", "")
            Inbox(a).Body = Replace(Inbox(a).Body, "</EM>", "")
            Inbox(a).Body = Replace(Inbox(a).Body, "<U>", "")
            Inbox(a).Body = Replace(Inbox(a).Body, "</U>", "")
            While Right(Inbox(a).Body, 2) = vbCrLf
                Inbox(a).Body = Left(Inbox(a).Body, Len(Inbox(a).Body) - 2)
            Wend
            While Left(Inbox(a).Body, 2) = vbCrLf
                Inbox(a).Body = Right(Inbox(a).Body, Len(Inbox(a).Body) - 2)
            Wend
            If DateToday = Inbox(a).rDate Then
                Fn = FreeFile
                Open Path & "tmpInbox.dat" For Append As #Fn
                    Print #Fn, Inbox(a).rTime & " :: " & Inbox(a).Name & " :: " & Inbox(a).Topic '& " :: " & Inbox(a).Body
                Close #Fn
            End If
        End If
    Next
    If FEx(Path & "tmpInbox.dat") Then
        If FEx(Path & "Inbox-" & InvDate & ".txt") Then
            oldInbox = vbCrLf
            Fn = FreeFile
            Open Path & "Inbox-" & InvDate & ".txt" For Input As #Fn
            While Not EOF(Fn)
                Line Input #Fn, ttInbox
                oldInbox = oldInbox & ttInbox & vbCrLf
            Wend
            Close #Fn
        End If
        Fn = FreeFile
        Open Path & "tmpInbox.dat" For Input As #Fn
        While Not EOF(Fn)
            Line Input #Fn, ttInbox
            tmpInbox = ttInbox & vbCrLf & tmpInbox
        Wend
        Close #Fn
        If Right(tmpInbox, 2) = vbCrLf Then tmpInbox = Left(tmpInbox, Len(tmpInbox) - 2)
        tmpInbox = Split(tmpInbox, vbCrLf)
        Fn = FreeFile
        Open Path & "Inbox-" & InvDate & ".txt" For Append As #Fn
        For a = UBound(tmpInbox) To 0 Step -1
            If InStr(1, oldInbox, tmpInbox(a)) = 0 Then Print #Fn, tmpInbox(a)
        Next
        Close #Fn
    End If
20  L 2, "Memory refreshed."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while reading messages from nordicmafia!"
End Sub
Function InvDate() As String
    InvDate = Split(Date$, "-")(2) & "-" & Split(Date$, "-")(0) & "-" & Split(Date$, "-")(1)
End Function
Function Month_NorToEng(ByVal vl As String) As String
    vl = Replace(vl, "Januar", "January")
    vl = Replace(vl, "Februar", "February")
    vl = Replace(vl, "Mars", "March")
    'vl = Replace(vl, "April", "April")
    vl = Replace(vl, "Mai", "May")
    vl = Replace(vl, "Juni", "June")
    vl = Replace(vl, "Juli", "July")
    'vl = Replace(vl, "August", "August")
    'vl = Replace(vl, "September", "September")
    vl = Replace(vl, "Oktober", "October")
    'vl = Replace(vl, "November", "November")
    vl = Replace(vl, "Desember", "December")
    Month_NorToEng = vl
End Function
Sub DeleteMessage()
    If COMPILED Then On Error GoTo hell
    If in_List.ListIndex = -1 Then MsgBox "Please select a message to delete.": Exit Sub
    vl = MsgBox( _
        Lng.sMsgDelThisMsg & vbCrLf & vbCrLf & _
        "From: " & Split(in_List.List(in_List.ListIndex), "> ")(0) & vbCrLf & _
        "Topic: " & Split(in_List.List(in_List.ListIndex), "> ")(1), _
        vbYesNo, "Delete?")
    If vl = vbYes Then
        Dim CurMSG As Mail
        CurMSG = GetMsgAry(LastReadMSG)
        L 1, "Deleting message " & CurMSG.ID & " - """ & CurMSG.Topic & """ from " & CurMSG.Name & "..."
        Nav "http://www.nordicmafia.net/nordic/index.php?side=pm_slett&id=" & CurMSG.ID
        L 1, "Deleted message " & CurMSG.ID & " - """ & CurMSG.Topic & """ from " & CurMSG.Name & "!"
        CurMSG.ID = 0
        SetMsgAry in_List.ItemData(in_List.ListIndex), CurMSG
        RefreshMemoryFromSite
        ReloadFromMemory
    End If
    Exit Sub
hell: wav "ohno": L 1, "Error occured while deleting message!"
End Sub
Sub ContactList_Refresh(Optional PlaySounds As Boolean = True)
    If COMPILED Then On Error GoTo hell
    If cntList.ListCount = 0 Then Exit Sub
    Dim IsOnline() As String
    ReDim IsOnline(cntList.ListCount - 1)
    For a = 0 To cntList.ListCount - 1
        IsOnline(a) = cntList.List(a)
    Next
    
    CharList = vbCrLf & GetAllChars & vbCrLf
    For a = 0 To UBound(IsOnline)
        If InStr(1, CharList, vbCrLf & Mid$(IsOnline(a), 3) & vbCrLf, vbTextCompare) = 0 Then
            lstOffline = lstOffline & Mid$(IsOnline(a), 3) & vbCrLf
        Else
            lstOnline = lstOnline & Mid$(IsOnline(a), 3) & vbCrLf
        End If
    Next
    cntList.Clear
    If Len(lstOffline) > 3 Then lstOffline = Left(lstOffline, Len(lstOffline) - 2)
    If Len(lstOnline) > 3 Then lstOnline = Left(lstOnline, Len(lstOnline) - 2)
    lstOffline = Split(lstOffline, vbCrLf)
    lstOnline = Split(lstOnline, vbCrLf)
    For a = 0 To UBound(lstOnline)
        cntList.AddItem "O " & lstOnline(a)
    Next
    For a = 0 To UBound(lstOffline)
        cntList.AddItem "X " & lstOffline(a)
    Next
    
    If PlaySounds = True Then
        For a = 0 To cntList.ListCount - 1
            ThisCharName = cntList.List(a)
            For b = 0 To UBound(IsOnline)
                If Mid$(ThisCharName, 2) = Mid$(IsOnline(a), 2) Then
                    If ThisCharName <> IsOnline(a) Then
                        If Left(ThisCharName, 1) = "X" Then wav "loff", 0 Else wav "lon", 0
                    End If
                    Exit For
                End If
            Next
        Next
    End If
    Exit Sub
hell: wav "ohno": L 1, "Error occured while refreshing contacts list!"
End Sub
Private Sub cntList_Click()
    If COMPILED Then On Error GoTo hell
    If cntSendCurr Then
        Comp_Name = Mid$(cntList.List(cntList.ListIndex), 3)
        Comp_Send_Click (0)
        Comp_Name = ""
    End If
    Exit Sub
hell: wav "ohno": L 1, "Error occured while sending message from contacts list!"
End Sub
Private Sub cntAdd_Click()
    If COMPILED Then On Error GoTo hell
    NewChar = InputBox(Lng.sMsgAddCnt, "Add contact")
    cntList.AddItem "? " & NewChar
    SaveContacts
    Exit Sub
hell: wav "ohno": L 1, "Error occured while adding contact!"
End Sub
Private Sub cntRem_Click()
    If COMPILED Then On Error GoTo hell
    CntToRem = Mid$(cntList.List(cntList.ListIndex), 3)
    vl = MsgBox(Replace(Lng.sMsgRemCnt, "%1", CntToRem), vbYesNo)
    If vl = vbNo Then Exit Sub
    cntList.RemoveItem (cntList.ListIndex)
    SaveContacts
    Exit Sub
hell: wav "ohno": L 1, "Error occured while removing contact!"
End Sub
Private Sub cntRefresh_Click()
    ContactList_Refresh
End Sub
Sub LoadContacts()
    If COMPILED Then On Error GoTo hell
    cntList.Clear
    If FEx("Contacts.txt") = False Then
        L 2, "No contacts to load. Resuming..."
        Exit Sub
    End If
    Fn = FreeFile
    Open "Contacts.txt" For Input As #Fn
    While Not EOF(Fn)
        Line Input #Fn, tmp
        cntList.AddItem "? " & tmp
    Wend
    Close #Fn
    Exit Sub
hell: wav "ohno": L 1, "Error occured while loading contacts list!"
End Sub
Sub SaveContacts()
    If COMPILED Then On Error GoTo hell
    If cntList.ListCount = 0 Then Exit Sub
    Fn = FreeFile
    Open "Contacts.txt" For Output As #Fn
    For a = 0 To cntList.ListCount - 1
        Print #Fn, Mid$(cntList.List(a), 3)
    Next
    Close #Fn
    Exit Sub
hell: wav "ohno": L 1, "Error occured while saving contacts list!"
End Sub
Private Sub cntSendPM_Click()
    If COMPILED Then On Error GoTo hell
    Comp_Name = Mid$(cntList.List(cntList.ListIndex), 3)
    Comp_Topic = ""
    Comp_Body = ""
    fShow frmComp
    L 2, "Composing a message to " & Comp_Name & "..."
    Comp_Topic.SetFocus
    Exit Sub
hell: wav "ohno": L 1, "Error occured while loading composer for reply!"
End Sub




Private Sub cmdAdvertise_Click()
    If COMPILED Then On Error GoTo hell
    L 1, "Confirming advertisement activation."
    vl = MsgBox(Lng.sMsgAdvSure, vbYesNo, "Thank you, but...")
    If vl = vbNo Then
        MsgBox "Thought so. Thanks, anyways.": Exit Sub
    End If
    L 1, "Reading player names..."
    CharList = GetAllChars
    L 1, "Preparing system for broadcasting..."
    L 2, "Removing moderators from possible targets..."
    Nav ROT(Website) & "Postal_PName.txt": PName = Split(Split(wSRC, "<PRE>")(1), "</PRE>")(0)
    Nav ROT(Website) & "Postal_Avo.txt": AvoidList = Split(Split(wSRC, "<PRE>")(1), "</PRE>")(0)
    AvoidList = Split(AvoidList, vbCrLf)
    For a = 0 To UBound(AvoidList)
        CharList = Replace(CharList, vbCrLf & AvoidList(a) & vbCrLf, vbCrLf)
    Next
    'MsgBox CharList: ExitAPP
    CharList = Split(CharList, vbCrLf)
    Nav ROT(Website) & "Postal_BCS.txt": WS_String = Split(Split(wSRC, "<PRE>")(1), "</PRE>")(0)
    bctopic = ROT(Split(WS_String, "(|)" & vbCrLf)(0))
    bcmessage = ROT(Split(WS_String, "(|)" & vbCrLf)(1))
    WS_String = ""
    If UBound(CharList) > 999 Then ToPick = 999 Else ToPick = UBound(CharList)
    L 1, "Picking players and adding messages to outbox."
        
        Comp_Name = PName
        Comp_Topic = ROT(bctopic)
        Comp_Body = ROT(bcmessage)
        Comp_Send_Click 1
    
    CharsPicked = vbCrLf: Randomize Timer
    Do While iPicked < ToPick
        L 2, "Player " & iPicked
        tmp = ((Rnd(1) * (UBound(CharList) - 1)) \ 1) + 1
        If InStr(1, CharsPicked, vbCrLf & CharList(tmp) & vbCrLf) = 0 Then
            For b = 0 To UBound(Outbox)
                If Outbox(b).ID = 0 Then
                    Outbox(b).ID = 1
                    Outbox(b).Name = CharList(tmp)
                    Outbox(b).Topic = ROT(bctopic)
                    Outbox(b).Body = ROT(bcmessage)
                    Exit For
                End If
            Next
            CharsPicked = CharsPicked & CharList(tmp) & vbCrLf
            iPicked = iPicked + 1
        End If
    Loop
    L 1, "Broadcasting initated! Thank you."
    Exit Sub
hell: wav "ohno": L 1, "Error occured while adding people to advertisement list!"
End Sub
Private Function GetAllChars() As String
    If COMPILED Then On Error GoTo hell
    CharsLeft = 1: CurCharStart = 0
    While CharsLeft > 0
        L 2, "Reading player names " & CurCharStart & " to " & CurCharStart + 600 & "..."
        tmp = GetChars(CharsLeft, CurCharStart)
        CurCharStart = CurCharStart + 600
        GetAllChars = GetAllChars & tmp & vbCrLf
    Wend
    If Len(CharList) > 3 Then CharList = Left(CharList, Len(CharList) - 2)
    L 2, "Read all online players."
    Exit Function
hell: wav "ohno": L 1, "Error occured while patching together player list!"
End Function
Private Function GetChars(CharsLeft As Variant, vl As Variant) As String
    If COMPILED Then On Error GoTo hell
    Nav "http://www.nordicmafia.net/nordic/index.php?side=online&start=" & vl
    SRC = wSRC
    vl1 = Split(Split(SRC, "Det er <FONT size=3><B>")(1), " spillere")(0)
    vl2 = Split(Split(SRC, "Viser <B>")(1), " spillere")(0)
    CharsLeft = (vl1 - vl) - 600: If CharsLeft < 0 Then CharsLeft = 0
    Chars = Split(Split(SRC, " spillere</B></P></DIV><BR>")(1), "<TABLE ")(0)
    Chars = Split(Chars, """>")
    For a = 1 To UBound(Chars)
        Chars(a) = Split(Chars(a), "</A>")(0)
        GetChars = GetChars & Chars(a) & vbCrLf
    Next
    If Len(GetChars) > 3 Then GetChars = Left(GetChars, Len(GetChars) - 2)
    Exit Function
hell: wav "ohno": L 1, "Error occured while getting single playerlist page!"
End Function
Private Function ROT(ByVal vl As String) As String
    If COMPILED Then On Error GoTo hell
    For a = 1 To Len(vl)
        tmp = Asc(Mid$(vl, a, 1))
        If tmp >= 65 And tmp <= 90 Then
            If tmp <= 77 Then tmp = tmp + 13 Else tmp = tmp - 13
        ElseIf tmp >= 97 And tmp <= 122 Then
            If tmp <= 109 Then tmp = tmp + 13 Else tmp = tmp - 13
        End If
        ROT = ROT & Chr(tmp)
    Next
    Exit Function
hell: wav "ohno": L 1, "Error occured while rotting!"
End Function
