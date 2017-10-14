VERSION 5.00
Begin VB.Form MapReaderfrm 
   Caption         =   "Map Reader for Tibia 7.92 By OsQu"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Refreshcmd 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   3840
      TabIndex        =   20
      Top             =   480
      Width           =   1095
   End
   Begin VB.ListBox Objectlst 
      Height          =   1230
      Left            =   2760
      TabIndex        =   18
      Top             =   2100
      Width           =   1995
   End
   Begin VB.TextBox ObjectInfotxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   2940
      Width           =   1215
   End
   Begin VB.TextBox ObjectIdtxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox PosZtxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox PosYtxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   1740
      Width           =   1215
   End
   Begin VB.TextBox PosXtxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox TileIdtxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox Counttxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TileNumbertxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Object List:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Object Info:"
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Object Id:"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   8
      Top             =   2700
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Objects:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   7
      Top             =   2430
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Position Z:"
      Height          =   195
      Index           =   5
      Left            =   420
      TabIndex        =   6
      Top             =   2100
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Position Y:"
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   5
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Position X:"
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   4
      Top             =   1500
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tile Id:"
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   1170
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Objects in Tile:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   855
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tile Number:"
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This program shows info about player's tile."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3690
   End
End
Attribute VB_Name = "MapReaderfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Objectlst_Click()
ObjectIdtxt.Text = Mid(Objectlst.List(Objectlst.ListIndex), 5, 20) 'Little bit dummy way to do it :D
ObjectInfotxt.Text = Objectlst.ItemData(Objectlst.ListIndex)
End Sub

Private Sub Refreshcmd_Click()
Dim TileNumber As Long, Count As Long, TileId As Long, X As Long, Y As Long, Z As Long, ObjectId(1 To 9) As Long, ObjectInfo(1 To 9) As Long
Dim PlayerTile As Long
Dim i As Long

If Tibia_Hwnd = 0 Then 'Check that there is tibia window
    MsgBox "Tibia Window Not Found!", vbCritical, "Error!"
End If

PlayerTile = Get_PlayerTile 'Get player tile number
'With tile number take all the data
TileNumber = Get_TileInfo(PlayerTile).TileNum
Count = Get_TileInfo(PlayerTile).MapCount
TileId = Get_TileInfo(PlayerTile).TileId
X = Get_TileInfo(PlayerTile).posX
Y = Get_TileInfo(PlayerTile).posY
Z = Get_TileInfo(PlayerTile).posZ

'Objects
Objectlst.Clear
For i = 1 To Count - 1
    ObjectId(i) = Get_TileInfo(PlayerTile).ObjectId(i)
    ObjectInfo(i) = Get_TileInfo(PlayerTile).ObjectInfo(i)
    
    Objectlst.AddItem i & " - " & ObjectId(i)
    Objectlst.ItemData(i - 1) = ObjectInfo(i)
Next i
TileNumbertxt.Text = TileNumber
Counttxt.Text = Count
TileIdtxt.Text = TileId
PosXtxt.Text = X
PosYtxt.Text = Y
PosZtxt.Text = Z
End Sub
