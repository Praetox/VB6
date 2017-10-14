VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shade of Black hider v2.0"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tHider 
      Interval        =   5
      Left            =   4440
      Top             =   240
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdHideSelf 
      Caption         =   "Hide me"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "1"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   2880
      Width           =   375
   End
   Begin VB.ListBox tskHide 
      Height          =   2595
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.ListBox tskFound 
      Height          =   2595
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label cnt 
      Alignment       =   2  'Center
      Caption         =   "0 tasks found."
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2910
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Tasks to hide"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   30
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tasks found"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub cmdHideSelf_Click()
    Me.Hide
End Sub
Private Sub cmdUpdate_Click()
    Update
End Sub
Private Sub form_load()
    Update
End Sub
Private Sub Update()
    Dim TSK As Long, intLen As Long, strTitle As String, pHvnd As String
    cnt = 0: tskFound.Clear
    TSK = GetWindow(Me.hWnd, 0)
    Do While TSK
        If TSK <> Me.hWnd And IsTask(TSK) Then
            intLen = GetWindowTextLength(TSK) + 1
            strTitle = Space(intLen)
            intLen = GetWindowText(TSK, strTitle, intLen)
            If intLen > 0 Then
                pHvnd = TSK
                Do While Len(pHvnd) < 7
                    pHvnd = "0" & pHvnd
                Loop
                tskFound.AddItem pHvnd & " " & strTitle
                cnt = cnt + 1
            End If
        End If
        TSK = GetWindow(TSK, 2)
    Loop
    cnt = cnt & " found."
End Sub
Public Function IsTask(TSK As Long) As Boolean
    Dim Style As Long
    Const ITStyle = &H10000000 Or &H800000
    Style = GetWindowLong(TSK, (-16))
    If (Style And ITStyle) = ITStyle Then IsTask = True
End Function

Private Sub tHider_Timer()
    If GetAsyncKeyState(16) = -32768 Then sDown = True Else sDown = False
    If GetAsyncKeyState(17) = -32768 Then cDown = True Else cDown = False
    If GetAsyncKeyState(18) = -32768 Then aDown = True Else aDown = False
    If GetAsyncKeyState(145) = -32767 Then
        If sDown And cDown And aDown Then
            Toggle (True)
        Else
            Toggle (False)
        End If
    End If
    If GetAsyncKeyState(123) = -32767 Then If sDown And cDown And aDown Then Me.Show
End Sub

Private Sub tskFound_Click()
    tskHide.AddItem tskFound.Text
End Sub
Private Sub tskHide_click()
    toGet = tskHide.Text
    For a = 0 To tskHide.ListCount - 1
        If toGet = tskHide.List(a) Then tskHide.RemoveItem (a)
    Next
End Sub
Private Sub Toggle(ByVal Show As Boolean)
    For a = 0 To tskHide.ListCount - 1
        tHv = Split(tskHide.List(a), " ")(0)
        If Show Then
            Call SetWindowPos(tHv, 0, 0, 0, 0, 0, &H43)
        Else
            Call SetWindowPos(tHv, 0, 0, 0, 0, 0, &H83)
        End If
    Next
End Sub
Private Sub cmdGo_Click()
    If cmdGo.Caption = 0 Then Toggle (True): cmdGo.Caption = 1 Else Toggle (False): cmdGo.Caption = 0
End Sub
