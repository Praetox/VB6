Attribute VB_Name = "General"
Public Const Compiled = True

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function MessageBeep Lib "user32.dll" (ByVal wType As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public doAlert As Boolean, Music As Boolean

Sub enAlert()
    doAlert = True
    frmMain.Main_Timer.Enabled = False
    frmMain.Main_Timer.Enabled = True
    frmMain.Main_Timer_Timer
End Sub

Sub ExitApp()
    Unload frmWhitelist
    Unload frmStart
    Unload frmFTP
    Unload frm1_General
    Unload frm2_Alerts
    Unload frm3_packets
    Unload frm4_Cavebot
    Unload frm5_Map
    Unload frmMain
    End
End Sub

Sub iSleep(ByVal t As Integer)
    tt = Timer * 1000
    While t + tt > Timer * 1000
        DoEvents
        Sleep (1)
    Wend
End Sub
Function l2b(num As Long, NByte As Long) As Byte
    If Compiled = True Then On Error Resume Next
    Dim byt(0 To 2) As Byte
    CopyMemory byt(0), ByVal VarPtr(num), Len(num)
    l2b = byt(NByte - 1)
End Function
Function hbol(Address As Long) As Byte
    If Compiled = True Then On Error Resume Next
    hbol = CByte(Address \ 256) ' high byte
End Function
Function lbol(Address As Long) As Byte
    If Compiled = True Then On Error Resume Next
    Dim h As Byte
    h = CByte(Address \ 256)
    lbol = CByte(Address - (CLng(h) * 256)) ' low byte
End Function
Function s2ba(ByVal vl As String) As Byte()
    Dim woot() As Byte, tmm As String
    tmp = Split(vl, " ")
    ReDim woot(UBound(tmp))
    For a = 0 To UBound(tmp)
        If Left(tmp(a), 1) = "0" Then tmp(a) = Right(tmp(a), 1)
        woot(a) = CByte("&H" & tmp(a))
    Next
    s2ba = woot
End Function
Function FEx(Filename As String) As Boolean
    On Error Resume Next
    FEx = (GetAttr(Filename) And vbDirectory) = 0
End Function

Sub GotoSafe(ByVal SafeZone As String)
    If Compiled = True Then On Error Resume Next
    Dim tX As Long, tY As Long, tZ As Long
    tmp = Split(SafeZone, ",")
    tX = tmp(0): tY = tmp(1): tZ = tmp(2)
    GotoXYZ tX, tY, tZ
End Sub

Function BL_Player() As Long
    If Compiled = True Then On Error Resume Next
    chrid = mReadLong(CH_ID)
    For a = BL_Start To BL_End Step BL_Dist
        If mReadLong(a + BL_ID) = chrid Then BL_Player = a
    Next
End Function

Sub Smsg(ByVal msg As String, Optional time As Long = 50)
    If Compiled = True Then On Error Resume Next
    mWriteString CH_TSt, msg
    mWriteLong CH_TTi, time
End Sub

Function cWithItem(ID As Long) As Long
    If Compiled = True Then On Error Resume Next
    cWithItem = -1: Dim a As Long
    For a = 1 To 25
        tmp = iPosInCont(ID, a)
        If tmp <> -1 Then
            cWithItem = a
            Exit Function
        End If
    Next
End Function
Function iPosInCont(ID As Long, cBP As Long) As Long
    If Compiled = True Then On Error Resume Next
    iPosInCont = -1
    Cont = CT_Start + (cBP * CTD_Container) - CTD_Container
    For a = 1 To mReadLong(Cont + CTD_ContainerAmount)
        If mReadLong(Cont + (a * CTD_ContainerItem) - CTD_ContainerItem + CTD_ContainerItemID) = ID Then
            iPosInCont = a
            Exit Function
        End If
    Next
End Function
Function nCont(ByVal Cont As Long, iCont As Long) As Integer
    nCont = mReadLong(CT_Start + ((Cont - 1) * CTD_Container) + ((iCont - 1) * CTD_ContainerItem) + CTD_ContainerItemCount)
End Function

Function GotoXYZ(X As Long, Y As Long, z As Long)
    If Compiled = True Then On Error Resume Next
    mWriteLong CH_gX, X
    mWriteLong CH_gY, Y
    mWriteLong CH_gZ, z
    tmp = BL_Player
    If tmp <> 0 Then
        mWriteLong tmp + BL_Wlk, 1
    End If
End Function

Function Stack_Items(ByVal ID As Long) As Boolean
    For a = 1 To mReadLong(CT_Start + CTD_ContainerAmount)
        If mReadLong(CT_Start + (a * CTD_ContainerItem) - CTD_ContainerItem + CTD_ContainerItemID) = ID Then
            FoundIn = FoundIn & a & "." & mReadLong(CT_Start + (a * CTD_ContainerItem) + CTD_ContainerItemCount - CTD_ContainerItem) & vbCrLf
        End If
    Next
    If FoundIn <> "" Then
        FoundIn = Left(FoundIn, Len(FoundIn) - 2)
        FoundIn = Split(FoundIn, vbCrLf)
        For a = UBound(FoundIn) To 0 Step -1
            If Split(FoundIn(a), ".")(1) < 100 Then
                MoveFrom = Split(FoundIn(0), ".")(0)
                moveTo = Split(FoundIn(a), ".")(0)
                MoveFn = Split(FoundIn(0), ".")(1)
                MoveTn = Split(FoundIn(a), ".")(1)
                If Int(MoveFn) <= (100 - Int(MoveTn)) Then MoveNum = Int(MoveFn) Else MoveNum = (100 - Int(MoveTn))
                Exit For
            End If
        Next
        If MoveNum > 0 And MoveFrom <> moveTo Then
            Stack_Items = True
            'frmMain.hdr = MoveFrom & " " & MoveTo & ", " & MoveNum
            sPck s2ba("F 0 78 FF FF 40 0 " & Hex(MoveFrom - 1) & " " & Hex(l2b(ID, 1)) & " " & Hex(l2b(ID, 2)) & " " & Hex(MoveFrom - 1) & " FF FF 40 0 " & Hex(moveTo - 1) & " " & Hex(MoveNum))
        End If
    End If
End Function
