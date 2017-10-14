Attribute VB_Name = "Packets"
Declare Sub SendPacket Lib "packet.dll" (ByVal ProcessID As Long, ByRef Packet() As Byte, Optional ByVal Encrypt As Byte = True, Optional ByVal SafeArray As Byte = True)

Sub sPck(pck() As Byte)
    If Compiled = True Then On Error Resume Next
    'For a = 0 To UBound(pck)
    '    tmp = tmp & Hex(pck(a)) & " · "
    'Next
    'MsgBox tmp
    GetWindowThreadProcessId tHvnd, P_ID
    SendPacket P_ID, pck
End Sub

Sub Attack(ID As Long)
    If Compiled = True Then On Error Resume Next
    Dim pck(6) As Byte
    pck(0) = &H5
    pck(1) = &H0
    pck(2) = &HA1
    pck(3) = l2b(ID, 1)
    pck(4) = l2b(ID, 2)
    pck(5) = l2b(ID, 3)
    pck(6) = &H40
    sPck pck
    mWriteLong BOX_3, ID
End Sub

Sub Logout()
    If Compiled = True Then On Error Resume Next
    'Dim pck(2) As Byte
    'pck(0) = &H1
    'pck(1) = &H0
    'pck(2) = &H14
    sPck s2ba("1 0 14") 'pck
End Sub

Sub aFollow()
    If Compiled = True Then On Error Resume Next
    'Dim pck(5) As Byte
    'pck(0) = &H4
    'pck(1) = &H0
    'pck(2) = &HA0
    'pck(3) = &H1
    'pck(4) = &H1
    'pck(5) = &H1
    sPck s2ba("4 0 A0 1 1 1") 'pck
End Sub

Sub use_WithCont(ID As Long, X As Long, Y As Long, Z As Long, Spot As Long, Cont As Long, Optional TileID As Long)
    If Compiled = True Then On Error Resume Next
    If Spot = -1 Or Cont = -1 Then Exit Sub
    Dim pck(18) As Byte
    pck(0) = &H11
    pck(1) = &H0
    pck(2) = &H83
    pck(3) = &HFF
    pck(4) = &HFF
    pck(5) = (Cont + 63)
    pck(6) = &H0
    pck(7) = Spot - 1
    pck(8) = lbol(ID)
    pck(9) = hbol(ID)
    pck(10) = Spot - 1
    pck(11) = lbol(X)
    pck(12) = hbol(X)
    pck(13) = lbol(Y)
    pck(14) = hbol(Y)
    pck(15) = Z
    If TileID = 0 Then 'mob
        pck(16) = &H63
        pck(17) = &H0
        pck(18) = &H1
    Else 'ground
        pck(16) = lbol(TileID)
        pck(17) = hbol(TileID)
        pck(18) = &H0
    End If
    sPck pck
End Sub

Sub lodItem(Slot As Integer, ID As Long, Cont As Long, iCont As Long, Optional iCount As Integer = 1)
    If Compiled = True Then On Error Resume Next
    Dim pck(16) As Byte
    pck(0) = &HF
    pck(1) = &H0
    pck(2) = &H78
    pck(3) = &HFF
    pck(4) = &HFF
    pck(5) = (Cont + 63)
    pck(6) = &H0
    pck(7) = iCont - 1
    pck(8) = lbol(ID)
    pck(9) = hbol(ID)
    pck(10) = iCont - 1
    pck(11) = &HFF
    pck(12) = &HFF
    pck(13) = Slot '2=amulet, 6=left, 9=ring, a=ammo
    pck(14) = &H0
    pck(15) = &H0
    pck(16) = iCount
    sPck pck
End Sub

Sub tosItem(Slot As Integer, ID As Long, Cont As Long, Optional iCount As Integer = 1)
    If Compiled = True Then On Error Resume Next
    Dim pck(16) As Byte
    pck(0) = &HF
    pck(1) = &H0
    pck(2) = &H78
    pck(3) = &HFF
    pck(4) = &HFF
    pck(5) = Slot '2=amulet, 6=left, 9=ring, a=ammo
    pck(6) = &H0
    pck(7) = &H0
    pck(8) = lbol(ID)
    pck(9) = hbol(ID)
    pck(10) = &H0
    pck(11) = &HFF
    pck(12) = &HFF
    pck(13) = (Cont + 63)
    pck(14) = &H0
    pck(15) = &H0
    pck(16) = iCount
    sPck pck
End Sub

Sub useItem(ID As Long, Cont As Long, iCont As Long)
    If Compiled = True Then On Error Resume Next
    Dim pck(11) As Byte
    pck(0) = &HA
    pck(1) = &H0
    pck(2) = &H82
    pck(3) = &HFF
    pck(4) = &HFF
    pck(5) = Cont + 63
    pck(6) = &H0
    pck(7) = iCont - 1
    pck(8) = lbol(ID)
    pck(9) = hbol(ID)
    pck(10) = iCont - 1
    pck(11) = Cont - 1
    sPck pck
End Sub

Sub OpenBody(X As Long, Y As Long, Z As Long, MobID As Long, RelPos As Long, OpenBPs As Long)
    If Compiled = True Then On Error Resume Next
    Dim pck(11) As Byte
    pck(0) = &HA
    pck(1) = &H0
    pck(2) = &H82
    pck(3) = l2b(X, 1)
    pck(4) = l2b(X, 2)
    pck(5) = l2b(Y, 1)
    pck(6) = l2b(Y, 2)
    pck(7) = Z
    pck(8) = l2b(MobID, 1)
    pck(9) = l2b(MobID, 2)
    pck(10) = RelPos  '2 if you're standing next to creature. 3 if you're standing on it
    pck(11) = OpenBPs 'Number of open backpacks
    sPck pck
End Sub
