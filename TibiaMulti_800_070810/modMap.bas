Attribute VB_Name = "Map"
Type TileData
    TileNum As Long
    MapCount As Long
    TileID As Long
    TopID As Long
    ObjectId(1 To 9) As Long
    ObjectInfo(1 To 9) As Long
    ObjectInfoEx(1 To 9) As Long
    posX As Long
    posY As Long
    posZ As Long
End Type

Function Map_PlayerTileNum() As Long
    Dim MapBegins As Long, StackSize As Long, ObjectId As Long, Data As Long, i As Long, h As Long, E As Long
    MapBegins = mReadLong(MAP_POINTER)
    PlayerID = mReadLong(CH_ID)
    For i = 0 To 2015
        h = (Map_TileDist * i)
        StackSize = mReadByte(MapBegins + h)
        If StackSize > 1 Then
            For E = 0 To StackSize - 1
                ObjectId = mReadLong(MapBegins + (i * Map_TileDist) + (E * Map_ObjectDist) + Map_ObjectIdOffset)
                If ObjectId = &H63 Then
                    Data = mReadLong(MapBegins + (i * Map_TileDist) + (E * Map_ObjectDist) + Map_ObjectDataOffset)
                    If Data = PlayerID Then
                        Map_PlayerTileNum = i
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function

Function Map_TileInfo(TileNum As Long) As TileData
    Dim MapTile_Address As Long, MapBegins As Long, N As Long
    MapTile_Address = mReadLong(MAP_POINTER) + (TileNum * Map_TileDist)
    With Map_TileInfo
        .TileNum = TileNum
        .MapCount = mReadLong(MapTile_Address)
        .TileID = mReadLong(MapTile_Address + Map_ObjectIdOffset)
        .TopID = mReadLong(MapTile_Address + Map_TopObject)
        .posX = Map_TilePos(TileNum, "x")
        .posY = Map_TilePos(TileNum, "y")
        .posZ = Map_TilePos(TileNum, "z")
        For N = 1 To (.MapCount - 1)
            .ObjectId(N) = mReadLong(MapTile_Address + Map_ObjectIdOffset + (Map_ObjectDist * N) + Map_ObjectIdDist)
            .ObjectInfo(N) = mReadLong(MapTile_Address + Map_ObjectIdOffset + (Map_ObjectDist * N) + Map_ObjectDataDist)
        Next N
    End With
End Function

Function Map_TilePos(TileNum As Long, xyz As String) As Long
    Dim X As Long, Y As Long, z As Long
    Select Case xyz
        Case "z"
            z = Fix(TileNum / (14 * 18))
            Map_TilePos = z
        Case "y"
            z = Fix(TileNum / (14 * 18))
            Y = Fix((TileNum - z * 14 * 18) / 18)
            Map_TilePos = Y
        Case "x"
            z = Fix(TileNum / (14 * 18))
            Y = Fix((TileNum - z * 14 * 18) / 18)
            X = Fix((TileNum - z * 14 * 18 - Y * 18))
            Map_TilePos = X
        Case Else
            Exit Function
    End Select
End Function

Sub Fish_ShowFishyWater()
    MapBegin = mReadLong(MAP_POINTER)
    Dim a As Long, b As Long, TileID As Long
    For a = 0 To 2015
        b = a * Map_TileDist
        b = b + MapBegin
        TileID = mReadLong(b + Map_ObjectIdOffset)
        If TileID >= TILE_WATER_NOFISH_BEGIN And TileID <= TILE_WATER_NOFISH_END Then
            mWriteLong b + Map_ObjectIdOffset, 408 'TILE_WATER_OLD
        End If
    Next
End Sub

Function Fish_Map() As Long
    frm5.FishDisp_Display.Cls
    Dim Chartile As Long, CharMapX As Long, CharMapY As Long, CharMapZ As Long, _
        ThisMapX As Long, ThisMapY As Long, ThisMapZ As Long
    Chartile = Map_PlayerTileNum
    CharMapX = Map_TilePos(Chartile, "x")
    CharMapY = Map_TilePos(Chartile, "y")
    CharMapZ = Map_TilePos(Chartile, "z")
    MapBegin = mReadLong(MAP_POINTER)
    Dim a As Long, b As Long, TileID As Long
    For a = 0 To 2015
        ThisMapZ = Map_TilePos(a, "z")
        If ThisMapZ = CharMapZ Then
            ThisMapX = Map_TilePos(a, "x") - CharMapX
            ThisMapY = Map_TilePos(a, "y") - CharMapY
            b = a * Map_TileDist
            b = b + MapBegin
            TileID = mReadLong(b + Map_ObjectIdOffset)
            ThisMapX = ThisMapX + 7
            ThisMapY = ThisMapY + 5
                                            '     ___________________
            If ThisMapX > 17 Then           '    /                   \
                ThisMapX = ThisMapX - 18    '   /                     |
            ElseIf ThisMapX < 0 Then        '--<  Horizontal wrapping |
                ThisMapX = ThisMapX + 18    '   \                     |
            End If                          '____\___________________/
            If ThisMapY > 13 Then           '    /                   \
                ThisMapY = ThisMapY - 14    '   /                     |
            ElseIf ThisMapY < 0 Then        '--<   Vertical wrapping  |
                ThisMapY = ThisMapY + 14    '   \                     |
            End If                          '    \___________________/
            
            If ThisMapX >= 0 And ThisMapX <= 14 Then
                If ThisMapY >= 0 And ThisMapY <= 10 Then
                    If TileID >= TILE_WATER_FISH_BEGIN And TileID <= TILE_WATER_FISH_END Then
                        'MsgBox "Closest fishy tile is " & ThisMapX & "x" & ThisMapY & "."
                        frm5.FishDisp_Display.PSet ((2 + (ThisMapX * 5)) * 15, (2 + (ThisMapY * 5)) * 15), &H8000
                        Fish_Map = Fish_Map + 1
                    Else
                        frm5.FishDisp_Display.PSet ((2 + (ThisMapX * 5)) * 15, (2 + (ThisMapY * 5)) * 15), &H80
                    End If
                End If
            End If
        End If
    Next
End Function

Sub Fish_CastRod(ByVal Randomcasts As Boolean, StopWhenLowCap As Boolean)
    Dim Chartile As Long, CharMapX As Long, CharMapY As Long, CharMapZ As Long, _
        ThisMapX As Long, ThisMapY As Long, ThisMapZ As Long
    If StopWhenLowCap Then If mReadLong(CH_Cap) < 5 Then Exit Sub
    Chartile = Map_PlayerTileNum
    CharMapX = Map_TilePos(Chartile, "x")
    CharMapY = Map_TilePos(Chartile, "y")
    CharMapZ = Map_TilePos(Chartile, "z")
    MapBegin = mReadLong(MAP_POINTER)
    Dim a As Long, b As Long, TileID As Long
    For a = 0 To 2015
        ThisMapZ = Map_TilePos(a, "z")
        If ThisMapZ = CharMapZ Then
            ThisMapX = Map_TilePos(a, "x") - CharMapX
            ThisMapY = Map_TilePos(a, "y") - CharMapY
            b = a * Map_TileDist
            b = b + MapBegin
            TileID = mReadLong(b + Map_ObjectIdOffset)
            ThisMapX = ThisMapX + 7
            ThisMapY = ThisMapY + 5
                                            '     ___________________
            If ThisMapX > 17 Then           '    /                   \
                ThisMapX = ThisMapX - 18    '   /                     |
            ElseIf ThisMapX < 0 Then        '--<  Horizontal wrapping |
                ThisMapX = ThisMapX + 18    '   \                     |
            End If                          '____\___________________/
            If ThisMapY > 13 Then           '    /                   \
                ThisMapY = ThisMapY - 14    '   /                     |
            ElseIf ThisMapY < 0 Then        '--<   Vertical wrapping  |
                ThisMapY = ThisMapY + 14    '   \                     |
            End If                          '    \___________________/
            
            If ThisMapX >= 0 And ThisMapX <= 14 Then
                If ThisMapY >= 0 And ThisMapY <= 10 Then
                    If TileID >= TILE_WATER_FISH_BEGIN And TileID <= TILE_WATER_FISH_END Then
                        ThisMapX = mReadLong(CH_X) + (ThisMapX - 7)
                        ThisMapY = mReadLong(CH_Y) + (ThisMapY - 5)
                        If Randomcasts Then
                            FishyTiles = FishyTiles & ThisMapX & "," & ThisMapY & "," & TileID & vbCrLf
                        Else
                            use_WithCont 3483, ThisMapX, ThisMapY, mReadLong(CH_Z), 1, -53, TileID
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next
    If Randomcasts = False Or FishyTiles = "" Then Exit Sub
        FishyTiles = Split(FishyTiles, vbCrLf)
        tiletouse = (Rnd(1) * (UBound(FishyTiles) - 1)) \ 1
        ThisMapX = Split(FishyTiles(tiletouse), ",")(0)
        ThisMapY = Split(FishyTiles(tiletouse), ",")(1)
        TileID = Split(FishyTiles(tiletouse), ",")(2)
        use_WithCont 3483, ThisMapX, ThisMapY, mReadLong(CH_Z), 1, -53, TileID
End Sub
