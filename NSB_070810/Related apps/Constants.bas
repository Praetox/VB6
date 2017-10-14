Attribute VB_Name = "Constants"
'=========Constants And Addresses=============
'Player Position
Public Const PLAYER_X = &H610C28
Public Const PLAYER_Y = &H610C24
Public Const PLAYER_Z = &H610C20

'Player Id
Public Const PLAYER_ID = &H6059D0

'Map Reading
Public Const MapTileDist = 172
Public Const MapObjectDist = 12
Public Const MapObjectIdDist = 0
Public Const MapObjectDataDist = 4
Public Const MapObjectIdOffset = 4
Public Const MapObjectDataOffset = 8

Public Const MAP_POINTER = &H615738

'Tile Data
Public Type TileData
    TileNum As Long
    MapCount As Long
    TileId As Long
    ObjectId(1 To 9) As Long
    ObjectInfo(1 To 9) As Long
    ObjectInfoEx(1 To 9) As Long
    posX As Long
    posY As Long
    posZ As Long
End Type
