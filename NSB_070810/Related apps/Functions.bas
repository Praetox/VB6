Attribute VB_Name = "Functions"
'Declarations
Option Explicit

'Win32 API
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Constants for reading memory
Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION
Public Const PROCESS_ALL_ACCESS = &H1F0FFF

'============Map Reading Functions===============
Public Function Get_PlayerTile() As Long 'Function for getting number of player tile
Dim MapBegins As Long, StackSize As Long, ObjectId As Long, Data As Long, i As Long, H As Long, E As Long

MapBegins = getMapPointer 'Let's get map pointer first
For i = 0 To 2015 ' Loop all the tiles trough
    H = (MapTileDist * i)
    StackSize = Memory_ReadByte(Tibia_Hwnd, getMapPointer + H) 'Read number of objects in tile
    If StackSize > 1 Then ' There must be at least 2 objects, so we no it's player's tile (Tile and charecter)
        For E = 0 To StackSize - 1 'Loop trough all the objects
            ObjectId = Memory_ReadLong(Tibia_Hwnd, getMapPointer + (i * MapTileDist) + (E * MapObjectDist) + MapObjectIdOffset)
            If ObjectId = &H63 Then '63h = 99d (creature id), so we know there is creature in tile
                Data = Memory_ReadLong(Tibia_Hwnd, getMapPointer + (i * MapTileDist) + (E * MapObjectDist) + MapObjectDataOffset)
                If Data = getPlayerId Then 'If objects id is equal with player's id
                    Get_PlayerTile = i
                    Exit Function 'We have it!
                End If
            End If
        Next
    End If
Next
End Function

Public Function Get_TileInfo(TileNum As Long) As TileData ' Function for reading Data from tile with Tile Number
Dim MapTile_Address As Long, MapBegins As Long, N As Long

'MapBegins = getMapPointer
MapTile_Address = getMapPointer + (TileNum * MapTileDist) + 4

With Get_TileInfo
    .TileNum = TileNum
    .MapCount = Memory_ReadLong(Tibia_Hwnd, MapTile_Address - 4)
    .TileId = Memory_ReadLong(Tibia_Hwnd, MapTile_Address)
    .posX = Tile_Coords(TileNum, "x")
    .posY = Tile_Coords(TileNum, "y")
    .posZ = Tile_Coords(TileNum, "z")
    
    For N = 1 To (.MapCount - 1) 'Looping trough all objects in tile ( count - 1 because first object is tile itself)
        .ObjectId(N) = Memory_ReadLong(Tibia_Hwnd, MapTile_Address + (MapObjectDist * N) + MapObjectIdDist)
        .ObjectInfo(N) = Memory_ReadLong(Tibia_Hwnd, MapTile_Address + (MapObjectDist * N) + MapObjectDataDist)
    Next N
End With
End Function

Public Function Tile_Coords(TileNum As Long, xyz As String) As Long 'Function for getting map coord from tilenumber
Dim Z As Long, Y As Long, X As Long
Select Case xyz
    Case "z"
        Z = Fix(TileNum / (14 * 18))
        Tile_Coords = Z
    Case "y"
        Z = Fix(TileNum / (14 * 18))
        Y = Fix((TileNum - Z * 14 * 18) / 18)
        Tile_Coords = Y
    Case "x"
        Z = Fix(TileNum / (14 * 18))
        Y = Fix((TileNum - Z * 14 * 18) / 18)
        X = Fix((TileNum - Z * 14 * 18 - Y * 18))
        Tile_Coords = X
    Case Else
        Exit Function
End Select
End Function
Public Function getMapPointer() As Long
getMapPointer = Memory_ReadLong(Tibia_Hwnd, MAP_POINTER)

End Function

Public Function getPlayerId() As Long
getPlayerId = Memory_ReadLong(Tibia_Hwnd, PLAYER_ID)
End Function


'===========Memory Reading Functions================

Public Function Tibia_Hwnd() As Long
    
  'Return the value of the Tibia Window
  'Use to find the window alot
    
  'Find Tibia's hwnd or Window
  Dim tibiaclient As Long
  tibiaclient = FindWindow("tibiaclient", vbNullString)
  
  'Return hwnd to function
  Tibia_Hwnd = tibiaclient
  
End Function

Public Function Memory_ReadByte(windowHwnd As Long, Address As Long) As Byte
  
   ' Declare some variables we need
   Dim PID As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Byte   ' Byte
    
   ' First get a handle to the "game" window
   If (windowHwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId windowHwnd, PID
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_VM_READ, False, PID)
   If (phandle = 0) Then Exit Function
   
   ' Read Long
   ReadProcessMemory phandle, Address, valbuffer, 1, 0&
       
   ' Return
   Memory_ReadByte = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function
Public Function Memory_ReadLong(windowHwnd As Long, Address As Long) As Long
  
   ' Declare some variables we need
   Dim PID As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Long   ' Long
    
   ' First get a handle to the "game" window
   If (windowHwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId windowHwnd, PID
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_VM_READ, False, PID)
   If (phandle = 0) Then Exit Function
   
   ' Read Long
   ReadProcessMemory phandle, Address, valbuffer, 4, 0&
       
   ' Return
   Memory_ReadLong = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function

