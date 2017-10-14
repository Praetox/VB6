Attribute VB_Name = "Memory"
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public tHvnd As Long

Public Function hVnd() As Long
    hVnd = FindWindow("tibiaclient", vbNullString)
End Function
Public Function mReadByte(Address As Long, Optional cVnd As Long = 0) As Byte
    If Compiled = True Then On Error Resume Next
    Dim PID As Long, pHnd As Long: If cVnd = 0 Then cVnd = tHvnd
    If cVnd = 0 Then Exit Function
    GetWindowThreadProcessId cVnd, PID
    pHnd = OpenProcess(&H10, False, PID)
    If pHnd = 0 Then Exit Function
    ReadProcessMemory pHnd, Address, mReadByte, 1, 0&
    CloseHandle pHnd
End Function
Public Function mReadLong(Address As Long, Optional cVnd As Long = 0) As Long
    If Compiled = True Then On Error Resume Next
    Dim PID As Long, pHnd As Long: If cVnd = 0 Then cVnd = tHvnd
    If cVnd = 0 Then Exit Function
    GetWindowThreadProcessId cVnd, PID
    pHnd = OpenProcess(&H10, False, PID)
    If pHnd = 0 Then Exit Function
    ReadProcessMemory pHnd, Address, mReadLong, 4, 0&
    CloseHandle pHnd
End Function
Public Function mReadString(Address As Long, Optional cVnd As Long = 0) As String
    If Compiled = True Then On Error Resume Next
    Dim PID As Long, pHnd As Long, str(255) As Byte: If cVnd = 0 Then cVnd = tHvnd
    If cVnd = 0 Then Exit Function
    GetWindowThreadProcessId cVnd, PID
    pHnd = OpenProcess(&H10, False, PID)
    If pHnd = 0 Then Exit Function
    ReadProcessMemory pHnd, Address, str(0), 255, 0&
    mReadString = chop(StrConv(str, vbUnicode))
    CloseHandle pHnd
End Function
Public Function mWriteString(Address As Long, ByVal ToWrite As String, Optional cVnd As Long = 0)
    If Compiled = True Then On Error Resume Next
    Dim PID As Long, pHnd As Long, str() As Byte: If cVnd = 0 Then cVnd = tHvnd
    If cVnd = 0 Then Exit Function
    GetWindowThreadProcessId cVnd, PID
    pHnd = OpenProcess(&H438, False, PID)
    If pHnd = 0 Then Exit Function
    str = StrConv(ToWrite & Chr(0), vbFromUnicode)
    WriteProcessMemory pHnd, Address, str(0), UBound(str) + 1, 0&
    CloseHandle pHnd
End Function
Public Sub mWriteByte(Address As Long, ToWrite As Byte, Optional cVnd As Long = 0)
    If Compiled = True Then On Error Resume Next
    Dim PID As Long, pHnd As Long: If cVnd = 0 Then cVnd = tHvnd
    If cVnd = 0 Then Exit Sub
    GetWindowThreadProcessId cVnd, PID
    pHnd = OpenProcess(&H438, False, PID)
    If pHnd = 0 Then Exit Sub
    WriteProcessMemory pHnd, Address, ToWrite, 1, 0&
    CloseHandle pHnd
End Sub
Public Sub mWriteLong(Address As Long, ToWrite As Long, Optional cVnd As Long = 0)
    If Compiled = True Then On Error Resume Next
    Dim PID As Long, pHnd As Long: If cVnd = 0 Then cVnd = tHvnd
    If cVnd = 0 Then Exit Sub
    GetWindowThreadProcessId cVnd, PID
    pHnd = OpenProcess(&H438, False, PID)
    If pHnd = 0 Then Exit Sub
    WriteProcessMemory pHnd, Address, ToWrite, 4, 0&
    CloseHandle pHnd
End Sub
Public Function chop(txt As String)
    If Compiled = True Then On Error Resume Next
    tmp = InStr(1, txt, Chr(0))
    If tmp > 0 Then chop = Left(txt, tmp - 1) Else chop = txt
End Function
Public Sub Window_Ontop(Widoww As Integer)
    If Compiled = True Then On Error Resume Next
    SetWindowPos Widoww, -1, 0, 0, 0, 0, &H1 Or &H10 Or &H2
End Sub
Public Sub Window_NotOntop(Widoww As Integer)
    If Compiled = True Then On Error Resume Next
    SetWindowPos Widoww, -2, 0, 0, 0, 0, &H1 Or &H10 Or &H2
End Sub
Sub Opacity(ByVal pHvnd As Long, ByVal Opacity As Integer)
    If Compiled = True Then On Error Resume Next
    SetWindowLong pHvnd, (-20), GetWindowLong(pHvnd, (-20)) Or &H80000
    SetLayeredWindowAttributes pHvnd, 0, Opacity, &H2
End Sub
Sub Opaque(ByVal pHvnd As Long)
    If Compiled = True Then On Error Resume Next
    SetWindowLong pHvnd, (-20), GetWindowLong(pHvnd, (-20)) And Not &H80000
    SetLayeredWindowAttributes pHvnd, 0, 0, &H2
End Sub
