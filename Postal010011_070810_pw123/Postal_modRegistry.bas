Attribute VB_Name = "modRegistry"
Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult&, lpdwDisposition&)
Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal dwRes&, ByVal dwType&, lpDataBuff As Any, ByVal nSize&)
Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const REG_SZ = 1&
Const REG_DWORD = 4&
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const Key_Write = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY

Sub RegSetValue(H_KEY&, RSubKey$, ValueName$, RegValue$)
    'H_KEY must be one of the Key Constants
    Dim lRtn&         'returned by registry functions, should be 0&
    Dim hKey&         'return handle to opened key
    Dim lpDisp&
    Dim Sec_Att As SECURITY_ATTRIBUTES
    Sec_Att.nLength = 12&
    Sec_Att.lpSecurityDescriptor = 0&
    Sec_Att.bInheritHandle = False
    If RegValue = "" Then RegValue = " "
    
        lRtn = RegCreateKeyEx(H_KEY, RSubKey, 0&, "", 0&, Key_Write, Sec_Att, hKey, lpDisp)
        If lRtn <> 0 Then
            Exit Sub       'No key open, so leave
        End If
        lRtn = RegSetValueEx(hKey, ValueName, 0&, REG_SZ, ByVal RegValue, CLng(Len(RegValue) + 1))
        lRtn = RegCloseKey(hKey)
End Sub

Sub ieImages(bShow As Boolean)
    Dim sShow As String: If bShow Then sShow = "yes" Else sShow = "no"
    modRegistry.RegSetValue &H80000001, "Software\Microsoft\Internet Explorer\Main", "Display Inline Images", sShow
End Sub
