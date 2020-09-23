Attribute VB_Name = "Module2"
'Private Const HKEY_CURRENT_CONFIG = &H80000005
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
'Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'Public Function GetAllValues(hKey As Long, strPath As String) As Boolean
'    Dim Cnt As Long, sSave As String
'    RegOpenKey hKey, strPath, hKey
'    Cnt = 0
'    Do
'        sSave = String(255, 0)
'        If RegEnumValue(hKey, Cnt, sSave, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
'        Cnt = Cnt + 1
'    Loop
'    RegCloseKey hKey
'    MsgBox Cnt
'End Function
