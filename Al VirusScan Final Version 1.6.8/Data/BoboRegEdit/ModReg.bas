Attribute VB_Name = "ModReg"
'Most functions here are standard Registry access functions
'with modifications to load Listview/Treeview

'API for registry access
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const MAX_PATH = 256&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const STANDARD_RIGHTS_WRITE = &H20000
Private Const STANDARD_RIGHTS_EXECUTE = &H20000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SEDataValue = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Public SubKColl As Collection
Public ValColl As Collection
Public ValTypeColl As Collection
Public HasSubKeys() As Boolean
Public Sub SaveKey(mPath As String, sfile As String)
    'Use real regedit to export as .reg files
    Dim temp As String
    FileAppend "", sfile
    temp = GetDosPath(sfile)
    If FileExists(temp) Then Kill temp
    Shell "regedit /E " & temp & " " & Chr(34) & mPath & Chr(34)
End Sub
Public Sub ImportNode(sInFile As String)
    'Use real regedit to import .reg files
    Dim temp As String
    temp = GetDosPath(sInFile)
    Shell "regedit /I /S " & temp
End Sub

Public Sub CreateKey(hKey As Long, strPath As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function DeleteKey(ByVal hKey As Long, ByVal strPath As String) As Boolean
    Dim lRegResult As Long
    lRegResult = RegDeleteKey(hKey, strPath)
    If lRegResult = 0 Then
        DeleteKey = True
    Else
        DeleteKey = False
    End If
End Function
Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, strValue)
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    'Upgraded to read REG_EXPAND_SZ
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    If Not IsEmpty(Default) Then
        GetSettingString = Default
    Else
        GetSettingString = ""
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Or REG_EXPAND_SZ Then
            strBuffer = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingString = Left$(strBuffer, intZeroPos - 1)
            Else
                GetSettingString = strBuffer
            End If
            If lValueType = REG_EXPAND_SZ Then GetSettingString = StripTerminator(ExpandEnvStr(GetSettingString))
        End If
    Else
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long
    Dim lRegResult As Long
    Dim lValueType As Long
    Dim lBuffer As Long
    Dim lDataBufferSize As Long
    Dim hCurKey As Long
    If Not IsEmpty(Default) Then
        GetSettingLong = Default
    Else
        GetSettingLong = 0
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lDataBufferSize = 4
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetSettingLong = lBuffer
        End If
    Else
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
    Dim lValueType As Long
    Dim byBuffer() As Byte
    Dim lDataBufferSize As Long
    Dim lRegResult As Long
    Dim hCurKey As Long
    ReDim byBuffer(0 To 1) As Byte
    byBuffer(0) = 0
    If Not IsEmpty(Default) Then
        If VarType(Default) = vbArray + vbByte Then
            GetSettingByte = Default
        Else
            GetSettingByte = 0
        End If
    Else
        GetSettingByte = 0
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_BINARY Then
            If lDataBufferSize = 0 Then
                ReDim byBuffer(0) As Byte
                byBuffer(0) = 0
                GetSettingByte = byBuffer
            Else
                ReDim byBuffer(lDataBufferSize - 1) As Byte
                lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
                GetSettingByte = byBuffer
            End If
        End If
    Else
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, byData() As Byte)
    Dim lRegResult As Long
    Dim hCurKey As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Sub SaveSettingEmptyByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String)
    Dim lRegResult As Long
    Dim hCurKey As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, 0&, 0&)
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetAllKeys(hKey As Long, strPath As String) As Variant
    'Modified by Ian Northwood - thanks Ian
    Dim lRegResult As Long
    Dim lCounter As Long
    Dim hCurKey As Long
    Dim hCurKey2 As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim strnames() As String
    Dim temp As String
    Dim intZeroPos As Integer
    Dim strDummy As String
    
    If Len(strPath) > 0 Then temp = "\"
    lCounter = 0
    ReDim strnames(lCounter) As String
    strnames(lCounter) = "  "
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    Do
        DoEvents
        lDataBufferSize = 255
        strBuffer = String(lDataBufferSize, " ")
        lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)
        If lRegResult = ERROR_SUCCESS Then
            ReDim Preserve strnames(lCounter) As String
            ReDim Preserve HasSubKeys(lCounter) As Boolean
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                strnames(UBound(strnames)) = Left$(strBuffer, intZeroPos - 1)
            Else
                strnames(UBound(strnames)) = strBuffer
            End If
            
            If Right$(strPath, 1) = "\" Then
                strDummy = strPath + strnames(lCounter)
            Else
                strDummy = strPath + temp + strnames(lCounter)
            End If

            lRegResult = RegOpenKey(hKey, strDummy, hCurKey2)
            lDataBufferSize = 255
            strBuffer = String(lDataBufferSize, " ")
            lRegResult = RegEnumKey(hCurKey2, 0, strBuffer, lDataBufferSize)
            If lRegResult = ERROR_SUCCESS Then
                HasSubKeys(UBound(HasSubKeys)) = True
            Else
                HasSubKeys(UBound(HasSubKeys)) = False
            End If
            lCounter = lCounter + 1
        Else
            Exit Do
        End If
    Loop
    GetAllKeys = strnames
End Function

Public Function GetAllValues(hKey As Long, strPath As String, Optional DontAddToTree As Boolean = False, Optional mTypes As Variant) As Variant
    'Loads up listview with values
    Dim lItem As ListItem
    Dim lRegResult As Long
    Dim hCurKey As Long
    Dim lValueNameSize As Long
    Dim strValueName As String
    Dim lCounter As Long
    Dim byDataBuffer(4000) As Byte
    Dim lDataBufferSize As Long
    Dim lValueType As Long
    Dim z As Long, zx As Long
    Dim strnames() As String
    Dim lTypes() As Long
    Dim temp As String
    Dim intZeroPos As Integer
    Dim byTemp() As Byte
    Dim byTemp2 As String
    Dim byTemp3 As String
    ReDim byTemp(0)
    Dim found As Boolean
    If Right(strPath, 1) = "\" Then strPath = Left(strPath, Len(strPath) - 1)
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    ReDim strnames(lCounter)
    strnames(lCounter) = "  "
    Do
      lValueNameSize = 255
      strValueName = String$(lValueNameSize, " ")
      lDataBufferSize = 4000
      lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
      If lRegResult = ERROR_SUCCESS Then
        ReDim Preserve strnames(lCounter) As String
        ReDim Preserve lTypes(lCounter) As Long
        lTypes(UBound(lTypes)) = lValueType
        intZeroPos = InStr(strValueName, Chr$(0))
        If intZeroPos > 0 Then
            strnames(UBound(strnames)) = Left$(strValueName, intZeroPos - 1)
            lCounter = lCounter + 1
        Else
            strnames(UBound(strnames)) = strValueName
            lCounter = lCounter + 1
        End If
      Else
        Exit Do
      End If
    Loop
    If DontAddToTree Then
        mTypes = lTypes
        GetAllValues = strnames
        Exit Function
    End If
    frmMain.LV.ListItems.Clear
    If lCounter = 0 Then
        Set lItem = fMainForm.LV.ListItems.Add(, , "(Default)", , 3)
        lItem.SubItems(1) = "REG_SZ"
        lItem.SubItems(2) = "(value not set)"
        lItem.Tag = hKey
        Exit Function
    End If
    For zx = 0 To UBound(strnames)
        If Len(strnames(zx)) = 0 Then
            found = True
            Select Case lTypes(zx)
                Case 1
                    Set lItem = fMainForm.LV.ListItems.Add(, , "(Default)", , 3)
                    lItem.SubItems(1) = "REG_SZ"
                    lItem.SubItems(2) = GetSettingString(hKey, strPath, strnames(zx), "")
                    If lItem.SubItems(2) = "" Then lItem.SubItems(2) = "(value not set)"
                Case 2
                    Set lItem = fMainForm.LV.ListItems.Add(, , "(Default)", , 3)
                    lItem.SubItems(1) = "REG_EXPAND_SZ"
                    lItem.SubItems(2) = GetSettingString(hKey, strPath, strnames(zx), "")
                    If lItem.SubItems(2) = "" Then lItem.SubItems(2) = "(value not set)"
                Case 3
                    Set lItem = fMainForm.LV.ListItems.Add(, , "(Default)", , 3)
                    lItem.SubItems(1) = "REG_BINARY"
                    byTemp = GetSettingByte(hKey, strPath, strnames(zx))
                    byTemp2 = CStr(byTemp(0))
                    If Len(byTemp2) = 0 Or (byTemp2 = "0" And UBound(byTemp) = 0) Then
                        byTemp2 = "(zero length binary)"
                    Else
                        byTemp2 = ""
                        For z = 0 To UBound(byTemp)
                            byTemp2 = byTemp2 + " " + Format(LCase(Hex$(byTemp(z))), "00")
                        Next
                    End If
                    lItem.SubItems(2) = byTemp2
                Case 4
                    Set lItem = fMainForm.LV.ListItems.Add(, , "(Default)", , 3)
                    lItem.SubItems(1) = "REG_DWORD"
                    byTemp2 = Trim(Str(GetSettingLong(hKey, strPath, strnames(zx), 0)))
                    If Len(byTemp2) = 0 Or byTemp2 = "0" Then
                        byTemp2 = "0x00000000 (0)"
                    Else
                        byTemp3 = Hex$(Val(byTemp2))
                        byTemp2 = LCase("0x" + String(8 - Len(byTemp3), "0") + Hex$(Val(byTemp2)) + " (" + byTemp2 + ")")
                    End If
                    lItem.SubItems(2) = byTemp2
            End Select
            lItem.Tag = hKey
        End If
    Next
    If Not found Then
        Set lItem = fMainForm.LV.ListItems.Add(, , "(Default)", , 3)
        lItem.SubItems(1) = "REG_SZ"
        lItem.SubItems(2) = "(value not set)"
        lItem.Tag = hKey
    End If
    For lCounter = 0 To UBound(strnames)
        temp = strnames(lCounter)
        If temp <> "" Then
            Select Case lTypes(lCounter)
                Case 1
                    Set lItem = fMainForm.LV.ListItems.Add(, , temp, , 3)
                    lItem.SubItems(1) = "REG_SZ"
                    lItem.SubItems(2) = GetSettingString(hKey, strPath, strnames(lCounter), "")
                    If lItem.SubItems(2) = "" Then lItem.SubItems(2) = "(value not set)"
                    lItem.Tag = hKey
                Case 2
                    Set lItem = fMainForm.LV.ListItems.Add(, , temp, , 3)
                    lItem.SubItems(1) = "REG_EXPAND_SZ"
                    lItem.SubItems(2) = GetSettingString(hKey, strPath, strnames(lCounter), "")
                    If lItem.SubItems(2) = "" Then lItem.SubItems(2) = "(value not set)"
                    lItem.Tag = hKey
                Case 3
                    Set lItem = fMainForm.LV.ListItems.Add(, , temp, , 4)
                    lItem.SubItems(1) = "REG_BINARY"
                    byTemp = GetSettingByte(hKey, strPath, strnames(lCounter))
                    byTemp2 = CStr(byTemp(0))
                    If Len(byTemp2) = 0 Or (byTemp2 = "0" And UBound(byTemp) = 0) Then
                        byTemp2 = "(zero length binary)"
                    Else
                        byTemp2 = ""
                        For z = 0 To UBound(byTemp)
                            byTemp2 = byTemp2 + " " + Format(LCase(Hex$(byTemp(z))), "00")
                        Next
                    End If
                    lItem.SubItems(2) = byTemp2
                    lItem.Tag = hKey
                Case 4
                    Set lItem = fMainForm.LV.ListItems.Add(, , temp, , 4)
                    lItem.SubItems(1) = "REG_DWORD"
                    byTemp2 = Trim(Str(GetSettingLong(hKey, strPath, strnames(lCounter), 0)))
                    If Len(byTemp2) = 0 Then
                        byTemp2 = "0x00000000 (0)"
                    Else
                        byTemp3 = Hex$(Val(byTemp2))
                        byTemp2 = LCase("0x" + String(8 - Len(byTemp3), "0") + Hex$(Val(byTemp2)) + " (" + byTemp2 + ")")
                    End If
                    lItem.SubItems(2) = byTemp2
                    lItem.Tag = hKey
            End Select
        End If
    Next
    lRegResult = RegCloseKey(hCurKey)

End Function

Public Function CountAllKeys(hKey As Long) As Boolean
    Dim lRegResult As Long
    Dim lCounter As Long
    Dim hCurKey As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim strnames() As String
    Dim intZeroPos As Integer
    lCounter = 0
    lRegResult = RegOpenKey(hKey, "", hCurKey)
    Do
    lDataBufferSize = 255
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        ReDim Preserve strnames(lCounter) As String
        intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
            strnames(UBound(strnames)) = Left$(strBuffer, intZeroPos - 1)
            lCounter = lCounter + 1
        Else
            strnames(UBound(strnames)) = strBuffer
            lCounter = lCounter + 1
        End If
    Else
        Exit Do
    End If
    If lCounter > 0 Then
    CountAllKeys = True
    Exit Do
    End If
Loop
End Function


Public Function SafeKeyName(mHkey As Long, NewName As String, mNode As Node) As String
    'Used to generate a unique name for the Treeview
    Dim temp As String, mNames() As String, z As Long, found As Boolean
    Dim TryName As String, cnt As Long
    cnt = 1
    TryName = NewName + Trim(Str(cnt))
    temp = Right(mNode.Key, Len(mNode.Key) - InStr(1, mNode.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    mNames = GetAllKeys(mHkey, temp)
    If mNames(0) = "  " Then
        SafeKeyName = TryName
    Else
        Do
             found = False
             For z = 0 To UBound(mNames)
                 If mNames(z) = TryName Then
                     cnt = cnt + 1
                     TryName = NewName + Trim(Str(cnt))
                     found = True
                     Exit For
                 End If
             Next
             If Not found Then Exit Do
         Loop
        SafeKeyName = TryName
    End If
End Function

Public Function IsSafeKeyName(mHkey As Long, NewName As String, mNode As Node) As Boolean
    'Used when user renames an item in the Treeview
    'to ensure a key with same name does not already exist
    Dim temp As String, mNames() As String, z As Long, found As Boolean
    temp = Right(mNode.Key, Len(mNode.Key) - InStr(1, mNode.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    mNames = GetAllKeys(mHkey, temp)
    found = False
    If mNames(0) = "  " Then
        found = False
    Else
        For z = 0 To UBound(mNames)
            If mNames(z) = NewName Then
                found = True
                Exit For
            End If
        Next
    End If
    IsSafeKeyName = Not found
End Function
Public Function SafeValueName(mHkey As Long, mNode As Node, NewName As String) As String
    'Used to generate a unique name for the Listview
    Dim temp As String, mNames() As String, z As Long, found As Boolean
    Dim TryName As String, cnt As Long
    cnt = 1
    TryName = NewName + Trim(Str(cnt))
    temp = Right(mNode.Key, Len(mNode.Key) - InStr(1, mNode.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    mNames = GetAllValues(mHkey, temp, True, mTypes)
    If mNames(0) = "  " Then
        SafeValueName = TryName
    Else
        Do
             found = False
             For z = 0 To UBound(mNames)
                If mNames(z) = TryName Then
                    cnt = cnt + 1
                    TryName = NewName + Trim(Str(cnt))
                    found = True
                    Exit For
                 End If
             Next
             If Not found Then Exit Do
         Loop
        SafeValueName = TryName
    End If
End Function
Public Function IsSafeValueName(mHkey As Long, mNode As Node, NewName As String) As Boolean
    'Used when user renames an item in the Listview
    'to ensure a key with same name does not already exist
    Dim temp As String, mNames() As String, z As Long, found As Boolean
    temp = Right(mNode.Key, Len(mNode.Key) - InStr(1, mNode.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    mNames = GetAllValues(mHkey, temp, True, mTypes)
    found = False
    If mNames(0) = "  " Then
        found = False
    Else
        For z = 0 To UBound(mNames)
            If mNames(z) = NewName Then
                found = True
                Exit For
            End If
        Next
    End If
    IsSafeValueName = Not found
End Function
Public Sub ListSubVals(mHkey As Long, mPath As String)
    Dim count As Long, tmpStr() As String, tmpValStr() As String, z As Long, zx As Long
    Dim tmpType() As Long
    tmpStr = GetAllKeys(mHkey, mPath)
    If tmpStr(0) <> "  " Then
        For z = 0 To UBound(tmpStr)
            If tmpStr(z) <> "" Then
                SubKColl.Add mPath & "\" + tmpStr(z)
                ListSubVals mHkey, mPath & "\" + tmpStr(z)
            End If
        Next
    End If
    tmpValStr = GetAllValues(mHkey, mPath, True, tmpType)
    If tmpValStr(0) <> "  " Then
        For zx = 0 To UBound(tmpValStr)
            If tmpValStr(zx) <> "" Then
                ValColl.Add mPath & "\" + tmpValStr(zx)
                ValTypeColl.Add tmpType(zx)
            End If
        Next
    End If
    DoEvents
End Sub


Private Function ExpandEnvStr(sData As String) As String
    'This is cool - borrowed this
    Dim c As Long, s As String
    s = ""
    c = ExpandEnvironmentStrings(sData, s, c)
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    ExpandEnvStr = s
End Function

Public Function IsAKey(ByVal hKey As Long, strPath As String) As Boolean
    Dim lRegErr As Long
    lRegErr = RegOpenKey(hKey, strPath, hCurKey)
    If lRegErr = 0 Then IsAKey = True
End Function
