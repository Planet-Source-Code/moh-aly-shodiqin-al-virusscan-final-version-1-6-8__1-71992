VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim ValIsi
Const ERROR_NO_MORE_ITEMS = 259&

Public Sub CreateKey(hKey As Long, strPath As String)
On Error Resume Next
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
If lRegResult <> ERROR_SUCCESS Then
  ' there is a problem
End If
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
On Error Resume Next
Dim lRegResult As Long
lRegResult = RegDeleteKey(hKey, strPath)
End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
On Error Resume Next
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegDeleteValue(hCurKey, strValue)
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
On Error Resume Next
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

' Set up default value
If Not IsEmpty(Default) Then
  GetSettingString = Default
Else
  GetSettingString = ""
End If

' Open the key and get length of string
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then
  If lValueType = REG_SZ Then
    ' initialise string buffer and retrieve string
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
    
    ' format string
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      GetSettingString = Left$(strBuffer, intZeroPos - 1)
      'Code Put Here
    Else
      GetSettingString = strBuffer
    End If

  End If

Else
  ' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
On Error Resume Next
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long
On Error Resume Next
Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long

' Set up default value
If Not IsEmpty(Default) Then
  GetSettingLong = Default
Else
  GetSettingLong = 0
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lDataBufferSize = 4       ' 4 bytes = 32 bits = long

lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_DWORD Then
    GetSettingLong = lBuffer
    'Code Put Here
  End If

Else
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
On Error Resume Next
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
On Error Resume Next
Dim lValueType As Long
Dim byBuffer() As Byte
Dim lDataBufferSize As Long
Dim lRegResult As Long
Dim hCurKey As Long

' setup default value
If Not IsEmpty(Default) Then
  If varType(Default) = vbArray + vbByte Then
    GetSettingByte = Default
  Else
    GetSettingByte = 0
  End If

Else
  GetSettingByte = 0
End If

' Open the key and get number of bytes
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = reg_binarY Then
    ' initialise buffers and retrieve value
    ReDim byBuffer(lDataBufferSize - 1) As Byte
    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
    GetSettingByte = byBuffer
    'Code Put Here
  End If

Else
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, byData As Long)
' Make sure that the array starts with element 0 before passing it!
' (otherwise it will not be saved!)
Dim lRegResult As Long
On Error Resume Next
Dim hCurKey As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
' Pass the first array element and length of array
lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, 3, byData, 4)
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetAllKeys(hKey As Long, strPath As String) As Variant
' Returns: an array in a variant of strings
On Error Resume Next
Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim strNames() As String
Dim intZeroPos As Integer

lCounter = 0

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do

  'initialise buffers (longest possible length=255)
  lDataBufferSize = 255
  strBuffer = String(lDataBufferSize, " ")
  lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)

  If lRegResult = ERROR_SUCCESS Then
  
    'tidy up string and save it
    ReDim Preserve strNames(lCounter) As String
    
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      strNames(UBound(strNames)) = Left$(strBuffer, intZeroPos - 1)
      ValIsi = strNames(UBound(strNames))
      'Code Put Here
    Else
      strNames(UBound(strNames)) = strBuffer
    End If

    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop

GetAllKeys = strNames
End Function

Public Function GetAllValues(hKey As Long, strPath As String) As Variant
' Returns: a 2D array.
' (x,0) is value name
' (x,1) is value type (see constants)
On Error Resume Next
Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim strValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim strNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do
  ' Initialise bufffers
  lValueNameSize = 255
  strValueName = String$(lValueNameSize, " ")
  lDataBufferSize = 4000
  
  lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
  
  If lRegResult = ERROR_SUCCESS Then
    
    ' Save the type
    ReDim Preserve strNames(lCounter) As String
    ReDim Preserve lTypes(lCounter) As Long
    lTypes(UBound(lTypes)) = lValueType
    
    'Tidy up string and save it
    intZeroPos = InStr(strValueName, Chr$(0))
    Dim A, b, c
    If intZeroPos > 0 Then
      strNames(UBound(strNames)) = Left$(strValueName, intZeroPos - 1)
      'GetSettingString hKey, strPath, strNames(UBound(strNames))
      'GetSettingByte hKey, strPath, strNames(UBound(strNames))
      'GetSettingLong hKey, strPath, strNames(UBound(strNames))
      'Set itmX = Form1.ListView1.ListItems.Add(1, , strNames(UBound(strNames)))
       'code here
    Else
      strNames(UBound(strNames)) = strValueName
    End If

    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop

'Move data into array
Dim Finisheddata() As Variant
ReDim Finisheddata(UBound(strNames), 0 To 1) As Variant

For lCounter = 0 To UBound(strNames)
  Finisheddata(lCounter, 0) = strNames(lCounter)
  Finisheddata(lCounter, 1) = lTypes(lCounter)
Next

GetAllValues = Finisheddata

End Function

' 6 Februari 2009
Sub GetEnumValue(oKey As Long, regPath As String, result As Collection)
    Dim hKey As Long, Cnt As Long, sName As String, sData As String, RET As Long, RetData As Long
    Const BUFFER_SIZE As Long = 255
    Dim data(3) As String
    
    If RegOpenKey(oKey, regPath, hKey) = 0 Then

        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        RET = BUFFER_SIZE
        RetData = BUFFER_SIZE

        While RegEnumValue(hKey, Cnt, sName, RET, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            If RetData > 0 Then
               data(0) = Left$(sName, RET)
               data(1) = Left$(sData, RetData - 1)
               data(2) = oKey
               data(3) = regPath
               result.Add data
            End If
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            RET = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        RegCloseKey hKey
    End If
End Sub



