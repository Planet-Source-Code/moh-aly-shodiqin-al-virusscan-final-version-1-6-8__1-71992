Attribute VB_Name = "mRegistry"
Option Explicit

Public CekSetting As Boolean, cekLoad As Boolean

Sub ForceCacheRefresh()

   Dim hKey As Long
   Dim dwKeyType As Long

   Dim dwDataType As Long
   Dim dwDataSize As Long

   Dim sKeyName As String
   Dim sValue As String
   Dim sData As String
   Dim sDataRet As String

   Dim tmp As Long
   Dim sNewValue As String
   Dim dwNewValue As Long
   Dim success As Long
   
'HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\ShellIconSize
'HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\Shell Icon Size

'1. open HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics
'2. get the type of value and its size stored at value "Shell Icon Size"
'3. get value of "Shell Icon Size"
'4. change this value (i.e. decrement the value by one)
'5. write it back to the registry
'6. call SendMessageTimeout HWND_BROADCAST
'7. return "Shell Icon Size" to its original setting
'8. call SendMessageTimeout HWND_BROADCAST again
'9. close the key

''''''''''''''''''''''''
'Sample Debug output
''''''''''''''''''''''''
'RegKeyOpen = 468
'RegGetStringSize = 3
'RegGetStringValue = 32
'Changing to = 31
'Changing back to = 32

  ''''''''''''''''''''''''
  '1. open key
   dwKeyType = HKEY_CURRENT_USER
   sKeyName = "Control Panel\Desktop\WindowMetrics"
   sValue = "Shell Icon Size"
   
   hKey = RegKeyOpen(HKEY_CURRENT_USER, sKeyName)
   
   If hKey <> 0 Then
   
      Debug.Print "RegKeyOpen = "; hKey
      
     ''''''''''''''''''''''''
     '2. Determine the size and type of data to be read.
     'In this case it should be a string (REG_SZ) value.
      dwDataSize = RegGetStringSize(ByVal hKey, sValue, dwDataType)
      
      Debug.Print "RegGetStringSize = "; dwDataSize
      
      If dwDataSize > 0 Then

        ''''''''''''''''''''''''
        '3. get the value for that key
         sDataRet = RegGetStringValue(hKey, sValue, dwDataSize)
         
        'if a value returned
         If sDataRet > "" Then
         
            Debug.Print "RegGetStringValue = "; sDataRet
            
           ''''''''''''''''''''''''
           '4, 5. convert sDataRet to a number and subtract 1,
           'convert back to a string, define the size
           'of the new string, and write it to the registry
            tmp = CLng(sDataRet)
            tmp = tmp - 1
            sNewValue = CStr(tmp) & Chr$(0)
            dwNewValue = Len(sNewValue)

            Debug.Print "Changing to = "; sNewValue
            
            If RegWriteStringValue(hKey, _
                                   sValue, _
                                   dwDataType, _
                                   sNewValue) = ERROR_SUCCESS Then
                                   
                                   
              ''''''''''''''''''''''''
              '6. because the registry was changed, broadcast
              'the fact passing SPI_SETNONCLIENTMETRICS,
              'with a timeout of 10000 milliseconds (10 seconds)
               Call SendMessageTimeout(HWND_BROADCAST, _
                                       WM_SETTINGCHANGE, _
                                       SPI_SETNONCLIENTMETRICS, _
                                       0&, SMTO_ABORTIFHUNG, _
                                       10000&, success)
                                       
              ''''''''''''''''''''''''
              '7. the desktop will have refreshed with the
              'new (shrunken) icon size. Now restore things
              'back to the correct settings by again writing
              'to the registry and posing another message.
               sDataRet = sDataRet & Chr$(0)
               
               Debug.Print "Changing back to = "; sDataRet
               
               Call RegWriteStringValue(hKey, _
                                       sValue, _
                                       dwDataType, _
                                       sDataRet)
                  
              ''''''''''''''''''''''''
              '8.  broadcast the change again
               Call SendMessageTimeout(HWND_BROADCAST, _
                                       WM_SETTINGCHANGE, _
                                       SPI_SETNONCLIENTMETRICS, _
                                       0&, SMTO_ABORTIFHUNG, _
                                       10000&, success)
               
            
            End If   'If RegWriteStringValue
         End If   'If sDataRet > ""
      End If   'If dwDataSize > 0
   End If   'If hKey > 0
         
  
  ''''''''''''''''''''''''
  '9. clean up
   Call RegCloseKey(hKey)

End Sub


Function RegGetStringSize(ByVal hKey As Long, _
                                  ByVal sValue As String, _
                                  dwDataType As Long) As Long

   Dim success As Long
   Dim dwDataSize As Long
   
   success = RegQueryValueEx(hKey, _
                             sValue, _
                             0&, _
                             dwDataType, _
                             ByVal 0&, _
                             dwDataSize)
         
   If success = ERROR_SUCCESS Then
      If dwDataType = REG_SZ Then
      
         RegGetStringSize = dwDataSize
         
      End If
   End If

End Function

Function RegKeyOpen(dwKeyType As Long, sKeyPath As String) As Long

   Dim hKey As Long
   Dim dwOptions As Long
   Dim SA As SECURITY_ATTRIBUTES
   
   SA.nLength = Len(SA)
   SA.bInheritHandle = False
   
   dwOptions = 0&
   If RegOpenKeyEx(dwKeyType, _
                   sKeyPath, dwOptions, _
                   KEY_ALL_ACCESS, hKey) = ERROR_SUCCESS Then
   
      RegKeyOpen = hKey
      
   End If

End Function

Function RegGetStringValue(ByVal hKey As Long, _
                                   ByVal sValue As String, _
                                    dwDataSize As Long) As String

   Dim sDataRet As String
   Dim dwDataRet As Long
   Dim success As Long
   Dim pos As Long
   
  'get the value of the passed key
   sDataRet = Space$(dwDataSize)
   dwDataRet = Len(sDataRet)
   
   success = RegQueryValueEx(hKey, sValue, _
                             ByVal 0&, dwDataSize, _
                             ByVal sDataRet, dwDataRet)

   If success = ERROR_SUCCESS Then
      If dwDataRet > 0 Then
      
         pos = InStr(sDataRet, Chr$(0))
         RegGetStringValue = Left$(sDataRet, pos - 1)
         
      End If
   End If
   
End Function

Public Function RegWriteStringValue(ByVal hKey, _
                                    ByVal sValue, _
                                    ByVal dwDataType, _
                                    sNewValue) As Long

   Dim success As Long
   Dim dwNewValue As Long
   
   dwNewValue = Len(sNewValue)
   
   If dwNewValue > 0 Then
      RegWriteStringValue = RegSetValueExString(hKey, _
                                                sValue, _
                                                0&, _
                                                dwDataType, _
                                                sNewValue, _
                                                dwNewValue)
                                           
   End If

End Function





