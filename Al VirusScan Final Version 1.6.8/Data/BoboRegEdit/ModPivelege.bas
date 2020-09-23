Attribute VB_Name = "ModPivelege"
Option Explicit

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(1) As LUID_AND_ATTRIBUTES
End Type

Private Const TOKEN_ADJUST_PRIVLEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const HKEY_USERS = &H80000003
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" _
(ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
                             TokenHandle As Long) As Long
                             
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias _
"LookupPrivilegeValueA" (ByVal lpSystemName As String, _
ByVal lpName As String, lpLuid As LUID) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
(ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long

Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) _
As Long

Private Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Retval As Long
Private strKeyName As String
Private MyToken As Long
Private TP As TOKEN_PRIVILEGES
Private RestoreLuid As LUID
Private BackupLuid As LUID


Public Sub EnableSavePiv()
    Retval = OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVLEGES _
       Or TOKEN_QUERY, MyToken)
'    If Retval = 0 Then MsgBox "OpenProcess: " & Err.LastDllError
    
    Retval = LookupPrivilegeValue(vbNullString, SE_RESTORE_NAME, _
       RestoreLuid)
'    If Retval = 0 Then MsgBox "LookupPrivileges: " & Err.LastDllError
    
    Retval = LookupPrivilegeValue(vbNullString, SE_BACKUP_NAME, BackupLuid)
'    If Retval = 0 Then MsgBox "LookupPrivileges: " & Retval
    
    TP.PrivilegeCount = 2
    TP.Privileges(0).pLuid = RestoreLuid
    TP.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    TP.Privileges(1).pLuid = BackupLuid
    TP.Privileges(1).Attributes = SE_PRIVILEGE_ENABLED
        
    Retval = AdjustTokenPrivileges(MyToken, vbFalse, TP, Len(TP), 0&, 0&)
'    If Retval = 0 Then MsgBox "AdjustTokenPrivileges: " & Err.LastDllError

End Sub
