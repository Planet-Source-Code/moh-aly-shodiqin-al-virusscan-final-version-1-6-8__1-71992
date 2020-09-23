Attribute VB_Name = "mShutDown"
'Module         : Shut Down Windows
'Date/Time      : 10 April 2009 11:01 AM
'----------------------------------------
Option Explicit

Private Const EWX_FORCE As Long = 4
Private Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Public Enum ExitWindows
    LOGOFF = 0
    SHUTDOWN = 1
    REBOOT = 2
    POWEROFF = 8
End Enum

#If False Then
Private LOGOFF, SHUTDOWN, REBOOT, POWEROFF
#End If

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
  
Private Sub AdjustToken()
    Dim hdlProcessHandle          As Long
    Dim hdlTokenHandle            As Long
    Dim tmpLuid                   As LUID
    Dim tkp                       As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored          As TOKEN_PRIVILEGES
    Dim lBufferNeeded             As Long

    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (&H20 Or &H8), hdlTokenHandle
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    With tkp
        .PrivilegeCount = 1
        .TheLuid = tmpLuid
        .Attributes = &H2
    End With
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub

Public Sub KillWindows(ByVal aOption As ExitWindows)
    AdjustToken
    Select Case aOption
      Case ExitWindows.LOGOFF
        ExitWindowsEx (ExitWindows.LOGOFF Or EWX_FORCE), &HFFFF
      Case ExitWindows.REBOOT
        ExitWindowsEx (ExitWindows.SHUTDOWN Or EWX_FORCE Or ExitWindows.REBOOT), &HFFFF
      Case ExitWindows.SHUTDOWN
        ExitWindowsEx (ExitWindows.SHUTDOWN Or EWX_FORCE), &HFFFF
      Case ExitWindows.POWEROFF
        ExitWindowsEx (ExitWindows.POWEROFF Or EWX_FORCE), &HFFFF
    End Select
End Sub
