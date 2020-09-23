Attribute VB_Name = "mAPI"
Option Explicit

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Const ANYSIZE_ARRAY = 1

Private Type LUID
    lowpart As Long
    highpart As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Type TOKEN_PRIVILEGES32
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Declare Function AdjustTokenPrivileges32 Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES32, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES32, ReturnLength As Long) As Long
Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

' Opens the access token associated with a thread ------------------
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadID As Long) As Long
Public Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean

Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000
Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16

Public Const THREAD_SUSPEND_RESUME As Long = &H2
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const PROCESS_SET_INFORMATION As Long = (&H200)

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

' -------------------------------------------
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (SEI As SHELLEXECUTEINFO) As Long

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    Hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

' Registry access
' mdlRegistry -----------------------------------
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const KEY_ALL_ACCESS As Long = &H3F
Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const WM_SETTINGCHANGE As Long = &H1A
Public Const SPI_SETNONCLIENTMETRICS As Long = &H2A
Public Const SMTO_ABORTIFHUNG As Long = &H2

Public Type SECURITY_ATTRIBUTES
   nLength                 As Long
   lpSecurityDescriptor    As Long
   bInheritHandle          As Long
End Type

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
    
Public Enum TipeReg
    REG_SZ = 1                         ' Unicode nul terminated string
    reg_binarY = 3                     ' Free form binary
    REG_DWORD = 4                      ' 32-bit number
End Enum

Public Const ERROR_SUCCESS = 0&

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPFLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000

Public Declare Sub SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long

' Set the priority class
' for the specified process -----------------------------
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long

'Public Enum PriorityClass
'   REALTIME_PRIORITY_CLASS = &H100
'   HIGH_PRIORITY_CLASS = &H80
'   NORMAL_PRIORITY_CLASS = &H20
'   IDLE_PRIORITY_CLASS = &H40
'End Enum

Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(256) As Byte
End Type

Public Enum IconRetrieve
    ricnLarge = 32
    ricnSmall = 16
End Enum

'User-Defined API Enum
Public Enum ESHGetFileInfoFlagConstants
    SHGFI_ATTRIBUTES = &H800        'get file attributes
    SHGFI_DISPLAYNAME = &H200       'get display name
    SHGFI_EXETYPE = &H2000          'get exe type
    SHGFI_ICON = &H100              'get icon handle and index
    SHGFI_LARGEICON = &H0           'get file's large icon
    SHGFI_LINKOVERLAY = &H8000      'add link overlay on the icon
    SHGFI_OPENICON = &H2            'get file's open icon
    SHGFI_SELECTED = &H10000        'blend icon with the system highlight color
    SHGFI_SHELLICONSIZE = &H4       'get shell-sized icon
    SHGFI_SMALLICON = &H1           'get file's small icon
    SHGFI_SYSICONINDEX = &H4000     'get icon index from system image list
    SHGFI_TYPENAME = &H400          'get file type description
    SHGFI_USEFILEATTRIBUTES = &H10  'use dwFileAttributes parameter
End Enum

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

