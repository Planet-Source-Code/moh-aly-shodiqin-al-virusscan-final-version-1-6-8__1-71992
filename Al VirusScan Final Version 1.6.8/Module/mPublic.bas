Attribute VB_Name = "mPublic"
Option Explicit

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Const ANYSIZE_ARRAY = 1

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As ESHGetFileInfoFlagConstants) As Long

Private Type SHFILEINFO
    hIcon           As Long ' : icon
    iIcon           As Long ' : icondex
    dwAttributes    As Long ' : SFGAO_ flags
    szDisplayName   As String * MAX_PATH ' : display name (or path)
    szTypeName      As String * 80 ' : type name
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As FILE_ATTRIBUTE) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Enum FILE_ATTRIBUTE
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_DIRECTORY = &H10
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_TEMPORARY = &H100
    FILE_ATTRIBUTE_COMPRESSED = &H800
End Enum

Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function BeepAPI Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const VER_PLATFORM_WIN32_NT = 2
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetUSERNAME Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Enum FOF_FILES
    FOF_NOCONFIRMATION = &H10
    FOF_SILENT = &H4
    FOF_ALLOWUNDO = &H40
End Enum
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Enum opFileWindows
    FO_MOVE = &H1
    FO_COPY = &H2
    FO_DELETE = &H3
    FO_RENAME = &H4
End Enum

Public SoundBuffer() As Byte
Public var_ClassID As Boolean

Public isStop As Boolean
Public UserCom As String
Public PCName As String
Public ComName As Long
Public GetIco As New cGetIconFile
Public REG As New cRegistry

Public Const vAppVersion = "1.6.8"
Public Const Copyright = "Copyright © 2008-2009 Moh Aly Shodiqin"
Public Const vScanEngine = "v1.3"
Public Const vScanWithVirusSample = "v1.4"
Public Const vProcessManager = "v1.6"
Public Const vAutorunLocation = "v1.1"
Public Const vRegistryTweak = "v1.5"
Public Const vRealtimeProtection = "v1.1"
Public vVirusDefinitions As String

Sub GetDefinitionDate(ByRef data As Collection)
    On Error Resume Next
    Dim iDB As Integer, iRC As Integer
    Dim rc As New ADODB.Recordset
    Dim mydata(1) As String
    Dim hSQL As String
    For iDB = 0 To UBound(myDatabase)
        If myDatabase(iDB).state = adStateOpen Then
            hSQL = "SELECT virus_definitions_date.autonum, virus_definitions_date.vdf_date From virus_definitions_date ORDER BY virus_definitions_date.vdf_date;"

            rc.Open hSQL, myDatabase(iDB), 3, 3
            If Not rc.EOF Then
               While Not rc.EOF
                  mydata(0) = NotNull(rc("vdf_date"))
                  data.Add mydata, NotNull(rc("autonum"))
                 rc.MoveNext
               Wend
            End If
            rc.Close
        End If
    Next iDB
End Sub

Sub VDFDate()
    Dim data As New Collection
    GetDefinitionDate data
    Dim i As Long, j As String

    If data.count > 0 Then
       For i = 1 To data.count
            vVirusDefinitions = data(i)(0)
       Next i
    End If
End Sub

Function FileDie(nFileName As String) As Boolean
    On Error GoTo Salah
    SetAttr nFileName, vbArchive + vbNormal
    Kill nFileName
    FileDie = True
    Exit Function
Salah:
End Function

Function TempWindow() As String
    Dim buff As String
    buff = String(255, 0)
    GetTempPath 255, buff
    TempWindow = nPath(Left(buff, InStr(1, buff, Chr(0)) - 1))
End Function

Function MyWindowDir() As String
    Dim buff As String
    buff = String(255, 0)
    GetWindowsDirectory buff, 255
    MyWindowDir = nPath(Left(buff, InStr(1, buff, Chr(0)) - 1))
End Function

Function MyWindowSys() As String
    Dim buff As String
    buff = String(255, 0)
    GetSystemDirectory buff, 255
    MyWindowSys = nPath(Left(buff, InStr(1, buff, Chr(0)) - 1))
End Function

Function nPath(mypath As String) As String
    If Right(mypath, 1) = "\" Then
       nPath = mypath
    Else
       nPath = mypath & "\"
    End If
End Function

Function file_getTitle(Filename As String) As String
    Dim Buffer() As String
    If InStr(1, Filename, ".", vbTextCompare) > 0 Then
       Buffer = Split(Filename, ".")
       If UBound(Buffer) > 0 Then
          file_getTitle = Buffer(UBound(Buffer))
       End If
    End If
End Function

Function file_getTitleName(Filename As String) As String
    Dim Buffer() As String
    If InStr(1, Filename, ".", vbTextCompare) > 0 Then
       Buffer = Split(Filename, ".")
       If UBound(Buffer) > 0 Then
          If UBound(Buffer) = 1 Then
            file_getTitleName = Buffer(0)
          Else
            ReDim Preserve Buffer(UBound(Buffer) - 1) As String
            file_getTitleName = Join(Buffer, ".")
          End If
       End If
    End If
End Function

Function file_getPath(Filename As String) As String
    On Error Resume Next
    Dim buff() As String
    buff() = Split(Filename, "\")
    If UBound(buff) > 0 Then
       ReDim Preserve buff(UBound(buff) - 1) As String
       file_getPath = nPath(Join(buff, "\"))
    End If
End Function

Function file_getName(Filename As String) As String
    On Error Resume Next
    Dim buff() As String
    buff() = Split(Filename, "\")
    If UBound(buff) > 0 Then
       file_getName = buff(UBound(buff))
    End If
End Function

Function file_getType(ByVal path As String) As String
    Dim FileInfo As SHFILEINFO, lngRet As Long
    
    lngRet = SHGetFileInfo(path, 0, FileInfo, Len(FileInfo), SHGFI_TYPENAME)
    If lngRet = 0 Then file_getType = Trim$(GetFileExtension(path) & " File"): Exit Function
    file_getType = Left$(FileInfo.szTypeName, InStr(1, FileInfo.szTypeName, vbNullChar) - 1)
End Function

Function GetFileExtension(ByVal path As String) As String
    Dim intRet As Integer
    intRet = InStrRev(path, ".")
    
    If intRet = 0 Then Exit Function
    GetFileExtension = UCase(Mid$(path, intRet + 1))
End Function

Function ReplacePathSystem(np As String) As String
    On Error Resume Next
    Dim buff As String
    buff = Replace(np, "\??\", "", , , vbTextCompare)
    buff = Replace(buff, "\\?\", "", , , vbTextCompare)
    buff = Replace(buff, "\SystemRoot\", MyWindowDir, , , vbTextCompare)
    buff = Replace(buff, "%systemroot%", MyWindowDir, , , vbTextCompare)
    buff = Replace(buff, "\\", "\", , , vbTextCompare)
    ReplacePathSystem = buff
End Function

Function GetMySetting(nApp As String, nKey As String, Optional nDefault As String)
    On Error GoTo Salah
    Dim ret As String
    ret = String(255, 0)
    GetPrivateProfileString nApp, nKey, nDefault, ret, 255, nPath(App.path) & "config.ini"
    GetMySetting = Left(ret, InStr(1, ret, Chr(0), vbTextCompare) - 1)
    Exit Function
Salah:
End Function

Function SetMySetting(nApp As String, nKey As String, nVal As String)
    On Error GoTo Salah
    WritePrivateProfileString nApp, nKey, nVal, nPath(App.path) & "config.ini"
    Exit Function
Salah:
End Function

Public Function IsFolderExist(Alamat As String) As Boolean
    If PathIsDirectory(Alamat) <> 0 Then
        IsFolderExist = True
    Else
        IsFolderExist = False
    End If
End Function

Sub QuarantineShow()
    If IsFolderExist(nPath(App.path) & "Quarantine") = False Then
        MkDir nPath(App.path) & "Quarantine"
    End If
End Sub

Function KillProcessById(p_lngProcessId As Long) As Long
On Error Resume Next
  Dim lnghProcess As Long
  Dim lngReturn As Long
    
    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES

    If IsWinNT Then
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY, hToken) = 0 Then
            CloseHandle hToken
        End If
        If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
            CloseHandle hToken
        End If
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
        If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, _
           ByVal 0&) = 0 Then
            CloseHandle hToken
        End If
    End If
    
    lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
    lngReturn = TerminateProcess(lnghProcess, 0&)
    KillProcessById = lngReturn
End Function

Public Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Sub Beep32(dwFreq As Long, dwDuration As Long)
    On Error Resume Next
    Dim H As String
    BeepAPI dwFreq, dwDuration
End Sub

Function GetTipeDrive(ndrv As String) As String
    GetTipeDrive = GetDriveType(ndrv)
    '    Select Case GetDriveType(ndrv)
    '        Case 2
    '            GetTipeDrive = "Removable"
    '        Case 3
    '            GetTipeDrive = "Drive Fixed"
    '        Case Is = 4
    '            GetTipeDrive = "Remote"
    '        Case Is = 5
    '            GetTipeDrive = "Cd-Rom"
    '        Case Is = 6
    '            GetTipeDrive = "Ram disk"
    '        Case Else
    '            GetTipeDrive = "Unrecognized"
    '    End Select
End Function

Function GetSpecialfolder(CSIDL As Long) As String
On Error Resume Next
    Dim r As Long, path As String
    Dim IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = 0 Then
        path = SPACE$(512)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal path)
        GetSpecialfolder = Left$(path, InStr(path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Public Function NameOfTheComputer(MachineName As String) As Long
    Dim NameSize As Long
    Dim x As Long
    MachineName = SPACE$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
End Function

Public Function GetUserCom() As String
    GetUserCom = Environ$("username")
    ComName = NameOfTheComputer(PCName)
    With frmConsole
        .Caption = "al VirusScan Console - " & PCName
'        .sbConsole.Panels(1).Text = Copyright
'        .sbConsole.Panels(2).Text = Format(Date, "ddd, dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") 'PCName
    End With
End Function

Function file_isFolder(path As String) As Long
    On Error GoTo Salah
    Dim ret As VbFileAttribute
    ret = GetAttr(path) And vbDirectory
    If ret = vbDirectory Then
        file_isFolder = 1
    Else
        file_isFolder = 0
    End If
    Exit Function
Salah:
    file_isFolder = -1
End Function

Sub CopyCOMCTL()
    On Error Resume Next
    Dim H As String
    H = Dir(nPath(App.path) & "COMCTL32.OCX", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If H = "" Then
        FileCopy nPath(MyWindowSys) & "COMCTL32.OCX", nPath(App.path) & "MSVBVM60.DLL"
    Else
        H = Dir(nPath(MyWindowSys) & "COMCTL32.OCX", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
        If H = "" Then
            FileCopy nPath(App.path) & "COMCTL32.OCX", nPath(MyWindowSys) & "COMCTL32.OCX"
        End If
    End If
End Sub

Public Function ShellWinFile(Operat As opFileWindows, nFlags As FOF_FILES, ParamArray vntFileName() As Variant)
    On Error Resume Next
    Dim i As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT
    
    For i = LBound(vntFileName) To UBound(vntFileName)
        sFileNames = sFileNames & vntFileName(i) & vbNullChar
    Next
    sFileNames = sFileNames & vbNullChar
    
    With SHFileOp
        .wFunc = Operat
        .pFrom = sFileNames
        .fFlags = nFlags Or FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    
    ShellWinFile = SHFileOperation(SHFileOp)
End Function

Public Sub AlwaysOnTop(hWnd As Long, SetOnTop As Boolean)
    If SetOnTop Then
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPFLAGS
    Else
        SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPFLAGS
    End If
End Sub

Public Sub SetOpagueForm(lMode As Boolean, F As Form)
    F.Enabled = False
    Dim i As Integer
    Select Case lMode
        Case False
            i = 255
            Do
                i = i - 5
                DoEvents
                MakeTransparent F.hWnd, i
                Sleep 10
            Loop While i >= 180
        Case True
            i = 192
            Do
                i = i + 5
                DoEvents
                MakeTransparent F.hWnd, i
                Sleep 10
            Loop Until i >= 255
    End Select
    F.Enabled = True
'    Me.Refresh
End Sub

Public Function MakeTransparent(hWnd As Long, Perc As Integer) As Long
    Dim msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
        MakeTransparent = 1
    Else
        msg = GetWindowLong(hWnd, GWL_EXSTYLE)
        msg = msg Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, msg
        SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
        MakeTransparent = 0
    End If
    If Err Then
        MakeTransparent = 2
    End If
End Function

Public Sub VirusAlert()
    On Error Resume Next
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "SoundWarning", 1) = 1 Then
        SoundBuffer = LoadResData("W2", "WAV")
        sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
    Else
        SoundBuffer = LoadResData("", "")
        sndPlaySound ByVal 0&, SND_NODEFAULT
    End If
End Sub

Public Sub AutoSizeListView(ByVal lv As ListView, Optional Column As ColumnHeader = Nothing)
    On Error Resume Next
    Dim c As ColumnHeader
    If Column Is Nothing Then
        For Each c In lv.ColumnHeaders
            SendMessage lv.hWnd, LVM_FIRST + 30, c.Index - 1, -1
        Next
    Else
        SendMessage lv.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    lv.Refresh
End Sub

Public Sub LogScan(sLog As String)
    On Error Resume Next
    Dim FF As Integer
    FF = FreeFile
    MkDir App.path & "\Log"
    Open App.path & "\Log\" & "VirusScanLog" & ".txt" For Append As #FF
        Print #FF, Date & vbTab & Time & "  " & sLog
    Close #FF
End Sub

Private Sub CL_AddItem(ByRef sCLParam() As String, ByVal sNewString As String, Optional sQuote As String = """")
On Error Resume Next
If sCLParam(0) <> vbNullString Then
    ReDim Preserve sCLParam(UBound(sCLParam) + 1) As String
End If
sCLParam(UBound(sCLParam)) = sNewString
If Left(sCLParam(UBound(sCLParam)), 1) = sQuote And Right(sCLParam(UBound(sCLParam)), 1) = sQuote Then
    sCLParam(UBound(sCLParam)) = Mid(sCLParam(UBound(sCLParam)), 2, Len(sCLParam(UBound(sCLParam))) - 2)
End If
End Sub

Private Function CL_CountChar(ByVal sInput As String, ByVal sChar As String) As Long
On Error Resume Next
Dim iPos As Long
CL_CountChar = 0
iPos = 1
Do Until InStr(iPos, sInput, sChar) = 0
    CL_CountChar = CL_CountChar + 1
    iPos = InStr(iPos, sInput, sChar) + Len(sChar)
Loop
End Function
Public Sub CL_Get(ByRef sCLParam() As String, nParam As String, Optional sSpace As String = ",", Optional sQuote As String = """")
On Error Resume Next
Dim sTemp As String
Dim iCounter As Long, iPos(1) As Long
ReDim sCLParam(0) As String
sTemp = nParam

If CL_CountChar(sTemp, sQuote) Mod 2 = 1 Then
    Exit Sub
Else
    iPos(0) = 1
    iPos(1) = 1
    Do Until InStr(iPos(0), sTemp, sSpace) = 0
        iPos(1) = InStr(iPos(0), sTemp, sSpace)
        If CL_CountChar(Mid(sTemp, iPos(0), iPos(1) - iPos(0)), sQuote) Mod 2 = 1 Then
            iPos(1) = InStr(iPos(1), sTemp, sQuote) + Len(sQuote)
        End If
        If iPos(1) > 0 Then
            CL_AddItem sCLParam, Mid(sTemp, iPos(0), iPos(1) - iPos(0)), sQuote
            iPos(0) = iPos(1) + 1
        Else
            Stop
        End If
    Loop
    If iPos(0) <= Len(sTemp) Then
        iPos(1) = Len(sTemp) + 1
        CL_AddItem sCLParam, Mid(sTemp, iPos(0), iPos(1) - iPos(0)), sQuote
    End If
End If
End Sub

Function GetFileNameFromParam(c As String) As String
Dim Z() As String, Y() As String
Z() = Split(c, "\")

If InStr(1, Z(UBound(Z)), " ", vbTextCompare) Then
   Y() = Split(Z(UBound(Z)), " ")
   ReDim Preserve Z(UBound(Z) - 1) As String
   GetFileNameFromParam = nPath(Join(Z, "\")) & Y(UBound(Y) - 1)
Else
  GetFileNameFromParam = c
End If
End Function

Public Function GetChecksum(sFile As String) As String
    On Error Resume Next
    Dim cb0 As Byte
    Dim cb1 As Byte
    Dim cb2 As Byte
    Dim cb3 As Byte
    Dim cb4 As Byte
    Dim cb5 As Byte
    Dim cb6 As Byte
    Dim cb7 As Byte
    Dim cb8 As Byte
    Dim cb9 As Byte
    Dim cb10 As Byte
    Dim cb11 As Byte
    Dim cb12 As Byte
    Dim cb13 As Byte
    Dim cb14 As Byte
    Dim cb15 As Byte
    Dim cb16 As Byte
    Dim cb17 As Byte
    Dim cb18 As Byte
    Dim cb19 As Byte
    Dim cb20 As Byte
    Dim cb21 As Byte
    Dim cb22 As Byte
    Dim cb23 As Byte
    Dim buff As String
    
    Open sFile For Binary Access Read As #1
        buff = SPACE$(1)
        Get #1, , buff
    Close #1
    
    Open sFile For Binary Access Read As #2
        Get #2, 512, cb0
        Get #2, 1024, cb1
        Get #2, 2048, cb2
        Get #2, 3000, cb3
        Get #2, 4096, cb4
        Get #2, 5000, cb5
        Get #2, 6000, cb6
        Get #2, 7000, cb7
        Get #2, 8192, cb8
        Get #2, 9000, cb9
        Get #2, 10000, cb10
        Get #2, 11000, cb11
        Get #2, 12288, cb12
        Get #2, 13000, cb13
        Get #2, 14000, cb14
        Get #2, 15000, cb15
        Get #2, 16384, cb16
        Get #2, 17000, cb17
        Get #2, 18000, cb18
        Get #2, 19000, cb19
        Get #2, 20480, cb20
        Get #2, 21000, cb21
        Get #2, 22000, cb22
        Get #2, 23000, cb23
    Close #2
    buff = cb0
    buff = buff & cb1
    buff = buff & cb2
    buff = buff & cb3
    buff = buff & cb4
    buff = buff & cb5
    buff = buff & cb6
    buff = buff & cb7
    buff = buff & cb8
    buff = buff & cb9
    buff = buff & cb10
    buff = buff & cb11
    buff = buff & cb12
    buff = buff & cb13
    buff = buff & cb14
    buff = buff & cb15
    buff = buff & cb16
    buff = buff & cb17
    buff = buff & cb18
    buff = buff & cb19
    buff = buff & cb20
    buff = buff & cb21
    buff = buff & cb22
    buff = buff & cb23
    Dim m_CRC As New cCRC32
    GetChecksum = m_CRC.StringChecksum(buff)
    Set m_CRC = Nothing
    Exit Function
ErrHandle:
    Close #2
End Function

' Advanced ProgressBar
Public Sub GetRGB(ByVal Collective As Long, Red As Long, Green As Long, Blue As Long)
    If Collective < 0 Then Collective = RGB(105, 105, 255) 'System color replacer
    Dim x As Long
    x = Int(Collective / 65536)
    Blue = x
    Collective = Collective - x * 65536
    x = Int(Collective / 256)
    Green = x
    Red = Collective - x * 256
End Sub
