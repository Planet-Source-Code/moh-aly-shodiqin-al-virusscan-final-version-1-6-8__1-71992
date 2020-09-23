Attribute VB_Name = "Module1"
'Mainly filehandling routines here
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Public Const GWL_STYLE = (-16)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const HDS_BUTTONS As Long = &H2
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Public Const SWP_DRAWFRAME As Long = &H20
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_FLAGS As Long = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Public fMainForm As frmMain
Public RootDummy As Node
Public Const OFN_OVERWRITEPROMPT = &H2
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngCC As Long
End Type

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function GetDosPath(LongPath As String) As String
    Dim s As String
    Dim i As Long
    Dim PathLength As Long
    i = Len(LongPath) + 1
    s = String(i, 0)
    PathLength = GetShortPathName(LongPath, s, i)
    GetDosPath = Left$(s, PathLength)

End Function
Sub Main()
    On Error Resume Next
    
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngCC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    
    On Error GoTo 0
    Set fMainForm = New frmMain
    fMainForm.Show
    fMainForm.cboAddress.ListIndex = 0
End Sub
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Function GetTempPathName() As String
    Dim sBuffer As String
    Dim lRet As Long
    sBuffer = String$(255, vbNullChar)
    lRet = GetTempPath(255, sBuffer)
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    GetTempPathName = sBuffer
End Function
Public Function GetTempFile(lpTempFilename As String, Optional mDir As String) As Boolean
    lpTempFilename = String(255, vbNullChar)
    If mDir = "" Then mDir = GetTempPathName
    GetTempFile = GetTempFileName(mDir, "bb", 0, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function

Public Function GetWinDir() As String
    Dim strSave As String
    strSave = String(200, Chr$(0))
    GetWinDir = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave)))
End Function
Public Function PathOnly(ByVal FilePath As String) As String
Dim temp As String
    temp = Mid$(FilePath, 1, InStrRev(FilePath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function


Public Function FileOnly(ByVal FilePath As String) As String
    FileOnly = Mid$(FilePath, InStrRev(FilePath, "\") + 1)
End Function
Public Sub FileAppend(Text As String, FilePath As String)
On Error Resume Next
Dim f As Integer
f = FreeFile
Dim Directory As String
              Directory$ = FilePath
    Open Directory$ For Append As #f
        Print #f, Text
    Close #f
Exit Sub
End Sub

Public Function HexToDec(ByVal HexStr As String) As Double
    'Borrowed from PSC
    Dim mult As Double
    Dim DecNum As Double
    Dim ch As String
    mult = 1
    DecNum = 0
    Dim i As Integer
    For i = Len(HexStr) To 1 Step -1
        ch = Mid(HexStr, i, 1)
        If (ch >= "0") And (ch <= "9") Then
            DecNum = DecNum + (Val(ch) * mult)
        Else
            If (ch >= "A") And (ch <= "F") Then
                DecNum = DecNum + ((Asc(ch) - Asc("A") + 10) * mult)
            Else
                If (ch >= "a") And (ch <= "f") Then
                    DecNum = DecNum + ((Asc(ch) - Asc("a") + 10) * mult)
                Else
                    HexToDec = 0
                    Exit Function
                End If
            End If
        End If
        mult = mult * 16
    Next i
    HexToDec = DecNum
End Function
Public Function FileExists(sSource As String) As Boolean
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function


Public Function GetRootHandle(mRoot As String) As Long
    'convert string to long constant
    Select Case mRoot
        Case "HKEY_CLASSES_ROOT"
            GetRootHandle = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            GetRootHandle = HKEY_CURRENT_CONFIG
        Case "HKEY_CURRENT_USER"
            GetRootHandle = HKEY_CURRENT_USER
        Case "HKEY_DYN_DATA"
            GetRootHandle = HKEY_DYN_DATA
        Case "HKEY_LOCAL_MACHINE"
            GetRootHandle = HKEY_LOCAL_MACHINE
        Case "HKEY_PERFORMANCE_DATA"
            GetRootHandle = HKEY_PERFORMANCE_DATA
        Case "HKEY_USERS"
            GetRootHandle = HKEY_USERS
        Case Else
            GetRootHandle = 0
    End Select
End Function
Public Function GetRootText(mRoot As Long) As String
    'convert long constant to string
    Select Case mRoot
        Case HKEY_CLASSES_ROOT
            GetRootText = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_CONFIG
            GetRootText = "HKEY_CURRENT_CONFIG"
        Case HKEY_CURRENT_USER
            GetRootText = "HKEY_CURRENT_USER"
        Case HKEY_DYN_DATA
            GetRootText = "HKEY_DYN_DATA"
        Case HKEY_LOCAL_MACHINE
            GetRootText = "HKEY_LOCAL_MACHINE"
        Case HKEY_PERFORMANCE_DATA
            GetRootText = "HKEY_PERFORMANCE_DATA"
        Case HKEY_USERS
            GetRootText = "HKEY_USERS"
    End Select
End Function

