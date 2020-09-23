VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShellLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Icon from picture:
Private Type PICTDESC
   cbSizeOfStruct As Long
   hGdiObj As Long
   hPalOrXYExt As Long
   picType As Long
End Type
Private Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type
Private Declare Sub OleCreatePictureIndirect Lib "olepro32.dll" ( _
    lpPictDesc As PICTDESC, riid As UUID, _
    ByVal fPictureOwnsHandle As Long, ipic As IPicture)
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

' IPersistFile:
Private Const STGM_DIRECT = 0

' ShowWindow constants:
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_MAXIMIZE = 3

' File:
Private Const MAX_PATH = 260
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpstring As String) As Long
' Win32 API Error Reporting:
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF


' Show command constants (a subset of constants for SetWindowPos API)
Public Enum EDisplayMode
    edmNormal = SW_NORMAL
    edmMinimized = SW_SHOWMINNOACTIVE
    edmMaximized = SW_MAXIMIZE
End Enum

Private link As New IShellLinkVB.cShellLink

'' Properties

'' Path of file represented by shortcut
Public Property Get TargetFile() As String
    Dim fd As WIN32_FIND_DATA, s As String
    s = String$(MAX_PATH, 0)
    link.GetPath s, MAX_PATH, fd, SLGP_UNCPRIORITY
    TargetFile = StrZToStr(s)
End Property

Public Property Let TargetFile(sPathA As String)
   ' Make sure file exists
   If Not FileExists(sPathA) Then
      Err.Raise 53
   Else
      link.SetPath sPathA
   End If
End Property

'' Startup directory for shortcut target
Public Property Get WorkingDirectory() As String
    Dim s As String
    s = String$(MAX_PATH, 0)
    link.GetWorkingDirectory s, MAX_PATH
    WorkingDirectory = StrZToStr(s)
End Property

Public Property Let WorkingDirectory(sWorkingA As String)
    link.SetWorkingDirectory sWorkingA
End Property

' Shortcut dialog ignores description, but we can save and restore it
Public Property Get Description() As String
    Dim s As String
    s = String$(MAX_PATH, 0)
    link.GetDescription s, MAX_PATH
    Description = StrZToStr(s)
End Property

Public Property Let Description(sDescription As String)
    link.SetDescription sDescription
End Property

'' Arguments for shortcut target
Public Property Get Arguments() As String
    Dim s As String
    s = String$(MAX_PATH, 0)
    link.GetArguments s, MAX_PATH
    Arguments = StrZToStr(s)
End Property

Public Property Let Arguments(sArgumentsA As String)
    link.SetArguments sArgumentsA
End Property

'' Display command can be Normal, Minimized, or Maximized
Public Property Get DisplayMode() As EDisplayMode
    DisplayMode = link.ShowCmd
End Property

Public Property Let DisplayMode(edm As EDisplayMode)
    Select Case edm
    Case SW_NORMAL
        ' Convert all these to normal: 0, 1, 4, 5, 8, 9, 10
        edm = edmNormal
    Case SW_SHOWMINNOACTIVE
        ' Convert all these to minimized: 2, 6, 7
        edm = edmMinimized
    Case SW_MAXIMIZE
        ' Pass maximize through: 3
        edm = edmMaximized
    Case Else
        ' Convert anything else to normal
        edm = edmNormal
    End Select
    link.ShowCmd = edm
End Property

Public Property Get HotKey() As KeyCodeConstants
   HotKey = link.HotKey
End Property

Public Property Let HotKey(kcc As KeyCodeConstants)
    link.HotKey = kcc
End Property

Public Property Get Icon() As IPicture
    Dim s As String, i As Long, hIcon As Long
    s = String$(MAX_PATH, 0)
    link.GetIconLocation s, MAX_PATH, i
    hIcon = ExtractIcon(App.hInstance, s, i)
    Set Icon = IconToPicture(hIcon)
End Property

Public Property Get IconFile() As String
   Dim s As String, i As Long, hIcon As Long
   s = String$(MAX_PATH, 0)
   link.GetIconLocation s, MAX_PATH, i
   s = StrZToStr(s)
   IconFile = s
End Property
Public Property Let IconFile(ByVal sFile As String)
Dim iIcon As Long
Dim s As String
   s = String$(MAX_PATH, 0)
   link.GetIconLocation s, MAX_PATH, iIcon
   link.SetIconLocation sFile, iIcon
End Property
Public Property Get IconIndex() As Long
Dim iIcon As Long
Dim s As String
   s = String$(MAX_PATH, 0)
   link.GetIconLocation s, MAX_PATH, iIcon
   IconIndex = iIcon
End Property
Public Property Let IconIndex(ByVal lIconIndex As Long)
Dim iIcon As Long
Dim s As String
   s = String$(MAX_PATH, 0)
   link.GetIconLocation s, MAX_PATH, iIcon
   link.SetIconLocation s, lIconIndex
End Property

Public Function Save(sLinkFile As String) As String
   Dim sLink As String

   
   ' Save the object to disk
   IPF(link).Save sLinkFile, 1
   Save = sLink
   
End Function

' Flags control behavior if LNK file reference can't be resolved:
'    SLR_ANY_MATCH - Display a dialog (with hWnd parameter as parent
'                    window) asking user whether to search for reference
'    SLR_NO_UI     - Search the disk for the time period specified by
'                    TimeOut parameter

Public Function Resolve(sFileA As String, _
            Optional flags As EShellLinkResolveFlags = SLR_ANY_MATCH, _
            Optional hWnd As Long = 0, _
            Optional TimeOut As Integer = 0) As Boolean
On Error GoTo Salah
   IPF(link).Load sFileA, STGM_DIRECT
   
   If flags = SLR_NO_UI And TimeOut > 0 Then
      Dim lDW As Long
      lDW = lDW * &H10000
      If (lDW And &H80000000) Then
         lDW = CLng(lDW And &H7FFFFFFF) Or &H80000000
      End If
      flags = flags Or lDW
   End If
   link.Resolve hWnd, flags Or SLR_NO_UI
   Resolve = True
   Exit Function
Salah:
End Function

Private Property Get IPF(link As IShellLinkVB.cShellLink) As IPersistFileVB
   Set IPF = link
End Property


Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, lstrlen(s))
End Function

Private Function FileExists(sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    FileExists = (Err = 0)
End Function

Private Function IconToPicture(ByVal hIcon As Long) As IPicture
   If hIcon = 0 Then
      Exit Function
   Else
      Dim ipic As IPicture
      Dim picdes As PICTDESC
      Dim iGuid As UUID
      
      ' Fill picture description
      picdes.cbSizeOfStruct = Len(picdes)
      picdes.picType = vbPicTypeIcon
      picdes.hGdiObj = hIcon
    
      ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
      With iGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
      End With
    
      ' Create picture from icon handle
      OleCreatePictureIndirect picdes, iGuid, True, ipic
      ' Result will be valid Picture or Nothing--either way set it
      Set IconToPicture = ipic
   End If
End Function


Private Function GetFullPath( _
      sFileName As String _
   ) As String

Dim c As Long, p As Long, sRet As String
    
   ' Get the path size, then create string of that size
   sRet = String(MAX_PATH, 0)
   c = GetFullPathName(sFileName, MAX_PATH, sRet, p)
   If c = 0 Then
      'Err.Raise WinApiError(Err.LastDllError)
   Else
      sRet = Left$(sRet, c)
      GetFullPath = sRet
   End If
    
End Function







