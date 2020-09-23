VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindFromPTR2 Lib "AVIGptr.dll" (PARAM1 As Any, ByVal PARAM2 As Any, PARAM3 As Any, ByVal PARAM4 As Any) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFileEx Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long, lpBaseMemPointer As Any) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseMemPointer As Any) As Long
Private Declare Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const MOVEFILE_REPLACE_EXISTING = &H1

Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_NORMAL = &H80

Const FILE_BEGIN = 0

Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FILE_SHARE_NONE = &H0
Const FILE_SHARE_READWRITE = &H3

Const CREATE_NEW = &H1 'Kreiraj novi bez brisanja starog
Const CREATE_ALWAYS = &H2 'Kreiraj novi preko postojeceg
Const OPEN_EXISTING = &H3 'Otvori postojeci,ako nema izbaci gresku
Const OPEN_ALWAYS = &H4 'Otvori postojeci,ako nema kreiraj novi
Const APPEND_EXISTING = &H5 'Piši na kraj filea

Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const GENERIC_READWRITE = &HC0000000
Const GENERIC_ALL = &H10000000

'Pristup Mapiranom file-u
Const PageNoAccess = &H1
Const PageReadOnly = &H2
Const PageReadWrite = &H4
Const PageWrite = &H8
Const PageExecute = &H10
Const PageExecuteRead = &H20
Const PageExecuteReadWrite = &H40
Const PageExecuteWrite = &H80
Const PageGuard = &H100
Const PageNoCache = &H200

'Pristup MAPI PODATAKA
Const FileMapCopy = &H1
Const FileMapQuery = &H1
Const FileMapWrite = &H2
Const FileMapRead = &H4
Const FileMapReadWrite = &H6
Const FileMapExtendedReadWrite = &H7
Const FileMapAll = &HF001F

Private PFileName As String
Private FileObjectName As String
Private FileLength As Long
Private FileHandle As Long 'HANDLE za otvoreni file
Private MapFileHandle As Long 'HANDLE za mapirani file
Private AccessMemory As Long 'Memorija pristup podacima kroz COPYMEMORY

Private Function OpenPageFile(ByVal Filename As String, Optional size As Long = 0, Optional Value As Byte = 0) As Boolean
On Error GoTo Salah
Dim tmpsz As Long
 ClosePageFile
 FileHandle = CreateFile(Filename, GENERIC_READWRITE, FILE_SHARE_READWRITE, ByVal 0&, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
 FileObjectName = Str(FileHandle)
 FileLength = GetFileSize(FileHandle, 0)
 tmpsz = FileLength
 FileLength = FileLength + size
 MapFileHandle = CreateFileMapping(FileHandle, ByVal 0&, PageReadWrite, 0, FileLength, FileObjectName)

 AccessMemory = MapViewOfFile(MapFileHandle, FileMapReadWrite, 0, 0, 0)
 If size <> 0 Then
    FillMemory ByVal AccessMemory + tmpsz, size, Value
 End If
 PFileName = Filename
 OpenPageFile = True
 Exit Function
Salah:
    OpenPageFile = False
End Function

Public Function ClosePageFile()
On Error Resume Next
    FlushViewOfFile ByVal AccessMemory, FileLength
    UnmapViewOfFile ByVal AccessMemory
    CloseHandle MapFileHandle
    CloseHandle FileHandle
    PFileName = ""
    FileObjectName = ""
    FileLength = 0
    FileHandle = 0
    MapFileHandle = 0
    AccessMemory = 0
End Function

Private Sub Class_Terminate()
    ClosePageFile
End Sub

Public Function SearchInOneFile(ByVal Filename As String, ByVal Pattern As String) As Long
On Error GoTo Dalje2
    If Len(Filename) = 0 Then Exit Function
    If Len(Pattern) > FileLen(Filename) Then
        SearchInOneFile = 0
        Exit Function
    End If
    If OpenPageFile(Filename) Then
      If AccessMemory = 0 Then
        SearchInOneFile = 0
ClosePageFile:
        Exit Function
      End If
      SearchInOneFile = SearchOneIn(Pattern)
      ClosePageFile
    End If
Exit Function
Dalje2:
    On Error GoTo 0
    SearchInOneFile = 0
    ClosePageFile
End Function

Private Function SearchOneIn(ByVal Pattern As String) As Long
On Error GoTo Salah
    Dim STARTADR As Long
    Dim RET      As Long
    Dim LASTADR  As Long
    
    STARTADR = AccessMemory
    LASTADR = STARTADR + FileLength
        
    RET = FindFromPTR2(ByVal STARTADR, LASTADR - STARTADR, ByVal Pattern, ByVal Len(Pattern))
    If RET = -1 Then
       SearchOneIn = 0
    Else
       SearchOneIn = 1
    End If
    Exit Function
Salah:
SearchOneIn = -1
End Function

Function isMeAvailabe() As Boolean
On Error Resume Next
Dim h As Long
Dim f As String
f = nPath(App.path) & "\check.ptr"
Open f For Output As #1
    Print #1, "vbbego community checked"
Close #1
h = SearchInOneFile(f, "community")
If h = 1 Then
   isMeAvailabe = True
Else
   isMeAvailabe = False
End If
Kill f
End Function

