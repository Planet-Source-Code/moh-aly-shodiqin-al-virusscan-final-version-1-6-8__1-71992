VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 20 Januari 2009
' 4:22 AM
'=======================================
' cFileSearch
'=======================================
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

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

Event onSearch(nFileName As String, nFileInfo As cFileInfo)
Dim varFileInfo As New cFileInfo
Dim StopScan    As Boolean
Dim PauseScan   As Boolean

Dim colTipeFile As Collection

Property Let SetFileType(value As Collection)
    Set colTipeFile = Nothing
    Set colTipeFile = value
End Property
    
Property Get SetFileType() As Collection
    Set SetFileType = colTipeFile
End Property

Private Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Private Function FindFilesAPI(path As String, cAllFiles As Boolean)
    On Error Resume Next
    Dim FileName    As String
    Dim DirName     As String
    Dim dirNames()  As String
    Dim nDir        As Integer
    Dim I           As Integer
    Dim hSearch     As Long
    Dim Cont        As Integer
    Dim wfd         As WIN32_FIND_DATA
    Dim cFind       As String
    If Right(path, 1) <> "\" Then path = path & "\"
    'Dim FileInfo As AVIGScnEng.CFileInfo
    'Dim m_cShortcut As AVIGScnEng.cShellLink
    
    If StopScan Then Exit Function
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", wfd)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
         DoEvents: If StopScan Then Exit Function
         If PauseScan Then
            While PauseScan
                DoEvents
            Wend
         End If
         DirName = StripNulls(wfd.cFileName)
        
        If (DirName <> ".") And (DirName <> "..") Then
            
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                nDir = nDir + 1
                With varFileInfo
                     .FileName = ""
                     .FilePath = path & DirName
                     .FileSize = 0
                     .FileTitle = ""
                     .FileAttr = GetFileAttributes(path & DirName)
                End With
                RaiseEvent onSearch(path & FileName, varFileInfo)
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, wfd) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    
    hSearch = FindFirstFile(path & "*.*", wfd)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            DoEvents: If StopScan Then Exit Function
            If PauseScan Then
               While PauseScan
                   DoEvents
               Wend
            End If
            
            FileName = StripNulls(wfd.cFileName)
            'SearchStr = Trim(LCase(SearchStr))
            If (FileName <> ".") And (FileName <> "..") Then
               If cAllFiles Then
                    With varFileInfo
                         .FileName = FileName
                         .FilePath = path
                         .FileSize = FileLen(path & FileName)
                         .FileTitle = file_getTitle(path & FileName)
                         .FileAttr = GetFileAttributes(path & FileName)
                    End With
      
                    RaiseEvent onSearch(path & FileName, varFileInfo)
               Else
                 cFind = LCase(Trim(file_getTitle(FileName)))
                 If Trim(cFind) <> "" Then
                    'If InStr(1, SearchStr, "." & cFind, vbTextCompare) > 0 Then
                    If isFileTypeExist(UCase(cFind)) Then
                       With varFileInfo
                            .FileName = FileName
                            .FilePath = path
                            .FileSize = FileLen(path & FileName)
                            .FileTitle = file_getTitle(path & FileName)
                            .FileAttr = GetFileAttributes(path & FileName)
                       End With
'
                      RaiseEvent onSearch(path & FileName, varFileInfo)
                    End If
                 End If
                 cFind = ""
               End If
            End If
            Cont = FindNextFile(hSearch, wfd) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    
    If nDir > 0 Then
        
        For I = 0 To nDir - 1
            DoEvents: If StopScan Then Exit Function
            If PauseScan Then
               While PauseScan
                   DoEvents
               Wend
            End If
            Call FindFilesAPI(path & dirNames(I) & "\", cAllFiles)
        Next I
    End If
End Function

Sub DoCmdSearchFile(SearchPath As String, Optional AllFiles As Boolean = True)
    On Error Resume Next
    StopScan = False
    PauseScan = False
    FindFilesAPI SearchPath, AllFiles
    StopScan = True
End Sub

Property Let StopSearch(value As Boolean)
    StopScan = value
    If value Then PauseScan = False
End Property
Property Get StopSearch() As Boolean
    StopSearch = StopScan
End Property

Property Let PauseSearch(value As Boolean)
    PauseScan = value
End Property
Property Get PauseSearch() As Boolean
    PauseSearch = PauseScan
End Property

Private Sub Class_Terminate()
    StopScan = True
End Sub

Private Function isFileTypeExist(ID As String) As Boolean
    On Error GoTo salah
    Dim buff As String
    buff = colTipeFile("#" & ID)
    isFileTypeExist = True
    Exit Function
salah:
End Function
