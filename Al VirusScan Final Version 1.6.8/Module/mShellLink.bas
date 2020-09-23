Attribute VB_Name = "mShellLink"
Option Explicit

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
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
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Function FileExists(ByVal sFile As String) As Boolean
    Dim tFnd As WIN32_FIND_DATA
    Dim hSearch As Long

    hSearch = FindFirstFile(sFile, tFnd)
    If Not (hSearch = -1) Then
        FindClose hSearch
        FileExists = True
    End If
End Function



