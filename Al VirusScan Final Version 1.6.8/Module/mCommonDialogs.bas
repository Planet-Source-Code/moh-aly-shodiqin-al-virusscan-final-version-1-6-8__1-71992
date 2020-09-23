Attribute VB_Name = "mCommonDialogs"
' 20 Januari 2009
' 2:51 AM
'=======================================
' Module Common Dialogs
'=======================================
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Const MAX_PATH = 260

' Const Flags the dialog box
Public Const BIF_NEWDIALOGSTYLE As Long = &H40
Public Const BIF_EDITBOX As Long = &H10
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_STATUSTEXT = &H4&
Public Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Public Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Public Const BIF_VALIDATE As Long = &H20
Public Const BIF_SHAREABLE As Long = &H8000

Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As SAVEFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As SAVEFILENAME) As Long

Private Type SAVEFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String
    On Error GoTo BrowseFolderErr
    Dim intNull As Integer
    Dim lngIDList As Long
    Dim lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo
    
    With udtBI
        .hwndOwner = lngHwnd
        .lpszTitle = lstrcat(strPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_EDITBOX
    End With
    lngIDList = SHBrowseForFolder(udtBI)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, 0)
        lngResult = SHGetPathFromIDList(lngIDList, strPath)
        Call CoTaskMemFree(lngIDList)
        intNull = InStr(strPath, vbNullChar)
        If intNull > 0 Then
            strPath = Left(strPath, intNull - 1)
        End If
    End If
    
    BrowseFolder = strPath
    Exit Function
    
BrowseFolderErr:
    BrowseFolder = Empty
End Function

Public Sub ShowProps(Filename As String, OwnerhWnd As Long)
    On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = Filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = App.hInstance
        .lpIDList = 0
    End With
    ShellExecuteEx SEI
End Sub

Function ShowSave(hWnd As Long, Optional extFile As String = "All files|*.*", Optional isOpen As Boolean = False, Optional nTitle As String = "") As String
    Dim OFName As SAVEFILENAME
    extFile = Replace(extFile, "|", Chr(0))
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = extFile
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = nPath(App.path)
    OFName.lpstrTitle = nTitle
    OFName.Flags = 0
    
    If isOpen = False Then
        If GetSaveFileName(OFName) Then
           ShowSave = Left(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr(0)) - 1)
        Else
           ShowSave = ""
        End If
    Else
        If GetOpenFileName(OFName) Then
           ShowSave = Left(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr(0)) - 1)
        Else
           ShowSave = ""
        End If
    End If
End Function

Function ShowSaveSample(hWnd As Long, Optional extFile As String = "Application|*.exe", Optional isOpen As Boolean = False, Optional nTitle As String = "") As String
    Dim OFName As SAVEFILENAME
    extFile = Replace(extFile, "|", Chr(0))
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = extFile
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = "C:"
    OFName.lpstrTitle = nTitle
    OFName.Flags = 0
    
    If isOpen = False Then
        If GetSaveFileName(OFName) Then
           ShowSaveSample = Left(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr(0)) - 1)
        Else
           ShowSaveSample = ""
        End If
    Else
        If GetOpenFileName(OFName) Then
           ShowSaveSample = Left(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr(0)) - 1)
        Else
           ShowSaveSample = ""
        End If
    End If
End Function

