Attribute VB_Name = "mCommonFunction"
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal mWnd As Long, ByVal aWnd As Long, data As String, parms As String, show As Boolean, nopause As Boolean) As Long

Public ViriOnCollect As New Collection
Public OnSelectDlg   As VbMsgBoxResult
Public DrvOnCollect  As Collection
Public WindCollect   As Collection

Function CallDLL(strLibraryName As String, functionName As String)
    Dim lb As Long, pa As Long
    'map 'user32' into the address space of the calling process.
    lb = LoadLibrary(strLibraryName)
    pa = GetProcAddress(lb, functionName)
    'CallWindowProc pa, Me.hWnd,
End Function

'Function FileExists(strFileName As String) As Boolean
'    On Error GoTo MakeF
'    'If file does Not exist, there will be an Error
'    Open strFileName For Input As #1
'    Close #1
'    'no error, file exists
'    FileExists = True
'    Exit Function
'MakeF:
'    'error, file does Not exist
'    FileExists = False
'    Exit Function
'
'End Function

Function JoinArray(thearray() As String, strDelim As String, start As Integer, Optional endx As Integer = -1) As String
    If endx = -1 Then endx = UBound(thearray) + 1
    Dim i As Integer, result As String
    
    For i = start - 1 To endx - 1
        If isStop = True Then Exit For
        If i = endx - 1 Then
            result = result & thearray(i)
        Else
            result = result & thearray(i) & strDelim
        End If
    Next i
    JoinArray = result
End Function

Function JoinArrayV(thearray(), strDelim As String, start As Integer, Optional endx As Integer = -1) As String
    If endx = -1 Then endx = UBound(thearray) + 1
    Dim i As Integer, result As String
    
    For i = start - 1 To endx - 1
        If isStop = True Then Exit For
        If i = endx - 1 Then
            result = result & thearray(i)
        Else
            result = result & thearray(i) & strDelim
        End If
    Next i
    JoinArrayV = result
End Function

Function TrimLeft(strText As String) As String
    Dim i As Integer
    For i = 1 To Len(strText)
        If isStop = True Then Exit For
        If Mid(strText, i, 1) <> " " And Mid(strText, i, 1) <> Chr(9) Then
            TrimLeft = Right(strText, Len(strText) - (i - 1))
            Exit Function
        End If
    Next i
End Function

Sub ExtractResource(nID, nType, nFileName As String)
    On Error Resume Next
    Kill nFileName
    Dim Buffer() As Byte
    Buffer() = LoadResData(nID, nType)
    Open nFileName For Binary As #1
        Put #1, , Buffer
    Close #1
End Sub

Public Sub DownloadFile(ByVal srcFileName As String, ByVal targetFileName As String)
  'This Downloads the latest version from the Internet
    Dim B() As Byte
    Dim FID As Byte

    Call frmUpdate.DownStatus("Connecting...")
    B() = frmUpdate.Inet.OpenURL(srcFileName, icByteArray)
    FID = FreeFile
    Open targetFileName For Binary Access Write As #FID
        Put #FID, , B()
    Close #FID
    Call frmUpdate.DownStatus("Unable to update virus definitions...")
    frmUpdate.sbStatus.Panels(1).Text = "Invalid URL"
    DoEvents
End Sub
