VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEngine32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Event onVirusFound(nFileName As String, nFileInfo As cFileInfo)
Event onEngineError(nstr As String)
Dim ColSol As Collection
Dim PFpage As cPaging
Dim PFPageUse As Boolean

Private Sub Class_Initialize()
On Error Resume Next
Dim hFile As String
Dim I As Integer
Dim filecol As New Collection

Set PFpage = New cPaging
PFPageUse = PFpage.isMeAvailabe

hFile = Dir(nPath(App.path) & "\scan.vdf", vbArchive Or vbNormal Or vbReadOnly Or vbSystem Or vbHidden)
If hFile <> "" Then
   While hFile <> ""
        I = I + 1
        filecol.Add nPath(App.path) & "\" & hFile
        hFile = Dir()
        DoEvents
   Wend
End If
If I > 0 Then
    ReDim myDatabase(I - 1) As New ADODB.Connection
    For I = 1 To filecol.count
        If LoadDatabase(myDatabase(I - 1), filecol(I)) = False Then
           RaiseEvent onEngineError("Can't load virus definitions file")
        End If
        Sleep 100
        DoEvents
    Next I
Else
   ReDim myDatabase(0) As New ADODB.Connection
   MsgBox "Can't load virus definition file, make sure you have correct files", 48, "AVS Error" & Chr(0)
End If

Set filecol = Nothing
'LoadPlugin

On Error Resume Next
'MkDir nPath(App.path) & "Quarantines"
'MkDir nPath(App.path) & "virdef"
End Sub

Function KillFile(Filename As String) As Boolean
    On Error GoTo Salah
    Kill Filename
    KillFile = True
    Exit Function
Salah:
    KillFile = False
End Function

Function CekOneFile(Filename As String, Optional hIDProcess As Long = 0, Optional ByRef var_viriName As String) As Boolean
    If var_ClassID = False Then Exit Function
    On Error Resume Next
    
    If Trim(Filename) <> "" Then
        Dim H() As Byte, rc As New ADODB.Recordset
        Dim I As Integer
        Dim nInfo As New cFileInfo
        Dim hSQL As String
        
        For I = 0 To UBound(myDatabase)
            If myDatabase(I).State = adStateOpen Then
                If rc.State = adStateOpen Then rc.Close
                hSQL = "SELECT vdb_virus_collection_head.id, vdb_virus_collection_head.virus_name, vdb_virus_collection_head.systems_affected, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, vdb_virus_collection_detail.removal_script, vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.default_action, vdb_virus_collection_detail.virus_date, vdb_virus_collection_detail.virus_like0, vdb_virus_collection_detail.virus_like1, vdb_virus_collection_detail.virus_like2, vdb_virus_collection_detail.virus_like3, vdb_virus_collection_detail.virus_like4, vdb_virus_collection_detail.virus_like5, vdb_virus_collection_detail.virus_like6, vdb_virus_collection_detail.virus_like7, vdb_virus_collection_detail.virus_like8, vdb_virus_collection_detail.virus_like9,vdb_virus_collection_detail.id_virus " & _
                        "FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id  " & _
                        "Where (((vdb_virus_collection_detail.virus_crc_check) = '" & GetSignatureFromFile(Filename) & "')) ORDER BY vdb_virus_collection_head.id;"
    
                rc.Open hSQL, myDatabase(I), 3, 3
                If rc.EOF Then
                    rc.Close
                    hSQL = "SELECT vdb_virus_collection_head.id, vdb_virus_collection_head.virus_name, vdb_virus_collection_head.systems_affected, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, vdb_virus_collection_detail.removal_script, vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.default_action, vdb_virus_collection_detail.virus_date, vdb_virus_collection_detail.virus_like0, vdb_virus_collection_detail.virus_like1, vdb_virus_collection_detail.virus_like2, vdb_virus_collection_detail.virus_like3, vdb_virus_collection_detail.virus_like4, vdb_virus_collection_detail.virus_like5, vdb_virus_collection_detail.virus_like6, vdb_virus_collection_detail.virus_like7, vdb_virus_collection_detail.virus_like8, vdb_virus_collection_detail.virus_like9,vdb_virus_collection_detail.id_virus " & _
                            "FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id  " & _
                            "Where (((vdb_virus_collection_detail.virus_crc_check2) = '" & GetSignatureFromFile(Filename, , 10000) & "')) ORDER BY vdb_virus_collection_head.id;"

                    rc.Open hSQL, myDatabase(I), 3, 3
                End If
                
                If Not rc.EOF Then
kembali:
                With nInfo
                    .Filename = Filename
                    If hIDProcess <> 0 Then
                       KillProcessById hIDProcess
                       KillProcessById hIDProcess
                       Sleep 500
                    End If
                    
                    QuarantineShow
                    Select Case UCase(NotNull(rc("default_action")))
                        Case "DELETE"
                             On Error Resume Next
                             RunningOneSolution NotNull(rc("id_virus")), Filename
                             SetFileAttributes Filename, FILE_ATTRIBUTE_NORMAL
                             .VirusClean = KillFile(Filename) ' IIf(ShellWinFile(FO_DELETE, FOF_ALLOWUNDO, Filename) = 0, True, False )
                        Case "QUARANTINE"
                             .VirusClean = QuarantineFile(NotNull(rc("virus_alias")), Filename)
                        Case "QUARANTINE-REMOVE"
                             .VirusClean = QuarantineFile(NotNull(rc("virus_alias")), Filename)
                             RunningOneSolution NotNull(rc("id_virus")), Filename
                        Case "BUNDLE"
                             RunningOneSolution NotNull(rc("id_virus")), Filename
                             .VirusClean = True
                    End Select
                    
                    var_viriName = NotNull(rc("virus_name"))
                    .VirusAction = UCase(NotNull(rc("default_action")))
                    .VirusAlias = NotNull(rc("virus_alias"))
                    .VirusCRC = NotNull(rc("virus_crc_check"))
                    .VirusID = NotNull(rc("id"))
                    .VirusIDDet = NotNull(rc("id_virus"))
                    .VirusName = NotNull(rc("virus_name"))
                    .VirusType = NotNull(rc("virus_type"))
                End With
                    RaiseEvent onVirusFound(Filename, nInfo)
                    CekOneFile = True
                    rc.Close
                    Exit For
            Else
                Select Case LCase(file_getTitle(Filename))
                    Case "cmd", "bat", "vbs", "htt", "inf", "bas", "html", "htm"
test_lagi:
                        'If ValidatePortableExecutable(filename) = False Then
                               
                        If FileLen(Filename) < 10000000 Then
                            If rc.State = adStateOpen Then rc.Close
                                'hSQL = "SELECT vdb_virus_collection_head.id, vdb_virus_collection_head.virus_name, vdb_virus_collection_head.systems_affected, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, vdb_virus_collection_detail.removal_script, vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.default_action, vdb_virus_collection_detail.virus_date, vdb_virus_collection_detail.virus_like0, vdb_virus_collection_detail.virus_like1, vdb_virus_collection_detail.virus_like2, vdb_virus_collection_detail.virus_like3, vdb_virus_collection_detail.virus_like4, vdb_virus_collection_detail.virus_like5, vdb_virus_collection_detail.virus_like6, vdb_virus_collection_detail.virus_like7, vdb_virus_collection_detail.virus_like8, vdb_virus_collection_detail.virus_like9, vdb_virus_collection_detail.id_virus " & _
                                       "FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id  " & _
                                       "ORDER BY vdb_virus_collection_head.id;"
                                       
                                       
                                hSQL = "SELECT vdb_virus_collection_head.id, vdb_virus_collection_head.virus_name, vdb_virus_collection_head.systems_affected, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, vdb_virus_collection_detail.removal_script, vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.default_action, vdb_virus_collection_detail.virus_date, vdb_virus_collection_detail.virus_like0, vdb_virus_collection_detail.virus_like1, vdb_virus_collection_detail.virus_like2, vdb_virus_collection_detail.virus_like3, vdb_virus_collection_detail.virus_like4, vdb_virus_collection_detail.virus_like5, vdb_virus_collection_detail.virus_like6, vdb_virus_collection_detail.virus_like7, vdb_virus_collection_detail.virus_like8, vdb_virus_collection_detail.virus_like9, vdb_virus_collection_detail.id_virus " & _
                                       " FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id " & _
                                       " WHERE (((vdb_virus_collection_detail.virus_like1)<>'') OR ((vdb_virus_collection_detail.virus_like2)<>'') OR ((vdb_virus_collection_detail.virus_like3)<>'') OR ((vdb_virus_collection_detail.virus_like4)<>'') OR ((vdb_virus_collection_detail.virus_like5)<>'') OR ((vdb_virus_collection_detail.virus_like6)<>'') OR ((vdb_virus_collection_detail.virus_like7)<>'') OR ((vdb_virus_collection_detail.virus_like8)<>'') OR ((vdb_virus_collection_detail.virus_like9)<>'')) " & _
                                       " ORDER BY vdb_virus_collection_head.id; "
        
                                rc.Open hSQL, myDatabase(I), 3, 3
                                       
                                    If Not rc.EOF Then
                                        Dim data(9) As String
                                        Dim nLoop As Integer
                                        While Not rc.EOF
                                            For nLoop = 0 To 9
                                                data(nLoop) = NotNull(rc("virus_like" & nLoop))
                                            Next nLoop
                                            
                                            If PFPageUse Then
                                                If NewisFileInfected(Filename, data) Then
                                                    CekOneFile = True
                                                    GoSub kembali
                                                    Exit Function
                                                End If
                                            Else
                                                If isFileInfected(Filename, data) Then
                                                    CekOneFile = True
                                                    GoSub kembali
                                                    Exit Function
                                                End If
                                            End If
                                            
                                            For nLoop = 0 To 9
                                                data(nLoop) = ""
                                            Next nLoop
                                            rc.MoveNext
                                        Wend
                                    End If
                                    rc.Close
                                End If
                        Case Else
                            CekOneFile = False
                End Select
                End If
              
            End If
        Next I
        Set rc = Nothing
        Set nInfo = Nothing
    End If
    DoEvents
End Function

Function GetSignatureFromFile(Filename As String, Optional StartOn As Long = 1000, Optional MAXSize As Long = 0) As String
    If var_ClassID = False Then Exit Function
    On Error GoTo Salah
    Dim mySize As Integer
    mySize = 16
    
    Dim Data0(15) As Byte
    Dim Data1(15) As Byte
    Dim Data2(15) As Byte
    Dim Data3(15) As Byte
    Dim Data4(15) As Byte
    Dim Data5(15) As Byte
    Dim Data6(15) As Byte
    Dim Data7(15) As Byte
    Dim Data8(15) As Byte
    Dim Data9(15) As Byte
    Dim Data10(15) As Byte
    Dim Data11(15) As Byte
    Dim Data12(15) As Byte
    Dim Data13(15) As Byte
    Dim Data14(15) As Byte
    Dim Data15(15) As Byte
    Dim Data16(15) As Byte
    Dim Data17(15) As Byte
    Dim Data18(15) As Byte
    Dim Data19(15) As Byte
    Dim Data20(15) As Byte
            
    
    Dim lenOfFile As Long
    lenOfFile = FileLen(Filename)
    If lenOfFile < StartOn Then StartOn = 1
    If lenOfFile > 100 Then
        Open Filename For Binary As #1
            If MAXSize <= 0 Then
                If lenOfFile >= mySize Then Get #1, 1, Data0
                If lenOfFile >= (mySize + 200) Then Get #1, 200, Data1
                If lenOfFile >= (mySize + 512) Then Get #1, 512, Data2
                If lenOfFile >= (mySize + 1000) Then Get #1, 1000, Data3
                If lenOfFile >= (mySize + 1024) Then Get #1, 1024, Data4
                If lenOfFile >= (mySize + 2048) Then Get #1, 2048, Data5
                If lenOfFile >= (mySize + 4096) Then Get #1, 4096, Data6
                If lenOfFile >= (mySize + (HasilKali(4096, 2))) Then Get #1, HasilKali(4096, 2), Data7
                If lenOfFile >= (mySize + (HasilKali(4096, 3))) Then Get #1, HasilKali(4096, 3), Data8
                If lenOfFile >= (mySize + (HasilKali(4096, 4))) Then Get #1, HasilKali(4096, 4), Data9
                If lenOfFile >= (mySize + (HasilKali(4096, 5))) Then Get #1, HasilKali(4096, 5), Data10
                If lenOfFile >= (mySize + (HasilKali(4096, 6))) Then Get #1, HasilKali(4096, 6), Data11
                If lenOfFile >= (mySize + (HasilKali(4096, 7))) Then Get #1, HasilKali(4096, 7), Data12
                If lenOfFile >= (mySize + (HasilKali(4096, 8))) Then Get #1, HasilKali(4096, 8), Data13
                If lenOfFile >= (mySize + (HasilKali(4096, 9))) Then Get #1, HasilKali(4096, 9), Data14
                If lenOfFile >= (mySize + (HasilKali(4096, 10))) Then Get #1, HasilKali(4096, 10), Data15
                If lenOfFile >= (mySize + (HasilKali(4096, 11))) Then Get #1, HasilKali(4096, 11), Data16
                If lenOfFile >= (mySize + (HasilKali(4096, 12))) Then Get #1, HasilKali(4096, 12), Data17
                If lenOfFile >= (mySize + (HasilKali(4096, 13))) Then Get #1, HasilKali(4096, 13), Data18
                If lenOfFile >= (mySize + (HasilKali(4096, 14))) Then Get #1, HasilKali(4096, 14), Data19
                If lenOfFile >= (mySize + (HasilKali(4096, 15))) Then Get #1, HasilKali(4096, 15), Data20
            Else
                If MAXSize >= mySize Then Get #1, 1, Data0
                If MAXSize >= 430 Then Get #1, 430, Data1
                If MAXSize >= 980 Then Get #1, 980, Data2
                If MAXSize >= 2000 Then Get #1, 2000, Data3
                If MAXSize >= 3000 Then Get #1, 3000, Data4
                If MAXSize >= 4000 Then Get #1, 4000, Data5
                If MAXSize >= 5000 Then Get #1, 5000, Data6
                If MAXSize >= 6000 Then Get #1, 6000, Data7
                If MAXSize >= 7000 Then Get #1, 7000, Data8
                If MAXSize >= 8000 Then Get #1, 8000, Data9
                If MAXSize >= 9000 Then Get #1, 9000, Data10
                If MAXSize >= 10000 Then Get #1, 10000, Data11
                If MAXSize >= 11000 Then Get #1, 11000, Data12
                If MAXSize >= 12000 Then Get #1, 12000, Data13
                If MAXSize >= 13000 Then Get #1, 13000, Data14
                If MAXSize >= 14000 Then Get #1, 14000, Data15
            End If
            Close #1
    End If
    
    Dim buff As String
    buff = StrConv(Data0, vbUnicode)
    buff = buff & StrConv(Data1, vbUnicode)
    buff = buff & StrConv(Data2, vbUnicode)
    buff = buff & StrConv(Data3, vbUnicode)
    buff = buff & StrConv(Data4, vbUnicode)
    buff = buff & StrConv(Data5, vbUnicode)
    buff = buff & StrConv(Data6, vbUnicode)
    buff = buff & StrConv(Data7, vbUnicode)
    buff = buff & StrConv(Data8, vbUnicode)
    buff = buff & StrConv(Data9, vbUnicode)
    buff = buff & StrConv(Data10, vbUnicode)
    buff = buff & StrConv(Data11, vbUnicode)
    buff = buff & StrConv(Data12, vbUnicode)
    buff = buff & StrConv(Data13, vbUnicode)
    buff = buff & StrConv(Data14, vbUnicode)
    buff = buff & StrConv(Data15, vbUnicode)
    buff = buff & StrConv(Data16, vbUnicode)
    buff = buff & StrConv(Data17, vbUnicode)
    buff = buff & StrConv(Data18, vbUnicode)
    buff = buff & StrConv(Data19, vbUnicode)
    buff = buff & StrConv(Data20, vbUnicode)
    
    Dim j As New cMD5
    j.MD5Init
    j.DigestStrToHexStr buff
    GetSignatureFromFile = j.GetValues
    Set j = Nothing
    Exit Function
Salah:
    RaiseEvent onEngineError("Can't get CRC from file " & Filename)
    Close #1
End Function

Function HasilKali(destA As Currency, desB As Currency) As Currency
    If var_ClassID = False Then Exit Function
    On Error GoTo Salah
    Dim desC As Currency
    desC = destA * desB
    HasilKali = desC
Salah:
End Function

Sub RunningSolution(col As Collection)
    On Error Resume Next
    If var_ClassID = False Then Exit Sub
    
    Set ColSol = New Collection
    Dim I As Long, cViri As Long
    Dim dataID(1) As Byte
    For I = 1 To col.count
        dataID(0) = "#" & col(I).VirusID
        dataID(1) = "#" & col(I).VirusIDDet
        AddToSolution dataID
    Next I
    
    If ColSol.count > 0 Then
    Dim rc As New ADODB.Recordset
    Dim myScript As New cScript
    Dim m As New cMain
    Dim paramlist(0)
    Dim hSQL As String
    
        For I = 0 To UBound(myDatabase)
            If myDatabase(I).State = adStateOpen Then
                For cViri = 1 To ColSol.count
                    If rc.State = adStateOpen Then rc.Close
                        hSQL = "SELECT *  FROM vdb_virus_collection_detail WHERE (id=" & Replace(ColSol(cViri)(0), "#", "") & ") AND (id_virus=" & Replace(ColSol(cViri)(1), "#", "") & ");"
                        rc.Open hSQL, myDatabase(I), 3, 3
              
                    If Not rc.EOF Then
                        Dim data As String
                        data = NotNull(rc("removal_script"))
                        Open TempWindow & "src_def.tmp" For Output As #1
                            Print #1, data
                        Close #1
                        m.NewScript
                        m.LoadScript m.ScriptCount, TempWindow & "src_def.tmp"
                        m.ExecuteAlias "main", paramlist()
                        Set m = Nothing
                        Exit For
                    End If
                    rc.Close
                Next cViri
            End If
        Next I
        Set rc = Nothing
    End If
    Set ColSol = Nothing
End Sub

Function AddToSolution(nVirusID() As Byte) As Boolean
    If var_ClassID = False Then Exit Function
    On Error GoTo Salah
    ColSol.Add nVirusID, nVirusID
    AddToSolution = True
    Exit Function
Salah:
End Function

Function QuarantineFile(virus As String, Filename As String) As Boolean
    If var_ClassID = False Then Exit Function
    On Error GoTo Salah
    Sleep 100
    Dim Length As Currency
    Length = FileLen(Filename)
    If Length > 10 Then
        SetFileAttributes Filename, FILE_ATTRIBUTE_NORMAL
       
        Dim Data1() As Byte
        Dim Data2() As Byte
           
        Dim first As Currency
        Dim second As Currency
          
        first = Int(Length / 2)
        second = (Length - first) - 2
        
        ReDim Data1(first) As Byte
        ReDim Data2(second) As Byte
        
        Open Filename For Binary As #1
             Get #1, , Data1
             Get #1, , Data2
        Close #1
        
        Kill Filename
             
        Dim OldName As String
        OldName = String(Len(Filename) + Len(virus) + 2, 0)
        OldName = virus & Chr(0) & Filename & Chr(0) & Chr(0)
        
        first = Len(OldName)
        second = (Length - first) + Len(OldName)
        
        Dim NewName As String
        QuarantineShow
        NewName = nPath(App.path) & "Quarantine\" & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & Int(Rnd * 255) & ".al"
        Open NewName For Binary Access Write As #1
             Put #1, , OldName
             Put #1, , Data2
             Put #1, , Data1
        Close #1
     
       ReDim Data1(0) As Byte
       ReDim Data2(0) As Byte
       QuarantineFile = True
    End If
        
    Exit Function
Salah:
    RaiseEvent onEngineError("Virus can't move to quarantine directory - " & Filename)
    Close #1
End Function

Function RestoreFiles(Filename As String, Optional Hwnd As Long = 0) As Boolean
    If var_ClassID = False Then Exit Function
    On Error GoTo Salah
        Dim Length As Long
        Length = FileLen(Filename)
        If Length > 1 Then
            Dim Data1() As Byte
            Dim Data2() As Byte
            Dim OldName As String

            Dim first As Currency
            Dim second As Currency
            Dim mark   As String
            Dim NewName As String

            mark = String(1024, 0)

            Open Filename For Binary Access Read As #1
                 Get #1, , mark
            Close #1
            mark = Left(mark, InStr(1, mark, Chr(0) & Chr(0)) - 1)
            Dim nInfo() As String
            nInfo() = Split(mark, Chr(0))

            If Trim(mark) <> "" Then
                Length = Length - (Len(mark) + 2)
                If isFileExists(nInfo(1)) = False Then
kembali:
                MakeSureDirectoryPathExists file_getPath(nInfo(1))
                first = Fix(Length / 2)
                second = (Length - first) - 2

                ReDim Data1(first) As Byte
                ReDim Data2(second) As Byte
                Open Filename For Binary Access Read As #1
                     Get #1, Len(mark) + 3, Data2
                     Get #1, second + Len(mark) + 4, Data1
                Close #1

                'ReDim Preserve Data2(1 To UBound(Data2)) As Byte

                If NewName <> "" Then mark = NewName

                Open nInfo(1) For Binary Access Write As #2
                      Put #2, , Data1
                      Put #2, , Data2
                Close #2

                RestoreFiles = True
                SetFileAttributes Filename, FILE_ATTRIBUTE_NORMAL
                Kill Filename
                ReDim Data1(0) As Byte
                ReDim Data2(0) As Byte
                mark = ""
'            Else
'                Dim msg As VbMsgBoxResult
'                msg = MsgBox(Replace(mark, Chr(0), "") & vbCrLf & vbCrLf & "File aready exist, YES button to replace, " & vbCrLf & "NO button to create new filename, CANCEL to abort", vbYesNoCancel + 32, "Confirm•" & Chr(0))
'                If msg = vbYes Then
'                    GoSub kembali
'                ElseIf msg = vbNo Then
'                    NewName = ShowSave(hwnd, "All files|*.*")
'                    If NewName <> "" Then
'                        GoSub kembali
'                        Return
'                    End If
'                Else
'                End If
                End If
           End If
        End If
        Exit Function
Salah:
    RaiseEvent onEngineError("Restore virus to original directory failed! - " & Filename)
    Close #1
    Close #2
End Function

Function isFileExists(Filename As String) As Boolean
    On Error Resume Next
    If Len(Filename) = 0 Then Exit Function
    If Dir(Filename, vbNormal + vbArchive + vbHidden + vbReadOnly + vbSystem) <> "" Then isFileExists = True
End Function

Sub GetVirusListInfo(ByRef data As Collection)
    On Error Resume Next
    If var_ClassID = False Then Exit Sub
    Dim iDB As Integer, iRC As Integer
    Dim rc As New ADODB.Recordset
    Dim mydata(1) As String
    Dim hSQL As String
    For iDB = 0 To UBound(myDatabase)
        If myDatabase(iDB).State = adStateOpen Then
            hSQL = "SELECT vdb_virus_collection_head.id, vdb_virus_collection_head.virus_name,vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, * " & _
                    "FROM vdb_virus_collection_head LEFT JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id ORDER BY vdb_virus_collection_detail.virus_alias;"
     
            rc.Open hSQL, myDatabase(iDB), 3, 3
            If Not rc.EOF Then
               While Not rc.EOF
                  mydata(0) = NotNull(rc("virus_alias"))
                  mydata(1) = NotNull(rc("virus_type"))
                  data.Add mydata, NotNull(rc("id_virus"))
                 rc.MoveNext
               Wend
            End If
            rc.Close
        End If
    Next iDB
End Sub

'Sub GetDefinitionDate(ByRef data As Collection)
'    On Error Resume Next
'    Dim iDB As Integer, iRC As Integer
'    Dim rc As New ADODB.Recordset
'    Dim mydata(1) As String
'    Dim hSQL As String
'    For iDB = 0 To UBound(myDatabase)
'        If myDatabase(iDB).state = adStateOpen Then
'            hSQL = "SELECT virus_definitions_date.autonum, virus_definitions_date.vdf_time, virus_definitions_date.vdf_date From virus_definitions_date ORDER BY virus_definitions_date.vdf_date;"
'
'            rc.Open hSQL, myDatabase(iDB), 3, 3
'            If Not rc.EOF Then
'               While Not rc.EOF
'                  mydata(0) = NotNull(rc("vdf_date"))
''                  mydata(1) = NotNull(rc("virus_type"))
'                  data.Add mydata, NotNull(rc("autonum"))
'                 rc.MoveNext
'               Wend
'            End If
'            rc.Close
'        End If
'    Next iDB
'End Sub

Function FindVirusOnly(Filename As String) As Boolean
    On Error Resume Next
    If var_ClassID = False Then Exit Function
    If Trim(Filename) <> "" Then
        Dim rc As New ADODB.Recordset
        Dim I As Integer
        Dim hSQL As String
    
        For I = 0 To UBound(myDatabase)
            If myDatabase(I).State = adStateOpen Then
                If rc.State = adStateOpen Then rc.Close
                hSQL = "SELECT vdb_virus_collection_head.id, vdb_virus_collection_head.virus_name, vdb_virus_collection_head.systems_affected, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, vdb_virus_collection_detail.removal_script, vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.default_action, vdb_virus_collection_detail.virus_date, vdb_virus_collection_detail.virus_like0, vdb_virus_collection_detail.virus_like1, vdb_virus_collection_detail.virus_like2, vdb_virus_collection_detail.virus_like3, vdb_virus_collection_detail.virus_like4, vdb_virus_collection_detail.virus_like5, vdb_virus_collection_detail.virus_like6, vdb_virus_collection_detail.virus_like7, vdb_virus_collection_detail.virus_like8, vdb_virus_collection_detail.virus_like9 " & _
                        "FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id  " & _
                        "Where (((vdb_virus_collection_detail.virus_crc_check) = '" & GetSignatureFromFile(Filename) & "')) ORDER BY vdb_virus_collection_head.id;"
    
                rc.Open hSQL, myDatabase(I), 3, 3
    
                If rc.EOF Then
                    rc.Close
                    hSQL = "SELECT vdb_virus_collection_head.id, vdb_virus_collection_head.virus_name, vdb_virus_collection_head.systems_affected, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, vdb_virus_collection_detail.removal_script, vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.default_action, vdb_virus_collection_detail.virus_date, vdb_virus_collection_detail.virus_like0, vdb_virus_collection_detail.virus_like1, vdb_virus_collection_detail.virus_like2, vdb_virus_collection_detail.virus_like3, vdb_virus_collection_detail.virus_like4, vdb_virus_collection_detail.virus_like5, vdb_virus_collection_detail.virus_like6, vdb_virus_collection_detail.virus_like7, vdb_virus_collection_detail.virus_like8, vdb_virus_collection_detail.virus_like9,vdb_virus_collection_detail.id_virus " & _
                            "FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id  " & _
                            "Where (((vdb_virus_collection_detail.virus_crc_check2) = '" & GetSignatureFromFile(Filename, , 10000) & "')) ORDER BY vdb_virus_collection_head.id;"
            
                    rc.Open hSQL, myDatabase(I), 3, 3
              End If
    
            If Not rc.EOF Then
                FindVirusOnly = True
                rc.Close
                Exit For
                  Else
                    Select Case LCase(file_getTitle(Filename))
                        Case "cmd", "bat", "vbs", "htm", "html", "htt", "asp", "jsp", "js", "php", "xml", "xsl", "bas", "mht", "hta", "pl", "reg", "dat", "inf", "tmp", "bak"
test_lagi:
                                
                            If rc.State = adStateOpen Then rc.Close
                                rc.Open "SELECT vdb_virus_collection_head.id, vdb_virus_collection_head.virus_name, vdb_virus_collection_head.systems_affected, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, vdb_virus_collection_detail.removal_script, vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.default_action, vdb_virus_collection_detail.virus_date, vdb_virus_collection_detail.virus_like0, vdb_virus_collection_detail.virus_like1, vdb_virus_collection_detail.virus_like2, vdb_virus_collection_detail.virus_like3, vdb_virus_collection_detail.virus_like4, vdb_virus_collection_detail.virus_like5, vdb_virus_collection_detail.virus_like6, vdb_virus_collection_detail.virus_like7, vdb_virus_collection_detail.virus_like8, vdb_virus_collection_detail.virus_like9, vdb_virus_collection_detail.id_virus " & _
                                    " FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id " & _
                                    " WHERE (((vdb_virus_collection_detail.virus_like1)<>'') OR ((vdb_virus_collection_detail.virus_like2)<>'') OR ((vdb_virus_collection_detail.virus_like3)<>'') OR ((vdb_virus_collection_detail.virus_like4)<>'') OR ((vdb_virus_collection_detail.virus_like5)<>'') OR ((vdb_virus_collection_detail.virus_like6)<>'') OR ((vdb_virus_collection_detail.virus_like7)<>'') OR ((vdb_virus_collection_detail.virus_like8)<>'') OR ((vdb_virus_collection_detail.virus_like9)<>'')) " & _
                                    " ORDER BY vdb_virus_collection_head.id; ", myDatabase(I), 3, 3
                                
                                If Not rc.EOF Then
                                   Dim data(9) As String
                                   Dim nLoop As Integer
                                   While Not rc.EOF
                                        For nLoop = 0 To 9
                                            data(nLoop) = NotNull(rc("virus_like" & nLoop))
                                        Next nLoop
                                        
                                        If PFPageUse Then
                                            If NewisFileInfected(Filename, data) Then
                                               FindVirusOnly = True
                                               rc.Close
                                               Set rc = Nothing
                                               Exit Function
                                            End If
                                        Else
                                            If isFileInfected(Filename, data) Then
                                               FindVirusOnly = True
                                               rc.Close
                                               Set rc = Nothing
                                               Exit Function
                                            End If
                                        End If
                                        For nLoop = 0 To 9
                                            data(nLoop) = ""
                                        Next nLoop
                                        rc.MoveNext
                                   Wend
                                End If
                                rc.Close
                        Case Else
                            FindVirusOnly = False
                End Select
              End If
           End If
       Next I
       Set rc = Nothing
    End If
    DoEvents
End Function

Function NewisFileInfected(Filename As String, nFind() As String) As Boolean
    If var_ClassID = False Then Exit Function
    On Error GoTo Salah
    'On Error Resume Next
    If FileLen(Filename) = 0 Then Exit Function
    Dim cc() As Boolean
    Dim tx() As String
    
    Dim I As Integer, jml As Integer
    
    
    For I = 0 To UBound(nFind)
        If Trim(nFind(I)) <> "" Then
           jml = jml + 1
        End If
    Next
    
    If jml = 0 Then Exit Function
    
    ReDim cc(jml) As Boolean
    ReDim tx(jml) As String
    
    For I = 0 To UBound(nFind)
        If Trim(nFind(I)) <> "" Then
           tx(I) = nFind(I)
        End If
    Next
        
           
    Dim myRes As Boolean, oncek As Long
    myRes = False
    
    For I = 0 To UBound(cc)
        If cc(I) = False Then
           If Trim(tx(I)) = "" Then
             cc(I) = False
           Else
            oncek = PFpage.SearchInOneFile(Filename, tx(I))
            If oncek = 1 Then
               cc(I) = True
            Else
               cc(I) = False
            End If
           End If
        End If
        If cc(0) = False Then
           myRes = True
           Exit For
        End If
    Next I
            
    Dim hRes As Boolean
    hRes = cc(0)
    For I = 1 To UBound(cc)
        hRes = hRes And cc(I)
        If cc(I) = False Then hRes = False: Exit For
    Next I
    
    NewisFileInfected = hRes
    Exit Function
Salah:
Close #1

End Function

Function isFileInfected(Filename As String, nFind() As String) As Boolean
    If var_ClassID = False Then Exit Function
    On Error GoTo Salah
'On Error Resume Next
    If FileLen(Filename) = 0 Then Exit Function
    Dim data As String
    If FileLen(Filename) > 100000 Then
       data = String(100000, 0)
    Else
       data = String(FileLen(Filename), 0)
    End If
    Dim cc() As Boolean
    Dim tx() As String

    Dim I As Integer, jml As Integer


    For I = 0 To UBound(nFind)
        If Trim(nFind(I)) <> "" Then
           jml = jml + 1
        End If
    Next

    If jml = 0 Then Exit Function

    ReDim cc(jml) As Boolean
    ReDim tx(jml) As String
    'If jml = 3 Then Stop
    For I = 0 To UBound(nFind)
        If Trim(nFind(I)) <> "" Then
           tx(I) = nFind(I)
        End If
    Next


    Dim myRes As Boolean
    myRes = False

    Open Filename For Binary Access Read As #1
    Do While Not EOF(1) And myRes = False
        Get #1, , data ' Read next record.
        'Line Input #1, Data
        For I = 0 To UBound(cc)
            If cc(I) = False Then If InStr(1, data, tx(I), vbTextCompare) > 0 Then cc(I) = True
            If cc(0) = False Then
               myRes = True
               Exit For
            End If
        Next I

        DoEvents
    Loop
    Close #1

    data = ""

    Dim hRes As Boolean
    hRes = cc(0)
    For I = 1 To UBound(cc)
        hRes = hRes And cc(I)
        If cc(I) = False Then hRes = False: Exit For
    Next I

    isFileInfected = hRes
    Exit Function
Salah:
Close #1

End Function

Function ContainsText(sFileName, sText) As Boolean
    Dim iFile As Integer
    Dim sRecordLine As String
    iFile = FreeFile
    On Error Resume Next
    
    ContainsText = False
    
    Open sFileName For Input As iFile
        Do While Not EOF(iFile)
            Line Input #iFile, sRecordLine
                If CBool(InStr(1, sRecordLine, sText, vbTextCompare)) Then
                        ContainsText = True
                        Exit Do
                End If
                DoEvents
        Loop
    Close iFile
End Function

Property Let ClassIDApartement(Value As String)
    On Error Resume Next
    If Value = Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(1) & Chr(255) Then
       var_ClassID = True
    Else
       var_ClassID = False
    End If
End Property

Sub CloseScanHandle()
    On Error Resume Next
    If var_ClassID = False Then Exit Sub
    Dim I As Integer
    For I = 0 To UBound(myDatabase)
       myDatabase(I).Close
    Next I
End Sub

Function GetViriListByUser(ByRef col As Collection) As Boolean
    If var_ClassID = False Then Exit Function
    Dim rc As New ADODB.Recordset
    On Error GoTo Salah
    
    Dim I As Integer
    For I = 0 To UBound(myDatabase)
       If myDatabase(I).State = adStateOpen Then
          Dim ID As String
          
          rc.Open "SELECT vdb_virus_collection_head.id, vdb_virus_collection_detail.virus_crc_check,vdb_virus_collection_detail.id_virus, vdb_virus_collection_head.virus_name, vdb_virus_collection_head.systems_affected, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type " & _
                  "FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id " & _
                  "WHERE (((vdb_virus_collection_head.virus_name)='ADD_BY_USER')) " & _
                  "ORDER BY vdb_virus_collection_detail.virus_alias;", myDatabase(I), 3, 3
    
          Dim data(3) As String
          If Not rc.EOF Then
             While Not rc.EOF
                 data(0) = NotNull(rc("id"))
                 data(1) = NotNull(rc("id_virus"))
                 data(2) = NotNull(rc("virus_alias"))
                 data(3) = NotNull(rc("virus_type"))
                 col.Add data, NotNull(rc("virus_crc_check"))
                rc.MoveNext
             Wend
          End If
          rc.Close
       End If
    Next I
    Set rc = Nothing
    GetViriListByUser = True
    Exit Function
Salah:
End Function

Function DeleteViriIDByUser(ID1 As String, ID2 As String) As Boolean
    If var_ClassID = False Then Exit Function
    Dim rc As New ADODB.Recordset
    On Error GoTo Salah
        Dim I As Integer
        For I = 0 To UBound(myDatabase)
            If myDatabase(I).State = 1 Then
                rc.Open "SELECT id, id_virus From vdb_virus_collection_detail " & _
                        "WHERE (vdb_virus_collection_detail.id=" & ID1 & ") AND (vdb_virus_collection_detail.id_virus=" & ID2 & ");", myDatabase(I), 3, 3
                        
                If Not rc.EOF Then
                   myDatabase(I).Execute "DELETE FROM vdb_virus_collection_detail WHERE (vdb_virus_collection_detail.id=" & ID1 & ") AND (vdb_virus_collection_detail.id_virus=" & ID2 & ");"
                   DeleteViriIDByUser = True
                End If
                rc.Close
             End If
        Next I
        Set rc = Nothing
     Exit Function
Salah:
    Set rc = Nothing
End Function

Sub RunningOneSolution(ID As String, Filename As String)
    On Error Resume Next
    If var_ClassID = False Then Exit Sub
    
    Dim rc As New ADODB.Recordset
    Dim myScript As New cScript
    Dim m As New cMain
    Dim paramlist(0)
    Dim hSQL As String
    
    Dim I As Integer
    For I = 0 To UBound(myDatabase)
        If myDatabase(I).State = adStateOpen Then
            hSQL = "SELECT *  FROM vdb_virus_collection_detail WHERE (id_virus=" & ID & ");"
            rc.Open hSQL, myDatabase(I), 3, 3
            If Not rc.EOF Then
                Dim data As String, myname As String
                myname = Filename
                myname = Replace(myname, "\", "\\", , , vbTextCompare)
                myname = Replace(myname, "$", "\$", , , vbTextCompare)
                myname = Replace(myname, "#", "\#", , , vbTextCompare)
                myname = Replace(myname, "%", "\%", , , vbTextCompare)
                myname = Replace(myname, "`", "\`", , , vbTextCompare)
                myname = Replace(myname, "~", "\~", , , vbTextCompare)
                myname = Replace(myname, "[", "\[", , , vbTextCompare)
                myname = Replace(myname, "]", "\]", , , vbTextCompare)
                
                data = NotNull(rc("removal_script"))
                data = Replace(data, "%filename%", myname, , , vbTextCompare)
                Open TempWindow & "src_def.tmp" For Output As #1
                    Print #1, data
                Close #1
'                Tunggu 0.2
                m.NewScript
                m.LoadScript m.ScriptCount, TempWindow & "src_def.tmp"
                m.ExecuteAlias "main", paramlist()
                Set m = Nothing
            End If
            rc.Close
        End If
    Next I
    Set rc = Nothing
End Sub

Sub AddUserDefinition(nFileName As String)
    UpdateUserDatabase nFileName
End Sub
