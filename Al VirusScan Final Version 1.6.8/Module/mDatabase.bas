Attribute VB_Name = "mDatabase"
Option Explicit

Public myDatabase() As New ADODB.Connection

Function NotNull(objStr) As String
    If IsNull(objStr) Then
       NotNull = ""
    Else
       NotNull = objStr
    End If
End Function

Function LoadDatabase(dbaCon As ADODB.Connection, Filename As String) As Boolean
    On Error GoTo Salah
    Dim ConStr As String
    Dim syslog As String
    LockUnlock Filename, False
    
    Dim H(1 To 8) As String * 1
    H(1) = Chr(222)
    H(2) = Chr(221)
    H(3) = Chr(222)
    H(4) = Chr(221)
    H(5) = "r"
    H(6) = "o"
    H(7) = "o"
    H(8) = "t"
    syslog = H(1) & H(2) & H(3) & H(4) & H(5) & H(6) & H(7) & H(8)
        
    ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Filename & ";Persist Security Info=False;Jet OLEDB:Database Password=" & syslog
    dbaCon.Open ConStr
    LockUnlock Filename, True
    LoadDatabase = True
    Exit Function
Salah:
End Function

Function UpdateUserDatabase(nFileName As String)
    On Error GoTo Salah
    Dim data As New ADODB.Connection
    If LoadDatabase(data, nFileName) Then
       Dim r1 As New ADODB.Recordset
       Dim r2 As New ADODB.Recordset
       r1.Open "SELECT vdb_virus_collection_head.id, vdb_virus_collection_detail.id_virus, vdb_virus_collection_detail.virus_alias, vdb_virus_collection_detail.virus_type, vdb_virus_collection_detail.removal_script, vdb_virus_collection_detail.virus_crc_check, vdb_virus_collection_detail.default_action, vdb_virus_collection_detail.virus_date, vdb_virus_collection_detail.virus_like0, vdb_virus_collection_detail.virus_like1, vdb_virus_collection_detail.virus_like2, vdb_virus_collection_detail.virus_like3, vdb_virus_collection_detail.virus_like4, vdb_virus_collection_detail.virus_like5, vdb_virus_collection_detail.virus_like6, vdb_virus_collection_detail.virus_like7, vdb_virus_collection_detail.virus_like8, vdb_virus_collection_detail.virus_like9, vdb_virus_collection_detail.virus_crc_check2 " & _
               "FROM vdb_virus_collection_head INNER JOIN vdb_virus_collection_detail ON vdb_virus_collection_head.id = vdb_virus_collection_detail.id " & _
               "WHERE (((vdb_virus_collection_head.virus_name)='ADD_BY_USER'));", data, 3, 3
    
        '-----------------------------------------------
        ' Request ID
           Dim ID As String
           Dim rc As New ADODB.Recordset
             rc.Open "SELECT virus_name From vdb_virus_collection_head " & _
                     "WHERE (((virus_name)='ADD_BY_USER'));", myDatabase(0), 3, 3
                     
             If rc.EOF Then
                myDatabase(0).Execute "INSERT INTO vdb_virus_collection_head (virus_name,systems_affected) VALUES('ADD_BY_USER','All Windows')"
             End If
             rc.Close
             
             '------------
             rc.Open "SELECT id,virus_name From vdb_virus_collection_head " & _
                     "WHERE (((vdb_virus_collection_head.virus_name)='ADD_BY_USER'));", myDatabase(0), 3, 3
             
             If Not rc.EOF Then
                ID = NotNull(rc("id"))
             End If
             rc.Close
        '-----------------------------------------------
    
       If Not r1.EOF Then
          r2.Open "vdb_virus_collection_detail", myDatabase(0), 3, 3
          While Not r1.EOF
            '-----------------------------------------------
              r2.AddNew
                   r2("id") = ID
                   r2("virus_type") = NotNull(r1("virus_type"))
                   r2("default_action") = NotNull(r1("default_action"))
                   r2("virus_alias") = NotNull(r1("virus_alias"))
                   r2("virus_crc_check") = NotNull(r1("virus_crc_check"))
                   r2("virus_crc_check2") = NotNull(r1("virus_crc_check2"))
                   r2("removal_script") = NotNull(r1("removal_script"))
                   r2("virus_like0") = NotNull(r1("virus_like0"))
                   r2("virus_like1") = NotNull(r1("virus_like1"))
                   r2("virus_like2") = NotNull(r1("virus_like2"))
                   r2("virus_like3") = NotNull(r1("virus_like3"))
                   r2("virus_like4") = NotNull(r1("virus_like4"))
                   r2("virus_like5") = NotNull(r1("virus_like5"))
                   r2("virus_like6") = NotNull(r1("virus_like6"))
                   r2("virus_like7") = NotNull(r1("virus_like7"))
                   r2("virus_like8") = NotNull(r1("virus_like8"))
                   r2("virus_like9") = NotNull(r1("virus_like9"))
              r2.Update
            r1.MoveNext
          Wend
          If r2.state = 1 Then r2.Close
       End If
       r1.Close
       data.Close
    End If
Salah:
End Function

' Encryption
Function LockUnlock(Filename As String, Locked As Boolean) As Boolean
    On Error Resume Next
    Dim data(0) As Byte
    Dim Data2(160) As Byte
    If Filename <> "" Then
      Dim H As String
      H = Dir(Filename, vbArchive + vbNormal)
      If H <> "" Then
         LockUnlock = True
        Open Filename For Binary As #1
            Dim inpwd(160) As Byte, i As Integer
            Dim shpwd(160) As Byte
            Get #1, 160, data
            Get #1, 1, inpwd
            If Locked Then
               If data(0) = 0 Then
                  For i = 0 To 160
                    shpwd(i) = inpwd(i) Xor 255 Xor 19 Xor 3 Xor 81
                  Next i
                  Put #1, 1, shpwd
               End If
            Else
               If data(0) = &HBE Then
                  For i = 0 To 160
                    shpwd(i) = inpwd(i) Xor 255 Xor 19 Xor 3 Xor 81
                  Next i
                  Put #1, 1, shpwd
               End If
            End If
            LockUnlock = True
         Close #1
         
      Else
        LockUnlock = False
      End If
    End If
End Function
