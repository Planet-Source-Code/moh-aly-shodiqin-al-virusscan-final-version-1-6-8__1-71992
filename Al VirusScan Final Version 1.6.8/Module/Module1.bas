Attribute VB_Name = "Module1"
Option Explicit

Public viri_col As New Collection
 
Public Sub Initialize_viri()
    On Error Resume Next
    Dim VirStr(1 To 4) As String
    Dim strArray(5) As String
    Dim i As Long
    
    strArray(1) = "TH.EXPLOIT.DCOM.A;CF586BC4FEE0BFD3A5B1338239F287C6;DELETE;Trojan Horse"
    strArray(2) = "RONTOKBRO;B47C07C30454B49AAA93708BC6173B8E;DELETE;Trojan Horse"
    strArray(3) = "KSpoold.XLS/DOC;35B0B817ECEFB83E2FCB951DDC010C4C;DELETE;Trojan Horse"
    strArray(4) = "KSpoold.Service;E22DFED2D50A1B5DBBACE3F9C535565B;DELETE;Trojan Horse"
    strArray(5) = "TH.EXPLOIT DCOM.B;BA5B994D260CD5F249F23C80687B0FD6;DELETE;Trojan Horse"
    
    For i = 1 To UBound(strArray)
        VirStr(1) = Split(strArray(i), ";")(0)
        VirStr(2) = Split(strArray(i), ";")(1)
        VirStr(3) = Split(strArray(i), ";")(2)
        VirStr(4) = Split(strArray(i), ";")(3)
        viri_col.Add VirStr
    Next i
End Sub
