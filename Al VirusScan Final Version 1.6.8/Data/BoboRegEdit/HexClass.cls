VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HexClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'I rushed th Hex Editor a bit
'Ended up wrapping into this class
'seems to work OK
Dim mHex() As String
Dim mBin() As Byte
Dim mAsc() As String
Dim HexString As String
Dim AscString As String
Public Property Get HexData() As Variant
    HexData = mHex
End Property
Public Property Let HexData(ByVal vNewValue As Variant)
    Dim z As Long
    Erase mHex
    Erase mAsc
    mHex = vNewValue
    ReDim mAsc(0 To UBound(mHex))
    For z = 0 To UBound(mHex)
        mAsc(z) = HexToAsc(mHex(z))
    Next
    HexString = Join(mHex, Chr(32))
    AscString = Join(mAsc, Chr(32))
End Property
Public Property Get BinData() As Variant
    BinData = mBin
End Property
Public Property Let BinData(ByVal vNewValue As Variant)
    mBin = vNewValue
End Property
Public Property Get AscData() As Variant
    AscData = mAsc
End Property
Public Property Let AscData(ByVal vNewValue As Variant)
    mAsc = vNewValue
End Property
Public Property Get HexStr() As String
    HexStr = HexString
End Property
Public Property Let HexStr(ByVal vNewValue As String)
    HexString = vNewValue
End Property
Public Property Get AscStr() As String
    AscStr = AscString
End Property
Public Property Let AscStr(ByVal vNewValue As String)
    AscString = vNewValue
End Property
Public Sub LoadRawHex(mSrc As String)
    Dim mvar As Variant, z As Long, z1 As Long
    HexString = Replace(mSrc, vbLf, Chr(32))
    HexString = Replace(HexString, vbCrLf, Chr(32))
    HexString = Replace(HexString, vbCr, Chr(32))
    mvar = Split(Trim(HexString), Chr(32))
    For z = 0 To UBound(mvar)
        If Len(mvar(z)) <> 0 Then z1 = z1 + 1
    Next
    ReDim mHex(0 To z1)
    ReDim mAsc(0 To z1)
    z1 = 0
    For z = 0 To UBound(mvar)
        If Len(mvar(z)) <> 0 Then
            mHex(z1) = mvar(z)
            mAsc(z1) = HexToAsc(mvar(z))
            z1 = z1 + 1
        End If
    Next
End Sub
Public Sub LoadRawAsc(mSrc As String)
    Dim z As Long, mvar As Variant
    AscString = Replace(mSrc, vbLf, Chr(32))
    AscString = Replace(AscString, vbCrLf, Chr(32))
    AscString = Replace(AscString, vbCr, Chr(32))
    mvar = Split(AscString, Chr(32))
    ReDim mHex(0 To UBound(mvar))
    ReDim mAsc(0 To UBound(mvar))
    For z = 0 To UBound(mvar)
        mHex(z) = Hex$(Asc((mvar(z))))
        If Len(mHex(z)) = 1 Then mHex(z) = "0" + mHex(z)
        If Len(mHex(z)) = 0 Then mHex(z) = "00"
        mAsc(z) = mvar(z)
    Next
    HexString = Join(mHex, Chr(32))
    AscString = Join(mAsc, Chr(32))
End Sub
Public Sub LoadRawBin(arrByte As Variant)
    Dim z As Long, q As Long
    ReDim mHex(0 To UBound(arrByte))
    ReDim mAsc(0 To UBound(arrByte))
    For z = 0 To UBound(arrByte)
        mHex(z) = Format(Hex$(arrByte(z)), "00")
        If Len(mHex(z)) = 1 Then mHex(z) = "0" + mHex(z)
        If Len(mHex(z)) = 0 Then mHex(z) = "00"
        mAsc(z) = Str(arrByte(z))
        q = Val(mAsc(z))
        If q < 33 Or (q > 126 And q < 144) Or (q > 147 And q < 161) Then
            mAsc(z) = Chr(46)
        Else
            mAsc(z) = Chr(q)
        End If
    Next
    HexString = Join(mHex, Chr(32))
    AscString = Join(mAsc, Chr(32))
End Sub
Public Function DataSplit(mWidth As Long, mSrc As Variant, Optional mSeparator As String = "") As String
    Dim z As Long, cnt As Long, temp As String
    For z = 0 To UBound(mSrc)
        If ((z + 1) Mod mWidth) = 0 Then
            temp = temp + mSrc(z) + vbLf
        Else
            temp = temp + mSrc(z) + mSeparator
        End If
    Next
    DataSplit = temp
End Function
Private Function HexToAsc(ByVal HexStr As String) As String
    Dim mult As Double
    Dim DecNum As Long
    Dim ch As String
    Dim tempASC As String
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
                    HexToAsc = ""
                    Exit Function
                End If
            End If
        End If
        mult = mult * 16
    Next i
    tempASC = Chr(DecNum)
    If DecNum < 33 Or (DecNum > 126 And DecNum < 144) Or (DecNum > 147 And DecNum < 161) Then
        tempASC = Chr(46)
    Else
        tempASC = Chr(DecNum)
    End If
    HexToAsc = tempASC
End Function
Public Sub EditByteByAsc(mStartIndex As Long, mSrc As Variant)
    Dim z As Long
    If mStartIndex > UBound(mAsc) Then
        z = mStartIndex
        If (mStartIndex - 1 Mod 2) <> 0 Then z = z + 1
        ReDim Preserve mAsc(0 To z)
        ReDim Preserve mHex(0 To z)
    End If
    For z = mStartIndex To mStartIndex + UBound(mSrc)
        mAsc(z) = mSrc(z - mStartIndex)
        mHex(z) = Hex$(Asc((mSrc(z - mStartIndex))))
        If Len(mHex(z)) = 1 Then mHex(z) = "0" + mHex(z)
        If Len(mHex(z)) = 0 Then mHex(z) = "00"
    Next
End Sub
Public Sub EditByteByHex(mStartIndex As Long, mSrc As Variant)
    Dim z As Long
    If mStartIndex + UBound(mSrc) > UBound(mAsc) Then
        z = mStartIndex + UBound(mSrc)
        ReDim Preserve mAsc(0 To z)
        ReDim Preserve mHex(0 To z)
    End If
    For z = mStartIndex To mStartIndex + UBound(mSrc)
        mAsc(z) = HexToAsc(mSrc(z - mStartIndex))
        mHex(z) = mSrc(z - mStartIndex)
        If Len(mHex(z)) = 1 Then mHex(z) = "0" + mHex(z)
        If Len(mHex(z)) = 0 Then mHex(z) = "00"
    Next
    HexData = mHex
End Sub
Public Sub RemoveBytes(mStart As Long, mFin As Long)
    Dim NewCnt As Long, tmpBytes() As String, z As Long, nz As Long
    NewCnt = UBound(mHex) - (mFin - mStart)
    ReDim tmpBytes(0 To NewCnt)
    For z = 0 To mStart - 1
        tmpBytes(z) = mHex(z)
    Next
    For nz = mFin + 1 To UBound(mHex)
        tmpBytes(z) = mHex(nz)
        z = z + 1
    Next
    HexData = tmpBytes
End Sub
Public Sub AddBytesAsBytes(mStart As Long, mInsert As Variant)
    Dim NewCnt As Long, tmpBytes() As String, z As Long, nz As Long, nzz As Long
    Dim cnt As Long
    NewCnt = UBound(mHex) + UBound(mInsert) + 1
    ReDim tmpBytes(0 To NewCnt)
    For z = 0 To mStart - 1
        tmpBytes(z) = mHex(z)
    Next
    For nz = mStart To mStart + UBound(mInsert)
        tmpBytes(nz) = mInsert(cnt)
        cnt = cnt + 1
    Next
    For nzz = mStart To UBound(mHex)
        tmpBytes(nz) = mHex(nzz)
        nz = nz + 1
    Next
    HexData = tmpBytes
End Sub
Public Sub AddBytes(mStart As Long, mInsert As String)
    Dim NewCnt As Long, tmpBytes() As String, z As Long, nz As Long, nzz As Long
    Dim tmpStr As String, temp As String
    tmpStr = Replace(mInsert, vbLf, "")
    tmpStr = Replace(tmpStr, vbCrLf, "")
    tmpStr = Replace(tmpStr, vbCr, "")
    NewCnt = UBound(mHex) + Len(tmpStr)
    ReDim tmpBytes(0 To NewCnt)
    For z = 0 To mStart
        tmpBytes(z) = mHex(z)
    Next
    For nz = 1 To Len(tmpStr)
        temp = Mid(tmpStr, nz, 1)
        If Trim(temp) <> "" Then mHex(z + nz - 1) = Hex$(Asc(temp))
        If Len(mHex(z + nz - 1)) = 1 Then mHex(z + nz - 1) = "0" + mHex(z + nz - 1)
        If Len(mHex(z + nz - 1)) = 0 Then mHex(z + nz - 1) = "00"
    Next
    For nzz = mStart + 1 To UBound(mHex)
        tmpBytes(z + nz - 1) = mHex(nzz)
        z = z + 1
    Next
    HexData = tmpBytes
End Sub
Public Sub AddEmptyLine()
    Dim tmpBytes() As String, z As Long
    Dim tmpStr As String, temp As String
    ReDim tmpBytes(0 To UBound(mHex) + 8)
    For z = 0 To UBound(mHex)
        tmpBytes(z) = mHex(z)
    Next
    For z = UBound(mHex) + 1 To UBound(mHex) + 8
        tmpBytes(z) = "00"
    Next
    HexData = tmpBytes
End Sub
Public Function GetBytes(mStart As Long, mEnd As Long) As Variant
    Dim z As Long, tmpStr() As String, cnt As Long
    ReDim tmpStr(0 To mEnd - mStart)
    For z = mStart To mEnd
        If z > UBound(mHex) Then Exit For
        tmpStr(cnt) = mHex(z)
        cnt = cnt + 1
    Next
    GetBytes = tmpStr
End Function
Public Function CreateByteList() As String
    Dim bList() As String, cnt As Long
    ReDim bList(0 To Int(UBound(mHex) / 8) + 1)
    For z = 0 To UBound(mHex) Step 8
        bList(cnt) = Hex$(z)
        If Len(bList(cnt)) = 1 Then bList(cnt) = "000" + bList(cnt)
        If Len(bList(cnt)) = 2 Then bList(cnt) = "00" + bList(cnt)
        If Len(bList(cnt)) = 3 Then bList(cnt) = "0" + bList(cnt)
        If Len(bList(cnt)) > 4 Then bList(cnt) = Right(bList(cnt), 4)
        cnt = cnt + 1
    Next z
    CreateByteList = Join(bList, vbLf)
End Function
