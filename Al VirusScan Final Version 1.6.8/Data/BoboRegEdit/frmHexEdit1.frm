VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHexEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Binary Data"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4440
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   3120
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.VScrollBar VS 
      Height          =   2655
      Left            =   5280
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   1
      Top             =   1080
      Width           =   5175
      Begin RichTextLib.RichTextBox RTFAscii 
         Height          =   3495
         Left            =   3960
         TabIndex        =   2
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   6165
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmHexEdit.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTFHex 
         Height          =   3495
         Left            =   840
         TabIndex        =   3
         Top             =   0
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   6165
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmHexEdit.frx":007D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTFLine 
         Height          =   3495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   6165
         _Version        =   393217
         BorderStyle     =   0
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmHexEdit.frx":00FA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Value name :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Value data :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmHexEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type
Dim selecting As Boolean
Dim onlyloading As Boolean
Private Const MM_TWIPS = 6

Public Function RTFHeight(RTFbox As RichTextBox) As Long
    Dim hdc As Long, z As Long
    Dim PscaleMode As Long
    Dim TextM As TEXTMETRIC
    hdc = GetWindowDC(RTFbox.hWnd)
    If hdc Then
        PscaleMode = SetMapMode(hdc, MM_TWIPS)
        GetTextMetrics hdc, TextM
        ReleaseDC hWnd, hdc
    End If
    RTFHeight = TextM.tmAscent * RTFbox.GetLineFromChar(Len(RTFbox.Text)) / Screen.TwipsPerPixelY
End Function



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim mvar As Variant, mByte() As Byte, z As Long, temp As String
    temp = Right(fMainForm.TV.SelectedItem.Key, Len(fMainForm.TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    mvar = Split(RTFHex.Text, " ")
    ReDim mByte(0 To UBound(mvar))
    For z = 0 To UBound(mvar)
        mByte(z) = Val(HexToDec(mvar(z)))
    Next
    SaveSettingByte Val(fMainForm.LV.SelectedItem.Tag), temp, fMainForm.LV.SelectedItem.Text, mByte
End Sub

Private Sub Form_Load()
    Dim arrByte() As Byte, arrStr() As String, mvar As Variant
    Dim mHandle As Long, z As Long, z1 As Long, temp As String, temp1 As String
    Dim tempASC As String, tempASC1 As String, q As Long
    Dim tempLine As String, tempLine1 As String
    onlyloading = True
'    mHandle = FreeFile
'    Open Label1.Caption For Binary As #mHandle
'        ReDim arrByte(1 To LOF(mHandle))
'        ReDim arrStr(1 To LOF(mHandle))
'        Get #mHandle, , arrByte
'    Close mHandle
    Label1.Caption = fMainForm.LV.SelectedItem.Text
    temp = Right(fMainForm.TV.SelectedItem.Key, Len(fMainForm.TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    mvar = GetSettingByte(Val(fMainForm.LV.SelectedItem.Tag), temp, fMainForm.LV.SelectedItem.Text)
    temp = ""
    ReDim arrByte(1 To UBound(mvar) + 1)
    For z = 0 To UBound(mvar)
        arrByte(z + 1) = mvar(z)
    Next
    ReDim arrStr(1 To UBound(arrByte))
    Dim AsStr() As String
    Dim HexStr() As String
    Dim cnt As Long
    cnt = UBound(arrByte)
    If (cnt Mod 8) = 0 Then
        cnt = cnt / 8
    Else
        If Int((cnt / 8)) * 8 < cnt Then
            cnt = cnt / 8 + 1
        Else
            cnt = cnt / 8
        End If
    End If
    ReDim HexStr(1 To cnt)
    ReDim AsStr(1 To cnt)
    cnt = 1
    tempLine = "0000"
    For z = 1 To UBound(arrByte)
        temp1 = Format(Hex$(arrByte(z)), "00")
        If Len(temp1) = 1 Then temp1 = "0" + temp1
        If Len(temp1) = 0 Then temp1 = "00"
        tempASC = Str(arrByte(z))
        q = Val(tempASC)
        If q < 33 Or (q > 126 And q < 144) Or (q > 147 And q < 161) Then
            tempASC = Chr(46)
        Else
            tempASC = Chr(q)
        End If
        If (z Mod 8) = 0 Then
            tempASC1 = tempASC1 + tempASC
            temp = temp + temp1
            AsStr(cnt) = Trim(tempASC1)
            HexStr(cnt) = Trim(temp)
            tempLine1 = Hex$(z)
            If Len(tempLine1) = 1 Then tempLine1 = "000" + tempLine1
            If Len(tempLine1) = 2 Then tempLine1 = "00" + tempLine1
            If Len(tempLine1) = 3 Then tempLine1 = "0" + tempLine1
            If Len(tempLine1) > 4 Then tempLine1 = Left(tempLine1, 4)
            tempLine = tempLine + vbCrLf + tempLine1
            cnt = cnt + 1
            tempASC1 = ""
            temp = ""
        Else
            tempASC1 = tempASC1 + tempASC
            temp = temp + temp1 + " "
        End If
    Next
'    z1 = (z Mod 8)
'    If (z Mod 8) <> 0 Then
'        tempLine1 = Hex$(z - (z Mod 8) + 8)
'        If Len(tempLine1) = 1 Then tempLine1 = "000" + tempLine1
'        If Len(tempLine1) = 2 Then tempLine1 = "00" + tempLine1
'        If Len(tempLine1) = 3 Then tempLine1 = "0" + tempLine1
'        If Len(tempLine1) > 4 Then tempLine1 = Left(tempLine1, 4)
'        tempLine = tempLine + vbCrLf + tempLine1
'        For z1 = z - (z Mod 8) To z - 1
'             temp1 = Format(Hex$(arrByte(z1)), "00")
'             If Len(temp1) = 1 Then temp1 = " " + temp1
'             If Len(temp1) = 0 Then temp1 = "  "
'             tempASC = Str(arrByte(z1))
'             q = Val(tempASC)
'             If q < 33 Or (q > 126 And q < 144) Or (q > 147 And q < 161) Then
'                 tempASC = Chr(46)
'             Else
'                 tempASC = Chr(q)
'             End If
'             tempASC1 = tempASC1 + tempASC
'             temp = temp + temp1 + " "
'             AsStr(cnt) = Trim(tempASC1)
'             HexStr(cnt) = Trim(temp)
'        Next z1
'    End If
    RTFHex.Text = Join(HexStr, vbCr)
    RTFAscii.Text = Join(AsStr, vbCr)
    RTFLine.Text = tempLine
    SetUpScrolling
    onlyloading = False

End Sub

Private Sub RTFAscii_Change()
    If Not onlyloading Then
        
    End If
End Sub

Private Sub RTFAscii_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim z As Long
    Select Case KeyCode
    Case vbKeyBack
        KeyCode = 0
        Exit Sub
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyEnd, vbKeyHome
        Exit Sub
    Case vbKeyDelete
        KeyCode = 32
    End Select
    RTFAscii.HideSelection = True
    selecting = True
    RTFAscii.SelLength = 1
    z = InStr(1, RTFAscii.SelText, vbCrLf)
    If z <> 0 Then RTFAscii.SelStart = RTFAscii.SelStart + z
    RTFAscii.SelLength = 1
End Sub

Private Sub RTFAscii_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim z As Long, zl As Long, asSt As Long, heSt As Long
    Select Case KeyCode
    Case vbKeyBack, vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyEnd, vbKeyHome
        KeyCode = 0
        Exit Sub
    Case vbKeyDelete
        KeyCode = 32
    End Select
    z = RTFAscii.GetLineFromChar(RTFAscii.SelStart - 1)
    asSt = RTFAscii.SelStart - 1 - z
    heSt = asSt * 3
    RTFHex.HideSelection = True
    RTFHex.SelStart = heSt
    RTFHex.SelLength = 2
    RTFHex.SelText = Hex$(KeyCode)
    RTFHex.HideSelection = False
    RTFAscii.HideSelection = False
    selecting = False
End Sub


Private Sub RTFAscii_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RTFHex.HideSelection = True
End Sub

Private Sub RTFAscii_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RTFHex.HideSelection = False
End Sub

Private Sub RTFAscii_SelChange()
    Dim z As Long, z1 As Long
    If Not selecting Then
        selecting = True
        z = RTFAscii.GetLineFromChar(RTFAscii.SelStart)
        z1 = RTFAscii.GetLineFromChar(RTFAscii.SelStart + RTFAscii.SelLength)
        z1 = z1 - z
        RTFHex.SelStart = (RTFAscii.SelStart - z) * 3
        RTFHex.SelLength = (RTFAscii.SelLength - z1) * 3
        FixHexSelection
        selecting = False
    End If
End Sub

Private Sub RTFHex_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim zx As Long, zx1 As Long, zx2 As Long, asSt As Long, heSt As Long, q As Long, q1 As Long
    Dim q2 As Long, temp As String, tmpText As String
    If (KeyCode < 48 Or KeyCode > 57) And (KeyCode < 65 Or KeyCode > 70) And (KeyCode < 97 Or KeyCode > 102) Then
        KeyCode = 0
        Exit Sub
    End If
    Select Case KeyCode
    Case vbKeyBack, vbKeyDelete, vbKeyNumpad0, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9
        KeyCode = 0
        Exit Sub
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyEnd, vbKeyHome
        Exit Sub
    End Select
    selecting = True
    zx = RTFHex.GetLineFromChar(RTFHex.SelStart)
    q = RTFHex.SelStart - zx * 24
    Select Case q
        Case 1, 4, 7, 10, 13, 16, 19, 22
            temp = Mid(RTFHex.Text, RTFHex.SelStart + 1, 2)
        Case 2, 5, 8, 11, 14, 17, 20
            RTFHex.SelStart = RTFHex.SelStart + 1
            temp = Mid(RTFHex.Text, RTFHex.SelStart, 2)
    End Select
    temp = HexToAsc(temp)
    If temp = "" Then temp = "."
    RTFHex.SelLength = 1
    z = InStr(1, RTFHex.SelText, vbCrLf)
    If z <> 0 Then RTFHex.SelStart = RTFHex.SelStart + z

    RTFHex.SelLength = 1
    asSt = RTFHex.SelStart + zx * 3
    heSt = asSt / 3
    RTFAscii.SelStart = heSt
    RTFAscii.SelLength = 1
    z = InStr(1, RTFAscii.SelText, vbCrLf)
    If z <> 0 Then RTFAscii.SelStart = RTFAscii.SelStart + z
    RTFAscii.SelLength = 1
    RTFAscii.SelText = temp
    selecting = False
End Sub

Private Sub RTFHex_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 65 Or KeyAscii > 70) And (KeyAscii < 97 Or KeyAscii > 102) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub RTFHex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RTFAscii.HideSelection = True
End Sub

Private Sub RTFHex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If RTFHex.SelLength = 0 Then RTFAscii.SelLength = 0
    RTFAscii.HideSelection = False
End Sub

Private Sub RTFHex_SelChange()
    Dim z As Long, z1 As Long
    If Not selecting Then
        selecting = True
        FixHexSelection
        z = RTFHex.GetLineFromChar(RTFHex.SelStart)
        z1 = RTFHex.GetLineFromChar(RTFHex.SelStart + RTFHex.SelLength)
        z1 = z1 - z
        RTFAscii.SelStart = RTFHex.SelStart / 3 + z
        RTFAscii.SelLength = (RTFHex.SelLength) / 3 + z1
        selecting = False
    End If
End Sub
Public Sub SetUpScrolling()
    Dim hh As Long
    RTFHex.Height = RTFHeight(RTFHex) / 5 * 4
    RTFAscii.Height = RTFHex.Height
    RTFLine.Height = RTFHex.Height
    VS.Max = RTFHex.Height - VS.Height
End Sub

Private Sub RTFLine_SelChange()
    RTFLine.SelLength = 0
End Sub

Private Sub VS_Scroll()
    RTFHex.Top = -VS.Value
    RTFAscii.Top = -VS.Value
    RTFLine.Top = -VS.Value
End Sub
Public Sub FixHexSelection()
    If RTFHex.SelLength = 0 Then Exit Sub
    q = RTFHex.SelStart
    q1 = q + RTFHex.SelLength
    q2 = q Mod 24
    Select Case q2
        Case 0, 3, 6, 9, 12, 15, 18, 21
            
        Case 1, 4, 7, 10, 13, 16, 19, 22
            q = q - 1
        Case 2, 5, 8, 11, 14, 17, 20, 23
            q = q + 1
    End Select
    q2 = q1 Mod 24
    Select Case q2
        Case 0, 3, 6, 9, 12, 15, 18, 21
            q1 = q1 - 1
        Case 1, 4, 7, 10, 13, 16, 19, 22
            q1 = q1 + 1
        Case 2, 5, 8, 11, 14, 17, 20, 23
        
    End Select
    RTFHex.SelStart = q
    If q > q1 Then
        RTFHex.SelLength = 0
    Else
        RTFHex.SelLength = q1 - q
    End If
End Sub


