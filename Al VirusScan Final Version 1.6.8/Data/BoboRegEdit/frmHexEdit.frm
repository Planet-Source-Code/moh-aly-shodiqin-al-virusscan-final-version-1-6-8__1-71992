VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHexEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Binary Value"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Value name"
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3870
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   2520
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox PicRtfBase 
      BackColor       =   &H80000009&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   0
      Top             =   1080
      Width           =   4995
      Begin RichTextLib.RichTextBox RTFLine 
         Height          =   3495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   6165
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmHexEdit.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.VScrollBar VS 
         Height          =   2835
         Left            =   4680
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin RichTextLib.RichTextBox RTFAscii 
         Height          =   3495
         Left            =   3600
         TabIndex        =   1
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   6165
         _Version        =   393217
         BorderStyle     =   0
         HideSelection   =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmHexEdit.frx":0080
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTFHex 
         Height          =   2775
         Left            =   600
         TabIndex        =   2
         Top             =   0
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   4895
         _Version        =   393217
         BorderStyle     =   0
         HideSelection   =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmHexEdit.frx":0100
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Value data :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Value name :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu mnuEditbase 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select All"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmHexEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Most of the Hex stuff is done in the class
'This code mainly deals with the interface
'which is made up of 3 Richtextboxes in a Picturebox
'Includes height calculation of a Richtextbox and
'scrolling all 3 Richtextboxes with a single VSScroll bar
'Selection of Hex and Ascii boxes simultaneously
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
Dim PrivateClipboard() As String
Dim ClipboardLoaded As Boolean
Private Const MM_TWIPS = 6
Dim HX As HexClass
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
    RTFHeight = TextM.tmHeight * RTFbox.GetLineFromChar(Len(RTFbox.Text)) / Screen.TwipsPerPixelY
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim mvar As Variant, mByte() As Byte, z As Long, temp As String
    temp = Right(fMainForm.TV.SelectedItem.Key, Len(fMainForm.TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    mvar = HX.HexData
    ReDim mByte(0 To UBound(mvar))
    For z = 0 To UBound(mvar)
        mByte(z) = Val(HexToDec(mvar(z)))
    Next
    z = Val(fMainForm.LV.SelectedItem.Tag)
    SaveSettingByte z, temp, fMainForm.LV.SelectedItem.Text, mByte
    fMainForm.LV.ListItems.Clear
    GetAllValues z, temp
    Unload Me
End Sub

Private Sub Form_Load()
    Dim mvar As Variant
    Dim temp As String
    Dim tmpByte() As String
    Dim byTemp2 As String
    Dim z As Long
    ReDim tmpByte(0 To 7) As String
    Me.Icon = fMainForm.Icon
    onlyloading = True
    RTFLine.Text = " "
    RTFLine.SelLength = 1
    RTFLine.SelColor = vbBlue
    Set HX = New HexClass
    Text1.Text = fMainForm.LV.SelectedItem.Text
    temp = Right(fMainForm.TV.SelectedItem.Key, Len(fMainForm.TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    mvar = GetSettingByte(Val(fMainForm.LV.SelectedItem.Tag), temp, fMainForm.LV.SelectedItem.Text)
    z = UBound(mvar)
    If z = 0 Then
        tmpByte(0) = mvar(0)
        For z = 1 To 7
            tmpByte(z) = 0
        Next
        HX.LoadRawBin tmpByte
    Else
        HX.LoadRawBin mvar
    End If
    RedrawData
    onlyloading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set HX = Nothing
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    'Allows cut and paste without using Windows Clipboard
    'which would complicate things needlessly
    Dim z As Long, p As Long, selp As Long
    z = RTFAscii.SelStart
    p = RTFAscii.GetLineFromChar(z)
    selp = RTFAscii.GetLineFromChar(z + RTFAscii.SelLength)
    Select Case Index
        Case 0
            PrivateClipboard = HX.GetBytes(z - p, z + RTFAscii.SelLength - selp)
            HX.RemoveBytes z - p, z + RTFAscii.SelLength - selp
            ClipboardLoaded = True
        Case 1
            PrivateClipboard = HX.GetBytes(z - p, z + RTFAscii.SelLength - selp)
            ClipboardLoaded = True
        Case 2
            HX.AddBytesAsBytes z - p, PrivateClipboard
        Case 3
            HX.RemoveBytes z - p, z + RTFAscii.SelLength - selp
        Case 5
            RTFAscii.SelStart = 0
            RTFAscii.SelLength = Len(RTFAscii.Text)
    End Select
    If Index <> 5 Then
        RedrawData
        RTFAscii.SelStart = z
    End If
End Sub

Private Sub mnuEditbase_Click()
    If ClipboardLoaded Then
        mnuEdit(2).Enabled = True
    Else
        mnuEdit(2).Enabled = False
    End If
    If RTFAscii.SelLength = 0 Then
        mnuEdit(0).Enabled = False
        mnuEdit(1).Enabled = False
        mnuEdit(3).Enabled = False
    Else
        mnuEdit(0).Enabled = True
        mnuEdit(1).Enabled = True
        mnuEdit(3).Enabled = True
    End If
End Sub


Private Sub RTFAscii_KeyPress(KeyAscii As Integer)
    Dim mSrc(0) As String, z As Long, p As Long
    If Shift <> 0 Then Exit Sub
    DoEvents
    mSrc(0) = Chr(KeyAscii)
    z = RTFAscii.SelStart
    p = RTFAscii.GetLineFromChar(z)
    If RTFAscii.SelLength > 0 Then
        selp = RTFAscii.GetLineFromChar(z + RTFAscii.SelLength)
        LockWindowUpdate Me.hWnd
        HX.RemoveBytes z - p, z + RTFAscii.SelLength - selp
        RedrawData
        LockWindowUpdate 0
    End If
    If KeyAscii = 22 Then
        HX.AddBytes z - p, Clipboard.GetText
        RedrawData
        RTFAscii.SelStart = z
    Else
        HX.EditByteByAsc z - p, mSrc
        LockWindowUpdate Me.hWnd
        RedrawData
        If RTFAscii.GetLineFromChar(z + 1) > p Then
            RTFAscii.SelStart = z + 2
        Else
            RTFAscii.SelStart = z + 1
        End If
    End If
    KeyAscii = 0
    LockWindowUpdate 0
End Sub


Private Sub RTFAscii_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuEditbase
    Else
        RTFHex.HideSelection = True
    End If
End Sub

Private Sub RTFAscii_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RTFHex.HideSelection = False
End Sub

Private Sub RTFAscii_SelChange()
    'This is tricky
    'Calculate selection of Hex based on selection of Ascii
    Dim z As Long, zl As Long, asSt As Long, heSt As Long
    Dim zx As Long, zx1 As Long, zx2 As Long, q As Long, q1 As Long
    Dim q2 As Long
    If Not selecting Then
        selecting = True
        z = RTFAscii.GetLineFromChar(RTFAscii.SelStart)
        zl = RTFAscii.GetLineFromChar(RTFAscii.SelStart + RTFAscii.SelLength)
        zl = zl - z
        asSt = RTFAscii.SelStart - z
        heSt = asSt * 3
        RTFHex.SelStart = heSt
        RTFHex.SelLength = (RTFAscii.SelLength - zl) * 3
        q1 = RTFHex.SelLength
        zx = RTFHex.GetLineFromChar(RTFHex.SelStart)
        zx1 = RTFHex.GetLineFromChar(RTFHex.SelStart + RTFHex.SelLength)
        zx2 = zx1
        zx1 = zx1 - zx
        q = RTFHex.SelStart - zx * 24
        Select Case q
            Case 1, 4, 7, 10, 13, 16, 19, 22
                RTFHex.SelStart = RTFHex.SelStart - 1
                If q1 <> 0 Then q1 = q1 + 1
            Case 2, 5, 8, 11, 14, 17, 20
                RTFHex.SelStart = RTFHex.SelStart + 1
                If q1 <> 0 Then q1 = q1 - 1
        End Select
        q2 = RTFHex.SelStart + q1 - zx2 * 24
        Select Case q2
            Case 1, 4, 7, 10, 13, 16, 19, 22
                If q1 <> 0 Then q1 = q1 + 1
            Case 0, 3, 6, 9, 12, 15, 18, 21
                If q1 <> 0 Then q1 = q1 - 1
        End Select
        RTFHex.SelLength = q1
        selecting = False
    End If
End Sub

Private Sub RTFHex_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode < 48 Or KeyCode > 57) And (KeyCode < 65 Or KeyCode > 70) And (KeyCode < 97 Or KeyCode > 102) Then
        KeyCode = 0
    End If

End Sub

Private Sub RTFHex_KeyPress(KeyAscii As Integer)
    Dim mSrc(0) As String, z As Long, p As Long, selp As Long, t As String
    DoEvents
    mSrc(0) = Chr(KeyAscii)
    FixHexSelection
    z = RTFHex.SelStart
    p = RTFHex.GetLineFromChar(z)
    If RTFHex.SelLength > 0 Then
        selp = RTFHex.GetLineFromChar(z + RTFHex.SelLength)
        LockWindowUpdate Me.hWnd
        HX.RemoveBytes z - p, z + RTFHex.SelLength - selp
        RTFHex.Text = HX.DataSplit(8, HX.HexData, Chr(32))
        RTFAscii.Text = HX.DataSplit(8, HX.AscData)
        LockWindowUpdate 0
    End If
    If z = 0 Then
        mSrc(0) = UCase(mSrc(0) + Mid(RTFHex.Text, 2, 1))
    Else
        If Mid(RTFHex.Text, z + p, 1) Like "[a-fA-F0-9]" Then
            mSrc(0) = UCase(Mid(RTFHex.Text, z + p, 1) + mSrc(0))
        Else
            mSrc(0) = UCase(mSrc(0) + Mid(RTFHex.Text, z + p + 2, 1))
        End If
    End If
    HX.EditByteByHex z / 3, mSrc
    KeyAscii = 0
    LockWindowUpdate Me.hWnd
    RedrawData
    If RTFHex.GetLineFromChar(z + 1) > p Then
        RTFHex.SelStart = z + 2
    Else
        RTFHex.SelStart = z + 1
    End If
    LockWindowUpdate 0
End Sub
Private Sub RTFHex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then RTFAscii.HideSelection = True

End Sub

Private Sub RTFHex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If RTFHex.SelLength = 0 Then RTFAscii.SelLength = 0
    If Button = 2 Then
        If RTFHex.SelLength = 0 Then
            LockWindowUpdate Me.hWnd
            selecting = True
            RTFHex.SelLength = 1
            FixHexSelection
            RTFHex.SelLength = 0
            selecting = False
            LockWindowUpdate 0
        End If
        Me.PopupMenu mnuEditbase
    Else
        RTFAscii.HideSelection = False
        FixHexSelection
    End If
End Sub

Private Sub RTFHex_SelChange()
    'This is easier
    'Calculate selection of Ascii based on selection of Hex
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
    If RTFHex.Height < PicRtfBase.Height Then RTFHex.Height = PicRtfBase.Height
    RTFAscii.Height = RTFHex.Height
    RTFLine.Height = RTFHex.Height
    VS.Max = RTFHex.Height - VS.Height
End Sub
Public Function FixHexPosition()
    Dim q As Long, q1 As Long, z As Long
    z = RTFHex.GetLineFromChar(RTFHex.SelStart)
    q = RTFHex.SelStart
    q1 = q Mod 24
    Select Case q1
        Case 2, 5, 8, 11, 14, 17, 20, 23
            q = q + 1
        Case 24
    End Select
    RTFHex.SelStart = q
    If RTFHex.SelStart + z > Len(RTFHex.Text) - 2 Then
        LockWindowUpdate Me.hWnd
        HX.AddEmptyLine
        RedrawData
        RTFHex.SelStart = q
        LockWindowUpdate 0
    End If
End Function
Public Sub FixHexSelection()
    Dim q As Long, q1 As Long, q2 As Long
    If RTFHex.SelLength = 0 Then
        FixHexPosition
        Exit Sub
    End If
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

Private Sub RTFLine_GotFocus()
    Dim z As Long
    z = RTFLine.GetLineFromChar(RTFLine.SelStart)
    RTFHex.SelStart = z * 24
    RTFHex.SetFocus
End Sub

Private Sub RTFLine_SelChange()
    If Not onlyloading Then RTFLine.SelLength = 0
End Sub

Private Sub VS_Scroll()
    RTFHex.Top = -VS.Value
    RTFAscii.Top = -VS.Value
    RTFLine.Top = -VS.Value
End Sub

Public Sub RedrawData()
    RTFHex.Text = HX.DataSplit(8, HX.HexData, Chr(32))
    RTFAscii.Text = HX.DataSplit(8, HX.AscData)
    RTFLine.Text = HX.CreateByteList
    SetUpScrolling
End Sub

