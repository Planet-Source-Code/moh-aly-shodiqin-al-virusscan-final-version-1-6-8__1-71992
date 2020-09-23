VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "VirusScan Registry Editor"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9660
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   1560
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   5
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   6150
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16510
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9660
      TabIndex        =   3
      Top             =   0
      Width           =   9660
      Begin VB.ComboBox cboAddress 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   720
         List            =   "frmMain.frx":0449
         TabIndex        =   5
         Text            =   "cboAddress"
         Top             =   30
         Width           =   2055
      End
      Begin MSComctlLib.Toolbar TB 
         Height          =   390
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               ImageIndex      =   2
            EndProperty
         EndProperty
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   2400
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":045A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":05B4
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":070E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1242
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":139C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3000
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2880
      Left            =   3360
      TabIndex        =   0
      Top             =   410
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   5080
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   2520
      Left            =   0
      TabIndex        =   1
      Top             =   410
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   4445
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2760
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import Registry File"
      End
      Begin VB.Menu mnuFileSaveKey 
         Caption         =   "Export &Key"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEditbase 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditNew 
         Caption         =   "&New"
         Enabled         =   0   'False
         Begin VB.Menu mnuEditNewKey 
            Caption         =   "&Key"
         End
         Begin VB.Menu mnuEditNewSP1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditNewValue 
            Caption         =   "&String Value"
            Index           =   0
         End
         Begin VB.Menu mnuEditNewValue 
            Caption         =   "&Long Value"
            Index           =   1
         End
         Begin VB.Menu mnuEditNewValue 
            Caption         =   "&Binary Vale"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Modify"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Rename"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Bookmark"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Reload this Key"
         Index           =   5
      End
      Begin VB.Menu mnuEditSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu mnuFavoritesBase 
      Caption         =   "Favorites"
      Begin VB.Menu mnuFavoritesAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuFavoritesOrganise 
         Caption         =   "Organise"
      End
      Begin VB.Menu mnuFavorites 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFavorites 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I never intended to submit this code to PSC
'It was written in under 3 hours to win a bet !!
'Consequently very little commenting of code
'90% of code is original with the rest being
'pretty standard and freely available code snippets
'Apologies to the authors of such code as I have no
'idea who you are.

'The Registry lies at the core of the OS
'All coders should be familiar with the
'layout and usage of the Registry - if
'your not then here it is on a plate !
'
'To newbies :
'    Sorry about the sparsity of comments
'    but it 's ALL here, you'll have to search to LEARN !
'    Anything you wanted to know or do in
'    Registry is here somewhere
'
'To more advanced coders:
'    Only the more peculiar functions have any comments
'    The Registry side of things is straight forward enough
'    In fact most of the code is Interface manipulation
'
'Demonstrates:
'How to implement all of Regedit's functionality
'Edits Binary,Strings,Unicode strings,Hex
'Improved Search facilities
'Includes simple Hex Editor which demostrates
'controlling multiple Richtextboxes including
'simultaneous selection, and scrolling all boxes
'with a single scrollbar
'
'As with all software using the registry
'BACKUP your registry before using
'(You shouldn't have any problems -
'the app can only do what you tell it to)
'Whilst it has not had a huge amount of testing
'everything seems to work - I won the bet !
'I suspect what bugs there are will be just
'"House-keeping" type interface glitches

'I never got around to making use of the
'Statusbar - dont know why I left it in really
Private Const TVS_NOTOOLTIPS = &H80
Private Const TVS_NOHSCROLL = &H8000
Dim ChildrenNodes As Collection
Dim DontUpdate As Boolean
Dim mbMoving As Boolean
Dim TVPopup As Boolean
Dim SlItem As ListItem
Const sglSplitLimit = 500
Dim PrelabelEdit As String
Private Sub cboAddress_Click()
    If Not DontUpdate Then
        TV.Scroll = False
        TV.Nodes(cboAddress.Text + "\").Selected = True
        TV.SetFocus
        TV_NodeClick TV.SelectedItem
        TV.Scroll = True
    End If
    cboAddress.SelStart = Len(cboAddress.Text)
    cboAddress.SelLength = 0
    CheckBackForward
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboAddress.AddItem cboAddress.Text
        TV.Scroll = False
        TV.Nodes(cboAddress.Text + "\").Selected = True
        TV.Nodes(cboAddress.Text + "\").Expanded = True
        TV.Nodes(cboAddress.Text + "\").EnsureVisible
        TV.SetFocus
        TV_NodeClick TV.SelectedItem
        TV.Scroll = True
    End If
End Sub

Private Sub Form_Initialize()
    InitCommonControls
'    Me.Height = 6855
End Sub

Private Sub Form_Load()
    Dim hHeader As Long, strnames() As String, z As Long, temp As String
    'Flat columnheaders for Listview
    hHeader = SendMessage(LV.hWnd, LVM_GETHEADER, 0, ByVal 0&)
    SetWindowLong hHeader, GWL_STYLE, GetWindowLong(hHeader, GWL_STYLE) Xor HDS_BUTTONS
    'Tooltips in Treeview get in the way - ditch them
    SetWindowLong TV.hWnd, GWL_STYLE, GetWindowLong(TV.hWnd, GWL_STYLE) Or TVS_NOTOOLTIPS
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 9780)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 7260)
    'Read our Favorites list - from where else but Registry
    strnames = GetAllValues(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\BoboRegEdit\Favorites", True)
    If strnames(0) <> "  " Then
        For z = 0 To UBound(strnames)
            If strnames(z) <> "" Then
                temp = GetSettingString(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\BoboRegEdit\Favorites", strnames(z), "")
                If temp <> "" Then
                    AddFave temp, strnames(z)
                End If
            End If
        Next
    End If
    DoEvents
    GetHKEYS
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    Dim z As Long
    If mnuFavorites(1).Visible Then
        For z = 1 To mnuFavorites.count - 1
            SaveSettingString HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\BoboRegEdit\Favorites", mnuFavorites(z).Caption, mnuFavorites(z).Tag
        Next
    End If
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 4500 Then Me.Width = 4500
    SizeControls imgSplitter.Left
    cboAddress.Width = Me.Width - cboAddress.Left - 150
End Sub
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub
Sub SizeControls(x As Single)
    On Error Resume Next
    If x < 2700 Then x = 2700
    If x > (Me.Width - 2000) Then x = Me.Width - 2000
    TV.Height = Me.Height - PicAddress.Height - SB.Height - 800
    TV.Width = x
    imgSplitter.Left = x
    LV.Left = x + 40
    LV.Width = Me.Width - (TV.Width + 140)
    LV.Height = TV.Height
    LV.ColumnHeaders(3).Width = LV.Width - LV.ColumnHeaders(1).Width - LV.ColumnHeaders(2).Width - 120
    imgSplitter.Top = TV.Top
    imgSplitter.Height = TV.Height
End Sub

Public Sub GetHKEYS()
    'Get the root keys
    Dim fred As Node, strnames() As String, z As Long
    Set fred = TV.Nodes.Add(, , "My Computer\", "My Computer", 1, 1)
    fred.Tag = "F"
    If CountAllKeys(HKEY_CLASSES_ROOT) Then
        Set fred = TV.Nodes.Add("My Computer\", tvwChild, "HKEY_CLASSES_ROOT\", "HKEY_CLASSES_ROOT", 1, 2)
        fred.Tag = "UF"
        Set fred = TV.Nodes.Add("HKEY_CLASSES_ROOT\", tvwChild, "HKEY_CLASSES_ROOT\Dummy", "Dummy", 1, 2)
        fred.Tag = "D"
    End If
    If CountAllKeys(HKEY_CURRENT_USER) Then
        Set fred = TV.Nodes.Add("My Computer\", tvwChild, "HKEY_CURRENT_USER\", "HKEY_CURRENT_USER", 1, 2)
        fred.Tag = "UF"
        Set fred = TV.Nodes.Add("HKEY_CURRENT_USER\", tvwChild, "HKEY_CURRENT_USER\Dummy", "Dummy", 1, 2)
        fred.Tag = "D"
    End If
    If CountAllKeys(HKEY_DYN_DATA) Then
        Set fred = TV.Nodes.Add("My Computer\", tvwChild, "HKEY_DYN_DATA\", "HKEY_DYN_DATA", 1, 2)
        fred.Tag = "UF"
        Set fred = TV.Nodes.Add("HKEY_DYN_DATA\", tvwChild, "HKEY_DYN_DATA\Dummy", "Dummy", 1, 2)
        fred.Tag = "D"
    End If
    If CountAllKeys(HKEY_LOCAL_MACHINE) Then
        Set fred = TV.Nodes.Add("My Computer\", tvwChild, "HKEY_LOCAL_MACHINE\", "HKEY_LOCAL_MACHINE", 1, 2)
        fred.Tag = "UF"
        Set fred = TV.Nodes.Add("HKEY_LOCAL_MACHINE\", tvwChild, "HKEY_LOCAL_MACHINE\Dummy", "Dummy", 1, 2)
        fred.Tag = "D"
    End If
    If CountAllKeys(HKEY_PERFORMANCE_DATA) Then
        Set fred = TV.Nodes.Add("My Computer\", tvwChild, "HKEY_PERFORMANCE_DATA\", "HKEY_PERFORMANCE_DATA", 1, 2)
        fred.Tag = "UF"
        Set fred = TV.Nodes.Add("HKEY_PERFORMANCE_DATA\", tvwChild, "HKEY_PERFORMANCE_DATA\Dummy", "Dummy", 1, 2)
        fred.Tag = "D"
    End If
    If CountAllKeys(HKEY_USERS) Then
        Set fred = TV.Nodes.Add("My Computer\", tvwChild, "HKEY_USERS\", "HKEY_USERS", 1, 2)
        fred.Tag = "UF"
        Set fred = TV.Nodes.Add("HKEY_USERS\", tvwChild, "HKEY_USERS\Dummy", "Dummy", 1, 2)
        fred.Tag = "D"
    End If
    If CountAllKeys(HKEY_CURRENT_CONFIG) Then
        Set fred = TV.Nodes.Add("My Computer\", tvwChild, "HKEY_CURRENT_CONFIG\", "HKEY_CURRENT_CONFIG", 1, 2)
        fred.Tag = "UF"
        Set fred = TV.Nodes.Add("HKEY_CURRENT_CONFIG\", tvwChild, "HKEY_CURRENT_CONFIG\Dummy", "Dummy", 1, 2)
        fred.Tag = "D"
    End If
    TV.Nodes(1).Expanded = True
End Sub

Private Sub LV_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim temp As String, mvar As Variant, mByte() As Byte, z As Long
    If InStr(1, NewString, "\") Then
        MsgBox "Invalid name. Try again"
        Cancel = 1
        LV.StartLabelEdit
        Exit Sub
    End If
    If NewString = PrelabelEdit Then Exit Sub
    temp = Right(TV.SelectedItem.Key, Len(TV.SelectedItem.Key) - InStr(1, TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    Select Case LV.SelectedItem.SubItems(1)
        Case "REG_SZ", "REG_EXPAND_SZ"
            SaveSettingString Val(LV.SelectedItem.Tag), temp, NewString, LV.SelectedItem.SubItems(2)
            DeleteValue Val(LV.SelectedItem.Tag), temp, PrelabelEdit
        Case "REG_DWORD"
            SaveSettingLong Val(LV.SelectedItem.Tag), temp, NewString, Val(LV.SelectedItem.SubItems(2))
            DeleteValue Val(LV.SelectedItem.Tag), temp, PrelabelEdit
        Case "REG_BINARY"
            mvar = GetSettingByte(Val(LV.SelectedItem.Tag), temp, PrelabelEdit)
            ReDim mByte(0 To UBound(mvar))
            For z = 0 To UBound(mvar)
                mByte(z) = mvar(z)
            Next
            SaveSettingByte Val(LV.SelectedItem.Tag), temp, NewString, mByte
            DeleteValue Val(LV.SelectedItem.Tag), temp, PrelabelEdit
    End Select
    
End Sub

Private Sub LV_BeforeLabelEdit(Cancel As Integer)
    PrelabelEdit = LV.SelectedItem.Text
End Sub

Private Sub LV_DblClick()
    Dim z As Long
    If Not SlItem Is Nothing Then mnuEdit_Click 0
        
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        Set SlItem = LV.HitTest(x, y)

End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim z As Long
    If Button = 2 And TV.SelectedItem.Key <> "My Computer\" Then
        Set SlItem = LV.HitTest(x, y)
        If SlItem Is Nothing Then
            EditMenuCheck True, LV, False
            Me.PopupMenu mnuEditbase
            EditMenuCheck False
        Else
            EditMenuCheck True, LV, True
            Me.PopupMenu mnuEditbase
            EditMenuCheck False
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Thank's to owner Bobo RegEdit.", vbInformation, "al VirusScan"
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Dim temp As String, mReturn As String, strnames() As String, CurHkey As Long
    Dim mChild As Node, Subfred As Node
    temp = Right(TV.SelectedItem.Key, Len(TV.SelectedItem.Key) - InStr(1, TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    If TVPopup Then
        Select Case Index
            Case 0
            Case 2
                If MsgBox("Are you sure you wish to delete this key ?", vbQuestion + vbOKCancel) = vbCancel Then Exit Sub
                Set ValColl = New Collection
                Set ValTypeColl = New Collection
                Set SubKColl = New Collection
                ListSubVals Val(LV.SelectedItem.Tag), temp
                For z = 1 To ValColl.count
                    DeleteValue Val(LV.SelectedItem.Tag), PathOnly(ValColl(z)), FileOnly(ValColl(z))
                Next z
                If TV.SelectedItem.Children > 0 Then RenameChildNodes TV.SelectedItem, TV.SelectedItem.Key
                zx = o
                Do
                    For z = SubKColl.count To 1 Step -1
                        If DeleteKey(Val(LV.SelectedItem.Tag), SubKColl(z)) Then SubKColl.Remove z
                    Next z
                    If SubKColl.count = 0 Then Exit Do
                    zx = zx + 1
                    If zx = 500 Then Exit Do
                Loop
                DeleteKey Val(LV.SelectedItem.Tag), temp
                TV.Nodes.Remove TV.SelectedItem.Index
                TV_NodeClick TV.SelectedItem
            Case 3
                TV.StartLabelEdit
            Case 4
                If TV.SelectedItem.Image = 5 Then
                    TV.SelectedItem.Image = 1
                    TV.SelectedItem.SelectedImage = 2
                Else
                    TV.SelectedItem.Image = 5
                    TV.SelectedItem.SelectedImage = 6
                End If
            Case 5
                LockWindowUpdate TV.hWnd
                CurHkey = Val(LV.SelectedItem.Tag)
                If Not IsAKey(CurHkey, temp) Then GoTo done
                If TV.SelectedItem.Tag = "UF" Then GoTo done1
                strnames = GetAllKeys(CurHkey, temp)
                If TV.SelectedItem.Children > 0 Then
                    Set mChild = TV.SelectedItem.Child
                    For z = mChild.LastSibling.Index To mChild.FirstSibling.Index Step -1
                        If TV.Nodes(z).Parent.Key = TV.SelectedItem.Key Then TV.Nodes.Remove z
                    Next
                End If
                For z = 0 To UBound(HasSubKeys)
                    Set mChild = TV.Nodes.Add(TV.SelectedItem.Key, tvwChild, TV.SelectedItem.Key + strnames(z) + "\", strnames(z), 1, 2)
                    mChild.Tag = "UF"
                    If HasSubKeys(z) = True Then
                        Set Subfred = TV.Nodes.Add(mChild, tvwChild, mChild.Key + "Dummy", "Dummy", 1, 2)
                        Subfred.Tag = "UF"
                    End If
                Next
done1:
                LockWindowUpdate 0
                Exit Sub
                
done:
                TV.Nodes.Remove TV.SelectedItem.Index
                LockWindowUpdate 0
        End Select
    Else
        Select Case Index
            Case 0
                Select Case LV.SelectedItem.SubItems(1)
                    Case "REG_SZ", "REG_EXPAND_SZ"
                        frmStringEdit.Show , Me
                        SetWindowPos frmStringEdit.hWnd, -1, 0, 0, 0, 0, 1 Or 2
                    Case "REG_DWORD"
                        frmLongEdit.Show , Me
                        SetWindowPos frmLongEdit.hWnd, -1, 0, 0, 0, 0, 1 Or 2
                    Case "REG_BINARY"
                        frmHexEdit.Show , Me
                        SetWindowPos frmHexEdit.hWnd, -1, 0, 0, 0, 0, 1 Or 2
                End Select
            Case 2
                If MsgBox("Are you sure you wish to delete this value ?", vbQuestion + vbOKCancel) = vbCancel Then Exit Sub
                DeleteValue Val(LV.SelectedItem.Tag), temp, LV.SelectedItem.Text
                LV.ListItems.Remove LV.SelectedItem.Index
            Case 3
                LV.StartLabelEdit
        End Select
    End If
End Sub

Private Sub mnuEditbase_Click()
    Dim IsAHkey As Boolean
    Select Case TV.SelectedItem.Key
        Case "HKEY_CLASSES_ROOT\": IsAHkey = True
        Case "HKEY_CURRENT_CONFIG\": IsAHkey = True
        Case "HKEY_CURRENT_USER\": IsAHkey = True
        Case "HKEY_DYN_DATA\": IsAHkey = True
        Case "HKEY_LOCAL_MACHINE\": IsAHkey = True
        Case "HKEY_PERFORMANCE_DATA\": IsAHkey = True
        Case "HKEY_USERS\": IsAHkey = True
        Case Else: IsAHkey = False
    End Select
    If IsAHkey Then
        mnuEdit(2).Enabled = False
        mnuEdit(3).Enabled = False
    Else
        mnuEdit(2).Enabled = True
        mnuEdit(3).Enabled = True
    End If

End Sub

Private Sub mnuEditFind_Click()
    frmSearch.Show , Me
    SetWindowPos frmSearch.hWnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub

Private Sub mnuEditNewKey_Click()
    Dim temp As String, mNode As Node, lItem As ListItem, NewName As String, CurHkey As Long
    temp = Right(TV.SelectedItem.Key, Len(TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    CurHkey = Val(LV.SelectedItem.Tag)
    NewName = SafeKeyName(CurHkey, "New Key #", TV.SelectedItem)
    CreateKey CurHkey, temp + "\" + NewName
    Set mNode = TV.Nodes.Add(TV.SelectedItem.Key, tvwChild, TV.SelectedItem.Key + NewName + "\", NewName, 1, 2)
    mNode.Tag = "F"
    mNode.Selected = True
    mNode.EnsureVisible
    LV.ListItems.Clear
    Set lItem = LV.ListItems.Add(, , "(Default)", , 3)
    lItem.SubItems(1) = "REG_SZ"
    lItem.SubItems(2) = "(value not set)"
    lItem.Tag = CurHkey
    lItem.Selected = True
    TV.StartLabelEdit
End Sub

Private Sub mnuEditNewValue_Click(Index As Integer)
    Dim lItem As ListItem, CurHkey As Long, CurKeyStr As String, temp As String
    Dim b() As Byte
    ReDim b(0)
    CurHkey = Val(LV.SelectedItem.Tag)
    CurKeyStr = Right(TV.SelectedItem.Key, Len(TV.SelectedItem.Key) - InStr(1, TV.SelectedItem.Key, "\"))
    Select Case Index
        Case 0
            temp = SafeValueName(CurHkey, TV.SelectedItem, "New Value #")
            Set lItem = LV.ListItems.Add(, , temp, , 3)
            lItem.SubItems(1) = "REG_SZ"
            lItem.SubItems(2) = "(value not set)"
            lItem.Tag = CurHkey
            SaveSettingString CurHkey, CurKeyStr, temp, ""
        Case 1
            temp = SafeValueName(CurHkey, TV.SelectedItem, "New Value #")
            Set lItem = LV.ListItems.Add(, , temp, , 4)
            lItem.SubItems(1) = "REG_DWORD"
            lItem.SubItems(2) = "0x00000000 (0)"
            lItem.Tag = CurHkey
            SaveSettingLong CurHkey, CurKeyStr, temp, 0
        Case 2
            temp = SafeValueName(CurHkey, TV.SelectedItem, "New Value #")
            Set lItem = LV.ListItems.Add(, , temp, , 4)
            lItem.SubItems(1) = "REG_BINARY"
            lItem.SubItems(2) = "(zero length binary)"
            lItem.Tag = CurHkey
            SaveSettingEmptyByte CurHkey, CurKeyStr, temp
    End Select
End Sub


Private Sub mnuFavorites_Click(Index As Integer)
    Dim z As Long, bud As Collection, temp As String
    Dim zx As Long, cnt As Long
    Set bud = New Collection
    temp = mnuFavorites(Index).Tag
    z = -1
    Do Until z = 0
        z = InStr(1, temp, "\")
        If z <> 0 Then
            temp = PathOnly(temp)
            bud.Add temp + "\"
        End If
    Loop
    If bud.count > 0 Then
        cnt = 1
        For z = bud.count To 1 Step -1
            For zx = cnt To TV.Nodes.count
                temp = bud(z)
                If TV.Nodes(zx).Key = bud(z) Then
                    TV.Nodes(zx).Expanded = True
                    cnt = zx
                    Exit For
                End If
            Next
        Next
        For zx = cnt To TV.Nodes.count
            If TV.Nodes(zx).Key = mnuFavorites(Index).Tag Then
                TV.Nodes(zx).Selected = True
                TV.Nodes(zx).EnsureVisible
                NodeClick TV.Nodes(zx)
                TV.SetFocus
                Exit For
            End If
        Next
    End If
End Sub

Private Sub mnuFavoritesAdd_Click()
    AddFave TV.SelectedItem.Key, TV.SelectedItem.Text
End Sub

Private Sub mnuFavoritesOrganise_Click()
    frmOrgFaves.Show vbModal, Me
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileImport_Click()
    Dim sfile As String
    On Error GoTo woops
    With cmndlg
        .Filter = "Registry file (*.reg)|*.reg"
        .DialogTitle = "Import Registry File"
        .Flags = 5
        .ShowOpen
        sfile = .FileName
        If Len(sfile) = 0 Then Exit Sub
    End With
    ImportNode sfile
woops:

End Sub

Private Sub mnuFileSaveKey_Click()
    Dim sfile As String
    On Error GoTo woops
    If TV.SelectedItem Is Nothing Or TV.SelectedItem.Index = 1 Then
        MsgBox "You need to select a key.", vbCritical
        Exit Sub
    End If
    With cmndlg
        .Filter = "Registry file (*.reg)|*.reg"
        .DialogTitle = "Export Registry File"
        .Flags = 5 Or OFN_OVERWRITEPROMPT
        .ShowSave
        sfile = .FileName
        If Len(sfile) = 0 Then Exit Sub
    End With
    If FileExists(sfile) Then Kill sfile
    SaveKey TV.SelectedItem.Key, sfile
woops:
End Sub
Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
    DontUpdate = False
    Select Case Button.Index
        Case 1
            cboAddress.ListIndex = cboAddress.ListIndex - 1
        Case 2
            cboAddress.ListIndex = cboAddress.ListIndex + 1
    End Select
End Sub

Private Sub TV_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim temp As String, mvar As Variant, z As Long, zx As Long
    Dim CurHkey As Long, CurHKeyStr As String, CurKeyStr As String
    Dim mByte() As Byte
    If InStr(1, NewString, "\") Then
        MsgBox "Invalid name. Try again"
        Cancel = 1
        LV.StartLabelEdit
        Exit Sub
    End If
    If NewString = PrelabelEdit Then Exit Sub
    If Not IsSafeKeyName(CurHkey, NewString, TV.SelectedItem) Then
        MsgBox "A key with that name already exists. Try again"
        Cancel = 1
        LV.StartLabelEdit
        Exit Sub
    End If
    Set ValColl = New Collection
    Set ValTypeColl = New Collection
    Set SubKColl = New Collection
    CurKeyStr = Right(TV.SelectedItem.Key, Len(TV.SelectedItem.Key) - InStr(1, TV.SelectedItem.Key, "\"))
    CurHKeyStr = Left(TV.SelectedItem.Key, InStr(1, TV.SelectedItem.Key, "\") - 1)
    If Right(CurKeyStr, 1) = "\" Then CurKeyStr = Left(CurKeyStr, Len(CurKeyStr) - 1)
    CurHkey = Val(LV.SelectedItem.Tag)
    temp = PathOnly(CurKeyStr)
    If temp <> "" Then temp = temp + "\"
    CreateKey CurHkey, temp + NewString
    ListSubVals CurHkey, CurKeyStr
    For z = 1 To SubKColl.count
        CreateKey CurHkey, Replace(SubKColl(z), CurKeyStr, temp + NewString)
    Next z
    For z = 1 To ValColl.count
        Select Case ValTypeColl(z)
        Case 1
            SaveSettingString CurHkey, PathOnly(Replace(ValColl(z), CurKeyStr, temp + NewString)), FileOnly(ValColl(z)), GetSettingString(CurHkey, PathOnly(ValColl(z)), FileOnly(ValColl(z)))
        Case 3
            mvar = GetSettingByte(CurHkey, PathOnly(ValColl(z)), FileOnly(ValColl(z)))
            ReDim mByte(0 To UBound(mvar)) As Byte
            For zx = 0 To UBound(mvar)
                mByte(zx) = mvar(zx)
            Next zx
            SaveSettingByte CurHkey, PathOnly(Replace(ValColl(z), CurKeyStr, temp + NewString)), FileOnly(ValColl(z)), mByte
        Case 4
            SaveSettingLong CurHkey, PathOnly(Replace(ValColl(z), CurKeyStr, temp + NewString)), FileOnly(ValColl(z)), GetSettingLong(CurHkey, PathOnly(ValColl(z)), FileOnly(ValColl(z)))
        End Select
    Next z
    For z = 1 To ValColl.count
        DeleteValue CurHkey, PathOnly(ValColl(z)), FileOnly(ValColl(z))
    Next z
    If TV.SelectedItem.Children > 0 Then RenameChildNodes TV.SelectedItem, TV.SelectedItem.Parent.Key + NewString + "\"
    TV.SelectedItem.Key = TV.SelectedItem.Parent.Key + NewString + "\"
    zx = o
    Do
        For z = SubKColl.count To 1 Step -1
            If DeleteKey(CurHkey, SubKColl(z)) Then SubKColl.Remove z
        Next z
        If SubKColl.count = 0 Then Exit Do
        zx = zx + 1
        If zx = 500 Then Exit Do
    Loop
    DeleteKey CurHkey, CurKeyStr
End Sub

Private Sub TV_BeforeLabelEdit(Cancel As Integer)
    PrelabelEdit = TV.SelectedItem.Text
End Sub

Private Sub TV_Expand(ByVal Node As MSComctlLib.Node)
    Dim temp As String
    Dim KeyStr As String
    Dim Subfred As Node
    Dim ff() As String, z As Long
    DontUpdate = True
    If Node.Tag = "UF" Then
        Screen.MousePointer = 11
        TV.Nodes.Remove (Node.Key + "Dummy")
        temp = Left(Node.Key, InStr(1, Node.Key, "\") - 1)
        KeyStr = Right(Node.Key, Len(Node.Key) - Len(temp) - 1)
        Select Case temp
            Case "HKEY_CLASSES_ROOT"
                ff = GetAllKeys(HKEY_CLASSES_ROOT, KeyStr)
            Case "HKEY_CURRENT_CONFIG"
                ff = GetAllKeys(HKEY_CURRENT_CONFIG, KeyStr)
            Case "HKEY_CURRENT_USER"
                ff = GetAllKeys(HKEY_CURRENT_USER, KeyStr)
            Case "HKEY_DYN_DATA"
                ff = GetAllKeys(HKEY_DYN_DATA, KeyStr)
            Case "HKEY_LOCAL_MACHINE"
                ff = GetAllKeys(HKEY_LOCAL_MACHINE, KeyStr)
            Case "HKEY_PERFORMANCE_DATA"
                ff = GetAllKeys(HKEY_PERFORMANCE_DATA, KeyStr)
            Case "HKEY_USERS"
                ff = GetAllKeys(HKEY_USERS, KeyStr)
        End Select
        Node.Tag = "F"
        For z = 0 To UBound(HasSubKeys)
            Set fred = TV.Nodes.Add(Node, tvwChild, Node.Key + ff(z) + "\", ff(z), 1, 2)
            fred.Tag = "UF"
            If HasSubKeys(z) = True Then
                Set Subfred = TV.Nodes.Add(fred, tvwChild, fred.Key + "Dummy", "Dummy", 1, 2)
                Subfred.Tag = "UF"
            End If
        Next
    End If
    DontUpdate = False
    Screen.MousePointer = 0
End Sub

Private Sub TV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim simon As Node
    If Button = 2 Then
        Set simon = TV.HitTest(x, y)
        If Not simon Is Nothing And TV.SelectedItem.Key <> "My Computer\" Then
            If simon.Image = 5 Then
                mnuEdit(4).Caption = "Remove Bookmark"
            Else
                mnuEdit(4).Caption = "Bookmark"
            End If
            mnuEdit(4).Visible = True
            EditMenuCheck True, TV, True
            TVPopup = True
            Me.PopupMenu mnuEditbase
            mnuEdit(4).Visible = False
            TVPopup = False
        End If
    End If
End Sub

Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim z As Long, temp As String, cnt As Long
    Dim KeyStr As String, hKeyboy As Long
    temp = Node.Key
    DontUpdate = True
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    If temp <> cboAddress.Text Then
        cboAddress.AddItem temp
        cboAddress.ListIndex = cboAddress.ListCount - 1
        cboAddress.SelStart = Len(cboAddress.Text)
        cboAddress.SelLength = 0
    End If
    If Node.Key = "My Computer\" Then
        LV.ListItems.Clear
        Exit Sub
    End If
    TV.Scroll = True
    Node.Expanded = True
    Node.EnsureVisible
    LockWindowUpdate LV.hWnd
    LV.ListItems.Clear
    temp = Left(Node.Key, InStr(1, Node.Key, "\") - 1)
    KeyStr = Right(Node.Key, Len(Node.Key) - Len(temp) - 1)
    If Right(KeyStr, 1) = "\" Then KeyStr = Left(KeyStr, Len(KeyStr) - 1)
    Select Case temp
        Case "HKEY_CLASSES_ROOT"
            hKeyboy = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            hKeyboy = HKEY_CURRENT_CONFIG
        Case "HKEY_CURRENT_USER"
            hKeyboy = HKEY_CURRENT_USER
        Case "HKEY_DYN_DATA"
            hKeyboy = HKEY_DYN_DATA
        Case "HKEY_LOCAL_MACHINE"
            hKeyboy = HKEY_LOCAL_MACHINE
        Case "HKEY_PERFORMANCE_DATA"
            hKeyboy = HKEY_PERFORMANCE_DATA
        Case "HKEY_USERS"
            hKeyboy = HKEY_USERS
    End Select
    cnt = GetAllValues(hKeyboy, KeyStr)
    DontUpdate = False
    LockWindowUpdate 0
End Sub

Private Function ApproveName(mName As String) As Boolean
    'No strange characters allowed thanks
    Dim z As Long, Test As Boolean, qq As String
    For z = 1 To Len(mName)
        qq = Mid(mName, z, 1)
        If qq Like "[a-zA-Z0-9]" Or qq = "." Then
            Test = True
        Else
            Test = False
            Exit For
        End If
    Next
    ApproveName = Test
End Function

Public Sub RenameChildNodes(mParent As Node, NewName As String)
    'Fix keys of children whose parent has been renamed
    Dim z As Long, fred As Node
    Set ChildrenNodes = New Collection
    GetAllChildren mParent
    If ChildrenNodes.count = 0 Then Exit Sub
    For z = 1 To ChildrenNodes.count
        Set fred = TV.Nodes(ChildrenNodes(z))
        fred.Key = Replace(fred.Key, mParent.Key, NewName)
    Next
End Sub

Public Sub GetAllChildren(mParent As Node)
    Dim z As Long, fred As Node
    If mParent.Children > 0 Then
        Set fred = mParent.Child
        For z = fred.FirstSibling.Index To fred.LastSibling.Index
            ChildrenNodes.Add mParent.Child.Index
            GetAllChildren TV.Nodes(z)
        Next
    End If
End Sub

Public Sub CheckBackForward()
    Dim z As Long
    z = cboAddress.ListCount - 1
    Select Case cboAddress.ListIndex
    Case 0
        TB.Buttons(1).Enabled = False
        If z > 0 Then TB.Buttons(2).Enabled = True
    Case z
        TB.Buttons(2).Enabled = False
        If z > 0 Then TB.Buttons(1).Enabled = True
    Case Else
        TB.Buttons(1).Enabled = True
        TB.Buttons(2).Enabled = True
    End Select

End Sub

Public Sub NodeClick(mNode As Node)
    TV_NodeClick mNode
End Sub

Public Sub AddFave(mPath As String, mName As String)
    Dim cnt As Long
    If mnuFavorites(1).Visible = False Then
        mnuFavorites(1).Tag = mPath
        mnuFavorites(1).Caption = mName
        mnuFavorites(1).Visible = True
        mnuFavorites(0).Visible = True
    Else
        cnt = mnuFavorites.count
        Load mnuFavorites(cnt)
        mnuFavorites(cnt).Tag = mPath
        mnuFavorites(cnt).Caption = mName
        mnuFavorites(cnt).Visible = True
    End If
End Sub

Public Sub EditMenuCheck(IsControl As Boolean, Optional mControl As Control, Optional Htest As Boolean)
    Dim z As Long
    mnuEditNew.Enabled = True
    mnuEditNew.Visible = True
    For z = 0 To mnuEdit.count - 1
        mnuEdit(z).Enabled = True
        mnuEdit(z).Visible = True
    Next
    mnuEditNewKey.Visible = True
    mnuEditNewSP1.Visible = True
    mnuEditFind.Visible = True
    mnuEditNewKey.Visible = True
    mnuEditSP1.Visible = True
    If IsControl Then
        Select Case mControl.Name
            Case LV.Name
                If Htest Then
                    mnuEdit(4).Visible = False
                Else
                    For z = 0 To mnuEdit.count - 1
                        mnuEdit(z).Visible = False
                    Next
                    mnuEditFind.Visible = False
                    mnuEditSP1.Visible = False
                End If
                
            Case TV.Name
                If Htest Then mnuEdit(0).Visible = False
        End Select
    Else
        mnuEdit(0).Visible = False
        mnuEdit(1).Visible = False
    End If
End Sub
