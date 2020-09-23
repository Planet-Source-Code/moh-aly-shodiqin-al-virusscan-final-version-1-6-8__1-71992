VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "frmDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   4515
      Index           =   0
      Left            =   300
      ScaleHeight     =   4515
      ScaleWidth      =   6990
      TabIndex        =   1
      Top             =   525
      Width           =   6990
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5475
         TabIndex        =   31
         Top             =   3990
         Width           =   1365
      End
      Begin VB.PictureBox picIco 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   150
         ScaleHeight     =   615
         ScaleWidth      =   510
         TabIndex        =   27
         Top             =   150
         Width           =   510
      End
      Begin VB.Label lblValue 
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   150
         TabIndex        =   33
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   1350
         TabIndex        =   32
         Top             =   1800
         Width           =   2565
      End
      Begin VB.Image Image1 
         Height          =   1950
         Left            =   5100
         Picture         =   "frmDetail.frx":08CA
         Top             =   1800
         Width           =   1950
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   1350
         TabIndex        =   26
         Top             =   4140
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   1350
         TabIndex        =   25
         Top             =   3900
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   1350
         TabIndex        =   24
         Top             =   3570
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   1350
         TabIndex        =   23
         Top             =   3330
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   1350
         TabIndex        =   22
         Top             =   3090
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   1350
         TabIndex        =   21
         Top             =   2850
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   1350
         TabIndex        =   20
         Top             =   2505
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   1350
         TabIndex        =   19
         Top             =   2265
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   1350
         TabIndex        =   18
         Top             =   2025
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   16
         Left            =   1350
         TabIndex        =   17
         Top             =   1215
         Width           =   5490
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   1350
         TabIndex        =   16
         Top             =   975
         Width           =   5490
      End
      Begin VB.Label lblValue 
         Caption         =   "Date Modified"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   150
         TabIndex        =   15
         Top             =   4140
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Started On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   150
         TabIndex        =   14
         Top             =   3900
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Memory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   150
         TabIndex        =   13
         Top             =   3570
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Threads"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   150
         TabIndex        =   12
         Top             =   3330
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Priority"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   150
         TabIndex        =   11
         Top             =   3090
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Process ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   150
         TabIndex        =   10
         Top             =   2850
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   150
         TabIndex        =   9
         Top             =   2505
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   150
         TabIndex        =   8
         Top             =   2265
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   150
         TabIndex        =   7
         Top             =   2025
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Directory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   150
         TabIndex        =   6
         Top             =   1215
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   975
         Width           =   1140
      End
      Begin VB.Label lblValue 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   975
         TabIndex        =   4
         Top             =   630
         Width           =   5865
      End
      Begin VB.Label lblValue 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   975
         TabIndex        =   3
         Top             =   390
         Width           =   5865
      End
      Begin VB.Label lblValue 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   975
         TabIndex        =   2
         Top             =   150
         Width           =   5865
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   4515
      Index           =   1
      Left            =   300
      ScaleHeight     =   4515
      ScaleWidth      =   6990
      TabIndex        =   28
      Top             =   525
      Width           =   6990
      Begin ComctlLib.ListView lvwModDetail 
         Height          =   4215
         Left            =   150
         TabIndex        =   29
         Top             =   150
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   7435
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ilsDetail"
         SmallIcons      =   "ilsDetail"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Modules Used"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Opened By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File Type"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Directory"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblNotFound 
         BackStyle       =   0  'Transparent
         Caption         =   "Cannot open file. Module not found."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   585
         TabIndex        =   30
         Top             =   4125
         Width           =   3015
      End
      Begin VB.Image imgEmpty 
         Height          =   240
         Left            =   225
         Picture         =   "frmDetail.frx":5A99
         Top             =   4125
         Width           =   240
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   4515
      Index           =   2
      Left            =   300
      ScaleHeight     =   4515
      ScaleWidth      =   6990
      TabIndex        =   35
      Top             =   525
      Width           =   6990
      Begin ComctlLib.ListView lvwVersion 
         Height          =   4215
         Left            =   150
         TabIndex        =   36
         Top             =   150
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   7435
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Value"
            Object.Width           =   6068
         EndProperty
      End
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   225
      TabIndex        =   37
      Top             =   6300
      Width           =   2865
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6000
      TabIndex        =   34
      Top             =   5250
      Width           =   1365
   End
   Begin ComctlLib.TabStrip tabDetail 
      Height          =   4965
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   8758
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "File Version"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ilsDetail 
      Left            =   3150
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As ESHGetFileInfoFlagConstants) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Private Type SHFILEINFO
    hIcon           As Long ' : icon
    iIcon           As Long ' : icondex
    dwAttributes    As Long ' : SFGAO_ flags
    szDisplayName   As String * MAX_PATH ' : display name (or path)
    szTypeName      As String * 80 ' : type name
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private shinfo As SHFILEINFO

Private Sub MakeInfo()
    On Error Resume Next
    Dim sFileName As String
    Dim hVer As VERHEADER
    Dim hIcoExt As Long, hIcoDraw As Long
    Dim fso As New FileSystemObject
    Dim FileInfo As file
    Dim sFile As String
    
    picIco.Cls
    sFileName = frmProcess.lvwProcess.SelectedItem.Tag
    Set FileInfo = fso.GetFile(sFileName)
    GetVerHeader sFileName, hVer
   
    If sFile <> sFileName Then
        lblValue(0) = hVer.FileDescription
        lblValue(1) = hVer.CompanyName
        lblValue(2) = hVer.FileVersion
        lblValue(8) = ": " & frmProcess.lvwProcess.SelectedItem.SubItems(1)
        lblValue(15) = ": " & file_getName(sFileName)
        lblValue(16) = ": " & file_getPath(sFileName)
        lblValue(17) = ": " & file_getType(sFileName)
        lblValue(18) = ": " & Format(FileLen(sFileName) / 1024, "###,####") & " KB"
        lblValue(19) = ": " & GetAttribute(sFileName)
        lblValue(21) = ": " & frmProcess.lvwProcess.SelectedItem.SubItems(3)
        lblValue(22) = ": " & frmProcess.lvwProcess.SelectedItem.SubItems(8)
        lblValue(23) = ": " & frmProcess.lvwProcess.SelectedItem.SubItems(4)
        lblValue(24) = ": " & frmProcess.lvwProcess.SelectedItem.SubItems(6)
        lblValue(25) = ": " & FileInfo.DateLastAccessed
        lblValue(26) = ": " & FileInfo.DateLastModified
'        hIcoExt = ExtractIcon(Me.hwnd, sFilename, 0)
'        hIcoDraw = DrawIcon(picIco.hdc, 0, 0, hIcoExt)
        RetrieveIcon sFileName, picIco, ricnLarge
'        picIco.Picture = GetIco.Icon(sFilename, LargeIcon)
    End If
End Sub

Private Sub cmdClose_Click()
    frmProcess.show
    Unload Me
End Sub

Private Sub cmdProperties_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 1 To frmProcess.lvwProcess.ListItems.count
      If frmProcess.lvwProcess.ListItems(i).Selected Then
         ShowProps frmProcess.lvwProcess.SelectedItem.Tag, Me.hWnd
      End If
    Next i
End Sub

Private Sub Form_Load()
    Me.Caption = "Details"
'    lvwStyleProcess lvwModDetail

    With List1
        .AddItem "Comment"
        .AddItem "Company Name"
        .AddItem "File Description"
        .AddItem "File Version"
        .AddItem "Internal Name"
        .AddItem "Legal Copyright"
        .AddItem "Legal Trademark"
        .AddItem "Original Filename"
        .AddItem "Product Name"
        .AddItem "Product Version"
        .AddItem "Private Build"
        .AddItem "Special Build"
    End With
    
    SetFlatHeaders lvwVersion.hWnd
    SetFlatHeaders lvwModDetail.hWnd
    
    MakeInfo
    GetModuleProcessID frmProcess.lvwProcess, 3, lvwModDetail, ilsDetail
    FillVersion

    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.hWnd, True
    Else
        AlwaysOnTop Me.hWnd, False
    End If
'    MemoryInfo lblInfo(0), lblInfo(1), lblInfo(2), lblInfo(3), lblInfo(4), lblInfo(5), lblInfo(6), lblInfo(7), lblInfo(8), lblInfo(9) ', ProgMemUsed ', Me.StatusBar1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdClose_Click
End Sub

Private Sub tabDetail_Click()
    Dim pic As PictureBox
    For Each pic In picDetail
        pic.Visible = (pic.Index = tabDetail.SelectedItem.Index - 1)
    Next
    Select Case tabDetail.SelectedItem.Caption
        Case "Advanced"
            CheckModule
        Case "General"
'            MakeInfo
    End Select
End Sub

Private Sub CheckModule()
    Dim sFileName As String
    
    sFileName = frmProcess.lvwProcess.SelectedItem.SubItems(2)
    If lvwModDetail.ListItems.count = 0 Then
        lvwModDetail.Visible = False
    Else
        lvwModDetail.Visible = True
    End If
End Sub

Sub RetrieveIcon(fName As String, DC As PictureBox, icnSize As IconRetrieve)
    Dim hImgLarge As Long  'the handle to the system image list
        
    If icnSize = ricnLarge Then
        hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    End If
End Sub

Sub FillVersion()
    Dim Cnt As Long
    Dim itmX As ListItem
    Dim Filename As String
    Dim hVer As VERHEADER
    
    Filename = frmProcess.lvwProcess.SelectedItem.Tag
    GetVerHeader Filename, hVer
    For Cnt = 0 To List1.ListCount - 1
        Set itmX = lvwVersion.ListItems.Add(, , List1.List(Cnt))
            itmX.SubItems(1) = hVer.Comments
      
        If Cnt = 1 Then
            itmX.SubItems(1) = hVer.CompanyName
        ElseIf Cnt = 2 Then
            itmX.SubItems(1) = hVer.FileDescription
        ElseIf Cnt = 3 Then
            itmX.SubItems(1) = hVer.FileVersion
        ElseIf Cnt = 4 Then
            itmX.SubItems(1) = hVer.InternalName
        ElseIf Cnt = 5 Then
            itmX.SubItems(1) = hVer.LegalCopyright
        ElseIf Cnt = 6 Then
            itmX.SubItems(1) = hVer.LegalTradeMarks
        ElseIf Cnt = 7 Then
            itmX.SubItems(1) = hVer.OrigionalFileName
        ElseIf Cnt = 8 Then
            itmX.SubItems(1) = hVer.ProductName
        ElseIf Cnt = 9 Then
            itmX.SubItems(1) = hVer.ProductVersion
        ElseIf Cnt = 10 Then
            itmX.SubItems(1) = hVer.PrivateBuild
        ElseIf Cnt = 11 Then
            itmX.SubItems(1) = hVer.PrivateBuild
        End If
    Next
    For Cnt = 1 To lvwVersion.ColumnHeaders.count
        LV_AutoSizeColumn lvwVersion, lvwVersion.ColumnHeaders.Item(Cnt)
    Next Cnt
End Sub

