VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmScanSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   Icon            =   "frmScanSample.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvwProcessList 
      Height          =   1515
      Left            =   1950
      TabIndex        =   24
      Tag             =   "Right click to scan this process..."
      Top             =   450
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   2672
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ilsDetected"
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Image Name"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Directory"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Choose from Process List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   1800
      TabIndex        =   28
      Top             =   150
      Width           =   2940
   End
   Begin VB.ListBox lstDrive 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   375
      Style           =   1  'Checkbox
      TabIndex        =   0
      Tag             =   "Select local drive..."
      Top             =   450
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Caption         =   "Drive"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   225
      TabIndex        =   23
      Top             =   150
      Width           =   1440
   End
   Begin alVirusScan.AdvProgressBar pbScanning 
      Height          =   165
      Left            =   1275
      TabIndex        =   22
      ToolTipText     =   "Scanning files..."
      Top             =   6593
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   291
      BarColor1       =   -2147483634
   End
   Begin ComctlLib.ListView lvwDetected 
      Height          =   1365
      Left            =   225
      TabIndex        =   21
      Tag             =   "Scan system detected"
      Top             =   4050
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   2408
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ilsDetected"
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Directory"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1536
      EndProperty
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Enabled         =   0   'False
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
      Left            =   5175
      TabIndex        =   13
      Tag             =   "Start scan..."
      Top             =   2250
      Width           =   1365
   End
   Begin ComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   6915
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8952
            MinWidth        =   6068
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5424
            Object.Tag             =   ""
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
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1665
      Left            =   4950
      ScaleHeight     =   1665
      ScaleWidth      =   2940
      TabIndex        =   3
      Top             =   375
      Width           =   2940
      Begin VB.TextBox txtValue 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1380
         Width           =   1560
      End
      Begin VB.TextBox txtValue 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1140
         Width           =   1560
      End
      Begin VB.TextBox txtValue 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   900
         Width           =   1560
      End
      Begin VB.PictureBox picSample 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   75
         ScaleHeight     =   540
         ScaleWidth      =   615
         TabIndex        =   4
         Top             =   75
         Width           =   615
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   5
         Left            =   75
         TabIndex        =   7
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   75
         TabIndex        =   6
         Top             =   1140
         Width           =   1125
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   3
         Left            =   75
         TabIndex        =   5
         Top             =   900
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
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
      Left            =   6600
      TabIndex        =   2
      Tag             =   "Click browse to choose virus sample from disk!"
      Top             =   2250
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Info"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   4875
      TabIndex        =   1
      Top             =   150
      Width           =   3090
   End
   Begin MSComCtl2.Animation aniScan 
      Height          =   690
      Left            =   225
      TabIndex        =   19
      Top             =   6075
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1217
      _Version        =   393216
      FullWidth       =   51
      FullHeight      =   46
   End
   Begin ComctlLib.ListView lvwProcess 
      Height          =   1140
      Left            =   225
      TabIndex        =   29
      Tag             =   "Scan process detected"
      Top             =   2775
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   2011
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ilsProcess"
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Image Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Directory"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Process ID"
         Object.Width           =   1536
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   840
      Left            =   5175
      TabIndex        =   30
      Top             =   4275
      Visible         =   0   'False
      Width           =   2640
      Begin VB.Timer tmrScan 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1500
         Top             =   225
      End
      Begin VB.Timer Timer2 
         Interval        =   10000
         Left            =   1050
         Top             =   225
      End
      Begin VB.PictureBox picProcess 
         AutoRedraw      =   -1  'True
         Height          =   315
         Left            =   75
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   31
         Top             =   225
         Visible         =   0   'False
         Width           =   315
      End
      Begin ComctlLib.ImageList ilsDetected 
         Left            =   1950
         Top             =   225
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
      End
      Begin ComctlLib.ImageList ilsProcess 
         Left            =   450
         Top             =   225
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   240
      Left            =   2475
      TabIndex        =   33
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   240
      Left            =   2100
      TabIndex        =   32
      Top             =   1500
      Width           =   990
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1275
      TabIndex        =   20
      Top             =   6555
      Width           =   6690
   End
   Begin VB.Label lblScan 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1215
      TabIndex        =   18
      Top             =   2505
      Width           =   45
   End
   Begin VB.Label lblScan 
      Caption         =   "Virus Found"
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
      Left            =   225
      TabIndex        =   17
      Top             =   2490
      Width           =   990
   End
   Begin VB.Label lblScan 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1215
      TabIndex        =   16
      Top             =   2250
      Width           =   45
   End
   Begin VB.Label lblScan 
      Caption         =   "File Scanned"
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
      Index           =   0
      Left            =   225
      TabIndex        =   15
      Top             =   2250
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   240
      Left            =   2250
      TabIndex        =   14
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label lblPath 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1275
      TabIndex        =   11
      Top             =   5790
      Width           =   6690
   End
   Begin VB.Label lblFilename 
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
      Left            =   1275
      TabIndex        =   10
      Top             =   5550
      Width           =   6690
   End
   Begin VB.Label Label1 
      Caption         =   "Scanning In"
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
      Left            =   225
      TabIndex        =   9
      Top             =   5790
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "File"
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
      Left            =   225
      TabIndex        =   8
      Top             =   5550
      Width           =   990
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuScan 
         Caption         =   "Scan..."
         Index           =   0
      End
      Begin VB.Menu mnuScan 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Terminate Process..."
         Index           =   2
      End
   End
   Begin VB.Menu mnuA 
      Caption         =   "Action"
      Visible         =   0   'False
      Begin VB.Menu mnuAction 
         Caption         =   "Delete..."
         Index           =   0
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Quarantine..."
         Index           =   1
      End
      Begin VB.Menu mnuAction 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Select All..."
         Index           =   3
      End
   End
   Begin VB.Menu mnuAP 
      Caption         =   "Action Process"
      Visible         =   0   'False
      Begin VB.Menu mnuProcess 
         Caption         =   "Terminate Process..."
         Index           =   0
      End
      Begin VB.Menu mnuProcess 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuProcess 
         Caption         =   "Quarantine Process..."
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmScanSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module     : Scan With Virus Sample®...
'           : ver 1.2
'           : 1 Maret 2009 10:22 PM
'           : Moh Aly Shodiqin
'           : Update 23 Maret 2009 11:21 AM
'           : ver 1.2
'           : Update 28 Maret 2009 7:45 PM
'--------------------------------------------------------------
'License : Freeware n Open Source
'--------------------------------------------------------------
'Original Author : Moh Aly Shodiqin
'Release date    : 9 April 2009 4:53 PM
'Author Contact  : felix_progressif@yahoo.com /
'                : http://fi5ly.blogspot.com
'--------------------------------------------------------------
'Great thanks to    : - Allah S.W.T
'                   : - Nabi Muhammad S.A.W
'                   : - My Parent
'                   : - My Soul
'                   : - www.planetsourcecode.com
'-------------------------------------------------------
'Don't forget to vote me
'--------------------------------------------------------------
Option Explicit

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As ESHGetFileInfoFlagConstants) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type SHFILEINFO
    hIcon           As Long ' : icon
    iIcon           As Long ' : icondex
    dwAttributes    As Long ' : SFGAO_ flags
    szDisplayName   As String * MAX_PATH ' : display name (or path)
    szTypeName      As String * 80 ' : type name
End Type

Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type

Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Dim WithEvents Sf As cFileSearch
Attribute Sf.VB_VarHelpID = -1
'Dim WithEvents Engine32 As cEngine32
Dim FileOnScan As Double
Dim FileViruses As Double
Dim FailedFile As Double
Dim CleanedFile As Double

Dim Scanning As Boolean
Dim Stopped As Boolean

Dim m_CRC As String
Dim size As String
Dim X As Integer
Private shinfo As SHFILEINFO

Sub LoadDrive()
'    On Error Resume Next
'    Dim LDs As Long, Cnt As Long, sDrives As String
'    LDs = GetLogicalDrives
'    For Cnt = 0 To 25
'        If (LDs And 2 ^ Cnt) <> 0 Then
'            Dim Serial As Long, VName As String, FSName As String, ndrvName As String
'            VName = String$(255, Chr$(0))
'            FSName = String$(255, Chr$(0))
'            GetVolumeInformation Chr$(65 + Cnt) & ":\", VName, 255, Serial, 0, 0, FSName, 255
'            VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
'            FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
'            ndrvName = ""
'            If VName = "" Then
'                Select Case GetTipeDrive(Chr$(65 + Cnt) & ":\")
'                       Case 2: ndrvName = "3½ Floppy (" & Chr$(65 + Cnt) & ":)"
'                       Case 5: ndrvName = "CDROM (" & Chr$(65 + Cnt) & ":)"
'                       Case Else: ndrvName = "Unknown (" & Chr$(65 + Cnt) & ":)"
'                End Select
'                If ndrvName <> "" Then
'                    lstDrive.AddItem Chr$(65 + Cnt) & ":\" & vbTab & ndrvName
'
'                End If
'            Else
'                ndrvName = VName & " (" & Chr$(65 + Cnt) & ":)"
'                lstDrive.AddItem Chr$(65 + Cnt) & ":\" & vbTab & ndrvName
'                'Chr$(65 + Cnt) & ":\", ndrvName)
'            End If
'        End If
'    Next Cnt

    On Error Resume Next
    Dim LDs As Long, Cnt As Long, sDrives As String
    LDs = GetLogicalDrives
    For Cnt = 0 To 25
        If (LDs And 2 ^ Cnt) <> 0 Then
            lstDrive.AddItem Chr(Cnt + 65) & ":\"
        End If
    Next Cnt
    
    Dim i As Integer
    For i = 0 To lstDrive.ListCount - 1
        lstDrive.Selected(i) = True
    Next
End Sub

Private Sub cmdBrowse_Click()
    On Error Resume Next
    Dim fullname As String
    Dim ver As VERHEADER
    
'    ResetMe
    fullname = ShowSaveSample(Me.hWnd, , True)
    If Trim(fullname) <> "" Then
        picSample.Cls
        GetVerHeader fullname, ver
        Label3 = fullname
        txtValue(0) = ": " & FileLen(fullname) \ 1024 & " KB"
        txtValue(1) = ": " & ver.CompanyName
        txtValue(2) = ": " & ver.InternalName
        RetrieveIcon fullname, picSample, ricnLarge
'        sb.Panels(2).Text = m_CRC.FileChecksum(Label3.Caption)
        Label4 = "If you're not sure the file is not virus click browse to choose from disk!"
        cmdScan.Enabled = True
    End If
    lvwProcess.ListItems.Clear
    lvwDetected.ListItems.Clear
    Call ProcessListSample(lvwProcessList, ilsDetected)
End Sub

Private Sub cmdBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels(1).Text = cmdBrowse.Tag
End Sub

Private Sub cmdScan_Click()
    If Label3.Caption = "" Then
        MsgBox "Please select virus sample to scan your disk!", vbExclamation, "al VirusScan"
        Exit Sub
    End If
  
    If Scanning = False Then
        lvwDetected.Enabled = False
        lvwDetected.ListItems.Clear
        
        ResetMe
        
        Scanning = True
        Stopped = False
        mnuHelp.Enabled = False
        cmdBrowse.Enabled = False
        m_CRC = GetChecksum(Label3.Caption)
'        Debug.Print m_CRC
'        cmdScan.SetFocus
        Label4.Visible = True
        cmdScan.Enabled = True
        cmdScan.Caption = "Stop"
        sb.Panels(2).Text = "Scanning files..."
        lstDrive.Enabled = False
        lvwProcessList.Enabled = False
        lvwProcess.Enabled = False
        Timer2.Enabled = False
'        txtValue(0).Enabled = False
'        txtValue(1).Enabled = False
'        txtValue(2).Enabled = False
        
        ' Process---------------------------------------------
        Dim enumerasi As Long
        Dim uProcess As PROCESSENTRY32
        Dim snap As Long
        Dim ID As Long
        Dim lv As ListItem
        
        snap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0)
        uProcess.dwSize = Len(uProcess)
        enumerasi = Process32First(snap, uProcess)
        
        ControlListView lvwProcess, ilsProcess, picProcess
        While enumerasi <> 0
            ID = uProcess.th32ProcessID
            If FileExists(ProcessPathByPID(uProcess.th32ProcessID)) = True Then
                If m_CRC = GetChecksum(ProcessPathByPID(uProcess.th32ProcessID)) Then
                    FileViruses = FileViruses + 1
                    LogScan "Virus Found On Process " & lvwProcessList.SelectedItem.SubItems(1)
'                    ilsDetected.ListImages.Add , Filename, GetIco.Icon(Filename, SmallIcon)
                    Set lv = lvwProcess.ListItems.Add(, , uProcess.szExeFile)
                    lv.SubItems(1) = ProcessPathByPID(uProcess.th32ProcessID)
                    lv.SubItems(2) = uProcess.th32ProcessID  'FileLen(uProcess.szExeFile) \ 1024 & " KB"
                    lvwProcess.ListItems(lvwProcess.ListItems.count).Selected = True
                End If
            End If
            enumerasi = Process32Next(snap, uProcess)
        Wend
        CloseHandle snap
        If lvwProcess.ListItems.count <> 0 Then GetIcons lvwProcess, ilsProcess, picProcess
        '---------------------------------------------------------------
        Dim i As Integer
        For i = 0 To lstDrive.ListCount - 1
            If Stopped = False Then
                If lstDrive.Selected(i) Then
                    aniScan.Stop
                    aniScan.Play
                    aniScan.Visible = True
                    If tmrScan.Enabled = False Then
                        tmrScan.Enabled = True
                        pbScanning.Visible = True
                    End If
                    Sf.DoCmdSearchFile lstDrive.List(i), True
                End If
            End If
        Next i
        '---------------------------------------------------------------
        lblFilename = ""
        lblPath = ""
        sb.Panels(2).Text = "Completed."
        tmrScan.Enabled = False
        pbScanning.Value = 0
        pbScanning.Visible = False
        lstDrive.Enabled = True
        lvwProcessList.Enabled = True
        cmdScan.Caption = "Scan"
        cmdScan.Enabled = False
        lvwDetected.Enabled = True
        lvwProcess.Enabled = True
        mnuHelp.Enabled = True
        cmdBrowse.Enabled = True
        Label4.Visible = False
        Timer2.Enabled = True
'        txtValue(0).Enabled = True
'        txtValue(1).Enabled = True
'        txtValue(2).Enabled = True
        aniScan.Stop
        aniScan.Visible = False
        Scanning = False
        Stopped = False
    Else
        Stopped = True
        Sf.StopSearch = True
    End If
End Sub

Private Sub cmdScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdScan.Caption = "Scan" Then
        sb.Panels(1).Text = cmdScan.Tag
    Else
        sb.Panels(1).Text = "Stop current scanning..."
       End If
End Sub

Private Sub Form_Load()
    Me.Caption = "al VirusScan Scan With Virus Sample " & vScanWithVirusSample
    Set Sf = New cFileSearch
    LoadDrive
    LoadAnim
    lvwStyle lvwDetected
    lvwStyle lvwProcess
    lvwStyle lvwProcessList
    SetFlatHeaders lvwDetected.hWnd
    SetFlatHeaders lvwProcess.hWnd
    ProcessListSample lvwProcessList, ilsDetected
    
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.hWnd, True
    Else
        AlwaysOnTop Me.hWnd, False
    End If

'    Set Engine32 = New cEngine32
'    Engine32.ClassIDApartement = Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(1) & Chr(255)

    pbScanning.Visible = False
    pbScanning.Style = DoubleColor
'    cmdBrowse.ToolTipText = "Click browse to choose virus sample from disk!"
    sb.Panels(1).Text = "Ready."
'    sb.Panels(2).Text = Copyright
End Sub

Sub RetrieveIcon(fName As String, DC As PictureBox, icnSize As IconRetrieve)
    Dim hImgLarge As Long  'the handle to the system image list
    
    If icnSize = ricnLarge Then
        hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    End If
End Sub

Sub ResetMe()
    FileOnScan = 0
    FileViruses = 0
    CleanedFile = 0
    FailedFile = 0
    lblScan(1) = ": " & FileOnScan & " file."
    Label4 = ""
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set Sf = Nothing
''    Set Engine32 = Nothing
'    Unload Me
    If cmdScan.Enabled = False Then
        Set Sf = Nothing
        Unload Me
    Else
        If MsgBox("Abort the current process...", vbQuestion + vbYesNo, "al VirusScan") = vbYes Then
            Set Sf = Nothing
            Unload Me
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub lstDrive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels(1).Text = lstDrive.Tag
End Sub

Private Sub lvwDetected_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels(1).Text = lvwDetected.Tag
End Sub

Private Sub lvwDetected_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lvwDetected.ListItems.count > 0 Then
            Label6.Caption = lvwDetected.SelectedItem.SubItems(1)
            PopupMenu mnuA
        End If
    End If
End Sub

Private Sub lvwProcess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels(1).Text = lvwProcess.Tag
End Sub

Private Sub lvwProcess_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lvwProcess.ListItems.count > 0 Then
            Label5.Caption = lvwProcess.SelectedItem.SubItems(1)
            PopupMenu mnuAP
        End If
    End If
End Sub

Private Sub lvwProcessList_Click()
    On Error Resume Next
    Dim fullname As String
    Dim ver As VERHEADER
    
    fullname = lvwProcessList.SelectedItem.SubItems(1)
    Label4.Caption = ""
    If Trim(fullname) <> "" Then
        picSample.Cls
        GetVerHeader fullname, ver
        Label3 = fullname
        txtValue(0) = ": " & FileLen(fullname) \ 1024 & " KB"
        txtValue(1) = ": " & ver.CompanyName
        txtValue(2) = ": " & ver.InternalName
        RetrieveIcon fullname, picSample, ricnLarge
        Label4 = "Choose file from process list be sure the file is virus / trojan. Don't make mistake!"
    End If
End Sub

Private Sub lvwProcessList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels(1).Text = lvwProcessList.Tag
End Sub

Private Sub lvwProcessList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lvwProcessList.ListItems.count > 0 Then
            Call lvwProcessList_Click
            Label3.Caption = lvwProcessList.SelectedItem.SubItems(1)
            PopupMenu mnuFile
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.show
End Sub

Private Sub mnuAction_Click(Index As Integer)
    Select Case Index
        Case 0: CleanVirus
        Case 1: Quarantine
        Case 3: All
    End Select
End Sub

Private Sub mnuProcess_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim i As Integer
            Dim Pesan As String, strFile As String
            Dim lExitCode As Long
            
            Pesan = "WARNING: Terminating a process can cause undesired" & vbCrLf & _
                    "results including loss of data and system instability. The" & vbCrLf & _
                    "process will not be given the chance to save its state or" & vbCrLf & _
                    "data before it is terminated. Are you sure you want to" & vbCrLf & _
                    "terminate the process?"
            If MsgBox(Pesan, vbYesNo + 48, "Process Manager Warning" & Chr(0)) = vbYes Then
                For i = 1 To lvwProcess.ListItems.count
                    If lvwProcess.ListItems(i).Selected Then
                        lExitCode = KillProcessById(CLng(lvwProcess.ListItems(i).SubItems(2)))
                        LogScan "Terminating Process... " & lvwProcess.SelectedItem.SubItems(1) & vbTab & "successfully."
                        lvwProcess.ListItems.Remove (i)
                        If lExitCode = 0 Then
                            MsgBox "Cannot terminate this process.", vbExclamation, "Unable To Terminate Process"
                            LogScan "Unable To Terminate Process... " & lvwProcess.SelectedItem.SubItems(1) & vbTab & "access is denied."
                        End If
                    End If
                Next i
            End If
            Call ProcessListSample(lvwProcessList, ilsDetected)
        Case 2
            Call QuaProcess
            Call ProcessListSample(lvwProcessList, ilsDetected)
    End Select
End Sub

Private Sub mnuScan_Click(Index As Integer)
    Select Case Index
        Case 0
'            cmdScan.Enabled = True
            Call cmdScan_Click
        Case 2
            Dim i As Integer
            Dim Pesan As String, strFile As String
            Dim lExitCode As Long
            
            Pesan = "WARNING: Terminating a process can cause undesired" & vbCrLf & _
                    "results including loss of data and system instability. The" & vbCrLf & _
                    "process will not be given the chance to save its state or" & vbCrLf & _
                    "data before it is terminated. Are you sure you want to" & vbCrLf & _
                    "terminate the process?"
            If MsgBox(Pesan, vbYesNo + 48, "Process Manager Warning" & Chr(0)) = vbYes Then
                For i = 1 To lvwProcessList.ListItems.count
                    If lvwProcessList.ListItems(i).Selected Then
                        lExitCode = KillProcessById(CLng(lvwProcessList.ListItems(i).Tag))
                        LogScan "Terminating Process... " & lvwProcessList.SelectedItem.SubItems(1) & vbTab & "successfully."
                        If lExitCode = 0 Then
                            MsgBox "Cannot terminate this process.", vbExclamation, "Unable To Terminate Process"
                            LogScan "Unable To Terminate Process... " & lvwProcessList.SelectedItem.SubItems(1) & vbTab & "access is denied."
                        End If
                    End If
                Next i
            End If
            Call ProcessListSample(lvwProcessList, ilsDetected)
    End Select
End Sub

Private Sub Sf_onSearch(nFileName As String, nFileInfo As cFileInfo)
    If Trim(nFileInfo.Filename) <> "" Then
        lblFilename = nFileInfo.Filename
        lblPath = nFileInfo.FilePath
        FileOnScan = FileOnScan + 1
        RunCMD nFileName 'Label3.Caption 'm_CRC.FileChecksum(Label3.Caption)
''    Else
''        lblFilename = "Scanning..."
    End If
    lblScan(1) = ": " & FileOnScan & " files."
'    If size = DungLuong(nFileInfo.FilePath & nFileInfo.Filename) Then
'        If m_CRC = GetChecksum(nFileInfo.FilePath & nFileInfo.Filename) Then
'            Dim lv As ListItem
''            Dim pf As String
''            pf = Filename
'            Set lv = lvwDetected.ListItems.Add(, , nFileInfo.Filename)
'            lv.SubItems(1) = nFileInfo.FilePath
'            lv.SubItems(2) = FileLen(nFileInfo.FilePath & nFileInfo.Filename) \ 1024 & " KB"
'        End If
'    End If
End Sub

Sub RunCMD(Filename As String)
    On Error Resume Next
    m_CRC = GetChecksum(Filename)
    If Trim(m_CRC) <> "" Then
        Dim lv As ListItem
        Dim pf As String
        pf = Filename
       
        ' System---------------------------------------------
        If m_CRC = GetChecksum(Label3.Caption) Then
            FileViruses = FileViruses + 1
            LogScan "Virus Found On Scan With Virus Sample " & vbTab & Label3.Caption
            With lvwDetected
                ilsDetected.ListImages.Add , Filename, GetIco.Icon(Filename, SmallIcon)
                Set lv = lvwDetected.ListItems.Add(, , file_getName(pf), , Filename)
                lv.SubItems(1) = pf
                lv.SubItems(2) = FileLen(pf) \ 1024 & " KB"
                lvwDetected.ListItems(lvwDetected.ListItems.count).Selected = True
            End With
        End If
    End If
    lblScan(3) = ": " & FileViruses
End Sub

Function IsFileExist(sPath As String) As Boolean
    If PathFileExists(sPath) = 1 And PathIsDirectory(sPath) = 0 Then
        IsFileExist = True
    Else
        IsFileExist = False
    End If
End Function

Public Function Cure(path As String)
    KillVirusNow (path)
'    SetFileAttributes path, FILE_ATTRIBUTE_NORMAL
'    Engine32.KillFile (path)
    CleanedFile = CleanedFile + 1
End Function

Public Function KillVirusNow(ByVal sPathDel As String) As Long
    On Error Resume Next
    SetAttr sPathDel, vbNormal
    TerminateExeName sPathDel
    LogScan "Virus Deleted... " & vbTab & sPathDel
    Kill (sPathDel)
    DoEvents
    If IsFileExist(sPathDel) = True Then
        LogScan "Unable to delete file... " & vbTab & sPathDel
        MsgBox "VirusScan cannot delete this file. Maybe this file is running in system process." & _
                vbCrLf & "You can terminate or quarantine the former system process, before delete the file.", vbExclamation + vbOKOnly, "Warning"
        Exit Function
    End If
End Function

Private Sub CleanVirus()
    On Error Resume Next
    Dim sClean As String
    Dim i As Long, lRet As Long
    
    With lvwDetected.ListItems
        For i = 1 To .count
            If .Item(i).Selected = True Then
                sClean = .Item(i).SubItems(1)
                VirusAlert
                SetFileAttributes sClean, FILE_ATTRIBUTE_NORMAL
                Sleep 200
                DoEvents
                Cure (sClean)
                If lRet <> 0 Then
                    .Item(i).Selected = False
                End If
                .Item(i).Selected = False
                CleanVirus
                .Remove (i)
                Label3.Caption = ""
                Label5.Caption = ""
                Label6.Caption = ""
                picSample.Cls
                txtValue(0) = ""
                txtValue(1) = ""
                txtValue(2) = ""
                Exit Sub
            End If
        Next i
    End With
    Call ProcessListSample(lvwProcessList, ilsDetected)
End Sub

Private Sub LoadAnim()
    On Error GoTo Salah
    Dim H As String
    H = App.path & "\Data\Findfile.avi"
    aniScan.Visible = False
    aniScan.Open H
Salah:
End Sub

Sub All()
    Dim i As Integer
    With lvwDetected.ListItems
        For i = 1 To .count
            .Item(i).Selected = True
        Next i
    End With
End Sub

Private Sub Timer2_Timer()
    If Not mnuFile.Visible Then
        ProcessListSample lvwProcessList, ilsDetected
        sb.Panels(1).Text = ""
        sb.Panels(2).Text = ""
    End If
End Sub

Private Sub tmrScan_Timer()
    X = X + 2
    If X > 100 Then X = 0

    pbScanning.Value = X
End Sub

Function QuarantineFile(Filename As String, Optional virus As String = "") As Boolean
'    If var_ClassID = False Then Exit Function
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
        
        LogScan "Virus Quarantined... " & vbTab & Filename
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
    LogScan "Unable to quarantine file... " & vbTab & Filename
    MsgBox "VirusScan cannot quarantine this file. Maybe this file is running in system process." & _
            vbCrLf & "You can terminate or quarantine the former system process, before quarantine the file.", vbExclamation + vbOKOnly, "Warning"
'    RaiseEvent onEngineError("Virus can't move to quarantine directory - " & Filename)
    Close #1
End Function

Private Sub Quarantine()
    On Error Resume Next
    Dim sNama As String
    Dim i As Long, lRet As Long
    Dim sFile As String
    Dim j As Long, lExitCode As Long
    
    With lvwDetected.ListItems
        For i = 1 To .count
            If .Item(i).Selected = True Then
                sFile = .Item(i).SubItems(1)
                QuarantineShow
                DoEvents
                VirusAlert
                TerminateExeName sFile
'                For j = 1 To lvwDetected.ListItems.count
'                    If lvwDetected.ListItems(i).Selected Then
'                        lExitCode = KillProcessById(CLng(lvwDetected.ListItems(j).SubItems(1)))
'                        If lExitCode = 0 Then MsgBox "Cannot terminate this process.", vbExclamation, "Unable To Terminate Process"
'                        Exit For
'                    End If
'                Next j
                sNama = QuarantineFile(sFile)
                LogScan "Quarantine File... " & vbTab & sFile
                If lRet <> 0 Then
                    .Item(i).Selected = False
                End If
                .Item(i).Selected = False
                Sleep 200
                Quarantine
                .Remove i
                lvwProcess.ListItems.Clear
                Label3.Caption = ""
                Label5.Caption = ""
                Label6.Caption = ""
                picSample.Cls
                txtValue(0) = ""
                txtValue(1) = ""
                txtValue(2) = ""
                Exit Sub
            End If
        Next i
    End With
    Call ProcessListSample(lvwProcessList, ilsDetected)
End Sub

Private Sub QuaProcess()
    On Error Resume Next
    Dim sNama As String
    Dim i As Long, lRet As Long
    Dim sFile As String
    Dim j As Long, lExitCode As Long
    
    With lvwProcess.ListItems
        For i = 1 To .count
            If .Item(i).Selected = True Then
                sFile = .Item(i).SubItems(1)
                QuarantineShow
                DoEvents
                VirusAlert
                TerminateExeName sFile
                For j = 1 To lvwProcess.ListItems.count
                    If lvwProcess.ListItems(i).Selected Then
                        lExitCode = KillProcessById(CLng(lvwProcess.ListItems(j).SubItems(2)))
                        If lExitCode = 0 Then MsgBox "Cannot terminate this process.", vbExclamation, "Unable To Terminate Process"
                    End If
                Next j
                sNama = QuarantineFile(sFile)
                LogScan "Quarantine Process... " & vbTab & sFile
                Kill sFile
                If lRet <> 0 Then
                    .Item(i).Selected = False
                End If
                .Item(i).Selected = False
'                Sleep 200
'                QuaProcess
                .Remove i
                lvwDetected.ListItems.Clear
                Label3.Caption = ""
                Label5.Caption = ""
                Label6.Caption = ""
                picSample.Cls
                txtValue(0) = ""
                txtValue(1) = ""
                txtValue(2) = ""
                Exit Sub
            End If
        Next i
    End With
End Sub

Function AddToColFileType(ID As String, col As Collection)
On Error GoTo Salah
   Dim buff As String
   col.Add ID, "#" & ID
   AddToColFileType = True
   Exit Function
Salah:
End Function

Public Sub GetIcon(icPath$, pDisp As PictureBox)
    pDisp.Cls
    Dim hImgSmall&: hImgSmall = SHGetFileInfo(icPath$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    'call SHGetFileInfo to return a handle to the icon associated with the specified file
    ImageList_Draw hImgSmall, shinfo.iIcon, pDisp.hdc, 0, 0, ILD_TRANSPARENT
     'Draw the icon to the specified picturebox control
End Sub

Public Sub GetIcons(lstView As ListView, imaList As ImageList, picTmp As PictureBox)
    Dim lsv As ListItem
    For Each lsv In lstView.ListItems
        picTmp.Cls
        GetIcon lsv.SubItems(1), picTmp
        imaList.ListImages.Add lsv.Index, , picTmp.Image
    Next
    
    With lstView
        .SmallIcons = imaList
        For Each lsv In .ListItems
            lsv.SmallIcon = lsv.Index
        Next
    End With
End Sub

Public Sub ControlListView(lstView As ListView, imaList As ImageList, picTmp As PictureBox)
    'On Local Error Resume Next
    picTmp.Cls
    picTmp.BackColor = &H8000000F
    lstView.BackColor = &HFFFFFF
    lstView.ListItems.Clear
    lstView.SmallIcons = Nothing
    imaList.ListImages.Clear
End Sub
