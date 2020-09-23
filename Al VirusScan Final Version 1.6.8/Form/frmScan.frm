VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6630
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin alVirusScan.dcButton cmdScanner 
      Height          =   390
      Index           =   0
      Left            =   150
      TabIndex        =   22
      Top             =   2625
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   688
      ButtonStyle     =   10
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   0
      PicNormal       =   "frmScan.frx":08CA
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin ComctlLib.ListView lvwDetected 
      Height          =   1890
      Left            =   150
      TabIndex        =   12
      Top             =   3150
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3334
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ilsVirus"
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Virus Name "
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "State"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filename"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Detection Type"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Directory"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date / Time"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5850
      Picture         =   "frmScan.frx":0A24
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtCmd 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6375
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5475
      Top             =   0
   End
   Begin ComctlLib.StatusBar sbScan 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   5190
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11633
            MinWidth        =   4410
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComCtl2.Animation aniScan 
      Height          =   690
      Left            =   225
      TabIndex        =   10
      Top             =   750
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1217
      _Version        =   393216
      FullWidth       =   51
      FullHeight      =   46
   End
   Begin alVirusScan.AdvProgressBar pbScanning 
      Height          =   165
      Left            =   1350
      TabIndex        =   19
      ToolTipText     =   "Scanning files..."
      Top             =   1305
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   291
      BarColor1       =   -2147483634
   End
   Begin alVirusScan.dcButton cmdScanner 
      Height          =   390
      Index           =   1
      Left            =   1140
      TabIndex        =   23
      Top             =   2625
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   688
      ButtonStyle     =   10
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   -2147483633
      PicAlign        =   0
      PicNormal       =   "frmScan.frx":0FAE
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin alVirusScan.dcButton cmdScanner 
      Height          =   390
      Index           =   2
      Left            =   2130
      TabIndex        =   24
      Top             =   2625
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   688
      ButtonStyle     =   10
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   -2147483633
      PicAlign        =   0
      PicNormal       =   "frmScan.frx":1548
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin alVirusScan.dcButton cmdScanner 
      Height          =   390
      Index           =   3
      Left            =   5475
      TabIndex        =   25
      Top             =   2625
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   688
      ButtonStyle     =   10
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   0
      PicNormal       =   "frmScan.frx":1AE2
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   ": "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   5400
      TabIndex        =   18
      Top             =   2205
      Width           =   990
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   "File Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   3750
      TabIndex        =   17
      Top             =   2205
      Width           =   1590
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   5400
      TabIndex        =   16
      Top             =   1950
      Width           =   75
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning Memory"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   3750
      TabIndex        =   15
      Top             =   1965
      Width           =   1590
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   5400
      TabIndex        =   14
      Top             =   1725
      Width           =   75
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Uncleaned"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   3750
      TabIndex        =   13
      Top             =   1725
      Width           =   1590
   End
   Begin ComctlLib.ImageList ilsVirus 
      Left            =   6150
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   1575
      TabIndex        =   9
      Top             =   2205
      Width           =   75
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Cleaned"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   225
      TabIndex        =   8
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   1575
      TabIndex        =   7
      Top             =   1965
      Width           =   75
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1575
      TabIndex        =   6
      Top             =   1725
      Width           =   75
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225
      X2              =   6450
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   225
      X2              =   6450
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1350
      TabIndex        =   5
      Top             =   225
      Width           =   5115
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Found"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   225
      TabIndex        =   4
      Top             =   1965
      Width           =   1215
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   "File Scanned"
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   3
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   "File "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   225
      Width           =   1065
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1350
      TabIndex        =   1
      Top             =   465
      Width           =   5115
   End
   Begin VB.Label lblvalue 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning in"
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   0
      Top             =   465
      Width           =   1065
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnExit 
         Caption         =   "Exit..."
      End
   End
   Begin VB.Menu mnuTask 
      Caption         =   "Scan Task"
      Begin VB.Menu mnuScan 
         Caption         =   "Scan Computer..."
         Index           =   0
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Scan Removable Disk..."
         Index           =   1
      End
      Begin VB.Menu mnuScan 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Scan Folder..."
         Index           =   3
      End
      Begin VB.Menu mnuScan 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Scan Running Process..."
         Index           =   5
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Scan System Recommended..."
         Index           =   6
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Scan Autorun Location..."
         Index           =   7
      End
      Begin VB.Menu mnuScan 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Start Scan..."
         Index           =   9
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================
' al VirusScan ver 1.0.0
'=======================================
' al VirusScan (AVS) akan menggantikan The Crying Machine
' dengan menggunakan scan engine yang baru
' dan dengan fitur-fitur yang baru.
' Semoga proyek ini cepat selesai ^_^
'#######################################
'
' Main Idea     : Moh Aly Shodiqin
' Company       : DQ Software
' Town          : Desa Campurejo RT 12/03 Panceng Gresik 61156 - Indonesia
'                 Copyright © 2008-2009 Moh Aly Shodiqin. All rights reserved
'
' History       :
'   1.0.0       : 20 Januari 2009
'               : Merancang form scan
'               : Scan Engine terbaru
'               : Engine AVIGEN
'               : Buat Icon AV
'               : Uflags Dialog BIF_RETURNONLYFSDIRS + BIF_EDITBOX + BIF_BROWSEINCLUDEFILES
'   1.0.1       : Status Bar ditambahkan untuk status scanning
'               : Uflags Dialog BIF_RETURNONLYFSDIRS + BIF_EDITBOX
'   1.0.5       : 29 Januari 2009
'               : Realtime Protection
'               : VirusScan Console
'               : Set Attribute File or Folder
'   1.0.6       : 30 Januari 2009
'               : Ditambahkan opsi Run when windows start pada VirusScan Console
'   1.6.6       : Update fungsi console
'               : Running on safe mode
'               : Perbaikan fungsi Realtime Protection v1.2
'               : Perbaikan pada fungsi Engine 1.3
'               : Ditambahkan fungsi Tweak Registry
'               : Ditambahkan thirdparty VirusScan Registry Editor
'               : Ditambahkan Update Online
'               : Semua opsi virusscan disimpan dalam registry
'               : update 15 Februari 2009
'               : list view database + column click
'               : nama database yang numerik di ganti scan.vdf
'   1.6.7       : 23 Februari 2009 10:57 AM
'               : Ditambahkan fungsi send the example of virus...
'               : ditambahkan Other System pada tweak registry
'               : fixed fungsi update online
'               : 27 Februari 2009 9:01 AM
'               : Changed icon dengan yang lebih matching dengan aplikasi ^_^
'=======================================
'
' Thanks To :
'
' Allah S.W.T
' Nabi Muhammad S.A.W
' My Parents
' My Soul
' AVIGEN - vbbego.com
' Noel A. Dacara
' Steve McMohan - www.VBAccelerator.com
' www.planetsourcecode.com
' Bagus Judistira
' Peradnya Dinata
' Thanks to all for the suggestions and comments.
'#######################################

'I'm sory my english language so bad ^_^

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As ESHGetFileInfoFlagConstants) As Long

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

Dim WithEvents SearchFile As cFileSearch
Attribute SearchFile.VB_VarHelpID = -1
Dim WithEvents Engine32 As cEngine32
Attribute Engine32.VB_VarHelpID = -1

Dim FileOnScan As Double
Dim FileViruses As Double
Dim FailedFile As Double
Dim CleanedFile As Double

Dim colTipeFile As Collection
Dim stateTipeFile As Boolean
Dim StateScan As String
Dim Y As Integer

Private shinfo As SHFILEINFO

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdScanner_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0:
            If StateScan <> "command" Then
                DoCmdScan "start"
            Else
                StateScan = ""
            End If
        Case 1:
            If StateScan <> "command" Then
                DoCmdScan "pause"
            Else
               StateScan = ""
            End If
        Case 2:
            If StateScan <> "command" Then
                DoCmdScan "stop"
            Else
                StateScan = ""
            End If
        Case 3
            Call mnExit_Click
    End Select
End Sub

Private Sub Engine32_onVirusFound(nFileName As String, nFileInfo As cFileInfo)
    On Error Resume Next
    FileViruses = FileViruses + 1
    lblValue(5) = ": " & FileViruses
    ViriOnCollect.Add nFileInfo
    Select Case UCase(nFileInfo.VirusAction)
        Case "DELETE"
            If nFileInfo.VirusClean Then
                VirusAlert
                CleanedFile = CleanedFile + 1
                addtoLView UCase(nFileInfo.VirusAlias), "DELETED", nFileInfo.Filename, nFileInfo.VirusType
                LogScan "Virus found " & nFileInfo.Filename & vbTab & nFileInfo.VirusAlias & vbTab & "DELETED"
            Else
                VirusAlert
                FailedFile = FailedFile + 1
                addtoLView UCase(nFileInfo.VirusAlias), "DELETED FAILED", nFileInfo.Filename, nFileInfo.VirusType, vbRed
                LogScan "Virus found " & nFileInfo.Filename & vbTab & nFileInfo.VirusAlias & vbTab & "DELETED FAILED"
            End If
        Case "QUARANTINE", "BUNDLE"
            If nFileInfo.VirusClean Then
                VirusAlert
                CleanedFile = CleanedFile + 1
                addtoLView UCase(nFileInfo.VirusAlias), "CLEAN + QUARANTINE", nFileInfo.Filename, nFileInfo.VirusType
                LogScan "Virus found " & nFileInfo.Filename & vbTab & nFileInfo.VirusAlias & vbTab & "CLEAN + QUARANTINE"
            Else
                VirusAlert
                FailedFile = FailedFile + 1
                addtoLView UCase(nFileInfo.VirusAlias), "CLEAN FAILED!", nFileInfo.Filename, nFileInfo.VirusType, vbRed
                LogScan "Virus found " & nFileInfo.Filename & vbTab & nFileInfo.VirusAlias & vbTab & "CLEAN FAILED!"
            End If
    End Select
    lblValue(7) = ": " & CleanedFile
    lblValue(9) = ": " & FailedFile
End Sub

Private Sub Form_Initialize()
    lvwStyle lvwDetected
'    SetFlatHeaders lvwDetected.hWnd
End Sub

Private Sub Form_Load()
    Me.Caption = "al VirusScan Scan With Virus Definitions" 'Beta Version " & "(" & vAppVersion & ")"
    LoadAnim
    cmdScanner(0).Caption = ""
    cmdScanner(1).Caption = ""
    cmdScanner(2).Caption = ""
    cmdScanner(3).Caption = ""
    cmdScanner(0).ToolTipText = "Start or resume scanning..."
    cmdScanner(1).ToolTipText = "Pause scanning..."
    cmdScanner(2).ToolTipText = "Stop scanning..."
    cmdScanner(3).ToolTipText = "Exit application..."
    cmdScanner(1).Enabled = False
    cmdScanner(2).Enabled = False
    Set Engine32 = New cEngine32
    Engine32.ClassIDApartement = Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(1) & Chr(255)

    Set SearchFile = New cFileSearch
    Set DrvOnCollect = New Collection
    
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.Hwnd, True
    Else
        AlwaysOnTop Me.Hwnd, False
    End If
    
    UpdateTypeFile

    ResetMe
    pbScanning.Visible = False
    pbScanning.Style = DoubleColor
    lblValue(11) = ": " & 0
'    sbScan.Panels(2).Text = Copyright
End Sub

Sub ResetMe()
    FileOnScan = 0
    FileViruses = 0
    CleanedFile = 0
    FailedFile = 0
    lblValue(4) = ": " & FileOnScan & " file."
    lblValue(5) = ": " & FileViruses
    lblValue(7) = ": " & CleanedFile
    lblValue(9) = ": " & FailedFile
    sbScan.Panels(1).Text = ""
    lvwDetected.ListItems.Clear
End Sub

Sub UpdateTypeFile()
    On Error Resume Next
    Dim H As String
    Dim I As Integer

    If REG.GetSettingString(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AllExtensions", "ALL") = "ALL" Then       'getFileExt = "*.*"
       lblValue(13) = ": ALL"
       lblValue(13).ToolTipText = ": ALL FILE *.*"
       stateTipeFile = True
    Else
       stateTipeFile = False
       Set colTipeFile = Nothing
       Set colTipeFile = New Collection
       
       H = REG.GetSettingString(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "ExtensionsSelected", "386|BAT|BIN|BTM|CLA|COM|CSC|DLL|DRV|EXE|EX_|OCX|OV?|PIF|SYS|VXD|CSH|DOC|DOT|HLP|HTA|HTM|HTML|HTT|INF|INI|JS|JSE|JTD|MDB|MP?|MSO|ODB|OBT|PL|PM|POT|PPS|PPT|RTF|SH|SHB|SHS|SMM|VBE|VBS|VSD|VSS|VST|WSF|WSH|XLA|XLS|SCR|SC_|TMP|JPG|REG")
       Dim inFIle() As String
       inFIle() = Split(H, "|", , vbTextCompare)
       For I = 1 To UBound(inFIle)
           AddToColFileType inFIle(I)
       Next I
       
       SearchFile.SetFileType = colTipeFile
       
       lblValue(13) = ": FILTERED"
       lblValue(13).ToolTipText = ": " & H
    End If
End Sub

Sub DoCmdScan(cmd As String, Optional ShowDlg As Boolean = True)
    On Error Resume Next
    Select Case LCase(cmd)
        Case "start"
            If cmdScanner(0).Tag = "" Then
                If ShowDlg Then
                    Load frmDrive
                    frmDrive.show 1, Me
                    If OnSelectDlg = vbCancel Then Exit Sub
                End If
                
                If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "ScanMemory", 1) = 1 Then
                    KillVirusProcessList True
                Else
                    lblValue(11) = ": " & 0
                End If
                lvwDetected.ListItems.Clear
                
                FileOnScan = 0
                FileViruses = 0
                CleanedFile = 0
                FailedFile = 0
                ResetMe
                Set SearchFile = New cFileSearch
                cmdScanner(0).Enabled = False
                cmdScanner(2).Enabled = True
                cmdScanner(1).Enabled = True
                Me.Caption = "al VirusScan Scan - "
                SearchFile.StopSearch = False
                Set ViriOnCollect = New Collection

'                StartTickCount = GetTickCount()
'
'                ---------------------------
                DisabledControl False
'                ---------------------------
                Dim I As Integer
                StateScan = "start"
                UpdateTypeFile
                
                If OnSelectDlg = vbOK Then
                    If DrvOnCollect.count > 0 Then
                        For I = 1 To DrvOnCollect.count
                            lblFilename = ""
                            lblPath = ""
                            Me.Caption = "al VirusScan Scan - " & DrvOnCollect(I)
                            If StateScan = "start" Then
                                If Trim(DrvOnCollect(I)) <> "" Then
                                    aniScan.Stop
                                    aniScan.Play
                                    aniScan.Visible = True
                                    If tmrScan.Enabled = False Then
                                        tmrScan.Enabled = True
                                        pbScanning.Visible = True
                                    End If
                                    SearchFile.DoCmdSearchFile DrvOnCollect(I), stateTipeFile
                                    LogScan "Scanning in " & vbTab & DrvOnCollect(I) & vbTab & "Memory Scanned " & lblValue(11) & "   File Scanned " & lblValue(4)
                                End If
                            End If
                        Next I
                    End If
                End If
                
                aniScan.Stop
                aniScan.Visible = False
                tmrScan.Enabled = False
                pbScanning.Value = 0
                pbScanning.Visible = False
                If StateScan = "stop" Then
                    sbScan.Panels(1).Text = "Aborted by user."
                End If
            
                cmdScanner(0).Enabled = True
                cmdScanner(2).Enabled = False
                cmdScanner(1).Enabled = False
                lblFilename = ""
                lblPath = ""
                sbScan.Panels(1).Text = "Completed."
                cmdScanner(0).Tag = ""
                Me.Caption = "al VirusScan Scan With Virus Definitions"
'                ---------------------------
                DisabledControl True
'                ---------------------------
    
                StateScan = "stop"
'                TotalElapsedMilliSec = TotalElapsedMilliSec + (GetTickCount() - StartTickCount)
'                TotalElapsedMilliSec = 0
                If ViriOnCollect.count > 0 Then
'                   ANVIBI_frmSolution.show 1, Me
                End If
                sbScan.Panels(1).Text = "Completed."
            Else
                aniScan.Play
                aniScan.Visible = True
                If tmrScan.Enabled = False Then
                    tmrScan.Enabled = True
                    pbScanning.Visible = True
                End If
                cmdScanner(0).Enabled = False
                cmdScanner(2).Enabled = True
                cmdScanner(1).Enabled = True
                StateScan = "pause"
                SearchFile.PauseSearch = False
            End If
       Case "stop"
            SearchFile.StopSearch = True
            StateScan = "stop"
'            cmdScanner(0).Enabled = True
'            cmdScanner(2).Enabled = False
            cmdScanner(1).Enabled = True
            sbScan.Panels(1).Text = "Aborted by user."
            cmdScanner(0).Tag = ""
            aniScan.Stop
            aniScan.Visible = False
            tmrScan.Enabled = False
            pbScanning.Value = 0
            pbScanning.Visible = False
        Case "pause"
            aniScan.Stop
            aniScan.Visible = False
            tmrScan.Enabled = False
            pbScanning.Value = 0
            pbScanning.Visible = False
            SearchFile.PauseSearch = True
            cmdScanner(0).Enabled = True
            cmdScanner(2).Enabled = True
            cmdScanner(1).Enabled = False
            sbScan.Panels(1).Text = "Paused."
            cmdScanner(0).Tag = "pause"
            StateScan = "pause"
            sbScan.Panels(1).Text = "Scan paused..."
    End Select
End Sub

Sub ScanOnlyFilesOrFolder(Optional isFolder As Boolean = False, Optional AddSpesifikDir As String = "")
    On Error Resume Next
    If isFolder = True Then
        Dim H As String
        lastPath = IIf(lastPath = "", nPath(App.path), lastPath)
        If AddSpesifikDir = "" Then
           H = BrowseFolder(Me.Hwnd, "Select direcotry to scan")
        Else
           H = AddSpesifikDir
        End If
        If Trim(H) <> "" Then
           lastPath = H
           Set DrvOnCollect = Nothing
           Set DrvOnCollect = New Collection
           DrvOnCollect.Add H, H
           OnSelectDlg = vbOK
           DoCmdScan "start", False
        End If
    End If
End Sub

Private Function AddToColFileType(ID As String)
    On Error GoTo Salah
    Dim buff As String
    colTipeFile.Add ID, "#" & ID
    AddToColFileType = True
    Exit Function
Salah:
End Function

Private Sub LoadAnim()
    On Error Resume Next
    Dim H As String
    H = App.path & "\Data\Findfile.avi"
    aniScan.Visible = False
    aniScan.Open H
End Sub

Sub addtoLView(ntext As String, nState As String, nfile As String, dt As String, Optional ncolor As Long = -1)
    On Error Resume Next
    Dim l As ListItem
    ControlListView lvwDetected, ilsVirus, pic
    Set l = lvwDetected.ListItems.Add(, , ntext)
        l.SubItems(1) = nState
        l.SubItems(2) = file_getName(nfile)
        l.SubItems(3) = dt
        l.SubItems(4) = file_getPath(nfile)
        l.SubItems(5) = Format(Date, "ddd, dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS")
        l.Tag = nfile
        lvwDetected.ListItems(lvwDetected.ListItems.count).EnsureVisible
        lvwDetected.ListItems(lvwDetected.ListItems.count).Selected = True
        If lvwDetected.ListItems.count <> 0 Then GetIcons lvwDetected, ilsVirus, pic
End Sub

Public Sub GetIcon(icPath$, pDisp As PictureBox)
    pDisp.Cls
    Dim hImgSmall&: hImgSmall = SHGetFileInfo(icPath$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    ImageList_Draw hImgSmall, shinfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
End Sub

Public Sub GetIcons(lstView As ListView, imaList As ImageList, picTmp As PictureBox)
    Dim lsv As ListItem
    For Each lsv In lstView.ListItems
        picTmp.Cls
        GetIcon lsv.Tag, picTmp
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
    picTmp.Cls
    picTmp.BackColor = vbWhite
    lstView.SmallIcons = Nothing
    imaList.ListImages.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call mnExit_Click
    If cmdScanner(0).Enabled = True Then
'        SearchFile.StopSearch = True
'        Engine32.CloseScanHandle
        Set SearchFile = Nothing
        Set Engine32 = Nothing
        Unload Me
    Else
        If MsgBox("Abort the current process...", vbQuestion + vbYesNo, "al VirusScan") = vbYes Then
'            SearchFile.StopSearch = True
'            Engine32.CloseScanHandle
            Set SearchFile = Nothing
            Set Engine32 = Nothing
            Unload Me
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub mnExit_Click()
'    Set SearchFile = Nothing
'    Set Engine32 = Nothing
    Unload Me
End Sub

Private Sub mnuAbout_Click()
    frmAbout.show
End Sub

Private Sub mnuScan_Click(Index As Integer)
    Select Case Index
        Case 0
            Set DrvOnCollect = Nothing
            Set DrvOnCollect = New Collection
       
            LDs = GetLogicalDrives
            For Cnt = 0 To 25
                If (LDs And 2 ^ Cnt) <> 0 Then
                   DrvOnCollect.Add Chr(65 + Cnt) & ":\", Chr(65 + Cnt) & ":\"
                End If
            Next Cnt
            OnSelectDlg = vbOK
            DoCmdScan "start", False
        Case 1
            Set DrvOnCollect = Nothing
            Set DrvOnCollect = New Collection
       
            LDs = GetLogicalDrives
            For Cnt = 0 To 25
                If (LDs And 2 ^ Cnt) <> 0 Then
                   If GetTipeDrive(Chr(65 + Cnt) & ":\") = 2 Then
                      DrvOnCollect.Add Chr(65 + Cnt) & ":\", Chr(65 + Cnt) & ":\"
                   End If
                End If
            Next Cnt
            OnSelectDlg = vbOK
            DoCmdScan "start", False
        Case 3
            ScanOnlyFilesOrFolder True
        Case 5
            KillVirusProcessList True
        Case 6
'            KillVirusProcessList True
'            ScanVirusFromRegistry
           
            'ScanOnlyFilesOrFolder True, MyWindowSys
            'C:\Documents and Settings\Administrator\My Documents
            Dim inDoc() As String
            inDoc() = Split(GetSpecialfolder(5), "\")
            If UBound(inDoc) > 0 Then
               ScanOnlyFilesOrFolder True, inDoc(0) & "\" & inDoc(1)
            End If
            ScanOnlyFilesOrFolder True, GetSpecialfolder(43) '43 - CommonFiles
            '------ Windows system
            inDoc() = Split(MyWindowSys, "\")
            If UBound(inDoc) > 0 Then
               ScanOnlyFilesOrFolder True, inDoc(0) & "\" & inDoc(1)
            End If
        Case 7
            ScanVirusFromRegistry
        Case 9
            DoCmdScan "start"
    End Select
End Sub

Private Sub SearchFile_onSearch(nFileName As String, nFileInfo As cFileInfo)
    If Trim(nFileInfo.Filename) <> "" Then
        lblFilename = nFileInfo.Filename
        lblPath = nFileInfo.FilePath
        If Engine32.CekOneFile(nPath(nFileInfo.FilePath) & nFileInfo.Filename) Then
            '
        End If
        FileOnScan = FileOnScan + 1
    Else
        sbScan.Panels(1).Text = "Scanning files..."
    End If
    lblValue(4) = ": " & FileOnScan & " files."
End Sub

Sub ScanVirusFromRegistry()
    On Error Resume Next
    Dim colRun As New Collection
    Dim data(3) As String
    
    Dim FileInfo As New cFileInfo
    
    Call REG.GetEnumValue(HKEY_CURRENT_USER, SMWC & "\Run", colRun)
    Call REG.GetEnumValue(HKEY_CURRENT_USER, SMWC & "\RunOnce", colRun)
    Call REG.GetEnumValue(HKEY_LOCAL_MACHINE, SMWC & "\Run", colRun)
    Call REG.GetEnumValue(HKEY_LOCAL_MACHINE, SMWC & "\RunOnce", colRun)
    Call REG.GetEnumValue(HKEY_LOCAL_MACHINE, SMWC & "\RunOnceEx", colRun)
    
    REG.SaveSettingString HKEY_CLASSES_ROOT, "exefile\shell\open\command", vbNullString, Chr(34) & "%1" & Chr(34) & " %*"
    
    Dim hDir As String, hStartUpPath As String
    hStartUpPath = nPath(ReplacePathSystem(GetSpecialfolder(&H7))) '&H18
    LogScan "Checking startup directory... " & hStartUpPath
    hDir = Dir(hStartUpPath & "*.*", vbDirectory + vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If Trim(hDir) <> "" Then
       While hDir <> ""
           If hDir <> ".." And hDir <> "." Then
              data(0) = hStartUpPath
              data(1) = Chr(34) & hStartUpPath & hDir & Chr(34)
              data(2) = ""
              data(3) = ""
              colRun.Add data
           End If
           hDir = Dir()
       Wend
    End If
    
    hStartUpPath = nPath(ReplacePathSystem(GetSpecialfolder(&H18)))
    LogScan "Checking startup directory... " & hStartUpPath
    hDir = Dir(hStartUpPath & "*.*", vbDirectory + vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If Trim(hDir) <> "" Then
       While hDir <> ""
           If hDir <> ".." And hDir <> "." Then
              data(0) = hStartUpPath
              data(1) = Chr(34) & hStartUpPath & hDir & Chr(34)
              data(2) = ""
              data(3) = ""
              colRun.Add data
           End If
           hDir = Dir()
       Wend
    End If
    '-------------
    Dim I As Integer
    Dim param() As String
    Dim myFileName As String
    If colRun.count > 0 Then
       LogScan "Checking registry entry on autorun locations... "
       For I = 1 To colRun.count
           If InStr(1, CStr(colRun(I)(1)), Chr(34), vbTextCompare) Then
              CL_Get param, CStr(colRun(I)(1)), " "
              myFileName = ReplacePathSystem(param(0))
           Else
              myFileName = ReplacePathSystem(GetFileNameFromParam(CStr(colRun(I)(1))))
           End If
           
           If Trim(myFileName) <> "" Then
              FileInfo.FilePath = myFileName
              LogScan "Scanning... " & myFileName
              If Engine32.CekOneFile(myFileName) = False Then
                '
              Else
                 If CStr(colRun(I)(2)) <> "" Then
                    REG.DeleteValue CLng(colRun(I)(2)), CStr(colRun(I)(3)), CStr(colRun(I)(0))
                 End If
              End If
           End If
       Next I
    End If
    sbScan.Panels(1).Text = "Completed."
    Set colRun = Nothing
End Sub

Sub KillVirusProcessList(Optional onboot As Boolean = False)
    On Error Resume Next
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    Dim namafile As String, lngModules(1 To 200) As Long
    Dim strModuleName As String, Xproses As Long
    Dim enumerasi As Long, strProcessName As String
    Dim lngSize As Long
    Dim lngReturn  As Long
    Set ViriOnCollect = New Collection
    Dim fileIsVirus As New Collection
                         
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    enumerasi = Process32First(hSnapShot, uProcess)
    lngSize = 500
    strModuleName = SPACE(MAX_PATH)
    FileOnScan = 0
    FileViruses = 0
    CleanedFile = 0
    FailedFile = 0
        
    Dim data(1) As String
    
    Do While enumerasi
        Xproses = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, uProcess.th32ProcessID)
        lngReturn = GetModuleFileNameExA(Xproses, lngModules(1), strModuleName, lngSize)
        strProcessName = ReplacePathSystem(Left(strModuleName, lngReturn))
        If strProcessName <> "" Then
            lblFilename = file_getName(strProcessName)
            lblPath = "Memory" 'file_getPath(strProcessName)
            If onboot Then
                Sleep 70
            End If
            If Engine32.FindVirusOnly(strProcessName) Then
                data(0) = strProcessName
                data(1) = uProcess.th32ProcessID
                fileIsVirus.Add data
                SuspenResumeThread uProcess.th32ProcessID, False
            End If
        End If
        namafile = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
        enumerasi = Process32Next(hSnapShot, uProcess)
        FileOnScan = FileOnScan + 1
        lblValue(4) = ": " & 0
        lblValue(11) = ": " & FileOnScan 'memory scan
        sbScan.Panels(1).Text = "Scanning memory..."
        lblFilename = ""
        lblPath = ""
    Loop
    CloseHandle hSnapShot
        
    If fileIsVirus.count > 0 Then
       Dim I As Integer
       For I = 1 To fileIsVirus.count
           If Engine32.CekOneFile(CStr(fileIsVirus(I)(0)), CLng(fileIsVirus(I)(1))) Then
              
           End If
          lblValue(5) = ": " & FileViruses 'virus found
       Next I
    End If
        
    FileOnScan = 0
                       
    If ViriOnCollect.count > 0 Then
        '
    End If
    sbScan.Panels(1).Text = "Completed."
End Sub

Sub DisabledControl(dControl As Boolean)
    mnuFile.Enabled = dControl
    mnuTask.Enabled = dControl
    mnuHelp.Enabled = dControl
'    lvwDetected.Enabled = dControl
End Sub

Private Sub tmrScan_Timer()
    Y = Y + 2
    If Y > 100 Then Y = 0

    pbScanning.Value = Y
End Sub

Sub RunScanFromParam(cmd As String)
    On Error Resume Next
    If LCase(StateScan) = "stop" Or Trim(StateScan) = "" Then
    Set DrvOnCollect = Nothing
    Set DrvOnCollect = New Collection
    Dim param() As String
    CL_Get param, cmd, " "
    StateScan = "start"
    If UBound(param) >= 0 Then
'        AddLog " þ Scanning file/folder...", &HC0FFC0
'        AddLog " ", &HC0FFC0
        FileOnScan = 0: FileViruses = 0
        CleanedFile = 0
        FailedFile = 0
              
        Dim fName As Integer
        For fName = 0 To UBound(param)
            If file_isFolder(param(fName)) Then
                DrvOnCollect.Add param(fName), param(fName)
            Else
'               AddLog " þ Scanning file " & param(fname), &HC0FFC0
                Me.Caption = "Scanning file : " & param(fName)
                If Engine32.CekOneFile(param(fName)) Then
                Else
'                AddLog " þ No virus detected on " & param(fname), &HFFFFC0
'                AddLog " "
                End If
            End If
        Next fName
        
        If DrvOnCollect.count > 0 Then
            OnSelectDlg = vbOK
            DoCmdScan "start", False
        End If
'        AddLog " þ Done."
    End If
    StateScan = "stop"
Else
'    AddLog " þ Can't scan now, current process still running..."
'    AddLog " þ please stop first and try again!"
End If
End Sub

Sub ScanCommandLine(data() As String)
    On Error Resume Next
    If LCase(StateScan) = "stop" Or Trim(StateScan) = "" Then
        StateScan = "start"
        Set DrvOnCollect = Nothing
        Set DrvOnCollect = New Collection
               
'        AddLog " þþþþþþþþ"
'        AddLog " þ Scanning from parameters...", &HC0FFC0
'        AddLog " ", &HC0FFC0
        FileOnScan = 0: FileViruses = 0
        CleanedFile = 0
        FailedFile = 0
              
        Dim fName As Variant
        Dim curName As String
        Dim FileInfo As CFileInfo32
        Dim m_cShortcut As cShellLink
        Dim I As Integer
    
        For I = 0 To UBound(data)
            If file_isFolder(data(I)) Then
                DrvOnCollect.Add data(I), data(I)
            Else
                curName = data(I)
                If LCase(file_getTitle(curName)) = "lnk" Then
                    Set FileInfo = New CFileInfo32
                    FileInfo.FullPathName = CStr(curName)
                    Select Case LCase(FileInfo.TypeName)
                        Case "shortcut"
                            Set m_cShortcut = New cShellLink
                            If m_cShortcut.Resolve(curName) Then
                                curName = m_cShortcut.TargetFile
                            End If
                    End Select
                    Set FileInfo = Nothing
                    Set m_cShortcut = Nothing
                End If
               
'                AddLog " þ Scanning file " & file_getName(CStr(curName)), &HC0FFC0
                Me.Caption = "Scanning file " & CStr(curName)
                If Engine32.CekOneFile(CStr(curName)) Then
                    '
                Else
'                    AddLog " þ No virus detected.", &HFFFFC0
'                    AddLog " "
                End If
            End If
        Next I
    
        If DrvOnCollect.count > 0 Then
            OnSelectDlg = vbOK
            DoCmdScan "start", False
        End If
'        AddLog " þ Done."
        StateScan = "stop"
    Else
'        AddLog " þ Can't scan now, current process still running..."
'        AddLog " þ please stop first and try again!"
    End If
End Sub
