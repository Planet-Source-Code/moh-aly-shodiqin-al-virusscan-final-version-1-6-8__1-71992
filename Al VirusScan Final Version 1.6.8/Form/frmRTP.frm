VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmRTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7650
   Icon            =   "frmRTP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   75
      Picture         =   "frmRTP.frx":08CA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   2025
      Visible         =   0   'False
      Width           =   315
   End
   Begin ComctlLib.ListView lvwRTP 
      Height          =   1290
      Left            =   0
      TabIndex        =   13
      Top             =   2400
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   2275
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ilsVirus"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
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
         Text            =   "Name "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "In Folder"
         Object.Width           =   8997
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Detected As"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Detection Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date / Time"
         Object.Width           =   4480
      EndProperty
   End
   Begin alVirusScan.IEWatch IEWatch1 
      Index           =   0
      Left            =   6750
      Top             =   3675
      _ExtentX        =   556
      _ExtentY        =   423
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   975
      ScaleHeight     =   1815
      ScaleWidth      =   6465
      TabIndex        =   1
      Top             =   300
      Width           =   6465
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
         Height          =   315
         Index           =   3
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1470
         Width           =   4740
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
         Height          =   315
         Index           =   2
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1155
         Width           =   4740
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
         Height          =   315
         Index           =   1
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   4740
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
         Height          =   315
         Index           =   0
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   525
         Width           =   4740
      End
      Begin VB.Label lblValue 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   75
         TabIndex        =   8
         Top             =   1470
         Width           =   1365
      End
      Begin VB.Label lblValue 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   1650
         TabIndex        =   7
         Top             =   75
         Width           =   4740
      End
      Begin VB.Label lblValue 
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   6
         Top             =   1425
         Width           =   1365
      End
      Begin VB.Label lblValue 
         Caption         =   "Detected As"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   5
         Top             =   1155
         Width           =   1365
      End
      Begin VB.Label lblValue 
         Caption         =   "In Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   4
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label lblValue 
         Caption         =   "Date and Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   3
         Top             =   525
         Width           =   1365
      End
      Begin VB.Label lblValue 
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   75
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   2115
      Left            =   900
      TabIndex        =   0
      Top             =   75
      Width           =   6615
   End
   Begin ComctlLib.ImageList ilsVirus 
      Left            =   150
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   75
      Picture         =   "frmRTP.frx":0E54
      Top             =   150
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuVirusScan 
         Caption         =   "al VirusScan..."
      End
      Begin VB.Menu mnuScan 
         Caption         =   "al VirusScan Scan With Virus Sample..."
      End
      Begin VB.Menu mnuB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsole 
         Caption         =   "VirusScan Console..."
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "Enable/Disable Realtime Protection..."
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "Monitoring Directory..."
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "Send The Example Of Virus..."
         Index           =   2
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "Cancel Menu..."
         Index           =   3
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "Exit Application..."
         Index           =   5
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "Update Now..."
         Index           =   7
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "Turn Off Computer..."
         Index           =   9
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuDisa 
         Caption         =   "About al VirusScan..."
         Index           =   11
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
      Begin VB.Menu mnuAct 
         Caption         =   "Select All"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAct 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuAct 
         Caption         =   "Remove Message from List"
         Index           =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuAct 
         Caption         =   "Remove All Messages"
         Index           =   3
      End
      Begin VB.Menu mnuAct 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuAct 
         Caption         =   "Close"
         Index           =   5
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 29 Januari 2009
' 10:50 PM
'
' Update 8 Februari 2009
' 8:53 PM
'=======================================
' Module Realtime Protection
'=======================================
Option Explicit

Dim isCompatch As Boolean

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

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

Dim iewindow As InternetExplorer

Dim isRunningOut As Boolean
Dim isRunningProcess As Boolean

Dim StopAll As Boolean
Dim WithEvents Engine32 As cEngine32
Attribute Engine32.VB_VarHelpID = -1
Dim isTerminate As Boolean
Dim isScannedOn As Boolean
Dim myTray As NOTIFYICONDATA

Dim WithEvents ShellIE As SHDocVw.ShellWindows
Attribute ShellIE.VB_VarHelpID = -1

Dim sTitle As String
Dim sMessage As String
Private Const AVS = "al VirusScan - Realtime Protection"
Private Const al = "al VirusScan"
Private shinfo As SHFILEINFO
Dim ViriOnCollect As Collection

Private Sub Engine32_onVirusFound(nFileName As String, nFileInfo As cFileInfo)
    On Error Resume Next
    Me.Caption = "VirusScan Scan Messages"
    Select Case UCase(nFileInfo.VirusAction)
        Case "DELETE"
            If nFileInfo.VirusClean Then
                VirusAlert
                lblValue(5) = ": VirusScan Alert!"
                txtValue(0) = ": " & Format(Date, "ddd, dd/mm/yyyy") & " " & Time
                txtValue(1) = ": " & nFileName
                txtValue(2) = ": " & nFileInfo.VirusAlias
                txtValue(3) = ": " & "DELETED"
                addtoLV UCase(nFileInfo.VirusAlias), "delete", "DELETED", nFileInfo.Filename, nFileInfo.VirusType
                LogScan "Virus found " & nFileInfo.Filename & vbTab & nFileInfo.VirusAlias & vbTab & "DELETED"
            Else
                VirusAlert
                lblValue(5) = ": VirusScan Alert!"
                txtValue(0) = ": " & Format(Date, "ddd, dd/mm/yyyy") & " " & Time
                txtValue(1) = ": " & nFileName
                txtValue(2) = ": " & nFileInfo.VirusAlias
                txtValue(3) = ": " & "DELETED FAILED"
                addtoLV UCase(nFileInfo.VirusAlias), "unclean", "DELETED FAILED", nFileInfo.Filename, nFileInfo.VirusType, vbRed
                LogScan "Virus found " & nFileInfo.Filename & vbTab & nFileInfo.VirusAlias & vbTab & "DELETED FAILED"
            End If
        Case "QUARANTINE", "BUNDLE"
            If nFileInfo.VirusClean Then
                VirusAlert
                lblValue(5) = ": VirusScan Alert!"
                txtValue(0) = ": " & Format(Date, "ddd, dd/mm/yyyy") & " " & Time
                txtValue(1) = ": " & nFileName
                txtValue(2) = ": " & nFileInfo.VirusAlias
                txtValue(3) = ": " & "CLEAN + QUARANTINE"
                addtoLV UCase(nFileInfo.VirusAlias), "clean", "CLEAN + QUARANTINE", nFileInfo.Filename, nFileInfo.VirusType
                LogScan "Virus found " & nFileInfo.Filename & vbTab & nFileInfo.VirusAlias & vbTab & "CLEAN + QUARANTINE"
            Else
                VirusAlert
                lblValue(5) = ": VirusScan Alert!"
                txtValue(0) = ": " & Format(Date, "ddd, dd/mm/yyyy") & " " & Time
                txtValue(1) = ": " & nFileName
                txtValue(2) = ": " & nFileInfo.VirusAlias
                txtValue(3) = ": " & "CLEAN FAILED!"
                addtoLV UCase(nFileInfo.VirusAlias), "unclean", "CLEAN FAILED!", nFileInfo.Filename, nFileInfo.VirusType, vbRed
                LogScan "Virus found " & nFileInfo.Filename & vbTab & nFileInfo.VirusAlias & vbTab & "CLEAN FAILED!"
            End If
    End Select
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Sub addtoLV(ntext As String, nicon, nState As String, nfile As String, dt As String, Optional ncolor As Long = -1)
    On Error Resume Next
    Dim l As ListItem
    ControlListView lvwRTP, ilsVirus, pic
    Set l = lvwRTP.ListItems.Add(, , file_getName(nfile))
        l.SubItems(1) = file_getPath(nfile)
        l.SubItems(2) = ntext
        l.SubItems(3) = dt
        l.SubItems(4) = nState
        l.SubItems(5) = Format(Date, "ddd, dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS")
        l.Tag = nfile
        lvwRTP.ListItems(lvwRTP.ListItems.count).EnsureVisible
        lvwRTP.ListItems(lvwRTP.ListItems.count).Selected = True
        If lvwRTP.ListItems.count <> 0 Then GetIcons lvwRTP, ilsVirus, pic
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

Private Sub Form_Initialize()
    CopyPlugin
    RegShell
    'Repair system windows
    REG.SaveSettingString HKEY_CLASSES_ROOT, "exefile\shell\open\command", vbNullString, Chr(34) & "%1" & Chr(34) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "piffile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "scrfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "regfile\shell\open\command", "", "regedit.exe %1"
    '--------------
    InitCommonControls
    lvwStyle lvwRTP
End Sub

Private Sub Form_Load()
    On Error GoTo Salah
    If App.PrevInstance Then
        MsgBox "al VirusScan is already run in your system.", vbExclamation, "al VirusScan"
        End
    End If
    Me.Caption = "VirusScan Scan Messages"
    Frame1.Caption = "VirusScan Message"
    Me.Top = 3000
    Me.Left = 4000
        
    Set Engine32 = New cEngine32
    Set ViriOnCollect = New Collection
    Set ShellIE = New SHDocVw.ShellWindows
    WatchIt

    Engine32.ClassIDApartement = Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(1) & Chr(255)
    KillVirusProcessList
    frmScan.ScanVirusFromRegistry
    
    If REG.GetSettingString(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "About Author", "") = "" Then
        On Error Resume Next
        frmAuthor.show
        REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "About Author", "1"
    End If
    If REG.GetSettingString(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "RealtimeProtection", "Enabled") = "Enabled" Then
        sTitle = AVS & " enabled"
        sMessage = "- Final Version " & vAppVersion & vbCrLf & _
                    "- Developed By Moh Aly Shodiqin" & vbCrLf & _
                    "- " & App.LegalCopyright & vbCrLf & vbCrLf & _
                    "- Virus Definitions : " & frmDatabase.lvwDB.ListItems.count
        mnuDisa(0).Checked = True
    Else
        sTitle = AVS & " disabled"
        sMessage = "- Final Version " & vAppVersion & vbCrLf & _
                    "- Developed By Moh Aly Shodiqin" & vbCrLf & _
                    "- Copyright © " & App.LegalCopyright & vbCrLf & vbCrLf & _
                    "- Virus Definitions : " & frmDatabase.lvwDB.ListItems.count
        mnuDisa(0).Checked = False
    End If
    If REG.GetSettingString(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "MonitoringDirectory", "Enabled") = "Enabled" Then
        mnuDisa(1).Checked = True
    Else
        mnuDisa(1).Checked = False
    End If

    SystrayOn Me, AVS
    PopupBalloon Me, sMessage, sTitle, NIIF_INFO
    If DateDiff("d", vVirusDefinitions, CDate(Date), vbUseSystemDayOfWeek, vbUseSystem) > 30 Then 'CDate(vVirusDefinitions) < Month(Date) Then 'DateDiff("d", vVirusDefinitions, CDate(Date)) > 5 Then
        Sleep 5000
        frmCheckUpdate.show
    End If
'    Debug.Print vVirusDefinitions
Salah:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim Action As Long
    If Me.ScaleMode = vbPixels Then
        Action = x
      Else
        Action = x / Screen.TwipsPerPixelX
    End If
    Select Case Action
      Case WM_RBUTTONUP
        PopupMenu mnuFile
    End Select
End Sub

Sub FindViriOnCurrentFolder(PathName As String)
    On Error Resume Next
    If isRunningOut = False Then
        isRunningOut = True
        Dim hDir As String
        hDir = Dir(nPath(PathName) & "*.*", vbHidden + vbNormal + vbReadOnly + vbSystem + vbArchive)
        If hDir <> "" Then
            While hDir <> ""
                If Engine32.CekOneFile(nPath(PathName) & hDir) = False Then
                    'KillVirusProcessList
                    'Call Engine32.CekOneFile(nPath(PathName) & hDir)
                End If
               
                hDir = Dir()
                DoEvents
                If StopAll Then Exit Sub
            Wend
        End If
        isRunningOut = False
    End If
End Sub

Sub KillVirusProcessList()
    On Error Resume Next
    If isRunningProcess = False Then
    isRunningProcess = True
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    Dim namafile As String, lngModules(1 To 200) As Long
    Dim strModuleName As String, Xproses As Long
    Dim enumerasi As Long, strProcessName As String
    Dim lngSize As Long
    Dim lngReturn  As Long
    
    Set ViriOnCollect = Nothing
    Set ViriOnCollect = New Collection
    Dim fileIsVirus As New Collection
                   
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    enumerasi = Process32First(hSnapShot, uProcess)
    lngSize = 500
    strModuleName = SPACE(MAX_PATH)
    
    Dim data(1) As String
    Do While enumerasi
        Xproses = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, uProcess.th32ProcessID)
        lngReturn = GetModuleFileNameExA(Xproses, lngModules(1), strModuleName, lngSize)
        strProcessName = ReplacePathSystem(Left(strModuleName, lngReturn))
        If strProcessName <> "" Then
            If Engine32.FindVirusOnly(strProcessName) Then
                data(0) = strProcessName
                data(1) = uProcess.th32ProcessID
                fileIsVirus.Add data
                SuspenResumeThread uProcess.th32ProcessID, False
            End If
        End If
        namafile = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
        enumerasi = Process32Next(hSnapShot, uProcess)
       
    Loop
    CloseHandle hSnapShot
       
    If fileIsVirus.count > 0 Then
       Dim i As Integer
       For i = 1 To fileIsVirus.count
           If Engine32.CekOneFile(CStr(fileIsVirus(i)(0)), CLng(fileIsVirus(i)(1))) Then
                 LogScan "Scan Memory Found " & fileIsVirus(i)(0)
           End If
       Next i
    End If
                           
    If ViriOnCollect.count > 0 Then
        Engine32.RunningSolution ViriOnCollect
    End If
    isRunningProcess = False
    End If
End Sub

Sub WatchIt()
    If isCompatch = False Then
        Dim i As Integer, Cnt As Integer
        Cnt = ShellIE.count - 1
        For i = 0 To Cnt
            If (IEWatch1.count - 1) < Cnt Then
                AddIEObj i
            End If
'            If IEWatch1(i).IEKey = 0 Then
                If FindID(ShellIE(i).Hwnd) = False Then
                    IEWatch1(i).EnabledMonitoring mnuDisa(0).Checked
                    IEWatch1(i).AddSubClass ShellIE(i)
                End If
'            End If
        Next i
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Select Case UnloadMode
        Case 0
            If isTerminate Then
                StopAll = True
                TerminateProcess GetCurrentProcess, 0
            Else
                Me.Hide
                Cancel = True
            End If
       Case Else
    End Select
End Sub

Private Sub IEWatch1_FileNameSeletedChange(Index As Integer, strFilename As String, Fullpath As String)
If mnuDisa(0).Checked Then
    If Trim(strFilename) <> "" Then
        If file_isFolder(Fullpath) = 0 Then
           If Engine32.CekOneFile(Fullpath) = False Then
              KillVirusProcessList
              Engine32.CekOneFile Fullpath
           End If
        End If
    End If
End If
End Sub

Sub AddIEObj(Index As Integer)
    On Error GoTo Salah
    Load IEWatch1(Index)
Salah:
End Sub
Private Sub IEWatch1_IEClosed(Index As Integer)
    CompactObject
End Sub

Private Sub IEWatch1_PathChange(Index As Integer, strPath As String)
    If mnuDisa(0).Checked Then
        If mnuDisa(1).Checked Then
            If Trim(strPath) <> "" Then
                If file_isFolder(strPath) Then
                   DoEvents
                   FindViriOnCurrentFolder strPath
                End If
            End If
        End If
    End If
End Sub

Function FindID(ID As Long) As Boolean
    On Error GoTo Salah
    Dim i As Integer
    For i = 0 To IEWatch1.count - 1
        If IEWatch1(i).IEKey = ID Then
            FindID = True
        End If
    Next i
Salah:
End Function

Private Sub lvwRTP_Click()
    On Error Resume Next
    txtValue(0) = ": " & lvwRTP.SelectedItem.SubItems(5)
    txtValue(1) = ": " & lvwRTP.SelectedItem.SubItems(1)
    txtValue(2) = ": " & lvwRTP.SelectedItem.SubItems(2)
    txtValue(3) = ": " & lvwRTP.SelectedItem.SubItems(4)
End Sub

Private Sub lvwRTP_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        If lvwRTP.ListItems.count > 0 Then
            PopupMenu mnuAction
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.show
End Sub

Private Sub mnuAct_Click(Index As Integer)
    Select Case Index
        Case 0: All
        Case 2: RemoveMessage
        Case 3: RemoveAll
        Case 5: Me.Hide
    End Select
End Sub

Private Sub mnuConsole_Click()
    frmConsole.show
End Sub

Private Sub mnuDisa_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            If mnuDisa(Index).Checked = False Then
                sTitle = "al VirusScan " 'Beta Version " & vAppVersion & " New"
                sMessage = "Realtime Protection enabled. " & al & " will protect your computer from viruses. "
                REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "RealtimeProtection", "Enabled"
                mnuDisa(Index).Checked = True
                SystrayOn Me, AVS
                PopupBalloon Me, sMessage, sTitle, NIIF_INFO
                isScannedOn = True
                KillVirusProcessList
            Else
                sTitle = "al VirusScan " 'Beta Version " & vAppVersion & " New"
                sMessage = "Realtime Protection disabled. Your computer may be at risk! "
                REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "RealtimeProtection", "Disabled"
                mnuDisa(Index).Checked = False
                SystrayOn Me, AVS
                PopupBalloon Me, sMessage, sTitle, NIIF_WARNING
                isScannedOn = False
            End If
            CompactObject
        Case 1
            If mnuDisa(Index).Checked = False Then
                mnuDisa(Index).Checked = True
                sTitle = "al VirusScan " 'Beta Version " & vAppVersion & " New"
                sMessage = "Monitoring directory enabled. " & al & " will protect your computer from viruses. "
                REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "MonitoringDirectory", "Enabled"
                SystrayOn Me, "Moh Aly Shodiqin"
                PopupBalloon Me, sMessage, sTitle, NIIF_INFO
            Else
                sTitle = "al VirusScan " 'Beta Version " & vAppVersion & " New"
                sMessage = "Monitoring directory disabled. Your computer may be at risk! "
                REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "MonitoringDirectory", "Disabled"
                SystrayOn Me, "Moh Aly Shodiqin"
                PopupBalloon Me, sMessage, sTitle, NIIF_WARNING
                mnuDisa(Index).Checked = False
            End If
            CompactObject
        Case 2
            ShellExecute Me.Hwnd, vbNullString, "mailto:felix_progressif@yahoo.com?subject=New Virus Sample...", vbNullString, "C:\", 1
        Case 5
            isTerminate = True
            StopAll = True
            SystrayOff Me
            Engine32.CloseScanHandle
            Set Engine32 = Nothing
            TerminateProcess GetCurrentProcess, 0 'End
        Case 7
            frmUpdate.show
        Case 9
            frmShutDown.show
        Case 11
            frmAbout.show
    End Select
End Sub

Private Sub mnuScan_Click()
    frmScanSample.show
End Sub

Private Sub mnuVirusScan_Click()
    frmScan.show
End Sub

Private Sub ShellIE_WindowRegistered(ByVal lCookie As Long)
    WatchIt
    Me.Caption = lCookie
End Sub

Sub CompactObject()
    On Error Resume Next
    isCompatch = True
    Dim i As Integer, Cnt As Integer
    For i = 0 To IEWatch1.count - 1
        IEWatch1(i).SetIENothing
    Next i
           
    Set ShellIE = Nothing
    For i = 1 To IEWatch1.count - 1
        Unload IEWatch1(i)
    Next i
       
    Set ShellIE = New SHDocVw.ShellWindows
    Cnt = ShellIE.count - 1
    For i = 0 To Cnt
        If i > 0 Then
            AddIEObj i
        End If
        'If IEWatch1(i).IEKey = 0 Then
              IEWatch1(i).AddSubClass ShellIE(i)
'        End If
    Next i
    isCompatch = False
End Sub

Sub CopyPlugin()
    On Error Resume Next
    Dim H As String
    H = Dir(nPath(MyWindowSys) & "comctl32.ocx", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If H = "" Then
        FileCopy nPath(App.path) & "\Data\comctl32.ocx", nPath(MyWindowSys) & "comctl32.ocx"
'        Shell "regsvr32 /s" & nPath(MyWindowSys) & "\comctl32.ocx", 0
    End If
    H = Dir(nPath(MyWindowSys) & "mscomct2.ocx", vbArchive + vbHidden + vbSystem + vbNormal + vbReadOnly)
    If H = "" Then
        FileCopy nPath(App.path) & "\Data\mscomct2.ocx", nPath(MyWindowSys) & "mscomct2.ocx"
'        Shell "regsvr32 /s" & nPath(MyWindowSys) & "\mscomct2.ocx", 0
    End If
    H = Dir(nPath(MyWindowSys) & "comdlg32.ocx", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If H = "" Then
        FileCopy nPath(App.path) & "\Data\comdlg32.ocx", nPath(MyWindowSys) & "comdlg32.ocx"
'        Shell "regsvr32 /s" & App.path & "\Data\comdlg32.ocx", 0
    End If
    H = Dir(nPath(MyWindowSys) & "mscomctl.ocx", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If H = "" Then
        FileCopy nPath(App.path) & "\Data\mscomctl.ocx", nPath(MyWindowSys) & "mscomctl.ocx"
'        Shell "regsvr32 /s" & nPath(MyWindowSys) & "\comdlg32.ocx", 0
    End If
End Sub

Sub RegShell()
    On Error Resume Next
    Dim myExe As String
    If LCase(Right(App.exename, 3)) = "exe" Then
       myExe = App.path & App.exename
    Else
       myExe = App.path & App.exename & ".exe"
    End If

    REG.SaveSettingString HKEY_CLASSES_ROOT, "*\Shell\al VirusScan", vbNullString, "Scan with al VirusScan..."
    REG.SaveSettingString HKEY_CLASSES_ROOT, "*\Shell\al VirusScan\command", vbNullString, myExe & " " & Chr(34) & "%1" & Chr(34)
    REG.SaveSettingString HKEY_CLASSES_ROOT, "Directory\Shell\al VirusScan", vbNullString, "Scan with al VirusScan..."
    REG.SaveSettingString HKEY_CLASSES_ROOT, "Directory\Shell\al VirusScan\command", vbNullString, myExe & " " & Chr(34) & "%1" & Chr(34)
'    REG.SaveSettingString HKEY_CLASSES_ROOT, "Drive\Shell\al VirusScan", vbNullString, "Scan with al VirusScan..."
'    REG.SaveSettingString HKEY_CLASSES_ROOT, "Drive\Shell\al VirusScan\command", vbNullString, myExe & " " & Chr(34) & "%1" & Chr(34)
    '-----------------------
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "Path", Chr(34) & App.path & "\al VirusScan.exe" & Chr(34)
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "alVirusScan", Chr(34) & App.path & "\al VirusScan.exe" & Chr(34) & " /RealtimeProtection"
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "SoundWarning", 1
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "ScanMemory", 1
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AllExtensions", IIf(frmConsole.optALL.Value, "ALL", "SELECTED")
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 0
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "Transparent", 0
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "RealtimeProtection", "Enabled"
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "MonitoringDirectory", "Enabled"
    frmConsole.chkOnTop.Value = 0
    frmConsole.chkTrans.Value = 0
End Sub

Sub All()
    Dim i As Integer
    With lvwRTP.ListItems
        For i = 1 To .count
            .Item(i).Selected = True
        Next i
    End With
End Sub

Sub RemoveAll()
    Dim i As Integer
    With lvwRTP.ListItems
        For i = 1 To .count
            If .Item(i).Selected Then
                .Item(i).Selected = False
                .Clear
                ClearText
                Exit Sub
            End If
        Next i
    End With
End Sub

Sub RemoveMessage()
    Dim i As Integer
    With lvwRTP.ListItems
        For i = 1 To .count
            If .Item(i).Selected Then
                .Item(i).Selected = False
                .Remove (i)
                If .count = 0 Then
                    ClearText
                End If
                Exit Sub
            End If
        Next i
    End With
End Sub

Sub ClearText()
    lblValue(5) = ": No item selected!"
    txtValue(0) = ":"
    txtValue(1) = ":"
    txtValue(2) = ":"
    txtValue(3) = ":"
End Sub
