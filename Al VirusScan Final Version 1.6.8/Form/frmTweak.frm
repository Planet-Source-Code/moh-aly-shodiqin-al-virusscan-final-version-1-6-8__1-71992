VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmTweak 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   Icon            =   "frmTweak.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTweak 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   5025
      Top             =   5850
   End
   Begin ComctlLib.StatusBar sbTweak 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   50
      Top             =   6390
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14270
            MinWidth        =   7920
            TextSave        =   ""
            Key             =   ""
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
   Begin VB.CommandButton cmdTweak 
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
      Index           =   3
      Left            =   6765
      TabIndex        =   49
      Tag             =   "Exit from tweak"
      Top             =   5850
      Width           =   1215
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "Fix Registry"
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
      TabIndex        =   48
      Tag             =   "Repair system registry"
      Top             =   5850
      Width           =   1215
   End
   Begin alVirusScan.AdvProgressBar pbFix 
      Height          =   240
      Left            =   150
      TabIndex        =   47
      Top             =   5475
      Width           =   7815
      _extentx        =   13785
      _extenty        =   423
      value           =   0
      barcolor1       =   -2147483634
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   225
      ScaleHeight     =   1290
      ScaleWidth      =   3690
      TabIndex        =   36
      Top             =   3975
      Width           =   3690
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable the Tools / Internet Options menu"
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
         Height          =   240
         Index           =   30
         Left            =   75
         TabIndex        =   41
         Tag             =   "NoBrowserOptions"
         Top             =   1035
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable of selecting a download directory"
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
         Height          =   240
         Index           =   29
         Left            =   75
         TabIndex        =   40
         Tag             =   "NoBrowserOptions"
         Top             =   795
         Width           =   3555
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable the Tools / Internet Options menu"
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
         Height          =   240
         Index           =   28
         Left            =   75
         TabIndex        =   39
         Tag             =   "NoBrowserOptions"
         Top             =   555
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable right-click context menu"
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
         Height          =   240
         Index           =   27
         Left            =   75
         TabIndex        =   38
         Tag             =   "NoBrowserContextMenu"
         Top             =   315
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable option of closing Internet Explorer"
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
         Height          =   240
         Index           =   26
         Left            =   75
         TabIndex        =   37
         Tag             =   "NoBrowserClose"
         Top             =   75
         Width           =   3540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Internet Explorer Security Restrictions"
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
      Height          =   1590
      Left            =   150
      TabIndex        =   35
      Top             =   3750
      Width           =   3840
   End
   Begin VB.CommandButton cmdTweak 
      Caption         =   "Apply"
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
      Index           =   0
      Left            =   150
      TabIndex        =   29
      Tag             =   "Apply tweak settings"
      Top             =   5850
      Width           =   1215
   End
   Begin VB.CommandButton cmdTweak 
      Caption         =   "Cek All"
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
      Index           =   1
      Left            =   1440
      TabIndex        =   28
      Tag             =   "Select all of tweak."
      Top             =   5850
      Width           =   1215
   End
   Begin VB.CommandButton cmdTweak 
      Caption         =   "Clear All"
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
      Index           =   2
      Left            =   2730
      TabIndex        =   27
      Tag             =   "Clear all of tweak."
      Top             =   5850
      Width           =   1215
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   225
      ScaleHeight     =   3165
      ScaleWidth      =   3690
      TabIndex        =   12
      Top             =   375
      Width           =   3690
      Begin VB.CheckBox chkSystem 
         Caption         =   "Show extensions for known file types"
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
         Height          =   240
         Index           =   16
         Left            =   75
         TabIndex        =   25
         Tag             =   "HideFileExt"
         Top             =   2910
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Show Hidden Folders And Files "
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
         Height          =   240
         Index           =   15
         Left            =   75
         TabIndex        =   24
         Tag             =   "Hidden "
         Top             =   2670
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Show File Hidden Operating System "
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
         Height          =   240
         Index           =   14
         Left            =   75
         TabIndex        =   23
         Tag             =   "ShowSuperHidden "
         Top             =   2430
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Properties My Computer"
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
         Height          =   240
         Index           =   13
         Left            =   75
         TabIndex        =   22
         Tag             =   "NoPropertiesMyComputer"
         Top             =   2190
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Hide And Support"
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
         Height          =   240
         Index           =   12
         Left            =   75
         TabIndex        =   21
         Tag             =   "NoSMHelp"
         Top             =   1950
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Registry Editor Tools"
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
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   20
         Tag             =   "DisableRegistryTools"
         Top             =   270
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Folder Options Menu"
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
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   19
         Tag             =   "NoFolderOptions"
         Top             =   510
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Menu Find"
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
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   18
         Tag             =   "NoFind"
         Top             =   750
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Menu Run"
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
         Height          =   240
         Index           =   4
         Left            =   75
         TabIndex        =   17
         Tag             =   "NoRun"
         Top             =   990
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Right-click on Desktop"
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
         Height          =   240
         Index           =   5
         Left            =   75
         TabIndex        =   16
         Tag             =   "NoViewContextMenu"
         Top             =   1230
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Show Windows Version on Desktop"
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
         Height          =   240
         Index           =   6
         Left            =   75
         TabIndex        =   15
         Tag             =   "PaintDesktopVersion"
         Top             =   1470
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Display Properties"
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
         Height          =   240
         Index           =   7
         Left            =   75
         TabIndex        =   14
         Tag             =   "NoDispCPL"
         Top             =   1710
         Width           =   3465
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Task Manager"
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
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   13
         Tag             =   "DisableTaskMgr"
         Top             =   75
         Width           =   3465
      End
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   4200
      ScaleHeight     =   1065
      ScaleWidth      =   3690
      TabIndex        =   6
      Top             =   375
      Width           =   3690
      Begin VB.CheckBox chkSystem 
         Caption         =   "Hide the Display Settings Page "
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
         Height          =   240
         Index           =   11
         Left            =   75
         TabIndex        =   10
         Tag             =   "NoDispSettingsPage"
         Top             =   795
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Hide the Screen Saver Settings Page "
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
         Height          =   240
         Index           =   10
         Left            =   75
         TabIndex        =   9
         Tag             =   "NoDispScrSavPage"
         Top             =   555
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Hide the Display Background Page "
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
         Height          =   240
         Index           =   9
         Left            =   75
         TabIndex        =   8
         Tag             =   "NoDispBackgroundPage"
         Top             =   315
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Hide the Display Appearance Page "
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
         Height          =   240
         Index           =   8
         Left            =   75
         TabIndex        =   7
         Tag             =   "NoDispAppearancePage"
         Top             =   75
         Width           =   3405
      End
   End
   Begin VB.PictureBox Picture8 
      BorderStyle     =   0  'None
      Height          =   3465
      Left            =   4275
      ScaleHeight     =   3465
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   1800
      Width           =   3615
      Begin VB.CheckBox chkSystem 
         Caption         =   "Show Full Path at Address Bar"
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
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   46
         Tag             =   "FullPathAddress"
         Top             =   3210
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Hide the File System Button "
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
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   45
         Tag             =   "NoFileSysPage"
         Top             =   2970
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Hide the Device Manager Page "
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
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   44
         Tag             =   "NoDevMgrPage"
         Top             =   2730
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Remove File menu from Explorer"
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
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   43
         Tag             =   "NoFileMenu"
         Top             =   2490
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Remove the Tildes in Short Filenames ""~"""
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
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   42
         Tag             =   "NameNumericTail"
         Top             =   2250
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Explorer's default context menu "
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
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   34
         Tag             =   "NoViewContextMenu"
         Top             =   1995
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Don't Save Settings at Exit "
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
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   33
         Tag             =   "NoSaveSettings"
         Top             =   1755
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Control Panel"
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
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   32
         Tag             =   "NoControlPanel"
         Top             =   1035
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Remove Username from the Start Menu"
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
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   31
         Tag             =   "NoUserNameInStartMenu"
         Top             =   1275
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Command Prompt and Batch Files"
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
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   30
         Tag             =   "DisableCMD"
         Top             =   1515
         Width           =   3405
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable System Tray "
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
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   4
         Tag             =   "NoTrayItemsDisplay"
         Top             =   795
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable Context Menus For the Taskbar"
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
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   3
         Tag             =   "NoTrayContextMenu"
         Top             =   555
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Hide the Network Neighborhood Icon"
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
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   2
         Tag             =   "NoNetHood"
         Top             =   315
         Width           =   3390
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Disable the Shut Down Command"
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
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   1
         Tag             =   "NoClose"
         Top             =   75
         Width           =   3390
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "System"
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
      Height          =   3540
      Left            =   150
      TabIndex        =   26
      Top             =   150
      Width           =   3840
   End
   Begin VB.Frame Frame14 
      Caption         =   "Windows Security Settings"
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
      Height          =   3765
      Left            =   4125
      TabIndex        =   5
      Top             =   1575
      Width           =   3840
   End
   Begin VB.Frame Frame13 
      Caption         =   "Display Properties Restrictions"
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
      Height          =   1365
      Left            =   4125
      TabIndex        =   11
      Top             =   150
      Width           =   3840
   End
End
Attribute VB_Name = "frmTweak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 6 Februari 2009
' 1:24 PM
' Tweak v1.2
'=======================================
' Module Update
'=======================================
Option Explicit

Public CekSetting As Boolean, cekLoad As Boolean
Dim X As Integer

Private Sub chkSystem_Click(Index As Integer)
    On Error Resume Next
    If cekLoad = True Then
        CekSetting = True
        cmdTweak(0).Enabled = True
        cmdTweak(0).Caption = "Apply"
    End If
End Sub

Sub Apply()
    SaveApp
    cmdTweak(0).Enabled = False
    cmdTweak(0).Caption = "No Changes"
    LockWindowUpdate (GetDesktopWindow())
    ForceCacheRefresh
    LockWindowUpdate (0)
End Sub

Sub Clear()
    Dim I As Integer
    On Error Resume Next
    With chkSystem
        For I = 0 To .count
            .Item(I).Value = 0
        Next I
    End With
End Sub

Sub Cek()
    Dim I As Integer
    On Error Resume Next
    With chkSystem
        For I = 0 To .count
            .Item(I).Value = 1
        Next I
    End With
End Sub

Private Sub chkSystem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbTweak.Panels(1).Text = chkSystem(Index).Caption
End Sub

Private Sub cmdFix_Click()
    If tmrTweak.Enabled = False Then
        If MsgBox("Are you sure want to repair registry ?", vbExclamation + vbYesNo, "al VirusScan") = vbYes Then
            sbTweak.Panels(1).Text = "Please wait take a few moment..."
            LockControl False
            tmrTweak.Enabled = True
            pbFix.Value = 0
        End If
    Else
        tmrTweak.Enabled = False
    End If
End Sub

Private Sub cmdFix_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbTweak.Panels(1).Text = cmdFix.Tag
End Sub

Private Sub cmdTweak_Click(Index As Integer)
    Select Case Index
        Case 0: Apply
        Case 1: Cek
        Case 2: Clear
        Case 3: Unload Me
    End Select
End Sub

Private Sub cmdTweak_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbTweak.Panels(1).Text = cmdTweak(Index).Tag
End Sub

Private Sub Form_Load()
    Me.Caption = "VirusScan Tweak Registry"
    cmdTweak(0).Enabled = False
    cekLoad = False
    CekSetting = False
    GetApp
    cekLoad = True
    
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.Hwnd, True
    Else
        AlwaysOnTop Me.Hwnd, False
    End If
    
    pbFix.Style = SmoothDoubleColor
'    sbTweak.Panels(2).Text = Copyright
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbTweak.Panels(1).Text = ""
End Sub

Private Sub tmrTweak_Timer()
    If pbFix.Value >= pbFix.Max Then
        tmrTweak.Enabled = False
        FixRegistry
        sbTweak.Panels(1).Text = "Registry have repairing by al VirusScan."
        LockControl True
        pbFix.Value = 0
    Else
        pbFix.Value = pbFix.Value + 1
    End If
End Sub

Sub LockControl(bLock As Boolean)
    cmdTweak(0).Enabled = False
    cmdTweak(1).Enabled = bLock
    cmdTweak(2).Enabled = bLock
    cmdTweak(3).Enabled = bLock
    cmdFix.Enabled = bLock
    Picture1.Enabled = bLock
    Picture6.Enabled = bLock
    Picture7.Enabled = bLock
    Picture8.Enabled = bLock
End Sub
