VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmConsole 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6750
   ClipControls    =   0   'False
   Icon            =   "frmConsole.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5325
      TabIndex        =   44
      Top             =   3300
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox picConsole 
      BorderStyle     =   0  'None
      Height          =   2265
      Index           =   0
      Left            =   225
      ScaleHeight     =   2265
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   525
      Width           =   6315
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1515
         Left            =   3150
         ScaleHeight     =   1515
         ScaleWidth      =   2940
         TabIndex        =   18
         Top             =   375
         Width           =   2940
         Begin VB.CheckBox chkMemory 
            Caption         =   "Scan Memory"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   75
            TabIndex        =   22
            Top             =   1170
            Width           =   2865
         End
         Begin VB.CheckBox chkSound 
            Caption         =   "Sound Warning"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   75
            TabIndex        =   21
            Top             =   900
            Width           =   2865
         End
         Begin VB.CheckBox chkTrans 
            Caption         =   "Transparent"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   75
            TabIndex        =   20
            Top             =   345
            Width           =   2865
         End
         Begin VB.CheckBox chkOnTop 
            Caption         =   "Always On Top"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   75
            TabIndex        =   19
            Top             =   75
            Width           =   2865
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            Index           =   1
            X1              =   75
            X2              =   2850
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   0
            X1              =   75
            X2              =   2775
            Y1              =   750
            Y2              =   750
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Window Settings"
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
         Height          =   1815
         Left            =   3075
         TabIndex        =   17
         Top             =   150
         Width           =   3090
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   225
         ScaleHeight     =   765
         ScaleWidth      =   2640
         TabIndex        =   10
         Top             =   1125
         Width           =   2640
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
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
            Height          =   615
            Left            =   1800
            TabIndex        =   13
            Top             =   60
            Width           =   765
         End
         Begin VB.OptionButton optSelected 
            Caption         =   "Selected Types"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   75
            TabIndex        =   12
            Top             =   450
            Width           =   1665
         End
         Begin VB.OptionButton optALL 
            Caption         =   "All Types"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   75
            TabIndex        =   11
            Top             =   150
            Width           =   1665
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Type File"
         ClipControls    =   0   'False
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
         Height          =   1065
         Left            =   150
         TabIndex        =   9
         Top             =   900
         Width           =   2790
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   225
         ScaleHeight     =   390
         ScaleWidth      =   2565
         TabIndex        =   4
         Top             =   375
         Width           =   2565
         Begin VB.CheckBox chkStartup 
            Caption         =   "Run when windows start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   75
            TabIndex        =   5
            Top             =   75
            Width           =   2265
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Startup"
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
         Height          =   690
         Left            =   150
         TabIndex        =   6
         Top             =   150
         Width           =   2790
      End
      Begin VB.Label lblDate 
         Caption         =   "Virus Definitions Update"
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
         Left            =   150
         TabIndex        =   43
         Top             =   2025
         Width           =   1815
      End
      Begin VB.Label lblDefDate 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   2025
         TabIndex        =   42
         Top             =   2025
         Width           =   60
      End
   End
   Begin VB.PictureBox picConsole 
      BorderStyle     =   0  'None
      Height          =   2265
      Index           =   1
      Left            =   225
      ScaleHeight     =   2265
      ScaleWidth      =   6315
      TabIndex        =   3
      Top             =   525
      Width           =   6315
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore Defaults"
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
         Left            =   4650
         TabIndex        =   41
         ToolTipText     =   "All options will be restore..."
         Top             =   1800
         Width           =   1515
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   3225
         ScaleHeight     =   690
         ScaleWidth      =   2865
         TabIndex        =   39
         Top             =   375
         Width           =   2865
         Begin VB.CheckBox chkSafeMode 
            Caption         =   "Running On Safe Mode"
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
            Height          =   195
            Left            =   75
            TabIndex        =   40
            Top             =   75
            Value           =   1  'Checked
            Width           =   2415
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Other"
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
         Height          =   990
         Left            =   3150
         TabIndex        =   38
         Top             =   150
         Width           =   3015
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   690
         Left            =   225
         ScaleHeight     =   690
         ScaleWidth      =   2715
         TabIndex        =   34
         Top             =   1425
         Width           =   2715
         Begin VB.CommandButton cmdLanguage 
            Caption         =   "OK"
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
            Left            =   1800
            TabIndex        =   37
            Top             =   315
            Width           =   765
         End
         Begin VB.OptionButton optIndo 
            Caption         =   "Indonesia"
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
            Left            =   75
            TabIndex        =   36
            Top             =   390
            Width           =   1440
         End
         Begin VB.OptionButton optEnglish 
            Caption         =   "English"
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
            Left            =   75
            TabIndex        =   35
            Top             =   75
            Width           =   1440
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Language"
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
         Height          =   990
         Left            =   150
         TabIndex        =   33
         Top             =   1200
         Width           =   2865
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   225
         ScaleHeight     =   690
         ScaleWidth      =   2715
         TabIndex        =   25
         Top             =   375
         Width           =   2715
         Begin VB.CheckBox chkRar 
            Caption         =   "RAR"
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
            Height          =   195
            Left            =   75
            TabIndex        =   27
            Top             =   375
            Width           =   2415
         End
         Begin VB.CheckBox chkZip 
            Caption         =   "ZIP"
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
            Height          =   195
            Left            =   75
            TabIndex        =   26
            Top             =   75
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Archive Scan"
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
         Height          =   990
         Left            =   150
         TabIndex        =   24
         Top             =   150
         Width           =   2865
      End
   End
   Begin VB.PictureBox picConsole 
      BorderStyle     =   0  'None
      Height          =   2265
      Index           =   2
      Left            =   225
      ScaleHeight     =   2265
      ScaleWidth      =   6315
      TabIndex        =   14
      Top             =   525
      Width           =   6315
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Delete Permanently"
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
         Left            =   3975
         TabIndex        =   16
         Top             =   75
         Value           =   1  'Checked
         Width           =   2190
      End
      Begin ComctlLib.ListView lvwQuarantines 
         Height          =   1815
         Left            =   75
         TabIndex        =   15
         Top             =   375
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "No"
            Object.Width           =   422
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Virus Name"
            Object.Width           =   3068
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Old Location"
            Object.Width           =   7832
         EndProperty
      End
      Begin VB.Label lblViri 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   75
         TabIndex        =   23
         Top             =   105
         Width           =   2940
      End
   End
   Begin VB.PictureBox picConsole 
      BorderStyle     =   0  'None
      Height          =   2265
      Index           =   3
      Left            =   225
      ScaleHeight     =   2265
      ScaleWidth      =   6315
      TabIndex        =   7
      Top             =   525
      Width           =   6315
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   840
         Left            =   375
         ScaleHeight     =   840
         ScaleWidth      =   5265
         TabIndex        =   29
         Top             =   1050
         Width           =   5265
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse..."
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
            Height          =   360
            Left            =   4125
            TabIndex        =   32
            Top             =   375
            Width           =   1065
         End
         Begin VB.TextBox txtLog 
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
            Height          =   285
            Left            =   75
            TabIndex        =   31
            Top             =   375
            Width           =   3915
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "Log to file"
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
            Left            =   75
            TabIndex        =   30
            Top             =   75
            Width           =   2715
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Log file"
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
         Height          =   1140
         Left            =   225
         TabIndex        =   28
         Top             =   825
         Width           =   5565
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   75
         Picture         =   "frmConsole.frx":08CA
         Top             =   75
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Configure the logging of virus activity. Specify the information to be captured for each log entry."
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
         Index           =   1
         Left            =   900
         TabIndex        =   8
         Top             =   150
         Width           =   5115
      End
   End
   Begin ComctlLib.TabStrip tabConsole 
      Height          =   2715
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   4789
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Option"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Virus Quarantines"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   ""
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
   Begin ComctlLib.StatusBar sbConsole 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2970
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11855
            MinWidth        =   5293
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
   Begin VB.Menu mnuFile 
      Caption         =   "File "
      Begin VB.Menu mnuSetAttr 
         Caption         =   "Set Attribute File or Folder..."
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowDB 
         Caption         =   "Show Virus Definitions..."
      End
      Begin VB.Menu mnuB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit..."
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "Debug"
      Visible         =   0   'False
      Begin VB.Menu mnuQuarantines 
         Caption         =   "Select All"
         Index           =   0
      End
      Begin VB.Menu mnuQuarantines 
         Caption         =   "Unselect"
         Index           =   1
      End
      Begin VB.Menu mnuQuarantines 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuQuarantines 
         Caption         =   "Delete"
         Index           =   3
      End
      Begin VB.Menu mnuQuarantines 
         Caption         =   "Restore"
         Index           =   4
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "Tools"
      Begin VB.Menu mnuTools 
         Caption         =   "VirusScan Registry Tweak..."
         Index           =   0
      End
      Begin VB.Menu mnuTools 
         Caption         =   "VirusScan Process Manager..."
         Index           =   1
      End
      Begin VB.Menu mnuTools 
         Caption         =   "VirusScan Autorun Location..."
         Index           =   2
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTools 
         Caption         =   "VirusScan Registry Editor..."
         Index           =   4
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      Begin VB.Menu mnuSysWindows 
         Caption         =   "Console Windows"
         Index           =   0
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "System Configurations"
         Index           =   1
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "System Restore"
         Index           =   2
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "Task Manager"
         Index           =   4
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "Registry Editor"
         Index           =   5
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "Performance"
         Index           =   7
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "Event Viewer"
         Index           =   8
      End
      Begin VB.Menu mnuSysWindows 
         Caption         =   "Service"
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
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 29 Januari 2009
' 3:50 PM
'
' Update 7 Februari 2009 12:44 PM
'=======================================
' Module Console
'=======================================
Option Explicit

Dim col As Collection
Dim Ondata As Collection

Dim WithEvents Engine32 As cEngine32
Attribute Engine32.VB_VarHelpID = -1
Const sLog = "\Log\VirusScanLog.txt"

Sub LoadVirusZone()
    On Error Resume Next
    Dim H As String
    Set Ondata = New Collection
    lvwQuarantines.ListItems.Clear
    
    Set Engine32 = New cEngine32
    Engine32.ClassIDApartement = Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(1) & Chr(255)
    
    H = Dir(nPath(App.path) & "Quarantine\*.al", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If Trim(H) <> "" Then
        While H <> ""
            GetInfoFromFile nPath(App.path) & "Quarantine\" & H, Ondata
            H = Dir()
        Wend
    End If
    
    If Ondata.count Then
        Dim I As Long
        Dim l As ListItem
        Dim j As Integer
    
        For I = 1 To Ondata.count
            Set l = lvwQuarantines.ListItems.Add(, , Format(I, "0#"))
                l.SubItems(1) = UCase(Ondata(I)(0))
                l.SubItems(2) = UCase(Ondata(I)(1))
        Next I
        lblViri.Caption = "Virus In Quarantines : " & Ondata.count
    End If
End Sub

Sub GetInfoFromFile(Filename As String, ByRef data As Collection)
    On Error GoTo Salah
    Dim mark   As String
    mark = String(1024, 0)
    Open Filename For Binary Access Read As #1
        Get #1, , mark
    Close #1
           
    mark = Left(mark, InStr(1, mark, Chr(0) & Chr(0)) - 1)
    Dim nInfo() As String
    nInfo() = Split(mark, Chr(0))
    ReDim Preserve nInfo(UBound(nInfo) + 1) As String
    nInfo(UBound(nInfo)) = Filename
    data.Add nInfo
    Exit Sub
Salah:
    Close #1
End Sub

Private Sub chkLog_Click()
    If chkLog.Value = 1 Then
        txtLog.Enabled = True
        cmdBrowse.Enabled = True
        REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "VirusScanLog", App.path & sLog
    Else
        txtLog.Enabled = False
        cmdBrowse.Enabled = False
        REG.DeleteValue HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "VirusScanLog"
    End If
End Sub

Private Sub chkMemory_Click()
    If chkMemory.Value = 1 Then
        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "ScanMemory", 1
    Else
        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "ScanMemory", 0
    End If
End Sub

Private Sub chkOnTop_Click()
    If chkOnTop.Value = 1 Then
        AlwaysOnTop Me.Hwnd, True
        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1
    Else
        AlwaysOnTop Me.Hwnd, False
        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 0
    End If
End Sub

Private Sub chkSafeMode_Click()
    On Error Resume Next
'    If chkSafeMode.Value = 1 Then
'        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "SafeMode", 1
'        REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe " & Chr(34) & App.path & "\al VirusScan.exe" & Chr(34)
'    Else
'        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "SafeMode", 0
'        REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
'        chkSafeMode.Value = 0
'    End If
End Sub

Private Sub chkSound_Click()
    If chkSound.Value = 1 Then
        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "SoundWarning", 1
    Else
        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "SoundWarning", 0
    End If
End Sub

Private Sub chkStartup_Click()
    If chkStartup.Value = 1 Then
        REG.SaveSettingByte HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "RunOnStartup", 1
        REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "alVirusScan", Chr(34) & App.path & "\al VirusScan.exe" & Chr(34) & " /RealtimeProtection"
'        REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "alVirusScan", App.path & "\al VirusScan.exe"
    Else
        REG.SaveSettingByte HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "RunOnStartup", 0
        REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "alVirusScan"
        chkStartup.Value = 0
    End If
End Sub

Private Sub chkTrans_Click()
    If chkTrans.Value = 1 Then
        SetOpagueForm False, Me
        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "Transparent", 1
    Else
        SetOpagueForm True, Me
        REG.SaveSettingLong HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "Transparent", 0
    End If
End Sub

Private Sub cmdBrowse_Click()
    txtLog = ShowSave(Me.Hwnd, "Log files|*.txt")
'    SetMySetting "Report", "LogFile", txtLog
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AllExtensions", IIf(optALL.Value, "ALL", "SELECTED")
    Dim I As Integer, buff As String
    For I = 0 To lstType.ListCount - 1
        buff = buff & lstType.List(I) & "|"
    Next I
    If Right(buff, 1) = "|" Then
        buff = Left(buff, Len(buff) - 1)
    End If
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "ExtensionsSelected", buff
    cmdOK.Enabled = False
End Sub

Private Sub cmdRestore_Click()
    With Me
        .chkLog.Value = 1
        .chkMemory.Value = 1
        .chkOnTop.Value = 0
        .chkSafeMode.Value = 1
        .chkSound.Value = 1
        .chkStartup.Value = 1
        .chkTrans.Value = 0
        .optALL.Value = True
        REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AllExtensions", IIf(optALL.Value, "ALL", "SELECTED")
        REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "RealtimeProtection", "Enabled"
        REG.SaveSettingString HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan", "MonitoringDirectory", "Enabled"
        .cmdOK.Enabled = False
        frmRTP.mnuDisa(0).Checked = True
        frmRTP.mnuDisa(1).Checked = True
    End With
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim H As String
    Dim I As Integer
    Dim data() As String
    
    lvwStyle lvwQuarantines
    LoadVirusZone
    VDFDate
    lblDefDate = ": " & vVirusDefinitions
'    lblDate(2) = "" '"<" & DateDiff("d", vVirusDefinitions, Date) & " days old.>"
    If CDate(Month(vVirusDefinitions)) < Month(Date) Then
'        lblDefDate.ForeColor = vbBlue
        lblDefDate.ToolTipText = "It is requiered to update your virus definitions..."
        LogScan "Virus Definitions Update " & vbTab & vVirusDefinitions & " It is requiered to update your virus definitions..."
    End If
        
    'AllExtensions
    If REG.GetSettingString(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AllExtensions", "ALL") = "ALL" Then
       optALL.Value = True
    Else
       optSelected.Value = True
       cmdOK.Enabled = False
    End If
    
    'ExtensionsSelected
    H = REG.GetSettingString(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "ExtensionsSelected", "386|BAT|BIN|BTM|CLA|COM|CSC|DLL|DRV|EXE|EX_|OCX|OV?|PIF|SYS|VXD|CSH|DOC|DOT|HLP|HTA|HTM|HTML|HTT|INF|INI|JS|JSE|JTD|MDB|MP?|MSO|ODB|OBT|PL|PM|POT|PPS|PPT|RTF|SH|SHB|SHS|SMM|VBE|VBS|VSD|VSS|VST|WSF|WSH|XLA|XLS|SCR|SC_|TMP|JPG|REG")
    data() = Split(H, "|", , vbTextCompare)
    
    For I = 0 To UBound(data)
        lstType.AddItem data(I)
    Next I
    
    'AlwaysOnTop
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        chkOnTop.Value = Checked
    Else
        chkOnTop.Value = Unchecked
    End If
    
    'Transparent
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "Transparent", 1) = 1 Then
        chkTrans.Value = Checked
    Else
        chkTrans.Value = Unchecked
    End If
    
    'SoundWarning
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "SoundWarning", 1) = 1 Then
        chkSound.Value = Checked
    Else
        chkSound.Value = Unchecked
    End If

    'ScanMemory
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "ScanMemory", 1) = 1 Then
        chkMemory.Value = Checked
    Else
        chkMemory.Value = Unchecked
    End If
  
    'Startup
    If REG.GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "alVirusScan") = Chr(34) & App.path & "\al VirusScan.exe" & Chr(34) & " /RealtimeProtection" Then
        chkStartup.Value = Checked
    Else
        chkStartup.Value = Unchecked
    End If
    
    'SafeMode
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "SafeMode", 1) = 1 Then
        chkSafeMode.Value = Checked
    Else
        chkSafeMode.Value = Unchecked
    End If
    GetUserCom
    cmdOK.Enabled = False
    txtLog = App.path & sLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Engine32 = Nothing
End Sub

Private Sub lvwQuarantines_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvwQuarantines.Sorted = True
    lvwQuarantines.SortKey = ColumnHeader.Index - 1
    If lvwQuarantines.SortOrder = lvwDescending Then
       lvwQuarantines.SortOrder = lvwAscending
    Else
       lvwQuarantines.SortOrder = lvwDescending
    End If
End Sub

Private Sub lvwQuarantines_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        If lvwQuarantines.ListItems.count > 0 Then
            PopupMenu mnuDebug
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuQuarantines_Click(Index As Integer)
    Select Case Index
        Case 0: SelectAll
        Case 1: Unselect
        Case 3
            If MsgBox("Are you sure want to delete quarantine files?", vbYesNo + 32, "Confirm") = vbYes Then
                DoDeleteVirus
            End If
        Case 4
            DoRestoreVirus
    End Select
End Sub

Private Sub mnuSetAttr_Click()
    frmAttr.show
End Sub

Private Sub mnuShowDB_Click()
    frmDatabase.show , Me
End Sub

Private Sub mnuSysWindows_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            ShellExecute Me.Hwnd, vbNullString, "cmd.exe", vbNullString, "C:\", 1
        Case 1
            ShellExecute Me.Hwnd, vbNullString, "msconfig.exe", vbNullString, "C:\", 1
        Case 2
            ShellExecute Me.Hwnd, vbNullString, MyWindowSys & "restore\rstrui.exe", vbNullString, "C:\", 1
        Case 4
            ShellExecute Me.Hwnd, vbNullString, "taskmgr.exe", vbNullString, "C:\", 1
        Case 5
            ShellExecute Me.Hwnd, vbNullString, "regedit.exe", vbNullString, "C:\", 1
        Case 7
            ShellExecute Me.Hwnd, vbNullString, "perfmon.msc", vbNullString, "C:\", 1
        Case 8
            ShellExecute Me.Hwnd, vbNullString, "eventvwr.msc", vbNullString, "C:\", 1
        Case 9
            ShellExecute Me.Hwnd, vbNullString, "services.msc", vbNullString, "C:\", 1
    End Select
End Sub

Private Sub mnuTools_Click(Index As Integer)
    Select Case Index
        Case 0: frmTweak.show
        Case 1: frmProcess.show
        Case 2: frmAutorun.show
        Case 4
            ShellExecute Me.Hwnd, vbNullString, nPath(App.path) & "\Data\avsregedit.exe", vbNullString, "C:\", 1
    End Select
End Sub

Private Sub optALL_Click()
    If optALL.Value = True Then
        cmdOK.Enabled = True
    End If
End Sub

Private Sub optSelected_Click()
    If optSelected.Enabled = True Then
        cmdOK.Enabled = True
    End If
End Sub

Private Sub tabConsole_Click()
    Dim pic As PictureBox
    For Each pic In picConsole
        pic.Visible = (pic.Index = tabConsole.SelectedItem.Index - 1)
    Next
End Sub

Private Sub SelectAll()
    Dim I As Integer
    With lvwQuarantines.ListItems
        For I = 1 To .count
            .Item(I).Selected = True
        Next I
    End With
End Sub

Private Sub Unselect()
    Dim I As Integer
    With lvwQuarantines.ListItems
        For I = 1 To .count
            .Item(I).Selected = False
        Next I
    End With
End Sub

Sub DoRestoreVirus()
    On Error Resume Next
    lvwQuarantines.Enabled = False
    Check1.Enabled = False
    
    Dim I As Integer
    For I = 1 To lvwQuarantines.ListItems.count
        If lvwQuarantines.ListItems(I).Selected Then
            Engine32.RestoreFiles CStr(Ondata(I)(2)), Me.Hwnd
            lvwQuarantines.ListItems(I).Selected = False
            Sleep 100
        End If
    Next I
    
    lvwQuarantines.Enabled = True
    Check1.Enabled = True
    
    LoadVirusZone
End Sub

Sub DoDeleteVirus()
    On Error Resume Next
    lvwQuarantines.Enabled = False
    Check1.Enabled = False
    
    Dim I As Integer
    For I = 1 To lvwQuarantines.ListItems.count
        If lvwQuarantines.ListItems(I).Selected Then
            If Check1.Value = 1 Then
                Kill CStr(Ondata(I)(2))
            Else
                Call ShellWinFile(FO_DELETE, 0, CStr(Ondata(I)(2)))
            End If
            lvwQuarantines.ListItems(I).Selected = False
        End If
    Next I
    lvwQuarantines.Enabled = True
    Check1.Enabled = True
    
    LoadVirusZone
End Sub
