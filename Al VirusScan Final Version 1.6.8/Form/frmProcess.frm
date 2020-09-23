VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7440
   Icon            =   "frmProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProc 
      BorderStyle     =   0  'None
      Height          =   4890
      Index           =   0
      Left            =   225
      ScaleHeight     =   4890
      ScaleWidth      =   6990
      TabIndex        =   2
      Top             =   450
      Width           =   6990
      Begin ComctlLib.ListView lvwProcess 
         Height          =   3090
         Left            =   150
         TabIndex        =   3
         Top             =   225
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   5450
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Image Name"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "User Name"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Location"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "PID"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Threads"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   1129
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Memory"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Attributes"
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   8
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Priority"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   9
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   1365
         Left            =   150
         TabIndex        =   4
         Top             =   3375
         Width           =   6690
         Begin VB.Timer Timer2 
            Interval        =   10000
            Left            =   0
            Top             =   900
         End
         Begin VB.PictureBox picProcess 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   150
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   6
            Top             =   225
            Width           =   540
         End
         Begin VB.Label lblValue 
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
            Left            =   900
            TabIndex        =   12
            Top             =   225
            Width           =   5640
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
            Index           =   1
            Left            =   900
            TabIndex        =   11
            Top             =   465
            Width           =   5640
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
            Index           =   2
            Left            =   900
            TabIndex        =   10
            Top             =   750
            Width           =   690
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
            Index           =   3
            Left            =   900
            TabIndex        =   9
            Top             =   990
            Width           =   690
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
            Index           =   4
            Left            =   1575
            TabIndex        =   8
            Top             =   750
            Width           =   4965
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
            Index           =   5
            Left            =   1575
            TabIndex        =   7
            Top             =   990
            Width           =   4965
         End
      End
   End
   Begin VB.PictureBox picProc 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   4890
      Index           =   1
      Left            =   225
      ScaleHeight     =   4890
      ScaleWidth      =   6990
      TabIndex        =   5
      Top             =   450
      Width           =   6990
      Begin ComctlLib.ProgressBar Prog 
         Height          =   225
         Left            =   1275
         TabIndex        =   52
         Top             =   4350
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmdOpt 
         Caption         =   "Optimize "
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
         Left            =   5775
         TabIndex        =   49
         ToolTipText     =   "RAM Optimizer"
         Top             =   4305
         Width           =   1065
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   225
         Left            =   150
         TabIndex        =   48
         ToolTipText     =   "Level 0-1"
         Top             =   4350
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   397
         _Version        =   327682
         Max             =   1
      End
      Begin VB.Frame Frame6 
         Caption         =   "Totals"
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
         Left            =   150
         TabIndex        =   40
         Top             =   1650
         Width           =   3240
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2025
            TabIndex        =   46
            Top             =   780
            Width           =   1065
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2025
            TabIndex        =   45
            Top             =   540
            Width           =   1065
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2025
            TabIndex        =   44
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Processes"
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
            Left            =   150
            TabIndex        =   43
            Top             =   780
            Width           =   1890
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
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
            Left            =   150
            TabIndex        =   42
            Top             =   540
            Width           =   1890
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Handles"
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
            Left            =   150
            TabIndex        =   41
            Top             =   300
            Width           =   1890
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "CPU Usage"
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
         Height          =   1440
         Left            =   150
         TabIndex        =   37
         Top             =   75
         Width           =   1290
         Begin VB.PictureBox picUsage 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H0000C000&
            Height          =   975
            Left            =   150
            ScaleHeight     =   61
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   62
            TabIndex        =   38
            Top             =   300
            Width           =   990
            Begin VB.Label lblCpuUsage 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   255
               Left            =   0
               TabIndex        =   39
               Top             =   675
               Width           =   930
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "CPU Usage History"
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
         Height          =   1440
         Left            =   1725
         TabIndex        =   35
         Top             =   75
         Width           =   5115
         Begin VB.PictureBox picGraph 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H00008000&
            Height          =   975
            Left            =   150
            ScaleHeight     =   100
            ScaleMode       =   0  'User
            ScaleWidth      =   99
            TabIndex        =   36
            Top             =   300
            Width           =   4815
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Kernel Memory"
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
         Left            =   3675
         TabIndex        =   28
         Top             =   2850
         Width           =   3240
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2100
            TabIndex        =   34
            Top             =   780
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2100
            TabIndex        =   33
            Top             =   540
            Width           =   990
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Paged"
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
            Left            =   150
            TabIndex        =   32
            Top             =   540
            Width           =   1665
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Non Paged"
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
            Left            =   150
            TabIndex        =   31
            Top             =   780
            Width           =   1665
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2100
            TabIndex        =   30
            Top             =   300
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   150
            TabIndex        =   29
            Top             =   300
            Width           =   1665
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Commit Charge "
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
         Left            =   150
         TabIndex        =   21
         Top             =   2850
         Width           =   3240
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Peak"
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
            Left            =   150
            TabIndex        =   27
            Top             =   780
            Width           =   1890
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Limit"
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
            Left            =   150
            TabIndex        =   26
            Top             =   540
            Width           =   1890
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Total "
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
            Left            =   150
            TabIndex        =   25
            Top             =   300
            Width           =   1890
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2025
            TabIndex        =   24
            Top             =   540
            Width           =   1065
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2025
            TabIndex        =   23
            Top             =   780
            Width           =   1065
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2025
            TabIndex        =   22
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Physical Memory"
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
         Left            =   3675
         TabIndex        =   14
         Top             =   1650
         Width           =   3240
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2025
            TabIndex        =   20
            Top             =   780
            Width           =   1065
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   2025
            TabIndex        =   19
            Top             =   540
            Width           =   1065
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   2025
            TabIndex        =   18
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Total "
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
            Left            =   150
            TabIndex        =   17
            Top             =   300
            Width           =   1665
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Available "
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
            Left            =   150
            TabIndex        =   16
            Top             =   540
            Width           =   1890
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "System Cache"
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
            Left            =   150
            TabIndex        =   15
            Top             =   780
            Width           =   1665
         End
      End
      Begin VB.Label lj 
         AutoSize        =   -1  'True
         Caption         =   "Progress:"
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
         Left            =   1275
         TabIndex        =   51
         Top             =   4650
         Width           =   705
      End
      Begin VB.Label ry 
         AutoSize        =   -1  'True
         Caption         =   "Done!"
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
         Left            =   2055
         TabIndex        =   50
         Top             =   4650
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "Please choose optimize level and then click 'Optimize' button to start optimization."
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
         Left            =   150
         TabIndex        =   47
         Top             =   4050
         Width           =   6615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   7800
      TabIndex        =   13
      Top             =   300
      Width           =   2115
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1575
         Top             =   225
      End
      Begin VB.Timer tmrMem 
         Interval        =   500
         Left            =   675
         Top             =   225
      End
      Begin VB.Timer tmrRefresh 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1125
         Top             =   225
      End
      Begin ComctlLib.ImageList ilsProcess 
         Left            =   75
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
   Begin ComctlLib.TabStrip tabProcess 
      Height          =   5340
      Left            =   150
      TabIndex        =   1
      Top             =   75
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   9419
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Processes"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Performance"
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
   Begin ComctlLib.StatusBar sbProcess 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5505
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2717
            MinWidth        =   2717
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3598
            MinWidth        =   3598
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3422
            MinWidth        =   3422
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3254
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
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewTask 
         Caption         =   "New Task (Run...)"
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit..."
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuEndProcess 
         Caption         =   "End Process"
      End
      Begin VB.Menu mnuB 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh Process"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "New Task (Run...)"
      End
      Begin VB.Menu mnuC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuThreads 
         Caption         =   "Threads"
         Begin VB.Menu mnuProcess 
            Caption         =   "Suspend Process"
            Index           =   0
         End
         Begin VB.Menu mnuProcess 
            Caption         =   "Resume Process"
            Index           =   1
         End
      End
      Begin VB.Menu mnuSetPrio 
         Caption         =   "Set Process Priority"
         Begin VB.Menu mnuSetPriority 
            Caption         =   "&1. Realtime Priority"
            Index           =   1
         End
         Begin VB.Menu mnuSetPriority 
            Caption         =   "&2. High Priority"
            Index           =   2
         End
         Begin VB.Menu mnuSetPriority 
            Caption         =   "&3. Normal Priority"
            Index           =   3
         End
         Begin VB.Menu mnuSetPriority 
            Caption         =   "&4. Idle Priority"
            Index           =   4
         End
      End
      Begin VB.Menu mnuD 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "Show Details..."
      End
      Begin VB.Menu mnuFindFile 
         Caption         =   "Find File Location"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuE 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Scanning Process..."
      End
   End
   Begin VB.Menu mnuShut 
      Caption         =   "Shut Down"
      Begin VB.Menu mnuShutDown 
         Caption         =   "Turn Off Computer"
         Index           =   0
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "Restart"
         Index           =   1
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "Log Off Windows"
         Index           =   2
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "Power Off"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 3 Februari 2009
' 12:28 AM
'=======================================
' Module Process Manager®
'=======================================
Option Explicit

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

Dim WithEvents Engine32 As cEngine32
Attribute Engine32.VB_VarHelpID = -1
Dim FileOnScan As Double
Dim FileViruses As Double
Dim FailedFile As Double
Dim CleanedFile As Double

Dim memory&
Dim FreeMEM
Dim OptS

Private shinfo As SHFILEINFO
Private QueryObject As Object

Private Sub cmdEndProcess_Click()
    Dim i As Integer
    Dim Pesan As String, strFile As String
    Dim lExitCode As Long
    
    Pesan = "WARNING: Terminating a process can cause undesired" & vbCrLf & _
            "results including loss of data and system instability. The" & vbCrLf & _
            "process will not be given the chance to save its state or" & vbCrLf & _
            "data before it is terminated. Are you sure you want to" & vbCrLf & _
            "terminate the process?"
    If MsgBox(Pesan, vbYesNo + 48, "Process Manager Warning" & Chr(0)) = vbYes Then
        lExitCode = TerminateProcessID(lvwProcess, 5)
        If lExitCode = 0 Then MsgBox "Cannot terminate this process.", vbExclamation, "Unable To Terminate Process"
        Call mnuRefresh_Click
    End If
End Sub

Private Sub cmdOpt_Click()
    If MsgBox("Optimize now...", vbExclamation + vbYesNo, "RAM Optimizer") = vbYes Then
        Timer1.Enabled = False
        cmdOpt.Enabled = False
        Select Case Slider1.Value
            Case 0
                OptS = 1000000
            Case 1
                OptS = 5000000
        End Select
        Call OptimizeRAM
    Else
        Exit Sub
    End If
End Sub

Private Sub Engine32_onVirusFound(nFileName As String, nFileInfo As cFileInfo)
    On Error Resume Next
    FileViruses = FileViruses + 1
    ViriOnCollect.Add nFileInfo
    Select Case UCase(nFileInfo.VirusAction)
        Case "DELETE"
            If nFileInfo.VirusClean Then
                CleanedFile = CleanedFile + 1
                VirusAlert
            Else
                FailedFile = FailedFile + 1
                VirusAlert
            End If
        Case "QUARANTINE", "BUNDLE"
            If nFileInfo.VirusClean Then
                CleanedFile = CleanedFile + 1
                VirusAlert
            Else
                FailedFile = FailedFile + 1
                VirusAlert
            End If
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = "VirusScan Process Manager"
    lvwStyle lvwProcess
    ProcessList lvwProcess, ilsProcess

    Set Engine32 = New cEngine32
    Engine32.ClassIDApartement = Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(1) & Chr(255)

    FileOnScan = 0
    FileViruses = 0
    CleanedFile = 0
    FailedFile = 0
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.hWnd, True
    Else
        AlwaysOnTop Me.hWnd, False
    End If
'    GetCPUInfo Me.sbProcess
'    sbProcess.Panels(2).Text = frmDetail.lblCpuUsage.Caption = CStr(Ret) + "%"
    '-----------------------------------------------------------------
    'set the Priority of this process to 'High'
    'this makes sure our program gets updated, even when
    'another process is consuming lots of CPU cycles
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'Initialize our Query object
    If IsWinNT Then
        Set QueryObject = New cCPUUsageNT
    Else
        Set QueryObject = New cCPUUsage
    End If
    'Initializing is necesarry for the correct values to be retrieved
    QueryObject.Initialize
    'start the timer
    tmrRefresh.Enabled = True
    'don't wait for the first interval to elapse
    tmrRefresh_Timer
    '-----------------------------------------------------------------
End Sub

Private Sub mnuShutDown_Click(Index As Integer)
    Select Case Index
        Case 0: KillWindows SHUTDOWN
        Case 1: KillWindows REBOOT
        Case 2: KillWindows LOGOFF
        Case 4: KillWindows POWEROFF
    End Select
End Sub

Private Sub tabProcess_Click()
    Dim pic As PictureBox
    For Each pic In picProc
        pic.Visible = (pic.Index = tabProcess.SelectedItem.Index - 1)
    Next
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = True
End Sub

Private Sub Timer2_Timer()
'==========================================================================================================='
' Queries the system for process information, and edits the result back to the form for display             '
'==========================================================================================================='
' We only want to do this if the popumenu is not visible, otherwise me might refresh at the wrong moment    '
'-----------------------------------------------------------------------------------------------------------'
    If Not mnuT.Visible Then
        ProcessList lvwProcess, ilsProcess
        picProcess.Cls
        lblValue(2) = ""
        lblValue(3) = ""
        lblValue(0) = ""
        lblValue(1) = ""
        lblValue(4) = ""
        lblValue(5) = ""
    End If
End Sub

Private Sub tmrRefresh_Timer()
    Dim ret As Long
    'query the CPU usage
    ret = QueryObject.Query
    If ret = -1 Then
        tmrRefresh.Enabled = False
        lblCpuUsage.Caption = ":("
        MsgBox "Error while retrieving CPU usage"
    Else
        DrawUsage ret, picUsage, picGraph
        lblCpuUsage.Caption = CStr(ret) + "%"
        sbProcess.Panels(2).Text = "CPU Usage : " & lblCpuUsage.Caption '= CStr(Ret) + "%"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Engine32.CloseScanHandle
    Set Engine32 = Nothing
    Unload Me
End Sub

Private Sub lvwProcess_Click()
    On Error Resume Next
    Dim ver As VERHEADER
    Dim sFile As String, sFileName As String
    picProcess.Cls
    sFileName = lvwProcess.SelectedItem.SubItems(2)
    If sFile <> sFileName Then
        file_getName (sFileName)
        GetVerHeader sFileName, ver
        lblValue(2) = "File"
        lblValue(3) = "Folder"
        lblValue(0) = ver.FileDescription
        lblValue(1) = ver.CompanyName
        lblValue(4) = ": " & file_getName(sFileName)
        lblValue(5) = ": " & file_getPath(sFileName)
        RetrieveIcon sFileName, picProcess, ricnLarge
    End If
End Sub

Private Sub lvwProcess_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvwProcess.Sorted = True
    lvwProcess.SortKey = ColumnHeader.Index - 1
    If lvwProcess.SortOrder = lvwDescending Then
       lvwProcess.SortOrder = lvwAscending
    Else
       lvwProcess.SortOrder = lvwDescending
    End If
End Sub

Private Sub lvwProcess_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        If lvwProcess.ListItems.count > 0 Then
            mnuFindFile.Caption = "Find File Location..."
            mnuFindFile.Tag = lvwProcess.SelectedItem.SubItems(2)
            mnuSetPriority(1).Checked = False
            mnuSetPriority(2).Checked = False
            mnuSetPriority(3).Checked = False
            mnuSetPriority(4).Checked = False
            lvwProcess_Click
            
            Dim priHwnd  As Long
            priHwnd = GetPriority(CLng(lvwProcess.SelectedItem.SubItems(3)))
            Select Case priHwnd
                   Case REALTIME_PRIORITY_CLASS
                        mnuSetPriority(1).Checked = True
                   Case HIGH_PRIORITY_CLASS
                        mnuSetPriority(2).Checked = True
                   Case NORMAL_PRIORITY_CLASS
                        mnuSetPriority(3).Checked = True
                   Case IDLE_PRIORITY_CLASS
                        mnuSetPriority(4).Checked = True
            End Select
            PopupMenu mnuT
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.show
End Sub

Private Sub mnuEndProcess_Click()
    Dim i As Integer
    Dim Pesan As String, strFile As String
    Dim lExitCode As Long
    
    Pesan = "WARNING: Terminating a process can cause undesired" & vbCrLf & _
            "results including loss of data and system instability. The" & vbCrLf & _
            "process will not be given the chance to save its state or" & vbCrLf & _
            "data before it is terminated. Are you sure you want to" & vbCrLf & _
            "terminate the process?"
    If MsgBox(Pesan, vbYesNo + 48, "Process Manager Warning" & Chr(0)) = vbYes Then
        lExitCode = TerminateProcessID(lvwProcess, 3)
        If lExitCode = 0 Then MsgBox "Cannot terminate this process.", vbExclamation, "Unable To Terminate Process"
        Call mnuRefresh_Click
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFileInfo_Click()
    frmDetail.show 'vbModal
    Me.Hide
End Sub

Private Sub mnuFindFile_Click()
    OpenFolderProcess lvwProcess, 2
End Sub

Private Sub mnuNewTask_Click()
    Dim sTitle As String, sPrompt As String
    sTitle = "New Process"
    sPrompt = "Type the name of a program, folder, document, or Internet resource."
                
    If IsWinNT Then
        SHRunDialog Me.hWnd, 0, 0, StrConv(sTitle, vbUnicode), StrConv(sPrompt, vbUnicode), 0
    Else
        SHRunDialog Me.hWnd, 0, 0, sTitle, sPrompt, 0
    End If
End Sub

Private Sub mnuProcess_Click(Index As Integer)
    Select Case Index
        Case 0: SetSuspendResumeThread lvwProcess, 3, True
        Case 1: SetSuspendResumeThread lvwProcess, 3, False
    End Select
End Sub

Private Sub mnuProperties_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 1 To lvwProcess.ListItems.count
      If lvwProcess.ListItems(i).Selected Then
         ShowProps lvwProcess.ListItems(i).SubItems(2), Me.hWnd
      End If
    Next i
End Sub

Private Sub mnuRefresh_Click()
    ProcessList lvwProcess, ilsProcess
    sbProcess.Panels(4).Text = ""
    sbProcess.Panels(3).Text = ""
End Sub

Private Sub mnuRun_Click()
    Call mnuNewTask_Click
End Sub

Private Sub mnuScan_Click()
    KillVirusProcessList True
End Sub

Private Sub mnuSetPriority_Click(Index As Integer)
    Dim lBase As Long
    Dim priHwnd  As Long
    Select Case Index
        Case 1
            lBase = REALTIME_PRIORITY_CLASS 'SetBasePriority(lvwProcess, 3, REALTIME_PRIORITY_CLASS)
        Case 2
            lBase = HIGH_PRIORITY_CLASS 'SetBasePriority(lvwProcess, 3, HIGH_PRIORITY_CLASS)
        Case 3
            lBase = NORMAL_PRIORITY_CLASS 'SetBasePriority(lvwProcess, 3, NORMAL_PRIORITY_CLASS)
        Case 4
            lBase = IDLE_PRIORITY_CLASS 'SetBasePriority(lvwProcess, 3, IDLE_PRIORITY_CLASS)
    End Select

    Dim i As Integer
    If lBase <> 0 Then
        For i = 1 To lvwProcess.ListItems.count
            If lvwProcess.ListItems(i).Selected Then
                priHwnd = OpenProcess(PROCESS_SET_INFORMATION, False, CLng(lvwProcess.SelectedItem.SubItems(3)))
                SetPriorityClass priHwnd, lBase
                CloseHandle priHwnd
            End If
        Next i
        Call mnuRefresh_Click
    End If
End Sub

Sub RetrieveIcon(fName As String, DC As PictureBox, icnSize As IconRetrieve)
    Dim hImgLarge As Long  'the handle to the system image list
        
    If icnSize = ricnLarge Then
        hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    End If
End Sub

Private Sub tmrMem_Timer()
'    UpdateValues sbProcess
'    MemoryInfo lblInfo(0), lblInfo(1), lblInfo(2), lblInfo(3), lblInfo(4), lblInfo(5), lblInfo(6), lblInfo(7), lblInfo(8), lblInfo(9), lblInfo(3)
    If Not mnuT.Visible Then
        MonitoringPerformance lblinfo(3), lblinfo(0), lblinfo(1), lblinfo(2), lblinfo(9), lblinfo(7), lblinfo(8), lblinfo(4), lblinfo(5), lblinfo(6), lblinfo(10), lblinfo(11)
        sbProcess.Panels(1).Text = "Processes : " & lblinfo(11)
    End If
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
        sbProcess.Panels(3).Text = "Scanning memory..."
        sbProcess.Panels(4).Text = ""
    Loop
    CloseHandle hSnapShot
        
    If fileIsVirus.count > 0 Then
       Dim i As Integer
       For i = 1 To fileIsVirus.count
           If Engine32.CekOneFile(CStr(fileIsVirus(i)(0)), CLng(fileIsVirus(i)(1))) Then
                LogScan "Scan Memory Found " & fileIsVirus(i)(0)
           End If
          sbProcess.Panels(4).Text = "Infected : " & FileViruses 'virus found
       Next i
    End If
        
    FileOnScan = 0
                       
    If ViriOnCollect.count > 0 Then
        '
    End If
    sbProcess.Panels(4).Text = "Infected : " & FileViruses
    sbProcess.Panels(3).Text = "Completed."
    ProcessList lvwProcess, ilsProcess
End Sub

Function OptimizeRAM()
    ReDim Buf(100)
    Dim i
    
    For i = 0 To 100
        Prog.Value = i
        Buf(i) = SPACE$(OptS)
        ry.Caption = "Optimizing..."
    Next i
    
    For i = 0 To 100
        Buf(i) = vbNull
    Next i
    
    Prog.Value = 0
    ry = "Completed!"
    Timer1.Enabled = True
    cmdOpt.Enabled = True
End Function
