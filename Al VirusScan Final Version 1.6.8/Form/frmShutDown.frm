VERSION 5.00
Begin VB.Form frmShutDown 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3150
      TabIndex        =   1
      Top             =   2100
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Restart"
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
      Left            =   3465
      TabIndex        =   4
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Turn Off"
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
      Left            =   2010
      TabIndex        =   3
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Log Off"
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
      Left            =   555
      TabIndex        =   2
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image imgShutdown 
      Height          =   480
      Index           =   1
      Left            =   2130
      MouseIcon       =   "frmShutDown.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmShutDown.frx":1D82
      Top             =   975
      Width           =   480
   End
   Begin VB.Image imgShutdown 
      Height          =   480
      Index           =   0
      Left            =   675
      MouseIcon       =   "frmShutDown.frx":264C
      MousePointer    =   99  'Custom
      Picture         =   "frmShutDown.frx":43CE
      Top             =   975
      Width           =   480
   End
   Begin VB.Image imgShutdown 
      Height          =   480
      Index           =   2
      Left            =   3585
      MouseIcon       =   "frmShutDown.frx":4C98
      MousePointer    =   99  'Custom
      Picture         =   "frmShutDown.frx":6A1A
      ToolTipText     =   "Restart"
      Top             =   975
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   1215
      Index           =   1
      Left            =   225
      Shape           =   4  'Rounded Rectangle
      Top             =   750
      Width           =   4365
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   1215
      Index           =   0
      Left            =   225
      Shape           =   4  'Rounded Rectangle
      Top             =   750
      Width           =   4365
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4110
      Picture         =   "frmShutDown.frx":72E4
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Turn Off Computer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   3315
   End
End
Attribute VB_Name = "frmShutDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    AlwaysOnTop Me.Hwnd, False
    Unload Me
End Sub

Private Sub Form_Load()
    lonRect = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 20, 20)
    SetWindowRgn Me.Hwnd, lonRect, True
    AlwaysOnTop Me.Hwnd, True
'    Sleep 500
    imgShutdown(0).ToolTipText = "Close your programs and ends your windows session"
    imgShutdown(1).ToolTipText = "Shuts down windows so that you can safely turn off the computer"
    imgShutdown(2).ToolTipText = "Shuts down windows and then start windows again"
End Sub

Private Sub imgShutdown_Click(Index As Integer)
    Select Case Index
        Case 0: KillWindows LOGOFF
        Case 1: KillWindows SHUTDOWN
        Case 2: KillWindows REBOOT
    End Select
End Sub
