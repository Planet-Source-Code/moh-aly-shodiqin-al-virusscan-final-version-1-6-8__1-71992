VERSION 5.00
Begin VB.Form frmCheckUpdate 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFadeOut 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   900
      Top             =   2625
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   450
      Top             =   2625
   End
   Begin VB.Timer tmrFadeIn 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   2625
   End
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
      Left            =   3240
      TabIndex        =   1
      Top             =   1350
      Width           =   1515
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Now..."
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
      Left            =   1650
      TabIndex        =   0
      Top             =   1350
      Width           =   1515
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
      Height          =   15
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   4890
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   675
      TabIndex        =   4
      Top             =   1875
      Width           =   4215
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
      Index           =   1
      Left            =   675
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   675
      X2              =   4875
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   675
      X2              =   4875
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmCheckUpdate.frx":0000
      Top             =   75
      Width           =   480
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
      Index           =   0
      Left            =   675
      TabIndex        =   2
      Top             =   75
      Width           =   4215
   End
   Begin VB.Label lblValue 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   675
      TabIndex        =   6
      Top             =   600
      Width           =   2040
   End
End
Attribute VB_Name = "frmCheckUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 1 April 2009 1:57 AM
' Module    : Check Update.
' --------------------------------------
Private Const SPI_GETWORKAREA As Long = 48&
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Type OSVERSIONINFO
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private m_iChangeSpeed    As Long         '/* The window's display speed
Private m_iCounter        As Long         '/* Display time in milliseconds
Private m_iScrnBottom     As Long         '/* Height of the screen - taskbar (if it is on the bottom)
Private m_bOnTop          As Boolean      '/* Form Z-Order Flag
Private m_iWindowCount    As Long         '/* Screen stop position multiplier (displaying more then 1 at a time)
Private m_bManualClose    As Boolean      '/* Manual close Flag
Private m_bCodeClose      As Boolean      '/* Prevent user close option
Private m_bFade           As Boolean      '/* Fade or move Flag
Private m_iOSver          As Byte         '/* OS 1=Win98/ME; 2=Win2000/XP
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim tg As Byte
Dim co As Byte

Private Sub cmdCancel_Click()
'    Unload Me
    tmrFadeOut.Enabled = True
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    cmdCancel_Click
    frmUpdate.show
End Sub

Private Sub Form_Load()
    Me.Caption = "Update Reminder"
    Dim rc         As RECT
    Dim scrnRight  As Long
    Dim OSV        As OSVERSIONINFO
    
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then m_iOSver = 1 '/* Win 98/ME
        If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then m_iOSver = 2  '/* Win 2000/XP
    End If

'    lonRect = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 20, 20)
'    SetWindowRgn Me.hWnd, lonRect, True
    '/* Get Screen and TaskBar size
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, rc, 0&)
    '/* Screen Height - Taskbar Height (if is is located at the bottom of the screen)
    m_iScrnBottom = rc.bottom * Screen.TwipsPerPixelY
    '/* Is the taskbar is located on the right side of the screen? (scrnRight < Screen.width)
    scrnRight = (rc.Right * Screen.TwipsPerPixelX)
    '/* Locate Form to bottom right and set default size
    Me.Move scrnRight - Me.Width, m_iScrnBottom, lblValue(3).Left + lblValue(3).Width + 100, cmdCancel.Top + 700
     
    tmrFadeIn.Enabled = True
    tg = 15
    lblValue(0) = "To get the latest update of al VirusScan Virus Definitions. " & vbCrLf & _
                "Update al VirusScan regularly are recommended."
    lblValue(4) = "Update Reminder!"
    lblValue(4).ForeColor = vbBlue
    lblValue(1) = "Your virus definitions is quite old. " & "(" & DateDiff("d", vVirusDefinitions, Date) & " days ago.)" & vbCrLf & _
                    "Please visit to http://fi5ly.blogspot.com"
'    lblValue(2) = "Please visit to http://fi5ly.blogspot.com" 'DateDiff("d", vVirusDefinitions, Date) & " days ago."
End Sub

Private Sub tmrFadeIn_Timer()
    If Me.Top < m_iScrnBottom - Me.Height Then
        tmrUpdate.Enabled = True
        tmrFadeIn.Enabled = False
    Else
        Me.Move Me.Left, Me.Top - 25, lblValue(3).Left + lblValue(3).Width + 100, cmdCancel.Top + 470
    End If
End Sub

Private Sub tmrFadeOut_Timer()
    If Me.Top > m_iScrnBottom Then
        Unload Me
    Else
        Me.Move Me.Left, Me.Top + 25, lblValue(3).Left + lblValue(3).Width + 100, cmdCancel.Top + 470
    End If
End Sub

Private Sub tmrUpdate_Timer()
    If tg - co > 0 Then
        co = co + 1
    Else
        tmrFadeOut.Enabled = True
    End If
End Sub
