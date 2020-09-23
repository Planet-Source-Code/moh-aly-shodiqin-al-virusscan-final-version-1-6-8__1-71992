VERSION 5.00
Begin VB.Form frmDrive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   Icon            =   "frmDrive.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDrive 
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
      Index           =   2
      Left            =   3300
      TabIndex        =   4
      Top             =   3000
      Width           =   1065
   End
   Begin VB.CommandButton cmdDrive 
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
      Height          =   390
      Index           =   1
      Left            =   2175
      TabIndex        =   3
      Top             =   3000
      Width           =   1065
   End
   Begin VB.ListBox lstDrive 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   150
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1125
      Width           =   4215
   End
   Begin alVirusScan.dcButton cmdDrives 
      Height          =   390
      Left            =   150
      TabIndex        =   5
      Top             =   600
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
      PicNormal       =   "frmDrive.frx":08CA
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Please select Drives or Directory To Scan From Viruses"
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
      Height          =   210
      Left            =   150
      TabIndex        =   2
      Top             =   75
      Width           =   4035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click OK To Continue, Cancel To Abort "
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
      Height          =   210
      Left            =   150
      TabIndex        =   1
      Top             =   300
      Width           =   2805
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "Select All..."
         Index           =   0
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "Unselect..."
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub LoadDrive()
    On Error Resume Next
    Dim LDs As Long, Cnt As Long, sDrives As String
    LDs = GetLogicalDrives
    For Cnt = 0 To 25
        If (LDs And 2 ^ Cnt) <> 0 Then
            Dim Serial As Long, VName As String, FSName As String, ndrvName As String
            VName = String$(255, Chr$(0))
            FSName = String$(255, Chr$(0))
            GetVolumeInformation Chr$(65 + Cnt) & ":\", VName, 255, Serial, 0, 0, FSName, 255
            VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
            FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
            ndrvName = ""
            If VName = "" Then
                Select Case GetTipeDrive(Chr$(65 + Cnt) & ":\")
                       Case 2: ndrvName = "3½ Floppy (" & Chr$(65 + Cnt) & ":)"
                       Case 5: ndrvName = "CDROM (" & Chr$(65 + Cnt) & ":)"
                       Case Else: ndrvName = "Unknown (" & Chr$(65 + Cnt) & ":)"
                End Select
                If ndrvName <> "" Then
                    lstDrive.AddItem Chr$(65 + Cnt) & ":\" & vbTab & ndrvName
                       
                End If
            Else
                ndrvName = VName & " (" & Chr$(65 + Cnt) & ":)"
                lstDrive.AddItem Chr$(65 + Cnt) & ":\" & vbTab & ndrvName
                'Chr$(65 + Cnt) & ":\", ndrvName)
            End If
        End If
    Next Cnt
End Sub

Private Sub cmdDrive_Click(Index As Integer)
    Select Case Index
        Case 1
            Set DrvOnCollect = Nothing
            Set DrvOnCollect = New Collection
            Dim I As Integer, myTab() As String
            For I = 0 To lstDrive.ListCount - 1
               If lstDrive.Selected(I) Then
                  myTab() = Split(lstDrive.List(I), vbTab)
                  DrvOnCollect.Add myTab(0), myTab(0)
               End If
            Next I
            OnSelectDlg = vbOK
            Unload Me
        Case 2
            OnSelectDlg = vbCancel
            Unload Me
    End Select
End Sub

Private Sub cmdDrives_Click()
    PopupMenu mnuFile, , cmdDrives.Left + 50, cmdDrives.Top + cmdDrives.Height
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = "Select Drive or Folder"
    LoadDrive
    cmdDrives.Caption = ""
    cmdDrives.ToolTipText = "Select local drives..."
    
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.Hwnd, True
    Else
        AlwaysOnTop Me.Hwnd, False
    End If
End Sub

Private Sub mnuSelect_Click(Index As Integer)
    On Error Resume Next
    Dim I As Integer
    Select Case Index
        Case 0
            For I = 0 To lstDrive.ListCount - 1
                lstDrive.Selected(I) = True
            Next I
        Case 1
            For I = 0 To lstDrive.ListCount - 1
                lstDrive.Selected(I) = False
            Next I
    End Select
End Sub
