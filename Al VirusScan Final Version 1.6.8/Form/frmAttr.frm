VERSION 5.00
Begin VB.Form frmAttr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "frmAttr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Archive"
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
      Left            =   3450
      TabIndex        =   15
      Top             =   975
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hidden"
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
      Left            =   4845
      TabIndex        =   14
      Top             =   975
      Width           =   1215
   End
   Begin VB.CommandButton cmdValue 
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
      Left            =   4695
      TabIndex        =   13
      Top             =   2325
      Width           =   1365
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "Stop"
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
      Index           =   2
      Left            =   3000
      TabIndex        =   12
      Top             =   2325
      Width           =   1365
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "Start"
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
      Index           =   1
      Left            =   1575
      TabIndex        =   11
      Top             =   2325
      Width           =   1365
   End
   Begin VB.CommandButton cmdValue 
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
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   2325
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Caption         =   "System"
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
      Left            =   4845
      TabIndex        =   9
      Top             =   1215
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Read Only"
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
      Left            =   3450
      TabIndex        =   8
      Top             =   1215
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   150
      X2              =   6075
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   6075
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SET ATTRIBUTE FOR FILE/FOLDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3435
      TabIndex        =   16
      Top             =   675
      Width           =   2625
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   2
      Left            =   1605
      TabIndex        =   7
      Top             =   1245
      Width           =   90
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
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
      Index           =   3
      Left            =   225
      TabIndex        =   6
      Top             =   1035
      Width           =   330
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folders"
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
      Index           =   5
      Left            =   225
      TabIndex        =   5
      Top             =   825
      Width           =   540
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   1
      Left            =   1605
      TabIndex        =   4
      Top             =   1035
      Width           =   90
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   0
      Left            =   1605
      TabIndex        =   3
      Top             =   825
      Width           =   90
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Failed"
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
      Index           =   8
      Left            =   225
      TabIndex        =   2
      Top             =   1245
      Width           =   420
   End
   Begin VB.Label lblDir 
      BackStyle       =   0  'Transparent
      Caption         =   "In Folder ?"
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
      Height          =   465
      Left            =   225
      TabIndex        =   1
      Top             =   1605
      Width           =   5835
   End
   Begin VB.Label lblfilename 
      BackStyle       =   0  'Transparent
      Caption         =   "||--"
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
      Height          =   450
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   5835
   End
End
Attribute VB_Name = "frmAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents SearchCls As cFileSearch
Attribute SearchCls.VB_VarHelpID = -1
Dim var_file As Long
Dim var_dir As Long
Dim var_filed As Long

Private Sub cmdValue_Click(Index As Integer)
    On Error Resume Next
    Dim H As String
    Select Case Index
        Case 0
            H = BrowseFolder(Me.hWnd, "Select directory to set attribute")
            If Trim(H) <> "" Then
                lblDir = H
                cmdValue(1).Enabled = True
            End If
        Case 1
            var_file = 0
            var_dir = 0
            var_filed = 0
            cmdValue(0).Enabled = False
            cmdValue(1).Enabled = False
            cmdValue(2).Enabled = True
            SearchCls.StopSearch = False
            Check1(0).Enabled = False
            Check1(1).Enabled = False
            Check1(2).Enabled = False
            Check1(3).Enabled = False
            cmdValue(3).Enabled = False
                
            SearchCls.DoCmdSearchFile lblDir, True
            cmdValue(2).Enabled = False
            cmdValue(1).Enabled = True
            cmdValue(0).Enabled = True
            SearchCls.StopSearch = True
            Check1(0).Enabled = True
            Check1(1).Enabled = True
            Check1(2).Enabled = True
            Check1(3).Enabled = True
            cmdValue(3).Enabled = True
        Case 2
            cmdValue(2).Enabled = False
            cmdValue(1).Enabled = True
            Check1(0).Enabled = True
            Check1(1).Enabled = True
            Check1(2).Enabled = True
            Check1(3).Enabled = True
            Check1(4).Enabled = True
            
            SearchCls.StopSearch = True
            cmdValue(0).Enabled = True
            cmdValue(3).Enabled = True
          Case 3
               Unload Me
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set SearchCls = New cFileSearch
    Me.Caption = "Set Attribute File Or Folder"
    
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.hWnd, True
    Else
        AlwaysOnTop Me.hWnd, False
    End If
End Sub

Private Sub SearchCls_onSearch(nFileName As String, nFileInfo As cFileInfo)
    On Error Resume Next
    lblfilename = "||-- " & nFileName
    
    Dim hRes As FILE_ATTRIBUTE
    
    If Check1(0).Value Then hRes = hRes + FILE_ATTRIBUTE_ARCHIVE
    If Check1(1).Value Then hRes = hRes + FILE_ATTRIBUTE_READONLY
    If Check1(2).Value Then hRes = hRes + FILE_ATTRIBUTE_HIDDEN
    If Check1(3).Value Then hRes = hRes + FILE_ATTRIBUTE_SYSTEM
       
    If SetFileAttributes(nFileName, hRes) = 0 Then var_filed = var_filed + 1
       
    If Trim(nFileInfo.Filename) <> "" Then
       var_file = var_file + 1
    Else
       var_dir = var_dir + 1
    End If
    
    lblvalue(0) = var_dir
    lblvalue(1) = var_file
    lblvalue(2) = var_filed
End Sub
