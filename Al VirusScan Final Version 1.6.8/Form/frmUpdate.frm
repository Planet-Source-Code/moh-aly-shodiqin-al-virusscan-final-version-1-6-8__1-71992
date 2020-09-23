VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   ControlBox      =   0   'False
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAuto 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3675
      Top             =   750
   End
   Begin ComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   1290
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
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
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3525
      Top             =   1650
   End
   Begin ComctlLib.ProgressBar pbUpdate 
      Height          =   165
      Left            =   75
      TabIndex        =   2
      Top             =   1650
      Visible         =   0   'False
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   291
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   4125
      TabIndex        =   1
      Top             =   750
      Width           =   1365
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   300
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblAuto 
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
      Left            =   75
      TabIndex        =   4
      Top             =   825
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5025
      Picture         =   "frmUpdate.frx":08CA
      Top             =   60
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   75
      X2              =   5550
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   75
      X2              =   5550
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   225
      Width           =   5415
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 5 Februari 2009
' 6:06 AM
'=======================================
' Module AutoUpdate
'=======================================
' 31 Maret 2009 10:47 AM
' fixed AutoUpdate

Option Explicit

Dim tg As Byte
Dim co As Byte

Sub DownStatus(ByVal strStatus As String)
    lblStatus.Caption = strStatus
End Sub

Private Sub cmdClose_Click()
    If cmdClose.Caption = "Cancel" Then
        tmrAuto.Enabled = True
        cmdClose.Caption = "Close"
        DownStatus "Closing update session."
        sbStatus.Panels(1).Text = "Done."
        Screen.MousePointer = 0
        tmrStatus.Enabled = False
    ElseIf cmdClose.Caption = "Close" Then
        Screen.MousePointer = 0
        tmrStatus.Enabled = False
        tmrAuto.Enabled = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "al VirusScan - Virus Definitions Update"
    DownStatus "Checking internet connection..."
'    On Error GoTo ErrDown
    sbStatus.Panels(1).Text = "Please wait for update to finish..."
    AlwaysOnTop Me.Hwnd, True
    
    Dim hMsg As Long
    hMsg = MsgBox("The Auto Update system will look for new available updates for the virus definitions file." & vbCrLf & _
        "This requires a connection to the internet." & vbCrLf & vbCrLf & _
        "Update Procedure : " & vbCrLf & _
        "1. Download Scan (rar)" & vbCrLf & _
        "2. Terminate process al VirusScan.exe" & vbCrLf & _
        "3. Extract Scan (rar) to al VirusScan Directory" & vbCrLf & _
        "4. Replace scan.vdf old version with scan.vdf last version" & vbCrLf & vbCrLf & _
        "Do you want to download update now?", vbQuestion + vbYesNo, Me.Caption)
    Select Case hMsg
        Case vbYes
            tmrStatus.Enabled = True
        Case vbNo
            Unload Me
    End Select
    tg = 30
    
'    DownStatus "Starting Download file..."
'    DownloadFile "http://localhost/DQ%20Soft/Update_al%20VirusScan/scan.vdf", App.path & "\temp.$$$"
'    If FileLen(App.path & "\temp.$$$") <> 0 Then
'        DownStatus "Download Completed."
'        Kill App.path & "\Data\scan.vdf"
'        FileCopy App.path & "\temp.$$$", App.path & "\Data\scan.vdf"
'        Kill App.path & "\temp.$$$"
'    End If
'    Exit Sub
'
'ErrDown:
'    DownStatus "Error occurred while downloading file..."
End Sub

Private Sub tmrAuto_Timer()
    If tg - co > 0 Then
        co = co + 1
        lblAuto.Caption = "Auto close in " & tg - co & " seconds."
    Else
        Unload Me
    End If
End Sub

Private Sub tmrStatus_Timer()
    On Error GoTo ErrDown
    With pbUpdate
        If .Value < 100 Then
            Screen.MousePointer = 13
            cmdClose.Caption = "Cancel"
            DoEvents
            .Value = .Value + 1
            If .Value = 30 Then
                DownStatus "Checking internet connection...": DoEvents
                Sleep 3000
            End If
            If .Value = 60 Then
                DownStatus "Starting Download file...": DoEvents
                DoEvents
                Sleep 3000
            End If
        Else
            .Value = 100
            DownStatus "Starting Download file..."
            DownloadFile "http://www.4shared.com/file/97970974/59c04359/scan.html", App.path & "\temp.$$$"
            If FileLen(App.path & "\temp.$$$") <> 0 Then
                DownStatus "Download Completed."
                sbStatus.Panels(1).Text = "Done."
'                Kill App.path & "\Data\scan.vdf"
                Dim H As String
                H = App.path & "\scan.rar"
                If Len(Dir$(H)) = 0 Then
                    FileCopy App.path & "\temp.$$$", H
                    Kill App.path & "\temp.$$$"
                Else
                    Kill H
                    FileCopy App.path & "\temp.$$$", H
                    Kill App.path & "\temp.$$$"
                End If
            End If
            cmdClose.Caption = "Close"
            tmrAuto.Enabled = True
            Screen.MousePointer = 0
            tmrStatus.Enabled = False
            Exit Sub
        End If
    End With
    Exit Sub
    
ErrDown:
    cmdClose.Caption = "Close"
    tmrAuto.Enabled = False
    cmdClose.Enabled = True
    Screen.MousePointer = 0
    tmrStatus.Enabled = False
    DownStatus "Error occurred while downloading file..."
End Sub
