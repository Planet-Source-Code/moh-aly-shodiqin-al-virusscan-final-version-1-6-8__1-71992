VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "al VirusScan String Creator"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7980
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboType 
      BackColor       =   &H80000009&
      Height          =   315
      ItemData        =   "frmMain.frx":08CA
      Left            =   5850
      List            =   "frmMain.frx":08E9
      TabIndex        =   14
      Top             =   2025
      Width           =   1890
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "Add to DB"
      Height          =   390
      Left            =   4950
      TabIndex        =   11
      Top             =   525
      Width           =   1440
   End
   Begin VB.ListBox lstFile 
      Height          =   5715
      ItemData        =   "frmMain.frx":0944
      Left            =   75
      List            =   "frmMain.frx":0946
      TabIndex        =   3
      Top             =   75
      Width           =   4215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   390
      Left            =   4950
      TabIndex        =   10
      Top             =   75
      Width           =   1440
   End
   Begin VB.TextBox txtHex 
      Height          =   1155
      Left            =   4575
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4500
      Width           =   3165
   End
   Begin VB.TextBox txtAsciiToAscii 
      Height          =   1155
      Left            =   4575
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2850
      Width           =   3165
   End
   Begin VB.PictureBox pRead 
      Height          =   240
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   7905
      TabIndex        =   4
      Top             =   6225
      Width           =   7965
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5850
      Width           =   7815
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read File"
      Height          =   390
      Left            =   6450
      TabIndex        =   1
      Top             =   525
      Width           =   1440
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   390
      Left            =   6450
      TabIndex        =   0
      Tag             =   "Find a file for read "
      Top             =   75
      Width           =   1440
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3450
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "ASCII"
      Enabled         =   0   'False
      Height          =   1590
      Index           =   0
      Left            =   4425
      TabIndex        =   7
      Top             =   2550
      Width           =   3465
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hex"
      Enabled         =   0   'False
      Height          =   1590
      Index           =   1
      Left            =   4425
      TabIndex        =   8
      Top             =   4185
      Width           =   3465
   End
   Begin VB.TextBox txtAscii 
      Height          =   495
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1650
      Width           =   2640
   End
   Begin VB.TextBox txtVirusName 
      Height          =   330
      Left            =   4575
      TabIndex        =   12
      Top             =   1575
      Width           =   3165
   End
   Begin VB.Frame Frame1 
      Caption         =   "Virus Name"
      Enabled         =   0   'False
      Height          =   1215
      Index           =   2
      Left            =   4425
      TabIndex        =   13
      Top             =   1275
      Width           =   3465
      Begin VB.Label Label1 
         Caption         =   "Virus Type"
         Height          =   315
         Left            =   450
         TabIndex        =   15
         Top             =   750
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module     : al VirusScan String Creator...
'           : ver 1.1
'           : 10 April 2009 6:40 PM
'           : Moh Aly Shodiqin
'--------------------------------------------------------------
'License : Freeware n Open Source
'--------------------------------------------------------------
'Original Author : Moh Aly Shodiqin
'Release date    : 10 April 2009 6:40 PM
'Author Contact  : felix_progressif@yahoo.com /
'                : http://fi5ly.blogspot.com
'--------------------------------------------------------------
'Great thanks to    : - Allah S.W.T
'                   : - Nabi Muhammad S.A.W
'                   : - My Parent
'                   : - My Soul
'                   : - www.planetsourcecode.com
'-------------------------------------------------------
'Don't forget to vote me
'--------------------------------------------------------------
'Credits
'                   : - Matt Combatti
'                   : - Someone for A Professional Hex Editor
'--------------------------------------------------------------
Dim filename As String
Dim nfile As Integer

Private Sub cmdBrowse_Click()
    cdl.Filter = "Application |*.exe|All Files |*.*"
    cdl.ShowOpen
    filename = cdl.filename
    cmdRead.Enabled = True
    txtPath.Text = filename
End Sub

Private Sub ReadFile()
    Dim buff As String
    Dim buffLen As Long
'    Dim nfile As Integer
    Dim temp As String, i As Long
    On Error Resume Next
    nfile = FreeFile

    pRead.Visible = True
    pRead.Cls
    Open filename For Binary Access Read As #nfile
        buffLen = 256
        Do Until EOF(nfile)
            If buffLen > LOF(nfile) - Loc(nfile) Then
                buffLen = LOF(nfile) - Loc(nfile)
                If buffLen < 1 Or Err Then Exit Do
            End If
            pRead.Line (0, 0)-(pRead.Width / LOF(nfile) * Loc(nfile), _
                        pRead.Height), &HFF00&, BF
            buff = Space(buffLen)
            Get #nfile, , buff
            For i = 1 To Len(buff)
                If InStr(txtAscii, LCase(Mid(buff, i, 1))) Then
                    temp = temp & Mid(buff, i, 1)
                Else
                    If Len(temp) > 3 Then
                        lstFile.AddItem Trim(temp)
                    End If
                    temp = ""
                End If
            Next i
            DoEvents
        Loop
    Close nfile
    If Err Then Err.Clear
    pRead.Cls
'    pRead.Visible =True
    cmdBrowse.Enabled = True
    cmdRead.Caption = "Read File"
    cmdRead.Enabled = False
    lstFile.Enabled = True
End Sub

Private Sub cmdDB_Click()
    On Error GoTo ErrH
    Dim virType As String
    If txtAsciiToAscii.Text = "" Then
        MsgBox "Make sure ascii value is not empty", vbExclamation, "al VirusScan"
        txtAsciiToAscii.SetFocus
        Exit Sub
    End If
    If txtVirusName.Text = "" Then
        MsgBox "Plese add virus name", vbExclamation, "al VirusScan"
        txtVirusName.SetFocus
        Exit Sub
    End If
    
    Select Case cboType.ListIndex
        Case 0: virType = "(Trojan Horse)"
        Case 1: virType = "(Backdoor"
        Case 2: virType = "(Script)"
        Case 3: virType = "(Hacktool)"
        Case 4: virType = "(Worm)"
        Case 5: virType = "(Macro)"
        Case 6: virType = "(Resident Memory)"
        Case 7: virType = "(Virus)"
        Case 8: virType = "(Generic)"
    End Select
    
    Open App.Path & "\stringscan.vdf" For Append As #1
        Print #1, txtHex.Text & ":S:" & StrConv(txtVirusName.Text, vbProperCase) & " " & virType
    Close #1
    MsgBox "String Signature added successfully", vbInformation, "al VirusScan"
    txtVirusName.Text = ""
    txtHex.Text = ""
    txtAsciiToAscii.Text = ""
ErrH:
End Sub

Private Sub cmdRead_Click()
    Select Case cmdRead.Caption
        Case "Read File"
            cmdRead.Caption = "Stop"
            cmdBrowse.Enabled = False
'            cmdRead.Enabled = False
            lstFile.Clear
            lstFile.Refresh
            lstFile.Enabled = False
            ReadFile
        Case "Stop"
            cmdRead.Caption = "Read File"
            Close #nfile
    End Select
End Sub

Private Sub cmdAbout_Click()
    MsgBox "Below is the Eicar Test String which is a Universally Accepted Test String to test antivirus software." & vbCrLf & _
            "Upon scanning a file containing this string whether it be in ascii or binary format," & vbCrLf & _
            "the antivirus should return the scanned file as being infected." & vbCrLf & vbCrLf & _
            "** Note that it is used for testing purposes only and is not actually virus code." & vbCrLf & vbCrLf & _
            "Moh Aly Shodiqin" & vbCrLf & _
            "felix_progressif@yahoo.com" & vbCrLf & _
            "http://fi5ly.blogspot.com", vbInformation, "al VirusScan - EICAR STANDARD ANTIVIRUS TEST FILE"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 32 To 128
        txtAscii.Text = txtAscii.Text & Chr(i) '& "    " & vbCrLf
    Next i
    For i = 160 To 255
        txtAscii.Text = txtAscii.Text & Chr(i) '& "    " & vbCrLf
    Next i
    cmdRead.Enabled = False
    txtAscii.Visible = False
'    pRead.Visible = False
End Sub

Private Sub lstFile_Click()
    Dim index As Long
    index = lstFile.ListIndex + 1
    If index Then
        txtAsciiToAscii.Text = lstFile.Text
        txtHex.Text = Ascii2Hex(txtAsciiToAscii.Text)
    End If
End Sub

Function Ascii2Hex(valAscii As String) As String
    Dim i As Integer
    For i = 1 To Len(valAscii)
        Ascii2Hex = Ascii2Hex & Hex(Asc(Mid(valAscii, i, 1)))
    Next i
End Function

Private Sub txtAsciiToAscii_Change()
    txtHex.Text = Ascii2Hex(txtAsciiToAscii.Text)
End Sub

Private Sub txtVirusName_Change()
    cboType.Text = "Trojan Horse"
    If txtVirusName = "" Then cboType.Text = ""
End Sub
