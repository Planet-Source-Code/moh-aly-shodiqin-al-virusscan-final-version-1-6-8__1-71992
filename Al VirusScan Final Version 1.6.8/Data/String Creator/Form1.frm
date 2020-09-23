VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl 
      Left            =   4575
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4800
      Width           =   6165
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read File"
      Height          =   390
      Left            =   4425
      TabIndex        =   2
      Top             =   675
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   390
      Left            =   4425
      TabIndex        =   1
      Top             =   225
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Height          =   4515
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   225
      Width           =   3840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filename As String

Private Sub Command1_Click()
    cdl.ShowOpen
    filename = cdl.filename
    Command2.Enabled = True
    Text2 = filename
End Sub

Private Sub ReadFile()
    Dim buff As String
    Dim buffLen As Long
    Dim nfile As Integer
    Dim temp As String, i As Long
'    On Error Resume Next
    nfile = FreeFile
    Open filename For Binary Access Read As #nfile
        buffLen = 255
        Do Until EOF(nfile)
            If buffLen > LOF(nfile) - Loc(nfile) Then
                buffLen = LOF(nfile) - Loc(nfile)
                If buffLen < 1 Or Err Then Exit Do
            End If
            buff = Space$(buffLen)
            Get #nfile, , buff
            For i = 1 To Len(buff)
                If InStr(Mid(buff, i, 1), Mid(buff, i, 1)) Then
                    temp = temp & Mid(buff, i, 1)
                    Text1.Text = Trim(temp)
                Else
'                    If Len(temp) > 3 Then
''                        Text1.Text = Trim(temp)
'                    End If
                    temp = ""
                End If
            Next i
            DoEvents
        Loop
    Close nfile
    If Err Then Err.Clear
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text1.SetFocus
    ReadFile
End Sub

Private Sub Form_Load()
    Command2.Enabled = False
End Sub
