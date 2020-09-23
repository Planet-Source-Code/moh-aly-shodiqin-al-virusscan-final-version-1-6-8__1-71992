VERSION 5.00
Begin VB.Form frmAuthor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   526
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   150
      Picture         =   "frmAuthor.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   6
      Top             =   150
      Width           =   720
   End
   Begin VB.CommandButton v 
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
      Left            =   6150
      TabIndex        =   5
      ToolTipText     =   "Click ok to exit about author al VirusScan..."
      Top             =   3975
      Width           =   1515
   End
   Begin VB.Label LBLB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   2820
      Width           =   285
   End
   Begin VB.Label LBLC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   540
      Left            =   225
      TabIndex        =   3
      Top             =   2625
      Width           =   285
   End
   Begin VB.Label LBLA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   105
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Caption         =   "al VirusScan"
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
      Left            =   6630
      TabIndex        =   1
      Top             =   150
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label LBLCR 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   0
      Top             =   4125
      Visible         =   0   'False
      Width           =   6495
   End
End
Attribute VB_Name = "frmAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Dim x As Integer

Private Sub Form_Load()
    On Local Error Resume Next
    v.Visible = False
    AlwaysOnTop Me.Hwnd, True
    LBLB.Left = Me.ScaleWidth / 2 - LBLB.Width / 2 - 1
    LBLB.Top = 100
    x = 1
    lonRect = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 20, 20)
    SetWindowRgn Me.Hwnd, lonRect, True
    FormFadeIn Me, 0, 240, 4
    SoundBuffer = LoadResData("W3", "WAV")
    sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_LOOP
    TA
End Sub

Private Sub FormFadeIn(ByRef nForm As Form, Optional ByVal nFadeStart As Byte = 0, Optional ByVal nFadeEnd As Byte = 255, Optional ByVal nFadeInSpeed As Byte = 5)
    Dim C
    Dim ne As Integer, EN(32767) As Boolean
    For Each C In nForm.Controls
     ne = ne + 1
     EN(ne) = C.Enabled
     C.Enabled = False
    Next
    If nFadeEnd = 0 Then
        nFadeEnd = 255
    End If
    If nFadeInSpeed = 0 Then
        nFadeInSpeed = 5
    End If
    If nFadeStart >= nFadeEnd Then
        nFadeStart = 0
    ElseIf nFadeEnd <= nFadeStart Then
        nFadeEnd = 255
    End If
    
       TransparentsForm nForm.Hwnd, 0
        nForm.show
        Dim I As Long
        For I = nFadeStart To nFadeEnd Step nFadeInSpeed
            TransparentsForm nForm.Hwnd, CByte(I)
            DoEvents
            Call Sleep(5)
        Next
        TransparentsForm nForm.Hwnd, nFadeEnd
        I = 0
    For Each C In nForm.Controls
     I = I + 1
     C.Enabled = EN(I)
    Next
End Sub

Private Function TransparentsForm(FormhWnd As Long, Alpha As Byte) As Boolean
    SetWindowLong FormhWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes FormhWnd, 0, Alpha, LWA_ALPHA
    LastAlpha = Alpha
End Function

Private Sub FormFadeOut(ByRef nForm As Form)
    On Local Error Resume Next
    Dim C
    Dim s As Integer
    For Each C In nForm.Controls
     C.Enabled = False
    Next
    
    Dim I As Long
    For I = 240 To 0 Step -5
        TransparentsForm nForm.Hwnd, CByte(I)
        DoEvents
        Call Sleep(5)
    Next
End Sub

Private Sub TA()
    On Local Error Resume Next
    Dim x As Long
    
    Me.FontSize = 25
    Me.FontBold = True
    Me.FontName = "Arial"
    Me.ForeColor = RGB(72, 123, 117)
    '--------------------------------------------
    For I = 0 To 12
        Me.CurrentX = 64
        Me.CurrentY = 15
        Me.Print Mid("al VirusScan", 1, CByte(I))
        DoEvents
        Call Sleep(50)
    Next
    '--------------------------------------------
    Me.FontBold = False
    Me.FontSize = 8
        Me.CurrentX = 255
        Me.CurrentY = 14
    Me.Print "®"
    '--------------------------------------------
    Me.ForeColor = 0
    Me.FontSize = 8
    Me.FontName = "Courier New"
    Me.CurrentX = 125
    Me.CurrentY = 45
    Me.Print "Final Version " & App.Major & "." & App.Minor & "." & App.Revision
    '--------------------------------------------
    For I = 0 To 5
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
'    For i = 0 To 33
'        Me.CurrentX = 270
'        Me.CurrentY = 45
'        Me.Print Mid("Virus Definitions : 25 Maret 2009", 1, CByte(i))
'        DoEvents
'        Call Sleep(40)
'    Next
'    Call Sleep(200)
    '--------------------------------------------
    Me.FontBold = False
    Me.ForeColor = 0
    Me.FontSize = 10
    Me.FontName = "Courier New"
    '--------------------------------------------
    For I = 0 To 12
        Me.CurrentX = 10
        Me.CurrentY = 65
        Me.Print Mid("Developed By", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlue
    For I = 0 To 19
        Me.CurrentX = 100
        Me.CurrentY = 65
        Me.Print Mid(" : Moh Aly Shodiqin", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
    '--------------------------------------------
    Me.ForeColor = vbBlack
    For I = 0 To 7
        Me.CurrentX = 10
        Me.CurrentY = 80
        Me.Print Mid("Address", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlue
    For I = 0 To 51
        Me.CurrentX = 100
        Me.CurrentY = 80
        Me.Print Mid(" : Ds Campurejo RT 12/03 Panceng Gresik - Indonesia", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
    '--------------------------------------------
    Me.ForeColor = vbBlack
    For I = 0 To 9
        Me.CurrentX = 10
        Me.CurrentY = 95
        Me.Print Mid("Copyright", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlue
    For I = 0 To 50
        Me.CurrentX = 100
        Me.CurrentY = 95
        Me.Print Mid(" : 2008-2009 Moh Aly Shodiqin. All rights reserved", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
    '--------------------------------------------
    Me.ForeColor = vbBlack
    For I = 0 To 7
        Me.CurrentX = 10
        Me.CurrentY = 110
        Me.Print Mid("Company", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlue
    For I = 0 To 10
        Me.CurrentX = 100
        Me.CurrentY = 110
        Me.Print Mid(" : DQ Soft", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(10)
    '--------------------------------------------
    Me.ForeColor = vbBlack
    For I = 0 To 5
        Me.CurrentX = 10
        Me.CurrentY = 125
        Me.Print Mid("Email", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlue
    For I = 0 To 29
        Me.CurrentX = 100
        Me.CurrentY = 125
        Me.Print Mid(" : felix_progressif@yahoo.com", 1, CByte(I)) '49 / fi5ly@yahoo.co.id
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
        Me.ForeColor = vbBlack
    For I = 0 To 4
        Me.CurrentX = 10
        Me.CurrentY = 140
        Me.Print Mid("Blog", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlue
    For I = 0 To 28
        Me.CurrentX = 100
        Me.CurrentY = 140
        Me.Print Mid(" : http://fi5ly.blogspot.com", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
    '--------------------------------------------
    Me.ForeColor = vbBlack
    For I = 0 To 50
        Me.CurrentX = 110
        Me.CurrentY = 151
        Me.Print Mid("---------------------------------------------------", 1, CByte(I))
        DoEvents
        Call Sleep(45)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlack
    For I = 0 To 45
        Me.CurrentX = 10
        Me.CurrentY = 164
        Me.Print Mid("Its use is not allowed weared for the purpose", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlack
    For I = 0 To 48
        Me.CurrentX = 10
        Me.CurrentY = 177
        Me.Print Mid("of profit or commercial, without permit from me.", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    Me.ForeColor = vbBlack
    For I = 0 To 65
        Me.CurrentX = 10
        Me.CurrentY = 190 '153 '141
        Me.Print Mid("See readme in About al VirusScan for detail informations...", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(200)
    '--------------------------------------------
    
    Call Sleep(50)
    Me.FontSize = 10
    Me.ForeColor = vbBlack
    Me.FontName = "Courier New"
    For I = 0 To 30
        Me.CurrentX = 10
        Me.CurrentY = 220
        Me.Print Mid("If you find any problems/bug.", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
    
    For I = 0 To 50
        Me.CurrentX = 10
        Me.CurrentY = 233
        Me.Print Mid("Any questions for this application, email to me...", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
    
    For I = 0 To 39
        Me.CurrentX = 10
        Me.CurrentY = 247
        Me.Print Mid("Free Software But Without Any Warranty.", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
    
    For I = 0 To 18
        Me.CurrentX = 10
        Me.CurrentY = 260
        Me.Print Mid("Made In Indonesia.", 1, CByte(I))
        DoEvents
        Call Sleep(40)
    Next
    Call Sleep(100)
    
    For x = 0 To 10
        Me.CurrentX = 430
        Me.CurrentY = 275
        Me.Print Mid(">>>>>>>>>>", 1, CByte(x))
        DoEvents
        Call Sleep(80)
    Next
    Call Sleep(80)
    
'    Me.FontSize = 9
'    Me.ForeColor = vbBlack
'    Me.FontName = "Courier New"
    Me.CurrentX = 10
    Me.CurrentY = 277
    Me.Print "© 2008-2009 Moh Aly Shodiqin. All right reserved "
    v.Visible = True
End Sub

Private Sub V_Click()
    FormFadeOut Me
    AlwaysOnTop Me.Hwnd, False
    SoundBuffer = LoadResData("W3", "WAV")
    sndPlaySound ByVal 0&, SND_NODEFAULT
    Unload Me
End Sub
