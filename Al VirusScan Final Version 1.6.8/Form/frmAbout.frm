VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAbout 
      BorderStyle     =   0  'None
      Height          =   3315
      Index           =   0
      Left            =   225
      ScaleHeight     =   3315
      ScaleWidth      =   5640
      TabIndex        =   3
      Top             =   525
      Width           =   5640
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   2
         X1              =   825
         X2              =   5325
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   225
         X2              =   5550
         Y1              =   2400
         Y2              =   2400
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
         Index           =   31
         Left            =   225
         TabIndex        =   43
         Top             =   2100
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
         Height          =   240
         Index           =   32
         Left            =   1875
         TabIndex        =   40
         Top             =   2100
         Width           =   3690
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
         Height          =   210
         Index           =   29
         Left            =   225
         TabIndex        =   36
         Top             =   2790
         Width           =   4695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   225
         X2              =   5550
         Y1              =   2400
         Y2              =   2400
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
         Index           =   27
         Left            =   1875
         TabIndex        =   34
         Top             =   2550
         Width           =   3690
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
         Index           =   26
         Left            =   225
         TabIndex        =   33
         Top             =   2550
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         Caption         =   "al VirusScan "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   825
         TabIndex        =   14
         Top             =   210
         Width           =   4740
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   75
         Picture         =   "frmAbout.frx":08CA
         Top             =   75
         Width           =   720
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
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   885
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
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   12
         Top             =   1125
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
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   11
         Top             =   1365
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
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   10
         Top             =   1605
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   1875
         MouseIcon       =   "frmAbout.frx":1794
         MousePointer    =   99  'Custom
         TabIndex        =   9
         ToolTipText     =   "Click here to see author al VirusScan..."
         Top             =   885
         Width           =   3690
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
         Left            =   1875
         TabIndex        =   8
         Top             =   1125
         Width           =   3690
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
         Index           =   6
         Left            =   1875
         TabIndex        =   7
         Top             =   1365
         Width           =   3690
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
         Index           =   7
         Left            =   1875
         TabIndex        =   6
         Top             =   1605
         Width           =   3690
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
         Index           =   9
         Left            =   225
         TabIndex        =   5
         Top             =   1860
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
         Height          =   240
         Index           =   10
         Left            =   1875
         TabIndex        =   4
         Top             =   1860
         Width           =   3690
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   825
         X2              =   5325
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.PictureBox picAbout 
      BorderStyle     =   0  'None
      Height          =   3315
      Index           =   1
      Left            =   225
      ScaleHeight     =   3315
      ScaleWidth      =   5640
      TabIndex        =   15
      Top             =   525
      Width           =   5640
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   35
         Left            =   2850
         TabIndex        =   44
         Top             =   1035
         Width           =   2640
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   34
         Left            =   2850
         TabIndex        =   42
         Top             =   1665
         Width           =   2640
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   33
         Left            =   2850
         TabIndex        =   41
         Top             =   795
         Width           =   2640
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   2850
         TabIndex        =   38
         Top             =   555
         Width           =   2640
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   28
         Left            =   2850
         TabIndex        =   35
         Top             =   315
         Width           =   2640
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   150
         TabIndex        =   25
         Top             =   2205
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   2850
         TabIndex        =   30
         Top             =   75
         Width           =   2640
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   150
         TabIndex        =   29
         Top             =   2925
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   150
         TabIndex        =   28
         Top             =   2685
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   22
         Left            =   2850
         TabIndex        =   27
         Top             =   1275
         Width           =   2640
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   150
         TabIndex        =   26
         Top             =   2445
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   150
         TabIndex        =   24
         Top             =   1965
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   150
         TabIndex        =   23
         Top             =   1725
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   150
         TabIndex        =   22
         Top             =   1515
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   150
         TabIndex        =   21
         Top             =   1275
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   150
         TabIndex        =   20
         Top             =   1035
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   150
         TabIndex        =   19
         Top             =   795
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   150
         TabIndex        =   18
         Top             =   555
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   150
         TabIndex        =   17
         Top             =   315
         Width           =   2565
      End
      Begin VB.Label lblValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   150
         TabIndex        =   16
         Top             =   75
         Width           =   2565
      End
   End
   Begin VB.PictureBox picAbout 
      BorderStyle     =   0  'None
      Height          =   3315
      Index           =   2
      Left            =   225
      ScaleHeight     =   3315
      ScaleWidth      =   5640
      TabIndex        =   31
      Top             =   525
      Width           =   5640
      Begin VB.TextBox txtReadme 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   32
         Top             =   75
         Width           =   5490
      End
   End
   Begin VB.PictureBox picAbout 
      BorderStyle     =   0  'None
      Height          =   3315
      Index           =   3
      Left            =   225
      ScaleHeight     =   3315
      ScaleWidth      =   5640
      TabIndex        =   37
      Top             =   525
      Width           =   5640
      Begin VB.TextBox txtVersion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   39
         Top             =   75
         Width           =   5490
      End
   End
   Begin ComctlLib.TabStrip tabAbout 
      Height          =   3765
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   6641
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About Author"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Great Thanks To"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Readme"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Version History"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4650
      TabIndex        =   0
      Top             =   4162
      Width           =   1290
   End
   Begin VB.Label lblValue 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   8
      Left            =   150
      TabIndex        =   1
      Top             =   4125
      Width           =   4290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   150
      X2              =   5925
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   5925
      Y1              =   4050
      Y2              =   4050
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 21 Januari 2009
' 12:50 PM
'=======================================
' Module Common Dialogs
'=======================================
Private Const APPName = "al VirusScan"

Private Sub cmdOK_Click()
'    SoundBuffer = LoadResData("W1", "WAV")
'    sndPlaySound ByVal 0&, SND_NODEFAULT
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & APPName '& " (AVS)"
    lblValue(4).ForeColor = vbBlue
    lblTitle = APPName & " Final Version " & vAppVersion '& " New"
    lblValue(0) = "Developed By"
    lblValue(1) = "Company"
    lblValue(2) = "Town"
    lblValue(3) = "Copyright"
    lblValue(4) = ": Moh Aly Shodiqin"
    lblValue(5) = ": " & App.CompanyName
    lblValue(6) = ": Gresik City East Java - Indonesia"
    lblValue(7) = ": 2008-2009 Moh Aly Shodiqin. All rights reserved" '" & App.LegalCopyright
    lblValue(8) = "al VirusScan will protect your computer from viruses. This application final version."
    lblValue(9) = "E-mail"
    lblValue(10) = ": felix_progressif@yahoo.com"
    lblValue(11) = "- Allah S.W.T"
    lblValue(12) = "- Nabi Muhammad S.A.W"
    lblValue(13) = "- My Parents"
    lblValue(14) = "- My Soul"
    lblValue(15) = "- AVIGEN - VBBego Team"
    lblValue(16) = "- Noel A. Dacara"
    lblValue(17) = "- www.VBAccelerator.com"
    lblValue(18) = "- www.planetsourcecode.com"
    lblValue(19) = "- Bagus Judistira"
    lblValue(20) = "- Peradnya Dinata"
    lblValue(21) = "- Fred.cpp"
    lblValue(22) = "- Thanks to all for the suggestions and comments."
    lblValue(23) = "- Cyber Chris"
    lblValue(24) = "- Bobo RegEdit"
    lblValue(25) = "- Dung Le Nguyen"
    lblValue(26) = "License"
    lblValue(27) = ": GNU General Public License"
    lblValue(28) = "- Indonesian Programmer"
    lblValue(29) = "See readme for detail informations..."
    lblValue(30) = "- Andrei O. Lisovoi"
    lblValue(31) = "Blog"
    lblValue(32) = ": http://fi5ly.blogspot.com"
    lblValue(33) = "- Umair_11D"
    lblValue(34) = "- Thank for the somebody of made the component which is wearing al VirusScan."
    lblValue(35) = "- Jason Hensley"
    '-------------------------------------------------------
    lblValue(11).ToolTipText = "Terima kasih tuhan engkau telah memberikan segala nikmat kepadaku, semoga aku selalu menjadi hamba-hambamu yang selalu bersyukur."
    lblValue(12).ToolTipText = "Nabi Muhammad engkau adalah panutanku"
    lblValue(13).ToolTipText = "Maafkan anakmu ini yang tidak selalu menuruti kata-katamu, yang selalu menghabiskan uangmu"
    lblValue(14).ToolTipText = "Kamu adalah belahan jiwaku."
    lblValue(15).ToolTipText = "Thank's for your AVIGEN Engine"
    lblValue(16).ToolTipText = "Your source code (VB6 CLASSES FOR SCANNING THE SYSTEM) is my inspiration"
    lblValue(17).ToolTipText = "It's house Steve McMahon"
    lblValue(19).ToolTipText = "Your source code is very cool"
    lblValue(20).ToolTipText = "Your source code is very cool"
    lblValue(18).ToolTipText = "The home millions of lines of source code"
    lblValue(21).ToolTipText = "Thank's for your button (isButton)"
    lblValue(23).ToolTipText = "Thank's for your language translate"
    lblValue(25).ToolTipText = "Thank's for your source code"
    lblValue(28).ToolTipText = "Viva programmer indonesia"
    lblValue(30).ToolTipText = "Your Advanced Progress Bar is very cool ^_^"
    lblValue(33).ToolTipText = "Your source code is very cool (StylerButton)"
    lblValue(35).ToolTipText = "Thanks for example contains APIs and Structures using PSAPI.dll"
    '-------------------------------------------------------
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.hWnd, True
    Else
        AlwaysOnTop Me.hWnd, False
    End If
    '-------------------------------------------------------
    Readme
'    VDFDate
    VersionHistory
'    Sleep 300
'    SoundBuffer = LoadResData("W1", "WAV")
'    sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_LOOP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdOK_Click
End Sub

Private Sub lblValue_Click(Index As Integer)
    If Not lblValue(Index).Caption = ": Moh Aly Shodiqin" Then
        Exit Sub
    Else
        frmAuthor.show
    End If
End Sub

Private Sub tabAbout_Click()
    Dim pic As PictureBox
    For Each pic In picAbout
        pic.Visible = (pic.Index = tabAbout.SelectedItem.Index - 1)
    Next
End Sub

Private Sub Readme()
    On Error Resume Next
    Dim sFile As String
    sFile = nPath(App.path) & "\readme.txt"
    Dim sData As String
    Open sFile For Binary Access Read As #1
        sData = String(LOF(1), Chr(0))
        Get #1, , sData
    Close #1
    txtReadme = sData
End Sub

Private Sub VersionHistory()
    With txtVersion
        .Text = ""
            .Text = .Text & "al VirusScan History :" & vbCrLf & vbCrLf
                .Text = .Text & ""
                .Text = .Text & " - Scan Engine" & vbTab & vbTab & vbTab & " : " & vScanEngine & vbCrLf
                .Text = .Text & " - Scan With Virus Sample Engine " & vbTab & " : " & vScanWithVirusSample & vbCrLf
                .Text = .Text & " - Realtime Protection Engine " & vbTab & vbTab & " : " & vRealtimeProtection & vbCrLf
                .Text = .Text & " - Virus Definitions" & vbTab & vbTab & vbTab & " : " & vVirusDefinitions & vbCrLf
                .Text = .Text & " - VirusScan Process Manager" & vbTab & " : " & vProcessManager & vbCrLf
                .Text = .Text & " - VirusScan Registry Tweak" & vbTab & vbTab & " : " & vRegistryTweak & vbCrLf & vbCrLf
                .Text = .Text & ""
                .Text = .Text & "What's new in final version 1.6.8" & vbCrLf
                .Text = .Text & "" & vbCrLf
                .Text = .Text & "al VirusScan 1.6.8" & vbTab & " - 1 Maret 2009 10:22 PM" & vbCrLf
                .Text = .Text & " - added fungsi Scan With Virus Sample with new algorithm" & vbCrLf
                .Text = .Text & " - added Version History pada about" & vbCrLf
                .Text = .Text & " - fixed pada Realtime Protection" & vbCrLf
                .Text = .Text & " - 3 Maret 2009 2:49 AM" & vbCrLf
                .Text = .Text & "      - added progressbar scanning file pada al VirusScan & al VirusScan Scan With Virus Sample" & vbCrLf
                .Text = .Text & "      - added Fix Registry untuk memperbaiki sistem windows yng telah dirusak oleh virus" & vbCrLf
                .Text = .Text & "      - fixed VirusScan Process Manager v1.4" & vbCrLf
                .Text = .Text & " - 6 Maret 2009 10:21 AM" & vbCrLf
                .Text = .Text & "      - fixed Version History" & vbCrLf
                .Text = .Text & "      - added About Author al VirusScan" & vbCrLf
                .Text = .Text & " - 23 Maret 2009 11:21 AM" & vbCrLf
                .Text = .Text & "      - added Quarantine file in scan with virus sample" & vbCrLf
                .Text = .Text & "      - added Autorun Location" & vbCrLf
                .Text = .Text & "      - added performance in VirusScan Process Manager " & vbCrLf
                .Text = .Text & "      - added Process List in scan with virus sample " & vbCrLf
                .Text = .Text & "      - fixed Scan With Virus Sample Engine" & vbCrLf
                .Text = .Text & "      - fixed Update Now" & vbCrLf
                .Text = .Text & "      - added Check Update " & vbCrLf
                .Text = .Text & " - 10 April 2009 2:02 AM - Final Version" & vbCrLf
                .Text = .Text & "      - Changed quarantine extension as al" & vbCrLf
                .Text = .Text & "      - added Shutdown windows with powerfull functions" & vbCrLf
                .Text = .Text & "      - Changed button scan, pause, stop, exit and select local drive" & vbCrLf & vbCrLf
                .Text = .Text & ""
                .Text = .Text & "al VirusScan 1.6.7" & vbTab & " - 23 Februari 2009 10:57 AM" & vbCrLf
                .Text = .Text & " - added fungsi Send The Example Of Virus" & vbCrLf
                .Text = .Text & " - added Windows Security Settings pada Tweak Registry" & vbCrLf
                .Text = .Text & " - fixed pada fungsi update online" & vbCrLf
                .Text = .Text & " - 27 Februari 2009 9:01 AM" & vbCrLf
                .Text = .Text & "      - Changed icon dengan yang lebih matching dengan aplikasi" & vbCrLf & vbCrLf
                .Text = .Text & ""
                .Text = .Text & "al VirusScan 1.6.6" & vbTab & " - 8 Februari 2009" & vbCrLf
                .Text = .Text & " - fixed fungsi console" & vbCrLf
                .Text = .Text & " - Running on safe mode" & vbCrLf
                .Text = .Text & " - fixed fungsi Realtime Protection v1.2" & vbCrLf
                .Text = .Text & " - fixed pada fungsi Scan Engine 1.3" & vbCrLf
                .Text = .Text & " - added thirdparty VirusScan Registry Editor" & vbCrLf
                .Text = .Text & " - added Update Online" & vbCrLf
                .Text = .Text & " - Semua opsi virusscan disimpan dalam registry" & vbCrLf & vbCrLf
                .Text = .Text & ""
                .Text = .Text & "al VirusScan 1.0.6" & vbTab & " - 30 Januari 2009" & vbCrLf
                .Text = .Text & " - added opsi Run when windows start pada VirusScan Console" & vbCrLf & vbCrLf
                .Text = .Text & ""
                .Text = .Text & "al VirusScan 1.0.5" & vbTab & " - 29 Januari 2009" & vbCrLf
                .Text = .Text & " - added Realtime Protection" & vbCrLf
                .Text = .Text & " - VirusScan Console" & vbCrLf
                .Text = .Text & " - added  Set Attribute File or Folder" & vbCrLf & vbCrLf
                .Text = .Text & ""
                .Text = .Text & "al VirusScan 1.0.1" & vbTab & " -" & vbCrLf
                .Text = .Text & " - added Status Bar untuk status scanning" & vbCrLf
                .Text = .Text & " - Uflags Dialog BIF_RETURNONLYFSDIRS + BIF_EDITBOX" & vbCrLf & vbCrLf
                .Text = .Text & ""
                .Text = .Text & "al VirusScan 1.0.0" & vbTab & " - 20 Januari 2009" & vbCrLf
                .Text = .Text & " - Merancang form scan" & vbCrLf
                .Text = .Text & " - Scan Engine terbaru menggunakan AVIGEN Engine" & vbCrLf
                .Text = .Text & " - Buat Icon al VirusScan" & vbCrLf
                .Text = .Text & " - Uflags Dialog BIF_RETURNONLYFSDIRS + BIF_EDITBOX + BIF_BROWSEINCLUDEFILES" ' & vbCrLf
    End With
End Sub

