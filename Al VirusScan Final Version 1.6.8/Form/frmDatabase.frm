VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Begin VB.Form frmDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvwDB 
      Height          =   3615
      Left            =   150
      TabIndex        =   0
      Top             =   750
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Virus Name"
         Object.Width           =   5363
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Virus Type"
         Object.Width           =   2716
      EndProperty
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
      Left            =   3975
      TabIndex        =   1
      Top             =   4500
      Width           =   1515
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   2175
      TabIndex        =   5
      Top             =   390
      Width           =   45
   End
   Begin VB.Label Label1 
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
      Index           =   2
      Left            =   2175
      TabIndex        =   4
      Top             =   150
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "Virus Definitions Update"
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
      Left            =   150
      TabIndex        =   3
      Top             =   390
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Virus Definitions"
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
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   1365
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 30 Januari 2009
' 1:08 PM
'=======================================
' Module Database Viruses
'=======================================
Option Explicit

Dim WithEvents Engine32 As cEngine32
Attribute Engine32.VB_VarHelpID = -1

Private Sub cmdClose_Click()
    Set Engine32 = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Database Viruses"
    InitDB
    
    If REG.GetSettingLong(HKEY_CURRENT_USER, "Software\DQ Soft\al VirusScan\Console", "AlwaysOnTop", 1) = 1 Then
        AlwaysOnTop Me.hWnd, True
    Else
        AlwaysOnTop Me.hWnd, False
    End If
End Sub

Private Sub InitDB()
    On Error Resume Next
    Set Engine32 = New cEngine32
    Engine32.ClassIDApartement = Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(1) & Chr(255)
    
    Dim data As New Collection, Data1 As New Collection
    Engine32.GetVirusListInfo data
'    Engine32.GetDefinitionDate Data1
'    GetDefinitionDate Data1
    If data.count > 0 Then
       Dim i As Long, j As String
       Dim lv As ListItem
       For i = 1 To data.count
            j = Format(i, "00#") & ". " & UCase(data(i)(0))
            Set lv = lvwDB.ListItems.Add(, , j)
                lv.SubItems(1) = UCase(data(i)(1))
       Next i
    End If
    '--------------------------------------------------
'    If Data1.count > 0 Then
'       For i = 1 To data.count
'            j = Data1(i)(0)
'       Next i
'    End If
'    VDFDate
    Me.Caption = "Virus List"
    Label1(2) = ": " & lvwDB.ListItems.count
    Label1(3) = ": " & vVirusDefinitions
    If CDate(Month(vVirusDefinitions)) < Month(Date) Then
'        Label1(3).ForeColor = vbBlue
        Label1(3).ToolTipText = "It is requiered to update your virus definitions..."
    End If
End Sub

Private Sub Form_Initialize()
    lvwStyle lvwDB
'    SetFlatHeaders lvwDB.hWnd
End Sub

' update 15 Februari 2009
Private Sub lvwDB_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvwDB.Sorted = True
    lvwDB.SortKey = ColumnHeader.Index - 1
    If lvwDB.SortOrder = lvwDescending Then
       lvwDB.SortOrder = lvwAscending
    Else
       lvwDB.SortOrder = lvwDescending
    End If
End Sub
