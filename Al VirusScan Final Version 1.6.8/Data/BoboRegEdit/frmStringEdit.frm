VERSION 5.00
Begin VB.Form frmStringEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit String"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Value name"
      Top             =   360
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Value data :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Value name :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmStringEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Just an inputbox
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim temp As String, tempName As String
    temp = Right(fMainForm.TV.SelectedItem.Key, Len(fMainForm.TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    tempName = fMainForm.LV.SelectedItem.Text
    If tempName = "(Default)" Then tempName = ""
    SaveSettingString Val(fMainForm.LV.SelectedItem.Tag), temp, tempName, Text2.Text
    fMainForm.LV.SelectedItem.SubItems(2) = Text2.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Dim temp As String
    Me.Icon = fMainForm.Icon
    temp = Right(fMainForm.TV.SelectedItem.Key, Len(fMainForm.TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    Text1.Text = fMainForm.LV.SelectedItem.Text
    Text2.Text = GetSettingString(Val(fMainForm.LV.SelectedItem.Tag), temp, fMainForm.LV.SelectedItem.Text, "")
End Sub
