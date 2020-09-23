VERSION 5.00
Begin VB.Form frmLongEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit DWORD value"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3960
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Base"
      Height          =   1095
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "Decimal"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hexadecimal"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Value name"
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Value name :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Value data :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmLongEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Really just a fancy inputbox
Dim CurVal As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim temp As String, temp2 As String
    temp = Right(fMainForm.TV.SelectedItem.Key, Len(fMainForm.TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    temp2 = Hex$(CurVal)
    If Len(temp2) > 6 Then
        Unload Me
        Exit Sub
    End If
    SaveSettingLong Val(fMainForm.LV.SelectedItem.Tag), temp, fMainForm.LV.SelectedItem.Text, CurVal
    If CurVal = 0 Then
        temp2 = "0x00000000 (0)"
    Else
        temp2 = LCase("0x" + String(8 - Len(temp2), "0") + temp2 + " (" + Trim(Str(CurVal)) + ")")
    End If
    fMainForm.LV.SelectedItem.SubItems(2) = temp2
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = fMainForm.Icon
    Text1.Text = fMainForm.LV.SelectedItem.Text
    GetCurSettings
End Sub

Private Sub Option1_Click()
    Dim temp As String
    temp = LCase(Hex$(CurVal))
    If Len(temp) < 7 Then
        Text2.Text = temp
    Else
        GetCurSettings
    End If
End Sub

Private Sub Option2_Click()
    Dim temp As String
    temp = LCase(Hex$(CurVal))
    If Len(temp) < 7 Then
        Text2.Text = CurVal
    Else
        GetCurSettings
        Text2.Text = CurVal
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Option1.Value Then
        If (KeyCode < 48 Or KeyCode > 57) And (KeyCode < 97 Or KeyCode > 102) Then
            If KeyCode <> 8 Then KeyCode = 0
        End If
    Else
        If (KeyCode < 48 Or KeyCode > 57) Then
            If KeyCode <> 8 Then KeyCode = 0
        End If
    End If

End Sub

Public Sub GetCurSettings()
    Dim temp As String, temp2 As String, z As Long
    temp = Right(fMainForm.TV.SelectedItem.Key, Len(fMainForm.TV.SelectedItem.Key) - InStr(1, fMainForm.TV.SelectedItem.Key, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    CurVal = GetSettingLong(fMainForm.LV.SelectedItem.Tag, temp, fMainForm.LV.SelectedItem.Text, 0)
    Text2.Text = Hex$(CurVal)

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Option1.Value Then
        CurVal = HexToDec(Text2.Text)
    Else
        CurVal = Val(Text2.Text)
    End If

End Sub
