VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrgFaves 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "        Organise Favorites"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmOrgFaves.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   840
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrgFaves.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrgFaves.frx":2014
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrgFaves.frx":232E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   6240
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrgFaves.frx":340C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2655
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Height          =   1770
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   3122
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Move Up"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Move Down"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remove from favorites"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOrgFaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    UpdateMenu
    Unload Me
End Sub

Private Sub Form_Load()
    Dim hHeader As Long
    Dim z As Long
    Dim lItem As ListItem
    'Flat column headers
    hHeader = SendMessage(LV.hWnd, LVM_GETHEADER, 0, ByVal 0&)
    SetWindowLong hHeader, GWL_STYLE, GetWindowLong(hHeader, GWL_STYLE) Xor HDS_BUTTONS
    If fMainForm.mnuFavorites(1).Visible Then
        For z = 1 To fMainForm.mnuFavorites.count - 1
            Set lItem = LV.ListItems.Add(, , fMainForm.mnuFavorites(z).Caption, , 1)
            lItem.SubItems(1) = fMainForm.mnuFavorites(z).Tag
        Next
    End If
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim tmpCaption(0 To 1) As String, tmpPath(0 To 1) As String
    Select Case Button.Index
        Case 1
            If LV.SelectedItem.Index = 1 Then Exit Sub
            tmpCaption(0) = LV.SelectedItem.Text
            tmpCaption(1) = LV.ListItems(LV.SelectedItem.Index - 1).Text
            tmpPath(0) = LV.SelectedItem.SubItems(1)
            tmpPath(1) = LV.ListItems(LV.SelectedItem.Index - 1).SubItems(1)
            LV.SelectedItem.Text = tmpCaption(1)
            LV.ListItems(LV.SelectedItem.Index - 1).Text = tmpCaption(0)
            LV.SelectedItem.SubItems(1) = tmpPath(1)
            LV.ListItems(LV.SelectedItem.Index - 1).SubItems(1) = tmpPath(0)
        Case 2
            If LV.SelectedItem.Index = LV.ListItems.count Then Exit Sub
            tmpCaption(0) = LV.SelectedItem.Text
            tmpCaption(1) = LV.ListItems(LV.SelectedItem.Index + 1).Text
            tmpPath(0) = LV.SelectedItem.SubItems(1)
            tmpPath(1) = LV.ListItems(LV.SelectedItem.Index + 1).SubItems(1)
            LV.SelectedItem.Text = tmpCaption(1)
            LV.ListItems(LV.SelectedItem.Index + 1).Text = tmpCaption(0)
            LV.SelectedItem.SubItems(1) = tmpPath(1)
            LV.ListItems(LV.SelectedItem.Index + 1).SubItems(1) = tmpPath(0)
        Case 3
            LV.ListItems.Remove LV.SelectedItem.Index
    End Select
End Sub

Public Sub UpdateMenu()
    Dim z As Long
    If fMainForm.mnuFavorites.count > 2 Then
        For z = fMainForm.mnuFavorites.count - 1 To 2 Step -1
            Unload fMainForm.mnuFavorites(z)
        Next
    End If
    fMainForm.mnuFavorites(1).Visible = False
    If LV.ListItems.count > 0 Then
        For z = 1 To LV.ListItems.count
            If z > 1 Then Load fMainForm.mnuFavorites(z)
            fMainForm.mnuFavorites(z).Visible = True
            fMainForm.mnuFavorites(z).Caption = LV.ListItems(z).Text
            fMainForm.mnuFavorites(z).Tag = LV.ListItems(z).SubItems(1)
        Next
    End If

End Sub
