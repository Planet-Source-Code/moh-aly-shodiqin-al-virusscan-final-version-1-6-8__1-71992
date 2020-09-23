VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaves 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organise Favorites"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "frmFaves.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFaves 
      Caption         =   "Close"
      Height          =   330
      Index           =   4
      Left            =   4755
      TabIndex        =   6
      Top             =   3720
      Width           =   1530
   End
   Begin VB.CommandButton cmdFaves 
      Caption         =   "Delete"
      Height          =   330
      Index           =   3
      Left            =   1845
      TabIndex        =   5
      Top             =   1440
      Width           =   1530
   End
   Begin VB.CommandButton cmdFaves 
      Caption         =   "Move to Folder..."
      Height          =   330
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1530
   End
   Begin VB.CommandButton cmdFaves 
      Caption         =   "Rename"
      Height          =   330
      Index           =   1
      Left            =   1845
      TabIndex        =   3
      Top             =   1050
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3435
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   6059
      _Version        =   393217
      Indentation     =   476
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFaves 
      Caption         =   "Create Folder"
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1050
      Width           =   1530
   End
   Begin MSComctlLib.ImageList LVImageList 
      Left            =   3120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFaves.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFaves.frx":2294
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   $"frmFaves.frx":282E
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmFaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
