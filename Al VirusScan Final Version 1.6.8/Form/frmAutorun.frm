VERSION 5.00
Begin VB.Form frmAutorun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   Icon            =   "frmAutorun.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   6600
      ScaleHeight     =   540
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   3150
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   225
      ScaleHeight     =   990
      ScaleWidth      =   1965
      TabIndex        =   8
      Top             =   3825
      Width           =   1965
      Begin VB.ComboBox Combo5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmAutorun.frx":08CA
         Left            =   75
         List            =   "frmAutorun.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   75
         Width           =   1815
      End
      Begin VB.CommandButton cmdKillAutorun 
         Caption         =   "Kill Autorun.inf"
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
         Left            =   75
         TabIndex        =   9
         Top             =   525
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Autorun Drive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   150
      TabIndex        =   7
      Top             =   3600
      Width           =   2115
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
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
      TabIndex        =   6
      Top             =   4500
      Width           =   1515
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
      Left            =   5700
      TabIndex        =   5
      Top             =   4500
      Width           =   1515
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2700
      Width           =   7065
   End
   Begin VB.ComboBox cboAutorun 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmAutorun.frx":08CE
      Left            =   150
      List            =   "frmAutorun.frx":08D0
      TabIndex        =   2
      Top             =   3150
      Width           =   3540
   End
   Begin VB.ListBox lstAutorun 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   7065
   End
   Begin VB.ListBox List6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2475
      TabIndex        =   1
      Top             =   1725
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.ListBox List5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2475
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuKill 
         Caption         =   "Kill Autorun.inf in selected drive"
         Index           =   0
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Kill Autorun.inf In All Drive"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmAutorun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As ESHGetFileInfoFlagConstants) As Long

Private Type SHFILEINFO
    hIcon           As Long ' : icon
    iIcon           As Long ' : icondex
    dwAttributes    As Long ' : SFGAO_ flags
    szDisplayName   As String * MAX_PATH ' : display name (or path)
    szTypeName      As String * 80 ' : type name
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Const KEY_QUERY_VALUE = &H1
'Private Const MAX_PATH = 260

Private Enum RegDataTypes
    REG_SZ = 1                         ' Unicode nul terminated string
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    REG_DWORD = 4                      ' 32-bit number
End Enum

Private Enum RegistryKeys
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Enum ValKey
    Values = 0
    Keys = 1
End Enum

Private Type ByteArray
  FirstByte As Byte
  ByteBuffer(255) As Byte
End Type

Dim baData As ByteArray
Private shinfo As SHFILEINFO

Private Function OpenKey(RegistryKey As RegistryKeys, Optional SubKey As String) As Long
    If OpenKey <> 0 Then RegCloseKey (OpenKey)
    RegOpenKeyEx RegistryKey, SubKey, 0, KEY_QUERY_VALUE, OpenKey
End Function

Private Function GetCount(RegisteryKeyHandle As Long, ValuesOrKeys As ValKey) As Long
    If ValuesOrKeys = Keys Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, GetCount, 0, 0, 0, 0, 0, 0, 0
    If ValuesOrKeys = Values Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, GetCount, 0, MAX_PATH + 1, 0, 0
End Function

Private Function EnumKey(RegisteryKeyHandle As Long, KeyIndex As Long) As String
    EnumKey = SPACE(MAX_PATH + 1)
    RegEnumKey RegisteryKeyHandle, KeyIndex, EnumKey, MAX_PATH + 1
    EnumKey = Trim(EnumKey)
End Function

Private Function EnumValue(RegisteryKeyHandle As Long, KeyIndex As Long) As String
    Dim lBufferLen As Long, i As Integer
    For i = 0 To 255
      baData.ByteBuffer(i) = 0
    Next
    lBufferLen = 255
    EnumValue = SPACE(MAX_PATH + 1)
    RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, 0, lValNameLen, lValLen, 0, 0
    RegEnumValue RegisteryKeyHandle, KeyIndex, EnumValue, MAX_PATH + 1, 0, 0, baData.FirstByte, lBufferLen
    EnumValue = Trim(EnumValue)
End Function

Private Function DeleteValue(RegisteryKeyHandle As Long, KeyName As String) As Long
    DeleteValue = RegDeleteValue(RegisteryKeyHandle, KeyName)
End Function

Private Function SetValue(RegisteryKeyHandle As RegistryKeys, SubRegistryKey As String, KeyName As String, newValue As String, Optional DataType As RegDataTypes)
    Dim lRetVal As Long
    lRetVal = OpenKey(RegisteryKeyHandle, SubRegistryKey)
    If DataType = 0 Then DataType = REG_SZ
    RegSetValueEx lRetVal, KeyName, 0, DataType, newValue, LenB(StrConv(SubKeyValue, vbFromUnicode))
End Function

Private Function GetKeyValue(hKey As Long, KeyName As String) As String
    Dim i As Long
    Dim rc As Long
    
    Dim hDepth As Long
    Dim sKeyVal As String
    Dim lKeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    rc = RegQueryValueEx(hKey, KeyName, 0, lKeyValType, tmpVal, KeyValSize)
    GetKeyValue = Trim(tmpVal)
    
End Function

Function InitAutorun()
    On Error Resume Next
    
    Dim hKey As Long
    Dim lCount As Long
    Dim i As Long
    
    lstAutorun.Clear
    List5.Clear
    List6.Clear
    
    Select Case cboAutorun.Text
        Case "HKLM - Run"
        hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
        lCount = GetCount(hKey, Values)
        For i = 0 To lCount - 1
            lstAutorun.AddItem EnumValue(hKey, i)
            List5.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
            List6.AddItem "HKLM - Run"
        Next i
    Case "HKLM - RunOnce"
        hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
        lCount = GetCount(hKey, Values)
        For i = 0 To lCount - 1
            lstAutorun.AddItem EnumValue(hKey, i)
            List5.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
            List6.AddItem "HKLM - RunOnce"
        Next i
    Case "HKLM - RunOnceEx"
        hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx")
        lCount = GetCount(hKey, Values)
        For i = 0 To lCount - 1
            lstAutorun.AddItem EnumValue(hKey, i)
            List5.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
            List6.AddItem "HKLM - RunOnceEx"
        Next i
    Case "HKLM - RunServices"
        hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
        lCount = GetCount(hKey, Values)
        For i = 0 To lCount - 1
            lstAutorun.AddItem EnumValue(hKey, i)
            List5.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
            List6.AddItem "HKLM - RunServices"
        Next i
    Case "HKCU - Run"
        hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
        lCount = GetCount(hKey, Values)
        For i = 0 To lCount - 1
            lstAutorun.AddItem EnumValue(hKey, i)
            List5.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
            List6.AddItem "HKCU - Run"
        Next i
    Case "HKCU - PoliciesRun"
        hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run")
        lCount = GetCount(hKey, Values)
        For i = 0 To lCount - 1
            lstAutorun.AddItem EnumValue(hKey, i)
            List5.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
            List6.AddItem "HKCU - PoliciesRun"
        Next i
    Case "HKCU - RunOnce"
        hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
        lCount = GetCount(hKey, Values)
        For i = 0 To lCount - 1
            lstAutorun.AddItem EnumValue(hKey, i)
            List5.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
            List6.AddItem "HKCU - RunOnce"
        Next i
    Case "HKCU - Windows"
        hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows")
        lCount = GetCount(hKey, Values)
        For i = 0 To lCount - 1
            lstAutorun.AddItem EnumValue(hKey, i)
            List5.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
            List6.AddItem "HKCU - Windows"
        Next i
    
    Dim fso As New FileSystemObject
    Dim sFolder As Folder
    Dim sFiles As Files
    Dim sFile As file
    
    Case "Scheduled Task"
        Set sFolder = fso.GetFolder("C:\Windows\Tasks")
        Set sFiles = sFolder.Files
        If sFiles.count > 0 Then
            For Each sFile In sFiles
                lstAutorun.AddItem (sFile.name)
                List5.AddItem sFile.path
                List6.AddItem "Scheduled Task"
            Next
        End If
    Case "User Startup"
        Dim strUserProfile As String
        strUserProfile = Environ$("UserProfile") & "\Start Menu\Programs\Startup"
        Set sFolder = fso.GetFolder(strUserProfile)
        Set sFiles = sFolder.Files
        If sFiles.count > 0 Then
            For Each sFile In sFiles
                lstAutorun.AddItem (sFile.name)
                List5.AddItem sFile.path
                List6.AddItem "User Startup"
            Next
        End If
    Case "Common Startup"
        Set sFolder = fso.GetFolder("C:\Documents and Settings\All Users\Start Menu\Programs\Startup")
        Set sFiles = sFolder.Files
        If sFiles.count > 0 Then
            For Each sFile In sFiles
                lstAutorun.AddItem (sFile.name)
                List5.AddItem sFile.path
                List6.AddItem "All Users Startup"
            Next
        End If
    Case Else
        txtPath.Text = "Please choose autorun location..."
    End Select
End Function

Private Sub cboAutorun_Click()
    txtPath.Text = "Please choose autorun location..."
    picSample.Cls
    cmdDelete.Enabled = False
    InitAutorun
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete Autorun Location...?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        ClearAutorun
        cmdDelete.Enabled = False
        txtPath.Text = ""
    End If
    cmdDelete.Enabled = False
End Sub

Private Sub cmdKillAutorun_Click()
    PopupMenu mnuFile, , cmdKillAutorun.Left + 220, cmdKillAutorun.Top + cmdKillAutorun.Height + 3800
End Sub

Private Sub Form_Load()
    With cboAutorun
        .Text = "HKLM - Run"
        .AddItem "HKCU - PoliciesRun"
        .AddItem "HKCU - Run"
        .AddItem "HKCU - RunOnce"
        .AddItem "HKCU - Windows"
        .AddItem "HKLM - Run"
        .AddItem "HKLM - RunOnce"
        .AddItem "HKLM - RunOnceEx"
        .AddItem "Scheduled Task"
        .AddItem "User Startup"
        .AddItem "Common Startup"
    End With
    Me.Caption = "Autorun Location"
    InitAutorun
    GetDrive
End Sub

Private Sub lstAutorun_Click()
    On Error GoTo Salah
    List5.Selected(lstAutorun.ListIndex) = True
    List6.Selected(lstAutorun.ListIndex) = True
    txtPath.Text = List5.Text
    If txtPath <> "" Then
        picSample.Cls
        RetrieveIcon List5.Text, picSample, ricnLarge
    End If
    If txtPath <> "" Then cmdDelete.Enabled = True
Salah:
End Sub

Function ClearAutorun()
    Dim i As Long
    Dim tmp As Long
    Dim fso As New FileSystemObject
    
'    For i = 1 To lstAutorun.ListCount - 1
'        If lstAutorun.Selected(i) = True Then
            Select Case List6.Text
                Case "HKLM - RunServices"
                    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", lstAutorun.Text
                Case "HKLM - Run"
                    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", lstAutorun.Text
                Case "HKCU - Run"
                    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", lstAutorun.Text
                Case "HKCU - PoliciesRun"
                    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", lstAutorun.Text
                Case "HKLM - RunOnce"
                    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce", lstAutorun.Text
                Case "HKLM - RunOnceEx"
                    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx", lstAutorun.Text
                Case "HKCU - RunOnce"
                    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", lstAutorun.Text
                Case "HKCU - Windows"
                    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", lstAutorun.Text
                Case Else
                    fso.DeleteFile txtPath.Text, True
            End Select
'        End If
'    Next i
    
    Set fso = Nothing
    Call InitAutorun
End Function

Public Function GetDrive()
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    
    On Error Resume Next
    
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
        If (drv.DriveType = 1) And (drv.DriveLetter <> "A") Then
            Combo5.AddItem drv.DriveLetter & ":\"
        End If
    Next
    If Combo5.ListCount = 0 Then
        Combo5.AddItem "None"
    End If
    Combo5.ListIndex = 0
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
End Function

Function ClearAuto()
    On Error Resume Next
    If IsFileExist(Combo5.Text & "autorun.inf") = True Then
        DeleteFile Combo5.Text & "autorun.inf"
        Call MsgBox("File deleted Successfully !", vbOKOnly + vbInformation, Me.Caption)
    Else
        MsgBox "Autorun.inf not found !", vbOKOnly + vbCritical, Me.Caption
    End If
End Function

Function ClearAutoAllDrive()
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    
    On Error Resume Next
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
        Kill drv.DriveLetter & ":\autorun.inf"
    Next
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
End Function

Function IsFileExist(sPath As String) As Boolean
    If PathFileExists(sPath) = 1 And PathIsDirectory(sPath) = 0 Then
        IsFileExist = True
    Else
        IsFileExist = False
    End If
End Function

Private Sub mnuKill_Click(Index As Integer)
    Select Case Index
        Case 0
            If MsgBox("Delete autorun.inf in drive " & Combo5.Text & " ?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                 ClearAuto
            End If
        Case 1
            If MsgBox("Delete autorun.inf in all drives...?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                 ClearAutoAllDrive
                 Call MsgBox("All autorun.inf was deleted!", vbOKOnly + vbInformation, Me.Caption)
            End If
    End Select
End Sub

Sub RetrieveIcon(fName As String, DC As PictureBox, icnSize As IconRetrieve)
    Dim hImgLarge As Long  'the handle to the system image list
    
    If icnSize = ricnLarge Then
        hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    End If
End Sub
