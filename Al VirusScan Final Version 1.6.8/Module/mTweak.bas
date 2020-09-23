Attribute VB_Name = "mTweak"
' Module    :   Tweak Registry...
'           :   ver 1.4
'           :   6 Februari 2009
'           :   1:24 PM
'           :   update v1.5 "5 Maret 2009 04:24 AM"
'           :   fixed SaveApp, GetApp & FixRegistry
'           :   Moh Aly Shodiqin
'--------------------------------------------------------------
Option Explicit

Public Const rExplorer = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            rEnum = "Software\Microsoft\Windows\CurrentVersion\Policies\NonEnum", _
            rWapp = "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", _
            rInternet = "Software\Policies\Microsoft\Internet Explorer\Restrictions", _
            rSystem = "Software\Microsoft\Windows\CurrentVersion\Policies\System", _
            rNetwork = "Software\Microsoft\Windows\CurrentVersion\Policies\Network", _
            rDesktop = "Control Panel\Desktop", _
            rAdvanced = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Const SMWC = "Software\Microsoft\Windows\CurrentVersion"
Dim REG As New cRegistry

Sub GetApp()
    On Error Resume Next
    Dim i As Integer, Isi As String, tmp
    
    With frmTweak
        For i = 0 To .chkSystem.count - 1
           Isi = Trim(.chkSystem(i).Tag)
           Select Case i
                Case 0, 1, 7, 8 To 11
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rSystem, Isi)
                    tmp = REG.GetSettingLong(HKEY_LOCAL_MACHINE, rSystem, Isi)
                Case 2 To 5, 12, 13, 17 To 20, 21, 22, 24, 25
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rExplorer, Isi)
                    tmp = REG.GetSettingLong(HKEY_LOCAL_MACHINE, rExplorer, Isi)
                Case 6
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rDesktop, Isi)
                Case 14
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rAdvanced, Isi)
                Case 15
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", Isi)
                    If Trim(tmp) <> 1 Then
                        tmp = 0
                    Else
                        tmp = 1
                    End If
                Case 16
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", Isi)
                    If Trim(tmp) = 0 Then
                        tmp = 1
                    Else
                        tmp = 0
                    End If
                Case 23
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", Isi)
                Case 26 To 30
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rInternet, Isi)
                Case 31
                    tmp = REG.GetSettingLong(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", Isi)
                Case 32
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rExplorer, Isi)
                Case 33, 34
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rSystem, Isi)
                Case 35
                    tmp = REG.GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", Isi)
            End Select
                .chkSystem(i).Value = Val(tmp)
           DoEvents
        Next i
    End With
End Sub

Sub SaveApp()
    On Error Resume Next
    Dim i As Integer, Isi As String
    
    With frmTweak
    For i = 0 To .chkSystem.count - 1
       Isi = Trim(.chkSystem(i).Tag)
       Select Case i
            Case 0, 1, 7, 8 To 11
                CekReg .chkSystem(i).Value, HKEY_CURRENT_USER, rSystem, Isi, 1
                CekReg .chkSystem(i).Value, HKEY_LOCAL_MACHINE, rSystem, Isi, 1
            Case 2 To 5, 12, 13, 17 To 20, 21, 22, 24, 25
                CekReg .chkSystem(i).Value, HKEY_CURRENT_USER, rExplorer, Isi, 1
                CekReg .chkSystem(i).Value, HKEY_LOCAL_MACHINE, rExplorer, Isi, 1
            Case 6
                CekReg .chkSystem(i).Value, HKEY_CURRENT_USER, rDesktop, Isi, 1
            Case 14
                CekReg .chkSystem(i).Value, HKEY_CURRENT_USER, rAdvanced, Isi, 1
            Case 15
                If .chkSystem(15).Value = 1 Then
                    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, Isi, 1
                Else
                    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, Isi, 2
                End If
            Case 16
                If .chkSystem(16).Value = 0 Then
                    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, Isi, 1
                Else
                    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, Isi, 0
                End If
            Case 23
                CekReg .chkSystem(i).Value, HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", Isi, 1
            Case 26 To 30
                CekReg .chkSystem(i).Value, HKEY_CURRENT_USER, rInternet, Isi, 1
            Case 31
                CekReg .chkSystem(i).Value, HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", Isi, 1
            Case 32
                CekReg .chkSystem(i).Value, HKEY_CURRENT_USER, rExplorer, Isi, 1
            Case 33, 34
                CekReg .chkSystem(i).Value, HKEY_CURRENT_USER, rSystem, Isi, 1
            Case 35
                If .chkSystem(35).Value = 1 Then
                    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", Isi, 1
                Else
                    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", Isi, 0
                End If
        End Select
       DoEvents
    Next i
    End With
End Sub

Function CekReg(Nm As Boolean, Root As Long, path As String, Value As String, Tipe As Byte)
    On Error Resume Next
    
    If Nm = True Then
       Select Case Tipe
              Case 1
                    REG.SaveSettingLong Root, path, Value, 1
              Case 2
                    REG.SaveSettingByte Root, path, Value, 1
              Case 3
                    REG.SaveSettingString Root, path, Value, 1
      End Select
    Else
       REG.DeleteValue Root, path, Value
    End If
    
End Function

Public Function FixRegistry()
    On Error Resume Next
    ComName = NameOfTheComputer(PCName)
    UserCom = GetUserCom()
    DoEvents
    
    ' Repair system windows-------------------------------------
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "exefile\shell\open\command", vbNullString, Chr(34) & "%1" & Chr(34) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "piffile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "scrfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "regfile\shell\open\command", "", "regedit.exe %1"
    REG.DeleteValue HKEY_CURRENT_USER, rSystem, "DisableTaskMgr"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rSystem, "DisableTaskMgr"
    REG.DeleteValue HKEY_CURRENT_USER, rSystem, "DisableRegistryTools"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rSystem, "DisableRegistryTools"
    REG.DeleteValue HKEY_CURRENT_USER, rExplorer, "NoFolderOptions"
    REG.DeleteValue HKEY_CURRENT_USER, rExplorer, "NoFind"
    REG.DeleteValue HKEY_CURRENT_USER, rExplorer, "NoRun"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rExplorer, "NoFolderOptions"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rExplorer, "NoFind"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rExplorer, "NoRun"
        
    ' Hidden files or folder-------------------------------------
    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, "Hidden", 2
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "CheckedValue", 2
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "DefaultValue", 2
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "Bitmap", "%SystemRoot%\system32\SHELL32.dll,4"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "Text", "@shell32.dll,-30499"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "Type", "group"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\NOHIDDEN", "CheckedValue", 2
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\NOHIDDEN", "DefaultValue", 2
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\NOHIDDEN", "Text", "@shell32.dll,-30501"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\NOHIDDEN", "Type", "radio"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\SHOWALL", "CheckedValue", 1
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\SHOWALL", "DefaultValue", 2
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\SHOWALL", "Text", "@shell32.dll,-30500"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\SHOWALL", "Type", "radio"

    ' Hide extensions--------------------------------------------
    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, "HideFileExt", 1
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "CheckedValue", 1
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "DefaultValue", 1
    REG.DeleteValue HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "HideFileExt"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "Text", "@shell32.dll,-30503"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "Type", "checkbox"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "UncheckedValue", 0

    ' Show super hiddens-----------------------------------------
    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, "ShowSuperHidden", 0
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto", "0"
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "@shell32.dll,-30508", "Hide protected operating system files (Recommended)"
    REG.SaveSettingString HKEY_USERS, "S-1-5-21-1417001333-1060284298-725345543-500\Software\Microsoft\Windows\ShellNoRoam\MUICache", "@shell32.dll,-30508", "Hide protected operating system files (Recommended)"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden", "Text", "@shell32.dll,-30508"

    ' Registered Organization & Registered Owner-----------------
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", UserCom
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", PCName

    ' Show Full Path at Address Bar------------------------------
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPathAddress", 1

    ' 4k51k4-----------------------------------------------------
    REG.DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    REG.DeleteKey HKEY_USERS, "S-1-5-21-1547161642-1343024091-725345543-500\Software\Policies\Microsoft\Windows\System"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", 0
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", 0
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    REG.SaveSettingString HKEY_CURRENT_USER, "Control Panel\Desktop\", "SCRNSAVE.EXE", ""
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 0
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 0
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe "
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "userinit.exe"

    ' Amburadul.Hokage Killer------------------------------------
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PaRaY_VM"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ConfigVir"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "NviDiaGT"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "NarmonVirusAnti"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "AVManager"
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Window Title", ""
    REG.DeleteValue HKEY_LOCAL_MACHINE, " SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA"
    REG.DeleteValue HKEY_CLASSES_ROOT, "exefile", "NeverShowExt"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msconfig.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\rstrui.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\wscript.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\mmc.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\procexp.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msiexec.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\taskkill.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\cmd..exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\tasklist.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\HokageFile.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Rin.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Obito.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\KakashiHatake.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\HOKAGE4.exe"

    ' Flu_Ikan--------------------------------------------------
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "kebodohan"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "pemalas"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "mulut_besar"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "otak_udang"
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.microsoft.com/isapi/redir.dll?prd={SUB_PRD}&clcid={SUB_CLSID}&pver={SUB_PVER}&ar=home"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.microsoft.com/isapi/redir.dll?prd={SUB_PRD}&clcid={SUB_CLSID}&pver={SUB_PVER}&ar=home"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmboot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmio.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmload.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\sermouse.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\sr.sys", "", "FSFilter System Recovery"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\vga.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\vgasave.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\dmboot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\dmiot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpcdd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpdd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpwd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\sermouse.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\tdpipe.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\tdtcp.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\vga.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\vgasave.sys", "", "Driver"

    LockWindowUpdate (GetDesktopWindow())
    ForceCacheRefresh
    LockWindowUpdate (0)
    DoEvents
End Function


