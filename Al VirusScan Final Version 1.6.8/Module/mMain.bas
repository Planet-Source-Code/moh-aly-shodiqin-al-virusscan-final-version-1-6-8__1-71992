Attribute VB_Name = "mMain"
Option Explicit

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngCC As Long
End Type

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Sub Main()
    Dim iccex As tagInitCommonControlsEx
    Select Case UCase$(Left$(Command, 2))
        Case "/RealtimeProtection"
'            REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "alVirusScan", nPath(App.path) & "\al VirusScan.exe /T"
            frmRTP.Hide
        Case Else
            With iccex
                .lngSize = LenB(iccex)
                .lngCC = ICC_USEREX_CLASSES
            End With
            InitCommonControlsEx iccex
            
            On Error GoTo 0
            frmRTP.Hide
    End Select
End Sub


