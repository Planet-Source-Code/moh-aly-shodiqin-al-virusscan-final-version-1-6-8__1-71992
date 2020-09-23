Attribute VB_Name = "mSystemTray"
    '******************************************************************************
    'Systray Module
    '
    'Mark Mokoski
    'markm@cmtelephone.com
    'www.cmtelephone.com
    '
    '6-NOV-2004
    '
    'Put App in SysTray, remove App from SysTray, Form on top, Balloon ToolTip code
    '
    'See Systray Form Code.txt in the ZIP file for form add-in's to make it all work
    '
    'Also see Microsoft Knowledge base http://support.microsoft.com/default.aspx?scid=kb;en-us;149276
    'for more information.
    '
    'This code is based on the Microsoft Knowledge Base code.
    '******************************************************************************

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private blnClick As Boolean
Private vbTray As NOTIFYICONDATA

'Public Const SWP_NOMOVE As Long = &H2
'Public Const SWP_NOSIZE As Long = &H1
Public Const flags As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONCLK As Long = &H204
Public Const WM_LBUTTONCLK As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_MOUSEMOVE As Long = &H200
Public Const NIM_ADD As Long = &H0
Public Const NIM_DELETE As Long = &H2
Public Const NIF_ICON As Long = &H2
Public Const NIF_MESSAGE As Long = &H1
Public Const NIM_MODIFY As Long = &H1
Public Const NIF_TIP As Long = &H4
Public Const NIF_INFO As Long = &H10
Public Const NIS_HIDDEN As Long = &H1
Public Const NIS_SHAREDICON As Long = &H2

Public Enum TypeBallon
    NIIF_NONE = &H0
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_INFO = &H1
    NIIF_GUID = &H4
End Enum

Private Const HWND_NOTOPMOST As Long = -2
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Public Sub SystrayOn(frm As Form, IconTooltipText As String)

    'Adds Icon to SysTray

        With vbTray
            .cbSize = Len(vbTray)
            .hwnd = frm.hwnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .szTip = Trim(IconTooltipText$) & vbNullChar
            .hIcon = frm.Icon
        End With
    
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
    
End Sub

Public Sub SystrayOff(frm As Form)

    'Removes Icon from SysTray

        With vbTray
            .cbSize = Len(vbTray)
            .hwnd = frm.hwnd
            .uID = vbNull
        End With
    
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
    
End Sub

Public Sub ChangeSystrayToolTip(frm As Form, IconTooltipText As String)

    'Changes the SysTray Balloon Tool Tip Text

        With vbTray
            .cbSize = Len(vbTray)
            .hwnd = frm.hwnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .szTip = Trim(IconTooltipText$) & vbNullChar
            .hIcon = frm.Icon
        End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)
    
End Sub

Public Sub FormOnTop(frm As Form)

    'Puts your form ontop of all the other windows!
'    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)

End Sub

Public Sub PopupBalloon(frm As Form, Message As String, Title As String, Optional balType As TypeBallon = NIIF_INFO)

    'Set a Balloon tip on Systray

    'Call RemoveBalloon(frm), This removes any current Balloon Tip that is active.
    'If you want Balloon Tips to "Stack up" and display in sequence
    'after each times out (or you click on the Balloon Tip to clear it),
    'comment out the Call below.

    Call RemoveBalloon(frm)

        With vbTray
            .cbSize = Len(vbTray)
            .hwnd = frm.hwnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = frm.Icon
            .dwState = 0
            .dwStateMask = 0
            .szInfo = Message & Chr(0)
            .szInfoTitle = Trim(Title) & vbNullChar
            'Choose the message icon below, NIIF_NONE, NIIF_WARNING, NIIF_ERROR, NIIF_INFO
            .dwInfoFlags = balType
        End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)

End Sub

Public Sub RemoveBalloon(frm As Form)

    'Kill any current Balloon tip on screen for referenced form
  
        With vbTray
            .cbSize = Len(vbTray)
            .hwnd = frm.hwnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = frm.Icon
            .dwState = 0
            .dwStateMask = 0
            .szInfo = Chr(0)
            .szInfoTitle = Chr(0)
            .dwInfoFlags = NIIF_NONE
        End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)

End Sub






