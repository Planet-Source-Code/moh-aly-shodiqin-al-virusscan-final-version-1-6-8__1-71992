VERSION 5.00
Begin VB.UserControl dcButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1365
   ClipBehavior    =   0  'None
   DefaultCancel   =   -1  'True
   HitBehavior     =   0  'None
   LockControls    =   -1  'True
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   91
   ToolboxBitmap   =   "dcButton.ctx":0000
End
Attribute VB_Name = "dcButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
                                                                                                                                                                #If False Then ' Ooopss!!! Now you see me :)

    -> BlendColors
    -> CalculateRects
    -> ColorScheme
    -> CreateButtonRegion
    -> DrawBackgroundFromParent
    -> DrawButton
    -> DrawButtonCrystalMac
    -> DrawButtonOffice2003
    -> DrawButtonOfficeXP
    -> DrawButtonOpera
    -> DrawButtonStandard
    -> DrawButtonXPStyle
    -> DrawButtonXPToolbar
    -> DrawButtonYahoo
    -> DrawCaptionEffect
    -> DrawGradientEx
    -> DrawIconEffect
    -> DrawShineEffect
    -> GetAccessKey
    -> OverrideColor
    -> PopupMenu
    -> SetButtonColors
    -> ShiftColor

' ################################################################################
' THIS CONTROL IS FREE FOR USE BY ANYONE. PLEASE READ README.TXT FOR OTHER DETAILS
                                                                                                                                                                #End If ' I bet you also like this idea, do you?
Option Explicit

#Const USE_CRYSTAL = True
#Const USE_MAC = True
#Const USE_MACOSX = True
#Const USE_OFFICE2003 = True
#Const USE_OFFICEXP = True
#Const USE_OPERABROWSER = True
#Const USE_STANDARD = True
#Const USE_XPBLUE = True
#Const USE_XPOLIVEGREEN = True
#Const USE_XPSILVER = True
#Const USE_XPTOOLBAR = True
#Const USE_YAHOO = True

' You can also unset some features not used in the control to save some more space
#Const USE_POPUPMENU = True
#Const USE_SPECIALEFFECTS = True

' About section APIs and a user-defined constant
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Const DC_URL As String = "http://dcbutton.dacarasoftwares.cjb.net/"
    Private Const SW_SHOWNORMAL As Long = 1

' Color convertion API
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

' Create transparent areas on the control
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Private Const RGN_OR As Long = 2
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' Cursor tracking APIs
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
    Private Type POINTAPI
        x As Long
        Y As Long
    End Type
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long ' Win98 or later
Private Declare Function TrackMouseEvent2 Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long ' Win95 w/ IE 3.0
    Private Const TME_LEAVE     As Long = &H2
    Private Const WM_ACTIVATE   As Long = &H6
    Private Const WM_MOUSELEAVE As Long = &H2A3
    Private Const WM_NCACTIVATE As Long = &H86
    Private Type TRACKMOUSEEVENTTYPE
        cbSize      As Long
        dwFlags     As Long
        hwndTrack   As Long
        dwHoverTime As Long
    End Type
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' Determines if a function is supported by a library
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
'Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

' Determines if the control's parent form/window is an MDI child window
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Const GWL_EXSTYLE    As Long = -20
    Private Const WS_EX_MDICHILD As Long = &H40&

' Determine if the system is in the NT platform (for unicode support)
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
    Private Const VER_PLATFORM_WIN32_NT As Long = 2
    Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion      As Long
        dwMinorVersion      As Long
        dwBuildNumber       As Long
        dwPlatformId        As Long
        szCSDVersion        As String * 128 ' Maintenance string for PSS usage
    End Type

' Drawing APIs (GDI32 library)
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Const SRCCOPY As Long = &HCC0020
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
#If USE_STANDARD Then
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
#End If
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Private Const PS_SOLID As Long = 0
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Const BI_RGB As Long = 0&
    Private Type BITMAPINFOHEADER
        biSize          As Long
        biWidth         As Long
        biHeight        As Long
        biPlanes        As Integer
        biBitCount      As Integer
        biCompression   As Long
        biSizeImage     As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed       As Long
        biClrImportant  As Long
    End Type
    Private Type RGBQUAD
        rgbBlue     As Byte
        rgbGreen    As Byte
        rgbRed      As Byte
        'rgbReserved As Byte ' Removed so that we can use RGBQUAD as
                             ' datatype for the bitmap data (lpBits)
    End Type
    Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
    End Type
Private Declare Function GetNearestColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal lpPoint As Any) As Long ' Modified
Private Declare Function PatBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
    Private Const PATCOPY As Long = &HF00021
Private Declare Function Polyline Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Const DIB_RGB_COLORS As Long = 0
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' Drawing APIs (User32 library)
Private Declare Function CopyRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Long
    Private Type RECT
        Left    As Long
        Top     As Long
        Right   As Long
        bottom  As Long
    End Type
#If USE_STANDARD Then
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
    Private Const BDR_RAISEDINNER As Long = &H4
    Private Const BDR_RAISEDOUTER As Long = &H1
    Private Const BDR_SUNKENINNER As Long = &H8
    Private Const BDR_SUNKENOUTER As Long = &H2
    Private Const BF_BOTTOM     As Long = &H8
    Private Const BF_LEFT       As Long = &H1
    Private Const BF_RIGHT      As Long = &H4
    Private Const BF_TOP        As Long = &H2
    Private Const BF_RECT       As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    Private Const BF_SOFT       As Long = &H1000
    Private Const EDGE_RAISED   As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    Private Const EDGE_SUNKEN   As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
#End If
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
    Private Const DI_NORMAL As Long = &H3
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' Drawing text in ansi/unicode
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long ' Modified
    Private Const DT_CALCRECT   As Long = &H400
    Private Const DT_CENTER     As Long = &H1
    Private Const DT_NOCLIP     As Long = &H100 ' Allow text to exceed specified drawing area (necessary for the vertical text effect)
    Private Const DT_WORDBREAK  As Long = &H10
    Private Const DT_CALCFLAG   As Long = DT_WORDBREAK Or DT_CALCRECT Or DT_NOCLIP Or DT_CENTER
    Private Const DT_DRAWFLAG   As Long = DT_WORDBREAK Or DT_NOCLIP Or DT_CENTER

' Load hand pointer as the control's cursor
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
    Private Const IDC_HAND As Long = 32649

' Restrict user from selecting other controls while spacebar is held down
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long

' SelfSub APIs and declarations
Private Declare Function CallWindowProcA Lib "user32.dll" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Const IDX_SHUTDOWN  As Long = 1
    Private Const IDX_HWND      As Long = 2
    Private Const IDX_WNDPROC   As Long = 9
    Private Const IDX_BTABLE    As Long = 11
    Private Const IDX_ATABLE    As Long = 12
    Private Const IDX_PARM_USER As Long = 13
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32.dll" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32.dll" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Const GWL_WNDPROC   As Long = -4
    Private Const WNDPROC_OFF   As Long = &H38
Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long) ' Modified
    Private Const ALL_MESSAGES  As Long = -1
    Private Const MSG_ENTRIES   As Long = 32
    Private Enum eMsgWhen
        MSG_BEFORE = 1
        MSG_AFTER = 2
        MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
    End Enum
    Private z_ScMem  As Long
    Private z_Sc(64) As Long
    Private z_Funk   As Collection

' Events
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over the button."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
'Occurs when the user presses and then releases a mouse button over the button.
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user clicks over the button twice."
Attribute DblClick.VB_UserMemId = -601
'Occurs when the user clicks over the button twice.
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while the button has the focus."
Attribute KeyDown.VB_UserMemId = -602
'Occurs when the user presses a key while the button has the focus.
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
'Occurs when the user presses and releases an ANSI key.
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while the button has the focus."
Attribute KeyUp.VB_UserMemId = -604
'Occurs when the user releases a key while the button has the focus.
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while the button has the focus."
Attribute MouseDown.VB_UserMemId = -605
'Occurs when the user presses the mouse button while the button has the focus.
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occrus when the cursor moves around the button for the first time."
'Occrus when the cursor moves around the button for the first time.
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the cursor leaves/moves outside the button."
'Occurs when the cursor leaves/moves outside the button.
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the cursor moves over the button."
Attribute MouseMove.VB_UserMemId = -606
'Occurs when the cursor moves over the button.
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while the button has the focus."
Attribute MouseUp.VB_UserMemId = -607
'Occurs when the user releases the mouse button while the button has the focus.

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private KeyCode, Shift, KeyAscii, Button, x, Y
#End If

Public Enum eButtonShapes
    ebsCutNone  ' Normal
    ebsCutLeft  '
    ebsCutRight '
    ebsCutSides ' Both left & right
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private ebsCutNone, ebsCutLeft, ebsCutRight, ebsCutSides
#End If

Private Enum eButtonStates
    ebsNormal   ' Normal state
    ebsHot      ' Cursor is over the control
    ebsDown     ' Mouse/key is down
    ebsDisabled ' Disabled state
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private ebsNormal, ebsHot, ebsDown, ebsDisabled
#End If

Public Enum eButtonStyles
    #If USE_CRYSTAL Then
    ebsCrystal
    #End If
    #If USE_MAC Then
    ebsMac
    #End If
    #If USE_MACOSX Then
    ebsMacOSX
    #End If
    #If USE_OFFICE2003 Then
    ebsOffice2003
    #End If
    #If USE_OFFICEXP Then
    ebsOfficeXP
    #End If
    #If USE_OPERABROWSER Then
    ebsOperaBrowser
    #End If
    #If USE_STANDARD Then
    ebsStandard
    #End If
    #If USE_XPBLUE Then
    ebsXPBlue
    #End If
    #If USE_XPOLIVEGREEN Then
    ebsXPOliveGreen
    #End If
    #If USE_XPSILVER Then
    ebsXPSilver
    #End If
    #If USE_XPTOOLBAR Then
    ebsXPToolbar
    #End If
    #If USE_YAHOO Then
    ebsYahoo
    #End If
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private ebsCrystal, ebsMac, ebsMacOSX, ebsOffice2003
    Private ebsOfficeXP, ebsOperaBrowser, ebsStandard, ebsXPBlue
    Private ebsXPOliveGreen, ebsXPSilver, ebsXPToolbar, ebsYahoo
#End If

#If USE_POPUPMENU Then
Public Enum eMenuAlignments
    emaBottom       ' Bottom of control, aligned to the left (default)
    emaLeft         ' Left side of the control, aligned to the top
    emaLeftBottom   ' Left side of control, aligned to the bottom
    emaRight        ' Right side of control, aligned to the top
    emaRightBottom  ' Right side of control, aligned to the bottom
    emaTop          ' Top of control
    emaTopLeft      ' Top of control, aligned to the left
    emaTopRight     ' Top of control, aligned to the right
End Enum
#End If

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private emaBottom, emaLeft, emaLeftBottom, emaRight
    Private emaRightBottom, emaTop, emaTopLeft, emaTopRight
#End If

Public Enum ePictureAlignments
    epaBehindText
    epaBottomEdge
    epaBottomOfCaption
    epaLeftEdge
    epaLeftOfCaption
    epaRightEdge
    epaRightOfCaption
    epaTopEdge
    epaTopOfCaption
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private epaBehindText, epaBottomEdge, epaBottomOfCaption
    Private epaLeftEdge, epaLeftOfCaption, epaRightEdge
    Private epaRightOfCaption, epaTopEdge, epaTopOfCaption
#End If

Public Enum ePictureSizes
    epsNormal ' Use original size of main picture
    eps16x16  ' Small icon size
    eps24x24  ' Standard toolbar icon size
    eps32x32  ' Standard icon size
    eps48x48  ' Explorer thumbnail size
    epsCustom ' Use picture size defined in property
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private epsNormal, eps16x16, eps24x24, eps32x32, eps48x48, epsCustom
#End If

#If USE_SPECIALEFFECTS Then
Public Enum eSpecialEffects ' Icon/Caption special effects
    eseNone     ' Normal effect (none)
    eseEmbossed ' Raised effect
    eseEngraved ' Sunken effect
    eseShadowed ' Shadow effect
End Enum
#End If

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private eseNone, eseEmbossed, eseEngraved, eseShadowed
#End If

Public Enum eUserColors ' User-defined colors
    eucDownColor    ' ---
    eucFocusBorder  ' Available for use with the OverrideColor procedure
    eucGrayColor    ' For advance users only
    eucGrayText     ' Not accessible using VB IDE's properties window
    eucHoverColor   ' Note: Color property usage may vary between button styles
    eucStartColor   ' ---
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private eucDownColor, eucFocusBorder, eucGrayColor
    Private eucGrayText, eucHoverColor, eucStartColor
#End If

Private Type tButtonColors  ' Safe button colors (Translated)
    BackColor   As Long     ' Normal
    DownColor   As Long     ' Down state
    FocusBorder As Long     ' Focus state
    ForeColor   As Long     ' Normal text color
    GrayColor   As Long     ' Disabled background color
    GrayText    As Long     ' Disabled text and border color
    HoverColor  As Long     ' Hot state
    MaskColor   As Long     ' Mask color for the picture
    StartColor  As Long     ' Start color for gradient effect
End Type

Private Type tButtonProperties ' Cached button properties
    BackColor   As Long
    Caption     As String
    CheckBox    As Boolean
    #If USE_SPECIALEFFECTS Then
    Effects     As eSpecialEffects
    #End If
    Enabled     As Boolean
    ForeColor   As Long
    HandPointer As Boolean
    MaskColor   As Long
    PicAlign    As ePictureAlignments
    PicDown     As StdPicture
    PicHot      As StdPicture
    PicNormal   As StdPicture
    PicOpacity  As Single
    PicSize     As ePictureSizes
    PicSizeH    As Long
    PicSizeW    As Long
    Shape       As eButtonShapes
    Style       As eButtonStyles
    UseMask     As Boolean
    Value       As Boolean
End Type

Private Type tButtonSettings        ' Runtime/designtime settings
    Button      As Integer          ' Last button clicked
    Caption     As RECT             ' Area where to draw caption
    Cursor      As Long             ' Handle to the sytem hand pointer
    Default     As Boolean          ' Is control set as DEFAULT button of a form?
    Focus       As RECT             ' Area where to draw FocusRect/ Shine object
    HasFocus    As Boolean          ' Is the control currently in focus
    Height      As Long             ' ScaleHeight (pixels)
    Picture     As RECT             ' Area where to draw icon/picture
    state       As eButtonStates    ' Current drawing state of the button
    Width       As Long             ' ScaleWidth (pixels)
End Type

#If USE_POPUPMENU Then
Private Type tPopupSettings
    Menu        As VB.Menu
    Align       As eMenuAlignments
    flags       As Long
    DefaultMenu As VB.Menu
End Type
#End If

Private Type tRGB
    r As Long
    G As Long
    b As Long
End Type

' Variables
Private m_bButtonHasFocus   As Boolean
Private m_bButtonIsDown     As Boolean
Private m_bCalculateRects   As Boolean
Private m_bControlHidden    As Boolean
Private m_bIsTracking       As Boolean
Private m_bIsPlatformNT     As Boolean
Private m_bMouseIsDown      As Boolean
Private m_bMouseOnButton    As Boolean
Private m_bParentActive     As Boolean
Private m_bRedrawOnResize   As Boolean
Private m_bSpacebarIsDown   As Boolean
Private m_bTrackHandler32   As Boolean

#If USE_POPUPMENU Then
Private m_bPopupInit        As Boolean
Private m_bPopupEnabled     As Boolean
Private m_bPopupShown       As Boolean
Private m_tPopupSettings    As tPopupSettings
#End If

#If USE_CRYSTAL Or USE_MAC Or USE_MACOSX Then
Private m_lpBordrPoints()   As POINTAPI ' Main border
#End If

Private m_tButtonProperty   As tButtonProperties
Private m_tButtonColors     As tButtonColors
Private m_tButtonSettings   As tButtonSettings

' //-- Properties --//

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used for the button."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
'   Returns/sets the background color used for the button.
    BackColor = m_tButtonProperty.BackColor
    
End Property

Public Property Let BackColor(Value As OLE_COLOR)
    m_tButtonProperty.BackColor = Value
    SetButtonColors ' Update color changes
    DrawButton Force:=True
    PropertyChanged "BackColor"
    
End Property

Public Property Get ButtonShape() As eButtonShapes
Attribute ButtonShape.VB_Description = "Returns/sets a value to determine the shape used to draw the button."
'   Returns/sets a value to determine the shape used to draw the button.
    ButtonShape = m_tButtonProperty.Shape
    
End Property

Public Property Let ButtonShape(Value As eButtonShapes)
    m_tButtonProperty.Shape = Value
    Me.Refresh
    PropertyChanged "ButtonShape"
    
End Property

Public Property Get ButtonStyle() As eButtonStyles
Attribute ButtonStyle.VB_Description = "Returns/sets a value to determine the style used to draw the button."
Attribute ButtonStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value to determine the style used to draw the button.
    ButtonStyle = m_tButtonProperty.Style
    
End Property

Public Property Let ButtonStyle(Value As eButtonStyles)
    m_tButtonProperty.Style = Value
    
    If (Not Ambient.UserMode) Then ' On IDE
        ColorScheme NoRedraw:=True ' Reset and set default colors
    Else
        SetButtonColors ' Set button colors using current color scheme
    End If
    
    Me.Refresh
    PropertyChanged "ButtonStyle"
    
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the button."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
'   Returns/sets the text displayed in the button.
    Caption = m_tButtonProperty.Caption
    
End Property

Public Property Let Caption(Value As String)
    m_tButtonProperty.Caption = Value
    UserControl.AccessKeys = GetAccessKey(Value)
    Me.Refresh
    PropertyChanged "Caption"
    
End Property

Public Property Get CheckBoxMode() As Boolean
Attribute CheckBoxMode.VB_Description = "Returns/sets the type of control the button will observe."
Attribute CheckBoxMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
'   Returns/sets the type of control the button will observe.
    CheckBoxMode = m_tButtonProperty.CheckBox
    
End Property

Public Property Let CheckBoxMode(Value As Boolean)
    m_tButtonProperty.CheckBox = Value
    
    If (Not Value) And (m_tButtonProperty.Value) Then
        m_tButtonProperty.Value = False ' Normal state
        PropertyChanged "Value"
    End If
    
    DrawButton Force:=True
    PropertyChanged "CheckBox"
    
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value to determine whether the button can respond to events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
'   Returns/sets a value to determine whether the button can respond to events.
    Enabled = m_tButtonProperty.Enabled
    
End Property

Public Property Let Enabled(Value As Boolean)
    m_tButtonProperty.Enabled = Value
    UserControl.Enabled = Value
    
    If (Not Value) Then ' Disabled
        DrawButton ebsDisabled
    ElseIf (m_bMouseOnButton) Then
        DrawButton ebsHot
    Else
        DrawButton ebsNormal
    End If
    
    PropertyChanged "Enabled"
    
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/sets the Font used to display text on the button."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
'   Returns/sets the Font used to display text on the button.
    Set Font = UserControl.Font
    
End Property

Public Property Set Font(Value As StdFont)
    Set UserControl.Font = Value
    Me.Refresh
    PropertyChanged "Font"
    
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font style."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontBold.VB_MemberFlags = "400"
'   Returns/sets bold font style.
    FontBold = UserControl.FontBold
    
End Property

Public Property Let FontBold(Value As Boolean)
    UserControl.FontBold = Value
    Me.Refresh
    
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font style."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontItalic.VB_MemberFlags = "400"
'   Returns/sets italic font style.
    FontItalic = UserControl.FontItalic
    
End Property

Public Property Let FontItalic(Value As Boolean)
    UserControl.FontItalic = Value
    Me.Refresh
    
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font used for the button caption."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontName.VB_MemberFlags = "400"
'   Specifies the name of the font used for the button caption.
    FontName = UserControl.FontName
    
End Property

Public Property Let FontName(Value As String)
    UserControl.FontName = Value
    Me.Refresh
    
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font used for the button caption."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontSize.VB_MemberFlags = "400"
'   Specifies the size (in points) of the font used for the button caption.
    FontSize = UserControl.FontSize
    
End Property

Public Property Let FontSize(Value As Single)
    UserControl.FontSize = Value
    Me.Refresh
    
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font style."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontStrikethru.VB_MemberFlags = "400"
'   Returns/sets strikethrough font style.
    FontStrikethru = UserControl.FontStrikethru
    
End Property

Public Property Let FontStrikethru(Value As Boolean)
    UserControl.FontStrikethru = Value
    Me.Refresh
    
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font style."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontUnderline.VB_MemberFlags = "400"
'   Returns/sets underline font style.
    FontUnderline = UserControl.FontUnderline
    
End Property

Public Property Let FontUnderline(Value As Boolean)
    UserControl.FontUnderline = Value
    Me.Refresh
    
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the text color of the button caption."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
'   Returns/sets the text color of the button caption.
    ForeColor = m_tButtonProperty.ForeColor
    
End Property

Public Property Let ForeColor(Value As OLE_COLOR)
    m_tButtonProperty.ForeColor = Value
    
    If (m_tButtonProperty.Enabled) Then ' Disabled text uses its own color
        SetButtonColors ' Update color changes
        DrawButton Force:=True
    End If
    
    PropertyChanged "ForeColor"
    
End Property

Public Property Get HandPointer() As Boolean
Attribute HandPointer.VB_Description = "Returns/sets a value to determine whether the control uses the system's hand pointer as its cursor."
Attribute HandPointer.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Returns/sets a value to determine whether the control uses the system's hand pointer as its cursor.
    HandPointer = m_tButtonProperty.HandPointer
    
End Property

Public Property Let HandPointer(Value As Boolean)
    m_tButtonProperty.HandPointer = Value
    PropertyChanged "HandPointer"
    
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle that uniquely identifies the control."
Attribute hWnd.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
'   Returns a handle that uniquely identifies the control.
    hWnd = UserControl.hWnd
    
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in a button's picture to be transparent."
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a color in a button's picture to be transparent.
    MaskColor = m_tButtonProperty.MaskColor
    
End Property

Public Property Let MaskColor(Value As OLE_COLOR)
    m_tButtonProperty.MaskColor = Value
    m_tButtonColors.MaskColor = TranslateColor(Value)
    DrawButton Force:=True
    PropertyChanged "MaskColor"
    
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon for the button."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Sets a custom mouse icon for the button.
    Set MouseIcon = UserControl.MouseIcon
    
End Property

Public Property Set MouseIcon(Value As IPictureDisp)
    Set UserControl.MouseIcon = Value       ' Set new cursor
                                            '
    If (Value Is Nothing) Then              '
        Me.MousePointer = 0 ' vbDefault     ' Apply appropriate
    Else                                    ' mouse pointer setting
        Me.MousePointer = 99 ' vbCustom     ' automatically
    End If                                  '
    
    PropertyChanged "MouseIcon"
    
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when cursor over the button."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Returns/sets the type of mouse pointer displayed when cursor over the button.
    MousePointer = UserControl.MousePointer
    
End Property

Public Property Let MousePointer(Value As MousePointerConstants)
    If (Not Value = vbCustom) And (m_tButtonProperty.HandPointer) Then
        Me.HandPointer = False          ' Unset hand pointer option
    End If                              ' on demand
                                        '
    UserControl.MousePointer = Value    ' Set new mouse pointer
    PropertyChanged "MousePointer"
    
End Property

Public Property Get PictureAlignment() As ePictureAlignments
Attribute PictureAlignment.VB_Description = "Returns/sets a value to determine where to draw the picture in the button."
Attribute PictureAlignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value to determine where to draw the picture in the button.
    PictureAlignment = m_tButtonProperty.PicAlign
    
End Property

Public Property Let PictureAlignment(Value As ePictureAlignments)
    m_tButtonProperty.PicAlign = Value
    Me.Refresh
    PropertyChanged "PicAlign"
    
End Property

Public Property Get PictureDown() As StdPicture
Attribute PictureDown.VB_Description = "Returns/sets the picture displayed when the control is pressed down or in checked state."
Attribute PictureDown.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the picture displayed when the control is pressed down or in checked state.
    Set PictureDown = m_tButtonProperty.PicDown
    
End Property

Public Property Set PictureDown(Value As StdPicture)
'   Main picture (PictureNormal) must be set first before this property (PictureDown)
'   Else the specified value will be set as the main picture instead.
    
    If (Value Is Nothing) Then
        Set m_tButtonProperty.PicDown = Nothing
        GoTo Jmp_Skip
    End If
    
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        Set Me.PictureNormal = Value
        Exit Property
    Else
        If (m_tButtonProperty.PicSize = epsNormal) Then
            If (Not m_tButtonProperty.PicNormal.Width = Value.Width) Or _
               (Not m_tButtonProperty.PicNormal.Height = Value.Height) Then
                ' If pictures do not have the same sizes (width or height)
                ' Use main picture's size as the standard for all pictures
                Me.PictureSize = epsCustom
            End If
        End If
        Set m_tButtonProperty.PicDown = Value
    End If
    
Jmp_Skip:
    If (m_bButtonIsDown) Then
        DrawButton Force:=True
    End If
    PropertyChanged "PicDown"
    
End Property

Public Property Get PictureHot() As StdPicture
Attribute PictureHot.VB_Description = "Returns/sets the picture displayed when the cursor is over the control."
Attribute PictureHot.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the picture displayed when the cursor is over the control.
    Set PictureHot = m_tButtonProperty.PicHot
    
End Property

Public Property Set PictureHot(Value As StdPicture)
'   Main picture (PictureNormal) must be set first before this property (PictureHot)
'   Else the specified value will be set as the main picture instead.
    
    If (Value Is Nothing) Then
        Set m_tButtonProperty.PicHot = Nothing
        GoTo Jmp_Skip
    End If
    
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        Set Me.PictureNormal = Value
        Exit Property
    Else
        If (m_tButtonProperty.PicSize = epsNormal) Then
            If (Not m_tButtonProperty.PicNormal.Width = Value.Width) Or _
               (Not m_tButtonProperty.PicNormal.Height = Value.Height) Then
                ' If pictures do not have the same sizes (width or height)
                ' Use main picture's size as the standard for all pictures
                Me.PictureSize = epsCustom
            End If
        End If
        Set m_tButtonProperty.PicHot = Value
    End If
    
Jmp_Skip:
    If (m_bMouseOnButton) Then
        DrawButton Force:=True
    End If
    PropertyChanged "PicHot"
    
End Property

Public Property Get PictureNormal() As StdPicture
Attribute PictureNormal.VB_Description = "Returns/sets the picture displayed on a normal state button."
Attribute PictureNormal.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the picture displayed on a normal state button.
    Set PictureNormal = m_tButtonProperty.PicNormal
    
End Property

Public Property Set PictureNormal(Value As StdPicture)
    
    If (Value Is Nothing) Then
        ' Cannot work without the main picture
        Set Me.PictureDown = Nothing
        Set Me.PictureHot = Nothing
            m_tButtonProperty.PicOpacity = 0
    ElseIf (m_tButtonProperty.PicOpacity = 0) Then
        m_tButtonProperty.PicOpacity = 1
    End If
    
    Set m_tButtonProperty.PicNormal = Value
        Me.PictureSize = m_tButtonProperty.PicSize ' Update picture sizes
        
    Me.Refresh
    PropertyChanged "PicNormal"
    PropertyChanged "PicOpacity"
    
End Property

Public Property Get PictureOpacity() As Long
Attribute PictureOpacity.VB_Description = "Returns/sets a value in percent how the pictures will be blended to the button."
Attribute PictureOpacity.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value in percent how the pictures will be blended to the button.
    PictureOpacity = m_tButtonProperty.PicOpacity * 100
    
End Property

Public Property Let PictureOpacity(Value As Long)
'   Below 10% means the picture will not be visible at all (or almost)
'   So why blend it this way if you could just remove the picture instead?
    m_tButtonProperty.PicOpacity = TranslateNumber(Value, 10, 100)
    m_tButtonProperty.PicOpacity = m_tButtonProperty.PicOpacity / 100
    DrawButton Force:=True
    PropertyChanged "PicOpacity"
    
End Property

Public Property Get PictureSize() As ePictureSizes
Attribute PictureSize.VB_Description = "Returns/sets a value to determine the size of the picture to draw."
Attribute PictureSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value to determine the size of the picture to draw.
    PictureSize = m_tButtonProperty.PicSize
    
End Property

Public Property Let PictureSize(Value As ePictureSizes)
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        m_tButtonProperty.PicSize = epsNormal
        m_tButtonProperty.PicSizeH = 0
        m_tButtonProperty.PicSizeW = 0
        GoTo Jmp_Skip
    End If
    
    m_tButtonProperty.PicSize = Value
    
    With m_tButtonProperty
    Select Case Value
        Case epsNormal
            .PicSizeH = ScaleY(m_tButtonProperty.PicNormal.Height, 8, 3) ' 8 = vbHimetric; 3 = vbPixels
            .PicSizeW = ScaleY(m_tButtonProperty.PicNormal.Width, 8, 3)
        Case eps16x16
            .PicSizeH = 16
            .PicSizeW = 16
        Case eps24x24
            .PicSizeH = 24
            .PicSizeW = 24
        Case eps32x32
            .PicSizeH = 32
            .PicSizeW = 32
        Case eps48x48
            .PicSizeH = 48
            .PicSizeW = 48
    End Select
    End With
    
Jmp_Skip:
    Me.Refresh
    
    PropertyChanged "PicSize"
    PropertyChanged "PicSizeH"
    PropertyChanged "PicSizeW"
    
End Property

Public Property Get PictureSizeH() As Long
Attribute PictureSizeH.VB_Description = "Returns/sets the standard/custom height of the defined pictures in pixels."
Attribute PictureSizeH.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the standard/custom height of the defined pictures in pixels.
    PictureSizeH = m_tButtonProperty.PicSizeH
    
End Property

Public Property Let PictureSizeH(Value As Long)
    m_tButtonProperty.PicSize = epsCustom   ' If modified then set size to custom
    m_tButtonProperty.PicSizeH = Value      '
    Me.Refresh                              '
    PropertyChanged "PicSize"               ' Notify the container that
    PropertyChanged "PicSizeH"              ' properties has been changed
    
End Property

Public Property Get PictureSizeW() As Long
Attribute PictureSizeW.VB_Description = "Returns/sets the standard/custom width of the defined pictures in pixels."
Attribute PictureSizeW.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the standard/custom width of the defined pictures in pixels.
    PictureSizeW = m_tButtonProperty.PicSizeW
    
End Property

Public Property Let PictureSizeW(Value As Long)
    m_tButtonProperty.PicSize = epsCustom   ' If modified then set size to custom
    m_tButtonProperty.PicSizeW = Value      '
    Me.Refresh                              '
    PropertyChanged "PicSize"               ' Notify the container that
    PropertyChanged "PicSizeW"              ' properties has been changed
    
End Property

#If USE_SPECIALEFFECTS Then
Public Property Get SpecialEffect() As eSpecialEffects
Attribute SpecialEffect.VB_Description = "Returns/sets a value to determine the effect applied to button caption and icon."
Attribute SpecialEffect.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value to determine the effect applied to button caption and icon.
    SpecialEffect = m_tButtonProperty.Effects
    
End Property

Public Property Let SpecialEffect(Value As eSpecialEffects)
    m_tButtonProperty.Effects = Value
    DrawButton Force:=True
    PropertyChanged "Effects"
    
End Property
#End If

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value to determine whether to use MaskColor to create transparent areas of the picture."
Attribute UseMaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value to determine whether to use MaskColor to create transparent areas of the picture.
    UseMaskColor = m_tButtonProperty.UseMask
    
End Property

Public Property Let UseMaskColor(Value As Boolean)
    m_tButtonProperty.UseMask = Value
    DrawButton Force:=True
    PropertyChanged "UseMask"
    
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the value or state of the button."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Returns/sets the value or state of the button.
    Value = m_tButtonProperty.Value
    
End Property

Public Property Let Value(Value As Boolean)
    m_tButtonProperty.Value = Value
    
    If (Value) And (Not m_tButtonProperty.CheckBox) Then
        If (Ambient.UserMode) Then
            m_tButtonSettings.Button = 1
            UserControl_Click ' Trigger click event
        Else
            m_tButtonProperty.Value = False
        End If
    Else ' Value is False or CheckBoxMode is True
        DrawButton Force:=True
    End If
    
    PropertyChanged "Value"
    
End Property

' //-- Public Procedures --//

Public Sub About()
Attribute About.VB_Description = "Shows information about the control and its author."
Attribute About.VB_UserMemId = -552
'   Shows information about the control and its author.
    ShellExecute UserControl.hWnd, "open", DC_URL, "", "", SW_SHOWNORMAL
    
End Sub

Public Sub ColorScheme( _
        Optional Style As eButtonStyles = -1, _
        Optional NoRedraw As Boolean)
Attribute ColorScheme.VB_Description = "Set BackColor to the default background color of a button style."
'   Set BackColor to the default background color of a button style.
    
    If (Style = -1) Then
        Style = m_tButtonProperty.Style ' Use current button style
    End If
    
    Select Case Style ' Only the background color will be set
        #If USE_CRYSTAL Then
        Case ebsCrystal
            m_tButtonProperty.BackColor = &HBA9EA0
        #End If
        #If USE_MAC Then
        Case ebsMac
            m_tButtonProperty.BackColor = &HFF9B48
        #End If
        #If USE_MACOSX Then
        Case ebsMacOSX
            m_tButtonProperty.BackColor = &HA19D9D
        #End If
        #If USE_OFFICE2003 Then
        Case ebsOffice2003
            m_tButtonProperty.BackColor = &HBA9EA0
        #End If
        #If USE_OFFICEXP Then
        Case ebsOfficeXP
            m_tButtonProperty.BackColor = &H8000000F ' vbButtonFace
        #End If
        #If USE_OPERABROWSER Then
        Case ebsOperaBrowser
            m_tButtonProperty.BackColor = &HD2CECF
        #End If
        #If USE_STANDARD Then
        Case ebsStandard
            m_tButtonProperty.BackColor = &H8000000F ' vbButtonFace
        #End If
        #If USE_XPBLUE Then
        Case ebsXPBlue
            m_tButtonProperty.BackColor = &HE6EBEC
        #End If
        #If USE_XPOLIVEGREEN Then
        Case ebsXPOliveGreen
            m_tButtonProperty.BackColor = &HE0F3F6
        #End If
        #If USE_XPSILVER Then
        Case ebsXPSilver
            m_tButtonProperty.BackColor = &HE4D1D2
        #End If
        #If USE_XPTOOLBAR Then
        Case ebsXPToolbar
            m_tButtonProperty.BackColor = &H8000000F ' vbButtonFace
        #End If
        #If USE_YAHOO Then
        Case ebsYahoo
            m_tButtonProperty.BackColor = &H12BCFF
        #End If
    End Select
    
    SetButtonColors
    
    If (Not NoRedraw) Then
        DrawButton Force:=True ' Apply new colors cheme
    End If
    
End Sub

Public Sub OverrideColor( _
        Property As eUserColors, _
        ByVal Color As Long, _
        Optional NoRedraw As Boolean)
Attribute OverrideColor.VB_Description = "Override the predefined color property of the button."
'   Override the predefined color property of the button. (For advance users only)
    
    Color = TranslateColor(Color)
    
    Select Case Property
        Case eucDownColor
            m_tButtonColors.DownColor = Color
        Case eucFocusBorder
            m_tButtonColors.FocusBorder = Color
        Case eucGrayColor
            m_tButtonColors.GrayColor = Color
        Case eucGrayText
            m_tButtonColors.GrayText = Color
        Case eucHoverColor
            m_tButtonColors.HoverColor = Color
        Case eucStartColor
            m_tButtonColors.StartColor = Color
        Case Else
            Err.Raise 5 ' Invalid procedure call or argument
            Exit Sub
    End Select
    
    If (Not NoRedraw) Then
        DrawButton Force:=True
    End If
    
End Sub

#If USE_POPUPMENU Then
Public Function SetPopupMenu( _
        Menu As Object, _
        Optional Align As eMenuAlignments, _
        Optional flags, _
        Optional DefaultMenu)
Attribute SetPopupMenu.VB_Description = "Set the control to handle popup operation to the specified menu and settings."
'   Set the control to handle popup operation to the specified menu and settings.
'   When set, the control does not send the regular mouse and keyboard events.
    
    If (Not Menu Is Nothing) Then
        If (TypeOf Menu Is VB.Menu) Then
            
            With m_tPopupSettings
                Set .Menu = Menu
                    .Align = Align
                    
                If (IsMissing(flags)) Then
                    .flags = 0
                Else
                    .flags = flags
                End If
                
                If (IsMissing(DefaultMenu)) Then
                    Set .DefaultMenu = Nothing
                Else
                    Set .DefaultMenu = DefaultMenu
                End If
            End With
            
            m_bPopupEnabled = True
            
        End If
    End If
    
End Function
#End If

#If USE_POPUPMENU Then
Public Sub UnsetPopupMenu()
Attribute UnsetPopupMenu.VB_Description = "Unset the control handling popup operation and restores regular button events."
'   Unset the control handling popup operation and restores regular button events.
    With m_tPopupSettings
        Set .Menu = Nothing
            .Align = 0
            .flags = 0
        Set .DefaultMenu = Nothing
    End With
    
    m_bPopupEnabled = False
    
End Sub
#End If

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of the button."
Attribute Refresh.VB_UserMemId = -550
'   Forces a complete repaint of the button.
    m_bCalculateRects = True
    DrawButton Force:=True
    
End Sub

' //-- Private Procedures --//

Private Function BlendColors( _
        Color1 As Long, _
        Color2 As Long, _
        Optional PercentInDecimal As Single = 0.5) As Long
'   Combines two colors together by how many percent.
    
    Dim Color1RGB As tRGB
    Dim Color2RGB As tRGB
    Dim Color3RGB As tRGB
    
    Color1RGB = GetRGB(Color1)
    Color2RGB = GetRGB(Color2)
    
    Color3RGB.r = Color1RGB.r + (Color2RGB.r - Color1RGB.r) * PercentInDecimal ' Percent should already
    Color3RGB.G = Color1RGB.G + (Color2RGB.G - Color1RGB.G) * PercentInDecimal ' be translated to decimal.
    Color3RGB.b = Color1RGB.b + (Color2RGB.b - Color1RGB.b) * PercentInDecimal ' Ex. 50% -> 50 / 100 = 0.5
    
    BlendColors = RGBEx(Color3RGB.r, Color3RGB.G, Color3RGB.b)
    
End Function

Private Function BlendRGBQUAD( _
        RGB1 As RGBQUAD, _
        RGB2 As RGBQUAD, _
        Optional PercentInDecimal As Single = 0.5) As RGBQUAD
'   Blend two colors particularly RGBQUAD structures together by how many percent.
    
    Dim RGB3 As tRGB
        RGB3.r = RGB2.rgbRed            ' An overflow will cause to occur
        RGB3.G = RGB2.rgbGreen          ' if we directly subtract RGBQUAD
        RGB3.b = RGB2.rgbBlue           ' values from each other.
                                        '
        RGB3.r = RGB3.r - RGB1.rgbRed   ' So instead, we store the first RGBQUAD
        RGB3.G = RGB3.G - RGB1.rgbGreen ' structure as the result then after
        RGB3.b = RGB3.b - RGB1.rgbBlue  ' we can safely subtract it the other.
        
        RGB3.r = RGB1.rgbRed + RGB3.r * PercentInDecimal    ' Percent should already
        RGB3.G = RGB1.rgbGreen + RGB3.G * PercentInDecimal  ' be translated to decimal.
        RGB3.b = RGB1.rgbBlue + RGB3.b * PercentInDecimal   ' Ex. 50% -> 50 / 100 = 0.5
        
        BlendRGBQUAD.rgbRed = RGB3.r
        BlendRGBQUAD.rgbGreen = RGB3.G
        BlendRGBQUAD.rgbBlue = RGB3.b
        
End Function

Private Sub CalculateRects( _
        Optional ByVal BorderH As Long = 4, _
        Optional ByVal BorderW As Long = 4)
'   Calculate areas where to draw the caption and the icon/picture
'   Uses the integer divisions because it does not round off the result when calculating,
'   thus making the result more accurate when performing specific calulations.
'   Although in theory, integer division is slower than the regular division, don't worry,
'   this procedure is called only when the control has encountered a decisive change.
    
    m_bCalculateRects = False
    
    Dim bh As Long          ' Height of button
    Dim bw As Long          ' Width of button
                            '
    Dim pr As RECT          ' Draw area of picture
                            '
    Dim S1 As Single        ' Temporary variables
    Dim S2 As Single        '
                            '
    Dim tn As Long          ' Length of text
    Dim tr As RECT          ' Draw area of caption
    Dim tx As String        ' Caption text
                            '
    With m_tButtonSettings  '
        bh = .Height        ' Get button size
        bw = .Width         '
                            '
        tx = m_tButtonProperty.Caption
        tn = Len(tx)        '
                            '
        SetRect .Focus, BorderW, BorderH, .Width - BorderW, .Height - BorderH
    End With                '
                            '
    If (tn > 0) Then        ' Set estimated drawing area of caption
        SetRect tr, 0, 0, bw, bh                    '
        If (m_bIsPlatformNT) Then                   '
            DrawTextW hdc, StrPtr(tx), tn, tr, DT_CALCFLAG
        Else                                        '
            DrawText hdc, tx, tn, tr, DT_CALCFLAG   ' Get width & height of area
        End If                                      ' that fits the text/caption
    End If                                          '
                                                    ' Move caption area to the center
    OffsetRect tr, (bw - tr.Right) \ 2, (bh - tr.bottom) \ 2
                                                    '
    CopyRect m_tButtonSettings.Caption, tr          ' Save changes
                                                    '
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        Exit Sub                                    ' Skip icon alignment when
    End If                                          ' no picture is to be aligned
    
    SetRect pr, 0, 0, m_tButtonProperty.PicSizeW, m_tButtonProperty.PicSizeH
    
    If (tn > 0) Then    ' Check if a caption is specified
        tn = 2          ' If set, then tn will be set to contain the value in
    Else                ' pixels the caption and the picture will be apart
        tn = 0          ' Else, just set to zero to retain the icon/picture
    End If              ' in place (center) when no caption is specified :)
    
    ' Note: OffsetRect moves the RECT coordinates by how many pixels
    '       from its current position. (it does not change its size)
    
    Select Case m_tButtonProperty.PicAlign                          ' Picture on center
        Case epaBehindText                                          ' but behind text
            OffsetRect pr, (bw - pr.Right) \ 2, (bh - pr.bottom) \ 2
                                                                    '
        Case epaBottomEdge, epaBottomOfCaption                      ' Picture on
            OffsetRect pr, (bw - pr.Right) \ 2, 0                   ' bottom portion
            OffsetRect tr, 0, -tr.Top                               ' of caption
                                                                    '
            If (m_tButtonProperty.PicAlign = epaBottomEdge) Then    '
                OffsetRect pr, 0, bh - pr.bottom - BorderH          '
                If (tn > 0) Then                                    '
                    OffsetRect tr, 0, pr.Top - tr.bottom '- tn      '
                    tn = tr.Top - BorderH                           '
                    If (tn > 1) Then                                '
                        OffsetRect tr, 0, -(tn \ 2)                 '
                    End If                                          '
                End If                                              '
            ElseIf (tn = 0) Then                                    '
                OffsetRect pr, 0, (bh - pr.bottom) \ 2              '
            Else                                                    '
                OffsetRect pr, 0, tr.bottom '+ tn                   '
                tn = (bh - pr.bottom) \ 2                           '
                OffsetRect pr, 0, tn                                '
                OffsetRect tr, 0, tn                                '
            End If                                                  '
                                                                    '
        Case epaLeftEdge, epaLeftOfCaption                          ' Picture on
            OffsetRect pr, 0, (bh - pr.bottom) \ 2                  ' left portion
            OffsetRect tr, -tr.Left, 0                              ' of caption
                                                                    '
            If (m_tButtonProperty.PicAlign = epaLeftEdge) Then      '
                OffsetRect pr, BorderW, 0                           '
                If (tn > 0) Then                                    '
                    OffsetRect tr, pr.Right + tn, 0                 '
                    tn = bw - BorderW - tr.Right                    '
                    If (tn > 1) Then                                '
                        OffsetRect tr, tn \ 2, 0                    '
                    End If                                          '
                End If                                              '
            ElseIf (tn = 0) Then                                    '
                OffsetRect pr, (bw - pr.Right) \ 2, 0               '
            Else                                                    '
                OffsetRect tr, pr.Right + tn, 0                     '
                tn = (bw - tr.Right) \ 2                            '
                OffsetRect tr, tn, 0                                '
                OffsetRect pr, tn, 0                                '
            End If                                                  '
                                                                    '
        Case epaRightEdge, epaRightOfCaption                        ' Picture on
            OffsetRect pr, 0, (bh - pr.bottom) \ 2                  ' right portion
            OffsetRect tr, -tr.Left, 0                              ' of caption
                                                                    '
            If (m_tButtonProperty.PicAlign = epaRightEdge) Then     '
                OffsetRect pr, bw - pr.Right - BorderW, 0           '
                If (tn > 0) Then                                    '
                    OffsetRect tr, pr.Left - tr.Right - tn, 0       '
                    tn = tr.Left - BorderW                          '
                    If (tn > 1) Then                                '
                        OffsetRect tr, -(tn \ 2), 0                 '
                    End If                                          '
                End If                                              '
            ElseIf (tn = 0) Then                                    '
                OffsetRect pr, (bw - pr.Right) \ 2, 0               '
            Else                                                    '
                OffsetRect pr, tr.Right + tn, 0                     '
                tn = (bw - pr.Right) \ 2                            '
                OffsetRect pr, tn, 0                                '
                OffsetRect tr, tn, 0                                '
            End If                                                  '
                                                                    '
        Case epaTopEdge, epaTopOfCaption                            ' Picture on
            OffsetRect pr, (bw - pr.Right) \ 2, 0                   ' top portion
            OffsetRect tr, 0, -tr.Top                               ' of caption
                                                                    '
            If (m_tButtonProperty.PicAlign = epaTopEdge) Then       '
                OffsetRect pr, 0, BorderH                           '
                If (tn > 0) Then                                    '
                    OffsetRect tr, 0, pr.bottom '+ tn               '
                    tn = bh - tr.bottom - BorderH                   '
                    If (tn > 1) Then                                '
                        OffsetRect tr, 0, tn \ 2                    '
                    End If                                          '
                End If                                              '
            ElseIf (tn = 0) Then                                    '
                OffsetRect pr, 0, (bh - pr.bottom) \ 2              '
            Else                                                    '
                OffsetRect tr, 0, pr.bottom '+ tn                   '
                tn = (bh - tr.bottom) \ 2                           '
                OffsetRect tr, 0, tn                                '
                OffsetRect pr, 0, tn                                '
            End If                                                  '
            
    End Select
    
    CopyRect m_tButtonSettings.Picture, pr  ' Save changes to drawing areas
    CopyRect m_tButtonSettings.Caption, tr  '
    
End Sub

#If USE_CRYSTAL Or USE_MAC Or USE_MACOSX Then
Private Sub CalculateRegionBorder( _
        Region As Long, _
        ByVal EllipseW As Long, _
        ByVal EllipseH As Long)
'   Calculate points needed to draw border of round-rectangular regions.
    
    ' We need to manually determine the points needed to draw the
    ' borders of a button with rounded corners properly because
    ' RoundRect API seems to be complicated for non-rectangular regions.
    
    ' The following code is a bit longer because all the four corners
    ' are scanned for points in separate loops. This is the only way
    ' to do the scanning fast and accurate because the points must be
    ' in a right order inorder for the resulting polygon to take shape.
    
    Dim aBordr() As POINTAPI
    Dim iBordr As Long
    
    Dim H As Long
    Dim W As Long
    Dim x As Long
    Dim Y As Long
        H = m_tButtonSettings.Height
        W = m_tButtonSettings.Width
        
        ' Corner's width/height is half of the specified ellipse size
        EllipseH = EllipseH / 2
        EllipseW = EllipseW / 2
        
        ' but does not exceed to one half of the button's size
        EllipseH = TranslateNumber(EllipseH, 0, H / 2)
        EllipseW = TranslateNumber(EllipseW, 0, W / 2)
        
    If (m_tButtonProperty.Shape = ebsCutLeft) Or _
       (m_tButtonProperty.Shape = ebsCutSides) Then
        
        ReDim Preserve aBordr(iBordr + 2)
        aBordr(iBordr).x = 0
        aBordr(iBordr).Y = 0
        iBordr = iBordr + 1
        aBordr(iBordr).x = 0
        aBordr(iBordr).Y = H - 1
        iBordr = iBordr + 1
    Else
        ' Note: While loop is faster than any other VB loop procedure
        ' Determine points on top-left corner
        While (Y < EllipseH)
            x = EllipseW
            While (x >= -1)
                If (PtInRegion(Region, x, Y) = 0) Then
                    ReDim Preserve aBordr(iBordr + 1)
                    aBordr(iBordr).x = x + 1
                    aBordr(iBordr).Y = Y
                    iBordr = iBordr + 1
                    If (x > 0) Then
                        aBordr(iBordr).x = x
                        aBordr(iBordr).Y = Y + 1
                        iBordr = iBordr + 1
                    End If
                    x = -1 ' Exit while
                End If
                x = x - 1
            Wend
            Y = Y + 1
        Wend
        
        Y = H - EllipseH
        
        ' Determine points on bottom-left corner
        While (Y < H)
            x = EllipseW
            While (x >= -1)
                If (PtInRegion(Region, x, Y) = 0) Then
                    ReDim Preserve aBordr(iBordr + 1)
                    If (x > 0) Then
                        aBordr(iBordr).x = x
                        aBordr(iBordr).Y = Y - 1
                        iBordr = iBordr + 1
                    End If
                    aBordr(iBordr).x = x + 1
                    aBordr(iBordr).Y = Y
                    iBordr = iBordr + 1
                    x = -1 ' Exit while
                End If
                x = x - 1
            Wend
            Y = Y + 1
        Wend
    End If
    
    If (m_tButtonProperty.Shape = ebsCutRight) Or _
       (m_tButtonProperty.Shape = ebsCutSides) Then
        
        ReDim Preserve aBordr(iBordr + 2)
        aBordr(iBordr).x = W - 1
        aBordr(iBordr).Y = H - 1
        iBordr = iBordr + 1
        aBordr(iBordr).x = W - 1
        aBordr(iBordr).Y = 0
        iBordr = iBordr + 1
    Else
        Y = H - 1
        
        ' Determine points on bottom-right corner
        While (Y > H - EllipseH)
            x = W - EllipseW
            While (x < W)
                If (PtInRegion(Region, x, Y) = 0) Then
                    ReDim Preserve aBordr(iBordr + 1)
                    aBordr(iBordr).x = x - 1
                    aBordr(iBordr).Y = Y
                    iBordr = iBordr + 1
                    aBordr(iBordr).x = x
                    aBordr(iBordr).Y = Y - 1
                    iBordr = iBordr + 1
                    x = W ' Exit while
                End If
                x = x + 1
            Wend
            Y = Y - 1
        Wend
        
        Y = EllipseH
        
        ' Determine points on top-right corner
        While (Y >= 0)
            x = W - EllipseW
            While (x < W)
                If (PtInRegion(Region, x, Y) = 0) Then
                    ReDim Preserve aBordr(iBordr + 1)
                    aBordr(iBordr).x = x
                    aBordr(iBordr).Y = Y + 1
                    iBordr = iBordr + 1
                    aBordr(iBordr).x = x - 1
                    aBordr(iBordr).Y = Y
                    iBordr = iBordr + 1
                    x = W ' Exit while
                End If
                x = x + 1
            Wend
            Y = Y - 1
        Wend
    End If
    
    ReDim Preserve aBordr(iBordr)               ' And finally
    aBordr(iBordr) = aBordr(0)                  ' to close the polygonss
                                                '
    Erase m_lpBordrPoints                       ' Clear existing points
          m_lpBordrPoints = aBordr              ' Copy new points
    Erase aBordr                                ' Array cleanup
    
End Sub
#End If

Private Sub CreateButtonRegion(EllipseW As Long, EllipseH As Long)
'   Create a button region with rounded corners.
    
    With m_tButtonSettings
        Dim hRgn  As Long
        Dim hRgn2 As Long
            ' Create button region with rounded corners as requested
            hRgn = CreateRoundRectRgn(0, 0, .Width + 1, .Height + 1, EllipseW, EllipseH)
            
        If (m_tButtonProperty.Shape = ebsCutLeft) Or _
           (m_tButtonProperty.Shape = ebsCutSides) Then
            hRgn2 = CreateRectRgn(0, 0, .Width / 2, .Height + 1)
            CombineRgn hRgn, hRgn, hRgn2, RGN_OR
            DeleteObject hRgn2
        End If
        
        If (m_tButtonProperty.Shape = ebsCutRight) Or _
           (m_tButtonProperty.Shape = ebsCutSides) Then
            hRgn2 = CreateRectRgn(.Width / 2, 0, .Width + 1, .Height + 1)
            CombineRgn hRgn, hRgn, hRgn2, RGN_OR
            DeleteObject hRgn2
        End If
        
        #If USE_CRYSTAL Then
        If (m_tButtonProperty.Style = ebsCrystal) Then
            CalculateRegionBorder hRgn, EllipseW, EllipseH
        End If
        #End If
        
        #If USE_MAC Then
        If (m_tButtonProperty.Style = ebsMac) Then
            CalculateRegionBorder hRgn, EllipseW, EllipseH
        End If
        #End If
        
        #If USE_MACOSX Then
        If (m_tButtonProperty.Style = ebsMacOSX) Then
            CalculateRegionBorder hRgn, EllipseW, EllipseH
        End If
        #End If
        
        SetWindowRgn hWnd, hRgn, True ' Set new window region
        DeleteObject hRgn
    End With
    
End Sub

#If USE_STANDARD Then
Private Function CreateCheckeredBrush( _
        hdc As Long, _
        Color1 As Long, _
        Color2 As Long) As Long
'   Create checkered brush. (Down state of standard checkbox)
    
    Dim hDCBrush As Long
    Dim hBMBrush As Long
    Dim hOBBrush As Long
    
    hDCBrush = CreateCompatibleDC(hdc)              ' Create new dc
    hBMBrush = CreateCompatibleBitmap(hdc, 2, 2)    ' Create bitmap for dc
                                                    '
    hOBBrush = SelectObject(hDCBrush, hBMBrush)     ' Select bitmap to dc
                                                    '
    SetPixelV hDCBrush, 0, 0, Color1                ' Draw checkered pattern
    SetPixelV hDCBrush, 1, 1, Color1                '
    SetPixelV hDCBrush, 1, 0, Color2                '
    SetPixelV hDCBrush, 0, 1, Color2                '
                                                    '
    hBMBrush = SelectObject(hDCBrush, hOBBrush)     ' Restore old bitmap
                                                    '
    CreateCheckeredBrush = CreatePatternBrush(hBMBrush)
                                                    '
    DeleteObject hBMBrush                           ' Delete objects
    DeleteObject hOBBrush                           '
    DeleteDC hDCBrush                               '
    
End Function
#End If

Private Sub DrawButton(Optional state As eButtonStates = -1, Optional Force As Boolean)
'   Draws the button itself
    If (m_bControlHidden) Then Exit Sub ' Why draw if control was hidden?
    
    If (state = -1) Then
        state = m_tButtonSettings.state ' Get current button state
    End If
    
    If (Not Force) Then ' Always update when forced
        ' Check if same as previous state, if so then exit
        If (state = m_tButtonSettings.state) Then
            If (m_tButtonSettings.HasFocus = m_bButtonHasFocus) Then
                ' If the button has changed its focus state from
                ' last drawing then button display must be updated
                Exit Sub
            End If
        End If
    End If
    
    #If DEBUG_BUTTON Then
        Debug.Print Caption & " - DrawButton", Choose(state + 1, "Normal", "Hot", "Down", "Disabled"), IIf(Force, "Forced", "")
    #End If
    
    UserControl.Cls
    
    m_tButtonSettings.HasFocus = m_bButtonHasFocus
    m_tButtonSettings.state = state ' Store current button state
    
    Select Case m_tButtonProperty.Style
        #If USE_CRYSTAL Then
        Case ebsCrystal:        DrawButtonCrystalMac state
        #End If
        #If USE_MAC Then
        Case ebsMac:            DrawButtonCrystalMac state
        #End If
        #If USE_MACOSX Then
        Case ebsMacOSX:         DrawButtonCrystalMac state
        #End If
        #If USE_OFFICE2003 Then
        Case ebsOffice2003:     DrawButtonOffice2003 state
        #End If
        #If USE_OFFICEXP Then
        Case ebsOfficeXP:       DrawButtonOfficeXP state
        #End If
        #If USE_OPERABROWSER Then
        Case ebsOperaBrowser:   DrawButtonOpera state
        #End If
        #If USE_STANDARD Then
        Case ebsStandard:       DrawButtonStandard state
        #End If
        #If USE_XPBLUE Then
        Case ebsXPBlue:         DrawButtonXPStyle state
        #End If
        #If USE_XPOLIVEGREEN Then
        Case ebsXPOliveGreen:   DrawButtonXPStyle state
        #End If
        #If USE_XPSILVER Then
        Case ebsXPSilver:       DrawButtonXPStyle state
        #End If
        #If USE_XPTOOLBAR Then
        Case ebsXPToolbar:      DrawButtonXPToolbar state
        #End If
        #If USE_YAHOO Then
        Case ebsYahoo:          DrawButtonYahoo state
        #End If
    End Select
    
End Sub

#If USE_CRYSTAL Or USE_MAC Or USE_MACOSX Then
Private Sub DrawButtonCrystalMac(DrawState As eButtonStates)
'   Drawing procedure for Crystal & Mac button styles.
    
    If (m_bCalculateRects) Then ' Calculate drawing areas only when needed
        With m_tButtonSettings
            Dim X1 As Long
            Dim X2 As Long
            Dim Y1 As Long
            Dim Y2 As Long
            
            #If USE_CRYSTAL Then
            If (m_tButtonProperty.Style = ebsCrystal) Then
                CreateButtonRegion .Height, .Width
                X1 = .Width / 4
                If (X1 > 10) Then X1 = 10 ' Max
                X2 = .Width - X1 - 1
                Y1 = 1
                Y2 = .Height / 2
            End If
            #End If
            
            #If USE_MAC Then
            If (m_tButtonProperty.Style = ebsMac) Then
                CreateButtonRegion 13, 13
                X1 = 2
                X2 = .Width - X1 - 1
                Y1 = 1
                Y2 = 7
            End If
            #End If
            
            #If USE_MACOSX Then
            If (m_tButtonProperty.Style = ebsMacOSX) Then
                CreateButtonRegion .Height, .Width
                X1 = 0
                X2 = .Width - X1 - 1
                Y1 = 0
                Y2 = .Height / 2
            End If
            #End If
            
            CalculateRects
            
            SetRect .Focus, X1, Y1, X2, Y2 ' Position shine object (use focus rect)
        End With
    End If
    
    With m_tButtonProperty
        If (.CheckBox And .Value And .Enabled) Then ' Button on checked state
            Select Case DrawState
                Case ebsNormal:     DrawState = ebsDown
                Case ebsHot:        DrawState = ebsDown
                Case ebsDown:       DrawState = ebsDown
            End Select
        ElseIf (DrawState = ebsNormal) And (m_bMouseIsDown) Then
                DrawState = ebsHot
        End If
    End With
    
    Dim Color1 As Long
    Dim Color2 As Long
    
    With m_tButtonColors
        Select Case DrawState ' Set background color
            Case ebsNormal:     Color1 = .BackColor
            Case ebsHot:        Color1 = .HoverColor
            Case ebsDown:       Color1 = .DownColor
            Case ebsDisabled:   Color1 = .GrayColor
        End Select
        
        If (.StartColor = -1) Then
            Color2 = BlendColors(Color1, &HFFFFFF, 0.8)
        Else
            Color2 = .StartColor
        End If
        
        #If USE_CRYSTAL Then
        If (m_tButtonProperty.Style = ebsCrystal) Then
            DrawGradientEx Color1, Color2, 60
            DrawShineEffect Color2, BlendColors(Color1, Color2, 0.6)
        End If
        #End If
        
        #If USE_MAC Then
        If (m_tButtonProperty.Style = ebsMac) Then
            DrawGradientEx Color1, Color2, 30
            DrawShineEffect Color2, BlendColors(Color1, Color2, 0.2)
        End If
        #End If
        
        #If USE_MACOSX Then
        If (m_tButtonProperty.Style = ebsMacOSX) Then
            DrawGradientEx Color1, Color2, 60
            DrawShineEffect Color2, BlendColors(Color1, Color2)
        End If
        #End If
        
        Color2 = ShiftColor(Color1, 0.2)
    End With
    
    With m_tButtonSettings
        If (.state = ebsDisabled) Then ' Disabled state
            If (m_tButtonProperty.Value) Then ' Down disabled
                DrawIcon m_tButtonProperty.PicNormal, 1, 1
                DrawCaption Color2, 2, 2
                DrawCaption m_tButtonColors.GrayText, 1, 1
            Else
                m_tButtonProperty.PicOpacity = 0.2
                DrawIcon m_tButtonProperty.PicNormal
                DrawCaption Color2, 1, 1
                DrawCaption m_tButtonColors.GrayText
                m_tButtonProperty.PicOpacity = 1
            End If
            
            'PolylineEx m_lpBordrPoints, BlendColors(Color1, Color2, 0.8)
        Else
            If (.state = ebsDown) Or (m_tButtonProperty.Value) Then
                If (m_tButtonProperty.Value) And (.state = ebsDown) Then
                    DrawIconEffect Color2, 2, 2
                    DrawCaptionEffect m_tButtonColors.ForeColor, Color2, 2, 2
                Else
                    DrawIconEffect Color2, 1, 1
                    DrawCaptionEffect m_tButtonColors.ForeColor, Color2, 1, 1
                End If
            Else
                DrawIconEffect Color2
                DrawCaptionEffect m_tButtonColors.ForeColor, Color2
            End If
            
            '' Draw main border
            'If (m_bParentActive) And (m_bButtonHasFocus Or .Default) Then
            '    PolylineEx m_lpBordrPoints, BlendColors(Color1, Color2, 0.5)
            'Else ' Border color is slightly brighter for normal state
            '    PolylineEx m_lpBordrPoints, BlendColors(Color1, Color2, 0.8)
            'End If
        End If
        
        PolylineEx m_lpBordrPoints, BlendColors(Color1, Color2, 0.8)
    End With
    
End Sub
#End If

#If USE_OFFICE2003 Then
Private Sub DrawButtonOffice2003(DrawState As eButtonStates)
'   Drawing procedure for the office 2003 button style.
    
    If (m_bCalculateRects) Then ' Recreate region
        With m_tButtonSettings ' Other settings may have changed button region
            CreateButtonRegion 0, 0
            CalculateRects
        End With
    End If
    
    Dim Color1 As Long
    Dim Color2 As Long
    
    With m_tButtonProperty
        If (.CheckBox And .Value And .Enabled) Then ' Checkbox on checked state
            Select Case DrawState
                Case ebsNormal
                    If (m_bButtonIsDown) Then ' Button is held down
                        DrawState = ebsDown
                    Else
                        Color1 = m_tButtonColors.HoverColor
                        Color2 = m_tButtonColors.DownColor
                        DrawState = -1
                    End If
                Case ebsHot
                    DrawState = ebsDown
            End Select
        ElseIf (DrawState = ebsNormal) And (m_bMouseIsDown) Then
                DrawState = ebsHot
        End If
    End With
    
    With m_tButtonColors
        Select Case DrawState
            Case ebsNormal, ebsDisabled
                Color1 = .StartColor
                Color2 = .BackColor
            Case ebsHot
                Color1 = .FocusBorder
                Color2 = .HoverColor
            Case ebsDown
                Color1 = .DownColor
                Color2 = .HoverColor
        End Select
    End With
    
    ' Draw button background
    DrawGradientEx Color1, Color2
    
    With m_tButtonSettings
        If (.state = ebsDisabled) Then ' Disabled state
            DrawIcon m_tButtonProperty.PicNormal
            DrawCaption m_tButtonColors.GrayText
        Else
            Color1 = BlendColors(Color1, Color2)
            DrawIconEffect Color1
            DrawCaptionEffect m_tButtonColors.ForeColor, Color1
            
            If (.state <> ebsNormal Or m_tButtonProperty.Value) Or _
               (.state = ebsNormal And m_bButtonIsDown) Then
                RectangleEx m_tButtonColors.ForeColor, 0, 0, .Width, .Height
            End If
        End If
    End With
    
End Sub
#End If

#If USE_OFFICEXP Then
Private Sub DrawButtonOfficeXP(DrawState As eButtonStates)
'   Drawing procedure for the office xp button style.
    
    If (m_bCalculateRects) Then ' Recreate region
        With m_tButtonSettings ' Other settings may have changed button region
            CreateButtonRegion 0, 0
            CalculateRects
        End With
    End If
    
    With m_tButtonProperty
        If (.CheckBox And .Value And .Enabled) Then ' Checkbox on checked state
            Select Case DrawState
                Case ebsNormal
                    If (m_bButtonIsDown) Then ' Button is held down
                        DrawState = ebsDown
                    End If
                Case ebsHot
                    DrawState = ebsDown
            End Select
        ElseIf (DrawState = ebsNormal) And (m_bMouseIsDown) Then
            DrawState = ebsHot
        End If
    End With
    
    Dim Color As Long
    
    With m_tButtonColors
        Select Case DrawState
            Case ebsNormal:     Color = .BackColor
            Case ebsHot:        Color = .HoverColor
            Case ebsDown:       Color = .DownColor
            Case ebsDisabled:   Color = .GrayColor
        End Select
        
        FillButtonEx Color ' Draw button background
    End With
    
    With m_tButtonSettings
        If (.state = ebsDisabled) Then ' Disabled state
            DrawIcon m_tButtonProperty.PicNormal
            DrawCaption m_tButtonColors.GrayText
        Else
            If (.state = ebsHot And Not m_tButtonProperty.Value And _
                m_tButtonProperty.PicHot Is Nothing And Not m_bMouseIsDown) Then
                
                DrawIcon m_tButtonProperty.PicNormal, 1, 1, &HC0C0C0
                DrawIcon m_tButtonProperty.PicNormal, -1, -1 ' Do not draw any effect
            Else
                If (.state = ebsDown) And (m_tButtonProperty.Value) Then
                    DrawIconEffect Color, 1, 1
                Else
                    DrawIconEffect Color
                End If
            End If
            
            If (.state = ebsDown) And (m_tButtonProperty.Value) Then
                DrawCaptionEffect m_tButtonColors.ForeColor, Color, 1, 1
            Else
                DrawCaptionEffect m_tButtonColors.ForeColor, Color
            End If
            
            If (.state <> ebsNormal Or m_tButtonProperty.Value) Or _
               (.state = ebsNormal And m_bButtonIsDown) Then
                RectangleEx m_tButtonColors.FocusBorder, 0, 0, .Width, .Height
            End If
        End If
    End With
    
End Sub
#End If

#If USE_OPERABROWSER Then
Private Sub DrawButtonOpera(DrawState As eButtonStates)
'   Drawing procedure for the opera browser button style.
    
    If (m_bCalculateRects) Then ' Recalculate drawing areas and recreate region
        CreateButtonRegion 0, 0
        CalculateRects
    End If
    
    With m_tButtonProperty
        If (.CheckBox And .Value And .Enabled) Then ' Checkbox on checked state
            DrawState = ebsHot
        End If
    End With
    
    Dim Color1 As Long
    Dim Color2 As Long
    
    With m_tButtonColors
        Select Case DrawState ' Draw button background
            Case ebsNormal
                DrawGradientEx .StartColor, .BackColor, 70
                Color1 = BlendColors(.StartColor, .BackColor)
            Case ebsHot, ebsDown
                Color1 = .HoverColor
                FillButtonEx Color1
            Case ebsDisabled
                DrawGradientEx .StartColor, .GrayColor, 70
        End Select
    End With
    
    With m_tButtonSettings
        If (.state = ebsDisabled) Then ' Disabled state
            If (m_tButtonProperty.Value) Then
                DrawIcon m_tButtonProperty.PicNormal, 1, 1 ' Force no special effect
                DrawCaption m_tButtonColors.GrayText, 1, 1
            Else
                DrawIcon m_tButtonProperty.PicNormal
                DrawCaption m_tButtonColors.GrayText
            End If
            
            Color1 = m_tButtonColors.GrayColor
            DrawLine ShiftColor(Color1, 0.1), 0, 0, .Width - 1, 0                       ' Top
            DrawLine ShiftColor(Color1, -0.1), 0, .Height - 1, .Width - 1, .Height - 1  ' Bottom
            DrawLine ShiftColor(Color1, -0.05), 0, 1, 0, .Height - 2                    ' Left
            DrawLine ShiftColor(Color1, -0.05), .Width - 1, 1, .Width - 1, .Height - 2  ' Right
        Else
            If (.state = ebsDown) Or (m_tButtonProperty.Value) Then
                If (.state = ebsDown) And (m_tButtonProperty.Value) Then
                    DrawIconEffect Color1, 2, 2
                    DrawCaptionEffect m_tButtonColors.ForeColor, Color1, 2, 2
                Else
                    DrawIconEffect Color1, 1, 1
                    DrawCaptionEffect m_tButtonColors.ForeColor, Color1, 1, 1
                End If
            Else
                DrawIconEffect Color1
                DrawCaptionEffect m_tButtonColors.ForeColor, Color1
            End If
            
            If (m_bParentActive) And (m_bButtonHasFocus Or .Default) Then   ' Draw borders
                Color2 = ShiftColor(Color1, 0.1)                            '
                DrawLine Color2, 1, 1, .Width - 2, 1                        ' Top
                DrawLine Color2, 2, 2, .Width - 3, 2                        '
                Color2 = ShiftColor(Color1, -0.1)                           '
                DrawLine Color2, 1, .Height - 2, .Width - 2, .Height - 2    ' Bottom
                DrawLine Color2, 2, .Height - 3, .Width - 3, .Height - 3    '
                Color2 = ShiftColor(Color1, -0.05)                          '
                DrawLine Color2, 1, 2, 1, .Height - 3                       ' Left
                DrawLine Color2, 2, 3, 2, .Height - 4                       '
                DrawLine Color2, .Width - 2, 2, .Width - 2, .Height - 3     ' Right
                DrawLine Color2, .Width - 3, 3, .Width - 3, .Height - 4     '
                
                RectangleEx m_tButtonColors.ForeColor, 0, 0, .Width, .Height
            Else
                DrawLine ShiftColor(Color1, 0.1), 0, 0, .Width - 1, 0                       ' Top
                DrawLine ShiftColor(Color1, -0.1), 0, .Height - 1, .Width - 1, .Height - 1  ' Bottom
                DrawLine ShiftColor(Color1, -0.05), 0, 1, 0, .Height - 2                    ' Left
                DrawLine ShiftColor(Color1, -0.05), .Width - 1, 1, .Width - 1, .Height - 2  ' Right
            End If
        End If
    End With
    
End Sub
#End If

#If USE_STANDARD Then
Private Sub DrawButtonStandard(DrawState As eButtonStates)
'   Drawing procedure for the standard command button.
    
    If (m_bCalculateRects) Then ' Used to reset button region
        CreateButtonRegion 0, 0
        CalculateRects 4, 4
    End If
    
    With m_tButtonProperty
        If (.CheckBox And .Value) Then ' Checkbox on checked state
            Dim hPen As Long
            Dim hOld As Long
            
            hPen = CreateCheckeredBrush(UserControl.hdc, &HFFFFFF, m_tButtonColors.DownColor)
            hOld = SelectObject(UserControl.hdc, hPen)
            
            ' Draw a checkered background for checkbox down state
            PatBlt UserControl.hdc, _
                   2, _
                   2, _
                   m_tButtonSettings.Width - 4, _
                   m_tButtonSettings.Height - 4, _
                   PATCOPY
            
            SelectObject UserControl.hdc, hOld
            DeleteObject hPen
        Else
            FillButtonEx m_tButtonColors.BackColor
        End If
    End With
    
    With m_tButtonSettings
        If (.state = ebsDisabled) Then ' Button is disabled
            If (m_tButtonProperty.CheckBox) Then
                If (m_tButtonProperty.Value) Then
                    DrawIcon m_tButtonProperty.PicNormal, 1, 1 ' Force no special effect
                    DrawCaption m_tButtonColors.GrayText, 1, 1
                Else
                    DrawIcon m_tButtonProperty.PicNormal
                    DrawCaption m_tButtonColors.GrayText
                End If
            Else
                DrawIcon m_tButtonProperty.PicNormal
                DrawCaption BlendColors(m_tButtonColors.GrayText, &HFFFFFF, 0.9), 1, 1
                DrawCaption m_tButtonColors.GrayText
            End If
        Else
            If (Not .state = ebsDown) And Not (m_tButtonProperty.Value) Then
                DrawIconEffect m_tButtonColors.BackColor
                DrawCaptionEffect m_tButtonColors.ForeColor, m_tButtonColors.BackColor
            Else
                DrawIconEffect m_tButtonColors.DownColor, 1, 1
                DrawCaptionEffect m_tButtonColors.ForeColor, m_tButtonColors.DownColor, 1, 1
            End If
            
            If (m_bParentActive) And (m_bButtonHasFocus Or .Default) Then ' Draw border
                If (Not m_tButtonProperty.CheckBox) Then
                    ' Focus border
                    RectangleEx m_tButtonColors.ForeColor, 0, 0, .Width, .Height
                ElseIf (Not m_tButtonProperty.Value) Then
                    RectangleEx m_tButtonColors.ForeColor, 0, 0, .Width, .Height
                End If
                
                If (m_bButtonHasFocus) Then
                    If (m_tButtonProperty.CheckBox And m_tButtonProperty.Value) Then
                        ' Draw a rectangle in exchange of the FocusRect for checkbox
                        RectangleEx m_tButtonColors.ForeColor, _
                                    .Focus.Left, _
                                    .Focus.Top, _
                                    .Focus.Right, _
                                    .Focus.bottom
                    Else
                        DrawFocusRect UserControl.hdc, .Focus
                    End If
                End If
            End If
        End If
        
        Select Case DrawState ' Draw button border
            Case ebsNormal, ebsDisabled, ebsHot
                If (m_tButtonProperty.Value) And (m_tButtonProperty.CheckBox) Then
                    DrawEdgeEx 0, 0, .Width, .Height, EDGE_SUNKEN
                Else
                    If (m_bButtonHasFocus) Or (.Default) Then
                        DrawEdgeEx 1, 1, .Width - 1, .Height - 1, EDGE_RAISED
                    Else
                        DrawEdgeEx 0, 0, .Width, .Height, EDGE_RAISED
                    End If
                End If
            Case ebsDown
                If (m_tButtonProperty.CheckBox) Then
                    DrawEdgeEx 0, 0, .Width, .Height, EDGE_SUNKEN
                Else
                    RectangleEx m_tButtonColors.ForeColor, 0, 0, .Width, .Height
                    RectangleEx m_tButtonColors.FocusBorder, 1, 1, .Width - 1, .Height - 1
                End If
        End Select
    End With
    
End Sub
#End If

#If USE_XPBLUE Or USE_XPOLIVEGREEN Or USE_XPSILVER Then
Private Sub DrawButtonXPStyle(DrawState As eButtonStates)
'   Drawing procedure for the xp button styles.
    
    If (m_bCalculateRects) Then ' Recreate region
        CreateButtonRegion 5, 5
        CalculateRects
    End If
    
    With m_tButtonProperty
        If (.CheckBox And .Value And .Enabled) Then ' Checkbox on checked state
            #If USE_XPSILVER Then                   '
            If (.Style = ebsXPSilver) Then          '
                DrawState = ebsDown                 '
            #If USE_XPBLUE Or USE_XPOLIVEGREEN Then ' There are some codes that
            ElseIf (DrawState = ebsDown) Then       ' are seemed to be repeated
                DrawState = ebsNormal               ' but they are intended to
            #End If                                 ' be arranged in that way...
            End If                                  '
            #Else                                   ' You don't need to fully
                If (DrawState = ebsDown) Then       ' understand why coz' it really
                    DrawState = ebsNormal           ' doesn't matter, they are just
                End If                              ' my way of excluding codes
            #End If                                 ' when a button style is not
        End If                                      ' actually needed or is unset.
    End With
    
    Dim Color1 As Long
    Dim Color2 As Long
    
    With m_tButtonSettings
        Select Case DrawState ' Draw button background
            Case ebsNormal, ebsHot
                If (m_tButtonProperty.Value) Then
                    Color2 = ShiftColor(m_tButtonColors.BackColor, -0.1)
                    DrawGradientEx Color2, m_tButtonColors.StartColor
                    Color2 = BlendColors(Color2, m_tButtonColors.StartColor)
                Else
                    DrawGradientEx m_tButtonColors.StartColor, m_tButtonColors.BackColor
                    Color2 = BlendColors(m_tButtonColors.StartColor, m_tButtonColors.BackColor)
                End If
            Case ebsDown
                #If USE_XPSILVER Then
                If (m_tButtonProperty.Style = ebsXPSilver) Then
                    DrawGradientEx m_tButtonColors.DownColor, m_tButtonColors.StartColor, 80
                    Color2 = BlendColors(m_tButtonColors.DownColor, m_tButtonColors.StartColor, 0.3)
                #If USE_XPBLUE Or USE_XPOLIVEGREEN Then                         '
                Else                                                            '
                    Color2 = m_tButtonColors.DownColor                          '
                    FillButtonEx Color2                                         '
                                                                                '
                    Color1 = ShiftColor(Color2, -0.05)                          '
                    DrawLine Color1, 2, 1, .Width - 3, 1                        ' Top
                    DrawLine Color1, 1, 2, 1, .Height - 3                       ' Left
                                                                                '
                    Color1 = ShiftColor(Color2, -0.03)                          '
                    DrawLine Color1, 2, 2, .Width - 3, 2                        ' Top
                    DrawLine Color1, 2, 3, 2, .Height - 4                       ' Left
                                                                                '
                    Color1 = ShiftColor(Color2, 0.02)                           '
                    DrawLine Color1, 2, .Height - 3, .Width - 3, .Height - 3    ' Bottom
                    DrawLine Color1, .Width - 3, 3, .Width - 3, .Height - 3     ' Right
                                                                                '
                    Color1 = ShiftColor(Color2, 0.03)                           '
                    DrawLine Color1, 2, .Height - 2, .Width - 3, .Height - 2    ' Bottom
                    DrawLine Color1, .Width - 2, 2, .Width - 2, .Height - 4     ' Right
                #End If                                                         '
                End If                                                          '
                #Else                                                           '
                    Color2 = m_tButtonColors.DownColor                          '
                    FillButtonEx Color2                                         '
                                                                                '
                    Color1 = ShiftColor(Color2, -0.05)                          '
                    DrawLine Color1, 2, 1, .Width - 3, 1                        ' Top
                    DrawLine Color1, 1, 2, 1, .Height - 3                       ' Left
                                                                                '
                    Color1 = ShiftColor(Color2, -0.03)                          '
                    DrawLine Color1, 2, 2, .Width - 3, 2                        ' Top
                    DrawLine Color1, 2, 3, 2, .Height - 4                       ' Left
                                                                                '
                    Color1 = ShiftColor(Color2, 0.02)                           '
                    DrawLine Color1, 2, .Height - 3, .Width - 3, .Height - 3    ' Bottom
                    DrawLine Color1, .Width - 3, 3, .Width - 3, .Height - 3     ' Right
                                                                                '
                    Color1 = ShiftColor(Color2, 0.03)                           '
                    DrawLine Color1, 2, .Height - 2, .Width - 3, .Height - 2    ' Bottom
                    DrawLine Color1, .Width - 2, 2, .Width - 2, .Height - 4     ' Right
                #End If                                                         '
            Case ebsDisabled
                FillButtonEx m_tButtonColors.GrayColor
        End Select
    End With
    
    With m_tButtonSettings
        If (.state = ebsDisabled) Then ' Disabled state
            DrawIcon m_tButtonProperty.PicNormal
            Color1 = m_tButtonColors.GrayText
            DrawCaption m_tButtonColors.GrayText
        Else
            DrawIconEffect Color2
            DrawCaptionEffect m_tButtonColors.ForeColor, Color2
            
            ' If button is held down and cursor is moved out of the control,
            ' the Hot state border is displayed until mouse button is released
            ' or cursor is moved back to the button
            
            If (.state = ebsHot) Or (m_bMouseIsDown And Not m_bMouseOnButton And Not DrawState = ebsDown) Then
                ' Hot state border (not shown when button is on checked state)
                If (Not m_tButtonProperty.Value) Then
                    Color1 = BlendColors(m_tButtonColors.HoverColor, &HFFFFFF)                  ' Prepare colors
                    Color2 = ShiftColor(m_tButtonColors.HoverColor, -0.02)                      '
                                                                                                '
                    DrawLine Color1, 1, 1, .Width - 2, 1                                        ' Top
                    DrawGradientEx Color1, Color2, , 1, 2, 1, .Height - 3                       ' Left
                    DrawGradientEx Color1, Color2, , .Width - 2, 2, .Width - 2, .Height - 3     ' Right
                                                                                                '
                    Color1 = ShiftColor(Color1, -0.05)                                          '
                    Color2 = m_tButtonColors.HoverColor                                         '
                                                                                                '
                    DrawLine Color1, 2, 2, .Width - 3, 2                                        ' Top
                    DrawGradientEx Color1, Color2, , 2, 2, 2, .Height - 3                       ' Left
                    DrawGradientEx Color1, Color2, , .Width - 3, 2, .Width - 3, .Height - 3     ' Right
                                                                                                '
                    DrawLine Color2, 2, .Height - 3, .Width - 3, .Height - 3                    ' Bottom
                    DrawLine ShiftColor(Color2, -0.05), 1, .Height - 2, .Width - 2, .Height - 2 ' Bottom
                #If USE_XPSILVER Then                                                           '
                ElseIf (m_tButtonProperty.Style = ebsXPSilver) Then                             '
                    RectangleEx m_tButtonColors.StartColor, 1, 1, .Width - 1, .Height           ' White border
                #End If                                                                         '
                End If
            Else
                #If USE_XPSILVER Then
                ' White border is drawn inside the main border on down state
                If (m_tButtonProperty.Style = ebsXPSilver) Then
                    If (m_bParentActive) And (m_bButtonHasFocus And Not .state = ebsDown And Not m_tButtonProperty.Value) Then
                        RectangleEx m_tButtonColors.StartColor, 2, 1, .Width - 2, .Height
                    Else
                        RectangleEx m_tButtonColors.StartColor, 1, 1, .Width - 1, .Height
                    End If
                End If
                #End If
                
                If (m_bParentActive) And (Not .state = ebsDown And (m_bButtonHasFocus Or .Default)) Then
                    ' Focus state border (not shown when button is in checked state)
                    If (Not m_tButtonProperty.Value) Then
                        Color1 = BlendColors(m_tButtonColors.FocusBorder, &HFFFFFF, 0.4)            ' Prepare colors
                        Color2 = m_tButtonColors.FocusBorder                                        '
                                                                                                    '
                        DrawLine ShiftColor(Color1, 0.05), 1, 1, .Width - 2, 1                      ' Top
                        DrawLine Color1, 2, 2, .Width - 3, 2                                        '
                        DrawLine Color2, 2, .Height - 3, .Width - 3, .Height - 3                    ' Bottom
                        DrawLine ShiftColor(Color2, -0.02), 1, .Height - 2, .Width - 2, .Height - 2 '
                                                                                                    '
                        #If USE_XPSILVER Then                                                       '
                        If (m_tButtonProperty.Style = ebsXPSilver) Then                             '
                            DrawGradientEx Color1, Color2, , 1, 2, 1, .Height - 3                   ' Left
                            DrawGradientEx Color1, Color2, , .Width - 2, 2, .Width - 2, .Height - 3 ' Right
                        #If USE_XPBLUE Or USE_XPOLIVEGREEN Then                                     '
                        Else                                                                        '
                            DrawGradientEx Color1, Color2, , 1, 2, 2, .Height - 3                   ' Left
                            DrawGradientEx Color1, Color2, , .Width - 3, 2, .Width - 2, .Height - 3 ' Right
                        #End If                                                                     '
                        End If                                                                      '
                        #Else                                                                       '
                            DrawGradientEx Color1, Color2, , 1, 2, 2, .Height - 3                   ' Left
                            DrawGradientEx Color1, Color2, , .Width - 3, 2, .Width - 2, .Height - 3 ' Right
                        #End If                                                                     '
                    End If
                End If
            End If
            ' Border color
            Color1 = BlendColors(m_tButtonColors.FocusBorder, &H0)
        End If
        
        ' Draw main button border
        RectangleEx Color1, 0, 0, .Width, .Height
        
        If Not (m_tButtonProperty.Shape = ebsCutLeft Or _
                m_tButtonProperty.Shape = ebsCutSides) Then
            SetPixelV hdc, 1, 1, Color1
            SetPixelV hdc, 1, .Height - 2, Color1
        End If
        
        If Not (m_tButtonProperty.Shape = ebsCutRight Or _
                m_tButtonProperty.Shape = ebsCutSides) Then
            SetPixelV hdc, .Width - 2, 1, Color1
            SetPixelV hdc, .Width - 2, .Height - 2, Color1
        End If
    End With
    
End Sub
#End If

#If USE_XPTOOLBAR Then
Private Sub DrawButtonXPToolbar(DrawState As eButtonStates)
'   Drawing procedure for the xp toolbar button style.
    
    If (m_bCalculateRects) Then ' Used to reset button region
        CreateButtonRegion 5, 5
        CalculateRects
    End If
    
    Dim Color1 As Long
    Dim Color2 As Long
    
    With m_tButtonProperty
        If (.Value And Not DrawState = ebsDown) Then ' Checkbox on checked state
            Color1 = BlendColors(m_tButtonColors.BackColor, &HFFFFFF, 0.8)
        ElseIf (DrawState = ebsDown) Then
            Color1 = m_tButtonColors.DownColor
        ElseIf (DrawState = ebsHot And (m_bMouseOnButton)) Then
            Color1 = m_tButtonColors.HoverColor
        Else
            Color1 = m_tButtonColors.BackColor
        End If
        
        FillButtonEx Color1
    End With
    
    With m_tButtonSettings
        If (.state = ebsDisabled) Then ' Button is disabled
            If (m_tButtonProperty.Value) Then
                DrawIcon m_tButtonProperty.PicNormal, 1, 1
                DrawCaption BlendColors(m_tButtonColors.GrayText, &HFFFFFF, 0.9), 2, 2
                DrawCaption m_tButtonColors.GrayText, 1, 1
                
                Color2 = ShiftColor(Color1, -0.2)
                RectangleEx Color2, 0, 0, .Width, .Height
            Else
                DrawIcon m_tButtonProperty.PicNormal
                DrawCaption BlendColors(m_tButtonColors.GrayText, &HFFFFFF, 0.9), 1, 1
                DrawCaption m_tButtonColors.GrayText
            End If
        Else
            If (.state = ebsDown) Then
                DrawIconEffect Color1, 1, 1
                DrawCaptionEffect m_tButtonColors.StartColor, Color1, 1, 1
            ElseIf (m_tButtonProperty.Value) Then
                DrawIconEffect Color1, 1, 1
                DrawCaptionEffect m_tButtonColors.ForeColor, Color1, 1, 1
            Else
                DrawIconEffect Color1
                DrawCaptionEffect m_tButtonColors.ForeColor, Color1
            End If
            
            If (Not m_tButtonProperty.Value) And (DrawState = ebsNormal) Then
                Exit Sub
            End If
            
            Color2 = ShiftColor(Color1, -0.2)
            
            If (DrawState = ebsHot) And (m_bMouseOnButton Or m_bSpacebarIsDown) Then
                DrawLine BlendColors(Color1, &HFFFFFF, 0.4), 1, 1, .Width - 3, 1                        ' Top
                DrawLine BlendColors(Color1, &HFFFFFF, 0.3), 1, 2, .Width - 3, 2                        '
                DrawLine BlendColors(Color1, Color2, 0.4), .Width - 2, 2, .Width - 2, .Height - 3       ' Right
                DrawLine BlendColors(Color1, Color2, 0.2), .Width - 3, 1, .Width - 3, .Height - 2       '
                DrawLine BlendColors(Color1, Color2, 0.5), 2, .Height - 2, .Width - 3, .Height - 2      ' Bottom
                DrawLine BlendColors(Color1, Color2, 0.3), 1, .Height - 3, .Width - 3, .Height - 3      '
                DrawLine BlendColors(Color1, Color2, 0.1), 1, .Height - 4, .Width - 4, .Height - 4      '
                RectangleEx Color2, 0, 0, .Width, .Height                                               ' Main border
            ElseIf (DrawState = ebsDown) Then                                                           '
                DrawLine BlendColors(Color1, Color2, 0.2), 1, 2, 1, .Height - 3                         ' Left
                DrawLine BlendColors(Color1, Color2, 0.1), 2, 1, 2, .Height - 2                         '
                DrawLine BlendColors(Color1, Color2, 0.2), 2, 1, .Width - 3, 1                          ' Top
                DrawLine BlendColors(Color1, Color2, 0.1), 3, 2, .Width - 2, 2                          '
                DrawLine BlendColors(Color1, &HFFFFFF, 0.2), 3, .Height - 2, .Width - 3, .Height - 2    ' Bottom
                DrawLine BlendColors(Color1, &HFFFFFF, 0.1), 3, .Height - 3, .Width - 2, .Height - 3    '
                RectangleEx Color2, 0, 0, .Width, .Height                                               ' Main border
            ElseIf (m_tButtonProperty.Value) Then                                                       '
                RectangleEx Color2, 0, 0, .Width, .Height                                               '
            End If                                                                                      '
        End If
        
        If (DrawState = ebsHot Or DrawState = ebsDown Or m_tButtonProperty.Value) Then
            If Not (m_tButtonProperty.Shape = ebsCutLeft Or _
                    m_tButtonProperty.Shape = ebsCutSides) Then
                SetPixelV hdc, 1, 1, Color2
                SetPixelV hdc, 1, .Height - 2, Color2
            End If
            
            If Not (m_tButtonProperty.Shape = ebsCutRight Or _
                    m_tButtonProperty.Shape = ebsCutSides) Then
                SetPixelV hdc, .Width - 2, 1, Color2
                SetPixelV hdc, .Width - 2, .Height - 2, Color2
            End If
        End If
    End With
    
End Sub
#End If

#If USE_YAHOO Then
Private Sub DrawButtonYahoo(DrawState As eButtonStates)
'   Drawing procedure for the yahoo style button
    
    If (m_bCalculateRects) Then ' Used to create/recreate button region
        CreateButtonRegion 5, 5
        CalculateRects
    End If
    
    With m_tButtonProperty
        If (.CheckBox And .Value And .Enabled) Then ' Checkbox on checked state
            DrawState = ebsDown
        End If
    End With
    
    Dim Color As Long
    
    With m_tButtonColors
        Select Case DrawState ' Draw button background
            Case ebsNormal:     Color = .BackColor
            Case ebsHot:        Color = .HoverColor
            Case ebsDown:       Color = .DownColor
            Case ebsDisabled:   FillButtonEx .GrayColor
        End Select
    End With
    
    If (Not DrawState = ebsDisabled) Then
        DrawGradientEx m_tButtonColors.StartColor, Color, 30
        Color = BlendColors(m_tButtonColors.StartColor, Color)
    End If
    
    With m_tButtonSettings
        If (.state = ebsDisabled) Then ' Disabled state
            If (m_tButtonProperty.Value) Then ' Down disabled
                DrawIcon m_tButtonProperty.PicNormal, 1, 1
                DrawCaption m_tButtonColors.GrayText, 1, 1
            Else
                DrawIcon m_tButtonProperty.PicNormal
                DrawCaption m_tButtonColors.GrayText
            End If
            
            Color = m_tButtonColors.GrayText ' Border color
        Else
            If (.state = ebsDown Or m_tButtonProperty.Value) Then
                If (.state = ebsDown And m_tButtonProperty.Value) Then
                    DrawIconEffect Color, 2, 2
                    DrawCaptionEffect m_tButtonColors.ForeColor, Color, 2, 2
                Else
                    DrawIconEffect Color, 1, 1
                    DrawCaptionEffect m_tButtonColors.ForeColor, Color, 1, 1
                End If
            Else
                DrawIconEffect Color
                DrawCaptionEffect m_tButtonColors.ForeColor, Color
            End If
            
            Color = m_tButtonColors.BackColor
            
            If (m_bParentActive) And (m_bButtonHasFocus Or .Default) Then
                DrawLine Color, 1, 1, .Width - 2, 1
                DrawLine Color, 1, .Height - 2, .Width - 2, .Height - 2
                
                If Not (m_tButtonProperty.Shape = ebsCutLeft Or _
                        m_tButtonProperty.Shape = ebsCutSides) Then
                    DrawLine Color, 1, 1, 1, .Height - 2
                    SetPixelV hdc, 2, 2, Color
                    SetPixelV hdc, 2, .Height - 3, Color
                End If
                
                If Not (m_tButtonProperty.Shape = ebsCutRight Or _
                        m_tButtonProperty.Shape = ebsCutSides) Then
                    DrawLine Color, .Width - 2, 1, .Width - 2, .Height - 2
                    SetPixelV hdc, .Width - 3, 2, Color
                    SetPixelV hdc, .Width - 3, .Height - 3, Color
                End If
            End If
        End If
        
        ' Draw border
        RectangleEx Color, 0, 0, .Width, .Height
        
        If Not (m_tButtonProperty.Shape = ebsCutLeft Or _
                m_tButtonProperty.Shape = ebsCutSides) Then
            SetPixelV hdc, 1, 1, Color
            SetPixelV hdc, 1, .Height - 2, Color
        End If
        
        If Not (m_tButtonProperty.Shape = ebsCutRight Or _
                m_tButtonProperty.Shape = ebsCutSides) Then
            SetPixelV hdc, .Width - 2, 1, Color
            SetPixelV hdc, .Width - 2, .Height - 2, Color
        End If
    End With
    
End Sub
#End If

Private Sub DrawCaption(Color As Long, Optional MoveX As Long, Optional MoveY As Long)
'   Draw the button's caption
    
    Dim tx As String                            '
    Dim tn As Long                              '
        tx = m_tButtonProperty.Caption          ' Get caption and its length
        tn = Len(tx)                            '
                                                '
    If (tn = 0) Then Exit Sub                   ' Bail procedure if no
                                                ' defined caption
    Dim rc As RECT                              '
        CopyRect rc, m_tButtonSettings.Caption  ' Get drawing area
                                                '
    If (Not MoveX = 0) Or (Not MoveY = 0) Then  '
        OffsetRect rc, MoveX, MoveY             ' Move drawing area for how many pixels
    End If                                      ' from the original area if set
                                                '
    SetTextColor hdc, Color                     ' Set text color
                                                '
    If (m_bIsPlatformNT) Then                   '
        DrawTextW hdc, StrPtr(tx), tn, rc, DT_DRAWFLAG
    Else                                        '
        DrawText hdc, tx, tn, rc, DT_DRAWFLAG   ' Draw caption as ansi/unicode
    End If                                      '
    
End Sub

Private Sub DrawCaptionEffect( _
        TextColor As Long, _
        BackColor As Long, _
        Optional MoveX As Long, _
        Optional MoveY As Long)
'   Draw the button's caption with special effect defined applied.
    
    #If USE_SPECIALEFFECTS Then
    If (Not m_tButtonProperty.Effects = eseNone) Then
        Dim Color1 As Long
        Dim Color2 As Long
            Color1 = ShiftColor(BackColor, 0.1)
            Color2 = ShiftColor(BackColor, -0.1)
    End If
    
    Select Case m_tButtonProperty.Effects ' Draw effects first
        Case eseEmbossed
            DrawCaption Color2, MoveX + 1, MoveY + 1
            DrawCaption Color1, MoveX - 1, MoveY - 1
        Case eseEngraved
            DrawCaption Color2, MoveX - 1, MoveY - 1
            DrawCaption Color1, MoveX + 1, MoveY + 1
        Case eseShadowed
            DrawCaption Color2, MoveX + 1, MoveY + 1
    End Select
    #End If
    
    DrawCaption TextColor, MoveX, MoveY
    
End Sub

#If USE_STANDARD Then
Private Sub DrawEdgeEx(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Edge As Long)
'   Perform customized DrawEdge API function
    Dim lpRect As RECT
    
    SetRect lpRect, X1, Y1, X2, Y2
    DrawEdge UserControl.hdc, lpRect, Edge, BF_RECT Or BF_SOFT
    
End Sub
#End If

Private Sub DrawGradientEx( _
        StartColor As Long, _
        EndColor As Long, _
        Optional Center As Single = 50, _
        Optional ByVal X1 As Long, _
        Optional ByVal Y1 As Long, _
        Optional ByVal X2 As Long = -1, _
        Optional ByVal Y2 As Long = -1) ' Center in percent
'   Draw vertical gradient effect on the control on specified coordinates.
    
    If (X2 = -1) Then X2 = m_tButtonSettings.Width - 1
    If (Y2 = -1) Then Y2 = m_tButtonSettings.Height - 1
    
    X2 = TranslateNumber(X2, X1, X2)            ' X2 must not be < than X1
    Y2 = TranslateNumber(Y2, Y1, Y2)            ' Y2 must not be < than Y1
                                                '
    Dim Color As Long                           '
    Dim Step As Single                          '
                                                '
    Dim RGB1 As tRGB                            ' Start color
    Dim RGB2 As tRGB                            ' End color
    Dim RGB3 As tRGB                            ' Mid color
    Dim RGB4 As tRGB                            ' Gradient color
                                                '
    Center = TranslateNumber(Center, 0, 100)    ' Center should not exceed 0 to 100
                                                '
    RGB1 = GetRGB(StartColor)                   ' Get RGB color values
    RGB2 = GetRGB(EndColor)                     '
                                                '
    RGB3.r = RGB1.r + (RGB2.r - RGB1.r) * 0.5   ' Blend start and end color
    RGB3.G = RGB1.G + (RGB2.G - RGB1.G) * 0.5   '
    RGB3.b = RGB1.b + (RGB2.b - RGB1.b) * 0.5   '
                                                '
    Center = (Y2 - Y1 - 1) * Center / 100       ' Converts center in percent
    Center = (Y1 + Center)                      ' to actual pixel coordinate
                                                '
    If (Center = 0) Then Center = 1             ' Avoid errors
                                                '
    While (Y1 <= Y2)                            ' Draw from top to bottom
        If (Y1 <= Center) Then                  '
            Step = Y1 / Center                  '
            RGB4.r = RGB1.r + (RGB3.r - RGB1.r) * Step
            RGB4.G = RGB1.G + (RGB3.G - RGB1.G) * Step
            RGB4.b = RGB1.b + (RGB3.b - RGB1.b) * Step
        Else                                    '
            Step = (Y1 - Center) / (Y2 - Center)
            RGB4.r = RGB3.r + (RGB2.r - RGB3.r) * Step
            RGB4.G = RGB3.G + (RGB2.G - RGB3.G) * Step
            RGB4.b = RGB3.b + (RGB2.b - RGB3.b) * Step
        End If                                  '
                                                '
        Color = RGBEx(RGB4.r, RGB4.G, RGB4.b)   ' Prepare color
                                                '
        If (X1 = X2) Then                       '
            SetPixelV hdc, X1, Y1, Color        ' Draw point to device
        Else                                    '
            DrawLine Color, X1, Y1, X2, Y1      ' Draw line to device
        End If                                  '
                                                '
        Y1 = Y1 + 1                             '
    Wend                                        '
    
End Sub

Private Sub DrawIcon( _
        Picture As StdPicture, _
        Optional MoveX As Long, _
        Optional MoveY As Long, _
        Optional BrushColor As Long = -1)
'   Draw icon to button with the specified settings.
    Dim hBMBck As Long
    Dim hDCBck As Long
    Dim hOBBck As Long
    
    Dim hBMDst As Long
    Dim hDCDst As Long
    Dim hOBDst As Long
    
    Dim hBMPic As Long
    Dim hDCPic As Long
    Dim hOBPic As Long
    
    Dim hBMSrc As Long
    Dim hDCSrc As Long
    Dim hOBSrc As Long
    
    Dim brRGB As tRGB
    Dim drawH As Long
    Dim drawW As Long
    Dim hBrsh As Long
    Dim lMask As Long
    Dim pictH As Long
    Dim pictW As Long
    Dim tBtmp As BITMAPINFO ' Drawing information for the picture/icon
    Dim tCrop As POINTAPI   ' Points where to start copying/cropping the image
    Dim tPict As RECT       ' Area where to draw the picture
    
    Dim aPict() As RGBQUAD                      ' Picture color bits
    Dim aBack() As RGBQUAD                      ' Background color bits
                                                '
    CopyRect tPict, m_tButtonSettings.Picture   ' Get drawing area
                                                '
    If (Not MoveX = 0) Or (Not MoveY = 0) Then  '
        OffsetRect tPict, MoveX, MoveY          ' Move if set
    End If                                      '
    
    If (m_tButtonProperty.PicSize = epsNormal) Then
        ' Crop drawing area not visible in the button
        If (tPict.Left < 0) Then tCrop.x = -tPict.Left: tPict.Left = 0
        If (tPict.Top < 0) Then tCrop.Y = -tPict.Top: tPict.Top = 0
        If (tPict.bottom > m_tButtonSettings.Height) Then tPict.bottom = m_tButtonSettings.Height
        If (tPict.Right > m_tButtonSettings.Width) Then tPict.Right = m_tButtonSettings.Width
    End If
    
    drawH = tPict.bottom - tPict.Top            ' Draw height
    drawW = tPict.Right - tPict.Left            ' Draw width
                                                '
    If (drawH < 1) Or (drawW < 1) Then Exit Sub ' Nowhere to draw
                                                '
    pictH = ScaleY(Picture.Height, 8, 3)        ' Picture height
    pictW = ScaleX(Picture.Width, 8, 3)         ' Picture width
                                                '
    hDCSrc = CreateCompatibleDC(hdc)            ' Create drawing DC
                                                '
    If (Picture.type = 1) Or (Picture.type > 1 And Not m_tButtonProperty.UseMask) Then
        hOBSrc = SelectObject(hDCSrc, Picture.Handle)
    End If                                      '
                                                '
    If (m_tButtonProperty.UseMask) Then         ' Check if we can use the maskcolor set
        lMask = m_tButtonColors.MaskColor       '
    ElseIf (Picture.type > 1) Then              ' if not then check if we have an icon
        lMask = GetPixel(hDCSrc, 0, 0)          ' if it is then get top-left pixel color
        DeleteObject SelectObject(hDCSrc, hOBSrc)
    Else                                        '
        lMask = -1                              ' if it is a bitmap then use no maskcolor
    End If                                      '
    
    If (Picture.type > 1) Then
        hBMSrc = CreateCompatibleBitmap(hdc, pictW, pictH)
        hOBSrc = SelectObject(hDCSrc, hBMSrc)
        hBrsh = CreateSolidBrush(lMask)
        
        ' Fill transparent areas of the icon with the defined maskcolor
        DrawIconEx hDCSrc, 0, 0, Picture.Handle, pictW, pictH, 0, hBrsh, DI_NORMAL
        DeleteObject hBrsh
    End If
    
    hDCBck = CreateCompatibleDC(hDCSrc)
    hDCDst = CreateCompatibleDC(hDCSrc)
    hDCPic = CreateCompatibleDC(hDCSrc)
    
    hBMBck = CreateCompatibleBitmap(hdc, drawW, drawH)
    hBMDst = CreateCompatibleBitmap(hdc, drawW, drawH)
    hBMPic = CreateCompatibleBitmap(hdc, drawW, drawH)
    
    hOBBck = SelectObject(hDCBck, hBMBck)
    hOBDst = SelectObject(hDCDst, hBMDst)
    hOBPic = SelectObject(hDCPic, hBMPic)
    
    If (tCrop.x > 0 Or tPict.Right = m_tButtonSettings.Width) Then pictW = drawW + tCrop.x
    If (tCrop.Y > 0 Or tPict.bottom = m_tButtonSettings.Height) Then pictH = drawH + tCrop.Y
    
    ' Copy image to destination DC. Crop/resize if necessary.
    StretchBlt hDCDst, 0, 0, drawW, drawH, hDCSrc, tCrop.x, tCrop.Y, pictW - tCrop.x, pictH - tCrop.Y, SRCCOPY
    
    If (Not Picture.type = 1) Then ' vbPicTypeBitmap
        DeleteObject SelectObject(hDCSrc, hOBSrc)
    End If
    
    DeleteDC hDCSrc
    
    ReDim aBack(0 To drawW * drawH * 1.5) As RGBQUAD
    ReDim aPict(0 To UBound(aBack)) As RGBQUAD
    
    ' Get background & picture bitmap image
    BitBlt hDCBck, 0, 0, drawW, drawH, hdc, tPict.Left, tPict.Top, SRCCOPY
    BitBlt hDCPic, 0, 0, drawW, drawH, hDCDst, 0, 0, SRCCOPY
    
    With tBtmp.bmiHeader
        .biBitCount = 24 ' bit
        .biCompression = BI_RGB ' = 0
        .biHeight = drawH
        .biPlanes = 1
        .biSize = Len(tBtmp.bmiHeader)
        .biWidth = drawW
    End With
    
    ' Get background & picture color bits
    GetDIBits hDCBck, hBMBck, 0, drawH, aBack(0), tBtmp, DIB_RGB_COLORS
    GetDIBits hDCPic, hBMPic, 0, drawH, aPict(0), tBtmp, DIB_RGB_COLORS
    
    DeleteObject SelectObject(hDCBck, hOBBck)       ' Clear bitmap objects from memory
    DeleteObject SelectObject(hDCPic, hOBPic)       ' immediately after being used
                                                    '
    DeleteDC hDCBck                                 ' Clear device context instances
    DeleteDC hDCPic                                 ' from memory immediately
                                                    '
    If (BrushColor > -1) Then                       ' Determine brush color
        brRGB = GetRGB(BrushColor)                  ' used to replace colors on the
    End If                                          ' image
                                                    '
    Dim lOpacity As Long                            '
    If (m_tButtonSettings.state = ebsDisabled) Then ' For disabled buttons
        lOpacity = m_tButtonProperty.PicOpacity     ' We will just blend the picture
        m_tButtonProperty.PicOpacity = 0.2          ' On the button by 20%
    End If                                          '
    
    If (lMask = -1) And (BrushColor = -1) And (m_tButtonProperty.PicOpacity = 1) Then
        ' Skip bit by bit processing of image when not really necessary.
        ' Helps make things faster especially when loading large image/s
        GoTo Jmp_DrawImage
    End If
    
    Dim x As Long
    Dim Y As Long
    Dim Z As Long
    
    While (Y < drawH)
        x = 0
        While (x < drawW)
            ' GetNearestColor returns the actual value identifying a color from the
            ' system palette that will be displayed when the specified color is used
            
            If (GetNearestColor(hDCDst, RGBEx(aPict(Z).rgbRed, _
                                              aPict(Z).rgbGreen, _
                                              aPict(Z).rgbBlue)) = lMask) Then
                                                '
                aPict(Z) = aBack(Z)             ' Replace to background pixel color
                                                ' to make it look like transparent
            Else                                '
                If (BrushColor > -1) Then       '
                    aPict(Z).rgbRed = brRGB.r   ' Change all pixel color values
                    aPict(Z).rgbGreen = brRGB.G ' of an image with the specified
                    aPict(Z).rgbBlue = brRGB.b  ' brush color when set
                    
                ElseIf (Not m_tButtonProperty.PicOpacity = 1) Then
                    ' Results to an effect that blends the picture on the control
                    aPict(Z) = BlendRGBQUAD(aBack(Z), aPict(Z), m_tButtonProperty.PicOpacity)
                    
                End If
            End If
            
            x = x + 1
            Z = Z + 1 ' bit counter
        Wend
        
        Y = Y + 1
    Wend
    
Jmp_DrawImage:
    
    Erase aBack
    DeleteObject SelectObject(hDCDst, hOBDst)
    DeleteDC hDCDst
    
    SetDIBitsToDevice hdc, _
                      tPict.Left, _
                      tPict.Top, _
                      drawW, _
                      drawH, _
                      0, _
                      0, _
                      0, _
                      drawH, _
                      aPict(0), _
                      tBtmp, _
                      DIB_RGB_COLORS                ' Draw optimized image to button
    Erase aPict                                     '
                                                    ' Clear color bit arrays from memory
    If (m_tButtonSettings.state = ebsDisabled) Then '
        m_tButtonProperty.PicOpacity = lOpacity     ' Restore opacity
    End If                                          '
    
End Sub

Private Sub DrawIconEffect( _
        BackColor As Long, _
        Optional MoveX As Long, _
        Optional MoveY As Long, _
        Optional BrushColor As Long = -1)
'   Draw icon to button with defined special effect being applied.
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        Exit Sub
    End If
    
    Dim Picture As StdPicture
    
    If (m_tButtonSettings.state = ebsHot And Not m_tButtonProperty.PicHot Is Nothing) Then
        Set Picture = m_tButtonProperty.PicHot
    ElseIf (m_tButtonSettings.state = ebsDown) Or (m_tButtonProperty.Value) Then
        If (m_tButtonProperty.PicDown Is Nothing) Then
            Set Picture = m_tButtonProperty.PicHot
        Else
            Set Picture = m_tButtonProperty.PicDown
        End If
    End If
    
    If (Picture Is Nothing) Then
        Set Picture = m_tButtonProperty.PicNormal
    End If
    
    #If USE_SPECIALEFFECTS Then
    If (Not m_tButtonProperty.Effects = eseNone) Then
        Dim Color1 As Long
        Dim Color2 As Long
            Color1 = ShiftColor(BackColor, 0.1)
            Color2 = ShiftColor(BackColor, -0.1)
    End If
    
    Select Case m_tButtonProperty.Effects ' Draw effects first
        Case eseEmbossed
            DrawIcon Picture, MoveX + 1, MoveY + 1, Color2
            DrawIcon Picture, MoveX - 1, MoveY - 1, Color1
        Case eseEngraved
            DrawIcon Picture, MoveX - 1, MoveY - 1, Color2
            DrawIcon Picture, MoveX + 1, MoveY + 1, Color1
        Case eseShadowed
            DrawIcon Picture, MoveX + 1, MoveY + 1, Color2
    End Select
    #End If
    
    DrawIcon Picture, MoveX, MoveY, BrushColor
    
End Sub

Private Sub DrawLine(Color As Long, X1 As Long, Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
'   Draw a line with the specified color and coordinates
    Dim hOld As Long
    Dim hPen As Long
        hPen = CreatePen(PS_SOLID, 1, Color)
        hOld = SelectObject(hdc, hPen)
        
    If (X1 = X2) Then Y2 = Y2 + 1   ' LineTo draws a line up to, but not
    If (Y1 = Y2) Then X2 = X2 + 1   ' including the defined point. But not now!
                                    '
    ' If (X1 = X2) Then             '
    '     If (Y2 >= Y1) Then Y2 = Y2 + 1 Else Y2 = Y2 - 1
    ' End If                        '
    ' If (Y1 = Y2) Then             '
    '     If (X2 >= X1) Then X2 = X2 + 1 Else X2 = X2 - 1
    ' End If                        '
                                    '
    MoveToEx hdc, X1, Y1, 0&        ' Set starting position of line
    LineTo hdc, X2, Y2              ' then draw a line to this point
                                    '
    SelectObject hdc, hOld          ' Restore the previous pen
    DeleteObject hPen               ' then clear pen object from memory
    
End Sub

#If USE_CRYSTAL Or USE_MAC Or USE_MACOSX Then
Private Sub DrawShineEffect(StartColor As Long, EndColor As Long)
'   Draw shine effect to button.
'   Before calling this procedure, the shine object (FocusRect coordinates) must
'   already be in position and size as it should appear on the control...
    
    Dim H As Long
    Dim W As Long
    Dim x As Long
    Dim Y As Long
    
    With m_tButtonSettings.Focus
        H = .bottom - .Top + 1
        W = .Right - .Left + 1
        x = .Left
        Y = .Top
                        ' Bug fixed: 07/28/06
        If H = 0 Then   ' Error on ReDim aGrad(0 To H - 1) As Long
            Exit Sub    ' Does not occur on my system no matter what I do
        End If          ' But some says it does occur on them, so we get rid of it
        
        If (m_tButtonProperty.Shape = ebsCutLeft) Or _
           (m_tButtonProperty.Shape = ebsCutSides) Then
            W = W + x
            x = 0
        End If
        
        If (m_tButtonProperty.Shape = ebsCutRight) Or _
           (m_tButtonProperty.Shape = ebsCutSides) Then
            W = W + .Left
        End If
    End With
    
    Dim A As Long
    Dim b As Long
    Dim c As Single
    
    Dim aGrad() As Long
    ReDim aGrad(0 To H - 1) As Long
    
    Dim tRGB1 As tRGB
    Dim tRGB2 As tRGB
    Dim tRGB3 As tRGB
        tRGB1 = GetRGB(StartColor)
        tRGB2 = GetRGB(EndColor)
        
    While (A < H)                                   ' Calculate gradient color bits
        c = A / H                                   ' Calculate color step value
        tRGB3.r = tRGB1.r + (tRGB2.r - tRGB1.r) * c '
        tRGB3.G = tRGB1.G + (tRGB2.G - tRGB1.G) * c '
        tRGB3.b = tRGB1.b + (tRGB2.b - tRGB1.b) * c '
        aGrad(A) = RGBEx(tRGB3.r, tRGB3.G, tRGB3.b) ' Get gradient color value
        A = A + 1                                   ' Get next color
    Wend                                            '
                                                    '
    Dim hRgn  As Long                               '
    Dim hRgn2 As Long                               '
        hRgn = CreateRoundRectRgn(0, 0, W, H, H, W) ' Create region(s)
                                                    '
        If (m_tButtonProperty.Shape = ebsCutLeft) Or _
           (m_tButtonProperty.Shape = ebsCutSides) Then
            hRgn2 = CreateRoundRectRgn(0, 0, W / 2, H, 0, 0)
            CombineRgn hRgn, hRgn, hRgn2, RGN_OR    '
            DeleteObject hRgn2                      '
        End If                                      '
                                                    '
        If (m_tButtonProperty.Shape = ebsCutRight) Or _
           (m_tButtonProperty.Shape = ebsCutSides) Then
            hRgn2 = CreateRoundRectRgn(W / 2, 0, W, H, 0, 0)
            CombineRgn hRgn, hRgn, hRgn2, RGN_OR    '
            DeleteObject hRgn2                      '
        End If                                      '
                                                    '
    Dim X1 As Long                                  '
    Dim X2 As Long                                  '
                                                    '
    While (b < H)                                   ' Draw from top to bottom
        A = 0                                       ' X at 0
        X1 = -1                                     '
        X2 = -1                                     '
                                                    '
        While (A < W)                               ' From left to right
            If (PtInRegion(hRgn, A, b)) Then        '
                If (X1 = -1) Then X1 = A            ' Set start point
            ElseIf (Not X1 = -1) And (X2 = -1) Then '
                X2 = A                              ' Set end point
                A = W ' Exit While                  '
            End If                                  '
            A = A + 1                               '
        Wend                                        '
        If (Not X1 = -1) And (Not X2 = -1) Then     ' Are points set?
            DrawLine aGrad(b), x + X1, Y + b, x + X2, Y + b
        End If                                      '
                                                    '
        b = b + 1                                   ' Draw gradient lines
    Wend                                            '
                                                    '
    DeleteObject hRgn                               ' Delete region object
    
End Sub
#End If

Private Sub FillButtonEx(Color As Long)
'   Fill the control with the specified color.
    Dim hOldBr As Long
    Dim hBrush As Long
    Dim lpRect As RECT
    
    hBrush = CreateSolidBrush(Color)    ' Create new brush with the specified color
    hOldBr = SelectObject(hdc, hBrush)  ' Use the new brush but save previous one
                                        '
    SetRect lpRect, 0, 0, m_tButtonSettings.Width, m_tButtonSettings.Height
                                        '
    FillRect hdc, lpRect, hBrush        '
                                        '
    SelectObject hdc, hOldBr            ' Restore the previous brush
    DeleteObject hBrush                 ' Remove brush instance from memory
    
End Sub

Private Function GetAccessKey(Caption As String) As String
'   Get accesskey from a caption
    Dim iMax As Integer
    Dim iPos As Integer
    Dim sChr As String * 1 ' Helps conserve memory I guess
    
    iMax = Len(Caption)
    
    ' Of course you can't assign an accesskey
    ' with less than 2 characters
    If (iMax < 2) Then Exit Function
    
    iMax = iMax - 1 ' Start from the second to the last character
    
    Do  ' An accesskey is found after the last
        ' ampersand character found on the string
        iPos = InStrRev(Caption, "&", iMax)
        
        If (iPos = 0) Then Exit Do
        
        If (iPos = 1) Then
            GetAccessKey = Mid$(Caption, iPos + 1, 1)
            Exit Do
        Else
            ' Check if the character before the
            ' ampersand is also an ampersand
            sChr = Mid$(Caption, iPos - 1, 1)
            
            ' A series of two ampersand characters will draw
            ' an ampersand character as part of the caption
            ' and will not be considered as an accesskey
            If (Not StrComp(sChr, "&") = 0) Then
                GetAccessKey = Mid$(Caption, iPos + 1, 1)
                Exit Do
            End If
            
            iMax = iPos - 2 ' Find another character
        End If
        
    Loop While (iMax > 0)
    
    GetAccessKey = LCase$(GetAccessKey)
    
End Function

Private Function GetRGB(Color As Long) As tRGB
'   Returns the RGB color value of the specified color.
    GetRGB.r = Color And 255
    GetRGB.G = (Color \ 256) And 255
    GetRGB.b = (Color \ 65536) And 255
    
End Function

Private Function IsFunctionSupported(sFunction As String, sModule As String) As Boolean
'   Determines if the passed function is supported by a library
    Dim hModule As Long
        hModule = GetModuleHandleA(sModule) ' GetModuleHandle
        
    If (hModule = 0) Then
        hModule = LoadLibrary(sModule)
    End If
    
    If (hModule) Then
        If (GetProcAddress(hModule, sFunction)) Then
            IsFunctionSupported = True
        End If
        FreeLibrary hModule
    End If
    
End Function

Private Function IsPlatformNT() As Boolean
'   Determines if the system currently running the program has an NT platform.
    Dim OSINFO As OSVERSIONINFO
        OSINFO.dwOSVersionInfoSize = Len(OSINFO)
        
    If (GetVersionEx(OSINFO)) Then
        IsPlatformNT = (OSINFO.dwPlatformId = VER_PLATFORM_WIN32_NT)
    End If
    
End Function

#If USE_CRYSTAL Or USE_MAC Or USE_MACOSX Then
Private Sub PolylineEx(Points() As POINTAPI, Color As Long)
'   Draws a series of line segments from the specified array of points.
    Dim hOld As Long
    Dim hPen As Long
        hPen = CreatePen(PS_SOLID, 1, Color)        ' Create new pen
        hOld = SelectObject(hdc, hPen)              ' Use new pen
                                                    '
        Polyline hdc, Points(0), UBound(Points) + 1 ' Draw lines
                                                    '
        SelectObject hdc, hOld                      ' Restore the previous pen
        DeleteObject hPen                           ' Remove new pen from memory
        
End Sub
#End If

Private Sub RectangleEx(Color As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
'   Draws a rectangle using the specified color and coordinates
    Dim hOld As Long
    Dim hPen As Long
        hPen = CreatePen(PS_SOLID, 1, Color)    ' Create new pen
        hOld = SelectObject(hdc, hPen)          ' Use new pen
                                                '
        Rectangle hdc, X1, Y1, X2, Y2           ' Draw shape
                                                '
        SelectObject hdc, hOld                  ' Restore previous pen
        DeleteObject hPen                       ' Remove new pen from memory
    
End Sub

Private Function RGBEx( _
        ByVal Red As Long, _
        ByVal Green As Long, _
        ByVal Blue As Long) As Long
'   Returns a whole number representing an RGB color value.
    RGBEx = Red + 256& * Green + 65536 * Blue
    
End Function

Private Sub SetButtonColors()
'   Set button colors and translate to vb safe colors
    
    With m_tButtonColors ' Common to all styles
        .BackColor = TranslateColor(m_tButtonProperty.BackColor)
        .ForeColor = TranslateColor(m_tButtonProperty.ForeColor)
        .GrayText = BlendColors(.ForeColor, &HFFFFFF, 0.6)
        .MaskColor = TranslateColor(m_tButtonProperty.MaskColor)
        .StartColor = &HFFFFFF ' vbWhite
        
        Select Case m_tButtonProperty.Style
            #If USE_CRYSTAL Then
            Case ebsCrystal
                .DownColor = ShiftColor(.BackColor, -0.15)
                .HoverColor = ShiftColor(.BackColor, 0.1)
                .BackColor = ShiftColor(.BackColor, -0.05)
                .GrayColor = .BackColor
                .StartColor = -1
            #End If
            #If USE_MAC Then
            Case ebsMac
                .DownColor = ShiftColor(.BackColor, -0.1)
                .GrayColor = .BackColor
                .HoverColor = ShiftColor(.BackColor, 0.1)
                .StartColor = -1
            #End If
            #If USE_MACOSX Then
            Case ebsMacOSX
                .DownColor = ShiftColor(.BackColor, -0.15)
                .HoverColor = ShiftColor(.BackColor, 0.1)
                .BackColor = ShiftColor(.BackColor, -0.05)
                .GrayColor = .BackColor
            #End If
            #If USE_OFFICE2003 Then
            Case ebsOffice2003
                .DownColor = &H4E91FE
                .HoverColor = &H8BCFFF ' Endcolor for Hot & Down state gradient
                .FocusBorder = &HCCF4FF ' Startcolor for Hot state gradient
            #End If
            #If USE_OFFICEXP Then
            Case ebsOfficeXP
                .DownColor = ShiftColor(.BackColor, -0.05)
                .FocusBorder = ShiftColor(.BackColor, -0.15)
                .GrayColor = .BackColor
                .HoverColor = ShiftColor(.BackColor, 0.05)
            #End If
            #If USE_OPERABROWSER Then
            Case ebsOperaBrowser
                .BackColor = .BackColor
                .GrayColor = BlendColors(.BackColor, &HFFFFFF, 0.3)
                .HoverColor = &HE6FFFF ' Background color on hover & down state
                .StartColor = ShiftColor(.BackColor, 0.1)
            #End If
            #If USE_STANDARD Then
            Case ebsStandard
                .DownColor = .BackColor ' For the checkered background pattern
                .FocusBorder = ShiftColor(.ForeColor, 0.5)
            #End If
            #If USE_XPBLUE Then
            Case ebsXPBlue
                .DownColor = ShiftColor(.BackColor, -0.05)
                .FocusBorder = &HE4AD89
                .GrayColor = BlendColors(.BackColor, &HFFFFFF, 0.8)
                .HoverColor = &H30B3F8
            #End If
            #If USE_XPOLIVEGREEN Then
            Case ebsXPOliveGreen
                .DownColor = ShiftColor(.BackColor, -0.05)
                .FocusBorder = &H54C190
                .GrayColor = BlendColors(.BackColor, &HFFFFFF, 0.8)
                .HoverColor = &H4F91E3
            #End If
            #If USE_XPSILVER Then
            Case ebsXPSilver
                .DownColor = ShiftColor(.BackColor, -0.1)
                .FocusBorder = &HE4AD89
                .GrayColor = BlendColors(.BackColor, &HFFFFFF, 0.8)
                .HoverColor = &H30B3F8
            #End If
            #If USE_XPTOOLBAR Then
            Case ebsXPToolbar
                .GrayColor = .BackColor
                .HoverColor = &HF3F7F8
                .DownColor = ShiftColor(.HoverColor, -0.05)
                ' .StartColor ' ForeColor for down state
            #End If
            #If USE_YAHOO Then
            Case ebsYahoo
                .DownColor = ShiftColor(.BackColor, -0.1)
                .FocusBorder = .BackColor
                .HoverColor = ShiftColor(.BackColor, 0.2)
                .GrayColor = BlendColors(.BackColor, &HFFFFFF, 0.8)
            #End If
        End Select
    End With
    
End Sub

Private Function ShiftColor(Color As Long, PercentInDecimal As Single) As Long
'   Add or remove a certain color quantity by how many percent.
    Dim RGB1 As tRGB
        RGB1 = GetRGB(Color)
        
        RGB1.r = RGB1.r + PercentInDecimal * 255 ' Percent should already
        RGB1.G = RGB1.G + PercentInDecimal * 255 ' be translated.
        RGB1.b = RGB1.b + PercentInDecimal * 255 ' Ex. 50% -> 50 / 100 = 0.5
        
    If (PercentInDecimal > 0) Then ' RGB values must be between 0-255 only
        If (RGB1.r > 255) Then RGB1.r = 255
        If (RGB1.G > 255) Then RGB1.G = 255
        If (RGB1.b > 255) Then RGB1.b = 255
    Else
        If (RGB1.r < 0) Then RGB1.r = 0
        If (RGB1.G < 0) Then RGB1.G = 0
        If (RGB1.b < 0) Then RGB1.b = 0
    End If
    
    ShiftColor = RGBEx(RGB1.r, RGB1.G, RGB1.b) ' Return shifted color value
    
End Function

#If USE_POPUPMENU Then
Private Sub ShowPopupMenu()
'   Displays a pop-up menu using the settings specified when set.
'   TrackPopupMenu flag/contant not included in VBs MenuControlContants Enum
    Const TPM_BOTTOMALIGN As Long = &H20&
    
    Dim Menu        As VB.Menu          ' Declarations
    Dim Align       As eMenuAlignments  '
    Dim flags       As Long             '
    Dim DefaultMenu As VB.Menu          '
                                        '
    With m_tPopupSettings               ' Get settings
        Set Menu = .Menu                '
            Align = .Align              '
            flags = .flags              '
        Set DefaultMenu = .DefaultMenu  '
    End With                            '
                                        '
    Dim x As Long                       '
    Dim Y As Long                       '
                                        '
    m_bPopupInit = True                 ' This flag prevents the mouseleave event
                                        ' to redraw the button
    Select Case Align                   '
        Case emaBottom                  '
            Y = m_tButtonSettings.Height
                                        '
        Case emaLeft, emaLeftBottom     '
            flags = flags Or vbPopupMenuRightAlign
                                        '
            If (Align = emaLeftBottom) Then
                Y = m_tButtonSettings.Height
            End If                      '
                                        '
        Case emaRight, emaRightBottom   '
            x = m_tButtonSettings.Width '
                                        '
            If (Align = emaRightBottom) Then
                Y = m_tButtonSettings.Height
            End If                      '
                                        '
        Case emaTop, emaTopRight, emaTopLeft
            flags = flags Or TPM_BOTTOMALIGN
                                        '
            If (Align = emaTopRight) Then
                x = m_tButtonSettings.Width
            ElseIf (Align = emaTopLeft) Then
                flags = flags Or vbPopupMenuRightAlign
            End If                      '
                                        '
        Case Else                       '
            m_bPopupInit = False        ' No popup menu will be shown
                                        '
    End Select                          '
                                        '
    If (m_bPopupInit) Then              ' Show popup menu while drawing
        DrawButton ebsDown              ' appropriate button state
                                        '
        If (DefaultMenu Is Nothing) Then
            UserControl.PopupMenu Menu, flags, x, Y
        Else                            '
            UserControl.PopupMenu Menu, flags, x, Y, DefaultMenu
        End If                          '
                                        ' The following instructions are executed
                                        ' after the popup menu has dismissed
        Dim lpPoint As POINTAPI         '
            GetCursorPos lpPoint        ' Get current cursor position
                                        '
        If (WindowFromPoint(lpPoint.x, lpPoint.Y) = UserControl.hWnd) Then
            m_bPopupShown = True        ' Mouse events will handle this later
        Else                            '
            m_bButtonIsDown = False     ' Restore button state to normal
            m_bMouseIsDown = False      ' when cursor is outside the button
            m_bMouseOnButton = False    ' while resetting other properties
            m_bSpacebarIsDown = False   ' and popup menu flags
            DrawButton ebsNormal        ' to allow control from responding
            m_bPopupInit = False        ' accurately to other events
            m_bPopupShown = False       ' such as mouse events
        End If                          '
    End If                              '
    
End Sub
#End If

Private Sub TrackMouseTracking(hWnd As Long)
'   Start tracking of mouse leave event
    Dim lpEventTrack As TRACKMOUSEEVENTTYPE
    
    With lpEventTrack
        .cbSize = Len(lpEventTrack)
        .dwFlags = TME_LEAVE
        .hwndTrack = hWnd
    End With
    
    If (m_bTrackHandler32) Then
        TrackMouseEvent lpEventTrack
    Else
        TrackMouseEvent2 lpEventTrack
    End If
    
End Sub

Private Function TranslateColor(Value As Long) As Long
'   Converts system color constants to real color values
    OleTranslateColor Value, 0&, TranslateColor
    
End Function

Private Function TranslateNumber( _
        ByVal Value As Single, _
        Minimum As Long, _
        Maximum As Long) As Single
'   Ensure a number does not exceed the specified limits.
    
    If (Value > Maximum) Then
        TranslateNumber = Maximum
    ElseIf (Value < Minimum) Then
        TranslateNumber = Minimum
    Else
        TranslateNumber = Value
    End If
    
End Function

' //-- Subclassing Procedure --//

Private Sub Subclass_Proc( _
        ByVal bBefore As Boolean, _
        ByRef bHandled As Boolean, _
        ByRef lReturn As Long, _
        ByVal lng_hWnd As Long, _
        ByVal uMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long, _
        ByRef lParamUser As Long)
        
    Select Case uMsg
        Case WM_MOUSELEAVE
            ' Triggered as the cursor moved out the button. If mouse button is held
            ' down, is it triggered when the button is released outside the button
            
            m_bMouseOnButton = False
            m_bIsTracking = False
            
            #If DEBUG_EVENTS Then
                Debug.Print Caption & " - UserControl_MouseLeave (Subclass)"
            #End If
            
            #If USE_POPUPMENU Then
            If (m_bPopupEnabled) Then       '
                If (m_bPopupInit) Then      ' Retain the down state of the button
                    m_bPopupInit = False    '
                    m_bPopupShown = True    '
                    Exit Sub                ' Drawing will be handled after the
                Else                        ' popup menu closes/exits
                    m_bPopupShown = False   '
                End If                      '
            End If                          '
            #End If
            
            If (Not m_bSpacebarIsDown) Then
                ' Force update if mouse button was held down on leave
                ' Neccessary for new button styles such as Office/XP styles
                If (m_tButtonProperty.Enabled) Then
                    DrawButton ebsNormal, (m_tButtonSettings.Button = 1)
                End If
                ' Bug fixed: When ENABLED is set to FALSE from its own event
                ' Still raise the MouseLeave event for the user
                RaiseEvent MouseLeave
            End If
            
        Case WM_ACTIVATE, WM_NCACTIVATE ' Parent form is activated or deactivated
            m_bParentActive = (Not wParam = 0)
            
            #If DEBUG_EVENTS Then
                Debug.Print Caption & " - UserControl_ParentActive", m_bParentActive
            #End If
            
            If (m_bParentActive) Then  ' Activated
                If (m_tButtonProperty.Enabled) Then
                    ' Force update if set to default or was on focus
                    If (m_bButtonHasFocus Or m_tButtonSettings.Default) Then
                        DrawButton ebsNormal, True
                    End If
                End If
            Else ' Deactivated
                Dim bFocus As Boolean
                Dim bForce As Boolean
                
                bFocus = m_bButtonHasFocus
                bForce = m_bButtonHasFocus Or m_tButtonSettings.Default Or m_bMouseOnButton
                
                m_bButtonHasFocus = False   ' Unset runtime settings
                m_bButtonIsDown = False     ' necessary to effectively
                m_bMouseIsDown = False      ' draw a normal button
                m_bMouseOnButton = False    '
                m_bSpacebarIsDown = False   '
                
                m_tButtonSettings.Default = False ' Temporary cancel DisplayAsDefault
                
                If (m_tButtonProperty.Enabled) Then
                    DrawButton ebsNormal, bForce
                End If
                
                ' Restore neccesary settings used when parent form is reactivated again
                m_tButtonSettings.Default = Ambient.DisplayAsDefault
                m_bButtonHasFocus = bFocus
            End If
    End Select
    
End Sub

' //-- UserControl Procedures --//

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
'   Triggered if the accesskey is pressed (Alt + underlined character of caption)
'   Also on ENTER key or ESCAPE if Cancel property is set to True
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_AccessKeyPress", KeyAscii
    #End If
    
    If (m_tButtonProperty.Enabled) Then
        If (m_bSpacebarIsDown) Then
            If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
                ReleaseCapture                      ' input processing
            End If                                  ' of the window
        End If
        
        If (m_tButtonProperty.CheckBox) Then
            ' Checkboxes does not respond to Enter & Escape keys
            If (KeyAscii = 13) Or (KeyAscii = 27) Then ' vbKeyReturn, vbKeyEscape
                Exit Sub
            End If
        End If
        
        m_bButtonIsDown = False         ' Release button
        m_tButtonSettings.Button = 1    '
                                        '
        #If USE_POPUPMENU Then          '
        If (Not m_bPopupEnabled) Then   '
            UserControl_Click           ' Trigger click event
        ElseIf (Not KeyAscii = 27) Then ' Escape not accepted
            If (Not m_bPopupShown) Then '
                ShowPopupMenu           ' Show popup menu on accesskey and enter
                m_bButtonIsDown = False '
                m_bMouseIsDown = False  '
                m_bPopupInit = False    '
                m_bSpacebarIsDown = False
            End If                      '
        End If                          '
        #Else                           '
            UserControl_Click           ' Trigger click event
        #End If                         '
    End If
    
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'   Usually triggered as the user changes focus on different controls on the window
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_AmbientChanged", PropertyName, IIf(StrComp(PropertyName, "DisplayAsDefault") = 0, Ambient.DisplayAsDefault, "")
    #End If
    
    ' DisplayAsDefault returns True if the current control
    ' on focus is not another button or this button itself.
    
    m_tButtonSettings.Default = Ambient.DisplayAsDefault
    
    If (StrComp(PropertyName, "DisplayAsDefault") = 0) Then
        m_bButtonIsDown = False
        m_bMouseIsDown = False
        m_bSpacebarIsDown = False
        
        ' Prevent unneccessary drawing updates
        If (m_tButtonProperty.Enabled) And (Not m_bMouseOnButton) Then
            ' GotFocus event will just update the button display later
            If (Not m_bButtonHasFocus) Then
                DrawButton Force:=True
            End If
        End If
    End If
    
End Sub

Private Sub UserControl_Click()
'   Triggered normally when user clicks the control and release it inside the button
    If (m_bButtonIsDown) Or (Not m_tButtonSettings.Button = 1) Then Exit Sub
        m_bMouseIsDown = False
        m_bSpacebarIsDown = False
    
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_Click", IIf(m_tButtonProperty.CheckBox, "Checked=" & Not m_tButtonProperty.Value, "")
    #End If
    
    With m_tButtonProperty                  '
        If (.CheckBox) Then                 ' Check if checkbox mode is on
            .Value = Not .Value             ' If so, then toggle button value
        End If                              '
                                            '
        If (Not m_bMouseOnButton) Then      ' Check if cursor is over the control
            DrawButton ebsNormal, True      ' Redraw is necessary if it is not :)
        Else                                '
            DrawButton ebsHot, .CheckBox    ' Force redraw for checkbox mode
        End If                              '
                                            '
        #If USE_POPUPMENU Then              '
        If (Not m_bPopupEnabled) Then       ' Raise no event if control is set to
            RaiseEvent Click                ' handle a popup menu
        End If                              '
        #Else                               '
            RaiseEvent Click                '
        #End If                             '
                                            '
        If (Not .CheckBox) Then             ' Sometimes VALUE property is set to True
            .Value = False                  ' to trigger the click event
        End If                              ' So we should reset value if set
    End With                                '
    
End Sub

Private Sub UserControl_DblClick()
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_DblClick"
    #End If
    
    If (m_tButtonProperty.HandPointer) Then
        SetCursor m_tButtonSettings.Cursor ' Set hand cursor
    End If
    
    ' Draw a down button state which helps to emulate multiple clicks
    If (m_tButtonSettings.Button = 1) Then
        m_bButtonIsDown = True          ' Bug fixed: 07/27/06
        m_bMouseIsDown = True           ' Draw HOT state on dblclick, hold then mousemove
        DrawButton ebsDown              '
        m_tButtonSettings.Button = 8    ' Just a double click flag for the MouseUp event
                                        '
        If (Not GetCapture = UserControl.hWnd) Then
            SetCapture UserControl.hWnd ' Send MouseUp event to the control
        End If                          '
                                        '
        #If USE_POPUPMENU Then          '
        If (Not m_bPopupEnabled) Then   ' Send no DblClick event when
            RaiseEvent DblClick         ' control is set to handle popup menus
        ElseIf (Not m_bPopupShown) Then '
            ShowPopupMenu               ' Show menu
        Else                            '
            m_bPopupShown = False       ' Just close the active menu displayed
        End If                          '
        #Else                           '
            RaiseEvent DblClick         '
        #End If                         '
    End If                              '
    
End Sub

Private Sub UserControl_GotFocus()
'   Only raised when last focused control is on the same window but not itself
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_GotFocus"
    #End If
    
    m_bButtonHasFocus = True
    
    If (Not m_bButtonIsDown) Then
        DrawButton ebsNormal
    End If
    
End Sub

Private Sub UserControl_Hide()
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_Hide"
    #End If
    
    m_bControlHidden = True
    
End Sub

Private Sub UserControl_Initialize()
'   Called on design and run-time; when the form is getting ready for display
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_Initialize"
    #End If
    
    m_bIsPlatformNT = IsPlatformNT() ' Needed for the unicode text support
    m_bRedrawOnResize = False
    
End Sub

Private Sub UserControl_InitProperties()
'   Called on design time only; everytime this control is added on the form
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_InitProperties"
    #End If
    
    With m_tButtonProperty
        .Caption = Ambient.DisplayName
        .CheckBox = False
        #If USE_SPECIALEFFECTS Then
        .Effects = 0
        #End If
        .Enabled = True
         UserControl.Font = Ambient.Font
        .ForeColor = Ambient.ForeColor
        .MaskColor = &HC0C0C0
        .PicAlign = epaLeftOfCaption
    Set .PicDown = Nothing
    Set .PicHot = Nothing
    Set .PicNormal = Nothing
        .PicOpacity = 1
        .PicSize = epsNormal
        #If USE_STANDARD Then
        .Style = ebsStandard
        #Else
        .Style = 0
        #End If
        .UseMask = True
        .Value = False
    End With
    
    ColorScheme NoRedraw:=True ' Prevent redraw, resize event will redraw later
    SetButtonColors
    
    m_bRedrawOnResize = True
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_KeyDown", KeyCode, Shift
    #End If
    
    Select Case KeyCode
        Case 32 ' vbKeySpace
            ' Some buttons seems to forget this thing, and it sucks me off!
            ' I find Alt+Space to be so useful especially in closing forms.
            If (Not Shift = 4) Then ' vbAltMask
                m_bButtonIsDown = True
                m_bSpacebarIsDown = True
                
                If (Not GetCapture = UserControl.hWnd) Then
                    SetCapture UserControl.hWnd ' Restrict user from selecting other
                End If                          ' controls while spacebar is held down
                                                '
                If (Not m_bMouseIsDown) Then    ' If mouse is up only
                    DrawButton ebsDown
                End If
            End If
            
            ' Normally, this event is raised before drawing the
            ' pressed button state when spacebar is pressed.
            ' I moved it here to draw the down state of the button
            ' before the user can add some additional commands.
            RaiseEvent KeyDown(KeyCode, Shift)
            
        Case 37, 38, 39, 40 ' vbKeyLeft, vbKeyUp, vbKeyRight, vbKeyDown
            If (Shift = 0) Then
                If (KeyCode = 37) Or (KeyCode = 38) Then ' vbKeyLeft, vbKeyUp
                    SendKeys "+{TAB}"
                Else
                    SendKeys "{TAB}"
                End If
            End If
            
            If (m_bSpacebarIsDown) Then
                If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
                    ReleaseCapture                      ' input processing
                End If                                  ' of the window
                
                ' If SPACEBAR is down then either of the arrow keys is pressed
                ' Process the arrow event first to transfer focus to the next
                ' available control then trigger the click event after
                DoEvents
                m_tButtonSettings.Button = 1
                UserControl_Click
            End If
            
        Case Else
            ' If spacebar is held down, then a key not included above is pressed
            ' should simulate a release to the button to an appropriate state.
            If (m_bSpacebarIsDown) Then
                m_bButtonIsDown = False
                m_bSpacebarIsDown = False
                
                If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
                    ReleaseCapture                      ' input processing
                End If                                  ' of the window
                
                If (m_bMouseOnButton) Then
                    DrawButton ebsHot
                Else
                    DrawButton ebsNormal
                End If
            End If
            
            RaiseEvent KeyDown(KeyCode, Shift)
    End Select
    
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_KeyPress", KeyAscii
    #End If
    
    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_KeyUp", KeyCode, Shift
    #End If
    
    Select Case KeyCode
        Case 32 ' vbKeySpace
            m_bButtonIsDown = m_bMouseIsDown
            
            If (m_bSpacebarIsDown) Then                 '
                m_bSpacebarIsDown = False               ' Space has been released
                                                        '
                #If USE_POPUPMENU Then                  '
                If (Not m_bPopupEnabled) Then           ' No event on popup menu
                    m_tButtonSettings.Button = 1        '
                    UserControl_Click                   ' Simulate a click event
                ElseIf (Not m_bPopupShown) Then         '
                    ShowPopupMenu                       ' Show popup menu on
                    m_bPopupInit = False                ' release of spacebar
                End If                                  '
                #Else                                   '
                    m_tButtonSettings.Button = 1        '
                    UserControl_Click                   '
                #End If                                 '
            End If                                      '
                                                        '
            If (m_bButtonIsDown) Then                   ' Raise Mouse_Up event
                If (Not GetCapture = UserControl.hWnd) Then
                    SetCapture UserControl.hWnd         ' When the left-mouse button
                End If                                  ' will be finally released
            Else                                        '
                If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
                    ReleaseCapture                      ' input processing
                End If                                  ' of the window
            End If                                      '
                                                        '
            #If USE_POPUPMENU Then                      '
            If (Not m_bPopupEnabled) Then               ' Raise no event when a
                RaiseEvent KeyUp(KeyCode, Shift)        ' popup menu has been displayed
            End If                                      '
            #Else                                       '
                RaiseEvent KeyUp(KeyCode, Shift)        '
            #End If                                     '
        Case Else
            RaiseEvent KeyUp(KeyCode, Shift)
    End Select
    
End Sub

Private Sub UserControl_LostFocus()
'   Only raised when another control of the same window is focused
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_LostFocus"
    #End If
    
    m_bButtonHasFocus = False           '
    m_bButtonIsDown = False             '
    m_bMouseIsDown = False              ' Release buttons
    m_bSpacebarIsDown = False           '
                                        '
    If (m_tButtonProperty.Enabled) Then '
        If (m_bParentActive) Then       ' We need to force redraw
            DrawButton ebsNormal, True  ' button on lost of control focus
        End If                          ' only when parent window is active
    End If                              ' the other way is handled by the subclass proc
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_MouseDown", Button, Shift
    #End If
    
    If (m_tButtonProperty.HandPointer) Then '
        SetCursor m_tButtonSettings.Cursor  ' Set hand cursor
    End If                                  '
                                            '
    m_tButtonSettings.Button = Button       ' Cache button to trigger double click event
                                            '
    If (Button = 1) Then                    ' vbLeftButton
        m_bButtonHasFocus = True            '
        m_bButtonIsDown = True              '
        m_bMouseIsDown = True               '
                                            '
        If (Not m_bSpacebarIsDown) Then     '
            DrawButton ebsDown              '
        ElseIf (Not m_bMouseOnButton) Then  ' If mouse button is pressed outside
            m_bMouseIsDown = False          ' the control while the spacebar is
            DrawButton ebsNormal            ' being held down then draw hot state
            m_bMouseIsDown = True           ' Unset/set m_bMouseIsDown to trick
        End If                              ' drawing procedures about the state
    End If                                  '
                                            '
    #If USE_POPUPMENU Then                  '
    If (Not m_bPopupEnabled) Then           ' No mouse event when popup is enabled
        RaiseEvent MouseDown(Button, Shift, x, Y)
    ElseIf (Not m_bPopupShown) Then         '
        ShowPopupMenu                       ' Show menu
    Else                                    '
        m_bPopupInit = False
        m_bPopupShown = False               ' Just close the active menu
    End If                                  '
    #Else                                   '
        RaiseEvent MouseDown(Button, Shift, x, Y)
    #End If                                 '
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Parent form is inactive
    If (m_tButtonProperty.HandPointer) Then
        SetCursor m_tButtonSettings.Cursor ' Set hand cursor
    Else ' Ensure cursor is what we expected
        UserControl.MousePointer = UserControl.MousePointer
    End If
    
    ' Restrict HOT state if parent window/form is not in focus
    ' If (Not m_bParentActive) Then Exit Sub
    ' Not now, I allow MouseOver event to show HOT state
    ' even if parent window is not in focus
    
    Dim lpPoint As POINTAPI
        GetCursorPos lpPoint
        
    If (Not WindowFromPoint(lpPoint.x, lpPoint.Y) = UserControl.hWnd) Then
        ' It reaches here if the left mouse button is held down
        ' as the user moves the cursor off the control
        If (m_bMouseOnButton) Then
            m_bMouseOnButton = False
            
            #If DEBUG_EVENTS Then
                Debug.Print Caption & " - UserControl_MouseLeave"
            #End If
            
            If (Not m_bSpacebarIsDown And Not m_bMouseIsDown) Then ' Retain down state
                DrawButton ebsNormal
            End If
            
            RaiseEvent MouseLeave
        ElseIf (m_bMouseIsDown) Then
            DrawButton ebsHot
        End If
        
        #If USE_POPUPMENU Then
            m_bPopupShown = False
        #End If
    Else
        #If DEBUG_EVENTS Then ' Must be before drawing the button
            If (Not m_bIsTracking) Or (Not m_bMouseOnButton) Then
                Debug.Print Caption & " - UserControl_MouseEnter"
            End If
        #End If
        
        m_bMouseOnButton = True
        
        If (Not m_bSpacebarIsDown) Then ' Check if spacebar is held down
            If (m_bButtonIsDown) Then   ' If not, draw appropriate button state
                DrawButton ebsDown      '
            Else                        ' If button should be in down state then
                DrawButton ebsHot       ' draw the down state else draw the hot state
            End If                      '
        ElseIf (m_bMouseIsDown) Then    ' Else, check if mouse is held down
            DrawButton ebsDown          ' as it moves over the button to
        End If                          ' draw the down state
                                        '
        If (Not m_bIsTracking) Then     ' Trigger MouseEnter event the first time
                m_bIsTracking = True    ' the cursor has moved on the control
            
            TrackMouseTracking UserControl.hWnd
            RaiseEvent MouseEnter
        Else
            #If DEBUG_EVENTS Then ' Must be before drawing the button
                Debug.Print Caption & " - UserControl_MouseMove"; x; Y
            #End If
            
            ' Succeeding move events will trigger the MouseMove event
            RaiseEvent MouseMove(Button, Shift, x, Y)
        End If
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_MouseUp", Button, Shift
    #End If
    
    If (m_tButtonProperty.HandPointer) And (m_bMouseOnButton) Then
        SetCursor m_tButtonSettings.Cursor ' Set hand cursor
    End If
    
    #If USE_POPUPMENU Then
    If (m_bPopupEnabled) Or (Button = 1 Or m_bPopupShown) Then ' vbLeftButton
    #Else
    If (Button = 1) Then ' vbLeftButton
    #End If
        m_bMouseIsDown = False
        m_bButtonIsDown = m_bSpacebarIsDown
        
        If (m_bSpacebarIsDown) And (Not m_bMouseOnButton) Then
            DrawButton ebsDown
        ElseIf (m_tButtonSettings.Button = 8) Then
            ' 8 -> Control had been double clicked
            If (m_bMouseOnButton) Then
                DrawButton ebsHot
            Else
                DrawButton ebsNormal
            End If
        End If
        
        If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
            ReleaseCapture                      ' input processing
        End If                                  ' of the window
                                                '
        #If USE_POPUPMENU Then                  '
        If (Not m_bPopupEnabled) Then           '
            RaiseEvent MouseUp(Button, Shift, x, Y)
        Else                                    '
            If (Not m_bMouseOnButton) Then      '
                DrawButton ebsNormal            ' This fix the defect found when
            End If                              ' SubClass failed to raise on MouseUp
            m_bPopupInit = False                ' after a menu has been dismissed...
            m_bPopupShown = False               '
        End If                                  '
        #Else                                   '
            RaiseEvent MouseUp(Button, Shift, x, Y)
        #End If                                 '
    End If
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'   Called everytime the properties are needed for display after being initialized
'   (On loading of form and before the is control is unloaded to design mode)
    
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_ReadProperties"
    #End If
    
    With PropBag
        m_tButtonProperty.BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_tButtonProperty.Shape = .ReadProperty("ButtonShape", 0)
        m_tButtonProperty.Style = .ReadProperty("ButtonStyle", 0)
        m_tButtonProperty.Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_tButtonProperty.CheckBox = .ReadProperty("CheckBox", False)
        #If USE_SPECIALEFFECTS Then
        m_tButtonProperty.Effects = .ReadProperty("Effects", 0)
        #End If
        m_tButtonProperty.Enabled = .ReadProperty("Enabled", True)
        m_tButtonProperty.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
        m_tButtonProperty.HandPointer = .ReadProperty("HandPointer", False)
        m_tButtonProperty.MaskColor = .ReadProperty("MaskColor", &HC0C0C0)
        m_tButtonProperty.PicAlign = .ReadProperty("PicAlign", epaLeftOfCaption)
    Set m_tButtonProperty.PicDown = .ReadProperty("PicDown", Nothing)
    Set m_tButtonProperty.PicHot = .ReadProperty("PicHot", Nothing)
    Set m_tButtonProperty.PicNormal = .ReadProperty("PicNormal", Nothing)
        m_tButtonProperty.PicOpacity = .ReadProperty("PicOpacity", 1) ' Do not blend
        m_tButtonProperty.PicSize = .ReadProperty("PicSize", 0) ' epsNormal
        m_tButtonProperty.PicSizeH = .ReadProperty("PicSizeH", 0)
        m_tButtonProperty.PicSizeW = .ReadProperty("PicSizeW", 0)
        m_tButtonSettings.state = .ReadProperty("State", 0) ' ebsNormal
        m_tButtonProperty.UseMask = .ReadProperty("UseMask", True)
        m_tButtonProperty.Value = .ReadProperty("Value", False)
        
    Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
    Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0) ' vbDefault
    End With
    
    UserControl.AccessKeys = GetAccessKey(m_tButtonProperty.Caption) ' Assign accesskey
    UserControl.Enabled = m_tButtonProperty.Enabled
    
    If (Ambient.UserMode) Then ' Start subclassing
        
        If (m_tButtonProperty.HandPointer) Then
            m_tButtonSettings.Cursor = LoadCursor(0, IDC_HAND) ' Load hand pointer
            ' If LoadCursor fails, the system does not support hand pointer option
            m_tButtonProperty.HandPointer = (Not m_tButtonSettings.Cursor = 0)
        End If
        
        m_bTrackHandler32 = IsFunctionSupported("TrackMouseEvent", "User32")
        
        If (Not m_bTrackHandler32) Then
            If (Not IsFunctionSupported("_TrackMouseEvent", "Comctl32")) Then
                Err.Raise -1, "System does not support TrackMouseEvent."
                ' ...which is really neccessary for this control to work properly
                GoTo Jmp_Skip
            End If
        End If
        
        sc_Subclass hWnd               ' Subclass the control
        sc_AddMsg hWnd, WM_MOUSELEAVE  ' Detect mouse leave event
        
        Dim hParent As Long
        Dim hWindow As Long
            If (TypeOf parent Is Form) Then             ' Check if parent is a form
                hParent = parent.hWnd                   ' Get parent form handle
            ElseIf (TypeOf parent.parent Is Form) Then  ' If not check parent of it
                hParent = parent.parent.hWnd            ' and so on...
            ElseIf (TypeOf parent.parent.parent Is Form) Then
                hParent = parent.parent.parent.hWnd     ' If here still fails then
            End If                                      ' I quit. We just simply skip
                                                        ' subclass for parent form :)
            If (hParent) Then                           '
                sc_Subclass hParent                     '
                hWindow = GetWindowLong(hParent, GWL_EXSTYLE)
                                                        '
                If (hWindow And WS_EX_MDICHILD) Then    ' Bug fixed:
                    sc_AddMsg hParent, WM_NCACTIVATE    ' Now we check if the
                Else                                    ' parent form is an MDI
                    sc_AddMsg hParent, WM_ACTIVATE      ' child the API way :)
                End If                                  '
                
            End If
    End If
    
Jmp_Skip:
    SetButtonColors
    m_bCalculateRects = True
    DrawButton Force:=True
    
    m_bRedrawOnResize = True
    
End Sub

Private Sub UserControl_Resize()

    
    With m_tButtonSettings
        Dim lpRect As RECT
            GetClientRect UserControl.hWnd, lpRect
        
       
        
        .Height = lpRect.bottom
        .Width = lpRect.Right
        
        Const MIN_HPX As Long = 15
        Const MIN_WPX As Long = 15
        
       
        
        If (.Height < MIN_HPX) Or (.Width < MIN_WPX) Then
            If (.Height < MIN_HPX) Then
                UserControl.Height = MIN_HPX * Screen.TwipsPerPixelY
            End If
            If (.Width < MIN_WPX) Then
                UserControl.Width = MIN_WPX * Screen.TwipsPerPixelX
            End If
            Exit Sub
        End If
        
        #If DEBUG_EVENTS Then
            Debug.Print Caption & " - UserControl_Resize", .Height, .Width
        #End If
        
        m_bCalculateRects = True
        
        If (Ambient.UserMode) Then      ' Always allow to redraw when running mode
            DrawButton Force:=True      '
        ElseIf (m_bRedrawOnResize) Then ' On IDE, some sort of filtering is done
            DrawButton Force:=True      ' to prevent control to redraw twice or more
        End If                          '
    End With
    
End Sub

Private Sub UserControl_Show()
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_Show"
    #End If
    
    m_bControlHidden = False
    
End Sub

Private Sub UserControl_Terminate()
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_Terminate"
    #End If
    
    If (m_tButtonProperty.HandPointer) Then
        DeleteObject m_tButtonSettings.Cursor
    End If
    
    On Error GoTo Jmp_Skip
                                
    If (Ambient.UserMode) Then
        sc_Terminate
    End If
                                
Jmp_Skip:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    
    #If DEBUG_EVENTS Then
        Debug.Print Caption & " - UserControl_WriteProperties"
    #End If
    
    With PropBag
        .WriteProperty "BackColor", m_tButtonProperty.BackColor, Ambient.BackColor
        .WriteProperty "ButtonShape", m_tButtonProperty.Shape, 0
        .WriteProperty "ButtonStyle", m_tButtonProperty.Style, 0
        .WriteProperty "Caption", m_tButtonProperty.Caption, Ambient.DisplayName
        .WriteProperty "CheckBox", m_tButtonProperty.CheckBox, False
         #If USE_SPECIALEFFECTS Then
        .WriteProperty "Effects", m_tButtonProperty.Effects, 0
         #End If
        .WriteProperty "Enabled", m_tButtonProperty.Enabled, True
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "ForeColor", m_tButtonProperty.ForeColor, Ambient.ForeColor
        .WriteProperty "HandPointer", m_tButtonProperty.HandPointer, False
        .WriteProperty "MaskColor", m_tButtonProperty.MaskColor, &HC0C0C0
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointer", UserControl.MousePointer, 0 '
        .WriteProperty "PicAlign", m_tButtonProperty.PicAlign, epaLeftOfCaption
        .WriteProperty "PicDown", m_tButtonProperty.PicDown, Nothing
        .WriteProperty "PicHot", m_tButtonProperty.PicHot, Nothing
        .WriteProperty "PicNormal", m_tButtonProperty.PicNormal, Nothing
        .WriteProperty "PicOpacity", m_tButtonProperty.PicOpacity, 1
        .WriteProperty "PicSize", m_tButtonProperty.PicSize, 0
        .WriteProperty "PicSizeH", m_tButtonProperty.PicSizeH, 0
        .WriteProperty "PicSizeW", m_tButtonProperty.PicSizeW, 0
        .WriteProperty "State", m_tButtonSettings.state, 0
        .WriteProperty "UseMask", m_tButtonProperty.UseMask, True
        .WriteProperty "Value", m_tButtonProperty.Value, False
    End With
    
End Sub



Private Function sc_Subclass(ByVal lng_hWnd As Long, Optional ByVal lParamUser As Long = 0, Optional ByVal nOrdinal As Long = 1, Optional ByVal oCallback As Object = Nothing, Optional ByVal bIdeSafety As Boolean = True) As Boolean
    Const CODE_LEN As Long = 260: Const MEM_LEN As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)): Const PAGE_RWX As Long = &H40&: Const MEM_COMMIT As Long = &H1000&: Const MEM_RELEASE As Long = &H8000&: Const IDX_EBMODE As Long = 3: Const IDX_CWP As Long = 4: Const IDX_SWL As Long = 5: Const IDX_FREE As Long = 6: Const IDX_BADPTR As Long = 7: Const IDX_OWNER As Long = 8: Const IDX_CALLBACK As Long = 10: Const IDX_EBX As Long = 16: Const SUB_NAME As String = "sc_Subclass"
    Dim nAddr As Long, nID As Long, nMyID As Long
    If IsWindow(lng_hWnd) = 0 Then Exit Function
    nMyID = GetCurrentProcessId
    GetWindowThreadProcessId lng_hWnd, nID
    If nID <> nMyID Then Exit Function
    If oCallback Is Nothing Then Set oCallback = Me
    nAddr = zAddressOf(oCallback, nOrdinal)
    If nAddr = 0 Then Exit Function
    If z_Funk Is Nothing Then
        Set z_Funk = New Collection
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
        z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&
        z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA"): z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA"): z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree"): z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)
    If z_ScMem <> 0 Then
        On Error GoTo ReleaseMemory
        z_Funk.Add z_ScMem, "h" & lng_hWnd
        On Error GoTo 0
        If bIdeSafety Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")
        z_Sc(IDX_EBX) = z_ScMem: z_Sc(IDX_HWND) = lng_hWnd: z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN: z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4): z_Sc(IDX_OWNER) = ObjPtr(oCallback): z_Sc(IDX_CALLBACK) = nAddr: z_Sc(IDX_PARM_USER) = lParamUser: nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)
        If nAddr = 0 Then GoTo ReleaseMemory
        z_Sc(IDX_WNDPROC) = nAddr: RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN: sc_Subclass = True
    End If
    Exit Function
ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE
End Function
Private Sub sc_Terminate()
    Dim i As Long
    If Not (z_Funk Is Nothing) Then
        For i = z_Funk.count To 1 Step -1
            z_ScMem = z_Funk.Item(i)
            If IsBadCodePtr(z_ScMem) = 0 Then sc_UnSubclass zData(IDX_HWND)
        Next i
        Set z_Funk = Nothing
    End If
End Sub
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
    If Not (z_Funk Is Nothing) Then
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
            zData(IDX_SHUTDOWN) = -1: zDelMsg ALL_MESSAGES, IDX_BTABLE: zDelMsg ALL_MESSAGES, IDX_ATABLE
        End If
        z_Funk.Remove "h" & lng_hWnd
    End If
End Sub
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        If When And MSG_BEFORE Then zAddMsg uMsg, IDX_BTABLE
        If When And MSG_AFTER Then zAddMsg uMsg, IDX_ATABLE
    End If
End Sub
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        If When And MSG_BEFORE Then zDelMsg uMsg, IDX_BTABLE
        If When And MSG_AFTER Then zDelMsg uMsg, IDX_ATABLE
    End If
End Sub
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam)
    End If
End Function
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        sc_lParamUser = zData(IDX_PARM_USER)
    End If
End Property
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal newValue As Long)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        zData(IDX_PARM_USER) = newValue
    End If
End Property
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long, nBase As Long, i As Long
    nBase = z_ScMem: z_ScMem = zData(nTable)
    If uMsg = ALL_MESSAGES Then
        nCount = ALL_MESSAGES
    Else
        nCount = zData(0)
        If nCount >= MSG_ENTRIES Then GoTo Bail
        For i = 1 To nCount
            If zData(i) = 0 Then
                zData(i) = uMsg: GoTo Bail
            ElseIf zData(i) = uMsg Then
                GoTo Bail
            End If
        Next i
        nCount = i: zData(nCount) = uMsg
    End If
    zData(0) = nCount
Bail:
    z_ScMem = nBase
End Sub
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long, nBase As Long, i As Long
    nBase = z_ScMem: z_ScMem = zData(nTable)
    If uMsg = ALL_MESSAGES Then
        zData(0) = 0
    Else
        nCount = zData(0)
        For i = 1 To nCount
            If zData(i) = uMsg Then
                zData(i) = 0: GoTo Bail
            End If
        Next i
    End If
Bail:
    z_ScMem = nBase
End Sub
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
End Function
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
    If Not (z_Funk Is Nothing) Then
        On Error GoTo Catch
        z_ScMem = z_Funk("h" & lng_hWnd): zMap_hWnd = z_ScMem
    End If
Catch:
End Function
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    Dim bSub As Byte, bVal As Byte, nAddr As Long, i As Long, j As Long
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4
    If Not zProbe(nAddr + &H1C, i, bSub) Then
        If Not zProbe(nAddr + &H6F8, i, bSub) Then
            If Not zProbe(nAddr + &H7A4, i, bSub) Then Exit Function
        End If
    End If
    i = i + 4: j = i + 1024
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4
        If IsBadCodePtr(nAddr) Then
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4: Exit Do
        End If
        RtlMoveMemory VarPtr(bVal), nAddr, 1
        If bVal <> bSub Then
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4: Exit Do
        End If
        i = i + 4
    Loop
End Function
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal As Byte, nAddr As Long, nLimit As Long, nEntry As Long
    nAddr = nStart: nLimit = nAddr + 32
    Do While nAddr < nLimit
        RtlMoveMemory VarPtr(nEntry), nAddr, 4
        If Not nEntry = 0 Then
            RtlMoveMemory VarPtr(bVal), nEntry, 1
            If bVal = &H33 Or bVal = &HE9 Then
                nMethod = nAddr: bSub = bVal: zProbe = True: Exit Function
            End If
        End If
        nAddr = nAddr + 4
    Loop
End Function
Private Property Get zData(ByVal nIndex As Long) As Long
    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property
Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property
Private Sub zWndProc1(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
   
    Subclass_Proc bBefore, bHandled, lReturn, lng_hWnd, uMsg, wParam, lParam, lParamUser
End Sub

