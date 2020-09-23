Attribute VB_Name = "mListView"
'=======================================
' Moh Aly Shodiqin
' felix_progressif@yahoo.com
' 18 Jan 2009
'---------------------------------------
' Set style List-View controls SP2
'---------------------------------------
' Control   LVS_EX_FULLROWSELECT
'           LVS_EX_GRIDLINES
'           LVS_EX_CHECKBOXES
'           LVS_EX_SUBITEMIMAGES
'           etc.
'
' Kode ini berdasarkan Microsoft Knowledge Base code.
' MSDN Library 2005
' URL : ms-help://MS.MSDNQTR.v80.en/MS.MSDN.v80/MS.WIN32COM.v10.en/shellcc/platform/commctls/listview/messages/lvm_setextendedlistviewstyle.htm
'=======================================
Option Explicit

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVS_EX_HEADERDRAGDROP = &H10
Public Const LVS_EX_TRACKSELECT = &H8
Public Const LVS_EX_ONECLICKACTIVATE = &H40
Public Const LVS_EX_TWOCLICKACTIVATE = &H80
Public Const LVS_EX_SUBITEMIMAGES = &H2

Public Const LVIF_STATE = &H8
 
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const GWL_STYLE        As Long = (-16)
Private Const LVM_GETHEADER    As Long = (LVM_FIRST + 31)
Private Const LVM_ARRANGE      As Long = (LVM_FIRST + 22)
Private Const HDS_BUTTONS      As Long = 2

Public Const LVIS_STATEIMAGEMASK As Long = &HF000

Public Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   state        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Public Const LVM_GETCOLUMN = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVCF_TEXT = &H4

Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cX As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

Function lvwStyle(lvStyle As ListView)
    Dim rStyle As Long
    Dim r As Long
    rStyle = SendMessageLong(lvStyle.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    rStyle = rStyle Xor LVS_EX_FULLROWSELECT
    r = SendMessageLong(lvStyle.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
End Function

Function lvwStyleProcess(lvStyle As ListView)
    Dim rStyle As Long
    Dim r As Long
    rStyle = SendMessageLong(lvStyle.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    rStyle = rStyle Xor LVS_EX_FULLROWSELECT Xor LVS_EX_GRIDLINES Xor LVS_EX_ONECLICKACTIVATE
    r = SendMessageLong(lvStyle.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
End Function

Public Sub SetFlatHeaders(LVhwnd As Long)
    Dim hHeader As Long
    Dim Style   As Long

    'get the handle to the listview header
    hHeader = SendMessage(LVhwnd, LVM_GETHEADER, 0, ByVal 0&)
    'set the new style
    Style = GetWindowLong(hHeader, GWL_STYLE)
    Style = Style And Not HDS_BUTTONS
    Call SetWindowLong(hHeader, GWL_STYLE, Style)
End Sub

Public Sub ArrangeLV(lstView As ListView)
    If lstView.View <> lvwReport Then
        Call SendMessage(lstView.hWnd, LVM_ARRANGE, 0, ByVal 0&)
    End If
End Sub
