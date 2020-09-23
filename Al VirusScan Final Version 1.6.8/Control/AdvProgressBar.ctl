VERSION 5.00
Begin VB.UserControl AdvProgressBar 
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   PropertyPages   =   "AdvProgressBar.ctx":0000
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ToolboxBitmap   =   "AdvProgressBar.ctx":0017
   Begin VB.PictureBox x 
      Height          =   105
      Left            =   810
      ScaleHeight     =   45
      ScaleWidth      =   1125
      TabIndex        =   2
      Top             =   435
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox PB 
      AutoRedraw      =   -1  'True
      Height          =   360
      Left            =   30
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   15
      Width           =   4815
   End
   Begin VB.Image CustomF 
      Height          =   210
      Left            =   3690
      Top             =   495
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image customI 
      Height          =   195
      Left            =   2985
      Top             =   495
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgXPC 
      Height          =   105
      Left            =   2505
      Picture         =   "AdvProgressBar.ctx":0329
      Top             =   540
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgXPF 
      Height          =   195
      Left            =   1035
      Picture         =   "AdvProgressBar.ctx":03F7
      Top             =   795
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   255
      TabIndex        =   1
      Top             =   390
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "AdvProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Enum
Enum BorderStyleConstantsH
vbNoBorder = 0
vbSingle = 1
End Enum
Enum AppearanceConstants
vbFlat = 0
vb3D = 1
End Enum
Private xp As Boolean ' if the last style was xp, this is true
Enum StyleConstants
Standart = 0
Smooth = 1
ValueDependant = 2 'Bar2 if low, bar1 if high
SmoothDoubleColor = 3
DoubleColor = 4
XPStyle = 5
CustomPictureTile = 6
CustomPictureShow = 7
CustomPictureStrech = 8
End Enum
'Default Property Values:
Const m_def_Style = 1
Const m_def_TextColor = vbBlack
Const m_def_BarColor2 = vbGreen
Const m_def_BarColor1 = &H8000000D
Const m_def_Value = 100
Const m_def_Max = 100
Const m_def_ShowText = False
'Property Variables:
Dim m_CustomFrame As Picture
Dim m_CustomPicture As Picture
Dim m_Style As StyleConstants
Dim m_TextColor As OLE_COLOR
Dim m_BarColor2 As OLE_COLOR
Dim m_BarColor1 As OLE_COLOR
Dim m_Value As Long
Dim m_Max As Long
Dim m_ShowText As Boolean
'Event Declarations:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
Attribute AutoRedraw.VB_ProcData.VB_Invoke_Property = "Settings"
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Private Sub PB_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub PB_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub PB_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub PB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X + PB.Left, Y + PB.Top)
End Sub

Private Sub PB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X + PB.Left, Y + PB.Top)
End Sub

Private Sub PB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X + PB.Left, Y + PB.Left)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=17,0,1,0
'Public Property Get Style() As StyleConstants
'    Style = m_Style
'End Property
'
'Public Property Let Style(ByVal New_Style As StyleConstants)
'    If Ambient.UserMode = False Then Err.Raise 383
'    Set m_Style = New_Style
'    PropertyChanged "Style"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=8,0,0,0
''Public Property Get Value() As Long
''    Value = m_Value
''End Property
''
''Public Property Let Value(ByVal New_Value As Long)
''    m_Value = New_Value
''    PropertyChanged "Value"
''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,0
Public Property Get Max() As Long
Attribute Max.VB_ProcData.VB_Invoke_Property = "APB"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    If Ambient.UserMode = False Then Err.Raise 381
    m_Max = New_Max
    PropertyChanged "Max"
    Call UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,1,0
Public Property Get ShowText() As Boolean
Attribute ShowText.VB_ProcData.VB_Invoke_Property = "Settings"
    ShowText = m_ShowText
End Property

Public Property Let ShowText(ByVal New_ShowText As Boolean)
    m_ShowText = New_ShowText
    lblVal.Visible = m_ShowText
    PropertyChanged "ShowText"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_Max = m_def_Max
    m_ShowText = m_def_ShowText
    m_Value = m_def_Value
    m_Style = m_def_Style
    m_TextColor = m_def_TextColor
    m_BarColor2 = m_def_BarColor2
    m_BarColor1 = m_def_BarColor1
    Set m_CustomFrame = LoadPicture("")
    Set m_CustomPicture = LoadPicture("")
End Sub
Private Sub UpdateText()
    lblVal = Int(m_Value / m_Max * 100) & "%"
    PB.CurrentX = PB.Width / 2 - lblVal.Width / 2
    PB.CurrentY = PB.Height / 2 - lblVal.Height / 2
    PB.ForeColor = m_TextColor
    If m_ShowText Then PB.Print lblVal
End Sub
Private Sub UserControl_Paint()
Dim percent As Integer
Static last As Double
Static old As Integer
Dim availp As Long, r1 As Long, g1 As Long, b1 As Long, r2 As Long, g2 As Long, b2 As Long
Dim rR As Single, gR As Single, bR As Single, maxH


    PB.ForeColor = m_TextColor
On Error GoTo error
Select Case Style
Case StyleConstants.Standart
    Dim i As Long
    If last > m_Value / m_Max Then PB.Cls
    last = m_Value / m_Max
    Do Until i >= Int(PB.Width * (m_Value / m_Max))
    PB.Line (i, 0)-(i + 13, PB.Height), m_BarColor1, BF 'Active Color
    PB.Line (i + 13, 0)-(i + 14, PB.Height), vbButtonFace, BF 'Erasal
    i = i + 15
    Loop
    If i > PB.Width / 2 Then
    PB.Line (i, 0)-(PB.Width - (PB.Width - i), PB.Height), vbButtonFace, BF 'Erase the rest
    Else
    PB.Line (i, 0)-(PB.Width - i, PB.Height), vbButtonFace, BF
    End If
Case StyleConstants.Smooth
    PB.Cls
    old = Int(m_Value / m_Max * PB.Width)
    PB.Line (0, 0)-(old, PB.Height), m_BarColor1, BF
    
Case StyleConstants.ValueDependant
    Call GetRGB(m_BarColor1, r1, g1, b1)
    Call GetRGB(m_BarColor2, r2, g2, b2)
    d = m_Value / m_Max * 100
    rR = (r2 - r1)
    gR = (g2 - g1)
    bR = (b2 - b1)
    PB.Cls
    On Error Resume Next
    PB.ForeColor = RGB(Int(r1 + rR * d / 100), Int(g1 + gR * d / 100), Int(b1 + bR * d / 100))
    PB.Line (0, 0)-(PB.Width * d / 100, PB.Height), , BF

Case StyleConstants.SmoothDoubleColor
    On Error Resume Next
    PB.Cls
    If m_Value = 0 Then Exit Sub
    Call GetRGB(m_BarColor1, r1, g1, b1)
    Call GetRGB(m_BarColor2, r2, g2, b2)
    maxH = Int(PB.Width * m_Value / m_Max)
    rR = (r2 - r1) / maxH
    gR = (g2 - g1) / maxH
    bR = (b2 - b1) / maxH
    For i = 0 To Int(PB.Width * m_Value / m_Max) Step 1
    PB.Line (i, 0)-(i + 5, PB.Height), RGB(r1 + rR * i, g1 + gR * i, b1 + bR * i), BF
    Next i
Case StyleConstants.DoubleColor
    'PB.Cls
    If last > m_Value / m_Max Then PB.Cls
    last = m_Value / m_Max
    If last = 0 Then Exit Sub
    On Error Resume Next
    Call GetRGB(m_BarColor1, r1, g1, b1)
    Call GetRGB(m_BarColor2, r2, g2, b2)
    maxH = Int(PB.Width * m_Value / m_Max) / 15
    rR = (r2 - r1) / maxH
    gR = (g2 - g1) / maxH
    bR = (b2 - b1) / maxH
    For i = 0 To Int(PB.Width * m_Value / m_Max) Step 15
        PB.Line (i, 0)-(i + 13, PB.Height), RGB(r1 + rR * i / 15, g1 + gR * i / 15, b1 + bR * i / 15), BF
    Next i
    If i > PB.Width / 2 Then
    PB.Line (i, 0)-(PB.Width - (PB.Width - i), PB.Height), vbButtonFace, BF 'Erase the rest
    Else
    PB.Line (i, 0)-(PB.Width - i, PB.Height), vbButtonFace, BF
    End If
Case StyleConstants.XPStyle
    If Not xp Then PB.PaintPicture imgXPF.Picture, 0, 0, PB.ScaleWidth, PB.ScaleHeight: xp = True ' Not to redraw the whole thing
    If Int(m_Value / m_Max * PB.ScaleWidth) = 0 Or m_Value = 0 Then i = 1: GoTo exXP
    Dim availSpace As Long ' Space on the bar actually waiting for this
    availSpace = Int((PB.ScaleWidth - 5) / (imgXPC.Width + 1)) - 1
    For i = 0 To Int(m_Value / m_Max * availSpace)
    PB.PaintPicture imgXPC.Picture, 3 + i + imgXPC.Width * i, 2, , imgXPF.Height - 4
    DoEvents
    Next i
exXP:
    If m_Value / m_Max <> 1 Then PB.Line (3 + (imgXPC.Width + 1) * i, 2)-(PB.ScaleWidth - 5, imgXPF.Height - 3), vbWhite, BF
Case StyleConstants.CustomPictureShow
    Vx = m_Value / m_Max * PB.ScaleWidth
    If Vx = 0 Then PB.Cls: Exit Sub
    PB.PaintPicture customI.Picture, 0, 0, , PB.ScaleHeight, , , Vx
    PB.Line (Vx, 0)-(PB.ScaleWidth, PB.ScaleHeight), vbButtonFace, BF
Case StyleConstants.CustomPictureStrech
    Vx = Int(m_Value / m_Max * PB.ScaleWidth)
    If Vx = 0 Then PB.Cls: Exit Sub
    PB.PaintPicture customI, 0, 0, Vx, PB.ScaleHeight
    PB.Line (Vx, 0)-(PB.ScaleWidth, PB.ScaleHeight), vbButtonFace, BF
Case StyleConstants.CustomPictureTile
    Dim aSize As Long ' Number of times to put this on.
    aSize = Int(m_Value / m_Max * (PB.ScaleWidth / (customI.Width + 1)))
    For i = 1 To aSize
    PB.PaintPicture customI.Picture, (i - 1) * customI.Width, 0, , PB.ScaleHeight
    PB.Line (i * (customI.Width + 1), 0)-(i * (customI.Width + 1) + 1, PB.ScaleHeight), vbButtonFace, BF
    Next i
    PB.Line (i * (customI.Width + 1), 0)-(PB.ScaleWidth, PB.ScaleHeight), vbButtonFace, BF
End Select
    UpdateText
    If Style <> XPStyle Then xp = False

Exit Sub
error:
Select Case Err.Number
Case 11
PB.Cls
Case 481
'INVALID PICTURE.
Style = SmoothDoubleColor
Exit Sub
'Case Else
'Err.Raise Err.Number, , Err.Description, Err.HelpFile, Err.HelpContext
Resume
End Select
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_ShowText = PropBag.ReadProperty("ShowText", m_def_ShowText)
    PB.Appearance = PropBag.ReadProperty("Appearance", 1)
    PB.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_TextColor = PropBag.ReadProperty("TextColor", m_def_TextColor)
    m_BarColor2 = PropBag.ReadProperty("BarColor2", m_def_BarColor2)
    m_BarColor1 = PropBag.ReadProperty("BarColor1", vbHighlightText)
    Set customI.Picture = PropBag.ReadProperty("CustomPicture", Nothing)
    Set CustomF.Picture = PropBag.ReadProperty("CustomFrame", Nothing)
    Set m_CustomPicture = PropBag.ReadProperty("CustomPicture", Nothing)
    Set m_CustomFrame = PropBag.ReadProperty("CustomFrame", Nothing)
    Call UserControl_Paint
End Sub

Private Sub UserControl_Resize()
UserControl.ScaleMode = vbTwips
PB.Height = UserControl.ScaleHeight
PB.Width = UserControl.ScaleWidth
PB.Top = 0
PB.Left = 0
lblVal.Top = (UserControl.ScaleHeight - lblVal.Height) / 2
lblVal.Left = (UserControl.ScaleWidth - lblVal.Width) / 2
If Style = XPStyle Then
PB.Height = imgXPF.Height
UserControl.Height = PB.Height
xp = False
End If
UserControl.ScaleMode = vbPixels
PB.Cls
Call UserControl_Paint
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("ShowText", m_ShowText, m_def_ShowText)
    Call PropBag.WriteProperty("Appearance", PB.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", PB.BorderStyle, 1)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("TextColor", m_TextColor, m_def_TextColor)
    Call PropBag.WriteProperty("BarColor2", m_BarColor2, m_def_BarColor2)
    Call PropBag.WriteProperty("BarColor1", m_BarColor1, m_def_BarColor1)
    Call PropBag.WriteProperty("CustomPicture", customI.Picture, Nothing)
    Call PropBag.WriteProperty("CustomFrame", CustomF.Picture, Nothing)
    Call PropBag.WriteProperty("CustomFrame", m_CustomFrame, Nothing)
    Call PropBag.WriteProperty("CustomPicture", m_CustomPicture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PB,PB,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = PB.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    PB.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PB,PB,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstantsH
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = PB.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstantsH)
    PB.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,3,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "What is the current progress"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    If New_Value > m_Max Then New_Value = m_Max
    m_Value = New_Value
    Call UserControl_Paint
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,1,2
Public Property Get Style() As StyleConstants
Attribute Style.VB_Description = "What kind of progress bar sshould this be?"
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As StyleConstants)
Static preXPhadBorders As Boolean ' If the previous style (before XP) had borders they are restored
Static postXP As Boolean ' This is true if XP was on before this.
    m_Style = New_Style
    If New_Style = XPStyle Then
    postXP = True
    preXPhadBorders = IIf(PB.BorderStyle > 0, True, False)
    BorderStyle = 0
    Call UserControl_Resize
    ElseIf NewStyle > XPStyle Then
    If postXP Then
    postXP = False
    If preXPhadBorders Then PB.BorderStyle = 1
    End If
    BorderStyle = 1
    End If
    PropertyChanged "Style"
    Call UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,1,vbgreen
Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "What is the color of the text (if showing it)"
    TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
    m_TextColor = New_TextColor
    PropertyChanged "TextColor"
    lblVal.ForeColor = New_TextColor
    Call UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbGreen
Public Property Get BarColor2() As OLE_COLOR
Attribute BarColor2.VB_Description = "What color should the end of the progress bar be (near 100 percent), works only with style set to DoubleColor"
    BarColor2 = m_BarColor2
End Property

Public Property Let BarColor2(ByVal New_BarColor2 As OLE_COLOR)
    m_BarColor2 = New_BarColor2
    Call UserControl_Paint
    PropertyChanged "BarColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbGreen
Public Property Get BarColor1() As OLE_COLOR
Attribute BarColor1.VB_Description = "What color should the bar be"
    BarColor1 = m_BarColor1
End Property

Public Property Let BarColor1(ByVal New_BarColor1 As OLE_COLOR)
    m_BarColor1 = New_BarColor1
    Call UserControl_Paint
    PropertyChanged "BarColor1"
End Property
'''
''''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''''MemberInfo=11,0,0,0
'''Public Property Get CustomPicture() As Picture
'''    Set CustomPicture = m_CustomPicture
'''End Property
'''
'''Public Property Set CustomPicture(ByVal New_CustomPicture As Picture)
'''    Set m_CustomPicture = New_CustomPicture
'''    PropertyChanged "CustomPicture"
'''End Property
'''
''''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''''MemberInfo=11,0,0,0
'''Public Property Get CustomFrame() As Picture
'''    Set CustomFrame = m_CustomFrame
'''End Property
'''
'''Public Property Set CustomFrame(ByVal New_CustomFrame As Picture)
'''    Set m_CustomFrame = New_CustomFrame
'''    PropertyChanged "CustomFrame"
'''End Property
'''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=imgXPC,imgXPC,-1,Picture
''Public Property Get CustomPicture() As Picture
''    Set CustomPicture = imgXPC.Picture
''End Property
''
''Public Property Set CustomPicture(ByVal New_CustomPicture As Picture)
''    Set imgXPC.Picture = New_CustomPicture
''    PropertyChanged "CustomPicture"
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=imgXPF,imgXPF,-1,Picture
''Public Property Get CustomFrame() As Picture
''    Set CustomFrame = imgXPF.Picture
''End Property
''
''Public Property Set CustomFrame(ByVal New_CustomFrame As Picture)
''    Set imgXPF.Picture = New_CustomFrame
''    PropertyChanged "CustomFrame"
''End Property
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=customI,customI,-1,Picture
'Public Property Get CustomPicture() As Picture
'    Set CustomPicture = customI.Picture
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=CustomF,CustomF,-1,Picture
'Public Property Get CustomFrame() As Picture
'    Set CustomFrame = CustomF.Picture
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=customI,customI,-1,Picture
'Public Property Get CustomPicture() As Picture
'    Set CustomPicture = customI.Picture
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get CustomFrame() As Picture
Attribute CustomFrame.VB_Description = "Sets the frame image, if the style is FrameMath."
    Set CustomFrame = m_CustomFrame
End Property

Public Property Set CustomFrame(ByVal New_CustomFrame As IPictureDisp)
    Set m_CustomPicture = New_CustomFrame
    Set customI.Picture = New_CustomFrame
    PropertyChanged "CustomFrame"
End Property

Public Property Let CustomFrame(ByVal New_CustomFrame As IPictureDisp)
    Set m_CustomPicture = New_CustomFrame
    Set customI.Picture = New_CustomFrame
    PropertyChanged "CustomFrame"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get CustomPicture() As Picture
Attribute CustomPicture.VB_Description = "Sets the unit picture if the style is CustomFrame, and sets the image in most other styles."
    Set CustomPicture = m_CustomPicture
End Property

Public Property Set CustomPicture(ByVal New_CustomPicture As IPictureDisp)
    Set m_CustomPicture = New_CustomPicture
    Set customI.Picture = New_CustomPicture
    PropertyChanged "CustomPicture"
End Property
Public Property Let CustomPicture(ByVal New_CustomPicture As IPictureDisp)
    Set m_CustomPicture = New_CustomPicture
    Set customI.Picture = New_CustomPicture
    PropertyChanged "CustomPicture"
End Property
