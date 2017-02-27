VERSION 5.00
Begin VB.UserControl XpB 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   ForeColor       =   &H80000015&
   MousePointer    =   99  'Custom
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   ToolboxBitmap   =   "Xpb.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   1245
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "XpB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private MouseDownPressed As Boolean    'был ли мд до апа, для клика

Dim UserScaleW As Long
Dim UserScaleH As Long

'mzt Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'mzt Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'mzt Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
'mzt Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
'Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As textparametreleri) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
'mzt Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'mzt Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
'Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type RGB
    Red As Byte
    Green As Byte
    Blue As Byte
End Type
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type textparametreleri
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Public Enum XBPicturePosition
    gbTOP = 0
    gbLEFT = 1
    gbRIGHT = 2
    gbBOTTOM = 3
End Enum
Public Enum XBButtonStyle
    gbStandard = 0
    gbFlat = 1
    gbWinXP = 3
End Enum

Dim mvarClientRect As RECT
Dim mvarPictureRect As RECT
Dim mvarCaptionRect As RECT
Dim mvarOrgRect As RECT
Dim g_FocusRect As RECT
Dim alan As RECT
Dim m_Picture As Picture
Dim m_PicturePosition As XBPicturePosition
Dim m_ButtonStyle As XBButtonStyle
Dim mvarDrawTextParams As textparametreleri
Dim m_Caption As String
'mzt Dim m_Enabled As Boolean 'sur
'mzt Dim m_BackColor As Long 'Sur
Dim m_FontBold As Boolean    'sur

Dim m_PictureWidth As Long
Dim m_PictureHeight As Long
Dim g_HasFocus As Byte
Dim g_MouseDown As Byte
Dim g_MouseIn As Byte
Dim m_ShowFocusRect As Boolean

Const mvarPadding As Byte = 4

Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseIn(Shift As Integer)
Event MouseOut(Shift As Integer)
Event Click()
Event ShiftClick(Shift As Integer)

Dim m_MaskColor As Long
Dim m_UseMaskColor As Long
Dim m_XPDefaultColors As Boolean
Dim m_XPColor_Pressed As Long
Dim m_XPColor_Hover As Long



Private Sub UserControl_AmbientChanged(PropertyName As String)
pInitialize
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    m_ShowFocusRect = .ReadProperty("ShowFocusRect", 1)
    m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
    m_PicturePosition = .ReadProperty("PicturePosition", 1)
    m_ButtonStyle = .ReadProperty("ButtonStyle", 2)
    m_PictureWidth = .ReadProperty("PictureWidth", 32)
    m_PictureHeight = .ReadProperty("PictureHeight", 32)
    Set m_Picture = .ReadProperty("Picture", Nothing)
    UserControl.Enabled = .ReadProperty("Enabled", True)
    UserControl.FontBold = .ReadProperty("FontBold", False)
    UserControl.BackColor = .ReadProperty("Backcolor", &HFFFFFF)     'Sur
    'UserControl.ForeColor = .ReadProperty("ForeColor", &H80000012)    'Sur
    m_XPColor_Pressed = .ReadProperty("XPColor_Pressed", &H80000014)
    m_XPColor_Hover = .ReadProperty("XPColor_Hover", &H80000016)
    m_XPDefaultColors = .ReadProperty("XPDefaultColors", 1)
    '    UserControl.MousePointer = .ReadProperty("MousePointer", 0)
    m_MaskColor = .ReadProperty("MaskColor", 0)
    m_UseMaskColor = .ReadProperty("UseMaskColor", 0)
End With

SetAccessKeys

pInitialize
End Sub

'Private Sub UserControl_Resize()
''If ExitSVC Then Exit Sub
'    pInitialize
'End Sub

Private Sub UserControl_Terminate()
'If ExitSVC Then Exit Sub
Set m_Picture = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call .WriteProperty("PicturePosition", m_PicturePosition, 1)
    Call .WriteProperty("ButtonStyle", m_ButtonStyle, 2)
    Call .WriteProperty("Picture", m_Picture, Nothing)
    Call .WriteProperty("PictureWidth", m_PictureWidth, 32)
    Call .WriteProperty("PictureHeight", m_PictureHeight, 32)
    Call .WriteProperty("Enabled", Enabled, True)
    Call .WriteProperty("BackColor", BackColor, &HFFFFFF)     'sur
    'Call .WriteProperty("ForeColor", ForeColor, &H80000012)     'sur

    Call .WriteProperty("FontBold", m_FontBold, False)  'sur
    '    Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call .WriteProperty("ShowFocusRect", m_ShowFocusRect, 1)

    Call .WriteProperty("XPColor_Pressed", m_XPColor_Pressed, &H80000014)
    Call .WriteProperty("XPColor_Hover", m_XPColor_Hover, &H80000016)
    Call .WriteProperty("XPDefaultColors", m_XPDefaultColors, 1)
    Call .WriteProperty("MaskColor", m_MaskColor, 0)
    Call .WriteProperty("UseMaskColor", m_UseMaskColor, 0)
End With
End Sub
Private Sub CalcRECTs()
Dim picWidth As Integer
Dim picHeight As Integer
Dim capWidth As Integer
Dim capHeight As Integer
With alan
    .left = 0
    .top = 0
    .right = UserScaleW - 1
    .bottom = UserScaleH - 1
End With

With mvarClientRect
    .left = alan.left + mvarPadding
    .top = alan.top + mvarPadding
    .right = alan.right - mvarPadding + 1
    .bottom = alan.bottom - mvarPadding + 1
End With

If Len(m_Caption) = 0 Then
    With mvarPictureRect
        .left = (((mvarClientRect.right - mvarClientRect.left) - m_PictureWidth) * 0.5) + mvarClientRect.left
        .top = (((mvarClientRect.bottom - mvarClientRect.top) - m_PictureHeight) * 0.5) + mvarClientRect.top
        .right = .left + m_PictureWidth
        .bottom = .top + m_PictureHeight
    End With
Else
    With mvarCaptionRect
        .left = mvarClientRect.left
        .top = mvarClientRect.top
        .right = mvarClientRect.right
        .bottom = mvarClientRect.bottom
    End With
    
    CalculateCaptionRect

    If m_Picture Is Nothing Then Exit Sub

    picWidth = m_PictureWidth
    picHeight = m_PictureHeight
    With mvarCaptionRect
        capWidth = .right - .left
        capHeight = .bottom - .top
    End With

    If m_PicturePosition = gbLEFT Then
        With mvarPictureRect
            .top = (((mvarClientRect.bottom - mvarClientRect.top) - picHeight) * 0.5) + mvarClientRect.top
            
            '.left = (((mvarClientRect.right - mvarClientRect.left) - (picWidth + mvarPadding + capWidth)) * 0.5) + mvarClientRect.left
            'левее
            .left = mvarClientRect.left + mvarPadding ' + mvarPadding
            
            .bottom = .top + picHeight
            .right = .left + picWidth
        End With
        With mvarCaptionRect
            .top = (((mvarClientRect.bottom - mvarClientRect.top) - capHeight) * 0.5) + mvarClientRect.top
            '.left = mvarPictureRect.right + mvarPadding
'.left = (((mvarClientRect.right - mvarClientRect.left) - (capWidth + picWidth) * 0.5)) '+ picWidth / 2 + mvarClientRect.left
'.left = (((mvarClientRect.right - mvarClientRect.left) - (picWidth + capWidth)) * 0.5) + mvarClientRect.left
'по центру
.left = (((mvarClientRect.right - mvarClientRect.left + picWidth) - capWidth) * 0.5) + mvarClientRect.left

            .bottom = .top + capHeight
            .right = .left + capWidth
        End With
        '        CalculateCaptionRect
    ElseIf m_PicturePosition = gbRIGHT Then
        With mvarCaptionRect
            .top = (((mvarClientRect.bottom - mvarClientRect.top) - capHeight) * 0.5) + mvarClientRect.top
            .left = (((mvarClientRect.right - mvarClientRect.left) - (picWidth + mvarPadding + capWidth)) * 0.5) + mvarClientRect.left
            .bottom = .top + capHeight
            .right = .left + capWidth
        End With
        With mvarPictureRect
            .top = (((mvarClientRect.bottom - mvarClientRect.top) - picHeight) * 0.5) + mvarClientRect.top
            .left = mvarCaptionRect.right + mvarPadding
            .bottom = .top + picHeight
            .right = .left + picWidth
        End With
    ElseIf m_PicturePosition = gbTOP Then
        With mvarPictureRect
            .top = (((mvarClientRect.bottom - mvarClientRect.top) - (picHeight + mvarPadding + capHeight)) * 0.5) + mvarClientRect.top
            .left = (((mvarClientRect.right - mvarClientRect.left) - picWidth) * 0.5) + mvarClientRect.left
            .bottom = .top + picHeight
            .right = .left + picWidth
        End With
        With mvarCaptionRect
            .top = mvarPictureRect.bottom + mvarPadding
            .left = (((mvarClientRect.right - mvarClientRect.left) - capWidth) * 0.5) + mvarClientRect.left
            .bottom = .top + capHeight
            .right = .left + capWidth
        End With
    ElseIf m_PicturePosition = gbBOTTOM Then
        With mvarCaptionRect
            .top = (((mvarClientRect.bottom - mvarClientRect.top) - (picHeight + mvarPadding + capHeight)) * 0.5) + mvarClientRect.top
            .left = (((mvarClientRect.right - mvarClientRect.left) - capWidth) * 0.5) + mvarClientRect.left
            .bottom = .top + capHeight
            .right = .left + capWidth
        End With
        With mvarPictureRect
            .top = mvarCaptionRect.bottom + mvarPadding
            .left = (((mvarClientRect.right - mvarClientRect.left) - picWidth) * 0.5) + mvarClientRect.left
            .bottom = .top + picHeight
            .right = .left + picWidth
        End With
    End If
End If
End Sub
Public Sub pInitialize()
If ExitSVC Then Exit Sub

ScaleMode = 3
PaletteMode = 3

UserScaleW = UserControl.ScaleWidth
UserScaleH = UserControl.ScaleHeight

If UserScaleW < 10 Then UserControl.Width = 150
If UserScaleH < 10 Then UserControl.Height = 150

With g_FocusRect
    .left = 4
    .right = UserScaleW - 4
    .top = 4
    .bottom = UserScaleH - 4
End With

Refresh
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
'Debug.Print KeyAscii
MouseDownPressed = False

If Enabled = False Then Exit Sub
RaiseEvent Click
End Sub
Private Sub UserControl_EnterFocus()
g_MouseIn = 0
g_HasFocus = 1
Refresh
End Sub
Private Sub UserControl_ExitFocus()
If ExitSVC Then Exit Sub

g_HasFocus = 0
g_MouseDown = 0
g_MouseIn = 0
Refresh
End Sub

Public Sub Refresh()
AutoRedraw = True
UserControl.Cls

If m_ButtonStyle = gbWinXP Then
    DrawWinXPButton g_MouseDown, g_MouseIn
Else
    If g_MouseDown = 1 Then
        DRAWRECT hdc, 0, 0, UserScaleW, UserScaleH, &H80000014
        DRAWRECT hdc, 0, 0, UserScaleW + 1, UserScaleH + 1, 0
    ElseIf g_MouseIn = 1 Or m_ButtonStyle = gbStandard Then
        DRAWRECT hdc, 0, 0, UserScaleW, UserScaleH, 0
        DRAWRECT hdc, 0, 0, UserScaleW + 1, UserScaleH + 1, &H80000014
    End If
End If
CalcRECTs
If Len(m_Caption) > 0 Then DrawCaption
If m_Picture Is Nothing = False Then DrawPicture
If g_HasFocus = 1 And m_ShowFocusRect And m_ButtonStyle <> gbWinXP Then DrawFocusRect hdc, g_FocusRect
AutoRedraw = False
End Sub
'Private Sub UserControl_DblClick()
'    SetCapture hwnd
'    UserControl_MouseDown g_Button, g_Shift, g_X, g_Y
'End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then MouseDownPressed = False

If g_MouseDown = 0 Then
    If KeyCode = 32 Then
        g_MouseDown = 1
        g_MouseIn = 1
        Refresh
    End If
End If
RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then MouseDownPressed = False

If KeyCode = 32 Then
    g_MouseDown = 0
    g_MouseIn = 0
    Refresh
    'UserControl_MouseUp 1, Shift, 0, 0
    UserControl_AccessKeyPress 32     'кликать на пробел
End If
RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    g_Button = Button: g_Shift = Shift: g_X = X: g_Y = Y

MouseDownPressed = True

If Button < 2 Then
    g_MouseDown = 1
    Refresh
End If
RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim p As POINTAPI
GetCursorPos p
If g_MouseIn = 0 Then
    ReleaseCapture
    g_MouseDown = Button
    g_MouseIn = 1
    RaiseEvent MouseIn(Shift)
    Refresh
    SetCapture UserControl.hwnd
ElseIf hwnd <> WindowFromPoint(p.X, p.Y) Then
    g_MouseIn = 0
    g_MouseDown = 0
    RaiseEvent MouseOut(Shift)
    Refresh
    ReleaseCapture
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Enabled = False Then Exit Sub

g_MouseDown = 0
g_MouseIn = 0
ReleaseCapture     'last open
If Button < 2 Then
    Refresh
    Dim p As POINTAPI
    GetCursorPos p
    If hwnd = WindowFromPoint(p.X, p.Y) Then
        If MouseDownPressed Then
            If Shift <> 0 Then
                RaiseEvent ShiftClick(Shift)
            Else
                RaiseEvent Click
            End If
        End If
    End If
End If
RaiseEvent MouseUp(Button, Shift, X, Y)

MouseDownPressed = False
End Sub
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bEnabled As Boolean)
'If bEnabled = UserControl.Enabled Then Exit Property
UserControl.Enabled = bEnabled
g_HasFocus = 0
g_MouseDown = 0
g_MouseIn = 0

Refresh
End Property


Public Property Get BackColor() As OLE_COLOR
BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(New_BackColor As OLE_COLOR)
UserControl.BackColor = New_BackColor
Refresh
End Property

'Public Property Get ForeColor() As OLE_COLOR
'ForeColor = UserControl.ForeColor
'End Property
'Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
'UserControl.ForeColor = New_ForeColor
'Refresh
'End Property

Public Property Get FontBold() As Boolean
FontBold = UserControl.FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
UserControl.FontBold = New_FontBold
Refresh
End Property

Public Property Get hwnd() As Long
hwnd = UserControl.hwnd
End Property
'Public Property Get MousePointer() As MousePointerConstants
'    MousePointer = UserControl.MousePointer
'End Property
'Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
'    UserControl.MousePointer = New_MousePointer
'End Property
Public Property Get ShowFocusRect() As Boolean
ShowFocusRect = m_ShowFocusRect
End Property
Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
m_ShowFocusRect = New_ShowFocusRect
Refresh
End Property
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
m_Caption = New_Caption
SetAccessKeys
Refresh
End Property
Public Property Get ButtonStyle() As XBButtonStyle
ButtonStyle = m_ButtonStyle
End Property
Public Property Let ButtonStyle(ByVal New_ButtonStyle As XBButtonStyle)
m_ButtonStyle = New_ButtonStyle
Refresh
End Property
Public Property Get PicturePosition() As XBPicturePosition
PicturePosition = m_PicturePosition
End Property
Public Property Let PicturePosition(ByVal New_PicturePosition As XBPicturePosition)
m_PicturePosition = New_PicturePosition
Refresh
End Property
Public Property Get Picture() As Picture
Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
Set m_Picture = New_Picture
With UserControl
    If m_Picture Is Nothing = False Then
        m_PictureWidth = .ScaleX(m_Picture.Width, 8, 3)
        m_PictureHeight = .ScaleY(m_Picture.Height, 8, 3)
    End If
End With
Refresh
End Property
Private Sub CalculateCaptionRect()
Dim mvarWidth As Long, mvarHeight As Long
With mvarDrawTextParams
    .iLeftMargin = 1
    .iRightMargin = 1
    .iTabLength = 1
    .cbSize = Len(mvarDrawTextParams)
End With
DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, 1045, mvarDrawTextParams
With mvarCaptionRect
    mvarWidth = .right - .left
    mvarHeight = .bottom - .top
    .left = mvarClientRect.left + (((mvarClientRect.right - mvarClientRect.left) - (.right - .left)) * 0.5)
    .top = mvarClientRect.top + (((mvarClientRect.bottom - mvarClientRect.top) - (.bottom - .top)) * 0.5)
    .right = .left + mvarWidth
    .bottom = .top + mvarHeight
End With
End Sub
Private Sub DrawCaption()
If Enabled Then
    SetTextColor hdc, CColor(&H80000012)
    mvarOrgRect = mvarCaptionRect
    If g_MouseDown = 1 Then
        With mvarCaptionRect
            .left = mvarCaptionRect.left + 1
            .top = mvarCaptionRect.top + 1
            .right = mvarCaptionRect.right + 1
            .bottom = mvarCaptionRect.bottom + 1
        End With
    End If
    DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, 21, mvarDrawTextParams
    mvarCaptionRect = mvarOrgRect
Else
    Dim g_tmpFontColor As Long
    g_tmpFontColor = UserControl.ForeColor

    SetTextColor hdc, CColor(&H80000014)
    Dim mvarCaptionRect_Iki As RECT
    With mvarCaptionRect_Iki
        .bottom = mvarCaptionRect.bottom
        .left = mvarCaptionRect.left + 1
        .right = mvarCaptionRect.right + 1
        .top = mvarCaptionRect.top + 1
    End With
    DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect_Iki, 21, mvarDrawTextParams

    SetTextColor hdc, CColor(&H80000010)
    DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, 21, mvarDrawTextParams

    SetTextColor hdc, CColor(g_tmpFontColor)
End If
End Sub
Private Sub DrawPicture()
mvarOrgRect = mvarPictureRect

With mvarPictureRect
    .left = .left + g_MouseDown
    .top = .top + g_MouseDown
    .right = .right + g_MouseDown
    .bottom = .bottom + g_MouseDown

    Dim DC2 As Byte

    If Enabled = 0 Then
        DC2 = 35
        If m_Picture.Type = 1 Then DC2 = 36
        DrawState hdc, 0, 0, m_Picture, 0, .left, .top, 0, 0, DC2
    ElseIf m_Picture.Type = 1 Then
        Picture1.AutoRedraw = True
        Picture1.Cls
        Picture1.PaintPicture m_Picture, 0, 0
        If m_UseMaskColor = False Then m_MaskColor = CColor(GetPixel(Picture1.hdc, 0, 0))
        DoEvents
        TransparentBlt _
                UserControl.hdc, .left, .top, .right - .left, .bottom - .top, _
                Picture1.hdc, 0, 0, .right - .left, .bottom - .top, _
                m_MaskColor
        UserControl.Refresh
        Picture1.AutoRedraw = False
    ElseIf m_Picture.Type = 3 Then
        UserControl.PaintPicture m_Picture, .left, .top, .right - .left, .bottom - .top, 0, 0, m_PictureWidth, m_PictureHeight
    End If
End With

mvarPictureRect = mvarOrgRect
End Sub
Public Property Get XPColor_Pressed() As OLE_COLOR
XPColor_Pressed = m_XPColor_Pressed
End Property
Public Property Let XPColor_Pressed(ByVal New_XPColor_Pressed As OLE_COLOR)
m_XPColor_Pressed = New_XPColor_Pressed
End Property
Public Property Get XPColor_Hover() As OLE_COLOR
XPColor_Hover = m_XPColor_Hover
End Property
Public Property Let XPColor_Hover(ByVal New_XPColor_Hover As OLE_COLOR)
m_XPColor_Hover = New_XPColor_Hover
End Property
Public Property Get XPDefaultColors() As Boolean
XPDefaultColors = m_XPDefaultColors
End Property
Public Property Let XPDefaultColors(ByVal New_XPDefaultColors As Boolean)
m_XPDefaultColors = New_XPDefaultColors
End Property
Public Property Get MaskColor() As OLE_COLOR
MaskColor = m_MaskColor
End Property
Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
m_MaskColor = New_MaskColor
Refresh
End Property
Public Property Get UseMaskColor() As Boolean
UseMaskColor = m_UseMaskColor
End Property
Public Property Let UseMaskColor(ByVal New_V As Boolean)
m_UseMaskColor = New_V
Refresh
End Property
Private Sub DRAWRECT(DestHDC As Long, ByVal RectLEFT As Long, _
        ByVal RectTOP As Long, _
        ByVal RectRIGHT As Long, ByVal RectBOTTOM As Long, _
        ByVal MyColor As Long, _
        Optional FillRectWithColor As Byte = 0)
Dim MyRect As RECT, Firca As Long
Firca = CreateSolidBrush(CColor(MyColor))
With MyRect
    .left = RectLEFT
    .top = RectTOP
    .right = RectRIGHT
    .bottom = RectBOTTOM
End With
If FillRectWithColor = 1 Then
    FillRect DestHDC, MyRect, Firca
Else
    FrameRect DestHDC, MyRect, Firca
End If
DeleteObject Firca
End Sub
Private Sub DrawWinXPButton(ByVal Press As Byte, Optional HOVERING As Byte)
Dim X As Long, Intg As Single, curBackColor As Long

curBackColor = COLOR_DarkenColor(CColor(&H8000000F), 48)
If Enabled Then
    If m_XPDefaultColors = True Then
        m_XPColor_Pressed = RGB(140, 170, 230)
        m_XPColor_Hover = RGB(225, 153, 71)
    End If

    If UserScaleH = 0 Then Exit Sub
    If Press = 0 Then
        Intg = 50 / UserScaleH
        For X = 1 To UserScaleH
            'Line (0, x)-(UserScaleW, x), COLOR_DarkenColor(vbWhite, -Intg * x) 'And UserControl.BackColor
            Line (0, X)-(UserScaleW, X), COLOR_DarkenColor(BackColor, -Intg * X)       'And UserControl.BackColor
        Next

        DRAWRECT hdc, 0, 0, UserScaleW, UserScaleH, &H80000015

        If HOVERING = 1 Or g_HasFocus = 1 Then
            Intg = CColor(IIf(HOVERING, m_XPColor_Hover, m_XPColor_Pressed))
            DRAWRECT hdc, 1, 2, UserScaleW - 1, UserScaleH - 2, Intg

            Line (2, UserScaleH - 2)-(UserScaleW - 2, UserScaleH - 2), COLOR_DarkenColor(Intg, -40)
            Line (2, 1)-(UserScaleW - 2, 1), COLOR_DarkenColor(Intg, 90)
            Line (1, 2)-(UserScaleW - 1, 2), COLOR_DarkenColor(Intg, 35)
            curBackColor = COLOR_DarkenColor(Intg, 20)
            Line (2, 3)-(2, UserScaleH - 3), curBackColor
            Line (UserScaleW - 3, 3)-(UserScaleW - 3, UserScaleH - 3), curBackColor
            SetPixel hdc, 3, UserScaleH - 4, Intg
            SetPixel hdc, UserScaleW - 4, UserScaleH - 4, Intg
            Intg = COLOR_DarkenColor(Intg, 35)
            SetPixel hdc, UserScaleW - 4, 3, Intg
            SetPixel hdc, 3, 3, Intg
        End If
    Else
        Intg = 25 / UserScaleH
        curBackColor = COLOR_DarkenColor(curBackColor, -32)
        For X = 1 To UserScaleH
            Line (0, UserScaleH - X)-(UserScaleW, UserScaleH - X), COLOR_DarkenColor(curBackColor, -Intg * X)
        Next

        DRAWRECT hdc, 0, 0, UserScaleW, UserScaleH, &H80000015
    End If
    Intg = &H80000015
Else
    DRAWRECT hdc, 0, 0, UserScaleW, UserScaleH, COLOR_DarkenColor(curBackColor, -24), 1
    DRAWRECT hdc, 0, 0, UserScaleW, UserScaleH, COLOR_DarkenColor(curBackColor, -84)
    Intg = COLOR_DarkenColor(curBackColor, -72)
End If

'закругления
curBackColor = CColor(&H8000000F)
Line (0, 0)-(1, 1), curBackColor, BF
SetPixel hdc, 1, 1, Intg

Line (0, UserScaleH - 2)-(1, UserScaleH), curBackColor, BF
SetPixel hdc, 1, UserScaleH - 2, Intg

Line (UserScaleW - 2, 0)-(UserScaleW, 1), curBackColor, BF
SetPixel hdc, UserScaleW - 2, 1, Intg

Line (UserScaleW - 2, UserScaleH - 2)-(UserScaleW, UserScaleH), curBackColor, BF
SetPixel hdc, UserScaleW - 2, UserScaleH - 2, Intg
End Sub
Private Sub SetAccessKeys()
Dim ampersandPos As Long
With UserControl
    If Len(m_Caption) > 1 Then
        ampersandPos = InStr(1, m_Caption, "&", vbTextCompare)
        If (ampersandPos < Len(m_Caption)) And (ampersandPos > 0) Then
            If Mid$(m_Caption, ampersandPos + 1, 1) <> "&" Then
                .AccessKeys = LCase(Mid$(m_Caption, ampersandPos + 1, 1))
            Else
                ampersandPos = InStr(ampersandPos + 2, m_Caption, "&", vbTextCompare)
                If Mid$(m_Caption, ampersandPos + 1, 1) <> "&" Then
                    .AccessKeys = LCase(Mid$(m_Caption, ampersandPos + 1, 1))
                Else
                    .AccessKeys = ""
                End If
            End If
        Else
            .AccessKeys = ""
        End If
    Else
        .AccessKeys = ""
    End If
End With
End Sub
Private Function CColor(ByVal clr As OLE_COLOR) As Long
' If it's a system color, get the RGB value.
If clr And &H80000000 Then CColor = GetSysColor(clr And (Not &H80000000)) Else CColor = clr
End Function
Private Function COLOR_DarkenColor(ByVal color As Long, ByVal Value As Long) As Long
Dim cc As RGB, R As Integer, g As Integer, b As Integer

CopyMemory ByVal VarPtr(cc), ByVal VarPtr(color), 3
With cc
    b = .Blue + Value
    g = .Green + Value
    R = .Red + Value
End With
If R < 0 Then R = 0
If R > 255 Then R = 255
If g < 0 Then g = 0
If g > 255 Then g = 255
If b < 0 Then b = 0
If b > 255 Then b = 255
COLOR_DarkenColor = RGB(R, g, b)
End Function
