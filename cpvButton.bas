Attribute VB_Name = "UCButtons"
Option Explicit

'Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long


'Public Declare Function TransparentBlt Lib "msimg32" _
                (ByVal hdcDst As Long, ByVal nXOriginDst As Long, _
                 ByVal nYOriginDst As Long, ByVal nWidthDst As Long, _
                 ByVal nHeightDst As Long, ByVal hdcSrc As Long, _
                 ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
                 ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
                 ByVal crTransparent As Long) As Long
                 
                 
'Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
' DrawIconEx constants
'Public Const DI_MASK = &H1
'Public Const DI_IMAGE = &H2
'Public Const DI_NORMAL = &H3
'Public Const DI_COMPAT = &H4
'Public Const DI_DEFAULTSIZE = &H8


'Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
' DrawText constants
'Public Const DT_BOTTOM = &H8
'Public Const DT_CALCRECT = &H400
'Public Const DT_CENTER = &H1
'Public Const DT_EXPANDTABS = &H40
'Public Const DT_EXTERNALLEADING = &H200
'Public Const DT_INTERNAL = &H1000
'Public Const DT_LEFT = &H0
'Public Const DT_NOCLIP = &H100
'Public Const DT_NOPREFIX = &H800
'Public Const DT_RIGHT = &H2
'Public Const DT_SINGLELINE = &H20
'Public Const DT_TABSTOP = &H80
'Public Const DT_TOP = &H0
'Public Const DT_VCENTER = &H4
'Public Const DT_WORDBREAK = &H10

'Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Public Type POINTAPI
'    X As Long
'    Y As Long
'End Type


Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'Public Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type

Public g_Light As OLE_COLOR
Public g_Shadow As OLE_COLOR
Public g_HighLight As OLE_COLOR
Public g_DarkShadow As OLE_COLOR
'Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'Public Const BUTTON_NONE = 0
'Public Const BUTTON_UP = 1
'Public Const BUTTON_DOWN = 2

'Public Const BACKGROUND_COLOR = &H80000010
'Public Const BDR_RAISEDOUTER = &H1
'Public Const BDR_SUNKENOUTER = &H2
'Public Const BDR_RAISEDINNER = &H4
'Public Const BDR_SUNKENINNER = &H8

'Public Const BDR_OUTER = &H3
'Public Const BDR_INNER = &HC
'Public Const BDR_RAISED = &H5
'Public Const BDR_SUNKEN = &HA

'Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
'Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
'Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
'Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

'Public Const BF_LEFT = &H1
'Public Const BF_TOP = &H2
'Public Const BF_RIGHT = &H4
'Public Const BF_BOTTOM = &H8

'Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
'Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
'Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
'Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
'Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

'Public Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
'Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
'Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
'Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
'Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

'Public Const BF_MIDDLE = &H800    ' Fill in the middle.
'Public Const BF_SOFT = &H1000     ' Use for softer buttons.
'Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
'Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
'Public Const BF_MONO = &H8000     ' For monochrome borders.

'Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
'Public Declare Function PtInRect Lib "user32" (RECT As RECT, ByVal lPtX As Long, ByVal lPtY As Long) As Integer




'——————————————————————————————————————————————————————————————————
' The next code has extracted from vbAccelerator:
' Hue Lightness and Saturation (HLS) Model and Manipulating Colours
' http://www.vbaccelerator.com/codelib/gfx/clrman1.htm
'——————————————————————————————————————————————————————————————————
' I use it to convert RGB to HLS, and varying
' L value, reconvert new HLS to RGB
'——————————————————————————————————————————————————————————————————

Public Sub RGBToHSL(R As Long, G As Long, B As Long, _
        H As Single, s As Single, l As Single)

Dim Max As Single
Dim min As Single
Dim Delta As Single
Dim rR As Single, rG As Single, rB As Single

rR = R / 255: rG = G / 255: rB = B / 255

'Given:   RGB each in [0,1].
'Desired: H in [0,360] and S in [0,1], except if S=0, then H=UNDEFINED.

Max = Maximum(rR, rG, rB)
min = Minimum(rR, rG, rB)
l = (Max + min) / 2     'Lightness

If Max = min Then

    'Acrhomatic case:

    s = 0
    H = 0

Else

    'Chromatic case:

    'First calculate the saturation

    If l <= 0.5 Then
        s = (Max - min) / (Max + min)
    Else
        s = (Max - min) / (2 - Max - min)
    End If

    'Next calculate the hue

    Delta = Max - min

    If rR = Max Then
        H = (rG - rB) / Delta         'Resulting color is between yellow and magenta
    ElseIf rG = Max Then
        H = 2 + (rB - rR) / Delta     'Resulting color is between cyan and yellow
    ElseIf rB = Max Then
        H = 4 + (rR - rG) / Delta     'Resulting color is between magenta and cyan
    End If

End If

End Sub

Public Sub HSLToRGB(H As Single, s As Single, l As Single, _
        R As Long, G As Long, B As Long)

Dim rR As Single, rG As Single, rB As Single
Dim min As Single, Max As Single

If s = 0 Then

    'Achromatic case:

    rR = l: rG = l: rB = l

Else

    'Chromatic case:

    'Delta = Max-Min
    If l <= 0.5 Then
        'S = (Max - Min) / (Max + Min)
        'Get Min value:
        min = l * (1 - s)
    Else
        'S = (Max - Min) / (2 - Max - Min)
        'Get Min value:
        min = l - s * (1 - l)
    End If
    'Get the Max value
    Max = 2 * l - min

    'Now depending on sector we can evaluate the H,L,S:
    If (H < 1) Then

        rR = Max

        If (H < 0) Then
            rG = min
            rB = rG - H * (Max - min)
        Else
            rB = min
            rG = H * (Max - min) + rB
        End If

    ElseIf (H < 3) Then

        rG = Max

        If (H < 2) Then
            rB = min
            rR = rB - (H - 2) * (Max - min)
        Else
            rR = min
            rB = (H - 2) * (Max - min) + rR
        End If

    Else

        rB = Max

        If (H < 4) Then
            rR = min
            rG = rR - (H - 4) * (Max - min)
        Else
            rG = min
            rR = (H - 4) * (Max - min) + rG
        End If

    End If

End If

R = rR * 255: G = rG * 255: B = rB * 255

End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
If (rR > rG) Then
 If (rR > rB) Then
  Maximum = rR
 Else
  Maximum = rB
 End If
Else
 If (rB > rG) Then
  Maximum = rB
 Else
  Maximum = rG
 End If
End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
If (rR < rG) Then
 If (rR < rB) Then
  Minimum = rR
 Else
  Minimum = rB
 End If
Else
 If (rB < rG) Then
  Minimum = rB
 Else
  Minimum = rG
 End If
End If
End Function


Public Sub Draw3DEffect(p As PictureBox, flag As Integer)
'0 - up
'1-down
'2-erase
Dim g_3DInc As Integer
'Dim t As Long, l As Long

'Get3DButtonColors p.BackColor


g_3DInc = 1 'Else g_3DInc = 0
    
    'Draw 'edge':
't = p.Top: l = p.Left
Select Case flag
Case 1
        
        'This style...:
'p.Line (g_3DInc, g_3DInc)-(p.ScaleWidth - g_3DInc, p.ScaleHeight - g_3DInc), g_Shadow, B
'p.Line (g_3DInc, g_3DInc)-(p.ScaleWidth - g_3DInc, p.ScaleHeight - g_3DInc), g_Light, B
'p.Line (g_3DInc, g_3DInc)-(p.ScaleWidth - g_3DInc, p.ScaleHeight - g_3DInc), g_DarkShadow, B
'p.Line (g_3DInc, g_3DInc)-(p.ScaleWidth - g_3DInc, p.ScaleHeight - g_3DInc), g_HighLight, B
        'or classic...:
p.Line (1, 1)-(p.ScaleWidth - 2, p.ScaleHeight - 2), g_Shadow, B
        'Remove property ShowBorderOnFocus (=True: Correct classic effect)
Case 0
    
p.Line (1 + g_3DInc, 1 + g_3DInc)-(p.ScaleWidth - g_3DInc - 1, p.ScaleHeight - g_3DInc - 1), g_Light, B
p.Line (0 + g_3DInc, 0 + g_3DInc)-(p.ScaleWidth - g_3DInc - 2, p.ScaleHeight - g_3DInc - 2), g_Shadow, B
p.Line (0 + g_3DInc, 0 + g_3DInc)-(p.ScaleWidth - g_3DInc - 0, p.ScaleHeight - g_3DInc - 0), g_HighLight, B
p.Line (-1 + g_3DInc, -1 + g_3DInc)-(p.ScaleWidth - g_3DInc - 1, p.ScaleHeight - g_3DInc - 1), g_DarkShadow, B
    
Case 2
p.Line (1 + g_3DInc, 1 + g_3DInc)-(p.ScaleWidth - g_3DInc - 1, p.ScaleHeight - g_3DInc - 1), p.BackColor, B
p.Line (0 + g_3DInc, 0 + g_3DInc)-(p.ScaleWidth - g_3DInc - 2, p.ScaleHeight - g_3DInc - 2), p.BackColor, B
p.Line (0 + g_3DInc, 0 + g_3DInc)-(p.ScaleWidth - g_3DInc - 0, p.ScaleHeight - g_3DInc - 0), p.BackColor, B
p.Line (-1 + g_3DInc, -1 + g_3DInc)-(p.ScaleWidth - g_3DInc - 1, p.ScaleHeight - g_3DInc - 1), p.BackColor, B

    End Select
    
    'Draw black border:
    
'Line (0, 0)-(p.ScaleWidth - 1, p.ScaleHeight - 1), vbBlack, B

End Sub
Public Sub Get3DButtonColors(color As OLE_COLOR)
        
    Dim g_R As Long, g_G As Long, g_B As Long
    Dim H As Single, s As Single, l As Single
    Dim R As Long, G As Long, B As Long
    
    'If SystemColorConstant then get color:
    If color < 0 Then color = GetSysColor(color And Not &H80000000)

    'Convert Long value to R,G,B values:
    g_R = (color And &HFF&)
    g_G = (color And &HFF00&) \ &H100&
    g_B = (color And &HFF0000) \ &H10000
    
    'Get H,S,L values:
    RGBToHSL g_R, g_G, g_B, H, s, l
    
    'Get 3DColor values (on L)
    HSLToRGB H, s, l + (1 - l) / 8, R, G, B
        g_Light = RGB(R, G, B)
    HSLToRGB H, s, l + (1 - l) / 2, R, G, B
        g_HighLight = RGB(R, G, B)
    HSLToRGB H, s, l / 1.5, R, G, B
        g_Shadow = RGB(R, G, B)
    HSLToRGB H, s, l / 3.5, R, G, B
        g_DarkShadow = RGB(R, G, B)
    
End Sub






