VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTipEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (ByRef TLPINITCOMMONCONTROLSEX As InitCommonControlsEx) As Long
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Enum gbHWTitleIco
    icoTTNone = &H0
    icoTTInfo = &H1
    icoTTWarning = &H2
    icoTTError = &H3
End Enum
Public Enum TOOLSTYLE
    TTS_ALWAYSTIP = &H1
    TTS_NOPREFIX = &H2
    TTS_NOANIMATE = &H10
    TTS_NOFADE = &H20
    TTS_BALLOON = &H40
End Enum
Public Enum gbHWTTShType
    TTF_IDISHWND = &H1
    TTF_CENTERTIP = &H2
    TTF_SUBCLASS = &H10
    TTF_TRACK = &H20
    TTF_ABSOLUTE = &H80
    TTF_TRANSPARENT = &H100
    TTF_DI_SETITEM = &H8000
End Enum
Public Enum gbHWTTDelay
    TTDT_AUTOMATIC = &H0
    TTDT_RESHOW = &H1
    TTDT_AUTOPOP = &H2
    TTDT_INITIAL = &H3
End Enum

'mzt Private Type InitCommonControlsEx
'    dwSize          As Long
'    dwICC           As Long
'End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type TOOLINFO
    cbSize As Long
    uFlags As gbHWTTShType
    hWnd As Long
    uId As Long
    rct As RECT
    hInst As Long
    lpszText As String
    lParam As Long
End Type

Private Const CW_USEDEFAULT As Long = &H80000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOPMOST As Long = &H8&


' ToolTip Messages
Private Const WM_USER = &H400
'Private Const TTM_NEWTOOLRECT = (WM_USER + 6)
'Private Const TTM_GETTOOLINFO = (WM_USER + 8)
'Private Const TTM_TRACKACTIVATE = WM_USER + 17
'Private Const TTM_TRACKPOSITION = WM_USER + 18

Private Const TTM_SETTITLE As Long = &H420
Private Const TTM_SETTIPBKCOLOR As Long = &H413
Private Const TTM_SETTIPTEXTCOLOR As Long = &H414
Private Const TTM_SETDELAYTIME As Long = &H403
Private Const TTM_ADDTOOL As Long = &H404
Private Const TTM_DELTOOL As Long = &H405
Private Const TTM_GETMAXTIPWIDTH As Long = &H419

Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_GETTOOLCOUNT As Long = &H40D
Private Const TTM_UPDATETIPTEXT As Long = &H40C
Private Const TTM_ENUMTOOLS As Long = &H40E
Private Const TTM_GETTIPBKCOLOR As Long = &H416
Private Const TTM_GETDELAYTIME As Long = &H415

Public hToolTip As Long

Public Sub CreateToolTip(ByVal FormHandle As Long, _
        Optional ByVal inTTStyle As TOOLSTYLE = &H0, _
        Optional ByVal inTTIcon As gbHWTitleIco = icoTTNone, _
        Optional ByVal strTTTitle As String = vbNullString, _
        Optional ByVal iTTWidth As Long = &HFFFF, _
        Optional ByVal inTTDelayTime As Long = &HFFFF, _
        Optional ByVal inTTVisibleTime As Long = &HFFFF, _
        Optional ByVal lBackColor As ColorConstants = &HFFFF, _
        Optional ByVal lForeColor As ColorConstants = &HFFFF)

'Dim iccex As InitCommonControlsEx
'With iccex
'   .dwSize = Len(iccex)
'    iccex.dwICC = &HFF
'End With
'Call InitCommonControlsEx(iccex)

hToolTip = CreateWindowEx(WS_EX_TOPMOST, "tooltips_class32", &H0, WS_POPUP Or inTTStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, FormHandle, &H0, App.hInstance, &H0)
'hToolTip = CreateWindowEx(0, "tooltips_class32", &H0, WS_POPUP Or inTTStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, FormHandle, &H0, App.hInstance, &H0)

Call SetWindowPos(hToolTip, &HFFFF, &H0, &H0, &H0, &H0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)

If Not inTTIcon = &H0 Or Len(strTTTitle) Then _
        Call SendMessage(hToolTip, TTM_SETTITLE, inTTIcon, ByVal strTTTitle)
'    If Not iTTWidth = &HFFFF Then _
     Call SendMessage(hToolTip, TTM_SETMAXTIPWIDTH, iTTWidth, &H0)
If Not iTTWidth = &HFFFF Then _
        Call SendMessage(hToolTip, TTM_SETMAXTIPWIDTH, 0, iTTWidth)    ', &H0)

If Not inTTDelayTime = &HFFFF Then _
        Call SendMessage(hToolTip, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal inTTDelayTime)
If Not inTTVisibleTime = &HFFFF Then _
        Call SendMessage(hToolTip, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal inTTVisibleTime)
If Not lBackColor = &HFFFF Then _
        Call SendMessage(hToolTip, TTM_SETTIPBKCOLOR, lBackColor, ByVal &H0)
If Not lForeColor = &HFFFF Then _
        Call SendMessage(hToolTip, TTM_SETTIPTEXTCOLOR, lForeColor, ByVal &H0)

'SetParent FrmMain.Picture1.hwnd, hToolTip
End Sub

Public Function AddToolTip(ByVal CtlhWnd As Long, _
        ByVal strTTText As String, _
        Optional ByVal inTTshParam As gbHWTTShType)

Dim ti As TOOLINFO     'mzt , rct As RECT

With ti
    .uFlags = TTF_SUBCLASS Or TTF_IDISHWND Or inTTshParam
    'TTF_IDISHWND Or TTF_TRACK Or TTF_ABSOLUTE Or TTF_SUBCLASS Or inTTshParam
    ''TTF_SUBCLASS Or TTF_IDISHWND Or inTTshParam
    .hWnd = CtlhWnd
    .uId = CtlhWnd
    .hInst = App.hInstance
    .lpszText = strTTText
    'Debug.Print .lpszText
    'Debug.Print strTTText
    Call GetClientRect(CtlhWnd, .rct)
    .cbSize = Len(ti)
End With

'SendMessage hToolTip, TTM_TRACKPOSITION, 0, MakLng(0, 0)

SendMessage hToolTip, TTM_ADDTOOL, &H0, ti

'SendMessage hToolTip, TTM_TRACKACTIVATE, ByVal True, ti

End Function

Public Sub DestroyToolTip()
Call DestroyWindow(hToolTip): hToolTip = &H0
End Sub

Public Function RemoveToolTip(CtlhWnd As Long)
Dim ti As TOOLINFO

If hToolTip = &H0 Then Exit Function
If Not GetToolInfo(CtlhWnd, ti) Then Exit Function

Call SendMessage(hToolTip, TTM_DELTOOL, &H0, ti)
RemoveToolTip = True
End Function


Public Property Get ToolTipCount() As Integer                   ' Количество зарегистрированных подсказок
If hToolTip = &H0 Then Exit Property

ToolTipCount = SendMessage(hToolTip, TTM_GETTOOLCOUNT, 0, 0)
End Property

Public Property Get MaxToolTipWidth() As Integer                ' Максимальная длинна подсказки
If hToolTip = &H0 Then Exit Property

MaxToolTipWidth = LoWord(SendMessage(hToolTip, TTM_GETMAXTIPWIDTH, 0, 0))
End Property
Public Property Let MaxToolTipWidth(ByVal cx As Integer)
If hToolTip = &H0 Then Exit Property

Call SendMessage(hToolTip, TTM_SETMAXTIPWIDTH, &H0, ByVal CLng(cx))
End Property

Public Property Get ToolTipText(CtrlhWnd As Long) As String     ' Текст подсказки
Dim ti As TOOLINFO
If Not GetToolInfo(CtrlhWnd, ti) Then Exit Property

ToolTipText = ti.lpszText
End Property
Public Property Let ToolTipText(CtrlhWnd As Long, sText As String)
Dim ti As TOOLINFO
If Not GetToolInfo(CtrlhWnd, ti) Then Exit Property

ti.lpszText = sText & vbNullChar
Call SendMessage(hToolTip, TTM_UPDATETIPTEXT, &H0, ti)
End Property

Public Property Get BackColor() As ColorConstants               ' Цвет фона
BackColor = SendMessage(hToolTip, TTM_GETTIPBKCOLOR, &H0, &H0)
End Property
Public Property Let BackColor(inColor As ColorConstants)
Call SendMessage(hToolTip, TTM_SETTIPBKCOLOR, inColor, &H0)
End Property

Public Property Get ForeColor() As ColorConstants               ' Цвет обводки и надписей
ForeColor = SendMessage(hToolTip, TTM_SETTIPTEXTCOLOR, &H0, &H0)
End Property
Public Property Let ForeColor(inColor As ColorConstants)
Call SendMessage(hToolTip, TTM_SETTIPTEXTCOLOR, inColor, &H0)
End Property

Public Property Get DelayTime(dwType As gbHWTTDelay) As Long    ' Задержка и ее тип
DelayTime = SendMessage(hToolTip, TTM_GETDELAYTIME, dwType, &H0)
End Property
Public Property Let DelayTime(dwType As gbHWTTDelay, dwMilliSecs As Long)
Call SendMessage(hToolTip, TTM_SETDELAYTIME, dwType, ByVal dwMilliSecs)
End Property


Private Function GetToolInfo(ByVal hWndToolTip As Long, _
        ByRef ti As TOOLINFO) As Boolean
Dim li As Long

ti.cbSize = Len(ti)
ti.lpszText = String$(81, 0)

For li = &H0 To ToolTipCount - &H1
    If SendMessage(hToolTip, TTM_ENUMTOOLS, li, ti) Then If hWndToolTip = ti.uId Then GetToolInfo = True: Exit For
Next
On Error Resume Next
ti.lpszText = Left$(ti.lpszText, InStr(1, ti.lpszText, vbNullChar) - &H1)
End Function

Private Function LoWord(DWord As Long) As Integer
LoWord = IIf(DWord And &H8000&, DWord Or &HFFFF0000, DWord And &HFFFF&)
End Function


Private Sub Class_Terminate()
Call DestroyToolTip
End Sub
Public Function MakLng(X As Integer, Y As Integer) As Long
MakLng = X + (Y * &H10000)
End Function
