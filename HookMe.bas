Attribute VB_Name = "ModHook"
Option Explicit

'да-нет автозапуску CD
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Const RegMsg As String = "QueryCancelAutoPlay"
Public m_RegMsg As Long


Public Declare Function GetProp Lib "user32" _
                                Alias "GetPropA" _
                                (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Declare Function CallWindowProc Lib "user32" _
                                       Alias "CallWindowProcA" _
                                       (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
                                        ByVal Msg As Long, ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long

Private Declare Function SetProp Lib "user32" _
                                 Alias "SetPropA" _
                                 (ByVal hwnd As Long, ByVal lpString As String, _
                                  ByVal hData As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
                                      Alias "SetWindowLongA" _
                                      (ByVal hwnd As Long, ByVal nIndex As Long, _
                                       ByVal wNewWord As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
                                       Alias "GetWindowLongA" _
                                       (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
                               Alias "RtlMoveMemory" _
                               (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GWL_WNDPROC As Long = (-4)

'для FrmMain -------------------------------------------------
Public Function HookFunc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Dim foo As Long
Dim obj As FrmMain     'MUST be the correct name of the form

foo = GetProp(hwnd, "ObjectPointer")
'
' Ignore "impossible" bogus case
'
If (foo <> 0) Then
    CopyMemory obj, foo, 4
    On Error Resume Next
    HookFunc = obj.WindowProc(hwnd, Msg, wp, lp)
    If (err) Then
        UnhookWindow hwnd
'Debug.Print "Unhook on Error, #"; CStr(err.Number)
'Debug.Print "  Desc: "; err.Description
'Debug.Print "  Message, hWnd: &h"; Hex(hWnd), "Msg: &h"; Hex(Msg), "Params:"; wp; lp
    End If
    '
    ' Make sure we don't get any foo->Release() calls
    '
    foo = 0
    CopyMemory obj, foo, 4
End If

End Function


Public Sub HookWindow(hwnd As Long, thing As Object)
Dim foo As Long
CopyMemory foo, thing, 4
Call SetProp(hwnd, "ObjectPointer", foo)
Call SetProp(hwnd, "OldWindowProc", GetWindowLong(hwnd, GWL_WNDPROC))
Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf HookFunc)
End Sub


Public Sub UnhookWindow(hwnd As Long)
Dim foo As Long
foo = GetProp(hwnd, "OldWindowProc")
If (foo <> 0) Then
    Call SetWindowLong(hwnd, GWL_WNDPROC, foo)
End If
End Sub


Public Function InvokeWindowProc(hwnd As Long, Msg As Long, wp As Long, lp As Long) As Long
InvokeWindowProc = CallWindowProc(GetProp(hwnd, "OldWindowProc"), hwnd, Msg, wp, lp)
End Function
'все с FrmMain -------------------------------------------------


Public Sub ForceTextBoxNumeric(TextBox As TextBox, Optional Force As Boolean = True)
    Dim style As Long
    Const GWL_STYLE = (-16)
    Const ES_NUMBER = &H2000
    
    ' get current style
    style = GetWindowLong(TextBox.hwnd, GWL_STYLE)
    If Force Then
        style = style Or ES_NUMBER
    Else
        style = style And Not ES_NUMBER
    End If
    ' enforce new style
    SetWindowLong TextBox.hwnd, GWL_STYLE, style
End Sub

'для frmAuto ------------------------------------------------
Public Function HookFuncAutoAdd(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Dim foo As Long
Dim obj As frmAuto     'MUST be the correct name of the form
foo = GetProp(hwnd, "ObjectPointer")
'
' Ignore "impossible" bogus case
'
If (foo <> 0) Then
    CopyMemory obj, foo, 4
    On Error Resume Next
    HookFuncAutoAdd = obj.WindowProc(hwnd, Msg, wp, lp)
    If (err) Then
        UnhookWindowAutoAdd hwnd

'Debug.Print "Unhook on Error, #"; CStr(err.Number)
'Debug.Print "  Desc: "; err.Description
'Debug.Print "  Message, hWnd: &h"; Hex(hWnd), "Msg: &h"; Hex(Msg), "Params:"; wp; lp
    End If
    '
    ' Make sure we don't get any foo->Release() calls
    '
    foo = 0
    CopyMemory obj, foo, 4
End If

End Function


Public Sub HookWindowAutoAdd(hwnd As Long, thing As Object)
Dim foo As Long
CopyMemory foo, thing, 4
Call SetProp(hwnd, "ObjectPointer", foo)
Call SetProp(hwnd, "OldWindowProc", GetWindowLong(hwnd, GWL_WNDPROC))
Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf HookFuncAutoAdd)
End Sub

Public Sub UnhookWindowAutoAdd(hwnd As Long)
Dim foo As Long
foo = GetProp(hwnd, "OldWindowProc")
If (foo <> 0) Then
    Call SetWindowLong(hwnd, GWL_WNDPROC, foo)
End If
End Sub

Public Function InvokeWindowProcAutoAdd(hwnd As Long, Msg As Long, wp As Long, lp As Long) As Long
InvokeWindowProcAutoAdd = CallWindowProc(GetProp(hwnd, "OldWindowProc"), hwnd, Msg, wp, lp)
End Function
'---------------------------------------------


'для frmEditor ------------------------------------------------
Public Function HookFuncEditor(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Dim foo As Long
Dim obj As frmEditor     'MUST be the correct name of the form
foo = GetProp(hwnd, "ObjectPointer")
'
' Ignore "impossible" bogus case
'
If (foo <> 0) Then
    CopyMemory obj, foo, 4
    On Error Resume Next
    HookFuncEditor = obj.WindowProc(hwnd, Msg, wp, lp)
    If (err) Then
        UnhookWindowEditor hwnd

'Debug.Print "Unhook on Error, #"; CStr(err.Number)
'Debug.Print "  Desc: "; err.Description
'Debug.Print "  Message, hWnd: &h"; Hex(hWnd), "Msg: &h"; Hex(Msg), "Params:"; wp; lp
    End If
    '
    ' Make sure we don't get any foo->Release() calls
    '
    foo = 0
    CopyMemory obj, foo, 4
End If

End Function

Public Sub HookWindowEditor(hwnd As Long, thing As Object)
Dim foo As Long
CopyMemory foo, thing, 4
Call SetProp(hwnd, "ObjectPointer", foo)
Call SetProp(hwnd, "OldWindowProc", GetWindowLong(hwnd, GWL_WNDPROC))
Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf HookFuncEditor)
End Sub

Public Sub UnhookWindowEditor(hwnd As Long)
Dim foo As Long
foo = GetProp(hwnd, "OldWindowProc")
If (foo <> 0) Then
    Call SetWindowLong(hwnd, GWL_WNDPROC, foo)
End If
End Sub

Public Function InvokeWindowProcEditor(hwnd As Long, Msg As Long, wp As Long, lp As Long) As Long
InvokeWindowProcEditor = CallWindowProc(GetProp(hwnd, "OldWindowProc"), hwnd, Msg, wp, lp)
End Function
'--------------------------------------
