Attribute VB_Name = "ModLVSubClass"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Contains (Most of) the Windows 32bit API declares    ''
'' required to make a custom drawn listview.            ''
''                                                      ''
'' Created By      : Sean Young                         ''
'' Additional Code : Bryan Stafford - See ReadMe        ''
'' Created on      : 14 Feburary 2002 (in its present   ''
''                   form)                              ''
''                                                      ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 'NOTE: If your app uses some of these declares elsewhere you may move them
 '      to other modules without problem, just ensure that they stay Private

Public LVHighlight As ItemColourType 'цвет выделения
Public LVItemColor As ItemColourType 'цвет строк, черезстрочно


 'Generic WM_NOTIFY notification codes for common controls
Private Enum WinNotifications
    NM_FIRST = (-0&)              ' (0U-  0U)       ' // generic to all controls
    NM_LAST = (-99&)              ' (0U- 99U)
    NM_OUTOFMEMORY = (NM_FIRST - 1&)
    NM_CLICK = (NM_FIRST - 2&)
    NM_DBLCLK = (NM_FIRST - 3&)
    NM_RETURN = (NM_FIRST - 4&)
    NM_RCLICK = (NM_FIRST - 5&)
    NM_RDBLCLK = (NM_FIRST - 6&)
    NM_SETFOCUS = (NM_FIRST - 7&)
    NM_KILLFOCUS = (NM_FIRST - 8&)
    NM_CUSTOMDRAW = (NM_FIRST - 12&)
    NM_HOVER = (NM_FIRST - 13&)
End Enum

 'constant used to get the address of the window procedure for the subclassed
 'window
Private Const GWL_WNDPROC As Long = (-4&)
 'The notification message
Private Const WM_NOTIFY As Long = &H4E&
Private Const WM_PARENTNOTIFY = &H210
'Private Const WM_CAPTURECHANGED As Long = &H215
'Private Const WM_VSCROLL = &H115
'Private Const WM_HSCROLL = &H114

 'Constants telling us whats going on
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_PREPAINT As Long = &H1&
'Private Const CDDS_POSTPAINT As Long = &H2&
Private Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
'Private Const CDDS_ITEMPOSTPAINT = (CDDS_ITEM Or CDDS_POSTPAINT)
'Private Const CDDS_SUBITEM As Long = &H20000

 'Constants we send to the control to tell it what we want it to do
'Private Const CDRF_NEWFONT As Long = &H2&
Private Const CDRF_DODEFAULT As Long = &H0&
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20&
'Private Const CDRF_NOTIFYPOSTPAINT As Long = &H10&

Private Const CDIS_FOCUS As Long = &H10
'Private Const CDIS_SELECTED As Long = &H1

 'The NMHDR structure contains information about a notification message.
 'The pointer to this structure is specified as the lParam member of a
 'WM_NOTIFY message.
Private Type NMHDR
    hWndFrom As Long ' Window handle of control sending message
    idFrom As Long   ' Identifier of control sending message
    code  As Long    ' Specifies the notification code
End Type
  
 'sub struct of the NMCUSTOMDRAW struct
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
  
 'generic customdraw struct
Private Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hdc As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
End Type
  
 'listview specific custom draw struct
Private Type NMLVCUSTOMDRAW
    nmcd As NMCUSTOMDRAW
    clrText As Long
    clrTextBk As Long
     'If Internet explorer 4.0 or higher is not present
     'do not use this member:
    'iSubItem As Integer
End Type
    
 'Function used to manipulate memory data
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
  
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc&, ByVal hObject&) As Long
  
 'Tells us which control has the focus
'Private Declare Function GetFocus Lib "user32" () As Long

 'API call to alter the class data for a window
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

 'Function used to call the next window procedure in the "chain" for the subclassed
 'window
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Holds the code for the custom drawn listview.        ''
''                                                      ''
'' Created By      : Sean Young                         ''
'' Additional Code : Bryan Stafford                     ''
'' Created on      : 14 Feburary 2002 (in its present   ''
''                   form)                              ''
''                                                      ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type ItemColourType
    ForeGround As Long
    BackGround As Long
End Type

 'this variable holds a pointer to the original message handler. We MUST save it so
 'that it can be restored before we exit the app.
Private g_addProcOld As Long

 'This is the listview currently being delt with
'Private CDLV As ListView
 'Stores the default item colour
Private ItemColour As ItemColourType
 'Stores the custom highlight Colour
Private HighLightColour As ItemColourType
 'Indicates whether a custom highlight is to be used
Private UseHighLight As Boolean
 'Indicates whether a custom item colour is to be used
Private UseCustomColour As Boolean 'черезстрочки
Private NoHLFrame As Boolean 'нет рамки выделения
 'Stores whether the current item should be highlighted
'Private IsItemHighlighted As Boolean

'lved
'Public Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type LVHITTESTINFO
    pt As POINTAPI
    Flags As Long
    iItem As Long
    iSubitem As Long
End Type
Public Declare Function GetMessagePos Lib "user32" () As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Const LVIR_LABEL = 2
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Function LOWORD(dwValue As Long) As Integer
  MoveMemory LOWORD, dwValue, 2
End Function
Public Function HIWORD(dwValue As Long) As Integer
  MoveMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
End Function

Public Sub Attach(ByVal frmhWnd As Long) ', ByRef NewLV As ListView)
  '  Set CDLV = NewLV
    g_addProcOld = SetWindowLong(frmhWnd, GWL_WNDPROC, AddressOf LV_WindowProc)
End Sub

Public Sub UnAttach(ByVal frmhWnd As Long)
    Call SetWindowLong(frmhWnd, GWL_WNDPROC, g_addProcOld)
End Sub

Public Sub UseAlternatingColour(ByVal Value As Boolean)
    UseCustomColour = Value
    'CDLV.Refresh
    If FrmMain.ListView.Visible Then FrmMain.ListView.Refresh
    If FrmMain.tvGroup.Visible Then FrmMain.tvGroup.Refresh
End Sub
Public Sub NoHighLightFrame(ByVal Value As Boolean)
    NoHLFrame = Value
    'CDLV.Refresh
    If FrmMain.ListView.Visible Then FrmMain.ListView.Refresh
    If FrmMain.tvGroup.Visible Then FrmMain.tvGroup.Refresh
End Sub

Public Sub UseCustomHighLight(ByVal Value As Boolean)
    UseHighLight = Value
    'CDLV.Refresh
    If FrmMain.ListView.Visible Then FrmMain.ListView.Refresh
    If FrmMain.tvGroup.Visible Then FrmMain.tvGroup.Refresh
End Sub

Public Sub SetCustomColour(ByRef Value As ItemColourType)
    ItemColour = Value
End Sub

Public Function GetCustomColour() As ItemColourType
    GetCustomColour = ItemColour
End Function

Public Sub SetHighLightColour(ByRef Value As ItemColourType)
    HighLightColour = Value
End Sub

Public Function GetHighLightColour() As ItemColourType
    GetHighLightColour = HighLightColour
End Function

 '---{The subs and functions that custom paint the listview.
 '---{Drawing Subs
Private Sub DrawCustomColour(ByRef Struct As NMLVCUSTOMDRAW)
    With Struct
        .clrText = ItemColour.ForeGround
        .clrTextBk = ItemColour.BackGround
    End With
End Sub

'Private Sub DrawCustomHighlight(ByRef Struct As NMLVCUSTOMDRAW, ByVal row As Integer, ByVal chwnd As Long)
'    IsItemHighlighted = True
'    With Struct
'        .clrText = HighLightColour.ForeGround
'        .clrTextBk = HighLightColour.BackGround
'    End With
'    EnableHighlighting row, False, chwnd
'End Sub

'---{Subs that determine messages sent and what to do with them
'Public Sub EnableHighlighting(ByVal row As Integer, ByVal bHighLight As Boolean, ByVal chwnd As Long)
''    CDLV.Refresh = False
''    CDLV.ListItems.Item(row + 1).Selected = bHighLight
'
''Select Case chwnd
''Case FrmMain.ListView.hWnd
'''LockWindowUpdate chwnd
'''If FrmMain.ActiveControl.name = "ListView" Then
''    FrmMain.ListView.ListItems.Item(row + 1).Selected = bHighLight
'''End If
''
''Case FrmMain.tvGroup.hWnd
''   FrmMain.tvGroup.ListItems.Item(row + 1).Selected = bHighLight
''End Select
'
''LockWindowUpdate 0
'End Sub

'Private Function IsRowSelected(ByVal row As Integer, ByVal chwnd As Long)
'On Error Resume Next
'Select Case chwnd
'Case FrmMain.ListView.hWnd
'    IsRowSelected = FrmMain.ListView.ListItems.Item(row + 1).Selected
'Case FrmMain.tvGroup.hWnd
'    IsRowSelected = FrmMain.tvGroup.ListItems.Item(row + 1).Selected
'End Select
'End Function

'Where the magic happens :)
Private Function LV_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, _
                               ByVal wParam As Long, ByVal lParam As Long) As Long

Dim RetVal As Long
RetVal = 0    'initialise to a zero

If FrmMain.FrameView.Visible Then

    'Determine which message was recieved
    'Debug.Print iMsg
    
    Select Case iMsg
    Case WM_NOTIFY
        'If it's a WM_NOTIFY message copy the data from the address pointed to
        'by lParam into a NMHDR struct
        Dim udtNMHDR As NMHDR
        CopyMemory udtNMHDR, ByVal lParam, 12&
        
        With udtNMHDR
            If .code = NM_CUSTOMDRAW Then
                'If the code member of the struct is NM_CUSTOMDRAW, copy the data
                'pointed to by lParam into a NMLVCUSTOMDRAW struct
                Dim udtNMLVCUSTOMDRAW As NMLVCUSTOMDRAW

                'This is now OUR copy of the struct
                CopyMemory udtNMLVCUSTOMDRAW, ByVal lParam, Len(udtNMLVCUSTOMDRAW)

                With udtNMLVCUSTOMDRAW.nmcd
                    'determine whether or not this is one of the messages we are
                    'interested in
                    Select Case .dwDrawStage
                    

                    Case CDDS_PREPAINT
                        'If its a pre paint message then tell the control
                        '(basically windows) that we want first say in item
                        'painting, then exit and prevent VB getting this msg.
                        LV_WindowProc = CDRF_NOTIFYITEMDRAW    'Or CDRF_DODEFAULT Or CDRF_NEWFONT
                        
                        'Debug.Print Time & "pp"
                        Exit Function

                    Case CDDS_ITEMPREPAINT
                        'IsItemHighlighted = False

                        'Нет фокусной рамки
                        If NoHLFrame Then
                            .uItemState = .uItemState And (Not CDIS_FOCUS)
                        End If
                        '.uItemState = .uItemState And (Not CDIS_SELECTED) 'нет ли проблем с мультиселектом? + replacing color filling with border

                        'Alternating colours
                        If UseCustomColour Then
                            If (.dwItemSpec Mod 2) Then
                                '   If .dwItemSpec = 3 Then
                                '   DrawCustomColour udtNMLVCUSTOMDRAW, .dwItemSpec, udtNMHDR.hWndFrom
                                DrawCustomColour udtNMLVCUSTOMDRAW
                                '   End If
                            End If
                        End If

'If udtNMHDR.hWndFrom = FrmMain.ListView.hwnd Then
''выделить цветом после полосатости
''читать при загрузке базы данные цвете строки, поместить в таблицу (ключ цвет) и здесь юзать
''но как тут проверить ключ? умеем только порядковый номер
'If .dwItemSpec = 2 Then
'        udtNMLVCUSTOMDRAW.clrText = RGB(200, 200, 100)
'        udtNMLVCUSTOMDRAW.clrTextBk = RGB(100, 100, 200)
'    End If
'End If

                        '                        'Change Highlight
                        '                        If UseHighLight Then
                        '            'Debug.Print udtNMHDR.hWndFrom
                        '
                        '                        If IsRowSelected(.dwItemSpec, udtNMHDR.hWndFrom) Then
                        '                            DrawCustomHighlight udtNMLVCUSTOMDRAW, .dwItemSpec, udtNMHDR.hWndFrom
                        '
                        '                        End If
                        '                        End If

                        'Copy OUR copy of the struct back to the memory
                        'address pointed to by lParam
                        CopyMemory ByVal lParam, udtNMLVCUSTOMDRAW, Len(udtNMLVCUSTOMDRAW)
                        'Tell the control we want to be told about any changes, don't
                        'allow VB to get this message

                        LV_WindowProc = CDRF_DODEFAULT    'Or CDRF_NEWFONT Or CDRF_NOTIFYPOSTPAINT
                        'Exit Function ' не работает цвет должника если не отдаться g_addProcOld

                        '                    Case CDDS_ITEMPOSTPAINT
                        '                        If UseHighLight And IsItemHighlighted Then
                        '                            'If the item was selected re-select it, since we already
                        '                            'painted the highlight our custom colour
                        '                            EnableHighlighting .dwItemSpec, True, udtNMHDR.hWndFrom
                        '                            'LV_WindowProc = CDRF_DODEFAULT
                        '                            LV_WindowProc = CDRF_DODEFAULT 'Or CDRF_NOTIFYPOSTPAINT
                        '                            Exit Function
                        '                        End If
                        '
                        '                    Case Else
                        '                            LV_WindowProc = CDRF_DODEFAULT 'Or CDRF_NEWFONT Or CDRF_NOTIFYPOSTPAINT
                        '                            Exit Function
                    End Select
                End With
            End If
        End With

        '    Case WM_CAPTURECHANGED
        'Debug.Print " WM_CAP                           "
        '                            LV_WindowProc = CDRF_DODEFAULT
        '                            Exit Function
        
    Case WM_PARENTNOTIFY
     'если кликнуть на lv (для убирания окна редактирования lv)
        FrmMain.LVScroll
    'Debug.Print Time & " ss"
    End Select

End If

'pass all messages on to VB and then return the value to windows
'И передаем сообщение дальше
LV_WindowProc = CallWindowProc(g_addProcOld, hwnd, iMsg, wParam, lParam)
End Function

