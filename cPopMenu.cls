VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIconMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

' =======================================================================
' MENU private declares:
' =======================================================================

' Menu flag constants:
'Private Const MF_APPEND = &H100&
'Private Const MF_BITMAP = &H4&
'Private Const MF_BYCOMMAND = &H0&
'Private Const MF_BYPOSITION = &H400&
'Private Const MF_CALLBACKS = &H8000000
'Private Const MF_CHANGE = &H80&
Private Const MF_CHECKED = &H8&
'Private Const MF_CONV = &H40000000
'Private Const MF_DELETE = &H200&
'Private Const MF_DISABLED = &H2&
'Private Const MF_ENABLED = &H0&
'Private Const MF_END = &H80
'Private Const MF_ERRORS = &H10000000
'Private Const MF_GRAYED = &H1&
'Private Const MF_HELP = &H4000&
Private Const MF_HILITE = &H80&
'Private Const MF_HSZ_INFO = &H1000000
'Private Const MF_INSERT = &H0&
'Private Const MF_LINKS = &H20000000
'Private Const MF_MASK = &HFF000000
'Private Const MF_MENUBARBREAK = &H20&
'Private Const MF_MENUBREAK = &H40&
'Private Const MF_MOUSESELECT = &H8000&
Private Const MF_OWNERDRAW = &H100&
'Private Const MF_POPUP = &H10&
'Private Const MF_POSTMSGS = &H4000000
'Private Const MF_REMOVE = &H1000&
'Private Const MF_SENDMSGS = &H2000000
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MF_SYSMENU = &H2000&
'Private Const MF_UNCHECKED = &H0&
'Private Const MF_UNHILITE = &H0&
'Private Const MF_USECHECKBITMAPS = &H200&
'Private Const MF_DEFAULT = &H1000&

Private Const MFT_STRING = MF_STRING
'Private Const MFT_BITMAP = MF_BITMAP
'Private Const MFT_MENUBARBREAK = MF_MENUBARBREAK
'Private Const MFT_MENUBREAK = MF_MENUBREAK
Private Const MFT_OWNERDRAW = MF_OWNERDRAW
Private Const MFT_RADIOCHECK = &H200&
Private Const MFT_SEPARATOR = MF_SEPARATOR
'Private Const MFT_RIGHTORDER = &H2000&

' New versions of the names...
Private Const MFS_GRAYED = &H3&
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_HILITE = MF_HILITE
'Private Const MFS_ENABLED = MF_ENABLED
'Private Const MFS_UNCHECKED = MF_UNCHECKED
'Private Const MFS_UNHILITE = MF_UNHILITE
'Private Const MFS_DEFAULT = MF_DEFAULT

' MenuItemInfo Mask constants
Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2&
'Private Const MIIM_SUBMENU = &H4&
'Private Const MIIM_CHECKMARKS = &H8&
Private Const MIIM_TYPE = &H10&
Private Const MIIM_DATA = &H20&

'Private Const SC_RESTORE = &HF120&
'Private Const SC_MOVE = &HF010&
'Private Const SC_SIZE = &HF000&
'Private Const SC_MAXIMIZE = &HF030&
'Private Const SC_MINIMIZE = &HF020&
'Private Const SC_CLOSE = &HF060&
     
'Private Const SC_ARRANGE = &HF110&
'Private Const SC_HOTKEY = &HF150&
'Private Const SC_HSCROLL = &HF080&
'Private Const SC_KEYMENU = &HF100&
'Private Const SC_MOUSEMENU = &HF090&
'Private Const SC_NEXTWINDOW = &HF040&
'Private Const SC_PREVWINDOW = &HF050&
'Private Const SC_SCREENSAVE = &HF140&
'Private Const SC_TASKLIST = &HF130&
'Private Const SC_VSCROLL = &HF070&
'Private Const SC_ZOOM = SC_MAXIMIZE
'Private Const SC_ICON = SC_MINIMIZE

' Owner draw information:
'Private Const ODS_CHECKED = &H8
'Private Const ODS_DISABLED = &H4
'Private Const ODS_FOCUS = &H10
'Private Const ODS_GRAYED = &H2
'Private Const ODS_SELECTED = &H1
'Private Const ODT_BUTTON = 4
'Private Const ODT_COMBOBOX = 3
'Private Const ODT_LISTBOX = 2
Private Const ODT_MENU = 1

' GetDeviceCaps
Private Const BITSPIXEL = 12

Private Type MEASUREITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemWidth As Long
   ItemHeight As Long
   ItemData As Long
End Type

Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   ItemData As Long
End Type

Private Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As Long
   cch As Long
End Type
Private Type MENUITEMINFO_STRINGDATA
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type

'Private Type MENUITEMTEMPLATE
'   mtOption As Integer
'   mtID As Integer
'   mtString As Byte
'End Type
'Private Type MENUITEMTEMPLATEHEADER
'   versionNumber As Integer
'   Offset As Integer
'End Type

'Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long

'Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
'Private Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
'Private Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfoStr Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO_STRINGDATA) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfoStr Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO_STRINGDATA) As Long
'Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
'Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

'Private Declare Function CreateMenu Lib "user32" () As Long
'Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
'Private Declare Function CreatePopupMenu Lib "user32" () As Long

'Private Declare Function AppendMenuBylong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
'Private Declare Function AppendMenuByString Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
'Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
'Private Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
'Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Private Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
'Private Declare Function InsertMenuByString Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
'Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

'Private Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
'Private Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
'Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
'Private Declare Function HiliteMenuItem Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long

'Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
'Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

'Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long


' =======================================================================
' GDI private declares:
' =======================================================================

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawEdgeAPI Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' Pen functions:
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'    Private Const PS_DASH = 1
'    Private Const PS_DASHDOT = 3
'    Private Const PS_DASHDOTDOT = 4
'    Private Const PS_DOT = 2
    Private Const PS_SOLID = 0
'    Private Const PS_NULL = 5
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lParam As Long, ByVal wParam As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal cx As Long, ByVal cy As Long, _
   ByVal fuFlags As Long) As Long

'/* Image type */
'Private Const DST_COMPLEX = &H0
'Private Const DST_TEXT = &H1
'Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
'Private Const DST_BITMAP = &H4

' /* State type */
'Private Const DSS_NORMAL = &H0
'Private Const DSS_UNION = &H10         ' /* Gray string appearance */
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
'Private Const DSS_RIGHT = &H8000

'Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
'Private Const DT_WORDBREAK = &H10
Private Const DT_VCENTER = &H4
'Private Const DT_TOP = &H0
'Private Const DT_TABSTOP = &H80
Private Const DT_SINGLELINE = &H20
'Private Const DT_RIGHT = &H2
'Private Const DT_NOCLIP = &H100
'Private Const DT_INTERNAL = &H1000
'Private Const DT_EXTERNALLEADING = &H200
'Private Const DT_EXPANDTABS = &H40
'Private Const DT_CHARSTREAM = 4
'Private Const DT_NOPREFIX = &H800
'Private Const DT_EDITCONTROL = &H2000&
'Private Const DT_PATH_ELLIPSIS = &H4000&
'Private Const DT_END_ELLIPSIS = &H8000&
'Private Const DT_MODIFYSTRING = &H10000
'Private Const DT_RTLREADING = &H20000
'Private Const DT_WORD_ELLIPSIS = &H40000

Private Const OPAQUE = 2
Private Const Transparent = 1

' DrawEdge:
'Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
'Private Const BDR_SUNKENINNER = &H8

'Private Const BDR_OUTER = &H3
'Private Const BDR_INNER = &HC
'Private Const BDR_RAISED = &H5
'Private Const BDR_SUNKEN = &HA

'Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
'Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
'Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
'Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

'Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
'Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
'Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
'Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const CLR_INVALID = -1

'Private Declare Function ImageList_GetIconSize Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        cX As Long, cY As Long _
    ) As Long
Private Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
    ) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, ByVal i As Long, _
        ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal fStyle As Long _
    ) As Long
' Create a new icon based on an image list icon:
Private Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
    ) As Long
'Private Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
    ) As Long
Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long
'Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
'Private Const ILD_FOCUS = 4
'Private Const ILD_MASK = &H10&
'Private Const ILD_IMAGE = &H20&
'Private Const ILD_ROP = &H40&
'Private Const ILD_OVERLAYMASK = 3840
'Private Const ILC_COLOR = &H0
Private Const ILC_COLOR32 = &H20
Private Const ILC_MASK = &H1&

' =======================================================================
' General Win private declares:
' =======================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

'Private Const HWND_DESKTOP = 0


' Window Messages
Private Const WM_DESTROY = &H2
'Private Const WM_SETTEXT = &HC
Private Const WM_WININICHANGE = &H1A
'Private Const WM_ACTIVATEAPP = &H1C
'Private Const WM_CANCELMODE = &H1F
'Private Const WM_SETCURSOR = &H20
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
'Private Const WM_STYLECHANGING = &H7C
'Private Const WM_STYLECHANGED = &H7D
'Private Const WM_NCCALCSIZE = &H83
'Private Const WM_NCHITTEST = &H84
'Private Const WM_NCPAINT = &H85
'Private Const WM_NCACTIVATE = &H86
'Private Const WM_NCLBUTTONDOWN = &HA1
'Private Const WM_NCLBUTTONUP = &HA2
'Private Const WM_NCLBUTTONDBLCLK = &HA3
'Private Const WM_KEYDOWN = &H100
'Private Const WM_COMMAND = &H111
'Private Const WM_SYSCOMMAND = &H112
Private Const WM_INITMENUPOPUP = &H117
'Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
'Private Const WM_MDIGETACTIVE = &H229
'Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212


' =======================================================================
' IMPLEMENTATION
' =======================================================================

Public Enum ECPHighlightStyleConstants
   ECPHighlightStyleStandard = 0
   ECPHighlightStyleGradient = 1
   ECPHighlightStyleButton = 2
End Enum

Private m_cMemDC As cMemDC
Private m_cBitmap As cMemDC
Private m_cNCM As cNCMetrics
Private m_cBrush As cDottedBrush

Private m_hWnd As Long
Private m_scmode As Long

Private m_oActiveMenuColor As OLE_COLOR
Private m_oInActiveMenuColor As OLE_COLOR
Private m_oMenuBackgroundColor As OLE_COLOR

Private m_fnt As StdFont
Private m_fntSymbol As StdFont
Private m_pic As IPicture

Private m_eHighlightStyle As ECPHighlightStyleConstants

Private m_iRestore As Long
Private m_hMenuRestore() As Long
Private m_iMenuPosition() As Long
Private m_tMIIS() As MENUITEMINFO_STRINGDATA
Private m_sCaption() As String 'текст, состо€щий  из всех пунктов меню
Private m_sShortCut() As String
Private m_sAccelerator() As String
Private m_lMenuTextSize() As Long 'получающа€с€ обща€ ширина дл€ конкретного меню
Private m_lMenuShortCutSize() As Long

Private m_iHaveSeenCount As Long
Private m_hMenuSeen() As Long

Private m_lMenuItemHeight As Long

Private m_colIcons As New Collection
Private m_hIml As Long
Private m_ptrVb6ImageList As Long
Private m_lIconSize As Long
Private m_OfficeXPStyle As Boolean

Implements ISubclass

Public Property Get OfficeXpStyle() As Boolean
Attribute OfficeXpStyle.VB_Description = "Gets/sets whether the menu draws using the Office XP style."
   OfficeXpStyle = m_OfficeXPStyle
End Property

Public Property Let OfficeXpStyle(ByVal bState As Boolean)
   m_OfficeXPStyle = bState
End Property

Public Property Let ImageList( _
        ByRef vImageList As Variant _
    )
Attribute ImageList.VB_Description = "Gets/sets the ImageList to use when drawing the icons."
    m_hIml = 0
    m_ptrVb6ImageList = 0
    If (VarType(vImageList) = vbLong) Then
        ' Assume a handle to an image list:
        m_hIml = vImageList
    ElseIf (VarType(vImageList) = vbObject) Then
        ' Assume a VB image list:
        On Error Resume Next
        ' Get the image list initialised..
        vImageList.ListImages(1).Draw 0, 0, 0, 1
        m_hIml = vImageList.hImageList
        If (err.Number = 0) Then
            ' Check for VB6 image list:
            If (TypeName(vImageList) = "ImageList") Then
               Dim o As Object
               Set o = vImageList
               m_ptrVb6ImageList = ObjPtr(o)
            End If
        Else
            Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
        End If
        On Error GoTo 0
    End If
    If (m_hIml <> 0) Then
        If (m_ptrVb6ImageList <> 0) Then
            m_lIconSize = vImageList.ImageHeight
        Else
            Dim rc As RECT
            ImageList_GetImageRect m_hIml, 0, rc
            m_lIconSize = rc.bottom - rc.top
        End If
    End If
End Property

Public Property Get IconIndex(ByVal sCaption As String) As Long
Attribute IconIndex.VB_Description = "Gets/sets the 0-based index of the icon to be displayed against the menu item with the specified text."
'возвратит индекс по надписи
   IconIndex = m_colIcons.Item(sCaption)
End Property
Public Property Get IconIndexForIndex(ByVal lItem As Long) As Long
   IconIndexForIndex = m_colIcons.Item(lItem)
End Property
'Public Property Let IconIndex(ByVal sCaption As String, ByVal nIndex As Long)
Public Property Let IconIndex(ByVal sCaption As String, ByVal nIndex As Long)
   If nIndex = -1 Then
      m_colIcons.Remove sCaption
   Else
      On Error Resume Next
      m_colIcons.Item sCaption
      If err.Number = 0 Then
         m_colIcons.Remove sCaption
      End If
      On Error GoTo 0
      m_colIcons.Add nIndex, sCaption
   End If
End Property
Public Sub IconItemCaptionChanged(ByVal sOldCaption As String, ByVal sNewCaption As String)
Attribute IconItemCaptionChanged.VB_Description = "If the caption of a menu item is changed in code, call this method to ensure any icon associated with the item still displays."
Dim lIdx As Long
   lIdx = -1
   On Error Resume Next
   lIdx = m_colIcons.Item(sOldCaption)
   If lIdx > -1 Then
      m_colIcons.Remove sOldCaption
      m_colIcons.Add lIdx, sNewCaption
   End If
End Sub
Public Property Get IconItemCount() As Long
Attribute IconItemCount.VB_Description = "Gets the total number of icons associated with menu items in the object."
   IconItemCount = m_colIcons.Count
End Property
Public Sub ClearIcons()
Attribute ClearIcons.VB_Description = "Clears all the icons associated with menu items in the control."
   Set m_colIcons = New Collection
End Sub
Public Property Get HighlightStyle() As ECPHighlightStyleConstants
Attribute HighlightStyle.VB_Description = "Gets/sets the highlight style of the menu items."
   HighlightStyle = m_eHighlightStyle
End Property
Public Property Let HighlightStyle(ByVal eStyle As ECPHighlightStyleConstants)
   m_eHighlightStyle = eStyle
End Property
Private Property Get hFontSymbol() As Long
Dim iFn As IFont
   Set iFn = m_fntSymbol
   hFontSymbol = iFn.hFont
End Property
Friend Property Let Font( _
      fntTHis As StdFont _
   )
   pSetFont fntTHis
End Property
Public Property Set Font( _
      fntTHis As StdFont _
   )
Attribute Font.VB_Description = "Gets/sets the font used to draw the menu text. Set to nothing to use the default menu font."
   pSetFont fntTHis
End Property
Public Property Get Font() As StdFont
Dim lHDC As Long
   If m_fnt Is Nothing Then
      lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      Set Font = m_cNCM.Font(lHDC, MenuFOnt)
      DeleteDC lHDC
   Else
      Set Font = m_fnt
   End If
End Property
Private Sub pSetFont(fntTHis As StdFont)
   Set m_fnt = fntTHis
   m_fntSymbol.Size = Font.Size * 1.2
End Sub
Public Property Get BackgroundPicture() As IPicture
Attribute BackgroundPicture.VB_Description = "Gets/sets a background bitmap to show in the menu."
   Set BackgroundPicture = m_pic
End Property
Public Property Let BackgroundPicture(ByRef ipic As IPicture)
   pSetPicture ipic
End Property
Public Property Set BackgroundPicture(ByRef ipic As IPicture)
   pSetPicture ipic
End Property
Private Sub pSetPicture(ByRef ipic As IPicture)
   If Not ipic Is Nothing Then
      Set m_pic = ipic
      Set m_cBitmap = New cMemDC
      m_cBitmap.CreateFromPicture ipic
   Else
      Set m_cBitmap = Nothing
   End If
End Sub
Public Property Let ActiveMenuForeColor(ByVal oColor As OLE_COLOR)
Attribute ActiveMenuForeColor.VB_Description = "Gets/sets the foreground colour of an active menu item. Set to -1 to use the default colours."
   m_oActiveMenuColor = oColor
End Property
Public Property Get ActiveMenuForeColor() As OLE_COLOR
   If m_oActiveMenuColor = CLR_INVALID Then
      If (m_OfficeXPStyle) Then
         ActiveMenuForeColor = vbMenuText
      Else
         ActiveMenuForeColor = vbHighlightText
      End If
   Else
      ActiveMenuForeColor = m_oActiveMenuColor
   End If
End Property
Public Property Let InActiveMenuForeColor(ByVal oColor As OLE_COLOR)
Attribute InActiveMenuForeColor.VB_Description = "Gets/sets the colour to draw inactive (non-selected) menu item captions.  Set to -1 to use the default colours."
   m_oInActiveMenuColor = oColor
End Property
Public Property Get InActiveMenuForeColor() As OLE_COLOR
   If m_oInActiveMenuColor = CLR_INVALID Then
      InActiveMenuForeColor = vbMenuText
   Else
      InActiveMenuForeColor = m_oInActiveMenuColor
   End If
End Property
Public Property Let MenuBackgroundColor(ByVal oColor As OLE_COLOR)
Attribute MenuBackgroundColor.VB_Description = "Gets/sets the background colour of the menu items.  Set to -1 to use the default colours."
   m_oMenuBackgroundColor = oColor
End Property
Public Property Get MenuBackgroundColor() As OLE_COLOR
   If m_oMenuBackgroundColor = CLR_INVALID Then
      MenuBackgroundColor = vbMenuBar
   Else
      MenuBackgroundColor = m_oMenuBackgroundColor
   End If
End Property
Private Property Get hFont() As Long
Dim iFn As IFont
   Set iFn = Font
   hFont = iFn.hFont
End Property
Public Sub Attach(ByVal lhWnd As Long, ByVal ScMode As Long)
Attribute Attach.VB_Description = "Connects this object to a form to and starts customising the menus."
   Detach
   m_hWnd = lhWnd
   m_scmode = ScMode
   AttachMessage Me, m_hWnd, WM_WININICHANGE
   AttachMessage Me, m_hWnd, WM_DRAWITEM
   AttachMessage Me, m_hWnd, WM_MEASUREITEM
   AttachMessage Me, m_hWnd, WM_MENUCHAR
   AttachMessage Me, m_hWnd, WM_INITMENUPOPUP
   AttachMessage Me, m_hWnd, WM_EXITMENULOOP
   AttachMessage Me, m_hWnd, WM_DESTROY
End Sub
Public Sub Detach()
Attribute Detach.VB_Description = "Detaches this object from a form it was previously attached to and clears up all menu customisation."
   If Not m_hWnd = 0 Then
      DetachMessage Me, m_hWnd, WM_WININICHANGE
      DetachMessage Me, m_hWnd, WM_DRAWITEM
      DetachMessage Me, m_hWnd, WM_MEASUREITEM
      DetachMessage Me, m_hWnd, WM_MENUCHAR
      DetachMessage Me, m_hWnd, WM_INITMENUPOPUP
      DetachMessage Me, m_hWnd, WM_EXITMENULOOP
      DetachMessage Me, m_hWnd, WM_DESTROY
   End If
End Sub

Private Sub OwnerDrawMenu(ByVal hMenu As Long)
Dim lC As Long
Dim tMIIS As MENUITEMINFO_STRINGDATA
Dim tMII As MENUITEMINFO
Dim iMenu As Long
Dim sCap As String
Dim sShortCut As String
Dim tR As RECT
'Dim iPos As Long
Dim lID As Long
Dim bHaveSeen As Boolean
Dim hFntOld As Long
Dim lMenuTextSize As Long
Dim lMenuShortCutSize As Long
Dim i As Long
                  
   ' Set OD flag on the fly...
   bHaveSeen = pbHaveSeen(hMenu)

   hFntOld = SelectObject(m_cMemDC.hdc, hFont)
   lC = GetMenuItemCount(hMenu)
   For iMenu = 0 To lC - 1
      
      If Not bHaveSeen Then

         tMIIS.fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
         tMIIS.cch = 127
         tMIIS.dwTypeData = String$(128, 0)
         tMIIS.cbSize = LenB(tMIIS)
         GetMenuItemInfoStr hMenu, iMenu, True, tMIIS
         'Debug.Print "New Item", tMIIS.dwTypeData

         lID = plAddToRestoreList(hMenu, iMenu, tMIIS)

         If Not (tMIIS.fType And MFT_OWNERDRAW) = MFT_OWNERDRAW Then
            ' Setting this flag causes tMIIS.dwTypeData to be
            ' overwritten with our own app-defined value:
            tMII.fType = tMIIS.fType Or MFT_OWNERDRAW And Not MFT_STRING
            tMII.dwItemData = lID
            tMII.cbSize = LenB(tMII)
            tMII.fMask = MIIM_TYPE Or MIIM_DATA
            SetMenuItemInfo hMenu, iMenu, True, tMII
         End If

      Else

         tMII.fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
         tMII.cbSize = Len(tMIIS)
         ReDim b(0 To 128) As Byte
         tMII.dwTypeData = VarPtr(b(0))
         GetMenuItemInfo hMenu, iMenu, True, tMII
         lID = tMII.dwItemData

         If Not ((tMII.fType And MFT_OWNERDRAW) = MFT_OWNERDRAW) Then

            lID = plReplaceIndex(hMenu, iMenu)

            'Debug.Print "VB has done something to it!", lID
            tMIIS.fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
            tMIIS.cch = 127
            tMIIS.dwTypeData = String$(128, 0)
            tMIIS.cbSize = LenB(tMIIS)
            GetMenuItemInfoStr hMenu, iMenu, True, tMIIS
'
            If lID = 0 Then
               ' New item,just added:
               lID = plAddToRestoreList(hMenu, iMenu, tMIIS)
            Else
               ' replacing existing:
               pReplaceRestoreList lID, hMenu, iMenu, tMIIS
            End If

            ' Setting this flag causes tMIIS.dwTypeData to be
            ' overwritten with our own app-defined value:
            tMII.fType = tMIIS.fType Or MFT_OWNERDRAW And Not MFT_STRING
            tMII.dwItemData = lID
            tMII.cbSize = LenB(tMII)
            tMII.fMask = MIIM_TYPE Or MIIM_DATA
            SetMenuItemInfo hMenu, iMenu, True, tMII

         End If

      End If

      If lID > 0 And lID <= m_iRestore Then
         sCap = m_sCaption(lID)
         sShortCut = m_sShortCut(lID)

'Debug.Print m_sCaption(lID), m_sShortCut(lID)

         DrawText m_cMemDC.hdc, sCap, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
         If tR.right - tR.left + 1 > lMenuTextSize Then
            lMenuTextSize = tR.right - tR.left + 1
         End If
         If Len(sShortCut) > 0 Then
            DrawText m_cMemDC.hdc, sShortCut, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
            If tR.right - tR.left + 1 > lMenuShortCutSize Then
               lMenuShortCutSize = tR.right - tR.left + 1
            End If
         End If
         m_lMenuItemHeight = tR.bottom - tR.top + 2
         If m_lMenuItemHeight < m_lIconSize - 1 Then
            m_lMenuItemHeight = m_lIconSize - 1
         End If

      Else
         'Debug.Print "ERROR! ERROR! ERROR!"
      End If
      
   Next iMenu
   
   For i = 1 To m_iRestore
      If m_hMenuRestore(i) = hMenu Then
         m_lMenuTextSize(i) = lMenuTextSize
         m_lMenuShortCutSize(i) = lMenuShortCutSize
      End If
   Next i
   
   SelectObject m_cMemDC.hdc, hFntOld
   
End Sub
Private Function pbHaveSeen(ByVal hMenu As Long) As Boolean
   
   ' When WM_INITMENUPOPUP fires, this may or not be
   ' a new menu.  We use an array to store which menus
   ' we've already worked on:

Dim i As Long
   
   For i = 1 To m_iHaveSeenCount
      If hMenu = m_hMenuSeen(i) Then
         pbHaveSeen = True
         Exit Function
      End If
   Next i
   m_iHaveSeenCount = m_iHaveSeenCount + 1
   ReDim Preserve m_hMenuSeen(1 To m_iHaveSeenCount) As Long
   m_hMenuSeen(m_iHaveSeenCount) = hMenu

End Function
Private Function plReplaceIndex(ByVal hMenu As Long, ByVal iMenu As Long)
Dim i As Long
   For i = 1 To m_iRestore
      If m_hMenuRestore(i) = hMenu Then
         If m_iMenuPosition(i) = iMenu Then
            plReplaceIndex = i
            Exit Function
         End If
      End If
   Next i
End Function
Private Function plAddToRestoreList(ByVal hMenu As Long, ByVal iMenu As Long, tMIIS As MENUITEMINFO_STRINGDATA) As Long
   
   ' Here we store information about a menu item.  When the
   ' menus are closed again we can reset things back to the
   ' way they were using this struct.

   m_iRestore = m_iRestore + 1
   ReDim Preserve m_hMenuRestore(1 To m_iRestore) As Long
   ReDim Preserve m_iMenuPosition(1 To m_iRestore) As Long
   ReDim Preserve m_tMIIS(1 To m_iRestore) As MENUITEMINFO_STRINGDATA
   ReDim Preserve m_sCaption(1 To m_iRestore) As String
   ReDim Preserve m_sShortCut(1 To m_iRestore) As String
   ReDim Preserve m_sAccelerator(1 To m_iRestore) As String
   ReDim Preserve m_lMenuTextSize(1 To m_iRestore) As Long
   ReDim Preserve m_lMenuShortCutSize(1 To m_iRestore) As Long
   pReplaceRestoreList m_iRestore, hMenu, iMenu, tMIIS
   plAddToRestoreList = m_iRestore

End Function
Private Sub pReplaceRestoreList(ByVal lIdx As Long, hMenu As Long, iMenu As Long, tMIIS As MENUITEMINFO_STRINGDATA)
Dim sCap As String
'Dim sShortCut As String
Dim iPos As Long

   m_hMenuRestore(lIdx) = hMenu
   m_iMenuPosition(lIdx) = iMenu
   LSet m_tMIIS(lIdx) = tMIIS
   If tMIIS.cch > 0 Then
      sCap = left$(tMIIS.dwTypeData, tMIIS.cch)
   Else
      sCap = ""
   End If
   
'режут tab (на название и гор€ч клавиши)
   iPos = InStr(sCap, vbTab)
   If iPos > 0 Then
      m_sShortCut(lIdx) = Mid$(sCap, iPos + 1)
      m_sCaption(lIdx) = left$(sCap, iPos - 1)
   Else
      m_sCaption(lIdx) = sCap
      m_sShortCut(lIdx) = ""
   End If
   iPos = InStr(m_sCaption(lIdx), "&")
   If iPos > 0 And iPos < Len(m_sCaption(lIdx)) Then
      m_sAccelerator(lIdx) = UCase$(Mid$(m_sCaption(lIdx), iPos + 1, 1))
   End If
End Sub
Private Function InternalIDForWindowsID(ByVal wID As Long) As Long
Dim i As Long
   ' linear search I'm afraid, but it is only called once
   ' per menu item shown (when WM_MEASUREITEM is fired)
   For i = 1 To m_iRestore
      If m_tMIIS(i).wID = wID Then
         InternalIDForWindowsID = i
         Exit Function
      End If
   Next i
End Function
Private Sub pRestoreList()
Dim i As Long
   'Debug.Print "RESTORELIST"
   ' erase the lot:
   For i = 1 To m_iRestore
      SetMenuItemInfoStr m_hMenuRestore(i), m_iMenuPosition(i), True, m_tMIIS(i)
   Next i
   m_iRestore = 0
   Erase m_hMenuRestore
   Erase m_iMenuPosition
   Erase m_tMIIS
   Erase m_sCaption()
   Erase m_sShortCut()
   Erase m_sAccelerator()
   m_iHaveSeenCount = 0
   Erase m_hMenuSeen()
End Sub

Private Sub Class_Initialize()
   Set m_cNCM = New cNCMetrics
   m_cNCM.GetMetrics
   Set m_cMemDC = New cMemDC
   m_cMemDC.Width = Screen.Width \ Screen.TwipsPerPixelY
   m_cMemDC.Height = 24
   m_oActiveMenuColor = CLR_INVALID
   m_oInActiveMenuColor = CLR_INVALID
   m_oMenuBackgroundColor = CLR_INVALID
   Set m_fntSymbol = New StdFont
   m_fntSymbol.name = "Marlett"
   m_fntSymbol.Size = Font.Size * 1.2
   Set m_cBrush = New cDottedBrush
   m_cBrush.Create
End Sub

Private Sub Class_Terminate()
   Set m_cMemDC = Nothing
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
   If _
      CurrentMessage = WM_DESTROY Or _
      CurrentMessage = WM_INITMENUPOPUP Or _
      CurrentMessage = WM_WININICHANGE Or _
      CurrentMessage = WM_EXITMENULOOP _
   Then
      ISubclass_MsgResponse = emrPreprocess
   Else
      ISubclass_MsgResponse = emrConsume
   End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Dim iMenu As Long
'Dim iLastDownOn As Long
'Dim iLastOver As Long
Dim lR As Long
Dim lFlag As Long
Dim hMenu As Long
Dim iChar As Long

   Select Case iMsg
   Case WM_INITMENUPOPUP
      If (lParam And &HFFFF0000) = 0 Then
         OwnerDrawMenu wParam
      End If
      
   Case WM_MEASUREITEM
      ISubclass_WindowProc = MeasureItem(wParam, lParam)
   
   Case WM_DRAWITEM
      DrawItem wParam, lParam
      
   Case WM_MENUCHAR
      ' Check that this is my menu:
      lFlag = wParam \ &H10000
      If ((lFlag And MF_SYSMENU) <> MF_SYSMENU) Then
         hMenu = lParam
         iChar = (wParam And &HFFFF&)
         ' See if this corresponds to an accelerator on the menu:
         lR = ParseMenuChar(hMenu, iChar)
         If lR > 0 Then
            ISubclass_WindowProc = lR
            Exit Function
         End If
      End If
      ISubclass_WindowProc = CallOldWindowProc(m_hWnd, WM_MENUCHAR, wParam, lParam)
   
   Case WM_WININICHANGE
      m_cNCM.GetMetrics
   
   Case WM_EXITMENULOOP
      pRestoreList
   
   Case WM_DESTROY
      Detach
   
   End Select
   
End Function
Private Function ParseMenuChar( _
        ByVal hMenu As Long, _
        ByVal iChar As Integer _
    ) As Long
Dim sChar As String
Dim l As Long
'Dim lH() As Long
'Dim sItems() As String
    
    sChar = UCase$(Chr$(iChar))
    For l = 1 To m_iRestore
        If (m_hMenuRestore(l) = hMenu) Then
            If (m_sAccelerator(l) = sChar) Then
               ParseMenuChar = &H20000 Or m_iMenuPosition(l)
               ' Debug.Print "Found Menu Char"
               Exit Function
            End If
        End If
    Next l
    
End Function

Private Function MeasureItem(ByVal wParam As Long, ByVal lParam As Long) As Long
'высота пунктов меню
Dim tMIS As MEASUREITEMSTRUCT
Dim lID As Long

CopyMemory tMIS, ByVal lParam, LenB(tMIS)
If tMIS.CtlType = ODT_MENU Then

    ' because we don't get the popup menu handle
    ' in the tMIS structure, we have to do an internal
    ' lookup to find info about this menu item.
    ' poor implementation of MEASUREITEMSTRUCT - it
    ' should have a .hWndItem field like DRAWITEMSTRUCT
    ' - spm
    lID = InternalIDForWindowsID(tMIS.itemID)
    'Debug.Print lID 'сколько пунктов меню видно (1,2,3,4...)

    ' Width:
    tMIS.itemWidth = 4 + 22 + m_lMenuTextSize(lID) + 4
    If m_lMenuShortCutSize(lID) > 0 Then
        tMIS.itemWidth = tMIS.itemWidth + 4 + m_lMenuShortCutSize(lID) + 4
    End If
    If (m_OfficeXPStyle) Then
        tMIS.itemWidth = tMIS.itemWidth + 4
    End If

    'm_tMIIS(lID).cch кол символов в этом меню
    'Debug.Print m_tMIIS(lID).wID
    'Debug.Print tMIS.itemID

    ' Height:
    If lID > 0 And lID <= m_iRestore Then
        If (m_tMIIS(lID).fType And MFT_SEPARATOR) = MFT_SEPARATOR Then
            If (m_OfficeXPStyle) Then
                tMIS.ItemHeight = 3
            Else
                tMIS.ItemHeight = 8
            End If
        Else
            ' menu item height is always the same
            Select Case tMIS.itemID
            Case 91 To 116    '109 86-111 'menuheight высота меню пунктов
                tMIS.ItemHeight = m_lMenuItemHeight + 2  'у групп нет иконок - поуже

            Case 117 To 142 '142 на вс€кий случай +1 141    '109 86-111 'дл€ хистори
                'Debug.Print m_tMIIS(lID).cch
                If m_tMIIS(lID).cch = 1 Then 'нет капчи (точнее 1 символ), скрыть мпункт
                    tMIS.ItemHeight = 0
                Else
                    tMIS.ItemHeight = m_lMenuItemHeight + 2  'поуже поменьше
                End If

            Case Else
                tMIS.ItemHeight = m_lMenuItemHeight + 6
            End Select
            '            If (m_OfficeXPStyle) Then
            '                'высота строки в XP стиле
            '                'tMIS.itemHeight = tMIS.itemHeight + 4
            '            End If
        End If
    Else
        ' problem.
    End If

    CopyMemory ByVal lParam, tMIS, LenB(tMIS)

Else
    MeasureItem = CallOldWindowProc(m_hWnd, WM_MEASUREITEM, wParam, lParam)
End If
End Function
Private Function DrawItem(ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tDIS As DRAWITEMSTRUCT
Dim hBr As Long
Dim tR As RECT, tTR As RECT, tWR As RECT
Dim lHDC As Long
Dim hFntOld As Long
Dim hFntsOld As Long
Dim tMII As MENUITEMINFO
Dim bRadioCheck As Boolean, bDisabled As Boolean, bChecked As Boolean, bHighlighted As Boolean
'Dim bIsTopLevel As Boolean
Dim lID As Long
Dim lSelLeft As Long
'Dim sCC As String
Dim lIconIndex As Long
Dim lX As Long, lY As Long
Dim hBrush As Long

   CopyMemory tDIS, ByVal lParam, LenB(tDIS)
   
   If tDIS.CtlType = ODT_MENU Then
      ' tDIS.hWndItem is the menu containing the item, tDIS.itemID is the wID
      
      m_cMemDC.Width = tDIS.rcItem.right - tDIS.rcItem.left + 1
      m_cMemDC.Height = tDIS.rcItem.bottom - tDIS.rcItem.top + 1
      lHDC = m_cMemDC.hdc
      hFntOld = SelectObject(lHDC, hFont)
      
      LSet tR = tDIS.rcItem
      OffsetRect tR, -tR.left, -tR.top
            
      ' Fill background:
      tTR.right = m_cMemDC.Width
      tTR.bottom = m_cMemDC.Height
      
      If m_cBitmap Is Nothing Then
         hBr = CreateSolidBrush(TranslateColor(MenuBackgroundColor))
         FillRect lHDC, tTR, hBr
         DeleteObject hBr
      Else
         TileArea lHDC, tR.left, tR.top, tR.right - tR.left + 1, tR.bottom - tR.top + 1, m_cBitmap.hdc, m_cBitmap.Width, m_cBitmap.Height, tDIS.rcItem.top
      End If
         
      If (m_OfficeXPStyle) Then
         Dim tSideRect As RECT
         LSet tSideRect = tTR
         tSideRect.right = m_lMenuItemHeight + 8
         fillWithLighterControlColour lHDC, tSideRect, tDIS.rcItem.top
      End If
      
      tR.top = tR.top + 1
            
      SetBkMode lHDC, Transparent
      
      ' Draw the text:
      tMII.cbSize = LenB(tMII)
      tMII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_DATA
      ReDim b(0 To 128) As Byte
      tMII.dwTypeData = VarPtr(b(0))
      GetMenuItemInfo tDIS.hwndItem, tDIS.itemID, False, tMII
      
      If (tMII.fType And MFT_SEPARATOR) = MFT_SEPARATOR Then
         ' Separator:
         LSet tWR = tR
         tWR.top = (tWR.bottom - tWR.top - 2) \ 2 + tWR.top
         tWR.bottom = tWR.top + 2
         InflateRect tWR, -12, 0
         If (m_OfficeXPStyle) Then
            Dim tWRS As RECT
            LSet tWRS = tWR
            tWRS.left = tSideRect.right + 4
            tWRS.right = tWRS.right + 20
            tWRS.top = tWRS.top + 1
            tWRS.bottom = tWRS.top
            DrawEdge lHDC, tWRS, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, True
         Else
            DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, False
         End If
      Else
         ' Text item:
         bRadioCheck = ((tMII.fType And MFT_RADIOCHECK) = MFT_RADIOCHECK)
         bDisabled = ((tMII.fState And MFS_DISABLED) = MFS_DISABLED)
         bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED)
         bHighlighted = ((tMII.fState And MFS_HILITE) = MFS_HILITE)
         
         lID = tMII.dwItemData

         ' Icon?
         lIconIndex = -1
         On Error Resume Next
         lIconIndex = m_colIcons.Item(m_sCaption(lID))
         On Error GoTo 0
         
         If bChecked Or lIconIndex > -1 Then
            lSelLeft = 4 + (tR.bottom - tR.top + 1 - 4)
         End If
                           
         If bHighlighted And Not bDisabled Then
            If m_eHighlightStyle = ECPHighlightStyleGradient Then
               ' Draw a gradient:
               LSet tWR = tR
               tWR.left = lSelLeft
               tWR.right = tWR.left + 4 + (tR.bottom - tR.top + 1 - 4)
               hBr = CreateSolidBrush(TranslateColor(vbHighlight))
               FillRect lHDC, tWR, hBr
               DeleteObject hBr
               LSet tWR = tR
               tWR.left = tWR.left + 4 + (tR.bottom - tR.top + 1 - 4)
               DrawGradient lHDC, tWR, TranslateColor(vbHighlight), TranslateColor(MenuBackgroundColor), False
            ElseIf m_eHighlightStyle = ECPHighlightStyleButton Then
               ' do nothing now
            Else
               ' standard:
               If (m_OfficeXPStyle) Then
                  LSet tWR = tR
                  tWR.left = tWR.left + 1
                  tWR.right = tWR.right - 2
                  fillWithLighterSelectedColour lHDC, tWR, tDIS.rcItem.top + tWR.top
                  DrawEdge lHDC, tWR, 0, 0, True
               Else
                  LSet tWR = tR
                  tWR.left = lSelLeft
                  fillWithHighlightBackColor lHDC, tWR, tDIS.rcItem.top + tWR.top
               End If
            End If
         End If
         
         If bDisabled Then
            SetTextColor lHDC, TranslateColor(vb3DHighlight)
         Else
            If bHighlighted Then
               SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
            Else
               SetTextColor lHDC, TranslateColor(InActiveMenuForeColor)
            End If
         End If
         
         ' Get the check/icon space:
         LSet tWR = tR
         If m_eHighlightStyle = ECPHighlightStyleButton Then
            InflateRect tWR, -2, -2
         Else
            tWR.left = tWR.left + 1
         End If
         tWR.right = tWR.left + (tWR.bottom - tWR.top + 1 - 2)
         
         ' Check:
         If bChecked Then
                        
            ' Colour in:
            If Not bHighlighted Then
               SetBkMode lHDC, OPAQUE
               If (m_OfficeXPStyle) Then
                  LSet tWRS = tWR
                  InflateRect tWRS, -1, -1
                  hBrush = CreateSolidBrush(BlendColor(vbHighlight, BlendColor(MenuBackgroundColor, vbButtonFace, 128), 40))
                  FillRect lHDC, tWRS, hBrush
                  DeleteObject hBrush
               Else
                  If (NoPalette) Then
                     hBrush = CreateSolidBrush(LighterColour(MenuBackgroundColor))
                     FillRect lHDC, tWR, hBrush
                     DeleteObject hBrush
                  Else
                     m_cBrush.Rectangle lHDC, tWR.left, tWR.top, tWR.right - tWR.left, tWR.bottom - tWR.top, 1, PATCOPY, True, MenuBackgroundColor, vb3DHighlight
                  End If
               End If
               SetBkMode lHDC, Transparent
                  
               If bDisabled Then
                  SetTextColor lHDC, TranslateColor(vb3DHighlight)
               End If
            ElseIf (m_OfficeXPStyle) Then
               LSet tWRS = tWR
               InflateRect tWRS, -1, -1
               hBrush = CreateSolidBrush(BlendColor(vbHighlight, MenuBackgroundColor, 128))
               FillRect lHDC, tWRS, hBrush
               DeleteObject hBrush
            End If
            If Not bDisabled Then
               If bHighlighted Then
                  SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
               Else
                  SetTextColor lHDC, TranslateColor(InActiveMenuForeColor)
               End If
            End If
            
            If (m_OfficeXPStyle) Then
               LSet tWRS = tWR
               InflateRect tWRS, -1, -1
               DrawEdge lHDC, tWRS, BDR_SUNKENOUTER, BF_RECT, m_OfficeXPStyle
            Else
               DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_RECT, m_OfficeXPStyle
            End If
                        
            If lIconIndex = -1 Then
               ' Draw the appropriate symbol:
               SelectObject lHDC, hFntOld
               hFntsOld = SelectObject(lHDC, hFontSymbol)
               If bHighlighted Then
                  SetTextColor lHDC, TranslateColor(InActiveMenuForeColor)
               End If
               If bRadioCheck Then
                  pDrawText lHDC, "h", tWR, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, bDisabled
               Else
                  pDrawText lHDC, "b", tWR, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, bDisabled
               End If
               SelectObject lHDC, hFntsOld
               hFntOld = SelectObject(lHDC, hFont)
               If bHighlighted Then
                  SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
               End If
            Else
               lX = tWR.left + (tWR.right - tWR.left + 1 - m_lIconSize) \ 2
               lY = tWR.top + (tWR.bottom - tWR.top + 1 - m_lIconSize) \ 2
               If bDisabled Then
                  ImageListDrawIconDisabled m_ptrVb6ImageList, lHDC, m_hIml, lIconIndex, lX, lY, m_lIconSize
               Else
                  ImageListDrawIcon m_ptrVb6ImageList, lHDC, m_hIml, lIconIndex, lX, lY
               End If
            End If
         Else
            If lIconIndex > -1 Then
               If bHighlighted And Not bDisabled Then
                  If (Not (m_eHighlightStyle = ECPHighlightStyleButton) And Not (m_OfficeXPStyle)) Then
                     DrawEdge lHDC, tWR, BDR_RAISEDINNER, BF_RECT, m_OfficeXPStyle
                  End If
               End If
               lX = tWR.left + (tWR.right - tWR.left + 1 - m_lIconSize) \ 2
               lY = tWR.top + (tWR.bottom - tWR.top + 1 - m_lIconSize) \ 2
               lX = lX + 2 * Abs(m_eHighlightStyle = ECPHighlightStyleButton)
               If bDisabled Then
                  ImageListDrawIconDisabled m_ptrVb6ImageList, lHDC, m_hIml, lIconIndex, lX, lY, m_lIconSize
               Else
                  If (m_OfficeXPStyle) Then
                     If (bHighlighted) Then
                        ImageListDrawIconDisabled m_ptrVb6ImageList, lHDC, m_hIml, lIconIndex, lX + 1, lY + 1, m_lIconSize, True
                        ImageListDrawIcon m_ptrVb6ImageList, lHDC, m_hIml, lIconIndex, lX - 1, lY - 1
                     Else
                        ImageListDrawIcon m_ptrVb6ImageList, lHDC, m_hIml, lIconIndex, lX, lY
                     End If
                  Else
                     ImageListDrawIcon m_ptrVb6ImageList, lHDC, m_hIml, lIconIndex, lX, lY
                  End If
               End If
            End If

         End If
         
                  
         ' Draw text:
         If m_eHighlightStyle = ECPHighlightStyleButton And Not (bDisabled) Then
            SetTextColor lHDC, TranslateColor(InActiveMenuForeColor)
         End If
         
         LSet tWR = tR
         tWR.left = 4 + (tR.bottom - tR.top + 1 - 4) + 2
         If (m_OfficeXPStyle) Then
            tWR.left = tWR.left + 4
         End If
         If lID > 0 And lID <= m_iRestore Then
            pDrawText lHDC, m_sCaption(lID), tWR, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, bDisabled
            If Len(m_sShortCut(lID)) > 0 Then
               tWR.left = tWR.left + m_lMenuTextSize(lID) + 4 + 4
               pDrawText lHDC, m_sShortCut(lID), tWR, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, bDisabled
            End If
         End If
         
         ' Highlighted:
         If bHighlighted And m_eHighlightStyle = ECPHighlightStyleButton And Not (bDisabled) Then
            LSet tWR = tR
            InflateRect tWR, -2, 0
            DrawEdge lHDC, tWR, BDR_RAISEDINNER, BF_RECT, m_OfficeXPStyle
         End If
         
      End If
      
      SelectObject lHDC, hFntOld
      
      BitBlt tDIS.hdc, tDIS.rcItem.left, tDIS.rcItem.top, tDIS.rcItem.right - tDIS.rcItem.left + 1, tDIS.rcItem.bottom - tDIS.rcItem.top + 1, lHDC, 0, 0, vbSrcCopy
      
   Else
      DrawItem = CallOldWindowProc(m_hWnd, WM_DRAWITEM, wParam, lParam)
   End If
   
End Function
Private Function pDrawText(ByVal lHDC As Long, ByVal sText As String, tR As RECT, ByVal dtFlags As Long, ByVal bDisabled As Boolean)
Dim tWR As RECT
   LSet tWR = tR
   If bDisabled Then
      If (m_OfficeXPStyle) Then
         SetTextColor lHDC, TranslateColor(vb3DShadow)
      Else
         SetTextColor lHDC, TranslateColor(vb3DHighlight)
         OffsetRect tWR, 1, 1
      End If
   End If
   DrawText lHDC, sText, -1, tWR, dtFlags
   If bDisabled Then
      If Not (m_OfficeXPStyle) Then
         SetTextColor lHDC, TranslateColor(vbButtonShadow)
         OffsetRect tWR, -1, -1
         DrawText lHDC, sText, -1, tWR, dtFlags
      End If
   End If
End Function

Private Sub DrawGradient( _
      ByVal hdc As Long, _
      ByRef rct As RECT, _
      ByVal lEndColour As Long, _
      ByVal lStartColour As Long, _
      ByVal bVertical As Boolean _
   )
Dim lStep As Long
Dim lPos As Long, lSize As Long
Dim bRGB(1 To 3) As Integer
Dim bRGBStart(1 To 3) As Integer
Dim dR(1 To 3) As Double
Dim dPos As Double ', d As Double
Dim hBr As Long
Dim tR As RECT
   
   LSet tR = rct
   If bVertical Then
      lSize = (tR.bottom - tR.top)
   Else
      lSize = (tR.right - tR.left)
   End If
   lStep = lSize \ 255
   If (lStep < 3) Then
       lStep = 3
   End If
       
   bRGB(1) = lStartColour And &HFF&
   bRGB(2) = (lStartColour And &HFF00&) \ &H100&
   bRGB(3) = (lStartColour And &HFF0000) \ &H10000
   bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
   dR(1) = (lEndColour And &HFF&) - bRGB(1)
   dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
   dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)
        
   For lPos = lSize To 0 Step -lStep
      ' Draw bar:
      If bVertical Then
         tR.top = tR.bottom - lStep
      Else
         tR.left = tR.right - lStep
      End If
      If tR.top < rct.top Then
         tR.top = rct.top
      End If
      If tR.left < rct.left Then
         tR.left = rct.left
      End If
      
      'Debug.Print tR.Right, tR.left, (bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1))
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hdc, tR, hBr
      DeleteObject hBr
            
      ' Adjust colour:
      dPos = ((lSize - lPos) / lSize)
      If bVertical Then
         tR.bottom = tR.top
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      Else
         tR.right = tR.left
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      End If
      
   Next lPos

End Sub

Private Sub TileArea( _
        ByVal hdcTo As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal hdcSrc As Long, _
        ByVal SrcWidth As Long, _
        ByVal SrcHeight As Long, _
        ByVal lOffsetY As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = (X Mod SrcWidth)
    lSrcStartY = ((Y + lOffsetY) Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = Y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (Y + Height)
        If (lDstY + lDstHeight) > (Y + Height) Then
            lDstHeight = Y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + Width)
            If (lDstX + lDstWidth) > (X + Width) Then
                lDstWidth = X + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdcTo, lDstX, lDstY, lDstWidth, lDstHeight, hdcSrc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop
End Sub

' Convert Automation color to Windows color
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
End Property


Private Function DrawEdge( _
      ByVal hdc As Long, _
      qrc As RECT, _
      ByVal edge As Long, _
      ByVal grfFlags As Long, _
      ByVal bOfficeXpStyle As Boolean _
   ) As Long
   If (bOfficeXpStyle) Then
      Dim junk As POINTAPI
      Dim hPenOld As Long
      Dim hPen As Long
      If (qrc.bottom > qrc.top) Then
         hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbHighlight))
      Else
         hPen = CreatePen(PS_SOLID, 1, TranslateColor(vb3DShadow))
      End If
      hPenOld = SelectObject(hdc, hPen)
      MoveToEx hdc, qrc.left, qrc.top, junk
      LineTo hdc, qrc.right - 1, qrc.top
      If (qrc.bottom > qrc.top) Then
         LineTo hdc, qrc.right - 1, qrc.bottom - 1
         LineTo hdc, qrc.left, qrc.bottom - 1
         LineTo hdc, qrc.left, qrc.top
      End If
      SelectObject hdc, hPenOld
      DeleteObject hPen
   Else
      DrawEdgeAPI hdc, qrc, edge, grfFlags
   End If
End Function


Private Sub ImageListDrawIcon( _
        ByVal ptrVb6ImageList As Long, _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        Optional ByVal bSelected As Boolean = False, _
        Optional ByVal bBlend25 As Boolean = False _
    )
Dim lFlags As Long
Dim lR As Long

    lFlags = ILD_TRANSPARENT
    If (bSelected) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    If (bBlend25) Then
        lFlags = lFlags Or ILD_BLEND25
    End If
    If (ptrVb6ImageList <> 0) Then
        Dim o As Object
        On Error Resume Next
        Set o = ObjectFromPtr(ptrVb6ImageList)
        If Not (o Is Nothing) Then
            If m_scmode = vbPixels Then
                o.ListImages(iIconIndex + 1).Draw hdc, lX, lY, lFlags
            Else
                o.ListImages(iIconIndex + 1).Draw hdc, lX * Screen.TwipsPerPixelX, lY * Screen.TwipsPerPixelY, lFlags
            End If
        End If
        On Error GoTo 0
    Else
        lR = ImageList_Draw( _
                hIml, _
                iIconIndex, _
                hdc, _
                lX, _
                lY, _
                lFlags)
        If (lR = 0) Then
            Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
        End If
    End If
End Sub
Private Sub ImageListDrawIconDisabled( _
        ByVal ptrVb6ImageList As Long, _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lSize As Long, _
        Optional ByVal asShadow As Boolean _
    )
Dim lR As Long
Dim hIcon As Long

   hIcon = 0
   If (ptrVb6ImageList <> 0) Then
      Dim o As Object
      On Error Resume Next
      Set o = ObjectFromPtr(ptrVb6ImageList)
      If Not (o Is Nothing) Then
         
         Dim lhDCDisp As Long
         Dim lHDC As Long
         Dim lhBmp As Long
         Dim lhBmpOld As Long
         Dim lhIml As Long
                  
         lhDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
         lHDC = CreateCompatibleDC(lhDCDisp)
         lhBmp = CreateCompatibleBitmap(lhDCDisp, o.ImageWidth, o.ImageHeight)
         DeleteDC lhDCDisp
         lhBmpOld = SelectObject(lHDC, lhBmp)
         o.ListImages.Item(iIconIndex + 1).Draw lHDC, 0, 0, 0
         SelectObject lHDC, lhBmpOld
         DeleteDC lHDC
         lhIml = ImageList_Create(o.ImageWidth, o.ImageHeight, ILC_MASK Or ILC_COLOR32, 1, 1)
         ImageList_AddMasked lhIml, lhBmp, TranslateColor(o.BackColor)
         DeleteObject lhBmp
         hIcon = ImageList_GetIcon(lhIml, 0, 0)
         ImageList_Destroy lhIml
         
      End If
      On Error GoTo 0
   Else
      hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
   End If
   If (hIcon <> 0) Then
      If (asShadow) Then
         Dim hBr As Long
         hBr = GetSysColorBrush(vb3DShadow And &H1F)
         lR = DrawState(hdc, hBr, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO)
         DeleteObject hBr
      Else
         lR = DrawState(hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
      End If
      DestroyIcon hIcon
   End If
   
End Sub
'Private Sub fillWithLighterBackColor(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long, ByVal bInfrequent As Boolean)
'Dim hBrush As Long
'   SetBkMode lHDC, OPAQUE
'   If (NoPalette) Then
'      If bInfrequent Then
'         hBrush = CreateSolidBrush(SlightlyLighterColour(MenuBackgroundColor))
'      Else
'         hBrush = CreateSolidBrush(LighterColour(MenuBackgroundColor))
'      End If
'      FillRect lHDC, tR, hBrush
'      DeleteObject hBrush
'   Else
'      m_cBrush.Rectangle lHDC, tR.left, tR.top, tR.right - tR.left + 1, tR.bottom - tR.top + 1, 1, PATCOPY, True, MenuBackgroundColor, vb3DHighlight
'   End If
'   SetBkMode lHDC, TRANSPARENT
'End Sub
Private Sub fillWithHighlightBackColor(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)
Dim hBr As Long
   hBr = CreateSolidBrush(TranslateColor(vbHighlight))
   FillRect lHDC, tR, hBr
   DeleteObject hBr
End Sub
'Private Sub fillWithNormalBackground(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)
'Dim hBrush As Long
'   hBrush = CreateSolidBrush(TranslateColor(MenuBackgroundColor))
'   FillRect lHDC, tR, hBrush
'   DeleteObject hBrush
'End Sub
Private Sub fillWithLighterControlColour(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)

Dim hBrush As Long
   SetBkMode lHDC, OPAQUE
   If (NoPalette) Then
      hBrush = CreateSolidBrush(BlendColor(MenuBackgroundColor, vbButtonFace, 128))
      FillRect lHDC, tR, hBrush
      DeleteObject hBrush
   Else
      m_cBrush.Rectangle lHDC, tR.left, tR.top, tR.right - tR.left + 1, tR.bottom - tR.top + 1, 1, PATCOPY, True, MenuBackgroundColor, vb3DHighlight
   End If
   SetBkMode lHDC, Transparent
End Sub

Private Sub fillWithLighterSelectedColour(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)

Dim hBrush As Long
   SetBkMode lHDC, OPAQUE
   If (NoPalette) Then
      hBrush = CreateSolidBrush(BlendColor(vbHighlight, MenuBackgroundColor, 80))
      FillRect lHDC, tR, hBrush
      DeleteObject hBrush
   Else
      m_cBrush.Rectangle lHDC, tR.left, tR.top, tR.right - tR.left + 1, tR.bottom - tR.top + 1, 1, PATCOPY, True, MenuBackgroundColor, vb3DHighlight
   End If
   SetBkMode lHDC, Transparent
End Sub
Private Property Get BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal alpha As Long = 128 _
   ) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = TranslateColor(oColorFrom)
   lCTo = TranslateColor(oColorTo)
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB( _
      ((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), _
      ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), _
      ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255) _
      )
      
End Property
'Private Property Get SlightlyLighterColour(ByVal oColor As OLE_COLOR) As Long
'Dim lC As Long
'Dim H As Single, s As Single, l As Single
'Dim lR As Long, lG As Long, lB As Long
'Static s_lColLast As Long
'Static s_lLightColLast As Long
'
'   lC = TranslateColor(oColor)
'   If (lC <> s_lColLast) Then
'      s_lColLast = lC
'      RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, H, s, l
'      If (l > 0.99) Then
'         l = l * 0.95
'      Else
'         l = l * 1.05
'         If (l > 1) Then
'            l = 1
'         End If
'      End If
'      HLSToRGB H, s, l, lR, lG, lB
'      s_lLightColLast = RGB(lR, lG, lB)
'   End If
'   SlightlyLighterColour = s_lLightColLast
'End Property

Private Property Get LighterColour(ByVal oColor As OLE_COLOR) As Long
Dim lC As Long
Dim H As Single, s As Single, l As Single
Dim lR As Long, lG As Long, lB As Long
Static s_lColLast As Long
Static s_lLightColLast As Long
   
   lC = TranslateColor(oColor)
   If (lC <> s_lColLast) Then
      s_lColLast = lC
      RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, H, s, l
      If (l > 0.99) Then
         l = l * 0.8
      Else
         l = l * 1.2
         If (l > 1) Then
            l = 1
         End If
      End If
      HLSToRGB H, s, l, lR, lG, lB
      s_lLightColLast = RGB(lR, lG, lB)
   End If
   LighterColour = s_lLightColLast
End Property
 
Private Property Get NoPalette(Optional ByVal bForce As Boolean = False) As Boolean
Static bOnce As Boolean
Static bNoPalette As Boolean
Dim lHDC As Long
Dim lBits As Long
   If (bForce) Then
      bOnce = False
   End If
   If Not (bOnce) Then
      lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      If (lHDC <> 0) Then
         lBits = GetDeviceCaps(lHDC, BITSPIXEL)
         If (lBits <> 0) Then
            bOnce = True
         End If
         bNoPalette = (lBits > 8)
         DeleteDC lHDC
      End If
   End If
   NoPalette = bNoPalette
End Property

Private Sub RGBToHLS( _
     ByVal R As Long, ByVal g As Long, ByVal b As Long, _
     H As Single, s As Single, l As Single _
     )
 Dim Max As Single
 Dim min As Single
 Dim Delta As Single
 Dim rR As Single, rG As Single, rB As Single

     rR = R / 255: rG = g / 255: rB = b / 255

 '{Given: rgb each in [0,1].
 ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
         Max = Maximum(rR, rG, rB)
         min = Minimum(rR, rG, rB)
             l = (Max + min) / 2 '{This is the lightness}
         '{Next calculate saturation}
         If Max = min Then
             'begin {Acrhomatic case}
             s = 0
             H = 0
             'end {Acrhomatic case}
         Else
             'begin {Chromatic case}
                 '{First calculate the saturation.}
             If l <= 0.5 Then
                 s = (Max - min) / (Max + min)
             Else
                 s = (Max - min) / (2 - Max - min)
             End If
             '{Next calculate the hue.}
             Delta = Max - min
             If rR = Max Then
                     H = (rG - rB) / Delta '{Resulting color is between yellow and magenta}
             ElseIf rG = Max Then
                 H = 2 + (rB - rR) / Delta '{Resulting color is between cyan and yellow}
             ElseIf rB = Max Then
                 H = 4 + (rR - rG) / Delta '{Resulting color is between magenta and cyan}
             End If
         'end {Chromatic Case}
     End If
 End Sub

 Private Sub HLSToRGB( _
     ByVal H As Single, ByVal s As Single, ByVal l As Single, _
     R As Long, g As Long, b As Long _
     )
 Dim rR As Single, rG As Single, rB As Single
 Dim min As Single, Max As Single

     If s = 0 Then
     ' Achromatic case:
     rR = l: rG = l: rB = l
     Else
     ' Chromatic case:
     ' delta = Max-Min
     If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         min = l * (1 - s)
     Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         min = l - s * (1 - l)
     End If
     ' Get the Max value:
     Max = 2 * l - min
     
     ' Now depending on sector we can evaluate the h,l,s:
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
     R = rR * 255: g = rG * 255: b = rB * 255
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




