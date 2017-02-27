Attribute VB_Name = "ModMSGBox"

'Programmer   : Chay Luna
'Email Addr   : chay_luna@yahoo.com
'Description  : Positional message box (myMsgBox)

Option Explicit

Private Const GWL_WNDPROC = (-4)

Private Const WM_NCACTIVATE = &H86
Private Const WA_INACTIVE = 0

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Enum myMsgBoxPositioning
  RelativeToOwer
  RelativeToScreen
End Enum

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Dim m_PrevProc&

Dim m_xPos As Long, m_yPos As Long
Dim m_xPosCenter As Boolean, m_yPosCenter As Boolean
Dim m_Positioning As myMsgBoxPositioning


Private Sub SubClass(ByVal hwnd As Long)
  m_PrevProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub UnSubClass(ByVal hwnd As Long)
  SetWindowLong hwnd, GWL_WNDPROC, m_PrevProc
End Sub


'This is the new window proc of the owner/parent window
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim nRet&, sClassName$
Dim rOwner As RECT, R As RECT

  WindowProc = CallWindowProc(m_PrevProc, hwnd, uMsg, wParam, lParam)
    
  If uMsg = WM_NCACTIVATE And wParam = WA_INACTIVE Then
      
    'Det window class name
    sClassName = Space$(128)
    nRet = GetClassName(lParam, sClassName, 128)
    sClassName = Left$(sClassName, nRet)
      
    'Check for dialog box
    If sClassName = "#32770" Then
      
      'Get owner rectangle
      If m_Positioning = RelativeToOwer Then
        GetWindowRect hwnd, rOwner
      Else
        GetWindowRect GetDesktopWindow(), rOwner
      End If
        
      'Get msgbox rectangle
      GetWindowRect lParam, R
                  
      'Compute x and y
      If m_xPosCenter Then
        m_xPos = ((rOwner.Right - rOwner.Left) - (R.Right - R.Left)) \ 2 + rOwner.Left
      Else
        m_xPos = m_xPos + rOwner.Left
      End If
      If m_yPosCenter Then
        m_yPos = ((rOwner.Bottom - rOwner.Top) - (R.Bottom - R.Top)) \ 2 + rOwner.Top
      Else
        m_yPos = m_yPos + rOwner.Top
      End If
                            
      'Position dialog box
      SetWindowPos lParam, 0&, m_xPos, m_yPos, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE
                
    End If
      
  End If
        
End Function
 
'This is the desired function
Public Function myMsgBox(ByVal sPrompt As String, Optional ByVal nButtons As VbMsgBoxStyle = vbOKOnly, _
                         Optional ByVal sTitle, Optional ByVal hWndOwner, _
                         Optional ByVal Positioning As myMsgBoxPositioning = RelativeToOwer, _
                         Optional ByVal xPos As Variant = "Center", Optional ByVal yPos As Variant = "Center") As VbMsgBoxResult

  'Default title
  If IsMissing(sTitle) Then sTitle = App.title
  
  If IsMissing(hWndOwner) Then
    
    'Simple call the msgbox when no hWndOwner
    myMsgBox = MsgBox(sPrompt, nButtons, sTitle)
    
  Else
        
    'Get values
    m_xPosCenter = StrComp(xPos, "Center", vbTextCompare) = 0
    m_yPosCenter = StrComp(yPos, "Center", vbTextCompare) = 0
    If Not m_xPosCenter Then m_xPos = xPos
    If Not m_yPosCenter Then m_yPos = yPos
    m_Positioning = Positioning
    
    'Subclass parent window
    SubClass hWndOwner
    
    'Simply call the normal vb msgbox function (See WindowProc for some actions)
    myMsgBox = MsgBox(sPrompt, nButtons, sTitle)
    
    'Remove subclassing
    UnSubClass hWndOwner
    
  End If
  
End Function
Public Function myInputBox(ByVal sPrompt As String, Optional ByVal sTitle, Optional ByVal hWndOwner, _
                         Optional ByVal Positioning As myMsgBoxPositioning = RelativeToOwer, _
                         Optional ByVal xPos As Variant = "Center", Optional ByVal yPos As Variant = "Center", _
                         Optional ByVal sDefault As String) _
                         As String

  'Default title
  If IsMissing(sTitle) Then sTitle = App.title
  
  If IsMissing(hWndOwner) Then
    'Simple call the msgbox when no hWndOwner
    myInputBox = InputBox(sPrompt, , sDefault)

  Else
        
    'Get values
    m_xPosCenter = StrComp(xPos, "Center", vbTextCompare) = 0
    m_yPosCenter = StrComp(yPos, "Center", vbTextCompare) = 0
    If Not m_xPosCenter Then m_xPos = xPos
    If Not m_yPosCenter Then m_yPos = yPos
    m_Positioning = Positioning
    
    'Subclass parent window
    SubClass hWndOwner
    
    'Simply call the normal vb msgbox function (See WindowProc for some actions)
    myInputBox = InputBox(sPrompt, , sDefault)
    
    'Remove subclassing
    UnSubClass hWndOwner
    
  End If
  
End Function

