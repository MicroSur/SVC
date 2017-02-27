VERSION 5.00
Begin VB.UserControl MyFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   ForwardFocus    =   -1  'True
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   198
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   Begin VB.Frame Fr 
      Caption         =   "Movies ("
      ClipControls    =   0   'False
      Height          =   2835
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "MyFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum GROUPBOXBORDERSTYLE
    None
    [Fixed Single]
End Enum

Public Enum GROUPBOXBACKSTYLE
    Transparent
    Solid
End Enum

Public Enum GROUPBOXAPPEARANCE
    Flat
    [3D]
End Enum

Private Sub UserControl_AmbientChanged(PropertyName As String)
'    DrawFrame
End Sub
Public Property Get hwnd()
    hwnd = UserControl.hwnd
End Property
Private Sub UserControl_Resize()
If ExitSVC Then Exit Sub
    DrawFrame
End Sub

Private Sub DrawFrame()

'Fr.Move 0, 0, ScaleX(UserControl.Width, vbTwips, vbPixels), ScaleY(UserControl.Height, vbTwips, vbPixels)
Fr.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawFrame
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
    DrawFrame
End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
'    m_BorderStyle = m_def_BorderStyle
'    m_Caption = Ambient.DisplayName
'    m_Appearance = m_def_Appearance
    Enabled = True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If ExitSVC Then Exit Sub

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Fr.Caption = PropBag.ReadProperty("Caption", "Movies (")
    Fr.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Fr.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Fr.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Fr.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 3133)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 2964)
End Sub

Private Sub UserControl_Show()
    DrawFrame
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", Fr.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Fr.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", Fr.Enabled, True)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 3133)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 2964)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Fr,Fr,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Fr.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Fr.Caption() = New_Caption
    PropertyChanged "Caption"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Fr,Fr,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Fr.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Fr.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Fr,Fr,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Fr.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Fr.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Fr,Fr,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Fr.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Fr.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Fr,Fr,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Fr.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Fr.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,hDC
'Public Property Get hdc() As Long
'    hdc = UserControl.hdc
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,hWnd
'Public Property Get hwnd() As Long
'    hwnd = UserControl.hwnd
'End Property

