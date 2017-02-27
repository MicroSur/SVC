VERSION 5.00
Begin VB.UserControl UCVMSUR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ControlContainer=   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   1500
   Begin VB.PictureBox PicVM 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   9570
      Left            =   0
      ScaleHeight     =   638
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      Begin VB.PictureBox picVMB1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   8
         Top             =   240
         Width           =   960
      End
      Begin VB.PictureBox picVMB2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   195
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   7
         Top             =   1500
         Width           =   795
      End
      Begin VB.PictureBox picVMB3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   210
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   6
         Top             =   2610
         Width           =   795
      End
      Begin VB.PictureBox picVMB4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   195
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   5
         Top             =   3645
         Width           =   795
      End
      Begin VB.PictureBox picVMB5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   210
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   4
         Top             =   4725
         Width           =   795
      End
      Begin VB.PictureBox picVMB6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   195
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   3
         Top             =   5760
         Width           =   795
      End
      Begin VB.PictureBox picVMB7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   210
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   2
         Top             =   6780
         Width           =   795
      End
      Begin VB.PictureBox picVMB8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   195
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   1
         Top             =   7770
         Width           =   795
      End
      Begin VB.Label LVMB 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "New"
         Height          =   210
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   3420
         Width           =   1050
      End
      Begin VB.Label LVMB 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Cover"
         Height          =   210
         Index           =   3
         Left            =   60
         TabIndex        =   15
         Top             =   4440
         Width           =   1050
      End
      Begin VB.Label LVMB 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "People"
         Height          =   210
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   5520
         Width           =   1050
      End
      Begin VB.Label LVMB 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Options"
         Height          =   210
         Index           =   5
         Left            =   60
         TabIndex        =   13
         Top             =   6540
         Width           =   1050
      End
      Begin VB.Label LVMB 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Statistic"
         Height          =   210
         Index           =   6
         Left            =   45
         TabIndex        =   12
         Top             =   7545
         Width           =   1050
      End
      Begin VB.Label LVMB 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "X"
         Height          =   210
         Index           =   7
         Left            =   60
         TabIndex        =   11
         Top             =   8550
         Width           =   1050
      End
      Begin VB.Label LVMB 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Edit"
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   2355
         Width           =   1050
      End
      Begin VB.Label LVMB 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "View"
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   1140
         Width           =   1290
      End
      Begin VB.Image imgVMB 
         Enabled         =   0   'False
         Height          =   720
         Index           =   0
         Left            =   1920
         Picture         =   "UCVMSUR.ctx":0000
         Top             =   300
         Width           =   720
      End
      Begin VB.Image imgVMB 
         Enabled         =   0   'False
         Height          =   720
         Index           =   1
         Left            =   1920
         Picture         =   "UCVMSUR.ctx":0ECA
         Top             =   1440
         Width           =   720
      End
      Begin VB.Image imgVMB 
         Enabled         =   0   'False
         Height          =   720
         Index           =   2
         Left            =   1860
         Picture         =   "UCVMSUR.ctx":1D94
         Top             =   2400
         Width           =   720
      End
      Begin VB.Image imgVMB 
         Enabled         =   0   'False
         Height          =   720
         Index           =   3
         Left            =   1860
         Picture         =   "UCVMSUR.ctx":2C5E
         Top             =   3420
         Width           =   720
      End
      Begin VB.Image imgVMB 
         Enabled         =   0   'False
         Height          =   720
         Index           =   4
         Left            =   1860
         Picture         =   "UCVMSUR.ctx":3B28
         Top             =   4455
         Width           =   720
      End
      Begin VB.Image imgVMB 
         Enabled         =   0   'False
         Height          =   720
         Index           =   5
         Left            =   1920
         Picture         =   "UCVMSUR.ctx":49F2
         Top             =   5520
         Width           =   720
      End
      Begin VB.Image imgVMB 
         Enabled         =   0   'False
         Height          =   720
         Index           =   6
         Left            =   1920
         Picture         =   "UCVMSUR.ctx":58BC
         Top             =   6795
         Width           =   720
      End
      Begin VB.Image imgVMB 
         Enabled         =   0   'False
         Height          =   720
         Index           =   7
         Left            =   1860
         Picture         =   "UCVMSUR.ctx":6786
         Top             =   7920
         Width           =   720
      End
   End
End
Attribute VB_Name = "UCVMSUR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private MouseDownPressed As Boolean 'был ли мд до апа, для клика

'VerticalMenu for SurVideoCatalog
'width = 1230
'PicVM.width = 81
Private WithEvents m_cVMB1 As cMouseTrack
Attribute m_cVMB1.VB_VarHelpID = -1
Private WithEvents m_cVMB2 As cMouseTrack
Attribute m_cVMB2.VB_VarHelpID = -1
Private WithEvents m_cVMB3 As cMouseTrack
Attribute m_cVMB3.VB_VarHelpID = -1
Private WithEvents m_cVMB4 As cMouseTrack
Attribute m_cVMB4.VB_VarHelpID = -1
Private WithEvents m_cVMB5 As cMouseTrack
Attribute m_cVMB5.VB_VarHelpID = -1
Private WithEvents m_cVMB6 As cMouseTrack
Attribute m_cVMB6.VB_VarHelpID = -1
Private WithEvents m_cVMB7 As cMouseTrack
Attribute m_cVMB7.VB_VarHelpID = -1
Private WithEvents m_cVMB8 As cMouseTrack
Attribute m_cVMB8.VB_VarHelpID = -1

'Events
Event MenuItemClick(MenuItem As Integer, Sh As Integer)
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hwnd()
    hwnd = PicVM.hwnd
End Property

Public Property Get hwnd1()
    hwnd1 = picVMB1.hwnd
End Property
Public Property Get hwnd2()
    hwnd2 = picVMB2.hwnd
End Property
Public Property Get hwnd3()
    hwnd3 = picVMB3.hwnd
End Property
Public Property Get hwnd4()
    hwnd4 = picVMB4.hwnd
End Property
Public Property Get hwnd5()
    hwnd5 = picVMB5.hwnd
End Property
Public Property Get hwnd6()
    hwnd6 = picVMB6.hwnd
End Property
Public Property Get hwnd7()
    hwnd7 = picVMB7.hwnd
End Property
Public Property Get hwnd8()
    hwnd8 = picVMB8.hwnd
End Property



Private Sub UserControl_Initialize()
SetClolor
SetVMBIcons
VMBArrange

End Sub

'Initialize Properties for User Control
Public Sub SetClolor()
Dim i As Integer

   For i = 0 To 7
   LVMB(i).BackColor = picVMB1.Container.BackColor
   LVMB(i).ForeColor = picVMB1.Container.ForeColor
   Next i
   
   picVMB1.BackColor = picVMB1.Container.BackColor
   picVMB2.BackColor = picVMB1.Container.BackColor
   picVMB3.BackColor = picVMB1.Container.BackColor
   picVMB4.BackColor = picVMB1.Container.BackColor
   picVMB5.BackColor = picVMB1.Container.BackColor
   picVMB6.BackColor = picVMB1.Container.BackColor
   picVMB7.BackColor = picVMB1.Container.BackColor
   picVMB8.BackColor = picVMB1.Container.BackColor

Get3DButtonColors picVMB1.Container.BackColor
End Sub

Private Sub SetVMBIcons()
Dim l As Long, t As Long


t = (picVMB1.ScaleWidth - imgVMB(0).Width) \ 2
l = (picVMB1.ScaleHeight - imgVMB(0).Height) \ 2

   Set imgVMB(0).Container = picVMB1
   imgVMB(0).Move t, l
   Set m_cVMB1 = New cMouseTrack
   m_cVMB1.AttachMouseTracking picVMB1
   
   Set imgVMB(1).Container = picVMB2
   imgVMB(1).Move t, l
   Set m_cVMB2 = New cMouseTrack
   m_cVMB2.AttachMouseTracking picVMB2

   Set imgVMB(2).Container = picVMB3
   imgVMB(2).Move t, l
   Set m_cVMB3 = New cMouseTrack
   m_cVMB3.AttachMouseTracking picVMB3

   Set imgVMB(3).Container = picVMB4
   imgVMB(3).Move t, l
   Set m_cVMB4 = New cMouseTrack
   m_cVMB4.AttachMouseTracking picVMB4

   Set imgVMB(4).Container = picVMB5
   imgVMB(4).Move t, l
   Set m_cVMB5 = New cMouseTrack
   m_cVMB5.AttachMouseTracking picVMB5

   Set imgVMB(5).Container = picVMB6
   imgVMB(5).Move t, l
   Set m_cVMB6 = New cMouseTrack
   m_cVMB6.AttachMouseTracking picVMB6

   Set imgVMB(6).Container = picVMB7
   imgVMB(6).Move t, l
   Set m_cVMB7 = New cMouseTrack
   m_cVMB7.AttachMouseTracking picVMB7

   Set imgVMB(7).Container = picVMB8
   imgVMB(7).Move t, l
   Set m_cVMB8 = New cMouseTrack
   m_cVMB8.AttachMouseTracking picVMB8
   



'm_cVMB1.optMethod(m_cVMB1.Method).Tag = "CODE"
'optMethod(m_cVMB1.Method).Value = True

End Sub

Private Sub VMBArrange()
Dim l As Long, t As Long
Dim w As Long, H As Long
Dim i As Integer, k As Integer, M As Integer, lH As Integer, o As Integer

On Error Resume Next
If kDPI <= 0 Then kDPI = 1
'UserControl.ScaleWidth = PicVM.Width
'PicVM.Height = UserControl.Height ' Screen.Height

t = (picVMB1.Container.ScaleWidth - picVMB1.ScaleWidth) \ 2
l = 26 * kDPI 'отступ до первой картинки
k = 1 + picVMB1.ScaleHeight 'отступ + высота картинки до лейбла
M = 5 'отступ после лейбла до след картинки

w = picVMB1.Width
H = picVMB1.Height
lH = LVMB(0).Height ' высота метки

o = picVMB1.ScaleHeight + lH + M 'высота картинка + метка

For i = 0 To 7
LVMB(i).Width = PicVM.ScaleWidth
LVMB(i).BackColor = picVMB1.BackColor
LVMB(i).ForeColor = vbWhite
Next

'1
picVMB1.Move t, l
LVMB(0).Move 0, l + k
l = l + o

'2
picVMB2.Height = H: picVMB2.Width = w
picVMB2.Move t, l
LVMB(1).Move 0, l + k
l = l + o

'3
picVMB3.Height = H: picVMB3.Width = w
picVMB3.Move t, l
LVMB(2).Move 0, l + k
l = l + o

'4
picVMB4.Height = H: picVMB4.Width = w
picVMB4.Move t, l
LVMB(3).Move 0, l + k
l = l + o

'5
picVMB5.Height = H: picVMB5.Width = w
picVMB5.Move t, l
LVMB(4).Move 0, l + k
l = l + o

'6
picVMB6.Height = H: picVMB6.Width = w
picVMB6.Move t, l
LVMB(5).Move 0, l + k
l = l + o

'7
picVMB7.Height = H: picVMB7.Width = w
picVMB7.Move t, l
LVMB(6).Move 0, l + k
l = l + o

'8
picVMB8.Height = H: picVMB8.Width = w
picVMB8.Move t, l
LVMB(7).Move 0, l + k
'l = l + o

End Sub
'vmb1
Private Sub m_cVMB1_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   ' Hover event, support for user32 and comctl32 methods:
   m_cVMB1.StartMouseTracking
End Sub
'vmb1
Private Sub m_cVMB1_MouseLeave()
   ' End tracking:
   Draw3DEffect picVMB1, 2
End Sub
'vmb1
Private Sub picVMB1_Click()
'RaiseEvent MenuItemClick(1, Shift)
End Sub

'vmb1
Private Sub picVMB1_DblClick()
Draw3DEffect picVMB1, 1
'picVMB1_Click
End Sub

'vmb1
Private Sub picVMB1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB1, 1
'MouseDownPressed = True
End Sub

'vmb1
Private Sub picVMB1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Tracking is initialised by entering the control:
   If Not (m_cVMB1.Tracking) Then
      m_cVMB1.StartMouseTracking
      Draw3DEffect picVMB1, 0
   End If

End Sub
'vmb1
Private Sub picVMB1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB1, 2 '0
'If (X > picVMB1.Left) And (X < (picVMB1.Width + picVMB1.Left)) Then
'If MouseDownPressed Then
RaiseEvent MenuItemClick(1, Shift)
'End If
End Sub

'vmb2
Private Sub m_cVMB2_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
m_cVMB2.StartMouseTracking
End Sub
'vmb2
Private Sub m_cVMB2_MouseLeave()
Draw3DEffect picVMB2, 2
End Sub
'vmb2
Private Sub picVMB2_Click()
'RaiseEvent MenuItemClick(2)
End Sub

'vmb2
Private Sub picVMB2_DblClick()
Draw3DEffect picVMB2, 1
'picVMB1_Click
End Sub

'vmb2
Private Sub picVMB2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB2, 1
End Sub

'vmb2
Private Sub picVMB2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not (m_cVMB2.Tracking) Then
      m_cVMB2.StartMouseTracking
      Draw3DEffect picVMB2, 0
   End If

End Sub
'vmb2
Private Sub picVMB2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB2, 2
RaiseEvent MenuItemClick(2, Shift)
End Sub



'vmb3
Private Sub m_cVMB3_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   ' Hover event, support for user32 and comctl32 methods:
   m_cVMB3.StartMouseTracking
End Sub
'vmb3
Private Sub m_cVMB3_MouseLeave()
   ' End tracking:
   Draw3DEffect picVMB3, 2
End Sub
'vmb3
Private Sub picVMB3_Click()
'RaiseEvent MenuItemClick(3
End Sub

'vmb3
Private Sub picVMB3_DblClick()
Draw3DEffect picVMB3, 1
End Sub

'vmb3
Private Sub picVMB3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB3, 1
End Sub

'vmb3
Private Sub picVMB3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Tracking is initialised by entering the control:
   If Not (m_cVMB3.Tracking) Then
      m_cVMB3.StartMouseTracking
      Draw3DEffect picVMB3, 0
   End If

End Sub
'vmb3
Private Sub picVMB3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB3, 2
RaiseEvent MenuItemClick(3, Shift)
End Sub

'vmb4
Private Sub m_cVMB4_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   ' Hover event, support for user32 and comctl32 methods:
   m_cVMB4.StartMouseTracking
End Sub
'vmb4
Private Sub m_cVMB4_MouseLeave()
   ' End tracking:
   Draw3DEffect picVMB4, 2
End Sub
'vmb4
Private Sub picVMB4_Click()
'RaiseEvent MenuItemClick(4)
End Sub

'vmb4
Private Sub picVMB4_DblClick()
Draw3DEffect picVMB4, 1
'picVMB4_Click
End Sub

'vmb4
Private Sub picVMB4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB4, 1
End Sub

'vmb4
Private Sub picVMB4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Tracking is initialised by entering the control:
   If Not (m_cVMB4.Tracking) Then
      m_cVMB4.StartMouseTracking
      Draw3DEffect picVMB4, 0
   End If

End Sub
'vmb4
Private Sub picVMB4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB4, 2
RaiseEvent MenuItemClick(4, Shift)
End Sub

'vmb5
Private Sub m_cVMB5_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   ' Hover event, support for user32 and comctl32 methods:
   m_cVMB5.StartMouseTracking
End Sub
'vmb5
Private Sub m_cVMB5_MouseLeave()
   ' End tracking:
   Draw3DEffect picVMB5, 2
End Sub
'vmb5
Private Sub picVMB5_Click()
'RaiseEvent MenuItemClick(5)
End Sub

'vmb5
Private Sub picVMB5_DblClick()
Draw3DEffect picVMB5, 1
'picVMB5_Click
End Sub

'vmb5
Private Sub picVMB5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB5, 1
End Sub

'vmb5
Private Sub picVMB5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Tracking is initialised by entering the control:
   If Not (m_cVMB5.Tracking) Then
      m_cVMB5.StartMouseTracking
      Draw3DEffect picVMB5, 0
   End If

End Sub
'vmb5
Private Sub picVMB5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB5, 2
RaiseEvent MenuItemClick(5, Shift)
End Sub

'vmb6
Private Sub m_cVMB6_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   ' Hover event, support for user32 and comctl32 methods:
   m_cVMB6.StartMouseTracking
End Sub
'vmb6
Private Sub m_cVMB6_MouseLeave()
   ' End tracking:
   Draw3DEffect picVMB6, 2
End Sub
'vmb6
Private Sub picVMB6_Click()
'RaiseEvent MenuItemClick(6)
End Sub

'vmb6
Private Sub picVMB6_DblClick()
Draw3DEffect picVMB6, 1
'picVMB6_Click
End Sub

'vmb6
Private Sub picVMB6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB6, 1
End Sub

'vmb6
Private Sub picVMB6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Tracking is initialised by entering the control:
   If Not (m_cVMB6.Tracking) Then
      m_cVMB6.StartMouseTracking
      Draw3DEffect picVMB6, 0
   End If

End Sub
'vmb6
Private Sub picVMB6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB6, 2
RaiseEvent MenuItemClick(6, Shift)
End Sub

'vmb7
Private Sub m_cVMB7_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   ' Hover event, support for user32 and comctl32 methods:
   m_cVMB7.StartMouseTracking
End Sub
'vmb7
Private Sub m_cVMB7_MouseLeave()
   ' End tracking:
   Draw3DEffect picVMB7, 2
End Sub
'vmb7
Private Sub picVMB7_Click()
'RaiseEvent MenuItemClick(7)
End Sub

'vmb7
Private Sub picVMB7_DblClick()
Draw3DEffect picVMB7, 1
'picVMB7_Click
End Sub

'vmb7
Private Sub picVMB7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB7, 1
End Sub

'vmb7
Private Sub picVMB7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Tracking is initialised by entering the control:
   If Not (m_cVMB7.Tracking) Then
      m_cVMB7.StartMouseTracking
      Draw3DEffect picVMB7, 0
   End If

End Sub
'vmb7
Private Sub picVMB7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB7, 2 '0
RaiseEvent MenuItemClick(7, Shift)
End Sub

'vmb8
Private Sub m_cVMB8_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   ' Hover event, support for user32 and comctl32 methods:
   m_cVMB8.StartMouseTracking
End Sub
'vmb8
Private Sub m_cVMB8_MouseLeave()
   ' End tracking:
   Draw3DEffect picVMB8, 2
End Sub
'vmb8
Private Sub picVMB8_Click()
'RaiseEvent MenuItemClick(8)
End Sub

'vmb8
Private Sub picVMB8_DblClick()
Draw3DEffect picVMB8, 1
'picVMB8_Click
End Sub

'vmb8
Private Sub picVMB8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB8, 1
End Sub

'vmb8
Private Sub picVMB8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Tracking is initialised by entering the control:
   If Not (m_cVMB8.Tracking) Then
      m_cVMB8.StartMouseTracking
      Draw3DEffect picVMB8, 0
   End If

End Sub
'vmb8
Private Sub picVMB8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw3DEffect picVMB8, 2
RaiseEvent MenuItemClick(8, Shift)
End Sub

Private Sub UserControl_Resize()
PicVM.Height = UserControl.Height
PicVM.Width = UserControl.Width
End Sub
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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 93)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 748)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 93)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 748)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

