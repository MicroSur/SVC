VERSION 5.00
Begin VB.Form FormShowPic 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "FormShowPic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   173
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TWait 
      Interval        =   30000
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer TMove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   660
      Top             =   120
   End
   Begin VB.PictureBox PicHB 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   300
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   1935
   End
End
Attribute VB_Name = "FormShowPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'&H0088FFFF& цвет шрифта желтенький

' таскать форму
Private Const LP_HT_CAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
'Private Const WM_NCLBUTTONUP = &HA2

'Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private FormDrag As Boolean
Private MDown As Boolean
'Private mcap As Boolean
'Private old_x As Single
'Private old_y As Single

Public WithEvents hb_cScroll As cScrollBars
Attribute hb_cScroll.VB_VarHelpID = -1
'scrsaver
Private XStep As Single
Private YStep As Single
Private hb_cScrollValue As Long
Private hsbHeight As Long 'размер скролла в системе

Private Sub Form_Activate()
'если это все убрать, мы не получим возврат фокуса листу, если листали с 'нет картинки' - в 'есть картинка'
If frmEditorFlag Then Exit Sub 'не в редакторе
If FrmMain.FrameView.Visible Then
 If Not PicManualFlag Then
  If Not MDown Then
  'фокус возвращался мейну, до первого клика на скролл, когда устанавливался ShowPicFocus
  If Not ShowPicFocus Then
  On Error Resume Next
    FrmMain.SetFocus
  End If
  End If
 End If
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'> 39
'< 37
'1 49 97 35
'2 50 98 40
'3 51 99 34
'If KeyCode = 49 Then
'Unload Me

'Debug.Print KeyCode
If Not (addflag Or editFlag) Then
 Select Case KeyCode
 Case 37 'l
  If hb_cScroll.Value(efsHorizontal) - 1 >= hb_cScroll.min(efsHorizontal) Then hb_cScroll.Value(efsHorizontal) = hb_cScroll.Value(efsHorizontal) - 1
 Case 39 'r
  If hb_cScroll.Value(efsHorizontal) + 1 <= hb_cScroll.Max(efsHorizontal) Then hb_cScroll.Value(efsHorizontal) = hb_cScroll.Value(efsHorizontal) + 1
 End Select

 'FrmMain.Timer2.Enabled = True
End If
 
TicScrSaver = GetTickCount
TMove.Enabled = False
ShowPicFocus = True

If KeyCode = 27 Then
    Unload Me
End If

End Sub

Private Sub Form_Load()
FormShowPicLoaded = True
PicHB.Visible = True
PicHB.Move Me.ScaleWidth - 97, Me.ScaleHeight - 21, 97, 21

Set hb_cScroll = New cScrollBars
hb_cScroll.Create PicHB.hwnd, True, False
hb_cScroll.style = efsEncarta
hb_cScroll.Orientation = efsoHorizontal
hb_cScroll.min(efsHorizontal) = 1
'hb_cScroll.Max(efsHorizontal) = 3 '2
hb_cScroll.LargeChange(efsHorizontal) = 1
hb_cScroll.SmallChange(efsHorizontal) = 1

hb_cScroll.Visible(efsHorizontal) = True
FormShowPic.FontBold = True

'узнать размеры системных скроллов
'Const SM_CXVSCROLL = 2
'Const SM_CYHSCROLL = 3
'hsbHeight = GetSystemMetrics(SM_CYHSCROLL)
ScaleMode = 3
hsbHeight = GetSystemMetrics(3)
End Sub

Private Sub Form_LostFocus()
MDown = False
'ShowPicFocus = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

MDown = True
'FormDrag = True

If Button = 2 Then
    FrmMain.PopupMenu FrmMain.PopShowPicHid ', DefaultMenu:=FrmMain.mnuKillPic просто толстит
    Exit Sub
End If

'If Not (addflag Or editFlag) Then
'    'FrmMain.Timer2.Enabled = True
'End If

TicScrSaver = GetTickCount
TMove.Enabled = False
'old_x = x: old_y = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

frmShow_xPos = FormShowPic.left
frmShow_yPos = FormShowPic.top

If IsCoverShowFlag Then
    CoverWindTop = FormShowPic.top
    CoverWindLeft = FormShowPic.left
Else
    ScrShotWindTop = FormShowPic.top
    ScrShotWindLeft = FormShowPic.left
End If

'Debug.Print ScrShotWindTop
'FormDrag = False

'таскать

If MDown Then
ReleaseCapture
    'DoEvents
    'FormDrag = True
    'If old_x <> x Then
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, LP_HT_CAPTION, ByVal 0&
End If
   '  SendMessage Me.hWnd, &HA3, LP_HT_CAPTION, ByVal 0&

'End If
    'mcap = True
'Else
'    ReleaseCapture
'    'DoEvents
'    'FormDrag = True
'
'    SendMessage Me.hWnd, WM_NCLBUTTONUP, LP_HT_CAPTION, ByVal 0&

''ReleaseCapture
''SendMessage Me.hWnd, WM_NCLBUTTONUP, LP_HT_CAPTION, ByVal 0&
''SendMessage Me.hWnd, WM_NCLBUTTONUP, 0&, ByVal 0&
'    'If mcap Then
'    'SetCapture (Me.hWnd): mcap = False
'    'ReleaseCapture
'    'SetCapture (Me.hWnd)
'    'GetCapture
'



MDown = False
TWait.Enabled = True

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If mcap = True Then Exit Sub
'только на быстрый клик
'If FormDrag = False Then
If Button = 1 Then Unload Me
'FormDrag = False
End Sub

Private Sub Form_Paint()
'scrsaver
XStep = Screen.TwipsPerPixelX * SgnRnd
YStep = Screen.TwipsPerPixelY * SgnRnd

TicScrSaver = GetTickCount
TMove.Enabled = False
End Sub

Private Sub Form_Resize()
Dim hsbH As Long

On Error Resume Next
'было 21

hsbH = hsbHeight + 3

'не показывать скролл, если маленькая картинка
If (PicHB.Height * 3) > Me.ScaleHeight Then
    PicHB.Visible = False
    Exit Sub
End If

'сжимать скролл, если не влезает
'PicHB.Move Me.ScaleWidth - 97, Me.ScaleHeight - 20, 97, 20 '20 мало - не видно бывает скролла
PicHB.Move Me.ScaleWidth - 97, Me.ScaleHeight - hsbH, 97, hsbH
'PicHB.Move 0, Me.ScaleHeight - 20, 97, 20
If PicHB.Width >= Me.ScaleWidth Then
    PicHB.Move 0, Me.ScaleHeight - hsbH, Me.ScaleWidth, hsbH
End If
'Pic.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)
If addflag Or editFlag Then
 FrmMain.Timer2.Enabled = False
Else
 'FrmMain.Timer2.Enabled = True
End If

FormShowPicLoaded = False
MDown = False
Set hb_cScroll = Nothing
FormShowPicIsModal = False
End Sub


Private Sub hb_cscroll_ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)
ShowPicFocus = True
Me.SetFocus
End Sub
Private Sub hb_cscroll_Scroll(eBar As EFSScrollBarConstants)
hb_cScrollValue = hb_cScroll.Value(efsHorizontal)
End Sub
Private Sub hb_cscroll_Change(eBar As EFSScrollBarConstants)
Dim tRes As String    'разрешение
Dim tOld As Boolean

hb_cScrollValue = hb_cScroll.Value(efsHorizontal)

'1  If hb_cScrollValue <> HScroll.Value Then hb_cScroll.Value(efsHorizontal) = HScroll.Value 'позиционирование
2 If addflag Or editFlag Or IsCoverShowFlag Then PicHB.Visible = False: Exit Sub
3 If PicManualFlag Then Exit Sub

'Debug.Print Time

With FrmMain
    tOld = .Timer2.Enabled
    .Timer2.Enabled = False

    Select Case hb_cScrollValue
    Case 1
        If NoPic1Flag Then
            If NoPic2Flag Then
                GetPic .PicTempHid(0), 1, "SnapShot3"
            Else
                GetPic .PicTempHid(0), 1, "SnapShot2"
            End If
        Else
            GetPic .PicTempHid(0), 1, "SnapShot1"
        End If

    Case 2
        If NoPic1Flag And NoPic2Flag Then
            GetPic .PicTempHid(0), 1, "SnapShot3"
        Else
            If NoPic1Flag Then
                GetPic .PicTempHid(0), 1, "SnapShot3"
            ElseIf NoPic2Flag Then
                GetPic .PicTempHid(0), 1, "SnapShot3"
            Else
                GetPic .PicTempHid(0), 1, "SnapShot2"
            End If
        End If

    Case 3
        GetPic .PicTempHid(0), 1, "SnapShot3"
    End Select


    'то же в ShowInShowPic
    SetFrmShowPicPicture 0


    'разрешение
    'FormShowPic.Print ScaleX(.PicTempHid(0).Width, vbTwips, vbPixels) & "x" & _
     'ScaleY(.PicTempHid(0).Height, vbTwips, vbPixels)

    'не показывать скролл и разрешение, если маленькая картинка
    If (PicHB.Height * 3) > Me.ScaleHeight Then
        PicHB.Visible = False
    Else
        'увеличенное разрешение
        PicHB.Visible = True

        '        FormShowPic.Print ScaleX(FormShowPic.Width, vbTwips, vbPixels) & "x" & _
                 ScaleY(FormShowPic.Height, vbTwips, vbPixels)
        tRes = FormShowPic.Width / Screen.TwipsPerPixelX & "x" & FormShowPic.Height / Screen.TwipsPerPixelY
        FormShowPic.Line (3, 3)-(FormShowPic.TextWidth(tRes) + 2, FormShowPic.TextHeight(tRes) + 2), 0, BF
        FormShowPic.CurrentX = 3: FormShowPic.CurrentY = 3
        FormShowPic.Print tRes

    End If

    .Timer2.Enabled = tOld
    '.Timer2_Timer

End With

End Sub
Private Sub hb_cscroll_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
hb_cScrollValue = hb_cScroll.Value(efsHorizontal)
End Sub


Private Sub TMove_Timer()
'If Not FormShowPicLoaded Then
If GetForegroundWindow <> Me.hwnd Then
TMove.Enabled = False
'Debug.Print Time & " exit" '& ActiveControl.name
'LockWindowUpdate 0
Exit Sub
End If
'End If
'LockWindowUpdate FrmMain.ListView.hwnd

'Debug.Print Time

FrmMain.Timer2.Enabled = False
If MDown Then Exit Sub

'If Me.Left < 500 Then TMove.Enabled = False: Exit Sub
'If Me.Top < 500 Then TMove.Enabled = False: Exit Sub
'If Me.Left - 500 > Screen.Width - Me.Width Then TMove.Enabled = False: Exit Sub
'If Me.Top - 500 > Screen.Height - Me.Height Then TMove.Enabled = False: Exit Sub



'If Me.Left < 0 Or Me.Left > Screen.Width - Me.Width Then
'XStep = -XStep
'Debug.Print "смена X знака"
'End If
'If Me.Top < 0 Or Me.Top > Screen.Height - Me.Height Then
'YStep = -YStep
'Debug.Print "смена Y знака"
'End If
    
If Me.left < 500 Or Me.left > Screen.Width - Me.Width - 500 Then XStep = -XStep
If Me.top < 500 Or Me.top > Screen.Height - Me.Height - 1000 Then YStep = -YStep
     
     
'Me.Move Me.Left + XStep, Me.Top + YStep
Me.left = Me.left + XStep
Me.top = Me.top + YStep

'Debug.Print "move: " & Me.Left + XStep, Me.Top + YStep

End Sub

Private Function SgnRnd() As Integer
Randomize Timer
If Int(2 * Rnd) = 0 Then SgnRnd = 1 Else SgnRnd = -1
End Function

Private Sub TWait_Timer()
If GetTickCount - TicScrSaver > TWait.Interval Then

If Me.left < 500 Then TMove.Enabled = False: TWait.Enabled = False: Exit Sub
If Me.top < 500 Then TMove.Enabled = False: TWait.Enabled = False: Exit Sub
If Me.left - 500 > Screen.Width - Me.Width Then TMove.Enabled = False: TWait.Enabled = False: Exit Sub
If Me.top - 500 > Screen.Height - Me.Height Then TMove.Enabled = False: TWait.Enabled = False: Exit Sub

TMove.Enabled = True

End If
End Sub
