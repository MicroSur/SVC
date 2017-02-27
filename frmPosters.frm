VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPosters 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   Icon            =   "frmPosters.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   427
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      Height          =   270
      Left            =   480
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   900
      Begin VB.PictureBox picProgressSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   360
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   94
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      Height          =   2880
      Left            =   120
      ScaleHeight     =   188
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   5
      Top             =   120
      Width           =   3525
      Begin VB.PictureBox picSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2475
         Left            =   120
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   2610
         Begin VB.OptionButton optThumb 
            Height          =   1920
            Index           =   0
            Left            =   210
            MaskColor       =   &H8000000F&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   1920
         End
      End
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   120
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6090
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   556
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   423
            MinWidth        =   423
            Picture         =   "frmPosters.frx":000C
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPosters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vsbWidth As Long    'размер скролла в системе
Private WithEvents poster_cScroll As cScrollBars
Attribute poster_cScroll.VB_VarHelpID = -1

Private Type POINTAPI
    X As Long
    Y As Long
End Type

'Private mbActive As Boolean
Private mlCurThumb As Long
Private Const SRCCOPY As Long = &HCC0020
Private Const STRETCH_HALFTONE As Long = &H4&
'Private Const SW_RESTORE As Long = &H9&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub optThumb_Click(Index As Integer)
DoEvents
Call SetPoster(Index)
End Sub

Private Sub poster_cScroll_Change(eBar As EFSScrollBarConstants)
poster_cScroll_Scroll eBar
End Sub
Private Sub poster_cScroll_Scroll(eBar As EFSScrollBarConstants)
If (eBar = efsHorizontal) Then
    picSlide.left = -Screen.TwipsPerPixelX * poster_cScroll.Value(eBar)
Else
    picSlide.top = -Screen.TwipsPerPixelY * poster_cScroll.Value(eBar)
End If
End Sub

Private Sub CreateThumbPic(picSource As PictureBox, picThumb As PictureBox)

'This sub uses the halftone stretch mode, which produces the highest
'quality possible, when stretching the bitmap.

Dim lRet As Long
Dim lLeft As Long
Dim lTop As Long
Dim lWidth As Long
Dim lHeight As Long
Dim lForeColor As Long
Dim hBrush As Long
Dim hDummyBrush As Long
Dim lOrigMode As Long
Dim fScale As Single
Dim uBrushOrigPt As POINTAPI

picThumb.Width = optThumb(0).Width - 10    '80 '64
picThumb.Height = optThumb(0).Height - 24    ' 80 '64
'picThumb.BackColor = vbButtonFace

picThumb.AutoRedraw = True
picThumb.Cls

If picSource.Width <= picThumb.Width - 2 And picSource.Height <= picThumb.Height - 2 Then
    fScale = 1
Else
    'без -2 нек картинки неверные
    fScale = IIf(picSource.Width > picSource.Height, (picThumb.Width - 2) / picSource.Width, (picThumb.Height - 2) / picSource.Height)
End If

lWidth = picSource.Width * fScale
lHeight = picSource.Height * fScale
lLeft = Int((picThumb.Width - lWidth) / 2)
lTop = Int((picThumb.Height - lHeight) / 2)

'Store the original ForeColor
lForeColor = picThumb.ForeColor

'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
lOrigMode = SetStretchBltMode(picThumb.hdc, STRETCH_HALFTONE)

'Realign the brush...
'Get picEdit's brush by selecting a dummy brush into the DC
hDummyBrush = CreateSolidBrush(lForeColor)
hBrush = SelectObject(picThumb.hdc, hDummyBrush)
'Reset the brush (This will force windows to realign it when it's put back)
lRet = UnrealizeObject(hBrush)
'Set picEdit's brush alignment coordinates to the left-top of the bitmap
lRet = SetBrushOrgEx(picThumb.hdc, lLeft, lTop, uBrushOrigPt)
'Now put the original brush back into the DC at the new alignment
hDummyBrush = SelectObject(picThumb.hdc, hBrush)

'Stretch the bitmap
lRet = StretchBlt(picThumb.hdc, lLeft, lTop, lWidth, lHeight, _
                  picSource.hdc, 0, 0, picSource.Width, picSource.Height, SRCCOPY)

'Set the stretch mode back to it's original mode
lRet = SetStretchBltMode(picThumb.hdc, lOrigMode)

'Reset the original alignment of the brush...
'Get picEdit's brush by selecting the dummy brush back into the DC
hBrush = SelectObject(picThumb.hdc, hDummyBrush)
'Reset the brush (This will force windows to realign it when it's put back)
lRet = UnrealizeObject(hBrush)
'Set the brush alignment back to the original coordinates
lRet = SetBrushOrgEx(picThumb.hdc, uBrushOrigPt.X, uBrushOrigPt.Y, uBrushOrigPt)
'Now put the original brush back into picEdit's DC at the original alignment
hDummyBrush = SelectObject(picThumb.hdc, hBrush)
'Get rid of the dummy brush
lRet = DeleteObject(hDummyBrush)

'Restore the original ForeColor
picThumb.ForeColor = lForeColor

'линии вокруг картинок
'picThumb.Line (lLeft - 1, lTop - 1)-Step(lWidth + 1, lHeight + 1), &H0&, B

End Sub

Public Sub CreateThumbs(lFilCnt As Long)

Dim iMaxLen As Integer
'Dim X As Long
'Dim Y As Long
Dim i As Long
Dim lPicCnt As Long
'Dim lFilCnt As Long
'Dim spath As String
Dim sText As String

Screen.MousePointer = vbHourglass

picSlide.Move 0, 0, optThumb(0).Width, optThumb(0).Height
picSlide.Visible = False
picSlide.BackColor = vbButtonFace
Set picSlide.Font = optThumb(0).Font
While optThumb.Count > 1
    Unload optThumb(optThumb.Count - 1)
Wend


DoEvents
On Error Resume Next


Call StartProgress

For i = 0 To lFilCnt    'filHidden.ListCount - 1

    If GetKeyState(vbKeyEscape) < 0 Then Exit For

    Call UpdateProgress((CSng(i + 1) / CSng(lFilCnt)) * 100)    ', filHidden.List(i))
    Set picLoad.Picture = Nothing    'LoadPicture()
    picLoad.Cls
    err.Clear

    'If InStr(1, LCase$(filHidden.List(i)), ".ico") > 0 _
     '   Or InStr(1, LCase$(filHidden.List(i)), ".cur") > 0 Then
    '    Set picLoad.Picture = LoadPicture(sPath & filHidden.List(i), vbLPLargeShell, vbLPDefault)
    'Else
    '        Set picLoad.Picture = LoadPicture(sPath & filHidden.List(i))

    GetURL2Pic SC.CodeObject.sPoster(i), picLoad
    sText = SC.CodeObject.tPoster(i)

    'End If
    If err.Number = 0 Then
        Call CreateThumbPic(picLoad, picThumb)
        If lPicCnt > 0 Then
            Load optThumb(lPicCnt)
            Set optThumb(lPicCnt).Container = picSlide
        End If
        'optThumb(lPicCnt).Tag = filHidden.List(i)
        Set optThumb(lPicCnt).Picture = picThumb.Image
        optThumb(lPicCnt).Value = False

        iMaxLen = optThumb(lPicCnt).Width - 15
        '            If picSlide.TextWidth(sText) > iMaxLen Then
        '                iMaxLen = iMaxLen - picSlide.TextWidth("...")
        '            End If
        While picSlide.TextWidth(sText) > iMaxLen
            sText = left$(sText, Len(sText) - 1)
        Wend
        '            If iMaxLen < optThumb(lPicCnt).Width - 15 Then
        '                sText = sText & "..."
        '            End If

        optThumb(lPicCnt).Caption = sText
        optThumb(lPicCnt).Visible = True
        lPicCnt = lPicCnt + 1
    End If
Next i

picProgress.Visible = False

'Free the unneeded resources
Set picLoad.Picture = LoadPicture()
Set picThumb.Picture = LoadPicture()
'optThumb(0).Value = True
mlCurThumb = 0

Call Form_Resize
picSlide.Visible = True
''''''''''''''''

Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long

If poster_cScroll Is Nothing Then Exit Sub
poster_cScroll.Value(efsVertical) = 0

lHeight = (picSlide.Height - picFrame.Height + sbrMain.Height) \ Screen.TwipsPerPixelY
If (lHeight > 0) Then
    lProportion = lHeight \ (picFrame.Height \ Screen.TwipsPerPixelY) + 1
    poster_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
    'poster_cScroll.SmallChange(efsVertical) = 1
    poster_cScroll.Max(efsVertical) = lHeight
    poster_cScroll.Visible(efsVertical) = True
Else
    poster_cScroll.Visible(efsVertical) = False
    picSlide.top = 0
    'Me.Width = optThumb(0).Width * Screen.TwipsPerPixelX * 3 + 180
End If
''''''''''''''
Call Form_Resize
'End If

Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
frmPostersFlag = True 'до слайдера
Set poster_cScroll = New cScrollBars
poster_cScroll.Create picFrame.hwnd, False, True
'poster_cScroll.style = efsRegular '= efsEncarta 'не важно, системное
'poster_cScroll.VBGColor = RGB(5, 255, 255)
poster_cScroll.Orientation = efsoVertical
poster_cScroll.Visible(efsVertical) = False

ScaleMode = 3
'Const SM_CXVSCROLL = 2
'Const SM_CYHSCROLL = 3
vsbWidth = GetSystemMetrics(2)

Me.Caption = "SurVideoCatalog - " & NamesStore(13)
Me.Icon = FrmMain.Icon

Me.Width = optThumb(0).Width * Screen.TwipsPerPixelX * 3 + 180 + vsbWidth * Screen.TwipsPerPixelX
Me.Height = optThumb(0).Height * Screen.TwipsPerPixelY * 3 + sbrMain.Height * Screen.TwipsPerPixelY + 600 + vsbWidth * Screen.TwipsPerPixelY

Me.top = frmEditor.top + (frmEditor.Height - Me.Height) / 2
Me.left = frmEditor.left + (frmEditor.Width - Me.Width) / 2


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim lIdx As Long
For lIdx = 1 To optThumb.Count - 1
    Unload optThumb(lIdx)
Next lIdx

Set poster_cScroll = Nothing
frmPostersFlag = False
End Sub

Private Sub Form_Resize()

Dim X As Long
Dim Y As Long
Dim lIdx As Long
Dim lCols As Long

On Error Resume Next

If Me.WindowState <> vbMinimized Then
    'Me.Width = optThumb(0).Width * Screen.TwipsPerPixelX * 3 + 180 + vsbWidth * Screen.TwipsPerPixelX
    'Me.Height = optThumb(0).Height * Screen.TwipsPerPixelY * 3 + 600 + vsbWidth * Screen.TwipsPerPixelY


    picFrame.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - sbrMain.Height
    'lCols = Int((picFrame.ScaleWidth - vsbWidth) / optThumb(0).Width)
    lCols = Int(picFrame.ScaleWidth / optThumb(0).Width)
    If lCols > 0 Then
        For lIdx = 0 To optThumb.Count - 1
            X = (lIdx Mod lCols) * optThumb(0).Width
            Y = Int(lIdx / lCols) * optThumb(0).Height
            optThumb(lIdx).Move X, Y
        Next lIdx
        picSlide.Width = lCols * optThumb(0).Width
        picSlide.Height = Int(optThumb.Count / lCols) * optThumb(0).Height
        If Int(optThumb.Count / lCols) < (optThumb.Count / lCols) Then
            picSlide.Height = picSlide.Height + optThumb(0).Height
        End If
    End If


    'Background 'после размеров
    If lngBrush <> 0 Then
        GetClientRect picFrame.hwnd, rctMain
        FillRect picFrame.hdc, rctMain, lngBrush
        GetClientRect picSlide.hwnd, rctMain
        FillRect picSlide.hdc, rctMain, lngBrush
    End If

End If


End Sub
Private Sub StartProgress()

With picProgress
    .Cls
    .BackColor = vbButtonFace
    .ForeColor = vbButtonText
    .Move sbrMain.left + sbrMain.Panels(2).left, sbrMain.top + 1, _
          sbrMain.Panels(2).Width, sbrMain.Height - 1
End With

With picProgressSlide
    .Cls
    .BackColor = vbHighlight
    .ForeColor = vbHighlightText
    .Move 0, 0, 1, picProgress.ScaleHeight
End With

picProgress.Visible = True

End Sub

Private Sub UpdateProgress(ByVal iPercent As Integer)    ', ByVal sCaption As String)

Dim lTextTop As Long

picProgress.Cls
picProgressSlide.Cls
picProgressSlide.Width = picProgress.ScaleWidth * (CSng(iPercent) / 100!)
lTextTop = (picProgress.ScaleHeight - picProgress.TextHeight(NamesStore(15))) / 2
picProgress.CurrentX = 3
picProgress.CurrentY = lTextTop
picProgress.Print NamesStore(15)
picProgressSlide.CurrentX = 3
picProgressSlide.CurrentY = lTextTop
picProgressSlide.Print NamesStore(15)

DoEvents

End Sub
