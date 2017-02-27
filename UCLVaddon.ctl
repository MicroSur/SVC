VERSION 5.00
Begin VB.UserControl UCLVaddon 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   6240
   ScaleWidth      =   6915
   Begin VB.TextBox tBIO 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   5040
      Width           =   3075
   End
   Begin VB.PictureBox picUCLV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   1140
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   32
      Top             =   4980
      Width           =   975
   End
   Begin VB.PictureBox PicLVAddon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5775
      ScaleWidth      =   6675
      TabIndex        =   15
      Top             =   0
      Width           =   6675
      Begin VB.TextBox textFile 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4800
         Width           =   5295
      End
      Begin VB.TextBox textDebtor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   5400
         Width           =   5295
      End
      Begin VB.TextBox textLabel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3600
         Width           =   5295
      End
      Begin VB.TextBox textGenre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox textMName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   60
         Width           =   5295
      End
      Begin VB.TextBox textCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   660
         Width           =   5295
      End
      Begin VB.TextBox textAuthor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5295
      End
      Begin VB.TextBox textRole 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   1260
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox textOther 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5100
         Width           =   5295
      End
      Begin VB.TextBox textCDN 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3300
         Width           =   5295
      End
      Begin VB.TextBox textTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   5295
      End
      Begin VB.TextBox textVideo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4200
         Width           =   5295
      End
      Begin VB.TextBox textAudio 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4500
         Width           =   5295
      End
      Begin VB.TextBox TextLang 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2700
         Width           =   5295
      End
      Begin VB.TextBox TextSubt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3000
         Width           =   5295
      End
      Begin VB.Image ImStars 
         Appearance      =   0  'Flat
         Height          =   180
         Left            =   1260
         Picture         =   "UCLVaddon.ctx":0000
         Top             =   3940
         Width           =   1515
      End
      Begin VB.Image ImStars0 
         Appearance      =   0  'Flat
         Height          =   180
         Left            =   1800
         Picture         =   "UCLVaddon.ctx":010D
         Top             =   3960
         Width           =   1515
      End
      Begin VB.Image imgType 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   120
         Top             =   2100
         Width           =   495
      End
      Begin VB.Label LFile 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "File(s)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   4800
         Width           =   405
      End
      Begin VB.Label LDebtor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Debtor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   5400
         Width           =   480
      End
      Begin VB.Label LLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   29
         Top             =   3600
         Width           =   390
      End
      Begin VB.Label LOther 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   5100
         Width           =   735
      End
      Begin VB.Label LAct 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Actors"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label LCountry 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Production"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   660
         Width           =   765
      End
      Begin VB.Label LRes 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Director"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label LMName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   300
      End
      Begin VB.Label LGenre 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Genre"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   23
         Top             =   360
         Width           =   435
      End
      Begin VB.Label LVideo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Video"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   22
         Top             =   4200
         Width           =   405
      End
      Begin VB.Label LAudio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Audio"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   4500
         Width           =   405
      End
      Begin VB.Label LTime 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   960
         Width           =   345
      End
      Begin VB.Label LNcd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CDN"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   3300
         Width           =   345
      End
      Begin VB.Label LLang 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lang"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   2700
         Width           =   360
      End
      Begin VB.Label LSubt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subt"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   3000
         Width           =   330
      End
      Begin VB.Label LRate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rating"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   3900
         Width           =   465
      End
   End
End
Attribute VB_Name = "UCLVaddon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''''''''''
'   при изменении полей
'Clear()
'UserControl_Resize()
'frmmain.LangChange()
'frmmain.FillLVAdd()
'Клик на label
''''''''''
Private WithEvents lvaddon_cScroll As cScrollBars
Attribute lvaddon_cScroll.VB_VarHelpID = -1

Event tActMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, s As String)
Event tActMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, s As String)
'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
'Private UCLastCtrl As String 'последний контрол текст с фокусом и выделением

Private FixTxtLeft As Long


'Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
'Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
'Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
'Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
'Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls

Public Property Get hwnd()
    hwnd = PicLVAddon.hwnd
End Property

Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property



Private Sub LAct_Click()
GotoCol dbActerInd + 1
End Sub

Private Sub LAudio_Click()
GotoCol dbAudioInd + 1
End Sub

Private Sub LCountry_Click()
GotoCol dbCountryInd + 1
End Sub

Private Sub LDebtor_Click()
GotoCol dbDebtorInd + 1
End Sub

Private Sub LFile_Click()
GotoCol dbFileNameInd + 1
End Sub

Private Sub LGenre_Click()
GotoCol dbGenreInd + 1
End Sub

Private Sub GotoCol(ind As Integer)
'принимаем инд требуемой колонки
Dim cleft As Single
'Dim pos As Long
Dim i As Integer

On Error Resume Next

With FrmMain
'заполнить массис соответствия   индекс в списке - позиция
'индекс массива - позиция, значение - индекс поля списка
Dim pArr() As Integer
ReDim pArr(.ListView.ColumnHeaders.Count) 'As Integer
For i = 1 To .ListView.ColumnHeaders.Count
    pArr(.ListView.ColumnHeaders(i).Position) = i
Next i

nScrollPos = GetScrollPos(.ListView.hwnd, SB_HORZ)
'pos = .ListView.ColumnHeaders(ind).Position
'Debug.Print "Width = " & .ListView.ColumnHeaders(ind).Width
'Debug.Print .ListView.ColumnHeaders(ind).Text & " инд. = " & ind & " поз. = " & pos
'cleft = .ListView.ColumnHeaders(.ListView.ColumnHeaders(ind).Position).Left / Screen.TwipsPerPixelX
'cleft = .ListView.ColumnHeaders(ind).Left / Screen.TwipsPerPixelX
'cleft = ScaleX(.ListView.ColumnHeaders(ind).Left, vbTwips, vbPixels) '/ Screen.TwipsPerPixelX

For i = 1 To .ListView.ColumnHeaders(ind).Position - 1
    cleft = cleft + .ListView.ColumnHeaders(pArr(i)).Width
Next i
    cleft = cleft / Screen.TwipsPerPixelX

'Debug.Print "где мы = " & nScrollPos
'Debug.Print "куда = " & cleft

ListViewScroll .ListView, cleft - nScrollPos, 0

End With
End Sub

Private Sub LLabel_Click()
GotoCol dbLabelInd + 1
End Sub

Private Sub LLang_Click()
GotoCol dbLanguageInd + 1
End Sub

Private Sub LMName_Click()
GotoCol dbMovieNameInd + 1
End Sub

Private Sub LNcd_Click()
GotoCol dbMediaTypeInd + 1
End Sub

Private Sub LOther_Click()
GotoCol dbOtherInd + 1
End Sub

Private Sub LRate_Click()
GotoCol dbRatingInd + 1
End Sub

Private Sub LRes_Click()
GotoCol dbDirectorInd + 1
End Sub

Private Sub LSubt_Click()
GotoCol dbSubTitleInd + 1
End Sub

Private Sub LTime_Click()
GotoCol dbTimeInd + 1
End Sub

Private Sub LVideo_Click()
GotoCol dbVideoInd + 1
End Sub

Private Sub PicLVAddon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ctrl As Control
'поймать фокус
On Error Resume Next
'If Not FormShowPicLoaded Then
If frmSRFlag Then
    If frmSR.Visible Then Exit Sub
End If

If FrmMain.txtEdit.Visible Then Exit Sub
If FrmMain.LstFiles.Visible Then Exit Sub
If FrmMain.TextItemHid.SelLength > 0 Then Exit Sub

If GetForegroundWindow = FrmMain.hwnd Then

    If ActiveControl.name <> "UCLV" Then

        For Each ctrl In UserControl.Controls
            If TypeOf ctrl Is TextBox Then
                If ActiveControl.name = ctrl.name Then
                    ''    Debug.Print ctrl.name
                    ''        Exit Sub
                    ''    End If
                    'Debug.Print "c:" & ctrl.name
                    If ctrl.SelLength <> 0 Then
                        ctrl.SetFocus
                        Exit Sub
                    End If
                Else
                    'UCLastCtrl = ctrl.name
                    ctrl.SelLength = 0    'убрать выделения не активных
                End If
            End If
        Next

        UserControl.SetFocus
    End If

End If
End Sub

Private Sub picUCLV_Click()
Dim tOld As Boolean
On Error Resume Next

With FrmMain
    If UCLVShowPersonFlag Then
        If .mnuShowThisActer.Enabled Then .mnuShowThisActer_Click
    Else
    
        If NoDBFlag Then Exit Sub
        If NoPicFrontFaceFlag Then Exit Sub
        If .PicFaceV.Picture = 0 Then GetPic .PicFaceV, 1, "FrontFace" '22
        If NoPicFrontFaceFlag Then Exit Sub
        
        
        'If Button = 2 Then Me.PopupMenu Me.popFaceHid: Exit Sub

        tOld = .Timer2.Enabled
        .Timer2.Enabled = False

        'If Button = 1 Then
        .PicTempHid(1).Picture = .PicFaceV.Image
        IsCoverShowFlag = True
        ViewScrShotFlag = False
        FormShowPic.Visible = False

        'не видим скролл
        'FormShowPic.hb_cScroll.Visible(efsHorizontal) = False
        FormShowPic.PicHB.Visible = False
        PicManualFlag = True
        ShowInShowPic 1, FrmMain
        'End If    'but 1

        .Timer2.Enabled = tOld
    End If

End With
End Sub

Private Sub TextAuthor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   ' Make VB discard the mouse capture.
    TextAuthor.Enabled = False
    TextAuthor.Enabled = True
    TextAuthor.SetFocus
    RaiseEvent tActMouseDown(Button, Shift, X, Y, TextAuthor.SelText)
Else
'ctrl выбор имени актера (не шифт)
    'If Shift = 2 Then SelectWordsGroup TextAuthor, TextAuthor.SelStart
    If Shift = 0 Then SelectWordsGroup TextAuthor, TextAuthor.SelStart
End If
End Sub

Private Sub TextAuthor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent tActMouseUp(Button, Shift, X, Y, TextAuthor.SelText)
End Sub

Private Sub TextRole_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   ' Make VB discard the mouse capture.
    TextRole.Enabled = False
    TextRole.Enabled = True
    TextRole.SetFocus
    RaiseEvent tActMouseDown(Button, Shift, X, Y, TextRole.SelText)
    
Else
'ctrl выбор имени актера
    'If Shift = 2 Then SelectWordsGroup TextRole, TextRole.SelStart
    If Shift = 0 Then SelectWordsGroup TextRole, TextRole.SelStart

End If


End Sub

Private Sub TextRole_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent tActMouseUp(Button, Shift, X, Y, TextRole.SelText)
End Sub
Private Function Max(in1 As Integer, in2 As Integer) As Integer
If in1 >= in2 Then Max = in1 Else Max = in2
End Function


Private Sub UserControl_Initialize()
Set lvaddon_cScroll = New cScrollBars
lvaddon_cScroll.Create UserControl.hwnd, False, True
'lvaddon_cScroll.style = efsRegular '= efsEncarta 'не важно, системное
'lvaddon_cScroll.VBGColor = RGB(5, 255, 255)
lvaddon_cScroll.Orientation = efsoVertical
PicLVAddon.Move 0, 0
picUCLV.Height = 2650
End Sub


Private Sub UserControl_Resize()
'If FirstActivateFlag Then Exit Sub

Dim txtL As Integer
Dim txtW As Integer
'mzt Dim txtH As Integer 'в ролях, прим
Dim lblL As Integer
Dim lblW As Integer
Dim ctrl As Control

For Each ctrl In UserControl.Controls
    'установка высоты текстов
    If TypeOf ctrl Is TextBox Then
        If (ctrl.name <> "textRole") And (ctrl.name <> "textOther") Then ctrl.Height = LMName.Height
    End If
Next



On Error Resume Next

lblL = 60

lblW = 1150

'txtL = lblL + lblW + 60
txtL = lblL + LSubt.Width + 100    '60
txtW = PicLVAddon.Width - lblW - 200

FixTxtLeft = txtL    'TextMName.Left
'If LMName.Font.Size > 10 Then FixTxtLeft = txtL + 150
'If LMName.Font.Size > 12 Then FixTxtLeft = txtL + 250
'If LMName.Font.Size > 14 Then FixTxtLeft = txtL + 500

TextMName.Move FixTxtLeft, TextMName.Top, txtW

'LGenre 'TextGenre
TextGenre.Move FixTxtLeft, TextGenre.Top, txtW

'LCountry 'TextCountry
TextCountry.Move FixTxtLeft, TextCountry.Top, txtW

textTime.Move FixTxtLeft, textTime.Top, txtW

TextAuthor.Move FixTxtLeft, TextAuthor.Top, txtW
'AntiFlick.Width = textAuthor.Width

TextRole.Move FixTxtLeft, TextRole.Top, txtW
TextLang.Move FixTxtLeft, TextLang.Top, txtW
TextSubt.Move FixTxtLeft, TextSubt.Top, txtW
TextCDN.Move FixTxtLeft, TextCDN.Top, txtW
'ImStars.Move ImStars.Left, ImStars.Top ',  txtW
ImStars.Move FixTxtLeft, ImStars.Top
ImStars0.Move FixTxtLeft, ImStars.Top
'imgType.Move txtW, ImStars.Top

textVideo.Move FixTxtLeft, textVideo.Top, txtW
textAudio.Move FixTxtLeft, textAudio.Top, txtW
textFile.Move FixTxtLeft, textFile.Top, txtW
TextOther.Move FixTxtLeft, TextOther.Top, txtW
TextLabel.Move FixTxtLeft, TextLabel.Top, txtW
textDebtor.Move FixTxtLeft, textDebtor.Top, txtW


'''''''''''''''''''
Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long

On Error Resume Next

PicLVAddon.Width = UserControl.ScaleWidth

'uc будет до FirstActivateFlag флага
picUCLV.Move lblL, PicLVAddon.Top + PicLVAddon.Height ', picUCLV.Height  ', PicLVAddon.Width - FixTxtLeft

If picUCLV.Top + picUCLV.Height > UserControl.Height Then
    Opt_UCLVPic_Vis = False
    picUCLV.Visible = False
    tBIO.Visible = False: tBIO.Text = vbNullString
    
Else
    If Not Opt_UCLVPic_Vis Then 'только если не было видно
    If FrmMain.ListView.ListItems.Count > 0 Then
    If Not FirstActivateFlag Then 'без этого и скриншоты маленькими показ. при старте
        If GetPic(FrmMain.PicTempHid(1), 1, "FrontFace") Then
            'ResizeWIA FrmMain.PicTempHid(1), picUCLV.ScaleWidth, picUCLV.ScaleHeight, aratio:=True
            ResizeWIA FrmMain.PicTempHid(1), picUCLV.ScaleHeight, picUCLV.ScaleHeight, aratio:=True
            picUCLV.Width = FrmMain.PicTempHid(1).Width
            picUCLV.Picture = FrmMain.PicTempHid(1).Picture
        Else
            picUCLV.Picture = Nothing
        End If
    End If
    End If
    End If
    Opt_UCLVPic_Vis = True    'каринка видна
    picUCLV.Visible = True
    If UCLVShowPersonFlag Then tBIO.Visible = True
End If

tBIO.Left = picUCLV.Left + picUCLV.Width + 90
tBIO.Move tBIO.Left, picUCLV.Top, PicLVAddon.Width - tBIO.Left - 60, picUCLV.Height


If lvaddon_cScroll Is Nothing Then Exit Sub

lHeight = (PicLVAddon.Height - UserControl.ScaleHeight) \ Screen.TwipsPerPixelY
If (lHeight > 0) Then
    lProportion = lHeight \ (UserControl.ScaleHeight \ Screen.TwipsPerPixelY) + 1
    lvaddon_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
    lvaddon_cScroll.Max(efsVertical) = lHeight
    lvaddon_cScroll.Visible(efsVertical) = True
Else
    lvaddon_cScroll.Visible(efsVertical) = False
    PicLVAddon.Top = 0
End If


End Sub
Public Sub Correct_tBIO()
If picUCLV.Picture = 0 Then
tBIO.Left = FixTxtLeft 'picUCLV.Left
Else
tBIO.Left = picUCLV.Left + picUCLV.Width + 90
End If
tBIO.Move tBIO.Left, picUCLV.Top, PicLVAddon.Width - tBIO.Left - 60, picUCLV.Height
End Sub

Private Sub lvaddon_cScroll_Change(eBar As EFSScrollBarConstants)
   lvaddon_cScroll_Scroll eBar
End Sub
Private Sub lvaddon_cScroll_Scroll(eBar As EFSScrollBarConstants)
   If (eBar = efsHorizontal) Then
      PicLVAddon.Left = -Screen.TwipsPerPixelX * lvaddon_cScroll.Value(eBar)
   Else
      PicLVAddon.Top = -Screen.TwipsPerPixelY * lvaddon_cScroll.Value(eBar)
   End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Dim ctrl As Control
    UserControl.BackColor() = New_BackColor

For Each ctrl In UserControl.Controls
If TypeOf ctrl Is TextBox Or TypeOf ctrl Is Label Or TypeOf ctrl Is PictureBox Then
    ctrl.BackColor = New_BackColor
End If
Next
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Dim ctrl As Control

    UserControl.ForeColor() = New_ForeColor
    
For Each ctrl In UserControl.Controls
If TypeOf ctrl Is TextBox Or TypeOf ctrl Is Label Then
    ctrl.ForeColor = New_ForeColor
End If
If TypeOf ctrl Is Line Then
ctrl.BorderColor = New_ForeColor
End If
Next
   
    PropertyChanged "ForeColor"
End Property

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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Dim ctrl As Control
    Set UserControl.Font = New_Font
For Each ctrl In UserControl.Controls
If (TypeOf ctrl Is TextBox) Or (TypeOf ctrl Is Label) Then
'ctrl.Font.Bold = False
    Set ctrl.Font = New_Font

'ctrl.Font.Size = 8
End If
Next
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl_Resize
End Sub

'Private Sub UserControl_Click()
'    RaiseEvent Click
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
  
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HC0FFFF)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Terminate()
Set lvaddon_cScroll = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HC0FFFF)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub



Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property
Public Sub ShowStars(v As Single)
ImStars.ToolTipText = " " & v & " "
If v = 0 Or v > 10 Then
ImStars0.Visible = False: ImStars.Visible = False
Exit Sub
End If

If kDPI <= 0 Then kDPI = 1
'UCLV.Controls("ImStars").Width = (v * 165) - (15 * v) + 15
ImStars.Width = (v * 150 + 15) / kDPI ' v * 150 / kDPI + 15 / kDPI
ImStars.Visible = True: ImStars0.Visible = True
End Sub
Public Sub Clear()
picUCLV.Picture = Nothing
tBIO = vbNullString
TextMName = vbNullString
'TextUser = vbNullString
textAudio = vbNullString
textVideo = vbNullString
textTime = vbNullString
TextCDN = vbNullString
TextOther = vbNullString
TextRole = vbNullString
TextAuthor = vbNullString
TextCountry = vbNullString
TextGenre = vbNullString

TextLang = vbNullString
TextSubt = vbNullString
ImStars.Visible = False: ImStars0.Visible = False
Set imgType.Picture = Nothing

TextLabel = vbNullString
textDebtor = vbNullString
textFile = vbNullString
End Sub
