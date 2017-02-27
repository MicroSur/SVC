VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAuto 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chAutoFiles 
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   37
      Top             =   4920
      Width           =   195
   End
   Begin VB.CheckBox chAutoFiles 
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   35
      Top             =   4140
      Width           =   195
   End
   Begin VB.CheckBox chAutoFiles 
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   33
      Top             =   3900
      Width           =   195
   End
   Begin VB.TextBox tDescrExt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1080
      TabIndex        =   32
      Text            =   "txt"
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox tPixExt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1080
      TabIndex        =   31
      Text            =   "jpg"
      Top             =   3660
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chAutoFiles 
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   28
      Top             =   4560
      Width           =   195
   End
   Begin VB.CheckBox chAutoFiles 
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   25
      Top             =   3540
      Width           =   195
   End
   Begin VB.CheckBox cEjectMedia 
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   6360
      Width           =   195
   End
   Begin VB.CheckBox cAutoClose 
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   6120
      Width           =   195
   End
   Begin VB.TextBox tAvi 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   960
      TabIndex        =   2
      Text            =   "avi"
      Top             =   2700
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox tDS 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   960
      TabIndex        =   3
      Text            =   "ds"
      Top             =   2940
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TextDebug 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   7020
      Width           =   6495
   End
   Begin MSComctlLib.ProgressBar pbAuto 
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   6780
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox chNoMess 
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   5880
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox chShots 
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   5340
      Width           =   195
   End
   Begin VB.CheckBox chSubFolders 
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   2460
      Width           =   195
   End
   Begin VB.CheckBox chAviHid 
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   2700
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox chDSHid 
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   2940
      Width           =   195
   End
   Begin SurVideoCatalog.XpB cCheckAllHid 
      Height          =   315
      Left            =   4620
      TabIndex        =   11
      Top             =   180
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      Caption         =   ""
      PicturePosition =   0
      ButtonStyle     =   3
      Picture         =   "FrmAuto.frx":0000
      PictureWidth    =   16
      PictureHeight   =   16
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin VB.ListBox LstFinded 
      Appearance      =   0  'Flat
      Height          =   2055
      ItemData        =   "FrmAuto.frx":059A
      Left            =   60
      List            =   "FrmAuto.frx":059C
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin SurVideoCatalog.XpB cUnCheckAllHid 
      Height          =   315
      Left            =   5580
      TabIndex        =   12
      Top             =   180
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   ""
      PicturePosition =   0
      ButtonStyle     =   3
      Picture         =   "FrmAuto.frx":059E
      PictureWidth    =   16
      PictureHeight   =   16
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB cScanGo 
      Height          =   435
      Left            =   4620
      TabIndex        =   14
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   767
      Caption         =   "ScanGo"
      ButtonStyle     =   3
      PictureWidth    =   0
      PictureHeight   =   0
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB cScan 
      Height          =   375
      Left            =   4380
      TabIndex        =   15
      Top             =   2760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Scan"
      ButtonStyle     =   3
      Picture         =   "FrmAuto.frx":0B38
      PictureWidth    =   16
      PictureHeight   =   16
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB cGo 
      Height          =   375
      Left            =   4380
      TabIndex        =   16
      Top             =   5220
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Go"
      ButtonStyle     =   3
      Picture         =   "FrmAuto.frx":10D2
      PictureWidth    =   16
      PictureHeight   =   16
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB cCheckDVD 
      Height          =   315
      Left            =   4620
      TabIndex        =   13
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Caption         =   "CheckDVD"
      ButtonStyle     =   3
      Picture         =   "FrmAuto.frx":166C
      PictureWidth    =   16
      PictureHeight   =   16
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB cNoAutoDups 
      Height          =   315
      Left            =   4620
      TabIndex        =   39
      Top             =   1020
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Caption         =   "NoDups"
      ButtonStyle     =   3
      Picture         =   "FrmAuto.frx":1C06
      PictureWidth    =   16
      PictureHeight   =   16
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000C&
      Height          =   915
      Left            =   60
      Top             =   5760
      Width           =   6495
   End
   Begin VB.Label lTxtTemplChange 
      BackStyle       =   0  'Transparent
      Caption         =   "ChTmpl"
      Height          =   195
      Left            =   540
      TabIndex        =   38
      Top             =   4920
      Width           =   5895
   End
   Begin VB.Label lPixTemplChange 
      BackStyle       =   0  'Transparent
      Caption         =   "ChTmpl"
      Height          =   195
      Left            =   540
      TabIndex        =   36
      Top             =   4140
      Width           =   5895
   End
   Begin VB.Label lAnyPix 
      BackStyle       =   0  'Transparent
      Caption         =   "AnyPix"
      Height          =   195
      Left            =   540
      TabIndex        =   34
      Top             =   3900
      Width           =   5895
   End
   Begin VB.Label lDescrExt 
      BackStyle       =   0  'Transparent
      Caption         =   "txt"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   480
      TabIndex        =   30
      Top             =   4680
      Width           =   5895
   End
   Begin VB.Label lDescr 
      BackStyle       =   0  'Transparent
      Caption         =   "Descr"
      Height          =   195
      Left            =   480
      TabIndex        =   29
      Top             =   4440
      Width           =   5895
   End
   Begin VB.Label lPixExt 
      BackStyle       =   0  'Transparent
      Caption         =   "jpg"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   480
      TabIndex        =   27
      Top             =   3660
      Width           =   5895
   End
   Begin VB.Label lblPix 
      BackStyle       =   0  'Transparent
      Caption         =   "Cover"
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Top             =   3420
      Width           =   5895
   End
   Begin VB.Label lEjectMedia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Eject"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   24
      Top             =   6360
      Width           =   5835
   End
   Begin VB.Label lAutoClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   23
      Top             =   6120
      Width           =   5775
   End
   Begin VB.Label lDSHid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ds"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   2940
      Width           =   3735
   End
   Begin VB.Label lAviHid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "avi"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   2700
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000C&
      Height          =   2415
      Left            =   60
      Top             =   3300
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   915
      Left            =   60
      Top             =   2340
      Width           =   6495
   End
   Begin VB.Label lNoMess 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NoMess"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   5880
      Width           =   5835
   End
   Begin VB.Label lShots 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shots"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   5340
      Width           =   3735
   End
   Begin VB.Label lSubFolders 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subs"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   2460
      Width           =   5895
   End
End
Attribute VB_Name = "frmAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private NStoreAuto(2) As String 'от 0, доп. фразы

'в паблик Private extAvi As String
'Private extDS As String


'добавление горизонтальной прокрутки listbox
Public Sub SetListboxScrollbar(lB As ListBox)
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

For i = 0 To lB.ListCount - 1
 new_len = 10 + ScaleX(TextWidth(lB.List(i)), ScaleMode, vbPixels)
 If max_len < new_len Then max_len = new_len
Next i

SendMessage lB.hwnd, LB_SETHORIZONTALEXTENT, max_len, 0
End Sub

Private Sub cAutoClose_Click()
ch_cAutoClose = cAutoClose.Value
End Sub

Private Sub cCheckAllHid_Click()
Call SendMessage(LstFinded.hwnd, LB_SETSEL, True, ByVal -1)
End Sub

Private Sub cCheckDVD_Click()

On Error Resume Next
If LstFinded.ListCount < 1 Then Exit Sub
Dim i As Long
Dim oldSel As Long

If LstFinded.ListIndex < 0 Then LstFinded.ListIndex = 0
oldSel = LstFinded.ListIndex

LstFinded.Visible = False

For i = 0 To LstFinded.ListCount - 1

    Select Case LCase$(right$(LstFinded.List(i), 6))
    Case "_1.vob"
        LstFinded.Selected(i) = True
    End Select

Next i
LstFinded.ListIndex = oldSel
LstFinded.Visible = True

End Sub

Private Sub cEjectMedia_Click()
ch_cEjectMedia = cEjectMedia.Value
End Sub

Private Sub chAutoFiles_Click(Index As Integer)
Select Case Index
Case 0: ch_chAutoFiles0 = chAutoFiles(Index).Value
Case 1: ch_chAutoFiles1 = chAutoFiles(Index).Value
Case 2: ch_chAutoFiles2 = chAutoFiles(Index).Value

Case 3: ch_chAutoFiles3 = chAutoFiles(Index).Value
Case 4: ch_chAutoFiles4 = chAutoFiles(Index).Value

'Case 5: ch_chAutoFiles5 = chAutoFiles(Index).Value
End Select

End Sub

Private Sub chAviHid_Click()
ch_chAviHid = chAviHid.Value
End Sub

Private Sub chDSHid_Click()
ch_chDSHid = chDSHid.Value
End Sub

Private Sub chSubFolders_Click()
ch_chSubFolders = chSubFolders.Value
End Sub

Private Sub cunCheckAllHid_Click()
Call SendMessage(LstFinded.hwnd, LB_SETSEL, False, ByVal -1)
End Sub
Private Sub chShots_Click()
If chShots.Value = vbChecked Then
    NoVideoProcess = False
Else
    NoVideoProcess = True
End If
ch_chShots = chShots.Value
End Sub


Private Sub cGo_Click()
Dim i As Integer, j As Integer
Dim Itm As ListItem

If LstFinded.SelCount < 1 Then Exit Sub

ToDebug "AProc..."

Dim UseTemplate As Boolean    'If .FrameAddEdit.Visible Then

'подготовить массивы с расширениями
Dim aPix() As String
Dim n_aPix As Long
Dim aTxt() As String
Dim n_aTxt As Long
Dim anyFile As String
Dim WeGetPicture As Boolean    'флаг получения картинки
Dim WeGetText As Boolean

Dim ClearAnnotFlag As Boolean    'в шаблоне не было, а мы забили из файла, - чистить
Dim ClearCoverFlag As Boolean    '-''-

n_aPix = Tokenize04(ExtPix, aPix(), " ", False)
n_aTxt = Tokenize04(ExtTxt, aTxt(), " ", False)




ToDebug "AutoFillInfo..."
TextDebug.Text = TextDebug.Text & Time & "> " & NStoreAuto(2) & vbCrLf
TextDebug.SelStart = Len(FormDebug.TextDebug.Text)

pbAuto.Max = LstFinded.SelCount
pbAuto.Value = 0

If Not (FrmMain.ListView.SelectedItem Is Nothing) Then
    '.ListView.SelectedItem.Selected = False 'снять текущую пометку
    For Each Itm In FrmMain.ListView.ListItems
        Itm.Selected = False
    Next
End If

If frmEditorFlag Then UseTemplate = True


' если в редакторе
If UseTemplate Then
    With frmEditor
        'грузить биг картинки
        If GetPic(.PicSS1Big, 1, "SnapShot1") Then
            .PicSS1Big.Picture = .PicSS1Big.Image
        End If
        If GetPic(.PicSS2Big, 1, "SnapShot2") Then
            .PicSS2Big.Picture = .PicSS2Big.Image
        End If
        If GetPic(.PicSS3Big, 1, "SnapShot3") Then
            .PicSS3Big.Picture = .PicSS3Big.Image
        End If

        'запомнить картинку шаблона
        Dim tmpPix As StdPicture
        If .PicFrontFace.Picture <> 0 Then
            Set tmpPix = .PicFrontFace.Picture
        End If

        'запомнить описание из шаблона
        Dim tmpAnnot As String
        tmpAnnot = .TextAnnotation
    End With
End If


AutoAddingFlag = True    'процесс из автодобавления
Mark2SaveFlag = False    'а нефиг краснить

FrmMain.ListView.MultiSelect = True


With frmEditor
    For i = 0 To LstFinded.ListCount - 1    '       Цикл по фильмам

        DoEvents
        If GetKeyState(vbKeyEscape) < 0 Then Exit For
        'If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For


        If LstFinded.Selected(i) Then
            LstFinded.ListIndex = i    'выделить (переместиться)

            addflag = True: editFlag = False    'нужно
            If rs.EditMode Then rs.CancelUpdate    'отменить, если в редактировании

            ClearVideo    'не чистить там обложку PicFrontFace
            If FrmMain.FrameView.Visible Then
                'nopic картинку
                FrmMain.Image0.Move 0, 0, FrmMain.FrameImageHid.Width, FrmMain.FrameImageHid.Height
                FrmMain.Image0.PaintPicture FrmMain.ImageList.ListImages(LastImageListInd).Picture, 0, 0, FrmMain.Image0.Width, FrmMain.Image0.Height    'nopic
            End If

            If Not UseTemplate Then
                'почистить поля, если не в редакторе
                Call ClearFields
                NoPicFrontFaceFlag = True: NoPic1Flag = True: NoPic2Flag = True: NoPic3Flag = True

            Else    'мы в редакторе - исп поля как шаблон (не чистить)

                If ClearAnnotFlag Then .TextAnnotation = vbNullString    'был изменен, а в шаблоне был пуст
                If ClearCoverFlag Then Set .PicFrontFace = Nothing    ' -''-

                'каждый раз установка флагов картинок (меняем флаг и далее...)

                If .PicFrontFace.Picture = 0 Then
                    NoPicFrontFaceFlag = True: SaveCoverFlag = False
                Else
                    NoPicFrontFaceFlag = False: SaveCoverFlag = True
                End If
                If .PicSS1Big.Picture = 0 Then
                    NoPic1Flag = True: SavePic1Flag = False
                Else
                    NoPic1Flag = False: SavePic1Flag = True
                End If
                If .PicSS2Big.Picture = 0 Then
                    NoPic2Flag = True: SavePic2Flag = False
                Else
                    NoPic2Flag = False: SavePic2Flag = True
                End If
                If .PicSS3Big.Picture = 0 Then
                    NoPic3Flag = True: SavePic3Flag = False
                Else
                    NoPic3Flag = False: SavePic3Flag = True
                End If

            End If

            If chNoMess.Value = vbChecked Then
                'не отвлекать
                CheckSameDisk = False    'не проверять сущ. ли диск в базе
            Else
                'CheckSameDisk = True
            End If

            If chShots.Value = vbChecked Then
                AutoShots = True
            Else
                AutoShots = False
            End If

            'открыть                                        (собственно процесс получения инфы)
            If OpenNewMovie(LstFinded.List(i)) Then

On Error GoTo err
                rs.AddNew
On Error GoTo 0

                '                                                   скриншоты
                If chShots.Value = vbChecked Then
                    If isAVIflag Then
                        'avi
                        If Not aferror Then .ComAutoScrShots_Click
                    Else
                        'не ави
                        If MPGCaptured Then .ComAutoScrShots_Click
                    End If
                End If




                Dim CurFilePathName As String    'c:\test без расширения
                Dim CurFilePath As String    'c:\

                'Dim CurExt As String
                'CurExt = getExtFromFile(LstFinded.List(i))
                'Dim CurFileName As String 'test.jpg
                'Dim CurFileNameOnly As String 'test без расширения
                'CurFileName = GetNameFromPathAndName(LstFinded.List(i))
                'GetExtensionFromFileName CurFileName, CurFileNameOnly
                'CurFilePath = GetPathFromPathAndName(LstFinded.List(i))
                'CurFilePathName = CurFilePath & CurFileNameOnly

                WeGetPicture = False: WeGetText = False
                CurFilePathName = vbNullString

                On Error Resume Next

                If ch_chAutoFiles0 Then    'взять картинку с тем же именем что и файл
                    If NoPicFrontFaceFlag Or ch_chAutoFiles3 Then    'нет картинки или можно переписывать
                        CurFilePathName = EraseExtFromFile(LstFinded.List(i))    'выделить путь
                        'поискать тут картину
                        If n_aPix > 0 Then
                            For j = 0 To n_aPix
                                If FileExists(CurFilePathName & "." & aPix(j)) Then
                                    .PicFrontFace.Picture = LoadPicture(CurFilePathName & "." & aPix(j))
                                    If NoPicFrontFaceFlag Then ClearCoverFlag = True    'почистить обложку
                                    NoPicFrontFaceFlag = False: SaveCoverFlag = True
                                    WeGetPicture = True
                                    Exit For    'взяли и хватит
                                End If
                            Next j
                        End If
                    End If    'NoPicFrontFaceFlag
                End If    'ch_chAutoFiles0

                If Not WeGetPicture Then    'снова если можно
                    If ch_chAutoFiles1 Then
                        If NoPicFrontFaceFlag Or ch_chAutoFiles3 Then    'нет картинки или можно переписывать 'взять любую
                            If n_aPix > -1 Then
                                For j = 0 To n_aPix
                                    CurFilePath = GetPathFromPathAndName(LstFinded.List(i))
                                    anyFile = GetFirstFileByExt(CurFilePath, aPix(j))
                                    If FileExists(anyFile) Then
                                        .PicFrontFace.Picture = LoadPicture(anyFile)
                                        If NoPicFrontFaceFlag Then ClearCoverFlag = True    'почистить обложку
                                        NoPicFrontFaceFlag = False: SaveCoverFlag = True
                                        WeGetPicture = True
                                        Exit For    'взяли
                                    End If
                                Next j
                            End If
                        End If
                    End If    'ch_chAutoFiles1
                End If    'NoPicFrontFaceFlag

                If UseTemplate Then    '
                    If Not (tmpPix Is Nothing) Then
                        If (Not WeGetPicture) And (tmpPix <> 0) Then
                            'установить взад картинку шаблона
                            Set .PicFrontFace.Picture = tmpPix
                        End If
                    End If
                End If

                If ch_chAutoFiles2 Then    'взять текст
                    If (Len(.TextAnnotation) = 0) Or ch_chAutoFiles4 Then    'если не занято шаблоном или переписывать
                        If Len(CurFilePathName) = 0 Then CurFilePathName = EraseExtFromFile(LstFinded.List(i))                     'выделить путь
                        'поискать тут текст
                        If n_aTxt > -1 Then
                            For j = 0 To n_aTxt
                                If FileExists(CurFilePathName & "." & aTxt(j)) Then

                                    Dim iFile As Integer, sfile As String
                                    iFile = FreeFile
                                    Open CurFilePathName & "." & aTxt(j) For Binary As #iFile
                                    sfile = AllocString_ADV(LOF(iFile))
                                    Get #iFile, , sfile
                                    Close #iFile

                                    If Len(.TextAnnotation) = 0 Then ClearAnnotFlag = True    'почистить описание
                                    .TextAnnotation = sfile
                                    WeGetText = True
                                    Exit For    'взяли и хватит
                                End If
                            Next j
                        End If
                    End If    'If Len(.TextAnnotation) = 0
                End If    'ch_chAutoFiles0

                If UseTemplate Then
                    If (Not WeGetText) And (Len(tmpAnnot) <> 0) Then
                        'установить взад описание из шаблона
                        .TextAnnotation = tmpAnnot
                    End If
                End If

                On Error GoTo 0

                'Придумать название
                .TextMName.Text = FileName2Title(LstFinded.List(i))

                '
                On Error Resume Next
                'сохранить
                SaveAutoAdd
                LstFinded.Selected(i) = False
                If err <> 0 Then ToDebug "Err.Go: " & err.Description
                On Error GoTo 0

            Else
                ToDebug "err_aadd: " & LstFinded.List(i)
                TextDebug.Text = TextDebug.Text & Time & "> " & "Error: " & LstFinded.List(i) & vbCrLf
                TextDebug.SelStart = Len(FormDebug.TextDebug.Text)
            End If
            pbAuto.Value = pbAuto.Value + 1
        End If    'помеченное
    Next

    AutoAddingFlag = False

    TextDebug.Text = TextDebug.Text & Time & "> " & "Done." & vbCrLf
    TextDebug.SelStart = Len(FormDebug.TextDebug.Text)

    pbAuto.Value = pbAuto.Max

    'в просмотр
    'If LastVMI <> 1 Then
    '    FrmMain.VerticalMenu_MenuItemClick 1, 0    '.ClearVideo Clear_mobjManager в .VerticalMenu_MenuItemClick 1, 0
    'Else
        addflag = False: editFlag = False
        ClearVideo
        Clear_mobjManager
        Set .PicFrontFace = Nothing
    'End If

End With
Screen.MousePointer = vbNormal

FrmMain.FrameView.Caption = FrameViewCaption & " " & FrmMain.ListView.ListItems.Count & " )"
'если выгружать - то выгружается автоформа If frmEditorFlag Then Unload frmEditor
'If Not frmEditorFlag Then
FrmMain.LVCLICK 'чтобы было что редактировать

If frmEditorFlag Then
editFlag = True
'загрузить в редактор последнюю
If frmEditorFlag Then
Screen.MousePointer = vbHourglass
        addflag = False: editFlag = True
        GetEditPix
        frmEditor.ImgPrCov.Picture = frmEditor.PicFrontFace.Picture
        Mark2SaveFlag = False
        GetFields
        Mark2SaveFlag = True
        frmEditor.ComSaveRec.BackColor = &HC0E0C0
        ToDebug "LV_Ed_Key: " & rs("key")
    Screen.MousePointer = vbNormal
End If
End If

'.Clear_mobjManager

Set tmpPix = Nothing
AutoShots = True


On Error Resume Next

'выдвинуть
If cEjectMedia.Value = vbChecked Then
    Dim tmp As String
    tmp = left$(LstFinded.List(0), 1) & ":"
    If apiDriveType(tmp) = 5 Then
        CDRomUnloadTray tmp
    End If
End If

'закрыть окно
If cAutoClose.Value = vbChecked Then Unload frmAuto

Exit Sub

err:
Screen.MousePointer = vbNormal
ToDebug "Err_AAGo: " & err.Description

End Sub


Private Sub chNoMess_Click()
If chNoMess.Value = vbChecked Then
    AutoNoMessFlag = True
Else
    AutoNoMessFlag = False
End If
ch_chNoMess = chNoMess.Value
End Sub



Private Sub cNoAutoDups_Click()
FillArrParse "filename", "undup"
End Sub

Private Sub cScan_Click()
Dim AutoAddDir As String
'mzt Dim i As Integer
Dim extensions As String

ToDebug "AScan..."
pbAuto.Value = pbAuto.min

Me.Enabled = False
AutoAddDir = BrowseForFolderByPath(FixPath(lastAutoAddfolderPath), NStoreAuto(1), frmAuto.hwnd)
Me.Enabled = True

If Len(AutoAddDir) = 0 Then Exit Sub
lastAutoAddfolderPath = AutoAddDir

LstFinded.Clear

If chAviHid.Value = vbChecked Then extensions = extAvi
If chDSHid.Value = vbChecked Then extensions = extensions & " " & extDS
extensions = Trim$(extensions)
extensions = Replace(extensions, "  ", " ")

ToDebug "ScanExt=" & extensions

'список фильмов
Screen.MousePointer = vbHourglass
FindFiles AutoAddDir, extensions, chSubFolders.Value
If LstFinded.ListCount < 1 Then Screen.MousePointer = vbNormal: Exit Sub

cCheckAllHid_Click 'пометить
LstFinded.ListIndex = 0

'прокрутка
SetListboxScrollbar LstFinded
pbAuto.Value = pbAuto.Max

CheckSameDisk = True 'типа вставили новый диск

Screen.MousePointer = vbNormal
End Sub


Private Sub cScanGo_Click()
cScan_Click
cGo_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'не канает, для остановки надо держать esc и происходит выгрузка
'If KeyAscii = 27 Then
'    If Not AutoAddingFlag Then Unload Me
'End If
End Sub

Private Sub Form_Load()
Dim tmp As String

frmAutoFlag = True
If chNoMess.Value = vbChecked Then
    AutoNoMessFlag = True
Else
    AutoNoMessFlag = False
End If
If chShots.Value = vbChecked Then
    NoVideoProcess = False
Else
    NoVideoProcess = True
End If

pbAuto.Value = pbAuto.Max

'текстовики для редактирования расширений
tAvi.Move lAviHid.left, lAviHid.top, lAviHid.Width, lAviHid.Height
tAvi.ZOrder 0
tDS.Move lDSHid.left, lDSHid.top, lDSHid.Width, lDSHid.Height
tDS.ZOrder 0

tPixExt.Move lPixExt.left, lPixExt.top, lPixExt.Width, lPixExt.Height
tPixExt.ZOrder 0
tDescrExt.Move lDescrExt.left, lDescrExt.top, lDescrExt.Width, lDescrExt.Height
tDescrExt.ZOrder 0


'расширения
tmp = VBGetPrivateProfileString("AutoAdd", "AVIsExt", iniFileName)
If Len(tmp) = 0 Then
    extAvi = "avi vid divx"
Else
    extAvi = tmp
End If
tmp = VBGetPrivateProfileString("AutoAdd", "DirectShowExt", iniFileName)
If Len(tmp) = 0 Then
    extDS = "mpg mpeg vob asf wmv mp4 mkv"
Else
    extDS = tmp
End If
lAviHid.Caption = extAvi
lDSHid.Caption = extDS

tmp = VBGetPrivateProfileString("AutoAdd", "CoverExt", iniFileName)
If Len(tmp) = 0 Then
    ExtPix = "jpg gif bmp"
Else
    ExtPix = tmp
End If
lPixExt.Caption = ExtPix

tmp = VBGetPrivateProfileString("AutoAdd", "DescrExt", iniFileName)
If Len(tmp) = 0 Then
    ExtTxt = "txt"
Else
    ExtTxt = tmp
End If
lDescrExt.Caption = ExtTxt



'язык
GetLangAuto

'галочки
chSubFolders.Value = ch_chSubFolders
chAviHid.Value = ch_chAviHid
chDSHid.Value = ch_chDSHid
chShots.Value = ch_chShots
chNoMess.Value = ch_chNoMess
cAutoClose.Value = ch_cAutoClose
cEjectMedia.Value = ch_cEjectMedia

chAutoFiles(0).Value = ch_chAutoFiles0 'cAddCoverExt
chAutoFiles(1).Value = ch_chAutoFiles1 'cAddCoverAny
chAutoFiles(2).Value = ch_chAutoFiles2 'cAddTXTDescr

chAutoFiles(3).Value = ch_chAutoFiles3
chAutoFiles(4).Value = ch_chAutoFiles4

'chAutoFiles(5).Value = ch_chAutoFiles5

'посерить, если без шаблона
If frmEditorFlag Then
Else
chAutoFiles(3).Enabled = False
chAutoFiles(4).Enabled = False
End If

'хук - no auto CD
If Not DebugMode Then
    If Opt_QueryCancelAutoPlay Then
       Call HookWindowAutoAdd(Me.hwnd, Me) 'Friend Function WindowProc
    End If
End If

End Sub
Friend Function WindowProc(hwnd As Long, Msg As Long, wp As Long, lp As Long) As Long
Dim result As Long
Select Case Msg
Case m_RegMsg       ' QueryCancelAutoPlay
    ' TRUE: cancel AutoRun
    ' *must* be 1, not -1!
    ' FALSE: allow AutoRun
    result = 1
ToDebug "CancelAutoPlayAuto"
Case Else
    ' Pass along to default window procedure.
    result = InvokeWindowProcAutoAdd(hwnd, Msg, wp, lp)
End Select
' Return desired result code to Windows.
WindowProc = result
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

Call UnhookWindowAutoAdd(Me.hwnd)

NoVideoProcess = False
CheckSameDisk = True

'сохранить галочки
WriteKey "AutoAdd", "chSubFolders", CStr(ch_chSubFolders), iniFileName
WriteKey "AutoAdd", "chAvi", CStr(ch_chAviHid), iniFileName
WriteKey "AutoAdd", "chDS", CStr(ch_chDSHid), iniFileName
WriteKey "AutoAdd", "chShots", CStr(ch_chShots), iniFileName
WriteKey "AutoAdd", "chNoMess", CStr(ch_chNoMess), iniFileName
WriteKey "AutoAdd", "cAutoClose", CStr(ch_cAutoClose), iniFileName
WriteKey "AutoAdd", "cEjectMedia", CStr(ch_cEjectMedia), iniFileName

WriteKey "AutoAdd", "cAddCoverExt", CStr(ch_chAutoFiles0), iniFileName
WriteKey "AutoAdd", "cAddCoverAny", CStr(ch_chAutoFiles1), iniFileName
WriteKey "AutoAdd", "cAddTXTDescr", CStr(ch_chAutoFiles2), iniFileName

WriteKey "AutoAdd", "cPixTemplChange", CStr(ch_chAutoFiles3), iniFileName
WriteKey "AutoAdd", "cTxtTemplChange", CStr(ch_chAutoFiles4), iniFileName

'WriteKey "AutoAdd", "cNoAutoDups", CStr(ch_chAutoFiles5), iniFileName

End Sub

Private Sub Form_Resize()
'Background
If lngBrush <> 0 Then
GetClientRect hwnd, rctMain
FillRect hdc, rctMain, lngBrush
End If
'If FrmAuto.WindowState = vbMinimized Then FrmMain.WindowState = vbMinimized
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAutoFlag = False
'?вообще не в фокусе MakeNormal FrmMain.hwnd 'это выведет окно в фокус
End Sub
Private Sub GetLangAuto()
Dim Contrl As Control
Dim i As Integer
'mzt Dim temp As String

On Error Resume Next

If Dir(lngFileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = vbNullString Or Len(lngFileName) = 0 Then
Call myMsgBox("Не найден файл локализации! Исправьте параметр LastLang в global.ini" & vbCrLf & "Language file not found: " & vbCrLf & lngFileName, vbCritical, , Me.hwnd)
End If

'ToDebug "Чтение файла локализации: " & lngFileName

For Each Contrl In frmAuto.Controls
If right$(Contrl.name, 3) <> "Hid" Then

If TypeOf Contrl Is Label Then '                           Label
Contrl.Caption = ReadLangAuto(Contrl.name & ".Caption", Contrl.Caption)
End If

If TypeOf Contrl Is XpB Then '                               XPB
 Contrl.Caption = ReadLangAuto(Contrl.name & ".Caption", Contrl.Caption)
 Contrl.pInitialize
End If

End If 'не Hid
Next 'Contrl

'                                                       NamesStoreStat()
For i = 0 To UBound(NStoreAuto)
NStoreAuto(i) = ReadLangAuto("NStoreAuto" & i)
Next i

Me.Caption = "SurVideoCatalog - " & NStoreAuto(0)
Me.Icon = FrmMain.Icon

LstFinded.BackColor = LVBackColor
LstFinded.ForeColor = LVFontColor
TextDebug.BackColor = LVBackColor
TextDebug.ForeColor = LVFontColor


End Sub


Private Sub lAviHid_Click()
tAvi.Visible = True: tAvi.SetFocus: tAvi = extAvi
End Sub

Private Sub lDescrExt_Click()
tDescrExt.Visible = True: tDescrExt.SetFocus: tAvi = ExtTxt
End Sub

Private Sub lDSHid_Click()
tDS.Visible = True: tDS.SetFocus: tDS = extDS
End Sub

Private Sub lPixExt_Click()
tPixExt.Visible = True: tPixExt.SetFocus: tPixExt = ExtPix
End Sub

Private Sub LstFinded_Click()
LstFinded.Refresh
End Sub

Private Sub tAvi_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    tAvi.Visible = False
    If extAvi <> tAvi Then
        WriteKey "AutoAdd", "AVIsExt", tAvi.Text, iniFileName
        extAvi = tAvi
        lAviHid.Caption = tAvi
    End If
Case 27
    tAvi.Visible = False
End Select
End Sub



Private Sub tAvi_LostFocus()
tAvi.Visible = False
End Sub

Private Sub tDescrExt_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    tDescrExt.Visible = False
    If ExtTxt <> tDescrExt Then
        WriteKey "AutoAdd", "DescrExt", tDescrExt.Text, iniFileName
        ExtTxt = tDescrExt
        lDescrExt.Caption = tDescrExt
    End If
Case 27
    tDescrExt.Visible = False
End Select

End Sub

Private Sub tDS_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    tDS.Visible = False
    If extDS <> tDS Then
        WriteKey "AutoAdd", "DirectShowExt", tDS.Text, iniFileName
        extDS = tDS
        lDSHid.Caption = tDS
    End If
Case 27
    tDS.Visible = False
End Select
End Sub

Private Sub tDS_LostFocus()
    tDS.Visible = False
End Sub

Private Sub tPixExt_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    tPixExt.Visible = False
    If ExtPix <> tPixExt Then
        WriteKey "AutoAdd", "CoverExt", tPixExt.Text, iniFileName
        ExtPix = tPixExt
        lPixExt.Caption = tPixExt
    End If
Case 27
    tPixExt.Visible = False
End Select

End Sub
Public Sub FillArrParse(fldName As String, WhatToDo As String)
'ака FillCombosParse
'заполняет массив потрошенными значениями
'fldName - поле базы
Dim tmp As String
Dim i As Long, j As Long
Dim R() As String
Dim strSQL As String
Dim rsTmp As DAO.Recordset
Dim rsArr() As String
Dim sDelim As String
Screen.MousePointer = vbHourglass

On Error Resume Next
strSQL = "Select " & fldName & " From Storage"
Set rsTmp = DB.OpenRecordset(strSQL)
If err Then
    ToDebug "Err_FArP: " & strSQL
    Screen.MousePointer = vbNormal
    Exit Sub
End If
On Error GoTo 0

If LCase$(fldName) = "filename" Then
    sDelim = "|"
Else
    sDelim = ",;"
End If

ReDim rsArr(0)    'заполнение с 0
If Not (rsTmp.BOF And rsTmp.EOF) Then
    rsTmp.MoveLast: rsTmp.MoveFirst

    For i = 1 To rsTmp.RecordCount
        If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For

        If IsNull(rsTmp(0)) Then
        Else

            If Tokenize04(rsTmp(0), R(), sDelim, True) > -1 Then
                For j = 0 To UBound(R)
                    rsArr(UBound(rsArr)) = R(j)
                    ReDim Preserve rsArr(UBound(rsArr) + 1)
                Next j
            End If
        End If
        rsTmp.MoveNext
    Next i
End If

If UBound(rsArr) > 0 Then
    TriQuickSortString rsArr        'sorts your string array
    remdups rsArr                   'removes dups

    'использовать по затребованному назначению
    Select Case WhatToDo
    Case "undup"    'анчек дублей по имени файла в автодабе
        If LstFinded.ListCount < 1 Then Screen.MousePointer = vbNormal: Exit Sub
'        Dim oldSel As Long
        If LstFinded.ListIndex < 0 Then LstFinded.ListIndex = 0
'        oldSel = LstFinded.ListIndex

        LstFinded.Visible = False
        
        For j = 0 To LstFinded.ListCount - 1
            For i = 0 To UBound(rsArr)
                If GetNameFromPathAndName(LstFinded.List(j)) = rsArr(i) Then
                    LstFinded.Selected(j) = False
                    Exit For    'i
                End If
            Next i
        Next j
'       LstFinded.ListIndex = oldSel
        LstFinded.Visible = True


    End Select
End If

Set rsTmp = Nothing

Screen.MousePointer = vbNormal
End Sub
