Attribute VB_Name = "modGraph"
Option Explicit
Public Function GetPic(lImage As PictureBox, Base As Integer, dfield As String) As Boolean
'true - картинка есть
'If FrameActer.Visible Then DoEvents

Dim img As ImageFile
Dim vec As Vector

Dim PicSize As Long
Dim b() As Byte

Dim TimerStatus As Boolean
Dim pwidth As Integer
Dim pheight As Integer

On Error Resume Next


If Base = 1 Then
    If NoDBFlag Then Exit Function
    If rs.RecordCount < 1 Then Exit Function
End If


GetPic = True
TimerStatus = FrmMain.Timer2.Enabled

Select Case Base
Case 1
    'rs base
    FrmMain.Timer2.Enabled = False

    Dim ForCiverFlag As Boolean
    If lImage.name = FrmMain.PicFaceV.name Then ForCiverFlag = True
    If lImage.name = frmEditor.PicFrontFace.name Then ForCiverFlag = True
    If (lImage.name = FrmMain.PicTempHid(1).name) Then
        If lImage.Index = 1 Then ForCiverFlag = True
    End If

    If ForCiverFlag Then
        'For face
        NoPicFrontFaceFlag = False
        PicSize = rs.Fields(dfield).FieldSize
        'Debug.Print lImage.name, "FielsSize=" & PicSize ', "ActualSize=" & rs!PicFaceV.ActualSize
        If PicSize = 0& Then
            NoPicFrontFaceFlag = True
            GetPic = False
            FrmMain.Timer2.Enabled = TimerStatus
            Exit Function
        End If
    Else
        'For ScreenShots
        If NoPic1Flag And NoPic2Flag And NoPic3Flag Then GetPic = False: Exit Function
        PicSize = rs.Fields(dfield).FieldSize
        'Debug.Print lImage.name, "FielsSize=" & PicSize
        If PicSize = 0& Then
            GetPic = False

            If FrmMain.FrameView.Visible Then FrmMain.Timer2.Enabled = True    'Timer2_Timer:Timer2.Enabled = True

            Exit Function
        Else
            SlideShowLastGoodPic = SlideShowLastFlag
            'Debug.Print SlideShowLastGoodPic
        End If

    End If

    ReDim b(PicSize - 1)
    b() = rs.Fields(dfield).GetChunk(0, PicSize)

Case 2
    ' ars base

    NoPicActFlag = False
    PicSize = ars.Fields(dfield).FieldSize
    If PicSize = 0& Then
        NoPicActFlag = True
        GetPic = False
        Exit Function
    End If
    ReDim b(PicSize - 1)
    b() = ars.Fields(dfield).GetChunk(0, PicSize)

End Select

'Set lImage = rs.Fields(dfield).Value 'PictureFromByteStream(B)
'   lPtr = VarPtr(B(0))

'Picture
'Set img = New ImageFile
Set vec = New Vector
vec.BinaryData = b
Set img = vec.ImageFile

If Not img Is Nothing Then
    'еще бывает нет img когда wia не зарегистрирован
    'это ошибка, но там вначале resume next

    If lImage.name = "PicSS1" _
       Or lImage.name = "PicSS2" _
       Or lImage.name = "PicSS3" _
       Or lImage.name = "Image0" Then


        Ratio = img.Width / img.Height
        If Ratio < 1 Then Ratio = 1.333    'допущение для картинок, уже записанных в базу с реальным (и кривым) разрешением
        lImage.Height = lImage.Width / Ratio
        'pwidth = ScaleX(img.Width, vbPixels, vbTwips)
        'pheight = ScaleY(img.Height, vbPixels, vbTwips)
        pwidth = img.Width * Screen.TwipsPerPixelX
        pheight = img.Height * Screen.TwipsPerPixelY
        FrmMain.PicTempHid(0).Width = pwidth
        FrmMain.PicTempHid(0).Height = pheight

        Set FrmMain.PicTempHid(0).Picture = img.ARGBData.Picture(img.Width, img.Height)

        lImage.Cls
        If lImage.name = "Image0" Then lImage.top = FrmMain.FrameImageHid.Height / 2 - lImage.Height / 2    ' скриншот по центру
        lImage.PaintPicture FrmMain.PicTempHid(0).Picture, 0, 0, lImage.Width, lImage.Height, 0, 0, pwidth, pheight, SRCCOPY

        If FrmMain.FrameView.Visible Or FrmMain.FrameActer.Visible Then DoEvents

    ElseIf lImage.name = "PicActFoto" Then
        pwidth = FrmMain.PicActFotoScroll.Width / Screen.TwipsPerPixelX - 5
        pheight = FrmMain.PicActFotoScroll.Height / Screen.TwipsPerPixelY - 5
        'lImage.Width = img.Width * Screen.TwipsPerPixelX * 2
        'lImage.Height = img.Height * Screen.TwipsPerPixelY * 2


        'FrmMain.PicActFotoScroll.ScaleHeight
        'FrmMain.PicTempHid(1).Width = img.Width * Screen.TwipsPerPixelX
        'FrmMain.PicTempHid(1).Height = img.Height * Screen.TwipsPerPixelY
        Set FrmMain.PicTempHid(1).Picture = img.ARGBData.Picture(img.Width, img.Height)    'реальный размер для копи mnuCopyFoto_Click mnuSaveFoto_Click

        If FrmMain.chActFotoScale.Value = vbChecked Then
            'увеличим до границ с пропорцией
            Dim IP As ImageProcess
            Set IP = New ImageProcess
            IP.Filters.Add IP.FilterInfos("Scale").FilterID
            IP.Filters(1).Properties("MaximumWidth").Value = pwidth  '640
            IP.Filters(1).Properties("MaximumHeight").Value = pheight    '480
            IP.Filters(1).Properties("PreserveAspectRatio").Value = True
            Set img = IP.Apply(img)
            Set lImage.Picture = img.ARGBData.Picture(img.Width, img.Height)
        Else
            Set lImage.Picture = FrmMain.PicTempHid(1).Picture
        End If

        'If FrmMain.FrameView.Visible Or FrmMain.FrameActer.Visible Then DoEvents
        If FrmMain.FrameActer.Visible Then DoEvents
    Else    '    картинки для обложки, ssbig редактора

        lImage.Width = img.Width * Screen.TwipsPerPixelX
        lImage.Height = img.Height * Screen.TwipsPerPixelY

        FrmMain.PicTempHid(1).Width = lImage.Width
        FrmMain.PicTempHid(1).Height = lImage.Height

        Set lImage.Picture = img.ARGBData.Picture(img.Width, img.Height)

        'If FrmMain.FrameView.Visible Or FrmMain.FrameActer.Visible Then DoEvents
        If FrmMain.FrameView.Visible Then DoEvents
        'lImage.Refresh
    End If

End If    '   If LoadJPGFromPtr

Erase b
Set img = Nothing
Set vec = Nothing
'Set m_cDib = Nothing
FrmMain.Timer2.Enabled = TimerStatus

End Function

Public Function LoadPictureWIA(iFile As String) As StdPicture
Dim img As ImageFile

On Error GoTo wiaerr

'Private vec As Vector
Set img = New ImageFile
img.LoadFile iFile

If Not img Is Nothing Then
    Set LoadPictureWIA = img.ARGBData.Picture(img.Width, img.Height)
End If

Exit Function
wiaerr:
ToDebug "Error LP_WIA: " & err.Description
'MsgBox err.Description, vbCritical
End Function

Private Sub GetPicFromUrlPut2Base(u As String)
'получить картинку из инета по абсолютному URL
'для InetGetPics
If BaseReadOnly Or BaseReadOnlyU Then Exit Sub

Dim tmp As String
tmp = LCase$(left$(u, 7))
If (tmp = "file://") Or (tmp = "http://") Or (tmp = "https:/") Then

    OpenURLProxy u, "pic" 'получили картинку в PicFrontFace
    If frmEditor.PicFrontFace.Picture <> 0 Then
        rs.Edit
            Pic2JPG frmEditor.PicFrontFace, 1, "FrontFace" 'положили в базу
        rs.Update
    End If
Else 'не абс путь

ToDebug msgsvc(48) & ": " & u

End If
End Sub


Public Sub InetGetPics(ch As Boolean)

'получить картинки, если поле dbCoverPathInd заполнено путем

If BaseReadOnly Or BaseReadOnlyU Then
    'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
    Exit Sub
End If

Dim ret As VbMsgBoxResult
Dim Itm As ListItem
Dim cp As String
Dim i As Long

With FrmMain
    .PBar.min = 0
    If ch Then
        .PBar.Max = CheckCount
    Else
        .PBar.Max = SelCount
    End If
    .PBar.Value = 0
    'TextItemHid.ZOrder 0
    .PBar.ZOrder 0

    .Timer2.Enabled = False

    ret = myMsgBox(msgsvc(47), vbYesNoCancel, , .hwnd)    'заменять, если есть
    If ret <> vbCancel Then
        i = 0
        If ch Then    'помеченные
            If CheckCount > 0 Then
                For Each Itm In .ListView.ListItems
                    DoEvents
                    If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
                    If Itm.Checked Then
                        i = i + 1
                        RSGoto Itm.Key
                        If (rs("FrontFace").FieldSize = 0) Or ((rs("FrontFace").FieldSize <> 0) And (ret = vbYes)) Then    'менять
                            cp = CheckNoNull("CoverPath")
                            If Len(cp) <> 0 Then    'заполнено поле CoverPath
                                GetPicFromUrlPut2Base frmEditor.cBasePicURL & cp
                            End If
                        End If
                        .PBar.Value = i
                    End If
                    If i = CheckCount Then Exit For
                Next
            End If    'CheckCount

        Else    'выделенные
            If SelCount > 0 Then
                For Each Itm In .ListView.ListItems
                    DoEvents
                    If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
                    If Itm.Selected Then
                        i = i + 1
                        RSGoto Itm.Key
                        If (rs("FrontFace").FieldSize = 0) Or ((rs("FrontFace").FieldSize <> 0) And (ret = vbYes)) Then    'менять
                            cp = rs("CoverPath")
                            If Len(cp) <> 0 Then    'заполнено поле CoverPath
                                GetPicFromUrlPut2Base frmEditor.cBasePicURL & cp
                            End If
                        End If
                        .PBar.Value = i
                    End If
                    If i = SelCount Then Exit For
                Next
            End If    'SelCount
        End If    'ch

    End If    'vbCancel

    .TextItemHid.ZOrder 0

    ' в лвклик RestoreBasePos
    'If Not MultiSel Then
    .Timer2.Enabled = True
    .LVCLICK
    'End If
End With
End Sub

Public Function SavePicFromPic(Pict As StdPicture, owHWND As Long, Optional myFileName As String) As Boolean
'положить картинку из picturebox в файл

Dim iFile As String
Dim img As ImageFile
Dim IP As ImageProcess
Dim vec As Vector
Dim WithWIA As Boolean
Dim sType As String

On Error GoTo err

iFile = pSaveDialog(owHWND, DTitle:=NamesStore(10), myFileName:=myFileName)
If iFile <> vbNullString Then

    WithWIA = True
    Select Case LCase$(right$(iFile, 3))
    Case "bmp"
        WithWIA = False
        'sType = wiaFormatBMP
        SavePicture Pict, iFile
    Case "jpg"
        sType = wiaFormatJPEG
    Case "gif"
        sType = wiaFormatGIF
    Case "png"
        sType = wiaFormatPNG
    Case "tif"
        sType = wiaFormatTIFF
    End Select
Else
    Exit Function    'нет файла
End If

If WithWIA Then

    Set vec = New Vector
    Set IP = New ImageProcess
    vec.BinaryData = Picture2Array(Pict)
    Set img = vec.ImageFile
    Set vec = Nothing

    While (IP.Filters.Count > 0)
        IP.Filters.Remove 1
    Wend

    IP.Filters.Add IP.FilterInfos("Convert").FilterID
    IP.Filters(1).Properties("FormatID").Value = sType     'FormatID 1

    Select Case sType
    Case wiaFormatJPEG
        IP.Filters(1).Properties("Quality").Value = QJPG    'Quality 2
    Case wiaFormatGIF
    Case wiaFormatPNG
    Case wiaFormatTIFF
        IP.Filters(1).Properties("Compression").Value = "Uncompressed"    'Compression 3   LZW CCITT3 CCITT4 RLE Uncompressed
    End Select

    Set img = IP.Apply(img)
    Set IP = Nothing

    If img Is Nothing Then
        'не вышло
        ToDebug "Err.SPFP: img Is Nothing"
        Exit Function
    Else
        'img.FileData.BinaryData
        If FileExists(iFile) Then Kill iFile
        img.SaveFile iFile

    End If
End If    'WithWIA

    ''exif info add
    'IP.Filters.Add IP.FilterInfos("Exif").FilterID
    'IP.Filters(1).Properties("ID") = 40091
    'IP.Filters(1).Properties("Type") = VectorOfBytesImagePropertyType
    'v.SetFromString "This Title tag written by Windows Image Acquisition Library v2.0"
    'IP.Filters(1).Properties("Value") = v
    'Set img = IP.Apply(img)

SavePicFromPic = True
Exit Function
err:
ToDebug "Err.SPFP: " & err.Description

End Function

Public Sub ShowInShowPic(w As Integer, frm As Form)
'Dim k As Integer ' коэф для выравнивания аспекта фильмов с 352 шириной
'w - 0 - sshot, 1 - face
'показ картинки из PicTempHid на форму FormShowPic
'FormShowPic.Visible = False
Dim tRes As String    'разрешение

If w = 1 Then FrmMain.PicTempHid(w).Picture = FrmMain.PicTempHid(w).Image
'w = 1 у скриншотов иначе путается размеры формы при разных размерах скриншотов

If FrmMain.PicTempHid(w).Picture = 0 Then Unload FormShowPic: Exit Sub

If Opt_CenterShowPic Then
    'центровать всегда глобально
    frmShow_xPos = (Screen.Width - (FrmMain.PicTempHid(w).Width)) \ 2
    frmShow_yPos = (Screen.Height - (FrmMain.PicTempHid(w).Height)) \ 2
    If frmShow_xPos < 0 Then frmShow_xPos = 0
    If frmShow_yPos < 0 Then frmShow_yPos = 0

Else
    'позиция 1:1 окна
    If IsCoverShowFlag Then
        frmShow_xPos = CoverWindLeft
        frmShow_yPos = CoverWindTop
    Else
        frmShow_xPos = ScrShotWindLeft
        frmShow_yPos = ScrShotWindTop
    End If

    If frmShow_xPos = 0 And frmShow_yPos = 0 Then    'немного не верно - не дает установить окно в 0,0
        'центровать при первой загрузке
        frmShow_xPos = (FrmMain.Width - (FrmMain.PicTempHid(w).Width)) \ 2 + FrmMain.left
        frmShow_yPos = (FrmMain.Height - (FrmMain.PicTempHid(w).Height)) \ 2 + FrmMain.top
    End If
End If    'Opt_CenterShowPic

'то же в HScroll_Change
SetFrmShowPicPicture w

'If Not FormShowPicFlag Then FormShowPic.Show , FrmMain 'модально

'MakeTopMost FormShowPic.hwnd
'SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
'SetWindowPos FormShowPic.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, TOPMOST_FLAGS Or SWP_NOACTIVATE

FormShowPic.Visible = True    'надо, могли прятать
'не показывать разрешение если маленькая картинка
If (FormShowPic.PicHB.Height * 3) < FormShowPic.ScaleHeight Then

    'FormShowPic.ForeColor = 49344
    ''On Error Resume Next
    ''FormShowPic.ForeColor = GetPixel(PicTempHid(w).hdc, 0, 0)
    'If GetPixel(PicTempHid(w).hdc, 0, 0) > 8000000 Then FormShowPic.ForeColor = &HC00000
    'Debug.Print FormShowPic.ForeColor
    'PicCoverTextWnd.Line (0, 0)-(PicCoverTextWnd.Width, PicCoverTextWnd.Height), CoverHorBackColor, BF
    tRes = FormShowPic.Width / Screen.TwipsPerPixelX & "x" & FormShowPic.Height / Screen.TwipsPerPixelY

    FormShowPic.Line (3, 3)-(FormShowPic.TextWidth(tRes) + 2, FormShowPic.TextHeight(tRes) + 2), 0, BF
    FormShowPic.CurrentX = 3: FormShowPic.CurrentY = 3
    FormShowPic.Print tRes
    'тоже в
    'SetFrmShowPicPicture '3,3
    'hb_cscroll_Change print
End If

If Not FormShowPicIsModal Then
    FormShowPic.Show 0, frm  'модально 1 раз
    FormShowPicIsModal = True
End If

End Sub


Public Sub ScrShotClick()
Dim tOld As Boolean

tOld = FrmMain.Timer2.Enabled
FrmMain.Timer2.Enabled = False
'SShotClickFlag = True 'для слайдера, чтоб не выполнялся

'If IsCoverShowFlag Then If FormShowPicFlag Then FormShowPic.Hide
If IsCoverShowFlag Then If FormShowPicLoaded Then FormShowPic.Hide
IsCoverShowFlag = False

If NoPic1Flag And NoPic2Flag And NoPic3Flag Then Exit Sub


'грузится FormShowPic и стартует первый скролл
DoEvents
Load FormShowPic
With FormShowPic

    .hb_cScroll.Max(efsHorizontal) = SShotsCount

    Select Case SlideShowLastGoodPic
    Case 0
        .hb_cScroll.Value(efsHorizontal) = 1
    Case 1
        .hb_cScroll.Value(efsHorizontal) = 2
    Case 2
        .hb_cScroll.Value(efsHorizontal) = 3
    End Select
End With


If FrmMain.PicTempHid(0).Picture <> 0 Then
    FormShowPic.PicHB.Visible = True    'used
    FormShowPic.hb_cScroll.Visible(efsHorizontal) = True
    ViewScrShotFlag = True
    ShowInShowPic 0, FrmMain
    
    
    'FormShowPic.Visible = True
End If

FrmMain.Timer2.Enabled = tOld
End Sub


