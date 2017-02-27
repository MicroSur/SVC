Attribute VB_Name = "ModAddEdit"
Option Explicit

Public sTimeSum As String 'сохранять для отката суммирования времени

Public AutoAddingFlag As Boolean    'работает автопоиск (не закрывать окно, нет звука)
Public AddEditCapt As String 'название фрейма редактора
Public MPGCaptured As Boolean    'захвачен ли кадр
Public MpegMediaOpen As Boolean    'открыт мпег
Public filltrueAdd As Boolean    'filltrue флаг добавлять
Public TextTimeMSHid As String    'время фильма в мс

''все поля редактора
'Public sTextMName As String    'TextMName
'Public sTextLabel As String    'TextLabel
'Public sTextGenre As String    'TextGenre
'Public sTextYear As String    'TextYear
'Public sTextCountry As String    'TextCountry
'Public sTextAuthor As String    'TextAuthor
'Public sTextRole As String    'TextRole
'Public sTextTimeHid As String    'TextTimeHid
'Public sTextResolHid As String    'TextResolHid
'Public sTextAudioHid As String    'TextAudioHid
'Public sTextFPSHid As String    'TextFPSHid
'Public sTextFilelenHid As String  'TextFilelenHid
'Public sTextCDN As String    'TextCDN
'Public sComboNos As String    'ComboNos
'Public sTextVideoHid As String    'TextVideoHid
'Public sTextSubt As String  'TextSubt
'Public sTextLang As String  'TextLang
'Public sTextRate As String    'TextRate
'Public sTextFileName As String  'TextFileName
'Public sTextUser As String    'TextUser
'Public sCDSerialCur As String    'CDSerialCur
'Public sTextOther As String    'TextOther
'Public sTextCoverURL As String  'TextCoverURL
'Public sTextMovURL As String    'TextMovURL
'Public sTextAnnotation  As String   'TextAnnotation



' для автозаполнения
Public m_bEditFromCode As Boolean

Public mobjManager As FilgraphManager
Public mobjSampleGrabber As ISampleGrabber
Public objPosition As IMediaPosition
Public objVideo As IBasicVideo
Public objAudio As IBasicAudio
Public objVideoW As IVideoWindow

Public m_cAVI As cAVIFrameExtract
Private Const MAXDWORD = &HFFFFFFFF    'для GetAllVobsSize

Private Const cstrGrabberName As String = "SampleGrabber"
Private Const cstrSVDName As String = "Universal Open Source MPEG Source"
'новый osmpegsplitter.1.0.0.3_nt.exe Private Const cstrSVDName As String = "Mpeg Source"
Private Const cstrVDName As String = "GPL MPEG-1/2 Decoder"
Private Const cstrMPG2DemName As String = "MPEG-2 Demultiplexer"
'Private Const cstrFVDName As String = "Universal Open Source MPEG Splitter"
'Private Const cstrFVDName As String = "Fraunhofer Video Decoder"
'Private Const cstrFADName As String = "Fraunhofer Audio Decoder"
'Private Const cstrM2SPLName As String = "MPEG-2 Splitter"
Private Const cstrM1SSName As String = "MPEG-I Stream Splitter"    'quartz.dll
Private Const cstrM1VDName As String = "MPEG Video Decoder"    'quartz.dll

Public Const cMPGRange As Integer = 10000    'For MPGPOSScroll

Public Sub ClearVideo()
Dim i As Integer
On Error Resume Next

'очистка (видео)
FrmMain.Timer2.Enabled = False
ToDebug "ClVid_proc..."

With frmEditor


.movie.Width = MovieEd_W
.movie.Height = MovieEd_H


isMPGflag = False: isAVIflag = False: isDShflag = False
pos1 = 0: pos2 = 0: pos3 = 0
.PositionP.Value = 0: .Position.Value = 0

MPGCodec = vbNullString
PixelRatio = 1: PixelRatioSS = 1

.ComRND(0).Enabled = False: .ComRND(1).Enabled = False: .ComRND(2).Enabled = False
.ComAutoScrShots.Enabled = False
For i = 0 To 2: .optAspect(i).Enabled = False: Next i
        
'Set PicFrontFace = Nothing 'нет для авто
'Set picCanvas = Nothing
'его теперь не почистить - видим  Set FrmMain.PicFaceV = Nothing
'Set Image0 = Nothing

'его теперь не почистить - видим Set .ImgPrCov = Nothing

If Not (m_cAVI Is Nothing) Then m_cAVI.filename = vbNullString 'unload

If MpegMediaOpen Then Call MpegMediaClose
'Debug.Print Time, "clear video"

Set .movie = Nothing
End With
End Sub


Public Sub MPGCaptureBasicVideo(Pict As PictureBox)
Dim objBVideo As KTLDirectShow.IBasicVideo
'Dim objBVideo As QuartzTypeLib.IBasicVideo
Dim lngLength As Long
Dim hGlobal As OLE_HANDLE
Dim lngPointer As Long
Dim tmp As Single

'If MediaState = 2 Then mobjManager.Pause 'играет, на паузу
ToDebug "DirectShow Capture..."
MPGCaptured = False
'при неудачном .GetCurrentImage - Атомейшн еррор
On Error Resume Next
Set objBVideo = mobjManager
If (objBVideo Is Nothing) Then Exit Sub
Call objBVideo.GetCurrentImage(lngLength, ByVal 0&)
If (0& < lngLength) Then
 hGlobal = GlobalAlloc(, lngLength)
 If (0& <> hGlobal) Then
  lngPointer = GlobalLock(hGlobal)
  If (0& <> lngPointer) Then
  
  'frmEditor.Position.Value = objPosition.CurrentPosition * 100
  'tmp = objPosition.CurrentPosition
  'Debug.Print tmp
  
   Call objBVideo.GetCurrentImage(lngLength, ByVal lngPointer)
   
   'objPosition.CurrentPosition = tmp
   'Debug.Print objPosition.CurrentPosition
   '   objPosition.CurrentPosition = frmEditor.Position.Value / 100
   
   Call GlobalUnlock(hGlobal)
   If (0& = err.Number) Then
    If (True = DIBDataCopy(hGlobal, Pict)) Then
     MPGCaptured = True 'и В DIBDataCopy (но там не всегда отрабатывала проверка из клипборда)
     Set objBVideo = Nothing
     ToDebug "захвачен кадр DShow"
     Exit Sub
    End If
   End If
  End If
  MPGCaptured = False
  'Debug.Print "ошибка захвата кадра (DirectShow)"
  ToDebug "ошибка захвата кадра DShow"
  Set objBVideo = Nothing
  Call GlobalFree(hGlobal)
 End If
End If

End Sub
Public Sub KeyPrev()

lastRendedAVI = m_cAVI.AVIStreamNearestPrevKeyFrame(lastRendedAVI)
'pRenderFrame lastRendedAVI
frmEditor.Position.Value = lastRendedAVI
PosScroll

End Sub
Public Sub KeyNext()

lastRendedAVI = m_cAVI.AVIStreamNearestNextKeyFrame(lastRendedAVI)
frmEditor.Position.Value = lastRendedAVI
PosScroll

End Sub
Public Function DIBDataCopy(ByVal hGlobal As OLE_HANDLE, Pict As PictureBox) As Boolean
'for MPGCaptureBasicVideo
If (0& = hGlobal) Then Exit Function
If (0& <> OpenClipboard()) Then
    If (0& <> EmptyClipboard()) Then
        DIBDataCopy = CBool(0& <> SetClipboardData(vbCFDIB, hGlobal))
    End If
    Call CloseClipboard
End If

If Clipboard.GetFormat(vbCFDIB) Then
    Pict.Picture = Clipboard.GetData
    If Pict.Picture <> 0 Then MPGCaptured = True 'Захватили ок (не полагаться только на клипборд)
'Debug.Print "ok"
End If
End Function



Public Function MediaState() As Integer
Dim lngState As Long
'mzt Dim objBasicVideo As IBasicVideo

On Error Resume Next

Call mobjManager.GetState(30&, lngState)
If (0& <> err.Number) Then Exit Function
Select Case lngState
    Case 1&     'Stopped
        'Set objBasicVideo = mobjManager
        'Set objBasicVideo = Nothing
        MediaState = 1
    Case 2&     'playing
        MediaState = 2
    Case Else     'pause ?
        MediaState = 3
End Select

End Function
Public Sub pRenderFrame(pos As Long)
'movie.Cls
'm_cAVI.FramePicture(pos).Render movie.hdc, 0, 0, 0, 0, 0, 0, 0, 0, 0
'Debug.Print m_cAVI.FrameDuration
'DoEvents

    m_cAVI.DrawFrame frmEditor.movie.hdc, pos, lWidth:=MovieWidth, lHeight:=MovieHeight, Transparent:=False
    'm_cAVI.DrawFrame movie.hDC, pos, lWidth:=ScaleX(MovieWidth, vbPixels, vbTwips), lHeight:=ScaleY(MovieHeight, vbPixels, vbTwips), Transparent:=False
    frmEditor.movie.Refresh

End Sub

Public Sub ComboKey(KeyCode As Integer, Shift As Integer)
'Shift = 1 -> shift
'Shift = 2 -> ctrl
'Shift = 4 -> alt
'устанавливает m_bEditFromCode, чтобы не делать автоселект

'm_bEditFromCode = True если нажаты управляющие кнопки
'Debug.Print KeyCode

'If Shift > 0 Then m_bEditFromCode = True: Exit Sub

Select Case KeyCode
Case 67
    If Shift = 2 Then m_bEditFromCode = True 'ctrl+с
Case vbKeyDelete
    m_bEditFromCode = True
Case vbKeyBack
    m_bEditFromCode = True
Case vbKeyControl
    m_bEditFromCode = True
Case vbKeyShift
    m_bEditFromCode = True
Case 20, 145, 19, 45, 36, 35, 144         'CAPS, SCROLL,
    m_bEditFromCode = True
Case vbKeyLeft, vbKeyRight
    m_bEditFromCode = True
End Select

End Sub
Public Function SavePicFromBase(Base As Integer, dfield As String, Optional myFileName As String) As Boolean
'положить картинку из базы в файл по возможности без пересжатия


Dim fname As String
Dim img As ImageFile
Dim IP As ImageProcess
Dim vec As Vector
Dim WithWIA As Boolean
Dim sType As String
Dim PicSize As Long
Dim b() As Byte
Dim LFile As Integer
'Dim PicExt As String    'разрешение файла в базе

On Error GoTo err

Select Case Base
Case 1
    PicSize = rs.Fields(dfield).FieldSize
    If PicSize = 0& Then Exit Function
    ReDim b(PicSize - 1)
    b() = rs.Fields(dfield).GetChunk(0, PicSize)

Case 2
    PicSize = ars.Fields(dfield).FieldSize
    If PicSize = 0& Then Exit Function
    ReDim b(PicSize - 1)
    b() = ars.Fields(dfield).GetChunk(0, PicSize)
End Select

'определить, что за файл
Set vec = New Vector
vec.BinaryData = b
Set img = vec.ImageFile
Set vec = Nothing

'Select Case img.FormatID
'    'без точки
'Case wiaFormatBMP: PicExt = "bmp"
'Case wiaFormatJPEG: PicExt = "jpg"
'Case wiaFormatGIF: PicExt = "gif"
'Case wiaFormatPNG: PicExt = "png"
'Case wiaFormatTIFF: PicExt = "tif"
'Case Else: PicExt = "pic"                        'для проверки, не рабочее расширение
'End Select

fname = pSaveDialog(FrmMain.hwnd, DTitle:=NamesStore(10), myFileName:=myFileName)

If fname <> vbNullString Then

    WithWIA = True
    Select Case LCase$(right$(fname, 3))
    Case "bmp"
        If img.FormatID = wiaFormatBMP Then 'формат в базе и выбранный совпадают
            WithWIA = False
        Else
            sType = wiaFormatBMP
        End If
    Case "jpg"
        If img.FormatID = wiaFormatJPEG Then 'формат в базе и выбранный совпадают
            WithWIA = False
        Else
            sType = wiaFormatJPEG
        End If
    Case "gif"
        If img.FormatID = wiaFormatGIF Then 'формат в базе и выбранный совпадают
            WithWIA = False
        Else
            sType = wiaFormatGIF
        End If
    Case "png"
        If img.FormatID = wiaFormatPNG Then 'формат в базе и выбранный совпадают
            WithWIA = False
        Else
            sType = wiaFormatPNG
        End If
    Case "tif"
        If img.FormatID = wiaFormatTIFF Then 'формат в базе и выбранный совпадают
            WithWIA = False
        Else
            sType = wiaFormatTIFF
        End If
    End Select
Else
    Exit Function    'нет файла
End If

If WithWIA Then

    Set vec = New Vector
    Set IP = New ImageProcess
    vec.BinaryData = b
    Set img = vec.ImageFile
    Set vec = Nothing

    While (IP.Filters.Count > 0)
        IP.Filters.Remove 1
    Wend

    IP.Filters.Add IP.FilterInfos("Convert").FilterID
    IP.Filters(1).Properties("FormatID").Value = sType     'FormatID 1

    Select Case sType
    Case wiaFormatBMP
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
        ToDebug "Err.SPFB: img Is Nothing"
        Exit Function

    Else
        'img.FileData.BinaryData
        If FileExists(fname) Then Kill fname
        img.SaveFile fname

    End If

Else

    'писать напрямую, выбранные разрешение и тип в базе совпадают
    LFile = FreeFile
    Open fname For Binary As #LFile
    Put #LFile, 1, b()
    Close #LFile

End If    'WithWIA

Erase b
SavePicFromBase = True
Exit Function
err:
ToDebug "Err.SPFB: " & err.Description
Close #LFile

End Function

Public Sub MpegMediaClose()
On Error Resume Next

MpegSizeAdjust True 'очистка окна

'mobjManager.Pause
mobjManager.Stop

'неплохо
If Not (objVideoW Is Nothing) Then Set objVideoW = Nothing
If Not (objPosition Is Nothing) Then Set objPosition = Nothing
If Not (objVideo Is Nothing) Then Set objVideo = Nothing
If Not (objAudio Is Nothing) Then Set objAudio = Nothing

Set mobjSampleGrabber = Nothing

'тормозит при выходе в окно просмотра списка, Set mobjManager = Nothing требует доступность открытого файла (не вынимать диск!)
'если включить, то в автодобавлении стоит .ClearVideo: Clear_mobjManager
'Clear_mobjManager

'If (True <> (mobjManager Is Nothing)) Then
'mobjManager.Stop
'DoEvents
'Set mobjManager = Nothing
'End If
    
MpegMediaOpen = False
ToDebug "VideoClosed"
End Sub
Public Function GetAllVobsSize(p As String, F As String) As String
'p - путь к VTS_YY_0.IFO
'f = VTS_YY_0.IFO
'Передавать только текущий файл!!! f может быть VTS_YY_0.IFO | VTS_YY_0.IFO | VTS_YY_0.IFO - f брать последний!
'надо найти файлы VTS_YY_X.vob

Dim pref As String 'VTS_YY_
Dim WFD As WIN32_FIND_DATA
Dim hfind As Long
Dim found As Boolean
Dim vSize As Double 'сумма длин файлов
Dim fname As String

pref = left$(F, 7) 'VTS_YY_
If right$(p, 1) <> "\" Then p = p & "\"
hfind = FindFirstFile(p & "*", WFD)
found = (hfind > 0)
    
Do While found
 fname = TrimNull(WFD.cFileName)
 If InStr(1, fname, pref, vbTextCompare) > 0 Then
  If StrComp(right$(fname, 3), "vob", vbTextCompare) > -1 Then
    If StrComp(right$(fname, 5), "0.vob", vbTextCompare) <> 0 Then
     vSize = vSize + ((WFD.nFileSizeHigh * MAXDWORD + 1) + WFD.nFileSizeLow) / 1024
    End If
  End If
 End If

found = FindNextFile(hfind, WFD)
Loop
    
GetAllVobsSize = CStr(Int(vSize))
End Function
Public Function AutoFill(Comb As ComboBox) As Boolean
'?True если комбик содержит в списке подходящее значение
' и если
'Dim i As Long, j As Long
'Dim strPartial As String, strTotal As String

'Prevent processing as a result of changes from code
If m_bEditFromCode Then
    m_bEditFromCode = False
    AutoFill = False
    Exit Function
End If

'AutoFill = False
'If Len(Comb.Text) = 0 And Comb.ListCount > 0 Then AutoFill = True
If Comb.ListCount > 0 Then AutoFill = True

'strPartial = Comb.Text
'i = SendMessage(Comb.hwnd, CB_FINDSTRINGEXACT, -1, ByVal strPartial)
'If i <> CB_ERR Then
''если найден точно такой-же, то не краснить - ? а следующий то другой...
'AutoFill = False
'End If

'With Comb
'    'Lookup list item matching text so far
'    strPartial = .Text
'    i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'CB_FINDSTRINGEXACT
'    'If match found, append unmatched characters
'    If i <> CB_ERR Then
'
'AutoFill = True
'
''        'Get full text of matching list item
''        strTotal = .List(i)
''        'Compute number of unmatched characters
''        j = Len(strTotal) - Len(strPartial)
''        If j <> 0 Then
''            'Append unmatched characters to string
''            m_bEditFromCode = True
''            .SelText = Right$(strTotal, j)
''            'Select unmatched characters
''            .SelStart = Len(strPartial)
''            .SelLength = j
''
''        Else
''            '*** Text box string exactly matches list item ***
''
''            'Note: The ListIndex is still -1. If you want to
''            'force the ListIndex to the matching item in the
''            'list, uncomment the following line. Note that
''            'PostMessage is required because Windows sets the
''            'ListIndex back to -1 once the Change event returns.
''            'Also note that the following line causes Windows to
''            'select the entire text, which interferes if the
''            'user wants to type additional characters.
''            '                PostMessage Combo1.hwnd, CB_SETCURSEL, i, 0
''        End If
'    End If
'End With

End Function


Public Sub AutoFillStore()
' в текущем сеансе запоминать вписываемое в редакторе
Dim i As Integer
Dim StoreFlag As Boolean

ToDebug " SaveEditorStore"

With frmEditor

    'название
    If Len(.TextMName.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextMName.ListCount - 1
            If .TextMName.List(i) = .TextMName.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextMName.AddItem .TextMName.Text
    End If

    'метка
    If Len(.TextLabel.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextLabel.ListCount - 1
            If .TextLabel.List(i) = .TextLabel.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextLabel.AddItem .TextLabel.Text
    End If

    'жанр
    If Len(.TextGenre.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextGenre.ListCount - 1
            If .TextGenre.List(i) = .TextGenre.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextGenre.AddItem .TextGenre.Text
    End If

    'произв
    If Len(.TextCountry.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextCountry.ListCount - 1
            If .TextCountry.List(i) = .TextCountry.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextCountry.AddItem .TextCountry.Text
    End If

    'год
    If Len(.TextYear.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextYear.ListCount - 1
            If .TextYear.List(i) = .TextYear.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextYear.AddItem .TextYear.Text
    End If

    'реж
    If Len(.TextAuthor.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextAuthor.ListCount - 1
            If .TextAuthor.List(i) = .TextAuthor.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextAuthor.AddItem .TextAuthor.Text
    End If

    'в ролях
    'StoreFlag = True
    'For i = 0 To TextRole.ListCount
    'If TextRole.List(i) = TextRole.Text Then StoreFlag = False: Exit For
    'Next i
    'If StoreFlag Then TextRole.AddItem TextRole.Text

    'должник
    If Len(.TextUser.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextUser.ListCount - 1
            If .TextUser.List(i) = .TextUser.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextUser.AddItem .TextUser.Text
    End If

    'Примеч TextOther
    If Len(.TextOther.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextOther.ListCount - 1
            If .TextOther.List(i) = .TextOther.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextOther.AddItem .TextOther.Text
    End If

    'Rating
    If Len(.TextRate.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextRate.ListCount - 1
            If .TextRate.List(i) = .TextRate.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextRate.AddItem .TextRate.Text
    End If

    'Lang
    If Len(.TextLang.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextLang.ListCount - 1
            If .TextLang.List(i) = .TextLang.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextLang.AddItem .TextLang.Text
    End If

    'Subt
    If Len(.TextSubt.Text) <> 0 Then
        StoreFlag = True
        For i = 0 To .TextSubt.ListCount - 1
            If .TextSubt.List(i) = .TextSubt.Text Then StoreFlag = False: Exit For
        Next i
        If StoreFlag Then .TextSubt.AddItem .TextSubt.Text
    End If

End With
End Sub

Public Sub Clear_mobjManager()
On Error Resume Next
'тормозит
If (True <> (mobjManager Is Nothing)) Then
    'mobjManager.Stop
    DoEvents
    'Call MpegSizeAdjust(True)
    'mobjManager.RenderFile

    'Dim objRegFilterInfo As IRegFilterInfo
    'Dim objFilterInfo As IFilterInfo
    'Dim GraphFilter As IFilterInfo
    'Dim objPin As IPinInfo
    'Dim SourceFilter As Boolean    'удачно ли прошел первый фильтр Universal Open Source MPEG Source"

    'Set mobjManager = New FilgraphManager
    'mobjManager.FilterCollection.Close
    '.RegFilterCollection = Nothing

    'Set objVideo = Nothing 'mobjManager
    'Set objAudio = Nothing 'mobjManager

    'For Each objRegFilterInfo In mobjManager.RegFilterCollection
    '    Set objRegFilterInfo = Nothing
    'Next

    '       For Each objPin In objFilterInfo.Pins
    '               objPin.Disconnect
    '       Next

    'mobjManager.StopWhenReady
    'If isMPGflag Or isDShflag Then нет

    Dim b As Boolean
    If FileExists(mpgName) Then b = True
    If FileExists(DShName) Then b = True

    If b Then
        Set mobjManager = Nothing
        'Debug.Print Time & " mobjManager = nothing"
        ToDebug "ClearMM_ok"
    Else
        ToDebug "ClearMM_NoFile"
    End If

End If
End Sub

Public Function filltrue(o As Object) As Boolean
'заполнять или нельзя
filltrueAdd = False
filltrue = False
With frmEditor
    If .ChInFilFl.Value = vbUnchecked Then    'разрешили все менять
        filltrue = True
        'Exit Function End With
    ElseIf .ChInFilFl.Value = vbChecked Then    'только если пусто
        If Len(o.Text) = 0 Then filltrue = True
    ElseIf .ChInFilFl.Value = vbGrayed Then    'добавлять
        If Len(o.Text) = 0 Then filltrue = True
        filltrueAdd = True
        'там проверять на filltrueAdd только если filltrue = False
    End If
End With
End Function

Public Sub OpenMovieForCapture(fname As String)
'открыть файл только для возможности снятия скриншотов
'fName - полный путь
Dim FileExt As String
Dim isAVIext As Boolean
Dim fSize As Long
Dim Handle As Long
Dim i As Integer
'Dim tmpdrive As String
'Dim tmpSerial As String
'Dim sVolumeName As String
'Dim temp As String
Dim MMI_Flag As Boolean    'узнал ли файл мминфо модуль


Set frmEditor.movie = Nothing

If Not Opt_AviDirectShow Then
    '                                                           AVI
    'если размер не 0
    If isWindowsNt Then
        Dim Pointer As Long, lpFSHigh As Currency
        Pointer = lopen(fname, OF_READ)
        GetFileSizeEx Pointer, lpFSHigh
        fSize = Int(lpFSHigh * 10000 / 1024)
        lclose Pointer
    Else
        fSize = Int(FileLen(fname) / 1024)
    End If
    If fSize = 0 Then
        ToDebug "err_нулевая длина: " & fname
        'Mark2SaveFlag = False
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    'раcширения avi
    FileExt = LCase$(getExtFromFile(fname))
    Select Case FileExt
    Case "avi"
        isAVIext = True
    Case "vid"
        isAVIext = True
    Case "divx"
        isAVIext = True
    Case "xvid"
        isAVIext = True
    End Select

    ToDebug "OpenFileExt=" & FileExt

    If isAVIext Then
        'если разширения подходят для avi
        Call PrepareAviForCupture(fname)
    End If

    If isAVIflag Then
        If aferror Then
            ToDebug "ssAVI, cupture error"
        Else
            OpenAddmovFlag = True
            ToDebug "AVI ready for cupture"
        End If
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

End If        'not Opt_AviDirectShow

'
'                                                                       mpeg?
On Error Resume Next
Handle = MediaInfo_New()
If err Then
    '            If Not AutoNoMessFlag Then
    myMsgBox msgsvc(51) & App.Path & "\MediaInfo.dll", vbCritical
    '            End If
    '    MsgBox "MediaInfo.dll not found. Reinstall SurVideoCatalog!", vbCritical
    ToDebug "Err_NoDLL: MediaInfo.dll"
    Screen.MousePointer = vbNormal
    Exit Sub
End If

Call MediaInfo_Open(Handle, StrPtr(fname))

If err = 0 Then
    i = MediaInfo_Count_Get(Handle, 1, -1)
    If i > 0 Then        'есть видео
        MMI_Flag = True
        MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))
        Select Case strMPEG
            Case "MPV1", "MPV2"
                isMPGflag = True
        End Select
    End If
    MediaInfo_Close Handle
Else
    err.Clear
End If
On Error GoTo 0

If isMPGflag Then
Else        '                                                          не MPV
    ''                                                             Обработка DS

    ToDebug "DShow: " & fname
    DShName = fname
    mpgName = fname
    '                                          DS INFO
    If DShGetInfoForCapture Then
        isDShflag = True
        OpenAddmovFlag = True
    Else
        isDShflag = False
    End If

    mpgName = vbNullString
End If

If isMPGflag Or isDShflag Then
Else
    ToDebug "err: Не поддерживается или уже открыт: " & fname
    If Not AutoNoMessFlag Then
        myMsgBox msgsvc(13) & fname, vbInformation, , FrmMain.hwnd
    End If
End If

'                                                               Обработка MPEG
If isMPGflag Then
    ToDebug "MPEG: " & fname
    mpgName = fname
    '                                               MPG INFO
    If MpgGetInfoForCapture Then
        isMPGflag = True
        OpenAddmovFlag = True
    Else
        isMPGflag = False
    End If
End If

Screen.MousePointer = vbNormal

On Error Resume Next: If Not frmAutoFlag Then FrmMain.SetFocus


End Sub

Public Function OpenNewMovie(Optional mn As String) As Boolean
'еще OpenMovieForCapture
Dim tmp As String    'полученное имя файла
Dim AVIInf As New clsAVIInfo
Dim Handle As Long
Dim i As Integer
Dim tmpdrive As String
Dim tmpSerial As String
Dim sVolumeName As String
Dim temp As String
Dim MMI_Flag As Boolean    'узнал ли файл мминфо модуль
Dim fSize As Long

ToDebug "OpenNew..."
OpenNewMovie = True    'если нет, укажем это потом
'Mark2SaveFlag = False 'не краснить save
NewDiskAddFlag = False


frmEditor.TxtIName.Text = vbNullString

If Len(mn) <> 0 Then
    tmp = mn    'передали уже имя файла
Else
    If Not AppendMovieFlag Then

        'Открыть новый фильм
    
        If Not FileExists(tmp) Then
            frmEditor.FrameAddEdit.Enabled = False    'чтоб не нажимались батоны
            'FrmMain.VerticalMenu.Enabled = False
            tmp = pLoadDialog     '(ComOpen.Caption)
            DoEvents
            'FrmMain.VerticalMenu.Enabled = True
            frmEditor.FrameAddEdit.Enabled = True
        End If

        If Len(tmp) = 0 Then
            Screen.MousePointer = vbNormal
            OpenNewMovie = False
            'Mark2SaveFlag = False
            ToDebug "...Cancel"
            Exit Function
        End If
    Else
        'добавление
        'взять текущее имя ави для добавления
        tmp = aviName
    End If
End If

DoEvents
ToDebug "NewMovie " & tmp

If MpegMediaOpen Then MpegMediaClose

'ClearVideo 'nah
'If addFlag Then ClearFields - а если сначала вводили тексты
Set frmEditor.movie = Nothing

'Auto name in inet name
temp = GetNameFromPathAndName(tmp)
GetExtensionFromFileName temp, temp
If Len(frmEditor.TxtIName.Text) = 0 Then frmEditor.TxtIName.Text = LCase$(temp)

If Not AppendMovieFlag Then    'если добавление - то в add
    'Серийник. CDSerialCur-поле
    frmEditor.CDSerialCur = vbNullString    ': SameCDLabel = vbNullString
    tmpdrive = left$(LCase$(tmp), 3)
    If DriveType(tmpdrive) = "CD-ROM" Then
        tmpSerial = Hex$(GetSerialNumber(tmpdrive, sVolumeName))
        If tmpSerial <> "0" Then    'надо

            If MediaSN = tmpSerial Then
                IsSameCdFlag = True    'это тот же носитель
            Else
                IsSameCdFlag = False
                MediaSN = tmpSerial    'запомнить серийник cd
            End If

            If CheckSameDisk Then
                'If SearchLVSimple(dbsnDiskInd, tmpSerial) Then    'есть ли уже в базе
                If SearchSNinbase(tmpSerial) Then    'есть ли уже в базе, там меняется SameCDLabel

                    If Not AutoNoMessFlag Then
                        If myMsgBox(msgsvc(28), vbYesNo, App.title, FrmMain.hwnd) = vbNo Then    'продолжить?
                            Set AVIInf = Nothing
                            OpenNewMovie = False
                            Screen.MousePointer = vbNormal
                            Exit Function
                        Else
                            CheckSameDisk = False    ' AutoNoMessFlag = True 'не спрашивать боле
                            frmEditor.CDSerialCur = tmpSerial
                            frmEditor.TextLabel = SameCDLabel
                        End If
                    Else    'без вопросов
                        frmEditor.CDSerialCur = tmpSerial
                        frmEditor.TextLabel = SameCDLabel
                    End If
                Else
                    'CheckSameDisk = False 'не нашли и не искать боле
                    SameCDLabel = vbNullString

                    frmEditor.CDSerialCur = tmpSerial
                    If Len(frmEditor.TextLabel) = 0 Then frmEditor.TextLabel = sVolumeName
                End If
            Else
                frmEditor.CDSerialCur = tmpSerial
                If Len(frmEditor.TextLabel) = 0 Then
                    If Len(SameCDLabel) = 0 Then
                        frmEditor.TextLabel = sVolumeName
                    Else
                        frmEditor.TextLabel = SameCDLabel
                    End If
                End If
            End If
        End If

        'Носитель
        If IsSameCdFlag Then    'если тот же носитель
            ToDebug "Носитель тот-же."
            frmEditor.ComboNos.Text = MediaType
        Else
            'определить тип носителя
            frmEditor.ComboNos.Text = GetOptoInfo(left$(tmp, 2))
            MediaType = frmEditor.ComboNos.Text
        End If

    Else
        If Opt_GetMediaType Then frmEditor.ComboNos.Text = DriveType(tmpdrive)
        tmpSerial = Hex$(GetSerialNumber(tmpdrive))
        If tmpSerial <> "0" Then frmEditor.CDSerialCur = tmpSerial
    End If    'drivetype
End If    'append


Screen.MousePointer = vbHourglass

'определить что за файл
isMPGflag = False: isAVIflag = False: isDShflag = False

OpenAddmovFlag = True

If Not Opt_AviDirectShow Then
    '                                                           AVI
    aviName = tmp

    'если размер не 0
    If isWindowsNt Then
        Dim Pointer As Long, lpFSHigh As Currency
        Pointer = lopen(tmp, OF_READ)
        GetFileSizeEx Pointer, lpFSHigh
        fSize = Int(lpFSHigh * 10000 / 1024)
        lclose Pointer
    Else
        fSize = Int(FileLen(aviName) / 1024)
    End If

    If fSize = 0 Then
        OpenNewMovie = False
        ToDebug "err: нулевая длина: " & aviName
        'Mark2SaveFlag = False
        Screen.MousePointer = vbNormal
        Exit Function
    End If

    'раcширения avi
    Dim FileExt As String
    Dim isAVIext As Boolean
    FileExt = LCase$(getExtFromFile(aviName))
    Select Case FileExt
    Case "avi"
        isAVIext = True
    Case "vid"
        isAVIext = True
    Case "divx"
        isAVIext = True
    Case "xvid"
        isAVIext = True
    End Select

    ToDebug "Open FileExt=" & FileExt
    '                                                               GetAviInfo
    If isAVIext Then GetAviInfo    'если разширения подходят для avi

    If isAVIflag Then
        If aferror Then
            ToDebug "VFW, Ошибка захвата кадра"
        Else
            ToDebug "VFW ok."
        End If
        Set AVIInf = Nothing
        Screen.MousePointer = vbNormal
        'Mark2SaveFlag = False
        Exit Function
    End If

End If    'not Opt_AviDirectShow

'                                                                       mpeg1/2?
On Error Resume Next
Handle = MediaInfo_New()
If err Then
    If Not AutoNoMessFlag Then
        myMsgBox msgsvc(51) & App.Path & "\MediaInfo.dll", vbCritical
    End If
    '    MsgBox "MediaInfo.dll not found. Reinstall SurVideoCatalog!", vbCritical
    ToDebug "Err_NoDLL: MediaInfo.dll"
    OpenNewMovie = False
    Screen.MousePointer = vbNormal
    Exit Function
End If

Call MediaInfo_Open(Handle, StrPtr(tmp))

If err = 0 Then
    i = MediaInfo_Count_Get(Handle, 1, -1)
    If i > 0 Then    'есть видео
        MMI_Flag = True
        MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))
        Select Case strMPEG
        Case "MPV1", "MPV2"
            isMPGflag = True
        End Select
    End If
    MediaInfo_Close Handle
Else
    err.Clear    ':On Error GoTo 0
End If

On Error GoTo 0

If isMPGflag Then
Else    '                                                          не MPV

    ''                                                               Обработка DS
    ''If isDShflag Then
    ToDebug "Info DShow файла: " & tmp
    DShName = tmp
    mpgName = tmp
    If DShGetInfo Then
        isDShflag = True
    Else
        OpenNewMovie = False
        isDShflag = False
    End If
    ''End If

    '    'срендерить автоматом
    '    mpgName = tmp
    ''исправить : DShGetInfo должна итти до RenderAuto!!! иначе не видит аспект
    '    If RenderAuto Then
    '        isDShflag = True
    '    End If

    mpgName = vbNullString
End If

If isMPGflag Or isDShflag Then
Else
    ToDebug "err: Не поддерживается или уже открыт: " & tmp
    OpenNewMovie = False
    If Not AutoNoMessFlag Then
        myMsgBox msgsvc(13) & tmp, vbInformation, , FrmMain.hwnd    'Этот файл не поддерживается или уже открыт
    End If
End If


'                                                               Обработка MPEG
If isMPGflag Then
    ToDebug "Обработка файла MPEG: " & tmp
    mpgName = tmp
    If Not MpgGetInfo Then OpenNewMovie = False
End If
'If MPGCaptured = False Then ToDebug "Ошибка: файл загружен, но есть ошибка захвата кадра: " & mpgName

''                                                               Обработка DS
'If isDShflag Then
'    ToDebug "Обработка DShow файла: " & tmp
'    DShName = tmp
'    If Not DShGetInfo Then OpenNewMovie = False
'End If

Screen.MousePointer = vbNormal

'On Error Resume Next : If Not frmAutoFlag Then FrmMain.SetFocus

End Function

Public Function MpgGetInfo() As Boolean
'MpgGetInfoForCapture
'это выполняется, когда mminfo опознал файл (нашел там видеопоток) и файл MPV

Dim WeGetTimeFromIfo As Boolean

Dim i As Integer, j As Integer
Dim temp As Currency    'Long
Dim Handle As Long
Dim TimeS As String
Dim tmps As String, tmp2s As String
Dim tmp As String, tmp2 As String
Dim tmpL As Long
Dim tmpa As String    'битр аудио
Dim MMI_Height As Integer    'из MMInfo
Dim MMI_Width As Integer
Dim objv_Height As Integer    'из objVideo.Source
Dim objv_Width As Integer

Dim ret As Long
Dim WFD As WIN32_FIND_DATA

'Dim ifo_fname As String
Dim ifo_handle As Long
Dim length_pgc(3) As Byte
Dim NumIFOChains As Long
Dim NumCells As Long

Dim AllVobsSize As String
Dim Pointer As Long, lpFSHigh As Currency
Dim rendMPV2 As Boolean    'срендерили как мпег2 с нашими фильтрами

Dim vstr_NumSubPic As Long
Dim IsVob As Boolean

Dim sAsp As String

With frmEditor

    MpgGetInfo = True

    If Not AppendMovieFlag Then
        .TextTimeHid = vbNullString
        .TextFPSHid = vbNullString
    End If

    'поиск файла ifo для данного воба
    If StrComp(right$(mpgName, 3), "vob", vbTextCompare) = 0 Then    'vob
        IsVob = True
        tmps = left$(mpgName, Len(mpgName) - 5) & "0.ifo"
        ret = FindFirstFile(tmps, WFD)
        If ret > 0 Then
            'open
            On Error Resume Next
            ifo_handle = ifoOpen(tmps, fio_USE_ASPI)
            If ifo_handle = 0 Then
                ToDebug ("ERR_vstrip.dll, не открыть: " & tmps)
                If Not AutoNoMessFlag Then
                    myMsgBox msgsvc(51) & App.Path & "\vstrip.dll", vbCritical
                End If
            End If
            On Error GoTo 0
        End If
    End If

    Handle = MediaInfo_New()
    Call MediaInfo_Open(Handle, StrPtr(mpgName))
    ToDebug "MediaInfo: " & bstr(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr("")))

    If ifo_handle = 0 Then

        If Not AppendMovieFlag Then
            'инфа из файла

            On Error Resume Next

            'звук                           звук
            .TextAudioHid = vbNullString
            tmpL = MediaInfo_Count_Get(Handle, MediaInfo_Stream_Audio, -1)
            If tmpL > 0 Then
                For i = 0 To tmpL - 1
                    tmps = vbNullString
                    tmps = bstr(MediaInfo_Get(Handle, 2, i, StrPtr("SamplingRate"), 1, 0)) & " "
                    tmps = tmps & Chnnls(bstr(MediaInfo_Get(Handle, 2, i, StrPtr("Channels"), 1, 0))) & " "
                    tmps = tmps & bstr(MediaInfo_Get(Handle, 2, i, StrPtr("Codec"), 1, 0)) & " "
                    
                    tmp = bstr(MediaInfo_Get(Handle, 2, i, StrPtr("BitRate"), 1, 0))
                    tmp = Replace(tmp, ".", ",")
                    If IsNumeric(tmp) Then
                        If Val(tmp) > 0 Then
                            tmp = tmp / 1000 & "kbps" 'звук
                            tmps = tmps & "(" & tmp & ")"
                        End If
                    End If

                    .TextAudioHid = tmps & ", " & .TextAudioHid
                Next i
                .TextAudioHid = left$(.TextAudioHid, Len(.TextAudioHid) - 2)
                .TextAudioHid = Trim$(.TextAudioHid)

            Else
                ToDebug "MMInfo не нашел звук"
            End If
            err.Clear: On Error GoTo 0

        End If    'append

        'видео mminfo
        '                           аспект
        If MediaInfo_Count_Get(Handle, MediaInfo_Stream_Video, -1) > 0 Then

            'строка 4/3  mminfo
            MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
            sAsp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))

            If Len(MMI_Format_str) = 0 Then
                MMI_Format_str = "4/3"
                ToDebug "MMInfo не нашел Format. = " & MMI_Format_str
            ElseIf MMI_Format_str = "0.000" Then
                MMI_Format_str = "4/3"
                ToDebug "MMInfo не дал Format. = " & MMI_Format_str
                
            Else
                ToDebug "MMInfo Format = " & MMI_Format_str
            End If
            'MMI_Format = CalcFormat(bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0)))
            MMI_Format = CalcFormat(MMI_Format_str, sAsp)

            MMI_Height = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Height"), 1, 0))
            MMI_Width = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Width"), 1, 0))
            'AspectRatio = GetAspectRatio(AspectRatioS, HeightS)
            tmps = vbNullString
            'Call MediaInfo_Option(Handle, StrPtr("Complete"), StrPtr("1"))
            MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))
            tmps = MyCodec(MPGCodec) & " "


            If Not AppendMovieFlag Then
                'fps  mminfo
                'tmps = tmps & AspectRatioS & " "
                tmp2s = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("FrameRate"), 1, 0))
                .TextFPSHid = Val(tmp2s)

                Select Case Int(Val(tmp2s))
                Case "25"
                    tmps = tmps & "PAL" & " "
                Case "29"    '29.97
                    tmps = tmps & "NTSC" & " "
                Case "23"    '23.976
                    tmps = tmps & "FILM" & " "
                Case Else
                    tmps = tmps & tmp2s & " "
                End Select

                'добавить аспект
                tmps = tmps & MMI_Format_str & " "
                'и битрейт видео
                tmp = vbNullString
                tmp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("BitRate"), 1, 0))
                tmp = Replace(tmp, ".", ",")
                If IsNumeric(tmp) Then
                    If Val(tmp) > 0 Then
                        'tmp = tmp / 1000 & "kbps"
                        tmp = Format$(tmp / 1000, "0") & "kbps"
                        tmps = tmps & "(" & tmp & ")"
                    End If
                End If
                tmps = Replace(tmps, "16/9", "16:9")
                tmps = Replace(tmps, "4/3", "4:3")
                tmps = Replace(tmps, " 0.000", vbNullString)
                .TextVideoHid = RTrim$(tmps)
            End If

        End If    'append

        'MediaInfo_Close Handle

    Else                                                   'инфу с ifo
        ToDebug "vstrip: " & tmps

        If Not AppendMovieFlag Then
            'video ifo
            .TextVideoHid = Trim$(IFObstr(ifoGetVideoDesc(ifo_handle)))

            'bitrate ifo video
            '        Handle = MediaInfo_New(): Call MediaInfo_Open(Handle, StrPtr(mpgName))
            tmp = vbNullString
            tmp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("BitRate"), 1, 0))
            tmp = Replace(tmp, ".", ",")
            If IsNumeric(tmp) Then
               ' tmp = tmp / 1000 & "kbps"
                tmp = Format$(tmp / 1000, "0") & "kbps"
            End If
            If Len(tmp) > 0 Then .TextVideoHid = .TextVideoHid & " (" & tmp & ")"

            'fps ifo
            .TextFPSHid = vbNullString
            If InStr(1, .TextVideoHid, "PAL", vbTextCompare) > 0 Then .TextFPSHid = "25"
            If InStr(1, .TextVideoHid, "NTSC", vbTextCompare) > 0 Then .TextFPSHid = "29.97"

            'audio & Language ifo
            j = ifoGetNumAudio(ifo_handle)
            If j > 0 Then
                tmps = vbNullString: tmp2s = vbNullString
                For i = 0 To j - 1
                    tmp = vbNullString
                    tmp = IFObstr(ifoGetAudioDesc(ifo_handle, i))
                    'bitrate audio
                    tmpa = vbNullString
                    tmpa = bstr(MediaInfo_Get(Handle, MediaInfo_Stream_Audio, i, StrPtr("BitRate"), 1, 0))
                    tmpa = Replace(tmpa, ".", ",")
                    If IsNumeric(tmpa) Then tmpa = tmpa / 1000 & "kbps"
                    If Len(tmpa) > 0 Then tmp = tmp & " (" & tmpa & ")"

                    'Debug.Print "VStrip Audio: " & tmp
                    'Debug.Print "Lang: " & IFObstr(ifoGetLangDesc(ifo_handle, i))
                    tmp2 = IFObstr(ifoGetLangDesc(ifo_handle, i))
                    'tmp = Replace(tmp, "[0]", vbNullString)
                    'tmp = Left$(tmp, InStr(1, tmp, "ch,", vbTextCompare) + 1) & ")"
                    tmps = tmps & tmp & ", "
                    tmp2s = tmp2s & tmp2 & ", "
                Next i
                'Lang
                If Len(tmps) > 2 Then .TextAudioHid = Trim$(left$(tmps, Len(tmps) - 2))
                If Len(tmp2s) > 2 Then tmp2s = Trim$(left$(tmp2s, Len(tmp2s) - 2))


                If filltrue(.TextLang) Then
                    .TextLang = CountryLocal(tmp2s)
                ElseIf filltrueAdd Then
                    .TextLang = .TextLang & ", " & CountryLocal(tmp2s)
                End If
            End If

            'SubTitle ifo
            tmp = vbNullString
            vstr_NumSubPic = ifoGetNumSubPic(ifo_handle)
            For i = 0 To vstr_NumSubPic - 1
                If i > 0 Then
                    tmp = tmp & ", " & IFObstr(ifoGetSubPicDesc(ifo_handle, i))
                Else
                    tmp = IFObstr(ifoGetSubPicDesc(ifo_handle, i))
                End If
            Next i

            If filltrue(.TextSubt) Then
                .TextSubt = CountryLocal(tmp)
            ElseIf filltrueAdd Then
                .TextSubt = .TextSubt & ", " & CountryLocal(tmp)
            End If

        End If    'not append

        '                                                           время фильма с ifo
        NumIFOChains = ifoGetNumPGCI(ifo_handle)

        If AppendMovieFlag Then .TextTimeHid = .TextTimeHid & ", "    'потом, если возможно, суммируется, если нет  (1 ифо на несколько фильмов ?) то добавиться

        For i = 0 To NumIFOChains - 1
            NumCells = ifoGetPGCIInfo(ifo_handle, i, length_pgc(0))    ' нужен length_pgc(0)
            .TextTimeHid = .TextTimeHid & Format$(length_pgc(0), "0#") & ":" & Format$(length_pgc(1), "0#") & ":" & Format$(length_pgc(2), "0#") & ", "    ' & "," & Format$(length_pgc(3), "0#")
            'Debug.Print FormatTime(TextTimeMSHid)
        Next i

        'заменим TextTimeHid, если можно суммировать
        If i > 1 Then    '1 ифо на несколько фильмов
            TextTimeMSHid = 0    'чтобы не суммировать потом DS
            WeGetTimeFromIfo = True
            'нажать плюсик
            Call .ComPlusHid_Click(2)
        Else
            If Not AppendMovieFlag Then
                .TextTimeHid = left$(.TextTimeHid, Len(.TextTimeHid) - 2)
                TextTimeMSHid = length_pgc(0) * 3600 + length_pgc(1) * 60 + length_pgc(2)
                WeGetTimeFromIfo = True
            Else
                'суммировать
                TextTimeMSHid = TextTimeMSHid + length_pgc(0) * 3600 + length_pgc(1) * 60 + length_pgc(2)
                .TextTimeHid = FormatTime(TextTimeMSHid)
                WeGetTimeFromIfo = True
            End If
        End If
        'убрать последние ", "
        If right$(.TextTimeHid, 2) = ", " Then .TextTimeHid = left$(.TextTimeHid, Len(.TextTimeHid) - 2)

        'close ifo
        If ifoClose(ifo_handle) Then ToDebug "ifo закрыт."

        If Not AppendMovieFlag Then
            If Len(.TextFPSHid) = 0 Then    'если попытка с ifo не прошла
                'fps mminfo
                .TextFPSHid = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("FrameRate"), 1, 0))
                'HeightS = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Height"), 1, 0))
                'WidthS = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Width"), 1, 0))
            End If
        End If

        MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))

        '    MMI_Ratio = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))
        '    If Len(MMI_Ratio) = 0 Then MMI_Ratio = "1.333" 'было 1

        'Debug.Print "MMI_Format=" & bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
        '4/3=?
        'MMI_Format = 1

        MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
        sAsp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))
        If Len(MMI_Format_str) = 0 Then
            MMI_Format_str = "4/3"
            ToDebug "MMInfo не нашел Format. = " & MMI_Format_str
        ElseIf MMI_Format_str = "0.000" Then
            MMI_Format_str = "4/3"
            ToDebug "MMInfo не дал Format. = " & MMI_Format_str
      
        Else
            ToDebug "MMInfo Format = " & MMI_Format_str
        End If
        MMI_Format = CalcFormat(MMI_Format_str, sAsp)

        MMI_Height = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Height"), 1, 0))
        MMI_Width = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Width"), 1, 0))

    End If    ' ifo

    If Handle <> 0 Then MediaInfo_Close Handle

    'Debug.Print "MMI_Ratio=" & MMI_Ratio
    'ToDebug "MMI_Ratio=" & MMI_Ratio
    'Debug.Print "MMI_Width=" & MMI_Width & " MMI_Height=" & MMI_Height

    ToDebug "MMI_Format=" & MMI_Format
    'взять MMI_Format из поля видео в редакторе, если нет - оставить текущий
    MMI_Format = GetAspectFromTextVideo
    'ToDebug "Ifo Format=" & MMI_Format
    ToDebug "Use Format=" & MMI_Format

End With

'****************************************************DIRECT X *********************



SVCDflag = False
If Not IsVob Then
    If MMI_Width = 480 And MMI_Height = 576 Then SVCDflag = True
    If MMI_Width = 480 And MMI_Height = 480 Then SVCDflag = True
End If
ToDebug "SVCD = " & SVCDflag
ToDebug "MMI_Codec: " & MPGCodec

Select Case strMPEG
Case "MPV2"
    If Opt_UseOurMpegFilters And Not SVCDflag Then        'нашим декодером
        ToDebug "Try RenderMPV2..."
        If RenderMPV2 Then    '                                   -----    RenderMPV2
            rendMPV2 = True
        Else
            ToDebug "Try RenderAuto..."
            If Not RenderAuto Then  '                            ----- RenderAuto
                'все плохо

                If Not AutoNoMessFlag Then
                    myMsgBox msgsvc(10) & mpgName    'Ошибка работы с файлом
                End If
                MpgGetInfo = False
                Exit Function
            End If
        End If
    Else        'сразу авто (svcd или не наши фильтры)
        ToDebug "сразу RenderAuto..."
        If Not RenderAuto Then    '                               ----- RenderAuto
            'все плохо
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(10) & mpgName
            End If
            MpgGetInfo = False
            Exit Function
        End If
    End If
Case "MPV1"
    If RenderMPV1 Then    '                               ----- RenderMPV1
        'rendMPV1 = True
    Else
        If Not RenderAuto Then    '                        ----- RenderAuto
            'все плохо
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(10) & mpgName
            End If
            MpgGetInfo = False
            Exit Function
        End If
    End If
Case Else
    MpgGetInfo = False
    Exit Function
End Select

On Error Resume Next
Set objPosition = mobjManager
objPosition.CurrentPosition = 5#
Set objVideo = mobjManager

With frmEditor
    'тестовый запуск
    .FrAdEdPixHid.Visible = True
    ToDebug "Run..."
    mobjManager.Stop
    mobjManager.Run
    'Sleep 300
    mobjManager.Pause
    ToDebug "Pause. Err=" & err.Number

    If AutoShots Then
        'проба захвата кадра
        'DoEvents
        MPGCaptureBasicVideo FrmMain.PicTempHid(0)
        If MPGCaptured = False Then
            'ToDebug "Ошибка: файл загружен, но есть ошибка захвата кадра: " & mpgName
            If Not AutoNoMessFlag Then myMsgBox msgsvc(38) & vbCrLf & mpgName
        End If
    End If

    err.Clear: On Error GoTo 0


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If StrComp(right$(mpgName, 3), "vob", vbTextCompare) = 0 Then    'сменить vob на ifo
        tmps = left$(mpgName, Len(mpgName) - 5) & "0.ifo"
        ret = FindFirstFile(tmps, WFD)
        If ret > 0 Then
            ToDebug "Структура DVD."
            'имена файлов
            If Not AppendMovieFlag Then
                .TextFileName.Text = GetEdTxtFileName(tmps)    'ifo
            Else
                .TextFileName.Text = .TextFileName.Text & " | " & GetEdTxtFileName(tmps)    'ifo
            End If
            'найти все вобы фильма и их размер
            AllVobsSize = GetAllVobsSize(GetPathFromPathAndName(tmps), GetNameFromPathAndName(tmps))    'нет! TextFileName.Text)
            ToDebug "VobsSize=" & AllVobsSize
        Else
            If Not AppendMovieFlag Then
                .TextFileName.Text = GetEdTxtFileName(mpgName)
            Else
                .TextFileName.Text = .TextFileName.Text & " | " & GetEdTxtFileName(mpgName)
            End If
        End If
    Else
        If Not AppendMovieFlag Then
            .TextFileName.Text = GetEdTxtFileName(mpgName)
        Else
            .TextFileName.Text = .TextFileName.Text & " | " & GetEdTxtFileName(mpgName)
        End If
    End If
    
    ToDebug "Editor/FileName=" & .TextFileName.Text

    'показывает из настроек декодера, те может быть неверно (50 вместо 25...)
    'TextFPSHid = Round(1 / objVideo.AvgTimePerFrame, 3)
    'FrameS = 1 / objVideo.AvgTimePerFrame * objPosition.Duration
    'If FrameS < 1 Then Exit Sub  'no frames

    On Error Resume Next    'automation error на нулевых вобах

    'Time dx - 1 файл
    If Not WeGetTimeFromIfo Then
        If Not AppendMovieFlag Then
            '    If Len(TextTimeHid) = 0 Then    'не узнали из ifo
            TextTimeMSHid = objPosition.Duration
            TimeS = FormatTime(TextTimeMSHid)    'mminfo.bas
            .TextTimeHid = TimeS
            '    End If
        Else
            'суммировать при добавлении если было к чему
            If TextTimeMSHid <> 0 Then
                TextTimeMSHid = TextTimeMSHid + objPosition.Duration
                .TextTimeHid = FormatTime(TextTimeMSHid)
            End If
        End If
    End If

    TimeL = objPosition.Duration
End With

'Debug.Print "TimeL = " & TimeL
If (TimeL = 0) Or (TimeL > 10000000) Then
    'ClearVideo
    ToDebug msgsvc(40) & " = " & TimeL
    If Not AutoNoMessFlag Then myMsgBox msgsvc(40) & " = " & TimeL, vbCritical, , FrmMain.hwnd
    MpgGetInfo = False
    Exit Function
End If

With frmEditor
    'file size
    If Len(AllVobsSize) = 0 Then
        'посчитать размер файла
        If isWindowsNt Then
            Pointer = lopen(mpgName, OF_READ)
            GetFileSizeEx Pointer, lpFSHigh
            If Not AppendMovieFlag Then
                .TextFilelenHid = Int(lpFSHigh * 10000 / 1024)
            Else
                .TextFilelenHid = Val(.TextFilelenHid) + Int(lpFSHigh * 10000 / 1024)
            End If
            lclose Pointer
        Else
            If Not AppendMovieFlag Then
                .TextFilelenHid = Int(FileLen(mpgName) / 1024)
            Else
                .TextFilelenHid = Val(.TextFilelenHid) + Int(FileLen(mpgName) / 1024)
            End If
        End If
    Else
        'или подставить размер всех вобов
        If Not AppendMovieFlag Then
            .TextFilelenHid = AllVobsSize
        Else
            .TextFilelenHid = Val(.TextFilelenHid) + Val(AllVobsSize)
        End If
    End If

    '                                                                 Frame Size
    objv_Width = objVideo.SourceWidth
    objv_Height = objVideo.SourceHeight

    If Not AppendMovieFlag Then
        .TextResolHid = Trim$(str$(objv_Width)) & " x " & Trim$(str$(objv_Height))
    End If

    '                                                                   aspect
    'PixelRatio = 1.333: PixelRatioSS = ScrShotEd_W / (4 / 3)    'дефолт 4:3
    'PixelRatioSS = ScrShotEd_W / (4 / 3)    'дефолт 4:3

    'то же MpgGetInfoForCapture
    PixelRatio = (objv_Height * MMI_Format) / objv_Width
    PixelRatioSS = ScrShotEd_W / MMI_Format


    'Debug.Print "PixelRatio=" & PixelRatio, "PixelRatioSS=" & PixelRatioSS
    'ToDebug "AspectFull=" & PixelRatio & " AspectMini=" & PixelRatioSS
    'TextResolHid = WidthS & " x " & HeightS

    'cds
    If Not AppendMovieFlag Then
        .TextCDN.Text = 1
    Else
        If NewDiskAddFlag Then
            .TextCDN.Text = Replace(.TextCDN, Val(.TextCDN), Val(.TextCDN) + 1)
        End If
    End If

    TimesX100 = TimeL * 100

    temp = objVideo.AvgTimePerFrame * 100

    .Position.min = 0
    .Position.Max = TimesX100    '- temp
    .Position.Value = 0

    .Position.TickFrequency = .Position.Max / 100
    .Position.SmallChange = temp * 100    '0.04*100 *1000
    .Position.LargeChange = temp * 1000

    PPMax = .Position.Value + cMPGRange    ' Const Range As Integer = 10000 в MpegPosScroll
    If PPMax > TimesX100 Then PPMax = .Position.Max    ' TimesX100

    .PositionP.min = 0    'PPMin
    .PositionP.TickFrequency = temp
    .PositionP.Max = PPMax
    .PositionP.SmallChange = temp
    .PositionP.LargeChange = temp * 10
    .PositionP.Value = 0

    .Position.Enabled = True: .PositionP.Enabled = True

    .ComKeyAvi(0).Enabled = False: .ComKeyAvi(1).Enabled = False

    If MPGCaptured Then
        For i = 0 To 2: .ComCap(i).Enabled = True: .ComRND(i).Enabled = True: Next i
        .ComAutoScrShots.Enabled = True

        'Отразить аспекты на кнопках                                                    4:3 16:9
        Call EdAspect2Buttons

    Else
        For i = 0 To 2: .ComCap(i).Enabled = False: .ComRND(i).Enabled = False: Next i
        .ComAutoScrShots.Enabled = False
        For i = 0 To 2: .optAspect(i).Enabled = False: Next i
    End If

    'ComAdd.Enabled = False
    .ComAdd.Enabled = True
End With
End Function

Public Sub EdAspect2Buttons()
'не должно помечать , если иной формат
'может исп MMI_Format сразу, а то бывает что используем формат MMI_Format а галочки стоят по MMI_Format_str
Dim i As Integer
If Len(MMI_Format_str) = 0 Then
    If MMI_Format = 1.333 Then
        MMI_Format_str = "4/3"
    ElseIf MMI_Format = 1.777 Then
        MMI_Format_str = "16/9"
        '    ElseIf MMI_Format = 1 Then
        '        MMI_Format_str = "1/1"
    Else
        MMI_Format_str = "4/3"
    End If
End If

'подгонять под mmi_format
With frmEditor

    For i = 0 To 2: .optAspect(i).Value = False: Next i

    ChangeFromCode_optAspect = True
    Select Case Trim$(MMI_Format_str)
    Case "4/3", "4:3"
        .optAspect(0).Value = True
    Case "16/9", "16:9"
        .optAspect(1).Value = True
        'Case "1/1", "1:1"
        '    optAspect(2).Value = True
    End Select

    If Not Opt_UseAspect Then    'если не юзать аспект, то не показывать кнопки
        'optAspect(2).Value = True 'w:h
        For i = 0 To 2: .optAspect(i).Enabled = False: Next i
    Else
        For i = 0 To 2: .optAspect(i).Enabled = True: Next i
    End If
    ChangeFromCode_optAspect = False

End With
End Sub

Public Function RenderAuto() As Boolean
'Рендер как есть автоматом
Dim GraphFilter As IFilterInfo
'Dim imEvent As IMediaEvent
Dim objRegFilterInfo As IRegFilterInfo
Dim objFilterInfo As IFilterInfo
Dim udtMediaType As TAMMediaType
'Dim objPinInfo As IPinInfo

If MpegMediaOpen Then
    MpegMediaClose
    Clear_mobjManager
End If

ToDebug " Creating RA_FGM..."
Set mobjManager = New FilgraphManager

On Error Resume Next    'RenderAuto()

'samplegrabber с ним понадежнее захват captureBasic
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrGrabberName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    'Set mobjManager = Nothing
    'Exit Function
    'раньше работало и без
Else
    Call objRegFilterInfo.Filter(objFilterInfo)
    Set objRegFilterInfo = Nothing
    Set mobjSampleGrabber = objFilterInfo.Filter
    If (mobjSampleGrabber Is Nothing) Then
        Set objFilterInfo = Nothing
        'Set mobjManager = Nothing
        'Call err.Clear
        'Exit Function
    Else

        With udtMediaType
            .MajorType = UUIDFromString(amIDMediaTypeVideo)
            .SubType = UUIDFromString(amIDMediaTypeVideoRGB24)
            .FormatType = UUIDFromString(amIDFormatVideoInfo)
        End With

        With mobjSampleGrabber
            .MediaType = udtMediaType
            Call .SetBufferSamples(0&)
            Call .SetOneShot(0&)
        End With
    End If
End If
''''''-sg

Call mobjManager.RenderFile(mpgName)
If (0& = err.Number) Then
    MpegMediaOpen = True

    'подогнать и показать                           видео окно
    MpegSizeAdjust False

Else
    ToDebug "Err_RA_proc: " & mpgName
    RenderAuto = False
    MpegMediaClose  'Set mobjManager = Nothing
    Exit Function
End If

'Set objPosition = mobjManager
'Set objVideo = mobjManager

'If Not AutoAddingFlag Then 'это не убирает звук, а просто не трогает
Set objAudio = mobjManager
On Error Resume Next
err.Clear
objAudio.Volume = -10000
If err <> 0 Then ToDebug "RA_CanNotPlaySound"
err.Clear
On Error GoTo 0
'Set objAudio = Nothing
'End If

On Error Resume Next
ToDebug "--- Render Auto Filters Begin"
For Each GraphFilter In mobjManager.FilterCollection
    ToDebug " " & GraphFilter.name
    'Debug.Print GraphFilter.Name
Next GraphFilter
ToDebug "--- Render Auto Filters End"
On Error GoTo 0

'после старта
'If AutoShots Then
'    'проба захвата кадра
'    DoEvents
'    MPGCaptureBasicVideo FrmMain.PicTempHid(0)
'    If MPGCaptured = False Then
'        If Not AutoNoMessFlag Then
'            myMsgBox msgsvc(38) & vbCrLf & mpgName
'            ToDebug "Err RA: " & msgsvc(38) & vbCrLf & mpgName
'        End If
'    End If
'End If

RenderAuto = True
ToDebug " RA_Ok"
'End With
End Function

Public Function DShGetInfoForCapture() As Boolean

Dim i As Integer
Dim temp As Currency
Dim Handle As Long
Dim sAsp As String
DShGetInfoForCapture = True    'если что, false и выход

'инфа из файла
On Error Resume Next
Handle = MediaInfo_New()
If err Then
    'If Not AutoNoMessFlag Then
    myMsgBox msgsvc(51) & App.Path & "\MediaInfo.dll", vbCritical
    'End If
    'MsgBox "MediaInfo.dll not found. Reinstall SurVideoCatalog!", vbCritical
    ToDebug "Err_NoDLL: MediaInfo.dll"
    DShGetInfoForCapture = False
    Exit Function
End If

Call MediaInfo_Open(Handle, StrPtr(DShName))
'ToDebug "MediaInfo: " & bstr(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr("")))
err.Clear: On Error GoTo 0

'видео
If MediaInfo_Count_Get(Handle, MediaInfo_Stream_Video, -1) > 0 Then
    MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
 sAsp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))
 
    If Len(MMI_Format_str) = 0 Then
        MMI_Format_str = "4/3"
        ToDebug "MMInfo_DS no format. = " & MMI_Format_str
    Else
        ToDebug "MMInfo_DS Format = " & MMI_Format_str
    End If
    MMI_Format = CalcFormat(MMI_Format_str, sAsp)

Else    'нет видео
    MMI_Format_str = "4/3"
    MMI_Format = CalcFormat(MMI_Format_str, sAsp)
    ToDebug "MMInfo не опознал файл. Format=4/3"
End If

If Handle <> 0 Then MediaInfo_Close Handle

'****************************************************DIRECT X *********************

If Not RenderAuto Then
    DShGetInfoForCapture = False
    Exit Function
End If

'тестовый запуск
Set objPosition = mobjManager
objPosition.CurrentPosition = 5#
Set objVideo = mobjManager

mobjManager.Stop
mobjManager.Run
'Do While MediaState <> 2: Loop
'Sleep 300  ' а то не видно после первого скролла
'Debug.Print Time, MediaState
mobjManager.Pause
'Debug.Print Time, MediaState

If AutoShots Then
    '                                                           проба захвата кадра
    'DoEvents
    MPGCaptureBasicVideo FrmMain.PicTempHid(0)
    If MPGCaptured = False Then
        If Not AutoNoMessFlag Then
            myMsgBox msgsvc(38) & vbCrLf & mpgName
            ToDebug "Err RA: " & msgsvc(38) & vbCrLf & mpgName
        End If
    End If
End If

On Error Resume Next    'automation error на нулевых вобах и др

'Time
TimeL = objPosition.Duration

If err Then
    ToDebug err.Description
    DShGetInfoForCapture = False
    Exit Function
End If


If (TimeL = 0) Or (TimeL > 10000000) Then
    ToDebug "Error: неприемлемая длительность."
    'ClearVideo
    DShGetInfoForCapture = False
    Exit Function
End If

' aspect
'    If Format(MMI_Format, "0.000") = 1.333 Then 'тут может ошибиться в расчетах ()и затем
PixelRatio = objVideo.SourceHeight * MMI_Format / objVideo.SourceWidth
PixelRatioSS = ScrShotEd_W / MMI_Format

TimesX100 = TimeL * 100
temp = objVideo.AvgTimePerFrame * 100

With frmEditor

    .Position.min = 0
    .Position.Max = TimesX100    '- temp
    .Position.Value = 0
    .Position.TickFrequency = .Position.Max / 100
    .Position.SmallChange = temp * 100  '0.04*100 *1000
    .Position.LargeChange = temp * 1000
    PPMax = .Position.Value + cMPGRange    ' Const Range As Integer = 10000 в MpegPosScroll
    If PPMax > TimesX100 Then PPMax = .Position.Max    'TimesX100
    .PositionP.min = 0    'PPMin
    .PositionP.TickFrequency = temp
    .PositionP.Max = PPMax
    .PositionP.SmallChange = temp / 2
    .PositionP.LargeChange = temp * 10
    .PositionP.Value = 0
    .Position.Enabled = True
    .PositionP.Enabled = True

    .ComKeyAvi(0).Enabled = False: .ComKeyAvi(1).Enabled = False
    If MPGCaptured Then
        For i = 0 To 2
            ': ComCap(i).Enabled = True:
            .ComRND(i).Enabled = True
        Next i
        .ComAutoScrShots.Enabled = True
        DShGetInfoForCapture = True
        'Отразить аспекты на кнопках                                                    4:3 16:9
        Call EdAspect2Buttons

    Else
        For i = 0 To 2
            'ComCap(i).Enabled = False
            .ComRND(i).Enabled = False
        Next i
        .ComAutoScrShots.Enabled = False
        For i = 0 To 2: .optAspect(i).Enabled = False: Next i
        DShGetInfoForCapture = False
    End If

End With
End Function


Public Function MpgGetInfoForCapture() As Boolean
'это выполняется, когда mminfo опознал файл (нашел там видеопоток) и файл MPV

Dim i As Integer    ', j As Integer
Dim temp As Currency    'Long
Dim Handle As Long
'Dim tmps As String    ', tmp2s As String
Dim MMI_Height As Integer    'из MMInfo
Dim MMI_Width As Integer
Dim objv_Height As Integer    'из objVideo.Source
Dim objv_Width As Integer
Dim sAsp As String
'Dim ret As Long
'Dim WFD As WIN32_FIND_DATA
Dim ifo_handle As Long    'пусть
Dim rendMPV2 As Boolean    'срендерили как мпег2 с нашими фильтрами
Dim IsVob As Boolean

MpgGetInfoForCapture = True

Handle = MediaInfo_New()
Call MediaInfo_Open(Handle, StrPtr(mpgName))
'ToDebug "MediaInfo: " & bstr(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr("")))
'инфа из файла
On Error Resume Next

'видео mminfo
'                           аспект
If MediaInfo_Count_Get(Handle, MediaInfo_Stream_Video, -1) > 0 Then

    'строка 4/3  mminfo
    MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
    sAsp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))
    If Len(MMI_Format_str) = 0 Then
        MMI_Format_str = "4/3"
        ToDebug "MMInfo no Format. = " & MMI_Format_str
    Else
        ToDebug "MMInfo Format = " & MMI_Format_str
    End If
    MMI_Format = CalcFormat(MMI_Format_str, sAsp)

    MMI_Height = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Height"), 1, 0))
    MMI_Width = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Width"), 1, 0))
    'AspectRatio = GetAspectRatio(AspectRatioS, HeightS)
    MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))

End If

If Handle <> 0 Then MediaInfo_Close Handle

ToDebug "MMI_Format=" & MMI_Format

''оставить текущий!
'MMI_Format = GetAspectFromTextVideo
ToDebug "Ifo Format=" & MMI_Format


'****************************************************DIRECT X *********************


SVCDflag = False
If Not IsVob Then
    If MMI_Width = 480 And MMI_Height = 576 Then SVCDflag = True
    If MMI_Width = 480 And MMI_Height = 480 Then SVCDflag = True
End If
ToDebug "SVCD = " & SVCDflag
ToDebug "MMI_Codec: " & MPGCodec

Select Case strMPEG
Case "MPV2"
    If Opt_UseOurMpegFilters And Not SVCDflag Then        'нашим декодером
        ToDebug "Try RenderMPV2..."
        If RenderMPV2 Then    '              -----    RenderMPV2
            rendMPV2 = True
        Else
            ToDebug "Try RenderAuto..."
            If Not RenderAuto Then    '     ----- RenderAuto
                'все плохо
                If Not AutoNoMessFlag Then
                    myMsgBox msgsvc(10) & mpgName        'Ошибка работы с файлом
                End If
                MpgGetInfoForCapture = False
                Exit Function
            End If
        End If
    Else        'сразу авто
        ToDebug "RenderAuto..."
        If Not RenderAuto Then   '  ----- RenderAuto
            'все плохо
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(10) & mpgName
            End If
            MpgGetInfoForCapture = False
            Exit Function
        End If
    End If

Case "MPV1"
    If RenderMPV1 Then  '  ----- RenderMPV1
        'rendMPV1 = True
    Else
        If Not RenderAuto Then  '  ----- RenderAuto
            'все плохо
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(10) & mpgName
            End If
            MpgGetInfoForCapture = False
            Exit Function
        End If
    End If

End Select


On Error Resume Next
Set objPosition = mobjManager
objPosition.CurrentPosition = 5#
Set objVideo = mobjManager

'тестовый запуск
frmEditor.FrAdEdPixHid.Visible = True

ToDebug "Run..."
mobjManager.Stop
mobjManager.Run
'Sleep 300
mobjManager.Pause
ToDebug "Pause. Err=" & err.Number

If AutoShots Then
    '                                           проба захвата кадра
    'DoEvents
    MPGCaptureBasicVideo FrmMain.PicTempHid(0)
    If MPGCaptured = False Then
        'ToDebug "Ошибка: файл загружен, но есть ошибка захвата кадра: " & mpgName
        If Not AutoNoMessFlag Then myMsgBox msgsvc(38) & vbCrLf & mpgName
    End If
End If


err.Clear: On Error GoTo 0


'------------------------------------
ToDebug "Editor/FileName=" & mpgName
TimeL = objPosition.Duration

'Debug.Print "TimeL = " & TimeL
If (TimeL = 0) Or (TimeL > 10000000) Then
    'ClearVideo
    ToDebug msgsvc(40) & " = " & TimeL
    If Not AutoNoMessFlag Then myMsgBox msgsvc(40) & " = " & TimeL, vbCritical, , FrmMain.hwnd
    MpgGetInfoForCapture = False
    Exit Function
End If

'                                                                 Frame Size
objv_Width = objVideo.SourceWidth
objv_Height = objVideo.SourceHeight

'                                                                   aspect
PixelRatio = 1.333: PixelRatioSS = ScrShotEd_W / (4 / 3)        'дефолт 4:3
'MMI_Format не менять без MMI_Format_str

'то же MpgGetInfo
PixelRatio = (objv_Height * MMI_Format) / objv_Width
PixelRatioSS = ScrShotEd_W / MMI_Format

TimesX100 = TimeL * 100
temp = objVideo.AvgTimePerFrame * 100

With frmEditor
    .Position.min = 0
    .Position.Max = TimesX100        '- temp
    .Position.Value = 0
    .Position.TickFrequency = .Position.Max / 100
    .Position.SmallChange = temp * 100        '0.04*100 *1000
    .Position.LargeChange = temp * 1000
    PPMax = .Position.Value + cMPGRange        ' Const Range As Integer = 10000 в MpegPosScroll
    If PPMax > TimesX100 Then PPMax = .Position.Max        ' TimesX100
    .PositionP.min = 0        'PPMin
    .PositionP.TickFrequency = temp
    .PositionP.Max = PPMax
    .PositionP.SmallChange = temp
    .PositionP.LargeChange = temp * 10
    .PositionP.Value = 0
    .Position.Enabled = True: .PositionP.Enabled = True

    .ComKeyAvi(0).Enabled = False: .ComKeyAvi(1).Enabled = False

    If MPGCaptured Then
        For i = 0 To 2: .ComCap(i).Enabled = True: .ComRND(i).Enabled = True: Next i
        .ComAutoScrShots.Enabled = True
        MpgGetInfoForCapture = True

        'Отразить аспекты на кнопках                                                    4:3 16:9
        Call EdAspect2Buttons

    Else
        For i = 0 To 2
            ': ComCap(i).Enabled = False:
            .ComRND(i).Enabled = False
        Next i
        .ComAutoScrShots.Enabled = False
        For i = 0 To 2: .optAspect(i).Enabled = False: Next i
        MpgGetInfoForCapture = False
    End If

End With
End Function

Public Function DShGetInfo() As Boolean
'еще DShGetInfoForCapture
Dim i As Integer
Dim temp As Currency    'Long
Dim tmps As String, tmp2s As String
Dim Handle As Long
Dim Pointer As Long, lpFSHigh As Currency
Dim tmp As String
Dim tmpL As Long
Dim sAsp As String

DShGetInfo = True        'если что, false и выход

'инфа из файла
On Error Resume Next
Handle = MediaInfo_New()
If err Then
    If Not AutoNoMessFlag Then
        myMsgBox msgsvc(51) & App.Path & "\MediaInfo.dll", vbCritical
    End If
    'MsgBox "MediaInfo.dll not found. Reinstall SurVideoCatalog!", vbCritical
    ToDebug "Err_NoDLL: MediaInfo.dll"
    DShGetInfo = False
    Exit Function
End If

With frmEditor

    Call MediaInfo_Open(Handle, StrPtr(DShName))
    ToDebug "MediaInfo: " & bstr(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr("")))

    'звук                                   звук
    .TextAudioHid = vbNullString
    tmpL = MediaInfo_Count_Get(Handle, MediaInfo_Stream_Audio, -1)
    If tmpL > 0 Then
        For i = 0 To tmpL - 1
            tmps = vbNullString
            tmps = bstr(MediaInfo_Get(Handle, 2, i, StrPtr("SamplingRate"), 1, 0)) & " "
            tmps = tmps & Chnnls(bstr(MediaInfo_Get(Handle, 2, i, StrPtr("Channels"), 1, 0))) & " "
            tmps = tmps & bstr(MediaInfo_Get(Handle, 2, i, StrPtr("Codec"), 1, 0))

            tmp = bstr(MediaInfo_Get(Handle, 2, i, StrPtr("BitRate"), 1, 0))
            tmp = Replace(tmp, ".", ",")
            If IsNumeric(tmp) Then
                If Val(tmp) > 0 Then
                    tmp = tmp / 1000 & "kbps" 'звук
                    tmps = tmps & " (" & tmp & ")"
                End If
            End If

            .TextAudioHid = tmps & ", " & .TextAudioHid
        Next i
        .TextAudioHid = Trim$(left$(.TextAudioHid, Len(.TextAudioHid) - 2))
        ToDebug "DS_Audio>" & .TextAudioHid

    End If
    err.Clear: On Error GoTo 0

    'видео
    If MediaInfo_Count_Get(Handle, MediaInfo_Stream_Video, -1) > 0 Then
        '?last14.12.2006 MMI_Ratio = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("PixelRatio"), 1, 0))
        '    MMI_Ratio = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))
        '    If Len(MMI_Ratio) = 0 Then
        '        '1выбор юзером аспект
        '        MMI_Ratio = "1.333"
        '        'NoAspectFlag = True
        '        ToDebug "MMI_Ratio_DS не найден. = " & MMI_Ratio
        '    Else
        '        ToDebug "MMI_Ratio_DS=" & MMI_Ratio
        '    End If

        'Debug.Print "MMI_Ratio_DS=" & MMI_Ratio
        'строка 4/3
        MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
        sAsp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))

        If Len(MMI_Format_str) = 0 Then
            MMI_Format_str = "4/3"
            ToDebug "MMInfo_DS не нашел Format. = " & MMI_Format_str
        Else
            ToDebug "MMInfo_DS Format = " & MMI_Format_str

        End If
        MMI_Format = CalcFormat(MMI_Format_str, sAsp)

        tmps = vbNullString
        tmps = MyCodec(bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))) & " "
        MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))
        ToDebug "DS_Video>" & MPGCodec

        tmp2s = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("FrameRate"), 1, 0))
        tmp2s = Replace(tmp2s, ".", ",")
        If IsNumeric(tmp2s) Then
            .TextFPSHid = Val(tmp2s)
            Select Case Int(Val(tmp2s))
            Case "25"
                tmps = tmps & "PAL" & " "
            Case "29"
                tmps = tmps & "NTSC" & " "
            Case "23"    '23.976
                tmps = tmps & "FILM" & " "
            Case Else
                'tmps = tmps & TextFPSHid & " "
            End Select
        Else
            'mmi не нашли видео
            .TextFPSHid = vbNullString
        End If

        'добавить аспект видео
        tmps = tmps & MMI_Format_str & " "
        tmps = Replace(tmps, "2.35", "2.35:1")

        'и битрейт видео
        tmp = vbNullString
        tmp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("BitRate"), 1, 0))
        tmp = Replace(tmp, ".", ",")
        If IsNumeric(tmp) Then
            If Val(tmp) > 0 Then
                'tmp = tmp / 1000 & "kbps"
                tmp = Format$(tmp / 1000, "0") & "kbps"
                tmps = tmps & "(" & tmp & ")"
            End If
        End If
        tmps = Replace(tmps, "16/9", "16:9")
        tmps = Replace(tmps, "4/3", "4:3")

        .TextVideoHid = RTrim$(tmps)

    Else
        '    MMI_Ratio = "1.333"
        MMI_Format_str = "4/3"
        MMI_Format = CalcFormat(MMI_Format_str, sAsp)
        ToDebug "MMInfo не опознал файл. Format=4/3"
    End If

    If Handle <> 0 Then MediaInfo_Close Handle
End With


    '****************************************************DIRECT X **********RenderAuto***********


If Not RenderAuto Then
    DShGetInfo = False
    Exit Function
End If

'тестовый запуск
On Error Resume Next

Set objPosition = mobjManager
objPosition.CurrentPosition = 5#
Set objVideo = mobjManager

ToDebug "dsh_Run..."
mobjManager.Stop
'If AutoShots Then
mobjManager.Run
'End If
'Do While MediaState <> 2: Loop
'Sleep 300  ' а то не видно после первого скролла
'Debug.Print Time, MediaState
mobjManager.Pause
'Debug.Print Time, MediaState
ToDebug "Pause. Err=" & err.Number

If AutoShots Then
    '                                                           проба захвата кадра
    'DoEvents
    MPGCaptureBasicVideo FrmMain.PicTempHid(0)
    If MPGCaptured = False Then
        If Not AutoNoMessFlag Then
            myMsgBox msgsvc(38) & vbCrLf & mpgName
            ToDebug "Err RA: " & msgsvc(38) & vbCrLf & mpgName
        End If
    End If
End If

With frmEditor
    If Not AppendMovieFlag Then
        .TextFileName.Text = GetEdTxtFileName(DShName)
    Else
        .TextFileName.Text = .TextFileName.Text & " | " & GetEdTxtFileName(DShName)
    End If

    On Error Resume Next    'automation error на нулевых вобах и др

    'Time
    If Not AppendMovieFlag Then
        TextTimeMSHid = objPosition.Duration
    Else
        TextTimeMSHid = TextTimeMSHid + objPosition.Duration
    End If
    .TextTimeHid = FormatTime(TextTimeMSHid)
    TimeL = objPosition.Duration
End With

If err Then
    ToDebug err.Description
    DShGetInfo = False
    Exit Function
End If


If (TimeL = 0) Or (TimeL > 10000000) Then
    ToDebug "Err: неприемлемый размер видео"
    DShGetInfo = False
    Exit Function
End If

With frmEditor
    'fps
    If objVideo.AvgTimePerFrame > 0 Then
        .TextFPSHid = Round(1 / objVideo.AvgTimePerFrame, 3)
    End If

    'file size

    If isWindowsNt Then
        Pointer = lopen(DShName, OF_READ)
        GetFileSizeEx Pointer, lpFSHigh
        If Not AppendMovieFlag Then
            .TextFilelenHid = Int(lpFSHigh * 10000 / 1024)
        Else
            .TextFilelenHid = Val(.TextFilelenHid) + Int(lpFSHigh * 10000 / 1024)
        End If
        lclose Pointer
    Else
        If Not AppendMovieFlag Then
            .TextFilelenHid = Int(FileLen(DShName) / 1024)
        Else
            .TextFilelenHid = Val(.TextFilelenHid) + Int(FileLen(DShName) / 1024)
        End If
    End If

    'FrSize
    .TextResolHid = Trim$(str$(objVideo.SourceWidth)) & " x " & Trim$(str$(objVideo.SourceHeight))

    ' aspect
    '    If Format(MMI_Format, "0.000") = 1.333 Then
    'тут может ошибиться в расчетах ()и затем
    PixelRatio = objVideo.SourceHeight * MMI_Format / objVideo.SourceWidth
    PixelRatioSS = ScrShotEd_W / MMI_Format


    'cds
    If Not AppendMovieFlag Then
        .TextCDN.Text = 1
    Else
        If NewDiskAddFlag Then
            .TextCDN.Text = Replace(.TextCDN, Val(.TextCDN), Val(.TextCDN) + 1)
        End If
    End If

    TimesX100 = TimeL * 100

    temp = objVideo.AvgTimePerFrame * 100

    .Position.min = 0
    .Position.Max = TimesX100    '- temp
    .Position.Value = 0
    .Position.TickFrequency = .Position.Max / 100
    .Position.SmallChange = temp * 100  '0.04*100 *1000
    .Position.LargeChange = temp * 1000
    PPMax = .Position.Value + cMPGRange    ' Const Range As Integer = 10000 в MpegPosScroll
    If PPMax > TimesX100 Then PPMax = .Position.Max    'TimesX100
    .PositionP.min = 0    'PPMin
    .PositionP.TickFrequency = temp
    .PositionP.Max = PPMax
    .PositionP.SmallChange = temp / 2
    .PositionP.LargeChange = temp * 10
    .PositionP.Value = 0
    .Position.Enabled = True
    .PositionP.Enabled = True

    .ComKeyAvi(0).Enabled = False: .ComKeyAvi(1).Enabled = False
    If MPGCaptured Then
        For i = 0 To 2: .ComCap(i).Enabled = True: .ComRND(i).Enabled = True: Next i
        .ComAutoScrShots.Enabled = True

        'Отразить аспекты на кнопках                                                    4:3 16:9
        Call EdAspect2Buttons

    Else
        For i = 0 To 2: .ComCap(i).Enabled = False: .ComRND(i).Enabled = False: Next i
        .ComAutoScrShots.Enabled = False
        For i = 0 To 2: .optAspect(i).Enabled = False: Next i
    End If

    .ComAdd.Enabled = True

End With
End Function

Public Function RenderMPV1() As Boolean
'вставка в граф
'"MPEG-I Stream Splitter" 'quartz.dll
'"MPEG Video Decoder" 'quartz.dll

Dim objRegFilterInfo As IRegFilterInfo
Dim objFilterInfo As IFilterInfo
'Dim udtMediaType As TAMMediaType
Dim GraphFilter As IFilterInfo
Dim udtMediaType As TAMMediaType
'Dim objPinInfo As IPinInfo

If MpegMediaOpen Then
    MpegMediaClose
    Clear_mobjManager
End If

ToDebug " Creating RMPV1_FGM..."
Set mobjManager = New FilgraphManager

'samplegrabber с ним понадежнее захват captureBasic
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrGrabberName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    'Set mobjManager = Nothing
    'Exit Function
    'раньше работало и без
Else
    Call objRegFilterInfo.Filter(objFilterInfo)
    Set objRegFilterInfo = Nothing
    Set mobjSampleGrabber = objFilterInfo.Filter
    If (mobjSampleGrabber Is Nothing) Then
        Set objFilterInfo = Nothing
        'Set mobjManager = Nothing
        'Call err.Clear
        'Exit Function
    Else

        With udtMediaType
            .MajorType = UUIDFromString(amIDMediaTypeVideo)
            .SubType = UUIDFromString(amIDMediaTypeVideoRGB24)
            .FormatType = UUIDFromString(amIDFormatVideoInfo)
        End With

        With mobjSampleGrabber
            .MediaType = udtMediaType
            Call .SetBufferSamples(0&)
            Call .SetOneShot(0&)
        End With
    End If
End If
''''''-sg


'впихнуть в граф наш декодер cstrFVDName
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrM1SSName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    ToDebug "Error: Необходимый фильтр (" & cstrM1SSName & ") не зарегистрирован"
    MpegMediaClose     'Set mobjManager = Nothing
    RenderMPV1 = False
    Exit Function
End If
'    On Error Resume Next
Call objRegFilterInfo.Filter(objFilterInfo)
Set objRegFilterInfo = Nothing

'впихнуть в граф наш декодер cstrM1SSName
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrM1VDName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    ToDebug "Error: Необходимый фильтр (" & cstrM1VDName & ") не зарегистрирован"
    MpegMediaClose     'Set mobjManager = Nothing
    RenderMPV1 = False
    Exit Function
End If
'    On Error Resume Next
Call objRegFilterInfo.Filter(objFilterInfo)
Set objRegFilterInfo = Nothing

err.Clear
mobjManager.RenderFile mpgName
If (0& = err.Number) Then

    MpegMediaOpen = True
    
    Call MpegSizeAdjust(False) '                                 настройка окна

    '    With mudtBitmapInfo
    '        .Size = Len(mudtBitmapInfo)
    '        .Planes = 1
    '        .BitCount = 24
    '    End With

    ' что-то не соединилось
    '   For Each objPinInfo In objFilterInfo.Pins
    '       If (objPinInfo.ConnectedTo Is Nothing) Then
    '           Set mobjSampleGrabber = Nothing
    '           Set objPinInfo = Nothing
    '           Exit For
    '       End If
    '       Set objPinInfo = Nothing
    '   Next

    err.Clear

Else
    ToDebug "Ошибка: не могу обработать файл: " & mpgName
    RenderMPV1 = False
    MpegMediaClose     'Set mobjManager = Nothing
    Exit Function


End If

'Set objPosition = mobjManager
'Set objVideo = mobjManager

'If Not AutoAddingFlag Then 'это не убирает звук, а просто не трогает
Set objAudio = mobjManager
On Error Resume Next: objAudio.Volume = -10000: On Error GoTo 0
'Set objAudio = Nothing
'End If

'список фильтров
'Debug.Print "MPV1"
ToDebug "Render MPV1 Filters"
For Each GraphFilter In mobjManager.FilterCollection
    'Debug.Print GraphFilter.Name
    ToDebug " " & GraphFilter.name
Next GraphFilter
ToDebug "Render MPV1 Filters End"

'If AutoShots Then
''проба захвата кадра
'DoEvents
'MPGCaptureBasicVideo FrmMain.PicTempHid(0)
'If MPGCaptured = False Then
'    'ToDebug "Ошибка: файл загружен, но есть ошибка захвата кадра: " & mpgName
'    If Not AutoNoMessFlag Then myMsgBox msgsvc(38) & vbCrLf & mpgName
'End If
'End If

RenderMPV1 = True
ToDebug " RMPV1_Ok"
End Function

Public Function RenderMPV2() As Boolean
'вставка в граф
'"Fraunhofer Video Decoder" (dvdvideo.ax из фри кодека) + audio
'"MPEG-2 Splitter" (mpg2splt.ax)

Dim objRegFilterInfo As IRegFilterInfo
Dim objFilterInfo As IFilterInfo
Dim GraphFilter As IFilterInfo
Dim objPin As IPinInfo
Dim SourceFilter As Boolean    'удачно ли прошел первый фильтр Universal Open Source MPEG Source"
Dim udtMediaType As TAMMediaType
'Dim objPinInfo As IPinInfo

On Error Resume Next    'надо

'If False Then

If MpegMediaOpen Then
    MpegMediaClose
    Clear_mobjManager
End If

ToDebug " Creating RMPV2_FGM..."
Set mobjManager = New FilgraphManager

'samplegrabber с ним понадежнее захват captureBasic
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrGrabberName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    'Set mobjManager = Nothing
    'Exit Function
    'раньше работало и без
Else
    Call objRegFilterInfo.Filter(objFilterInfo)
    Set objRegFilterInfo = Nothing
    Set mobjSampleGrabber = objFilterInfo.Filter
    If (mobjSampleGrabber Is Nothing) Then
        Set objFilterInfo = Nothing
        'Set mobjManager = Nothing
        'Call err.Clear
        'Exit Function
    Else

        With udtMediaType
            .MajorType = UUIDFromString(amIDMediaTypeVideo)
            .SubType = UUIDFromString(amIDMediaTypeVideoRGB24)
            .FormatType = UUIDFromString(amIDFormatVideoInfo)
        End With

        With mobjSampleGrabber
            .MediaType = udtMediaType
            Call .SetBufferSamples(0&)
            Call .SetOneShot(0&)
        End With
    End If
End If
''''''-sg

'впихнуть в граф наш декодер
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrVDName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    ToDebug "Error: Необходимый фильтр (" & cstrVDName & ") не зарегистрирован"
Else
    objRegFilterInfo.Filter objFilterInfo
End If

'ТУТ не надо - ошибка при проигрывании cstrMPG2DemName демультиплексер mpg2splt.ax
''' ^ впихнули

''впихнуть в граф наш декодер cstrSVDName декодер
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrSVDName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    ToDebug "Error: Необходимый фильтр (" & cstrSVDName & ") не зарегистрирован"
    SourceFilter = False
Else
    SourceFilter = True
End If

If SourceFilter Then
    objRegFilterInfo.Filter objFilterInfo
    If err.Number <> 0 Then
        ToDebug "Err: переустановите Universal Open Source MPEG Source"
        SourceFilter = False
    End If

    If SourceFilter Then
        'для source фильтра "Universal Open Source MPEG Source"
        'Рендерим
        If Not objRegFilterInfo Is Nothing Then
            objFilterInfo.filename = mpgName
            For Each objPin In objFilterInfo.Pins
                objPin.Render 'не все, некоторые отсутствуют (on error)
            Next
        End If
        err.Clear     'если нет субтитров у "Universal Open Source MPEG Source"
        If objFilterInfo.filename = mpgName Then
            'вышло с соурсом
            SourceFilter = True
        Else
            SourceFilter = False
        End If
    End If
End If

'End If 'надо sourse

If Not SourceFilter Then
    'проба с другим нашим фильтром (не соурс)
    '                                                       все с начала
If MpegMediaOpen Then
    MpegMediaClose
    Clear_mobjManager
End If

    Set mobjManager = New FilgraphManager

'samplegrabber с ним понадежнее захват captureBasic
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrGrabberName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    'Set mobjManager = Nothing
    'Exit Function
    'раньше работало и без
Else
    Call objRegFilterInfo.Filter(objFilterInfo)
    Set objRegFilterInfo = Nothing
    Set mobjSampleGrabber = objFilterInfo.Filter
    If (mobjSampleGrabber Is Nothing) Then
        Set objFilterInfo = Nothing
        'Set mobjManager = Nothing
        'Call err.Clear
        'Exit Function
    Else

        With udtMediaType
            .MajorType = UUIDFromString(amIDMediaTypeVideo)
            .SubType = UUIDFromString(amIDMediaTypeVideoRGB24)
            .FormatType = UUIDFromString(amIDFormatVideoInfo)
        End With

        With mobjSampleGrabber
            .MediaType = udtMediaType
            Call .SetBufferSamples(0&)
            Call .SetOneShot(0&)
        End With
    End If
End If
''''''-sg

    'впихнуть в граф наш декодер
    For Each objRegFilterInfo In mobjManager.RegFilterCollection
        If (cstrVDName = objRegFilterInfo.name) Then
            Exit For
        End If
        Set objRegFilterInfo = Nothing
    Next
    If (objRegFilterInfo Is Nothing) Then
        ToDebug "Error: Необходимый фильтр (" & cstrVDName & ") не зарегистрирован"
        MpegMediaClose    'Set mobjManager = Nothing
        RenderMPV2 = False
        Exit Function
    End If

    err.Clear
    objRegFilterInfo.Filter objFilterInfo

    If err.Number <> 0 Then
        ToDebug "Error: наш фильтр не смог"
        MpegMediaClose    'Set mobjManager = Nothing
        RenderMPV2 = False
        Exit Function
    End If

    ''впихнуть в граф наш декодер cstrMPG2DemName демультиплексер mpg2splt.ax
    For Each objRegFilterInfo In mobjManager.RegFilterCollection
        If (cstrMPG2DemName = objRegFilterInfo.name) Then
            Exit For
        End If
        Set objRegFilterInfo = Nothing
    Next
    If (objRegFilterInfo Is Nothing) Then
        ToDebug "Error: Необходимый фильтр (" & cstrMPG2DemName & ") не зарегистрирован"
    Else
        objRegFilterInfo.Filter objFilterInfo
    End If



    'срендерить
    Call mobjManager.RenderFile(mpgName)

End If    'Not SourceFilter

Set objRegFilterInfo = Nothing

If (0& = err.Number) Then
    MpegMediaOpen = True

    Call MpegSizeAdjust(False)

Else
    ToDebug "Ошибка: Не могу обработать файл: " & mpgName
    err.Clear
    '        Set mobjSampleGrabber = Nothing
    MpegMediaClose     'Set mobjManager = Nothing
    RenderMPV2 = False
    Exit Function
End If

'Set objPosition = mobjManager
'Set objVideo = mobjManager

'If Not AutoAddingFlag Then 'это не убирает звук, а просто не трогает
Set objAudio = mobjManager
On Error Resume Next: objAudio.Volume = -10000: On Error GoTo 0
'Set objAudio = Nothing
'End If

'список фильтров

'Debug.Print "MPV2"
ToDebug "Render MPV2 Filters"
For Each GraphFilter In mobjManager.FilterCollection
    'Debug.Print GraphFilter.Name
    ToDebug vbTab & GraphFilter.name
Next GraphFilter
ToDebug "Render MPV2 Filters End"

'If AutoShots Then
'    'проба захвата кадра
'    DoEvents
'    MPGCaptureBasicVideo FrmMain.PicTempHid(0)
'    If MPGCaptured = False Then
'        'ToDebug "Ошибка: файл загружен, но есть ошибка захвата кадра: " & mpgName
'        If Not AutoNoMessFlag Then myMsgBox msgsvc(38) & vbCrLf & mpgName
'    End If
'End If

RenderMPV2 = True
ToDebug " RMPV2_Ok"
End Function

Public Sub GetAviInfo()
'смотри PrepareAviForCupture

Dim AVIInf As New clsAVIInfo
Dim TimeS As String
Dim i As Integer
Dim temp As String
Dim Pointer As Long, lpFSHigh As Currency

Dim mtag() As Long
Dim InfoChar As String
Dim InfoText As String

'для суммирующихся
Dim tmpLang As String
Dim tmpCountry As String
Dim tmpGenre As String

AVIInf.ReadFile (aviName)

If AVIInf.NumStreams <> 0 Then        'avi?
    isAVIflag = True
Else
    isAVIflag = False
    Set AVIInf = Nothing
    Exit Sub
End If

If Not AppendMovieFlag Then        'Новый файл

    With frmEditor
        Frames = AVIInf.numFrames
        If Frames < 1 Then m_cAVI.filename = vbNullString: Exit Sub    'no frames
        ToDebug "AVI_Frames=" & Frames

        'имя файла
        .TextFileName.Text = GetEdTxtFileName(aviName)

        'Time
        TimeS = FormatTime(AVIInf.PlayLength)
        TextTimeMSHid = AVIInf.PlayLength
        .TextTimeHid = TimeS

        'fps
        .TextFPSHid = Round(AVIInf.FrameRate, 3)
        ToDebug "AVI_FPS=" & .TextFPSHid

        'file size
        If isWindowsNt Then
            Pointer = lopen(aviName, OF_READ)
            'size of the file
            GetFileSizeEx Pointer, lpFSHigh
            .TextFilelenHid = Int(lpFSHigh * 10000 / 1024)
            lclose Pointer
        Else
            .TextFilelenHid = Int(FileLen(aviName) / 1024)
        End If
        ToDebug "AVI_fSize=" & .TextFilelenHid

        'аудио
        'TextAudioHid = AVIInf.SamplesPerSec & " " & AVIInf.Channels & " " & AVIInf.AudioFormat
        .TextAudioHid = AVIInf.AllAudio
        If Trim$(.TextAudioHid) = "0" Then .TextAudioHid = "-"

        'видео
        temp = AVIInf.VideoCodec & " - " & AVIInf.CodecToName(AVIInf.VideoCodec) & AVIInf.VideoCodec2 & AVIInf.VideoBitrate
        '    tmp = AVIInf.VideoCodec
        '    If Len(tmp) <> 0 Then
        '        temp = tmp & " - " & AVIInf.CodecToName(tmp) & AVIInf.VideoCodec2 & AVIInf.VideoBitrate
        '    Else
        '        temp = "- " & AVIInf.CodecToName(AVIInf.VideoCodec22) & AVIInf.VideoCodec2 & AVIInf.VideoBitrate
        '    End If

        Do While InStr(1, temp, "  ")
            temp = Replace(temp, "  ", " ")
        Loop
        .TextVideoHid = temp

        'resolution
        AviWidth = AVIInf.Width
        AviHeight = AVIInf.Height
        Ratio = AviWidth / AviHeight

        'FrSize
        .TextResolHid = Trim$(str$(AviWidth)) & " x " & Trim$(str$(AviHeight))
        '
        .TextCDN.Text = 1

        'Мета данные (Info) meta
        '
        If AVIInf.GetInfoList(mtag) > 0 Then

            For i = 0 To UBound(mtag)
                If Len(mtag(i)) = 4 Then
                    InfoChar = AVIInf.LongToFourCC(mtag(i))
                    InfoText = AVIInf.QueryInfo(mtag(i))
                    If Len(InfoText) <> 0 Then

                        Select Case InfoChar
                        Case "INAM"    'Name/Title
                            If filltrue(.TextMName) Then
                                .TextMName = InfoText
                            ElseIf filltrueAdd Then
                                .TextMName = .TextMName & ", " & InfoText
                            End If
                        Case "IART"    'Artist , Director
                            If filltrue(.TextAuthor) Then
                                .TextAuthor = InfoText
                            ElseIf filltrueAdd Then
                                .TextAuthor = .TextAuthor & ", " & InfoText
                            End If
                        Case "ISTR"    'Starring
                            If filltrue(.TextRole) Then .TextRole = InfoText
                            If filltrue(.TextRole) Then
                                .TextRole = InfoText
                            ElseIf filltrueAdd Then
                                .TextRole = .TextRole & ", " & InfoText
                            End If
                        Case "IAS1", "IAS2", "IAS3", "IAS4", "IAS5", "IAS6", "IAS7", "IAS8", "IAS9"    'First -9 Language
                        If LCase$(LastLanguage) = "Русский" Then InfoText = Replace(InfoText, "Russian", "русский", Compare:=vbTextCompare)
                            tmpLang = InfoText & ", " & tmpLang
                            'Case "ILNG" 'Language ?
                            '    If filltrue(TextLang) Then TextLang = TextLang & ", " & InfoText
                            
                        Case "ICNT", "ISTD"    'Страна, 'Production studio
                            tmpCountry = InfoText & ", " & tmpCountry
                        Case "IGNR", "ISGN"    'Genre, 'Secondary genre
                            tmpGenre = InfoText & ", " & tmpGenre
                        Case "IWEB"    'Internet address
                            If filltrue(.TextMovURL) Then
                                .TextMovURL = InfoText
                            ElseIf filltrueAdd Then
                                .TextMovURL = .TextMovURL & ", " & InfoText
                            End If
                        Case "ICRD"    '  Creation Date (YYYYMMDD) наверно это год фильма
                            If filltrue(.TextYear) Then
                                .TextYear = left$(InfoText, 4)
                            ElseIf filltrueAdd Then
                                .TextYear = .TextYear & ", " & left$(InfoText, 4)
                            End If
                        Case "ISFT"  'Software (vdub)
                            'ToDebug AVIInf.QueryInfo(mtag(i))
                        Case "ICMT"    'comments
                            If filltrue(.TextOther) Then
                                .TextOther = InfoText
                            ElseIf filltrueAdd Then
                                .TextOther = .TextOther & ", " & InfoText
                            End If

                        End Select

                        'Debug.Print AVIInf.QueryInfo(mtag(i))
                    End If    '0 text
                End If    '4
            Next i

            'поместить суммы, кроме пустых
            If Len(tmpLang) <> 0 Then
                If filltrue(.TextLang) Then
                    .TextLang = tmpLang
                ElseIf filltrueAdd Then
                    .TextLang = .TextLang & ", " & tmpLang
                End If
            End If

            If Len(tmpCountry) <> 0 Then
                If filltrue(.TextCountry) Then
                    .TextCountry = tmpCountry
                ElseIf filltrueAdd Then
                    .TextCountry = .TextCountry & ", " & tmpCountry
                End If
            End If

            If Len(tmpGenre) <> 0 Then
                If filltrue(.TextGenre) Then
                    .TextGenre = tmpGenre
                ElseIf filltrueAdd Then
                    .TextGenre = .TextGenre & ", " & tmpGenre
                End If
            End If

            If right$(.TextLang, 2) = ", " Then .TextLang = left$(.TextLang, Len(.TextLang) - 2)
            If right$(.TextCountry, 2) = ", " Then .TextCountry = left$(.TextCountry, Len(.TextCountry) - 2)
            If right$(.TextGenre, 2) = ", " Then .TextGenre = left$(.TextGenre, Len(.TextGenre) - 2)
        End If    ' AVIInf.GetInfoList(mtag) > 0 есть метаданные
    End With
Else            '                                      Append
    'Debug.Print "Append AVI"

    Frames = AVIInf.numFrames
    If Frames < 1 Then Exit Sub            'no frames

    With frmEditor
        TextTimeMSHid = Int(TextTimeMSHid) + AVIInf.PlayLength
        .TextTimeHid = FormatTime(TextTimeMSHid)

        'size
        If isWindowsNt Then
            Pointer = lopen(aviName, OF_READ)
            GetFileSizeEx Pointer, lpFSHigh
            temp = Int(lpFSHigh * 10000 / 1024)
            lclose Pointer
            .TextFilelenHid = Int(Val(.TextFilelenHid)) + Int(Val(temp))
        Else
            .TextFilelenHid = Int(Val(.TextFilelenHid)) + Int(FileLen(aviName) / 1024)
        End If

        If NewDiskAddFlag Then .TextCDN.Text = Replace(.TextCDN, Val(.TextCDN), Val(.TextCDN) + 1)
        .TextFileName.Text = .TextFileName.Text & " | " & GetEdTxtFileName(aviName)
    End With
End If

With frmEditor
    ''Отразить аспекты на кнопках     ави - 1:1                                               4:3 16:9
    For i = 0 To 2
        .optAspect(i).Value = False
        .optAspect(i).Enabled = False
    Next i

    'Позиции
    .Position.min = 0
    If Frames > 1 Then .Position.Max = Frames - 1 Else .Position.Max = 1

    .Position.TickFrequency = .Position.Max / 100
    .Position.SmallChange = .Position.TickFrequency / 5
    .Position.LargeChange = .Position.Max / 100

    .PositionP.min = 0     'PPMin
    PPMax = .Position.Value + 1000
    If PPMax > Frames Then PPMax = Frames
    .PositionP.Max = PPMax
    .PositionP.SmallChange = 1     'PositionP.TickFrequency / 1
    .PositionP.LargeChange = 50     'PositionP.Max / 100

    'MovieWidth = ScaleX(movie.Width, vbTwips, vbPixels)
    MovieWidth = .movie.Width / Screen.TwipsPerPixelX

    If Ratio < 1 Then Ratio = 1.333
    MovieHeight = MovieWidth / Ratio
    'movie.Height = ScaleY(MovieHeight, vbPixels, vbTwips)
    .movie.Height = MovieHeight * Screen.TwipsPerPixelX

    Set AVIInf = Nothing

    If Not NoVideoProcess Then
        'видео окно
        Set m_cAVI = New cAVIFrameExtract
        m_cAVI.filename = aviName

        If Not aferror Then
        
        'If InStr(.TextVideoHid, "xvid") And (Frames > 5000) Then
        'xvid не любит скриншоты ключевых кадров, попытка пару раз напасть не на ключ
            If Frames > 5000 Then
                    'lastRendedAVI = m_cAVI.AVIStreamNearestPrevKeyFrame(5000)
                If Not m_cAVI.AVIStreamIsKeyFrame(5000) Then
                    lastRendedAVI = 5000
                    pRenderFrame lastRendedAVI
                    '.Position.Value = lastRendedAVI
                    '.PosScroll
                ElseIf Not m_cAVI.AVIStreamIsKeyFrame(4900) Then
                    lastRendedAVI = 4900
                    pRenderFrame lastRendedAVI
                    '.Position.Value = lastRendedAVI
                    '.PosScroll
                End If
            Else
                pRenderFrame 0
                lastRendedAVI = 0
'                .Position.Value = 0
'                .PositionP.Value = 0

            End If
            
           .Position.Value = 0 'пусть всегда в нулях остается для единообразия
            .PositionP.Value = 0

            .Position.Enabled = True
            .PositionP.Enabled = True
            .ComKeyAvi(0).Enabled = True: .ComKeyAvi(1).Enabled = True
            For i = 0 To 2: .ComCap(i).Enabled = True: .ComRND(i).Enabled = True: Next i
            .ComAdd.Enabled = True    ': ComAddHid.Enabled = True
            .ComAutoScrShots.Enabled = True
        Else
            m_cAVI.filename = vbNullString    'unload

        End If

    End If    'видео процесс

End With
End Sub

Public Sub PrepareAviForCupture(fname As String)
'смотри GetAviInfo
'берет инфо, готовит видео окно и слайдеры
'генерит флаг ошибки aferror, если была
Dim AVIInf As New clsAVIInfo
'Dim temp As String

AVIInf.ReadFile (fname)

If AVIInf.NumStreams <> 0 Then        'avi?
    isAVIflag = True
Else
    isAVIflag = False
    Set AVIInf = Nothing
    Exit Sub
End If

Frames = AVIInf.numFrames
If Frames < 1 Then
    m_cAVI.filename = vbNullString
    Set AVIInf = Nothing
    Exit Sub    'no frames
End If

'resolution
AviWidth = AVIInf.Width
AviHeight = AVIInf.Height
Ratio = AviWidth / AviHeight

With frmEditor
    'Позиции
    .Position.min = 0
    If Frames > 1 Then .Position.Max = Frames - 1 Else .Position.Max = 1

    .Position.TickFrequency = .Position.Max / 100
    .Position.SmallChange = .Position.TickFrequency / 5
    .Position.LargeChange = .Position.Max / 100

    .PositionP.min = 0     'PPMin
    PPMax = .Position.Value + 1000
    If PPMax > Frames Then PPMax = Frames
    .PositionP.Max = PPMax
    .PositionP.SmallChange = 1     'PositionP.TickFrequency / 1
    .PositionP.LargeChange = 50     'PositionP.Max / 100

    'MovieWidth = ScaleX(.movie.Width, vbTwips, vbPixels)
    MovieWidth = .movie.Width / Screen.TwipsPerPixelX

    If Ratio < 1 Then Ratio = 1.333
    MovieHeight = MovieWidth / Ratio
    '.movie.Height = ScaleY(MovieHeight, vbPixels, vbTwips)
    .movie.Height = MovieHeight * Screen.TwipsPerPixelX

    Set AVIInf = Nothing

    If Not NoVideoProcess Then
        'видео окно
        Set m_cAVI = New cAVIFrameExtract
        m_cAVI.filename = fname

        If Not aferror Then
        
        'If InStr(.TextVideoHid, "xvid") And (Frames > 5000) Then
        'xvid не любит скриншоты ключевых кадров, попытка пару раз напасть не на ключ
            If Frames > 5000 Then
                    'lastRendedAVI = m_cAVI.AVIStreamNearestPrevKeyFrame(5000)
                If Not m_cAVI.AVIStreamIsKeyFrame(5000) Then
                    lastRendedAVI = 5000
                    pRenderFrame lastRendedAVI
                    '.Position.Value = lastRendedAVI
                    '.PosScroll
                ElseIf Not m_cAVI.AVIStreamIsKeyFrame(4900) Then
                    lastRendedAVI = 4900
                    pRenderFrame lastRendedAVI
                    '.Position.Value = lastRendedAVI
                    '.PosScroll
                End If
            Else
                pRenderFrame 0
                lastRendedAVI = 0
'                .Position.Value = 0
'                .PositionP.Value = 0

            End If
            
            .Position.Value = 0 'пусть всегда в нулях остается для единообразия
            .PositionP.Value = 0
            
            .Position.Enabled = True
            .PositionP.Enabled = True
            .ComKeyAvi(0).Enabled = True: .ComKeyAvi(1).Enabled = True
            .ComRND(0).Enabled = True: .ComRND(1).Enabled = True: .ComRND(2).Enabled = True
            .ComAutoScrShots.Enabled = True
        Else
            m_cAVI.filename = vbNullString    'unload
        End If
    End If    'видео процесс

End With
End Sub

Public Function MpegSizeAdjust(cl As Boolean) As Boolean
'MMI_Format переменная тут заменена на локальную movieAspect
Dim movieAspect As Currency 'Single

'паблик Dim objVideoW As IVideoWindow
Dim lngWidth As Long
Dim lngHeight As Long

On Error Resume Next

MpegSizeAdjust = True
Set objVideoW = mobjManager

If cl Then 'очистка
 If (True <> (objVideoW Is Nothing)) Then objVideoW.Visible = False ': objVideoW.Owner = 0
 Set objVideoW = Nothing
 ToDebug "Clear video"
 Exit Function
End If
    
If (True <> (objVideoW Is Nothing)) Then
'Debug.Print "objVideoW.Width.Orig=" & objVideoW.Width & " objVideoW.Height.Orig=" & objVideoW.Height

ToDebug "Установка окна..."
VideoStand = "PAL"
If InStr(1, frmEditor.TextVideoHid, "PAL", vbTextCompare) > 0 Then
    VideoStand = "PAL"
ElseIf InStr(1, frmEditor.TextVideoHid, "NTSC", vbTextCompare) > 0 Then
    VideoStand = "NTSC"
ElseIf InStr(1, frmEditor.TextVideoHid, "FILM", vbTextCompare) > 0 Then
    VideoStand = "FILM"
End If

movieAspect = MMI_Format 'по умолчанию

Select Case VideoStand
Case "PAL"
    Select Case strMPEG
     Case "MPV2"
      If SVCDflag Then
       'svcd
       movieAspect = 1.333
       'movieAspect 'objVideo.SourceHeight * movieAspect / objVideo.SourceWidth
       'PixelRatioSS = ScrShotEd_W / movieAspect
      Else
       'dvd
      End If
     Case "MPV1"
     'vcd
      movieAspect = 1.333
    End Select
Case "NTSC"
    Select Case strMPEG
     Case "MPV2"
      If SVCDflag Then
       'svcd
       movieAspect = 1.333
      Else
       'dvd
      End If
     Case "MPV1"
     'vcd
     movieAspect = 1.333
    End Select
Case "FILM"
    movieAspect = MMI_Format
End Select

If movieAspect = 0 Then movieAspect = 1.333 'mp4


 On Error GoTo err
 
  objVideoW.Owner = frmEditor.movie.hwnd
'  .WindowStyle = enWindowStyles.WS_VISIBLE '&H80000000    'WS_CHILD
  objVideoW.WindowStyle = WS_CHILD 'так надо
  lngWidth = objVideoW.Width: lngHeight = objVideoW.Height
  '.AutoShow = True
  '.FullScreenMode = True
  'MovieWidth = ScaleX(movie.Width, vbTwips, vbPixels)
  MovieWidth = frmEditor.movie.Width / Screen.TwipsPerPixelX
  MovieHeight = MovieWidth / movieAspect
  'movie.Height = ScaleY(MovieHeight, vbPixels, vbTwips)
  frmEditor.movie.Height = MovieHeight * Screen.TwipsPerPixelY
  Call objVideoW.SetWindowPosition(0&, 0&, MovieWidth, MovieHeight)
  'Call .SetWindowPosition(0&, 0&, lngWidth, lngHeight)
  
  objVideoW.Visible = True
End If

'Debug.Print "objVideoW.Width.movie=" & objVideoW.Width & " objVideoW.Height.movie=" & objVideoW.Height

ToDebug "видео-окно: " & movieAspect

err:
'Set objVideoW = Nothing
If err.Number <> 0 Then
    ToDebug "Error in MSAcl: " & err.Description
    MpegSizeAdjust = False
End If
End Function

Public Sub AutoScrShots(F As Long)
'автоскриншоты для ави
'F - всего фреймов
Dim tmp As Long

With frmEditor
    Set .PicSS1 = Nothing
    Set .PicSS2 = Nothing
    Set .PicSS3 = Nothing


    .PicSS1.Height = ScrShotEd_W * .movie.Height / .movie.Width
    .PicSS1.Width = ScrShotEd_W
    .PicSS2.Height = .PicSS1.Height
    .PicSS2.Width = .PicSS1.Width
    .PicSS3.Height = .PicSS1.Height
    .PicSS3.Width = .PicSS1.Width

    tmp = (F - F * 0.05) / 3    '4
    pos1 = 1 + (Rnd() * tmp)    '1-4
    pos2 = tmp + (Rnd() * tmp) + 1    '4-8
    pos3 = tmp * 2 + (Rnd() * tmp) + 1    '8-12

    pos1 = m_cAVI.AVIStreamNearestNextKeyFrame(pos1)
    pos2 = m_cAVI.AVIStreamNearestNextKeyFrame(pos2)
    pos3 = m_cAVI.AVIStreamNearestNextKeyFrame(pos3)

    If Opt_PicRealRes Then    'большие
        '.PicSS1Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
        .PicSS1Big.Width = AviWidth * Screen.TwipsPerPixelX
        '.PicSS1Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
        .PicSS1Big.Height = AviHeight * Screen.TwipsPerPixelY

        .PicSS2Big.Width = .PicSS1Big.Width
        .PicSS2Big.Height = .PicSS1Big.Height
        .PicSS3Big.Width = .PicSS1Big.Width
        .PicSS3Big.Height = .PicSS1Big.Height

        'тестовый
        'tmp = m_cAVI.AVIStreamNearestNextKeyFrame(pos1)
        'm_cAVI.DrawFrame PicSS1Big.hdc, tmp, 0, 0, Transparent:=False

        m_cAVI.DrawFrame .PicSS1Big.hdc, pos1, 0, 0, Transparent:=False
        m_cAVI.DrawFrame .PicSS2Big.hdc, pos2, 0, 0, Transparent:=False
        m_cAVI.DrawFrame .PicSS3Big.hdc, pos3, 0, 0, Transparent:=False

        .PicSS2Big.Picture = .PicSS2Big.Image
        .PicSS1Big.Picture = .PicSS1Big.Image
        .PicSS3Big.Picture = .PicSS3Big.Image

        'small
        .PicSS1.PaintPicture .PicSS1Big, 0, 0, .PicSS1.ScaleWidth, .PicSS1.ScaleHeight
        .PicSS1.Picture = .PicSS1.Image
        .PicSS2.PaintPicture .PicSS2Big, 0, 0, .PicSS2.ScaleWidth, .PicSS2.ScaleHeight
        .PicSS2.Picture = .PicSS2.Image
        .PicSS3.PaintPicture .PicSS3Big, 0, 0, .PicSS3.ScaleWidth, .PicSS3.ScaleHeight
        .PicSS3.Picture = .PicSS3.Image

    Else
        'only small
        m_cAVI.DrawFrame .PicSS1.hdc, pos1, lWidth:=.PicSS1.ScaleWidth, lHeight:=.PicSS1.ScaleHeight, Transparent:=False
        .PicSS1.Picture = .PicSS1.Image
        m_cAVI.DrawFrame .PicSS2.hdc, pos2, lWidth:=.PicSS2.ScaleWidth, lHeight:=.PicSS2.ScaleHeight, Transparent:=False
        .PicSS2.Picture = .PicSS2.Image
        m_cAVI.DrawFrame .PicSS3.hdc, pos3, lWidth:=.PicSS3.ScaleWidth, lHeight:=.PicSS3.ScaleHeight, Transparent:=False
        .PicSS3.Picture = .PicSS3.Image

    End If

End With
'Position.Value = lastRendedAVI
ToDebug "AScrShotsPos Avi: " & pos1 & " " & pos2 & " " & pos3
End Sub

Public Sub AutoScrShotsN(F As Long, ss As Integer)
'по кнопке autoN

With frmEditor
    Select Case ss
    Case 1

        Set .PicSS1 = Nothing
        .PicSS1.Height = ScrShotEd_W * .movie.Height / .movie.Width
        .PicSS1.Width = ScrShotEd_W
        pos1 = 1 + (Rnd() * F)
        pos1 = m_cAVI.AVIStreamNearestNextKeyFrame(pos1)

        If Opt_PicRealRes Then    'большие
            '.PicSS1Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
            .PicSS1Big.Width = AviWidth * Screen.TwipsPerPixelX
            '.PicSS1Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
            .PicSS1Big.Height = AviHeight * Screen.TwipsPerPixelY
            m_cAVI.DrawFrame .PicSS1Big.hdc, pos1, 0, 0, Transparent:=False
            .PicSS1Big.Picture = .PicSS1Big.Image
        End If

        m_cAVI.DrawFrame .PicSS1.hdc, pos1, lWidth:=.PicSS1.ScaleWidth, lHeight:=.PicSS1.ScaleHeight, Transparent:=False
        .PicSS1.Picture = .PicSS1.Image

    Case 2

        Set .PicSS2 = Nothing
        .PicSS2.Height = ScrShotEd_W * .movie.Height / .movie.Width
        .PicSS2.Width = ScrShotEd_W

        pos2 = 1 + (Rnd() * F)
        pos2 = m_cAVI.AVIStreamNearestNextKeyFrame(pos2)

        If Opt_PicRealRes Then    'большие
            '.PicSS2Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
            .PicSS2Big.Width = AviWidth * Screen.TwipsPerPixelX
            '.PicSS2Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
            .PicSS2Big.Height = AviHeight * Screen.TwipsPerPixelY
            m_cAVI.DrawFrame .PicSS2Big.hdc, pos2, 0, 0, Transparent:=False
            .PicSS2Big.Picture = .PicSS2Big.Image
        End If

        m_cAVI.DrawFrame .PicSS2.hdc, pos2, lWidth:=.PicSS1.ScaleWidth, lHeight:=.PicSS1.ScaleHeight, Transparent:=False
        .PicSS2.Picture = .PicSS2.Image

    Case 3

        Set .PicSS3 = Nothing
        .PicSS3.Height = ScrShotEd_W * .movie.Height / .movie.Width
        .PicSS3.Width = ScrShotEd_W

        pos3 = 1 + (Rnd() * F)
        pos3 = m_cAVI.AVIStreamNearestNextKeyFrame(pos3)

        If Opt_PicRealRes Then    'большие
            'PicSS3Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
            .PicSS3Big.Width = AviWidth * Screen.TwipsPerPixelX
            'PicSS3Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
            .PicSS3Big.Height = AviHeight * Screen.TwipsPerPixelY
            m_cAVI.DrawFrame .PicSS3Big.hdc, pos3, 0, 0, Transparent:=False
            .PicSS3Big.Picture = .PicSS3Big.Image
        End If

        m_cAVI.DrawFrame .PicSS3.hdc, pos3, lWidth:=.PicSS1.ScaleWidth, lHeight:=.PicSS1.ScaleHeight, Transparent:=False
        .PicSS3.Picture = .PicSS3.Image

    End Select
End With
End Sub

'Public Sub GetPixDD()
''ToDebug "Картинки - из базы"
'
'Set PicSS1 = Nothing: Set PicSS1Big = Nothing
'If GetPic(PicSS1Big, 1, "SnapShot1") Then PicSS1Big.Picture = PicSS1Big.Image
'    Set PicSS2 = Nothing: Set PicSS2Big = Nothing
'    If GetPic(PicSS2Big, 1, "SnapShot2") Then PicSS2Big.Picture = PicSS2Big.Image
'        Set PicSS3 = Nothing: Set PicSS3Big = Nothing
'        If GetPic(PicSS3Big, 1, "SnapShot3") Then PicSS3Big.Picture = PicSS3Big.Image
'
'Set PicFrontFace = Nothing
'Set picCanvas = Nothing
'If GetPic(PicFrontFace, 1, "FrontFace") Then
'    PicFrontFace.Picture = PicFrontFace.Image
'End If
'
'End Sub

Public Sub GetFields()
'ToDebug "Записи - из базы" для редактора
Dim tmp As String
With frmEditor
    Mark2SaveFlag = False    'не делать Mark2Save

    .TextMName.Text = CheckNoNullStr("MovieName"): .FrameAddEdit.Caption = .TextMName.Text

    .TextLabel.Text = CheckNoNullStr("Label")
    .TextGenre.Text = CheckNoNullStr("Genre")
    .TextYear.Text = CheckNoNullStr("Year")
    .TextCountry.Text = CheckNoNullStr("Country")
    .TextAuthor.Text = CheckNoNullStr("Director")
    .TextRole.Text = CheckNoNullStr("Acter")
    .TextTimeHid.Text = CheckNoNullStr("Time")
    .TextResolHid.Text = CheckNoNullStr("Resolution")
    .TextAudioHid.Text = CheckNoNullStr("Audio")
    .TextFPSHid.Text = CheckNoNullStr("FPS")
    .TextFilelenHid.Text = CheckNoNullStr("FileLen")
    .TextCDN.Text = CheckNoNullStr("CDN")
    .ComboNos.Text = CheckNoNullStr("MediaType")
    .TextVideoHid.Text = CheckNoNullStr("Video")
    .TextSubt.Text = CheckNoNullStr("SubTitle")
    .TextLang.Text = CheckNoNullStr("Language")
    .TextRate.Text = CheckNoNullStr("Rating")
    .TextFileName.Text = CheckNoNullStr("FileName")
    .TextUser.Text = CheckNoNullStr("Debtor")
    .CDSerialCur.Text = CheckNoNullStr("SNDisk")
    .TextOther.Text = CheckNoNullStr("Other")
    .TextCoverURL.Text = CheckNoNullStr("CoverPath")
    .TextMovURL.Text = CheckNoNullStr("MovieURL")
    .TextAnnotation.Text = CheckNoNullStr("Annotation")

    'для инет поиска
    tmp = Replace(.TextMName.Text, "(", vbNullString)
    tmp = Replace(tmp, ")", vbNullString)
    tmp = Replace(tmp, "/", vbNullString)
    tmp = Replace(tmp, ".", vbNullString)
    tmp = Replace(tmp, "[", vbNullString)
    tmp = Replace(tmp, "]", vbNullString)

    If Len(.TextMName.Text) <> 0 Then .TxtIName.Text = LCase$(tmp)

    'почистить комбо со списками
    .ComboGenre.Text = vbNullString
    .ComboCountry.Text = vbNullString
    .ComboOther.Text = vbNullString
    ' в .ComboNos данные

    Mark2SaveFlag = True    'вернуть Mark2Save
End With
End Sub


Public Sub ClearFields()
'очистка полей в редакторе
With frmEditor
    .TextFilelenHid.Text = "0"
    .TextTimeHid.Text = vbNullString
    'TextCompanyHid = rs.Fields("")
    .TextVideoHid.Text = vbNullString
    .TextFPSHid.Text = vbNullString
    .TextResolHid.Text = vbNullString
    .TextAudioHid.Text = vbNullString
    'TextFramesHid = rs.Fields("")
    .TextMName.Text = vbNullString
    .TextLabel.Text = vbNullString
    .TextGenre.Text = vbNullString
    .TextCountry.Text = vbNullString
    .TextYear.Text = vbNullString
    .TextAuthor.Text = vbNullString
    .TextRole.Text = vbNullString
    .TextUser.Text = vbNullString
    .TextCDN.Text = "0"
    .TextFileName.Text = vbNullString
    .TextAnnotation.Text = vbNullString
    .TextOther.Text = vbNullString
    .CDSerialCur.Text = vbNullString
    .TextRate.Text = vbNullString
    .TextLang.Text = vbNullString
    .TextSubt.Text = vbNullString
    .TextCoverURL.Text = vbNullString
    .TextMovURL.Text = vbNullString
    
    'lbInetMovieList.Clear
    
    .ComboSites.Text = vbNullString
    .ComboGenre.Text = vbNullString
    .ComboCountry.Text = vbNullString
    .ComboOther.Text = vbNullString
    
    .ComboNos.Text = vbNullString
    
    TextTimeMSHid = "0"


End With
End Sub

Public Function PutFields() As Long
'текстовые поля в базу, поля в базу, из редактора
'возвратить rs("Key") до апдейта
'Dim tmp As String

On Error GoTo err
With frmEditor

'replace2regional тут нужен? нет/ не всегда

'комбосы
If Len(.TextLabel.Text) > 255 Then rs.Fields("Label") = left$(.TextLabel.Text, 255) Else rs.Fields("Label") = .TextLabel.Text
If Len(.TextMName.Text) > 255 Then rs.Fields("MovieName") = left$(.TextMName.Text, 255) Else rs.Fields("MovieName") = .TextMName.Text
If Len(.TextGenre.Text) > 255 Then rs.Fields("Genre") = left$(.TextGenre.Text, 255) Else rs.Fields("Genre") = .TextGenre.Text
If Len(.TextYear.Text) > 255 Then rs.Fields("Year") = left$(.TextYear.Text, 255) Else rs.Fields("Year") = .TextYear.Text
If Len(.TextCountry.Text) > 255 Then rs.Fields("Country") = left$(.TextCountry.Text, 255) Else rs.Fields("Country") = .TextCountry.Text
If Len(.TextAuthor.Text) > 255 Then rs.Fields("Director") = left$(.TextAuthor.Text, 255) Else rs.Fields("Director") = .TextAuthor.Text
If Len(.TextSubt.Text) > 255 Then rs.Fields("SubTitle") = left$(.TextSubt.Text, 255) Else rs.Fields("SubTitle") = .TextSubt.Text
If Len(.TextLang.Text) > 255 Then rs.Fields("Language") = left$(.TextLang.Text, 255) Else rs.Fields("Language") = .TextLang.Text
If Len(.TextRate.Text) > 255 Then rs.Fields("Rating") = left$(.TextRate.Text, 255) Else rs.Fields("Rating") = .TextRate.Text
If Len(.TextUser.Text) > 255 Then rs.Fields("Debtor") = left$(.TextUser.Text, 255) Else rs.Fields("Debtor") = .TextUser.Text

'memo
rs.Fields("Other") = .TextOther.Text 'combo
rs.Fields("FileName") = .TextFileName.Text 'txt
rs.Fields("Annotation") = .TextAnnotation.Text 'txt mult
rs.Fields("Acter") = .TextRole.Text 'txt mult

'текст, длина текстбоксов 255
rs.Fields("Time") = .TextTimeHid.Text
rs.Fields("Resolution") = .TextResolHid.Text
rs.Fields("Audio") = .TextAudioHid.Text
rs.Fields("FPS") = .TextFPSHid.Text
rs.Fields("CDN") = .TextCDN.Text
rs.Fields("MediaType") = .ComboNos.Text
rs.Fields("Video") = .TextVideoHid.Text
' дык не опустошить... вспомнить зачем эо было нужно? If CDSerialCur <> vbNullString Then
rs.Fields("snDisk") = .CDSerialCur.Text
rs.Fields("CoverPath") = .TextCoverURL.Text
rs.Fields("MovieURL") = .TextMovURL.Text

'число
rs.Fields("FileLen") = Val(.TextFilelenHid.Text)

rs.Update

'ошибка, если сохранена первая запись в базе  PutFields = rs("Key") 'ключ на котором были до добавления нового (чисто для интереса)

End With

ToDebug " Текст в базе"
Exit Function

err:
If err <> 0 Then MsgBox err.Description
ToDebug "Err_PutField:" & err.Description
On Error Resume Next
Resume Next
End Function
Public Sub GetEditPix()
'ToDebug "Картинки - из базы" 'для редактора

With frmEditor
Set .PicSS1 = Nothing: NoPic1Flag = False
Set .PicSS1Big = Nothing
If GetPic(.PicSS1, 1, "SnapShot1") Then .PicSS1.Picture = .PicSS1.Image

Set .PicSS2 = Nothing: NoPic2Flag = False
Set .PicSS2Big = Nothing
If GetPic(.PicSS2, 1, "SnapShot2") Then
    .PicSS2.Picture = .PicSS2.Image
Else
    .PicSS2.Height = .PicSS1.Height
End If

Set .PicSS3 = Nothing: NoPic3Flag = False
Set .PicSS3Big = Nothing
If GetPic(.PicSS3, 1, "SnapShot3") Then
    .PicSS3.Picture = .PicSS3.Image
Else
    .PicSS3.Height = .PicSS1.Height
End If


Set .PicFrontFace = Nothing
Set .picCanvas = Nothing
If GetPic(.PicFrontFace, 1, "FrontFace") Then
    .PicFrontFace.Picture = .PicFrontFace.Image
    DrawCoverEdit
End If
End With
End Sub
Public Sub DrawCoverEdit()
'обложка редактора
'BitBlt picCanvas.hdc, 0, 0, PicFrontFace.Width, PicFrontFace.Height, _
  PicFrontFace.hdc, 0, 0, SRCCOPY
'picCanvas.Refresh

Dim PRatio As Double
Dim CanvasW As Single
Dim CanvasH As Single
Dim CanvasHalfW As Single
Dim CanvasHalfH As Single
Dim k As Single
Dim chOr As Boolean

With frmEditor
CanvasW = .picCanvas.Width / Screen.TwipsPerPixelX
CanvasH = .picCanvas.Height / Screen.TwipsPerPixelX
CanvasHalfW = CanvasW / 2
CanvasHalfH = CanvasH / 2

If .PicFrontFace.Picture <> 0 Then

PRatio = .PicFrontFace.Height / .PicFrontFace.Width
k = CanvasH / CanvasW
If k < PRatio Then chOr = True

If chOr Then If PRatio > 1 Then PRatio = 1 / PRatio
    If chOr Then
        'centre hor
            .picCanvas.PaintPicture .PicFrontFace.Picture, CanvasHalfW - (CanvasW * PRatio * k) / 2, 0, CanvasW * PRatio * k, CanvasH
    Else
        'centre VERT
            .picCanvas.PaintPicture .PicFrontFace.Picture, 0, CanvasHalfH - (CanvasH * PRatio / k) / 2, CanvasW, CanvasH * PRatio / k
    End If


End If
End With
End Sub

Public Sub SetFromScript()
'для редактороа
On Error Resume Next

With SC.CodeObject

    If Len(.MTitle) <> 0 Then
        If filltrue(frmEditor.TextMName) Then
            frmEditor.TextMName = .MTitle
        ElseIf filltrueAdd Then
            frmEditor.TextMName = frmEditor.TextMName & " / " & .MTitle
        End If
    End If
    If Len(.MYear) <> 0 Then
        If filltrue(frmEditor.TextYear) Then
            frmEditor.TextYear = .MYear
        ElseIf filltrueAdd Then
            frmEditor.TextYear = frmEditor.TextYear & ", " & .MYear
        End If
    End If
    If Len(.MGenre) <> 0 Then
        If filltrue(frmEditor.TextGenre) Then
            frmEditor.TextGenre = .MGenre
        ElseIf filltrueAdd Then
            frmEditor.TextGenre = frmEditor.TextGenre & ", " & .MGenre
        End If
    End If
    If Len(.MDirector) <> 0 Then
        If filltrue(frmEditor.TextAuthor) Then
            frmEditor.TextAuthor = .MDirector
        ElseIf filltrueAdd Then
            frmEditor.TextAuthor = frmEditor.TextAuthor & ", " & .MDirector
        End If
    End If
    If Len(.MActors) <> 0 Then
        If filltrue(frmEditor.TextRole) Then
            frmEditor.TextRole = .MActors
        ElseIf filltrueAdd Then
            frmEditor.TextRole = frmEditor.TextRole & ", " & .MActors
        End If
    End If
    If Len(.MDescription) <> 0 Then
        If filltrue(frmEditor.TextAnnotation) Then
            frmEditor.TextAnnotation = .MDescription
        ElseIf filltrueAdd Then
            frmEditor.TextAnnotation = frmEditor.TextAnnotation & vbCrLf & .MDescription
        End If
    End If
    If Len(.MCountry) <> 0 Then
        If filltrue(frmEditor.TextCountry) Then
            frmEditor.TextCountry = .MCountry
        ElseIf filltrueAdd Then
            frmEditor.TextCountry = frmEditor.TextCountry & ", " & .MCountry
        End If
    End If
    If Len(.MRating) <> 0 Then
        If filltrue(frmEditor.TextRate) Then
            frmEditor.TextRate = .MRating
        ElseIf filltrueAdd Then
            'TextRate = (Val(TextRate) + Val(.MRating)) / 2
            frmEditor.TextRate = (Str2Val(frmEditor.TextRate) + Str2Val(.MRating)) / 2
        End If
    End If
    If Len(.MLang) <> 0 Then
        If filltrue(frmEditor.TextLang) Then
            frmEditor.TextLang = .MLang
        ElseIf filltrueAdd Then
            frmEditor.TextLang = frmEditor.TextLang & ", " & .MLang
        End If
    End If
    If Len(.MSubt) <> 0 Then
        If filltrue(frmEditor.TextSubt) Then
            frmEditor.TextSubt = .MSubt
        ElseIf filltrueAdd Then
            frmEditor.TextSubt = frmEditor.TextSubt & ", " & .MSubt
        End If
    End If
    If Len(.MPicURL) <> 0 Then 'замещать
            frmEditor.TextCoverURL = .MPicURL
    End If
    If Len(.MOther) <> 0 Then
        If filltrue(frmEditor.TextOther) Then
            frmEditor.TextOther = .MOther
        ElseIf filltrueAdd Then
            frmEditor.TextOther = frmEditor.TextOther & vbCrLf & .MOther
        End If
    End If

    'сайт , замещать
    If Len(frmEditor.ComboSites.Text) <> 0 Then frmEditor.TextMovURL = frmEditor.ComboSites.Text

    'pix
    If (frmEditor.PicFrontFace.Picture = 0) Or (frmEditor.ChInFilFl.Value <> vbChecked) Then
        If Len(.MPicURL) <> 0 Then
            OpenURLProxy .MPicURL, "pic"
            
            'Debug.Print SC.CodeObject.MPicURL
        Else
            Set frmEditor.ImgPrCov = Nothing: Set frmEditor.PicFrontFace = Nothing: Set frmEditor.picCanvas = Nothing
            NoPicFrontFaceFlag = True
        End If
    End If

End With

If err.Number <> 0 Then ToDebug "SFS.Errors=" & CStr(err.Number <> 0)        'да нет

End Sub

Public Sub PosScroll()
Dim temp As Long

On Error Resume Next
With frmEditor
    temp = .Position.Value + 1000

    If temp <= PPMin Then
        PPMin = .Position.Value - 1000
        If PPMin < 0 Then PPMin = 0
        .PositionP.min = PPMin
        PPMax = .Position.Value + 1000
        If PPMax > Frames Then PPMax = Frames
        .PositionP.Max = PPMax
    Else
        PPMax = .Position.Value + 1000
        If PPMax > Frames Then PPMax = Frames - 1
        .PositionP.Max = PPMax
        PPMin = .Position.Value - 1000
        If PPMin < 0 Then PPMin = 0
        .PositionP.min = PPMin
    End If

    'If Position.Value > 1000 Then
    ' PositionP.Value = Position.Value 'PositionP.Min + 1000
    'Else
    .PositionP.Value = .Position.Value  'PositionP.Min
    'End If

    Screen.MousePointer = vbHourglass
    If .Position.Value = 0 Then
        pRenderFrame 1
    Else
        pRenderFrame CDbl(.Position.Value)
    End If

    Screen.MousePointer = vbNormal
End With
End Sub

Public Sub MPGPosScroll()
'Dim objPosition As IMediaPosition
'Dim objBasicVideo As IBasicVideo
'objBasicVideo.AvgTimePerFrame

'   On Error Resume Next
'    Set objPosition = mobjManager
'    Set objBasicVideo = mobjManager
'mobjManager.Stop
'Debug.Print (Position.Value - 1) / 1000
'mobjManager.StopWhenReady
'mobjManager.Pause

Dim temp As Long
'MPGPosScroll = True
'MPGPosScrollFlag = True

On Error Resume Next
With frmEditor

temp = .Position.Value + cMPGRange

If temp <= PPMin Then
    PPMin = .Position.Value - cMPGRange
    If PPMin < 0 Then PPMin = 0
    .PositionP.min = PPMin
    PPMax = .Position.Value + cMPGRange
    If PPMax > TimesX100 Then PPMax = TimesX100
    .PositionP.Max = PPMax
Else
    PPMax = .Position.Value + cMPGRange
    If PPMax > TimesX100 Then PPMax = TimesX100
    .PositionP.Max = PPMax
    PPMin = .Position.Value - cMPGRange
    If PPMin < 0 Then PPMin = 0
    .PositionP.min = PPMin
End If
.PositionP.Value = .Position.Value  'PositionP.Min

If .Position.Value = .Position.Max Then .PositionP.Value = .PositionP.Max

'Screen.MousePointer = vbHourglass

If .Position.Value = 1 Then
    objPosition.CurrentPosition = 0     '(Position.Value - 1) / 1000 + 0.04
    'mobjManager.Pause
Else
    objPosition.CurrentPosition = .Position.Value / 100     '(Position.Value - 1) / 1000 + 0.04
    'Debug.Print objPosition.CurrentPosition
    'movie.Refresh
    'objPosition.CurrentPosition = CLng(Position.Value / 100)
    'mobjManager.Pause
End If

Screen.MousePointer = vbNormal
End With
End Sub

Public Sub SaveAutoAdd()
Dim curKey As String

FirstLVFill = False

With frmEditor
    ToDebug "SaveAutoAdd"
    'ToDebug "EditMode: " & rs.EditMode

    curKey = rs("Key")    'ключ добавляемого поля. Важно переместится на него после апдейта

    ToDebug "SaveAutKey=" & curKey

    If SavePic1Flag Then
        If NoPic1Flag Then
            rs.Fields("SnapShot1") = vbNullString
            ToDebug "Pic1-no"
        Else
            If Opt_PicRealRes Then    'большую
                Pic2JPG .PicSS1Big, 1, "SnapShot1"
                ToDebug "Pic1-big"
            Else    'мелкую
                Pic2JPG .PicSS1, 1, "SnapShot1"
                ToDebug "Pic1-sm"
            End If
        End If
    End If
    If SavePic2Flag Then
        If NoPic2Flag Then
            rs.Fields("SnapShot2") = vbNullString
            ToDebug "Pic2-no"
        Else
            If Opt_PicRealRes Then    'большую
                Pic2JPG .PicSS2Big, 1, "SnapShot2"
                ToDebug "Pic2-big"
            Else    'мелкую
                Pic2JPG .PicSS2, 1, "SnapShot2"
                ToDebug "Pic2-sm"
            End If
        End If
    End If
    If SavePic3Flag Then
        If NoPic3Flag Then
            rs.Fields("SnapShot3") = vbNullString
            ToDebug "Pic3-no"
        Else
            If Opt_PicRealRes Then    'большую
                Pic2JPG .PicSS3Big, 1, "SnapShot3"
                ToDebug "Pic3-big"
            Else    'мелкую
                Pic2JPG .PicSS3, 1, "SnapShot3"
                ToDebug "Pic3-sm"
            End If
        End If
    End If
    If SaveCoverFlag Then
        If NoPicFrontFaceFlag Then
            rs.Fields("FrontFace") = vbNullString
            ToDebug "Cover-no"
        Else
            Pic2JPG .PicFrontFace, 1, "FrontFace"
            ToDebug "Cover-yes"
        End If
    End If
    SavePic1Flag = False: SavePic2Flag = False: SavePic3Flag = False: SaveCoverFlag = False

    'положить в базу поля
    PutFields    'там апдейт с возвратом позиции на до добавления


    RSGoto curKey    'встать на добавленный

    FrmMain.ListView.Sorted = False

    ReDim Preserve lvItemLoaded(FrmMain.ListView.ListItems.Count + 1)    ' 1
    Add2LV FrmMain.ListView.ListItems.Count, FrmMain.ListView.ListItems.Count + 1    '2

    CurLVKey = rs("Key") & """"

    CurSearch = GotoLV(CurLVKey)
    'пометить все автодобавляемые
    If FrmMain.ListView.ListItems.Count > 0 Then
        Set FrmMain.ListView.SelectedItem = FrmMain.ListView.ListItems(CurSearch)
    End If


    'если была сортировка - произвести ее       - путь будут в конце и без настройки Opt_SortLVAfterEdit
    'If LVSortColl > 0 Then LVSOrt (LVSortColl)
    'If LVSortColl = -1 Then SortByCheck 0, True

    If FrmMain.FrameView.Visible Then FrmMain.ListView.SelectedItem.EnsureVisible    ': LVCLICK

End With
ToDebug "...saved"
End Sub

Public Sub Mark2Save()
'краснить при изменении в редакторе
If rs.RecordCount < 1 Then Exit Sub
If BaseReadOnly Or BaseReadOnlyU Then Exit Sub

With frmEditor
If Mark2SaveFlag Then
    If .FrameAddEdit.Visible Then
        If rs.EditMode = 0 Then
        rs.Edit
        .ComSaveRec.BackColor = &HC0C0E0 'покраснить
        End If
    End If
End If
End With
End Sub

Public Function GetEdTxtFileName(s As String) As String
'кинуть в поле редактора только файл или с полным путем
If Opt_FileWithPath Then
GetEdTxtFileName = s
Else
GetEdTxtFileName = GetNameFromPathAndName(s)
End If
End Function

Public Sub DelFromEditor()
'удаление из редактора
'после встаем на след по базе (пред, если нельзя)
Dim okk As Integer
Dim ComDelEnabled As Boolean
Dim itmX As ListItem
Dim i As Long
Dim temp As Long

okk = myMsgBox(msgsvc(15), vbOKCancel, , FrmMain.hwnd)
If okk = 2 Then Exit Sub


frmEditor.ComDel.Enabled = False: ComDelEnabled = True    'не нажимать пока не закончено
ToDebug "DelKey=" & rs("Key")

ClearVideo

If rs.RecordCount > 0 Then
    With FrmMain
        '?Удалить запись в LV в соответствии ключу базы
        For Each itmX In .ListView.ListItems
            If Val(itmX.Key) = rs("Key") Then .ListView.ListItems.Remove itmX.Index: Exit For
        Next

        '?Пометить как удаленное - убрать ключ
        'For Each itmX In ListView.ListItems
        '    If Val(itmX.Key) = rs("Key") Then itmX.Key = "": Exit For
        'Next

        '.ListView.Sorted = False 'это не меняет порядок
        If LVSortColl <> 0 Then LVSOrt (lvHeaderIndexPole)    'меняем порядок на как в базе
        'перенумеровать все индексы в LV
        For i = 1 To .ListView.ListItems.Count
            .ListView.ListItems(i).SubItems(lvIndexPole) = i - 1
        Next i

        'если была сортировка - произвести ее
        If Opt_SortLVAfterEdit Then
            If LVSortColl > 0 Then LVSOrt (LVSortColl)
            If LVSortColl = -1 Then SortByCheck 0
        End If
        
        rs.Delete    'удалить в базе
    End With
    If rs.RecordCount < 1 Then
        ComDelEnabled = False: FrmMain.VerticalMenu.SetFocus
        NoListClear
        FrmMain.VerticalMenu_MenuItemClick 1, 0
        Exit Sub
    End If
Else    'rs.RecordCount <= 0

    Exit Sub
End If

'на след запись или на пред.
rs.MoveNext
If rs.EOF Then
    rs.MovePrevious
    If rs.BOF Then
        ComDelEnabled = False
        FrmMain.VerticalMenu.SetFocus
    End If
End If

If rs.EditMode Then
    editFlag = True    ' думаем, что продолжаем редактирование
Else
    'позеленить
    frmEditor.ComSaveRec.BackColor = &HC0E0C0
End If

ToDebug "...удалили"

'загрузить след.
ToDebug "LoadKey=" & rs("Key")
GetEditPix

Mark2SaveFlag = False
GetFields
Mark2SaveFlag = True

'пометить
temp = GotoLVLong(rs("Key"))
If temp > -1 Then
    Set FrmMain.ListView.SelectedItem = FrmMain.ListView.ListItems(temp)
End If

If ComDelEnabled Then frmEditor.ComDel.Enabled = True: frmEditor.ComDel.SetFocus    'вернуть
End Sub

Public Sub SaveFromEditor()
Dim curKey As String

FirstLVFill = False

ToDebug "SaveClick..."
ToDebug " EditMode: " & rs.EditMode

If rs.EditMode Then
    If Not addflag Then editFlag = True    ' думаем, что продолжаем редактирование
    'editFlag = True
Else
    If rs.RecordCount < 1 Then Exit Sub
    rs.Edit
    editFlag = True
End If

ToDebug " AddIs: " & addflag
ToDebug " EdtIs: " & editFlag

curKey = rs("Key")    'ключ добавляемого поля. Важно переместится на него после апдейта

ToDebug " SaveRecKey=" & curKey

With frmEditor
    If SavePic1Flag Then
        If NoPic1Flag Then
            rs.Fields("SnapShot1") = vbNullString
            ToDebug " Pic1 - no"
        Else
            If Opt_PicRealRes Then    'большую
                Pic2JPG .PicSS1Big, 1, "SnapShot1"
                ToDebug " Pic1 - big"
            Else    'мелкую
                Pic2JPG .PicSS1, 1, "SnapShot1"
                ToDebug " Pic1 - small"
            End If
        End If
    End If
    If SavePic2Flag Then
        If NoPic2Flag Then
            rs.Fields("SnapShot2") = vbNullString
            ToDebug " Pic2 - no"
        Else
            If Opt_PicRealRes Then    'большую
                Pic2JPG .PicSS2Big, 1, "SnapShot2"
                ToDebug " Pic2 - big"
            Else    'мелкую
                Pic2JPG .PicSS2, 1, "SnapShot2"
                ToDebug " Pic2 - small"
            End If
        End If
    End If
    If SavePic3Flag Then
        If NoPic3Flag Then
            rs.Fields("SnapShot3") = vbNullString
            ToDebug " Pic3 - no"
        Else
            If Opt_PicRealRes Then    'большую
                Pic2JPG .PicSS3Big, 1, "SnapShot3"
                ToDebug " Pic3 - big"
            Else    'мелкую
                Pic2JPG .PicSS3, 1, "SnapShot3"
                ToDebug " Pic3 - small"
            End If
        End If
    End If
    If SaveCoverFlag Then
        If NoPicFrontFaceFlag Then
            rs.Fields("FrontFace") = vbNullString
            ToDebug " Cover - no"
        Else
            Pic2JPG .PicFrontFace, 1, "FrontFace"
            ToDebug " Cover - yes"
        End If
    End If
    SavePic1Flag = False: SavePic2Flag = False: SavePic3Flag = False: SaveCoverFlag = False

    'положить в базу поля
    'Dim pf As Long
    'pf =
    PutFields  'там апдейт с возвратом позиции на до добавления
    
    'If addflag Then ToDebug " BeforeAddKey=" & pf

    .FrameAddEdit.Caption = .TextMName.Text

    If addflag Then    '

        RSGoto curKey
        FrmMain.ListView.Sorted = False
        ReDim Preserve lvItemLoaded(FrmMain.ListView.ListItems.Count + 1)    ' 1
        Add2LV FrmMain.ListView.ListItems.Count, FrmMain.ListView.ListItems.Count + 1    '2
        CurLVKey = rs("Key") & """"
        CurSearch = GotoLV(CurLVKey)
        'пометить
        'If Not frmAutoFlag Then
        If FrmMain.ListView.ListItems.Count > 0 Then
            'встать на новую
            If CurSearch <> -1 Then Set FrmMain.ListView.SelectedItem = FrmMain.ListView.ListItems(CurSearch)
        End If
        'End If

        FrmMain.LVCLICK 'отразить
        addflag = False
    End If

    If editFlag Then
        FrmMain.ListView.Sorted = False
        'редактировать lv
        CurLVKey = rs("Key") & """"
        CurSearch = GotoLV(CurLVKey)
        EditLV CurSearch

        'сказать, что еще не показаны сабы
        If Opt_LoadOnlyTitles = True Then    ' только названия
            lvItemLoaded(CurSearch) = False
        End If

        'пометить 'нужно тк мб тут же редактирование и запись в CurSearch
        If Not (FrmMain.ListView.SelectedItem Is Nothing) Then
        Set FrmMain.ListView.SelectedItem = FrmMain.ListView.ListItems(CurSearch)
        End If
        
        FrmMain.LVCLICK 'отразить (надо. нет InitFlag - нет клика из меню просмотра)
    End If    'editflag

    'если была сортировка - произвести ее
    If Opt_SortLVAfterEdit Then
        If LVSortColl > 0 Then LVSOrt (LVSortColl)
        If LVSortColl = -1 Then SortByCheck 0, True
    End If

If Not (FrmMain.ListView.SelectedItem Is Nothing) Then
    If FrmMain.FrameView.Visible Then FrmMain.ListView.SelectedItem.EnsureVisible    ': LVCLICK
End If

    'On Error GoTo 0

    AutoFillStore    'запомнить поля в списки для последующего авто ввода

    .ComDel.Visible = True: .ComDel.Enabled = True: .ComSaveRec.BackColor = &HC0E0C0
End With
ToDebug "...сохранили"
End Sub
