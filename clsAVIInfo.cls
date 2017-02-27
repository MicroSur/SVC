VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAVIInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'       AVI info version 1.042
'  Written by Mike D Sutton of EDais
'     Microsoft Visual Basic MVP
'
' E-Mail: EDais@mvps.org
' WWW: Http://www.mvps.org/EDais/
'
' Written: 25/06/2002
' Last edited: 29/07/2003


Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Type typMainAVIHeader    ' 56 bytes
    dwMicroSecPerFrame As Long
    dwMaxBytesPerSec As Long
    dwReserved1 As Long
    dwFlags As Long
    dwTotalFrames As Long
    dwInitialFrames As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
End Type

Private Type typAVIStreamHeader    '48 bytes
    fccType As String * 4
    fccHandler As String * 4
    dwFlags As Long
    dwReserved1 As Long
    dwInitialFrames As Long
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
End Type

Private Type typRIFFFormHeader
    rfhSignature As String * 4
    rfhSize As Long
    rfhSpecialization As String * 4
End Type

Private Type typChunkHeader
    ChName As String * 4
    chSize As Long
End Type

Private Type typListHeader
    lhSignature As String * 4
    lhCount As Long
    lhIdentity As String * 4
End Type

'Public Type typCodecInf
'    FourCC As String * 4
'    Description As String
'End Type

Private Type BITMAPINFOHEADER    ' 40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD    ' 32-bit colour
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As RGBQUAD
End Type

Private Type WAVEFORMAT    ' 14 bytes
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
End Type

Private Type PCMWAVEFORMAT    ' 16 bytes
    wf As WAVEFORMAT
    wBitsPerSample As Integer
End Type

Private Type typInfoEntry
    ieTag As Long
    ieData As String
End Type

'Public CodecList() As typCodecInf
Public NumCodec As Long
Dim InfoEntryList() As typInfoEntry
Dim NumInfos As Long

' Member variables
Dim m_Width As Long
Dim m_Height As Long
Dim m_FrameRate As Double
Dim m_PlayLength As Currency    'Long
Dim m_NumFrames As Long
Dim m_NumStreams As Integer
Dim m_HasVideo As Boolean
Dim m_HasAudio As Boolean
Dim m_VideoCodec As String

Dim m_StreamVideo_dwLength As Long    'sur длина видеопотока фреймов

Dim m_VideoCodec2 As String
Dim m_AudioCodec As String
Dim m_BitDepth As Integer
Dim m_Channels As Integer
Dim m_SamplesPerSec As Long
Dim m_AudioFormat As String
Dim m_FileName As String
Dim m_FileSize As Long
Dim m_KBitPerSec As Long    'Currency
Dim v_KBitPerSec As Long    'Currency
Dim m_IndexOffset As Long
Dim m_MovieOffset As Long
Dim m_PadBytes As Long
Dim m_All_Audios As String
Dim m_rate As Long
Dim m_BytePerSec As Long


Private Const RIFFHeader As String = "RIFF"
Private Const AVIHeader As String = "AVI "
Private Const ListChunk As String = "LIST"
Private Const FormatChunk As String = "hdrl"
Private Const HeaderChunk As String = "avih"
Private Const StreamHeader As String = "strh"
Private Const StreamFormat As String = "strf"
Private Const VideoTag As String = "vids"
Private Const AudioTag As String = "auds"
Private Const JunkTag As String = "JUNK"    ' Standard junk tag
Private Const JunQTag As String = "JUNQ"    ' Premiere junk tag
Private Const MovieChunk As String = "movi"
Private Const IndexChunk As String = "idx1"

Private Const InfoList As String = "INFO"    ' Info list
''интересные
'Private Const InfoName      As String = "INAM" 'Name/Title
'Private Const InfoArti      As String = "IART" 'Artist , Director
'Private Const InfoStarr     As String = "ISTR" 'Starring
'Private Const InfoLang1     As String = "IAS1" 'First Language
'Private Const InfoLang2     As String = "IAS2"
'Private Const InfoLang3     As String = "IAS3"
'Private Const InfoLang4     As String = "IAS4"
'Private Const InfoLang5     As String = "IAS5"
'Private Const InfoLang6     As String = "IAS6"
'Private Const InfoLang7     As String = "IAS7"
'Private Const InfoLang8     As String = "IAS8"
'Private Const InfoLang9     As String = "IAS9"
'Private Const InfoLng       As String = "ILNG" 'Language ?
'Private Const InfoCountry   As String = "ICNT" 'Страна
'Private Const InfoStudio    As String = "ISTD" 'Production studio
'Private Const InfoRating    As String = "IRTD" 'Rating 'наверно возростной
'Private Const InfoGenr      As String = "IGNR" 'Genre
'Private Const InfoGenre2    As String = "ISGN" 'Secondary genre
'Private Const InfoWeb       As String = "IWEB" 'Internet address
'Private Const InfoSoft      As String = "ISFT" 'Software
'Private Const InfoCrea      As String = "ICRD" '  Creation Date (YYYYMMDD) наверно это год фильма

''еще мета данные
'Private Const InfoCinema  As String = "ICNM" 'Кинематограф
'Private Const InfoArch    As String = "IARL" 'Archival location
'Private Const InfoAudDef  As String = "ICAS" 'Default audio stream
'Private Const InfoCostum  As String = "ICDS" 'Дизайнер по костюмам
'Private Const InfoDigit   As String = "IDIT" 'Дата оцифровки
'Private Const InfoDistr   As String = "IDST" 'Дистрибютер
'Private Const InfoEdit    As String = "IEDT" 'Редактор
'Private Const InfoEncod   As String = "IENC" 'Encoded by
'Private Const InfoFrames  As String = "IFRM" 'Total NN Frames
'Private Const InfoLogo    As String = "ILGU" 'Logo URL
'Private Const InfoLogoI   As String = "ILIU" 'Logo Icon URL
'Private Const InfoMoreI   As String = "IMBI" 'More info banner icon url
'Private Const InfoMore    As String = "IMBU" 'More info banner url
'Private Const InfoTextI   As String = "IMIT" 'More info text
'Private Const InfoText    As String = "IMIU" 'More info url
'Private Const InfoMus     As String = "IMUS" 'Music by
'Private Const InfoDesc    As String = "IPDS" 'Production descr
'Private Const InfoProduce As String = "IPRO" 'Produced by
'Private Const InfoPart    As String = "IPRT" 'Part
'Private Const InfoRip     As String = "IRIP" 'Ripped by
'Private Const InfoWMark   As String = "IWMU" 'Watermark url
'Private Const InfoWrit    As String = "IWRI" 'Written by
'Private Const InfoNFOA    As String = "NFO/" 'Info (Adobe Premiere)
'Private Const InfoComi    As String = "ICMS" ' Info -> Commissioned
'Private Const InfoComm    As String = "ICMT" ' Info -> Comments
'Private Const InfoCopy    As String = "ICOP" ' Info -> Copyright
'Private Const InfoCrop    As String = "ICRP" ' Info -> Cropped
'Private Const InfoDimn    As String = "IDIM" ' Info -> Dimensions
'Private Const InfoDots    As String = "IDPI" ' Info -> Dots Per Inch
'Private Const InfoEngi    As String = "IENG" ' Info -> Engineer
'Private Const InfoIDTh    As String = "IID3" ' Info -> ID3 tag
'Private Const InfoKeyw    As String = "IKEY" ' Info -> Keywords
'Private Const InfoLigt    As String = "ILGT" ' Info -> Lightness
'Private Const InfoLyrc    As String = "ILYC" ' Info -> Lyrics
'Private Const InfoMedi    As String = "IMED" ' Info -> Medium
'Private Const InfoMPTh    As String = "IMP3" ' Info -> MP3 information
'Private Const InfoPalt    As String = "IPLT" ' Info -> Palette Setting
'Private Const InfoProd    As String = "IPRD" ' Info -> Product
'Private Const InfoSubj    As String = "ISBJ" ' Info -> Subject
'Private Const InfoShrp    As String = "ISHP" ' Info -> Sharpness
'Private Const InfoSrce    As String = "ISRC" ' Info -> Source
'Private Const InfoSrcF    As String = "ISRF" ' Info -> Source Form
'Private Const InfoTech    As String = "ITCH" ' Info -> Technician, encoded by(2)
'Private Const InfoDisp    As String = "DISP" ' Info -> Sound Scheme Title

Public Event ReadError(ByVal outDesc As String)

' Public interface to member variables
Public Property Get Width() As Long
Width = m_Width
End Property

Public Property Get Height() As Long
Height = m_Height
End Property

Public Property Get BitDepth() As Integer
If (m_HasVideo) Then BitDepth = m_BitDepth
End Property

Private Function Chnnls() As String
If (m_HasAudio) Then
    Select Case m_Channels
    Case 1
        Chnnls = "Mono"
    Case 2
        Chnnls = "Stereo"
    Case Else
        Chnnls = m_Channels & "ch"
    End Select
End If
End Function
Public Property Get Channels() As String
If (m_HasAudio) Then
    Select Case m_Channels
    Case 1
        Channels = "Mono"
    Case 2
        Channels = "Stereo"
    Case Else
        Channels = m_Channels & "ch"
    End Select
End If
End Property

Public Property Get SamplesPerSec() As Long
If (m_HasAudio) Then SamplesPerSec = m_SamplesPerSec
End Property

Public Property Get FrameRate() As Double
If (m_HasVideo) Then FrameRate = m_FrameRate
End Property
Public Property Get AllAudio() As String
If (m_HasAudio) Then AllAudio = m_All_Audios
End Property

Public Property Get PlayLength() As Long
PlayLength = m_PlayLength
End Property

Public Property Get numFrames() As Long
If (m_HasVideo) Then numFrames = m_NumFrames
End Property

Public Property Get NumStreams() As Integer
NumStreams = m_NumStreams
End Property

Public Property Get HasVideo() As Boolean
HasVideo = m_HasVideo
End Property

Public Property Get HasAudio() As Boolean
HasAudio = m_HasAudio
End Property
Public Property Get Rate() As String
If (m_HasVideo) Then Rate = m_rate
End Property



Public Property Get VideoCodec() As String
If (m_HasVideo) Then VideoCodec = m_VideoCodec
End Property
Public Property Get VideoCodec2() As String
If (m_HasVideo) Then
    If Len(m_VideoCodec2) <> 0 Then
        'VideoCodec2 = " (" & m_VideoCodec2 & ")"
        VideoCodec2 = " / " & m_VideoCodec2    '& ")"
    Else
        VideoCodec2 = vbNullString
    End If
End If
End Property
Public Property Get VideoCodec22() As String
If (m_HasVideo) Then
    If Len(m_VideoCodec2) <> 0 Then
        VideoCodec22 = m_VideoCodec2
    Else
        VideoCodec22 = vbNullString
    End If
End If
End Property
Public Property Get AudioCodec() As String
If (m_HasAudio) Then AudioCodec = m_AudioCodec
End Property

'Public Property Get NumCodecs() As Long
'    If (NumCodec = 0) Then Call MakeCodecList
'    NumCodecs = NumCodec
'End Property

Public Property Get AudioFormat() As String
If (m_HasAudio) Then AudioFormat = m_AudioFormat
End Property
Public Property Get VideoBitrate() As String
If (m_HasVideo) Then
    If Len(v_KBitPerSec) <> 0 Then
        VideoBitrate = " (~" & v_KBitPerSec & "kbps)"
    Else
        VideoBitrate = vbNullString
    End If
End If
End Property



Public Property Get filename() As String
filename = m_FileName
End Property

Public Property Get FileSize() As Long
FileSize = m_FileSize
End Property

Public Property Get KBytePerSec() As Long
If (m_HasAudio) Then KBytePerSec = m_KBitPerSec
End Property

Public Property Get IndexOffset() As Long
IndexOffset = m_IndexOffset
End Property

Public Property Get MovieOffset() As Long
MovieOffset = m_MovieOffset
End Property

Public Property Get PadBytes() As Long
PadBytes = m_PadBytes
End Property

' Public methods
Public Function ReadFile(ByRef inFile As String) As Boolean
Dim VBRHeader(4) As Byte     'sur

Dim FNum As Integer
Dim FileSize As Long
Dim AVIInfo As typMainAVIHeader
Dim AVIHead As typRIFFFormHeader
Dim AVIList As typListHeader
Dim AVIChunk As typChunkHeader
Dim AVIStream As typAVIStreamHeader
'mzt Dim PadData() As Byte
Dim AVIBMI As BITMAPINFO
Dim AVIWave As PCMWAVEFORMAT
Dim GetPal As Long
Dim LastPos As Long
Dim InfoLeft As Long
Dim ReadInfo As String
Dim AudioStreamsKBitPerSec As Long     ' Currency
Dim plusik As String
'mzt Dim addon As Integer     'непонятная компенсация при расчете битрейта
Dim HeadersSize As Long     'Currency 'Long
Dim MovieChunkSize As Long     'Currency 'Long
'    Dim AudioStreamsSize As Long ''Currency
Dim MovieSize As Long     'Currency

'MovieSize = MovieChunkSize - HeadersSize - AudioStreamsSize
'    Dim AudioStreamSizeCurrent As Long
Dim iaud As Integer    'кол-во аудио стримов

'AudioStreamSizeCurrent = 0
AudioStreamsKBitPerSec = 0

Const OneMillion As Long = 1000000

Call ClearInfo

On Error Resume Next


'    If isWindowsNt Then '- один фиг не понимает огромные файлы
'        Dim Pointer As Long, lpFSHigh As Currency
'        Pointer = lopen(inFile, OF_READ)
'        'size of the file
'        GetFileSizeEx Pointer, lpFSHigh
'        FileSize = lpFSHigh * 10000
'        lclose Pointer
'    Else
FileSize = FileLen(inFile)
'    End If

If (FileSize > 0) Then
    FNum = FreeFile()

    '        Open inFile For Binary Access Read Lock Write As #FNum
    Open inFile For Binary Access Read As #FNum

    On Error GoTo err

    Get #FNum, , AVIHead         ' Grab first set of AVI file headers
    Get #FNum, , AVIList
    Get #FNum, , AVIChunk
    Get #FNum, , AVIInfo

    ' Debug.Print AVIHead.rfhSize


    ' Validate headers
    If ((AVIHead.rfhSignature <> RIFFHeader) Or _
        (AVIHead.rfhSpecialization <> AVIHeader) Or _
        (AVIList.lhSignature <> ListChunk) Or _
        (AVIList.lhIdentity <> FormatChunk) Or _
        (AVIChunk.ChName <> HeaderChunk)) Then
        ToDebug "Err_IncAVIheader"

    Else         ' Assume by this point that this is a valid AVI file
        LastPos = Seek(FNum)

        Do    ' Start IFF chunk parser
            DoEvents
            Get #FNum, , AVIChunk

            ' Word align chunk size
            AVIChunk.chSize = ((AVIChunk.chSize + 1) \ 2) * 2

            ' Check for 'bad' chunk headers (<0 bytes)
            If (AVIChunk.chSize < 0) Then Exit Do

            Select Case AVIChunk.ChName
            Case ListChunk
                AVIList.lhSignature = AVIChunk.ChName
                AVIList.lhCount = AVIChunk.chSize
                Get #FNum, , AVIList.lhIdentity

                Select Case AVIList.lhIdentity
                Case MovieChunk

                    'попытка найти vbr 1
                    Get #FNum, Seek(FNum) + 44, VBRHeader()
                    If left$(StrConv(VBRHeader(), vbUnicode, LCID), 4) = "Xing" Then
                        'ToDebug "Audio - VBR"
                    Else
                        m_All_Audios = Replace(m_All_Audios, "(~", "(")
                    End If
                    'Debug.Print Left$(StrConv(VBRHeader(), vbUnicode, LCID), 4)
                    'If Left$(StrConv(VBRHeader(), vbUnicode, LCID), 4) = "Xing" Then VBR = " VBR"
                    'Он находится позже звука, блин. И еще тестить если несколько аудио

                    m_MovieOffset = Seek(FNum) - 5
                    LastPos = LastPos + AVIList.lhCount + 8

                    'попытка найти vbr 2
                    'Get #FNum, m_MovieOffset + 49, VBRHeader()
                    '


                    '''
                    MovieChunkSize = AVIChunk.chSize
                    'MovieChunkSize = LastPos - m_MovieOffset
                    Seek #FNum, LastPos
                Case InfoList                 ' Info chunk - Don't offset since we want to read the entries
                    InfoLeft = AVIChunk.chSize
                End Select

                LastPos = LastPos + Len(AVIList)

            Case StreamHeader
                Get #FNum, , AVIStream

                'Debug.Print AVIStream.dwLength
                'Debug.Print AVIStream.dwFlags, AVIStream.dwInitialFrames, AVIStream.dwLength, _
                 AVIStream.dwQuality, AVIStream.dwRate, AVIStream.dwReserved1, AVIStream.dwSampleSize, _
                 AVIStream.dwScale, AVIStream.dwStart, AVIStream.dwSuggestedBufferSize, _
                 AVIStream.fccHandler, AVIStream.fccType

            Case StreamFormat
                'Video
                If (AVIStream.fccType = VideoTag) Then
                    m_HasVideo = True
                    m_VideoCodec = AVIStream.fccHandler

                    m_StreamVideo_dwLength = AVIStream.dwLength
                    'Debug.Print AVIStream.dwLength 'длительность видео потока (фреймов)

                    'Debug.Print "AVIStream"
                    'Debug.Print AVIStream.dwRate
                    'Debug.Print AVIStream.dwFlags
                    'Debug.Print AVIStream.dwInitialFrames
                    'Debug.Print AVIStream.dwLength 'фреймов ?
                    'Debug.Print AVIStream.dwQuality
                    'Debug.Print AVIStream.dwSampleSize
                    'Debug.Print AVIStream.dwScale


                    Get #FNum, , AVIBMI.bmiHeader

                    With AVIBMI.bmiHeader
                        If (.biBitCount < 9) Then             ' Read Palette
                            If (.biClrUsed < 2) Then .biClrUsed = (2 ^ .biBitCount)
                            ReDim AVIBMI.bmiColors(.biClrUsed - 1) As RGBQUAD

                            For GetPal = 0 To .biClrUsed - 1
                                Get #FNum, , AVIBMI.bmiColors(GetPal)
                            Next GetPal
                        End If

                        'If the video codec was not set in the stream header, use the one defined here
                        m_VideoCodec2 = LongToFourCC(.biCompression)
                        If (m_VideoCodec = String(4, vbNullChar)) Then m_VideoCodec = m_VideoCodec2
                        If (m_VideoCodec = String(4, " ")) Then m_VideoCodec = m_VideoCodec2

                    End With

                    'аудио
                ElseIf AVIStream.fccType = AudioTag Then

                    m_HasAudio = True
                    m_AudioCodec = AVIStream.fccHandler
                    'Debug.Print AVIStream.dwLength 'длительность потока ацдио

                    Get #FNum, , AVIWave.wf
                    If (AVIChunk.chSize >= 16) Then Get #FNum, , AVIWave.wBitsPerSample
                    m_AudioFormat = AudioCodecFromFormat(AVIWave.wf.wFormatTag)

                    With AVIWave.wf
                        m_Channels = .nChannels
                        m_SamplesPerSec = .nSamplesPerSec
                        m_KBitPerSec = .nAvgBytesPerSec / 1000 * 8
                    End With
                    AudioStreamsKBitPerSec = AudioStreamsKBitPerSec + m_KBitPerSec        ' для видеобитрейта

                    'Debug.Print AVIStream.dwLength

                    iaud = iaud + 1
                    If m_All_Audios = vbNullString Then plusik = vbNullString Else plusik = ", "
                    ' сумма аудио данных
                    m_All_Audios = m_All_Audios & plusik & m_SamplesPerSec & " " & Chnnls & " " & m_AudioFormat & " (~" & m_KBitPerSec & "kbps)"

                End If             'AVIStream.fccType


            Case IndexChunk             ' Found movie index tag
                m_IndexOffset = Seek(FNum) - (Len(AVIChunk) + 1)
                LastPos = LastPos - (Len(AVIChunk) + 4)             ' *Shrug* dont ask, it doesn't work without it..

            Case JunkTag, JunQTag             ' Found junk tag, increment padding count
                m_PadBytes = m_PadBytes + AVIChunk.chSize

            Case Else             ' Un-defined chunk
                If (InfoLeft > 0) Then             ' Info chunk
                    ReadInfo = Space$(AVIChunk.chSize)
                    Get #FNum, , ReadInfo              ' Read chunk data and add to list
                    ReDim Preserve InfoEntryList(NumInfos) As typInfoEntry
                    InfoEntryList(NumInfos).ieTag = FourCCToLong(AVIChunk.ChName)
                    InfoEntryList(NumInfos).ieData = TrimNull(ReadInfo)
                    NumInfos = NumInfos + 1
                End If
            End Select

            ' Avoid overflows by checking to make sure we can accomodate the next chunk
            If (AVIChunk.chSize >= ((2 ^ 31) - (LastPos + 1))) Then
                ' RaiseEvent ReadError("File too large to parse")
                Exit Do
            End If

            If (AVIChunk.ChName <> ListChunk) Then         ' Offset to next chunk
                LastPos = LastPos + AVIChunk.chSize + Len(AVIChunk)

                If (InfoLeft) Then
                    InfoLeft = InfoLeft - (AVIChunk.chSize + Len(AVIChunk))
                    If (InfoLeft < 1) Then InfoLeft = 0
                End If

                Seek #FNum, LastPos
            End If
            ' ^^^^^^^^^^
        Loop Until LastPos >= FileSize


        If (m_MovieOffset) Then
            With AVIInfo         ' Set information about the AVI file to the class properties

                'Debug.Print .dwMaxBytesPerSec = 0
                'Debug.Print .dwRate =0
                'm_rate = .dwInitialFrames=0

                m_Width = .dwWidth
                m_Height = .dwHeight
                m_FrameRate = 1 / .dwMicroSecPerFrame * 1000000

                If .dwTotalFrames <> m_StreamVideo_dwLength Then
                    'применить длительность видеопотока
                    m_PlayLength = CCur((.dwMicroSecPerFrame / OneMillion) * m_StreamVideo_dwLength)
                    m_NumFrames = m_StreamVideo_dwLength
                Else
                    m_PlayLength = CCur((.dwMicroSecPerFrame / OneMillion) * .dwTotalFrames)
                    m_NumFrames = .dwTotalFrames
                End If

                m_NumStreams = .dwStreams
            End With

            m_BitDepth = AVIBMI.bmiHeader.biBitCount

            m_FileName = inFile
            m_FileSize = FileSize

            'VideoStreamSize = (m_FileSize / 1000 * 8 - AudioStreamsSize - HeadersSizeBit) ' - AVIStream.dwLength / 8)
            'битрейт видео
            ' v_KBitPerSec = VideoStreamSize / m_PlayLength

            HeadersSize = m_NumFrames * 8 * m_NumStreams
            MovieSize = MovieChunkSize - HeadersSize    '- AudioStreamsSize
            v_KBitPerSec = MovieSize / m_PlayLength / 1000 * 8 - AudioStreamsKBitPerSec
            'Debug.Print "1 A: " & AudioStreamsKBitPerSec & " V: " & v_KBitPerSec

        Else         ' Clear anything that was set during the parse
            Call ClearInfo
        End If         'm_MovieOffset

        ReadFile = m_MovieOffset


    End If         'Validate headers

    Close #FNum
Else
    ToDebug "Файл очень большой (или не AVI) для WFV"
    'RaiseEvent ReadError("Файл не найден или пуст.")
End If     'FileSize

Exit Function
err:
ToDebug "aviReadFile: " & err.Description
End Function

Public Function CodecToName(ByRef inCodec As String) As String
Dim inCodecSvc As String
'    Dim LoopCodec As Long
If (Len(inCodec) <> 4) Then Exit Function
'    If (NumCodec = 0) Then Call MakeCodecList
'    For LoopCodec = 0 To NumCodec - 1
'        If (StrComp(UCase$(inCodec), CodecList(LoopCodec).FourCC) = 0) Then
'            CodecToName = CodecList(LoopCodec).Description
'            Exit For
'        End If
'    Next LoopCodec
'    If (CodecToName = vbNullString) Then CodecToName = "?"
inCodecSvc = inCodec
If inCodec = String(4, vbNullChar) Then inCodecSvc = "____"
If InStr(inCodec, " ") Then inCodecSvc = Replace(inCodec, " ", "_")

CodecToName = VBGetPrivateProfileString("VIDEO", inCodecSvc, CodecsFileName, "?")

End Function

Public Function LongToFourCC(ByVal inLong As Long) As String
Dim TempArr(3) As Byte
Call RtlMoveMemory(ByVal VarPtr(TempArr(0)), ByVal VarPtr(inLong), 4)
LongToFourCC = StrConv(TempArr(), vbUnicode, LCID)

Dim i As Long
Dim txtlen As Long
Dim ch As String

txtlen = Len(LongToFourCC)
For i = 1 To txtlen
    ch = Mid$(LongToFourCC, i, 1)
    If ch < " " Then Mid$(LongToFourCC, i, 1) = " "
    'If ch < vbCr Then Mid$(LongToFourCC, i, 1) = " "
    'If ch = vbCr Then Mid$(txt, i, 1) = " " 'vbCrLf
    'Or ch > "~"
Next i

LongToFourCC = Trim$(LongToFourCC)
End Function

Public Function FourCCToLong(ByVal inString As String) As Long
Dim TempArr() As Byte

If (Len(inString) = 4) Then
    ReDim TempArr(3) As Byte
    TempArr() = StrConv(inString, vbFromUnicode, LCID)
    Call RtlMoveMemory(ByVal VarPtr(FourCCToLong), ByVal VarPtr(TempArr(0)), 4)
End If
End Function

Public Function GetInfoList(ByRef outArr() As Long) As Long
Dim LoopInfos As Long

If (NumInfos) Then
    ReDim outArr(NumInfos - 1) As Long

    For LoopInfos = 0 To NumInfos - 1
        outArr(LoopInfos) = InfoEntryList(LoopInfos).ieTag
    Next LoopInfos
End If

GetInfoList = NumInfos
End Function

Public Function QueryInfo(ByVal inInfo As Long) As String
Dim LoopInfos As Long
'mzt     Dim ConvTag As Long

If (Len(inInfo) <> 4) Then Exit Function    ' All info tags are 4 characters

If (NumInfos) Then
    For LoopInfos = 0 To NumInfos - 1
        If (inInfo = InfoEntryList(LoopInfos).ieTag) Then
            QueryInfo = InfoEntryList(LoopInfos).ieData
            Exit For
        End If
    Next LoopInfos
End If
End Function

'Public Function GetInfoName(ByVal inInfo As Long) As String
'    Dim InfoChar As String
'    If (Len(inInfo) <> 4) Then Exit Function
'    InfoChar = LongToFourCC(inInfo)
'    Select Case InfoChar
'        Case InfoList: GetInfoName = "Info list"
'        Case InfoArch: GetInfoName = "Archival location"
'        Case InfoArti: GetInfoName = "Artist"
'        Case InfoComi: GetInfoName = "Commissioned"
'        Case InfoComm: GetInfoName = "Comments"
'        Case InfoCopy: GetInfoName = "Copyright"
'        Case InfoCrea: GetInfoName = "Creation date"
'        Case InfoCrop: GetInfoName = "Cropped"
'        Case InfoDimn: GetInfoName = "Dimensions"
'        Case InfoDots: GetInfoName = "Dots per inch"
'        Case InfoEngi: GetInfoName = "Engineer"
'        Case InfoGenr: GetInfoName = "Genre"
'        Case InfoIDTh: GetInfoName = "ID3 tag"
'        Case InfoKeyw: GetInfoName = "Keywords"
'        Case InfoLigt: GetInfoName = "Lightness"
'        Case InfoLyrc: GetInfoName = "Lyrics"
'        Case InfoMedi: GetInfoName = "Medium"
'        Case InfoMPTh: GetInfoName = "MP3"
'        Case InfoName: GetInfoName = "Name/Title"
'        Case InfoPalt: GetInfoName = "Palette setting"
'        Case InfoProd: GetInfoName = "Product"
'        Case InfoSubj: GetInfoName = "Subject"
'        Case InfoSoft: GetInfoName = "Software"
'        Case InfoShrp: GetInfoName = "Sharpness"
'        Case InfoSrce: GetInfoName = "Source"
'        Case InfoSrcF: GetInfoName = "Source form"
'        Case InfoTech: GetInfoName = "Technician"
'        Case InfoDisp: GetInfoName = "Sound Scheme title"
'        Case Else ' Unknown info tag..
'            If (UCase$(Left$(inInfo, 1)) = "I") Then _
             '                GetInfoName = Right$(inInfo, 3) _
             '            Else GetInfoName = inInfo
'    End Select
'End Function

Private Function ClearInfo()
m_Width = 0    ' Clear any current (Public) information
m_Height = 0
m_FrameRate = 0
m_rate = 0
m_PlayLength = 0
m_NumFrames = 0
m_NumStreams = 0
m_HasVideo = False
m_HasAudio = False
m_VideoCodec = vbNullString
m_VideoCodec2 = vbNullString
m_AudioCodec = vbNullString
m_BitDepth = 0
m_Channels = 0
m_SamplesPerSec = 0
m_AudioFormat = vbNullString
m_FileName = vbNullString
m_FileSize = 0
m_BytePerSec = 0
m_IndexOffset = 0
m_MovieOffset = 0
m_PadBytes = 0
NumInfos = 0
ReDim InfoEntryList(0) As typInfoEntry
End Function

Private Function TrimNull(ByVal inString As String) As String
Dim NullPos As Long

NullPos = InStr(1, inString, vbNullChar)

If (NullPos) Then _
   TrimNull = left$(inString, NullPos - 1) _
   Else TrimNull = inString
    End Function

Private Function AudioCodecFromFormat(ByVal inFormat As Integer) As String
Dim UnsignedFormat As Long
Call RtlMoveMemory(ByVal VarPtr(UnsignedFormat), ByVal VarPtr(inFormat), 2)

'    Select Case UnsignedFormat ' /* WAVE form wFormatTag IDs */ -- Converted from "mmreg.h" version 1.58 on 10/07/2002
AudioCodecFromFormat = VBGetPrivateProfileString("AUDIO", "&H" & Hex(UnsignedFormat), CodecsFileName, "?")

End Function

' Private methods
'Private Function MakeCodecList()
'    NumCodec = 0
'Call AddCodec(CodecList(), NumCodec, String(4, vbNullChar), "Uncompressed")
'    ' Interpreted in part from "http://www.webartz.com/fourcc/" on 26/06/2002
'    MakeCodecList = NumCodec
'End Function

'Private Sub AddCodec(ByRef inList() As typCodecInf, ByRef inSize As Long, _
 '    ByVal inFourCC As String, Optional ByVal inDescription As String = vbNullString)
'    ReDim Preserve inList(inSize) As typCodecInf
'    inList(inSize).FourCC = inFourCC
'    If (inDescription = vbNullString) Then _
     '        inList(inSize).Description = inList(inSize - 1).Description _
     '        Else inList(inSize).Description = inDescription
'    inSize = inSize + 1
'End Sub

