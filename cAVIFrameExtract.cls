VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAVIFrameExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'mzt Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
'mzt Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
'mzt Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
'mzt Private Const FORMAT_MESSAGE_FROM_STRING = &H400
'mzt Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
'mzt Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
'mzt Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
'mzt Private Declare Function FormatMessage Lib "KERNEL32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
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

'mzt Private Type BITMAPINFO2
'    bmiHeader As BITMAPINFOHEADER
'    bmiColors(0 To 1) As RGBQUAD
'End Type
'mzt Private Type BITMAPINFO16
'    bmiHeader As BITMAPINFOHEADER
'    bmiColors(0 To 15) As RGBQUAD
'End Type
'mzt Private Type BITMAPINFO256
'mzt     bmiHeader As BITMAPINFOHEADER
'mzt     bmiColors(0 To 255) As RGBQUAD
'mzt End Type

'mzt Private Const ERR_BASE As Long = vbObjectError + 1024 + 77561
'mzt Private Const STREAM_TYPE_VIDEO = &H73646976 ' reads "vids"

Private Const OF_READ = &H0
Private Const OF_SHARE_EXCLUSIVE = &H10

'mzt Private Const AVIIF_LIST = &H1                 ' // chunk is a 'LIST'
'mzt Private Const AVIIF_KEYFRAME = &H10      '// this frame is a key frame.
'mzt Private Const AVIIF_FIRSTPART = &H20     '// this frame is the start of a partial frame.
'mzt Private Const AVIIF_LASTPART = &H40      '// this frame is the end of a partial frame.
'mzt Private Const AVIIF_MIDPART = (AVIIF_LASTPART Or AVIIF_FIRSTPART)
'mzt Private Const AVIIF_NOTIME = &H100     '// this frame doesn't take any time
'mzt Private Const AVIIF_COMPUSE = &HFFF0000      ' // these bits are for compressor use

Private Type TAVISTREAMINFO ' this is the ANSI version
    fccType As Long
    fccHandler As Long
    dwFlags As Long         '/* Contains AVITF_* flags */
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long ' /* dwRate / dwScale == samples/second */
    dwStart As Long
    dwLength As Long '; /* In units above... */
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As RECT
    dwEditCount As Long
    dwFormatChangeCount As Long
    szName(0 To 63) As Byte
End Type

Private Declare Sub AVIFileInit Lib "avifil32.dll" ()
Private Declare Sub AVIFileExit Lib "avifil32.dll" ()
'mzt Private Declare Function AVIStreamOpenFromFile Lib "avifil32.dll" Alias "AVIStreamOpenFromFileA" ( _
      ppavi As Any, ByVal szFile As String, _
      ByVal fccType As Long, ByVal lParam As Long, _
      ByVal mode As Long, pclsidHandler As Any _
   ) As Long
'mzt Private Declare Function AVIFileCreateStream Lib "avifil32.dll" ( _
      pfile As Any, _
      ppavi As Any, _
      psi As TAVISTREAMINFO _
   ) As Long
'mzt Private Declare Function AVIStreamSetFormat Lib "avifil32.dll" ( _
      pavi As Any, _
      ByVal lPos As Long, _
      lpFormat As Any, _
      ByVal cbFormat As Long _
   ) As Long
'Private Declare Function AVIStreamRelease Lib "avifil32.dll" (pavi As Any) As Long
'Private Declare Function AVIStreamLength Lib "avifil32.dll" (pavi As Any) As Long
'Private Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" ( _
         pavi As Any, lpbiWanted As Any _
      ) As Long
Private Declare Function AVIStreamGetFrameClose Lib "avifil32.dll" (pg As Any) As Long
'Private Declare Function AVIStreamStart Lib "avifil32.dll" (pavi As Any) As Long
''Private Declare Function AVIStreamSampleToTime Lib "avifil32.dll" (pavi As Any, ByVal lSample As Long) As Long
'Private Declare Sub AVIStreamInfo Lib "avifil32.dll" Alias "AVIStreamInfoA" (pavi As Any, psi As TAVISTREAMINFO, ByVal lSize As Long)
'Private Declare Function AVIStreamRead Lib "avifil32.dll" ( _
      pavi As Any, _
      ByVal lStart As Long, _
      ByVal lSamples As Long, _
      lpBuffer As Any, _
      ByVal cbBuffer As Long, _
      plBytes As Long, _
      plSamples As Long _
   ) As Long
'Private Declare Function AVIStreamWrite Lib "avifil32.dll" ( _
      pavi As Any, _
      ByVal lStart As Long, _
      ByVal lSamples As Long, _
      lpBuffer As Any, _
      ByVal cbBuffer As Long, _
      ByVal dwFlags As Long, _
      plSampWritten As Long, _
      plBytesWritten As Long _
   ) As Long
'mzt Private Declare Function AVIStreamReadFormat Lib "avifil32.dll" ( _
      pavi As Any, _
      ByVal lPos As Long, _
      lpFormat As Any, _
      ByRef lpcbFormat As Long) As Long
Private Declare Function AVIStreamGetFrame Lib "avifil32.dll" (pg As Any, ByVal lPos As Long) As Long

Private Declare Function DrawDibOpen Lib "MSVFW32.dll" () As Long
Private Declare Function DrawDibClose Lib "MSVFW32.dll" (ByVal hDD As Long) As Long
Private Declare Function DrawDibDraw Lib "MSVFW32.dll" (ByVal hDD As Long, ByVal hdc As Long, _
      ByVal xDst As Long, ByVal yDst As Long, ByVal dxDst As Long, ByVal dyDst As Long, _
      lpbi As Any, lpBits As Any, _
      ByVal xSrc As Long, ByVal ySrc As Long, ByVal dxSrc As Long, ByVal dySrc As Long, _
      ByVal wFlags As Long) As Long
   
Private Declare Function GetPixelAPI Lib "gdi32" Alias "GetPixel" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" ( _
    ByVal lpDriverName As String, lpDeviceName As Any, _
   lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

'mzt Private Type TICINFO
'mzt     dwSize As Long ';                 // sizeof(ICINFO)
'mzt     fccType As Long ';                // compressor type     'vidc' 'audc'
'mzt     fccHandler As Long ';             // compressor sub-type 'rle ' 'jpeg' 'pcm '
'mzt     dwFlags As Long ';                // flags LOWORD is type specific
'mzt     dwVersion As Long ';              // version of the driver
'mzt     dwVersionICM As Long ';           // version of the ICM used
'mzt     '//
'mzt     '// under Win32, the driver always returns UNICODE strings.
'mzt     '//
'mzt     'WCHAR   szName[16];             // short name
'mzt     szName(0 To 31) As Byte
'mzt     'WCHAR   szDescription[128];     // long name
'mzt     szDescription(0 To 255) As Byte
'mzt     'WCHAR   szDriver[128];          // driver that contains compressor
'mzt     szDriver(0 To 255) As Byte
'mzt End Type

'mzt Private Declare Function ICInfo Lib "MSVFW32.dll" ( _
    ByVal fccType As Long, _
    ByVal fccHandler As Long, _
    lpicinfo As Any _
    ) As Long

Public Enum EBitmapCompressionTypes
   BI_RGB = 0&
   BI_RLE8 = 1&
End Enum
   
Private m_sFileName As String
Private m_pAS As Long
Private m_pGF As Long
Private m_hDrawDib As Long
Private m_lFrames As Long
Private m_lFrameDuration As Long
Private m_tBMIH As BITMAPINFOHEADER
Private m_tBMIHBlank As BITMAPINFOHEADER
'mzt Private m_tBMI256 As BITMAPINFO256
'Private m_cPal As cPalette
Private m_tAVI As TAVISTREAMINFO
'mzt Private m_tICI As TICINFO
Private pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)

'SUR add for find key frames
Private Declare Function AVIStreamFindSample Lib "avifil32.dll" (ByVal pavi As Long, _
                                                                ByVal lPos As Long, _
                                                                ByVal lFlags As Long) As Long
'{-- Flags for AVIStreamFindSample --------------------------------------------}
'mzt Private Const FIND_DIR As Long = &HF        'direction
Private Const FIND_NEXT As Long = &H1       ' go forward
Private Const FIND_PREV As Long = &H4       'go backward
'mzt Private Const FIND_FROM_START As Long = &H8 ' start at the logical beginning
'mzt Private Const FIND_TYPE As Long = &HF0      ' type mask
Private Const FIND_KEY As Long = &H10       'find key frame.
'mzt Private Const FIND_ANY As Long = &H20       'find any (non-empty) sample
'mzt Private Const FIND_FORMAT As Long = &H40    ' find format change
'mzt Private Const FIND_RET As Long = &HF000     ' return mask
'mzt Private Const FIND_POS As Long = &H0        ' return logical position
'mzt Private Const FIND_LENGTH As Long = &H1000  ' return logical size
'mzt Private Const FIND_OFFSET As Long = &H2000  ' return physical position
'mzt Private Const FIND_SIZE As Long = &H3000    ' return physical size
'mzt Private Const FIND_INDEX As Long = &H4000   ' return physical index position

'mzt Private Const DDF_UPDATE = &H2    '/* re-draw the last DIB */
'mzt Private Const DDF_SAME_HDC = &H4    '/* HDC same as last call (all setup) */
'mzt Private Const DDF_SAME_DRAW = &H8    '/* draw params are the same */
'mzt Private Const DDF_DONTDRAW = &H10   '/* dont draw frame, just decompress */
'mzt Private Const DDF_ANIMATE = &H20   '/* allow palette animation */
'mzt Private Const DDF_BUFFER = &H40   '/* always buffer image */
'mzt Private Const DDF_JUSTDRAWIT = &H80   '/* just draw it with GDI */
'mzt Private Const DDF_FULLSCREEN = &H100  '/* use DisplayDib */
'mzt Private Const DDF_BACKGROUNDPAL = &H200  '/* Realize palette in background */
'mzt Private Const DDF_NOTKEYFRAME = &H400  '/* this is a partial frame update, hint */
'mzt Private Const DDF_HURRYUP = &H800  '/* hurry up please! */
'mzt Private Const DDF_HALFTONE = &H1000 '/* always halftone */
'mzt Private Const DDF_PREROLL = DDF_DONTDRAW '/* Builing up a non-keyframe */
'mzt Private Const DDF_SAME_DIB = DDF_SAME_DRAW
'mzt Private Const DDF_SAME_SIZE = DDF_SAME_DRAW


Public Property Get filename() As String
   filename = m_sFileName
End Property

Public Property Let filename(ByVal Value As String)
   
   If Value = vbNullString Then Unload: Exit Property
   
   m_sFileName = Value
   Dim lErr As Long
'   If (FileExists(Value, lErr)) Then
      Load
'   Else
'      pErr 1
'   End If
End Property

'Private Function FileExists(ByVal file As String, ByRef lErr As Long) As Boolean
'Dim sTest As String
'   On Error Resume Next
'   sTest = Dir(file)
'   lErr = err.Number
'   FileExists = ((lErr = 0) And Len(sTest) > 0)
'   On Error GoTo 0
'End Function

Private Sub Unload()
Dim bS As Boolean
'    If m_pGF <> 0 Then
'        Call AVIStreamGetFrameClose(m_pGF) '//deallocates the GetFrame resources and interface
'    End If
 '   If pAVIStream <> 0 Then
 '       Call AVIStreamRelease(pAVIStream) '//closes video stream
 '   End If
'    If m_pAS <> 0 Then
'        Call AVIFileRelease(m_pAS) '// closes the file
'    End If

   ' close frame interface
   If (m_pGF) Then
      bS = (AVIStreamGetFrameClose(ByVal m_pGF) = 0)
      m_pGF = 0
   End If

   ' release avi stream
   If Not (pAVIStream = 0) Then
      AVIStreamRelease ByVal pAVIStream
      pAVIStream = 0
   End If
   
   ' release avi file
   If Not (m_pAS = 0) Then
      AVIFileRelease ByVal m_pAS
      m_pAS = 0
   End If
   
   m_lFrames = 0
   m_lFrameDuration = 0
   
   ' reset bitmap info details:
   LSet m_tBMIH = m_tBMIHBlank
   
   ' clear up the palette
   'Set m_cPal = Nothing
  
End Sub

Private Sub Load()
Dim hR As Long
'mzt Dim totalTime As Long
'mzt Dim hIC As Long
'Dim pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)
Dim numFrames As Long       'number of frames in video stream
Dim firstFrame As Long      'position of the first video frame
Dim fileInfo As AVI_FILE_INFO           'file info struct
Dim streamInfo As AVI_STREAM_INFO       'stream info struct
Dim bih As BITMAPINFOHEADER     'infoheader to pass to GetFrame functions
'mzt     Dim pGetFrameObj As Long    'pointer to GetFrame interface

aferror = False

'   hR = AVIStreamOpenFromFile(m_pAS, m_sFileName, STREAM_TYPE_VIDEO, _
    0, OF_READ Or OF_SHARE_EXCLUSIVE, ByVal 0&)
'hR = AVIStreamOpenFromFile(m_pAS, m_sFileName, STREAM_TYPE_VIDEO, _
 0, OF_READ, ByVal 0&)

'Open the AVI File and get a file interface pointer (PAVIFILE)
hR = AVIFileOpen(m_pAS, m_sFileName, OF_SHARE_DENY_WRITE, ByVal 0&)
If hR <> AVIERR_OK Then GoTo ErrorOut


'   If FAILED(hR) Then
'      m_pAS = 0
'      Unload
'      pErr 2
'   Else
' open frames

''''''''''''''''''''''''''''''
'Get the first available video stream (PAVISTREAM)
hR = AVIFileGetStream(m_pAS, pAVIStream, streamtypeVIDEO, 0)
If hR <> AVIERR_OK Then GoTo ErrorOut

'get the starting position of the stream (some streams may not start simultaneously)
firstFrame = AVIStreamStart(pAVIStream)
If firstFrame = -1 Then GoTo ErrorOut     'this function returns -1 on error

'get the length of video stream in frames
numFrames = AVIStreamLength(pAVIStream)
If numFrames = -1 Then GoTo ErrorOut     ' this function returns -1 on error

'    MsgBox "PAVISTREAM handle is " & pAVIStream & vbNewline & _
     '            "Video stream length - " & numFrames & vbNewline & _
     '            "Stream starts on frame #" & firstFrame & vbNewline & _
     '            "File and Stream info will be written to Immediate Window (from IDE - Ctrl+G to view)", vbInformation, App.title
'
'get file info struct (UDT)
hR = AVIFileInfo(m_pAS, fileInfo, Len(fileInfo))
If hR <> AVIERR_OK Then GoTo ErrorOut

'Debug.Print fileInfo.dwMaxBytesPerSecond =0

'    Call DebugPrintAVIFileInfo(fileInfo)

'get stream info struct (UDT)
hR = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
If hR <> AVIERR_OK Then GoTo ErrorOut

'   Call DebugPrintAVIStreamInfo(streamInfo)
'Debug.Print streamInfo.dwSampleSize=0
'Debug.Print streamInfo.fccHandler

'set bih attributes which we want GetFrame functions to return
With bih
    .biBitCount = 24
    .biClrImportant = 0
    .biClrUsed = 0
    .biCompression = BI_RGB
    .biHeight = streamInfo.rcFrame.Bottom - streamInfo.rcFrame.Top
    .biPlanes = 1
    .biSize = 40
    .biWidth = streamInfo.rcFrame.Right - streamInfo.rcFrame.Left
    .biXPelsPerMeter = 0
    .biYPelsPerMeter = 0
    .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight     'calculate total size of RGBQUAD scanlines (DWORD aligned)
End With

'init AVISTreamGetFrame* functions and create GETFRAME object
'pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, ByVal AVIGETFRAMEF_BESTDISPLAYFMT) 'tell AVIStream API what format we expect and input stream
'   m_pGF = AVIStreamGetFrame(pAVIStream, bih)  'force function to return 24bit DIBS

' m_pGF = AVIStreamGetFrameOpen(pAVIStream, ByVal 1) 'tell AVIStream API what format we expect and input stream
m_pGF = AVIStreamGetFrameOpen(pAVIStream, bih)    'force function to return 24bit DIBS

If m_pGF = 0 Then
    m_pGF = AVIStreamGetFrameOpen(pAVIStream, ByVal 0&)     'no force
End If

If m_pGF = 0 Then
    If Not AutoNoMessFlag Then MsgBox msgsvc(9) & " (" & frmEditor.TextVideoHid & ")", vbInformation

    'If MsgBox(msgsvc(9) & " (" & TextVideoHid & ")", vbOKCancel) = vbOK Then _
     'Shell App.Path & "\avic.exe", vbNormalFocus
    GoTo ErrorOut
End If
From_m_pGF = m_pGF    'out

'' get number of frames
'm_lFrames = numFrames
'
'Dim lSample As Long
'lSample = AVIStreamStart(ByVal pAVIStream) + AVIStreamLength(ByVal pAVIStream)
'totalTime = AVIStreamSampleToTime(ByVal pAVIStream, lSample)
'm_lFrameDuration = (totalTime / m_lFrames)
'
'            ' Get the AVI format:
'            Dim lSize As Long
'            hR = AVIStreamReadFormat(ByVal pAVIStream, 0, ByVal 0&, lSize)
'            If (FAILED(hR) Or lSize < LenB(m_tBMIH)) Then
'               pErr 5
'            Else
'    '           Debug.Print lSize
'               lSize = LenB(m_tBMIH)
'
'               hR = AVIStreamReadFormat(ByVal pAVIStream, 0, m_tBMIH, lSize)
'
'               ' Get driver info
'               m_tICI.dwSize = LenB(m_tICI)
'               ICInfo m_tAVI.fccType, m_tAVI.fccHandler, m_tICI
'
'            End If

'    Set bih = Nothing

Exit Sub
ErrorOut:
aferror = True
If (hR <> AVIERR_OK) Then
    ToDebug Err.Description
    If Not AutoNoMessFlag Then MsgBox msgsvc(10) & vbNewLine & m_sFileName & vbNewLine & Err.Description, vbInformation, App.title
End If
End Sub

Public Property Get name() As String
Dim sName As String
Dim iPos As Long
   sName = StrConv(m_tAVI.szName, vbUnicode, LCID)
   iPos = InStr(sName, vbNullChar)
   If (iPos > 0) Then
      sName = Left$(sName, iPos - 1)
   End If
   name = sName
End Property

'Public Property Get Palette() As cPalette
'   Set Palette = m_cPal
'End Property

Public Property Get Width() As Long
   Width = m_tBMIH.biWidth
End Property

Public Property Get Height() As Long
   Height = m_tBMIH.biHeight
End Property

Public Property Get BitsPerPixel() As Long
   BitsPerPixel = m_tBMIH.biBitCount
End Property

Public Property Get Compression() As EBitmapCompressionTypes
   Compression = m_tBMIH.biCompression
End Property

Public Property Get FrameCount() As Long
   FrameCount = m_lFrames
End Property

Public Property Get FrameDuration() As Long
   FrameDuration = m_lFrameDuration
End Property

Public Property Get VideoHandlerFourCC() As Long
   VideoHandlerFourCC = m_tAVI.fccHandler
End Property

Public Property Get VideoHandlerFourCCString() As String
   VideoHandlerFourCCString = FourCCToString(m_tAVI.fccHandler)
End Property

Public Property Get VideoTypeFourCC() As Long
   VideoTypeFourCC = m_tAVI.fccType
End Property

Public Property Get VideoTypeFourCCString() As String
   VideoTypeFourCCString = FourCCToString(m_tAVI.fccType)
End Property


' Research these
'Public Property Get VideoHandlerName() As String
'   VideoHandlerName = m_tICI.szName
'End Property

'Public Property Get VideoHandlerDescription() As String
'   VideoHandlerDescription = m_tICI.szDescription
'End Property

'Public Property Get VideoHandlerDriver() As String
'   VideoHandlerDriver = m_tICI.szDriver
'End Property
' End Research these

Public Property Get FrameBitmap( _
      ByVal Index As Long, _
      Optional ByVal lWidth As Long = -1, _
      Optional ByVal lHeight As Long = -1, _
      Optional ByVal NewBackColor As OLE_COLOR = -1 _
   ) As cBmp
Dim hDCDisplay As Long
Dim hDCComp As Long
Dim hBmp As Long
Dim hBmpOld As Long
Dim tR As RECT
Dim hBr As Long
   
   ' Prepare parameters:
   If (lWidth < 0) Then lWidth = Width
   If (lHeight < 0) Then lHeight = Height
   
   ' Create a DC and bitmap to draw to:
   hDCDisplay = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   hDCComp = CreateCompatibleDC(hDCDisplay)
   hBmp = CreateCompatibleBitmap(hDCDisplay, lWidth, lHeight)
   hBmpOld = SelectObject(hDCComp, hBmp)
   DeleteDC hDCDisplay
   
   ' If back color set, then fill:
   If Not (NewBackColor = -1) Then
      tR.Right = lWidth
      tR.Bottom = lHeight
      hBr = CreateSolidBrush(TranslateColor(NewBackColor))
      FillRect hDCComp, tR, hBr
      DeleteObject hBr
   End If
   
   ' Draw the frame into the DC:
   DrawFrame hDCComp, Index, , , lWidth, lHeight, (NewBackColor = -1)
   
   ' Select the bitmap out:
   SelectObject hDCComp, hBmpOld
   
   Dim cb As New cBmp
   cb.Init hBmp
   Set FrameBitmap = cb
   
   ' Clear up
   DeleteDC hDCComp
   
End Property
Public Property Get FramePicture( _
      ByVal Index As Long, _
      Optional ByVal lWidth As Long = -1, _
      Optional ByVal lHeight As Long = -1, _
      Optional ByVal NewBackColor As OLE_COLOR = -1 _
   ) As StdPicture
   
   Dim cb As cBmp
   Dim hBmp As Long
   Set cb = FrameBitmap(Index, lWidth, lHeight, NewBackColor)
   hBmp = cb.ExtracthBmp()
   Set FramePicture = BitmapToPicture(hBmp)
   
End Property

Public Sub DrawFrame( _
        ByVal lHDC As Long, _
        ByVal Index As Long, _
        Optional ByVal x As Long = 0, _
        Optional ByVal y As Long = 0, _
        Optional ByVal lWidth As Long = -1, _
        Optional ByVal lHeight As Long = -1, _
        Optional ByVal Transparent As Boolean = False _
        )
Dim lpbi As Long
Dim hDCComp As Long
Dim hBmp As Long
Dim hBmpOld As Long

' Prepare parameters:
'   If (lWidth < 0) Then lWidth = Width
'   If (lHeight < 0) Then lHeight = Height

'Index = Index - 1

' Get the uncompressed frame:
lpbi = AVIStreamGetFrame(ByVal m_pGF, Index)
If (lpbi) Then

    If (Transparent) Then
        ' Create a work dc
        hDCComp = CreateCompatibleDC(lHDC)
        hBmp = CreateCompatibleBitmap(lHDC, lWidth, lHeight)
        hBmpOld = SelectObject(hDCComp, hBmp)
        DrawDibDraw m_hDrawDib, hDCComp, 0, 0, _
                lWidth, lHeight, ByVal lpbi, ByVal 0&, _
                0, 0, -1, -1, 0
        DrawTransparent hDCComp, lHDC, x, y, lWidth, lHeight
        SelectObject hDCComp, hBmpOld
        DeleteObject hBmp
        DeleteDC hDCComp
    Else
        ' Draw it directly onto the display:
        DrawDibDraw m_hDrawDib, lHDC, x, y, _
                lWidth, lHeight, ByVal lpbi, ByVal 0&, _
                0, 0, -1, -1, 0

                
'mzt Private Const DDF_UPDATE = &H2    '/* re-draw the last DIB */
'mzt Private Const DDF_SAME_HDC = &H4    '/* HDC same as last call (all setup) */
'mzt Private Const DDF_SAME_DRAW = &H8    '/* draw params are the same */
'mzt Private Const DDF_DONTDRAW = &H10   '/* dont draw frame, just decompress */
'mzt Private Const DDF_ANIMATE = &H20   '/* allow palette animation */
'mzt Private Const DDF_BUFFER = &H40   '/* always buffer image */
'mzt Private Const DDF_JUSTDRAWIT = &H80   '/* just draw it with GDI */
'mzt Private Const DDF_FULLSCREEN = &H100  '/* use DisplayDib */
'mzt Private Const DDF_BACKGROUNDPAL = &H200  '/* Realize palette in background */
'mzt Private Const DDF_NOTKEYFRAME = &H400  '/* this is a partial frame update, hint */
'mzt Private Const DDF_HURRYUP = &H800  '/* hurry up please! */
'mzt Private Const DDF_HALFTONE = &H1000 '/* always halftone */
'mzt Private Const DDF_PREROLL = DDF_DONTDRAW '/* Builing up a non-keyframe */
'mzt Private Const DDF_SAME_DIB = DDF_SAME_DRAW
'mzt Private Const DDF_SAME_SIZE = DDF_SAME_DRAW

        '            DrawDibDraw m_hDrawDib, lhDC, x, y, _
                     lWidth, lHeight, ByVal lpbi, ByVal 0&, _
                     0, 0, -1, -1, 0

    End If
End If


End Sub

Private Sub DrawTransparent( _
      ByVal hdcSrc As Long, _
      ByVal hDCDest As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long _
   )
Dim hDCMask As Long
Dim hBmpMask As Long
Dim hBmpMaskOld As Long
Dim hDCMem As Long
Dim hBmpMem As Long
Dim hBmpMemOld As Long
   
   ' Prepare the mask DC & bitmap
   hDCMask = CreateCompatibleDC(0)
   hBmpMask = CreateCompatibleBitmap(hDCMask, lWidth, lHeight)
   hBmpMaskOld = SelectObject(hDCMask, hBmpMask)
   
   ' Prepare the output DC & bitmap
   hDCMem = CreateCompatibleDC(hDCDest)
   hBmpMem = CreateCompatibleBitmap(hDCDest, lWidth, lHeight)
   hBmpMemOld = SelectObject(hDCMem, hBmpMem)
   
   ' Copy the background into the output:
   BitBlt hDCMem, 0, 0, lWidth, lHeight, hDCDest, x, y, vbSrcCopy
   
   ' Set background colour of source to the top-left pixel of the AVI
   SetBkColor hdcSrc, GetPixelAPI(hdcSrc, 0, 0)
   ' Copy source onto the mask bitmap:
   BitBlt hDCMask, 0, 0, lWidth, lHeight, hdcSrc, 0, 0, vbSrcCopy

   ' Set fore/back colour of source DC to black/white
   SetBkColor hdcSrc, &H0&
   SetTextColor hdcSrc, &HFFFFFF
   ' AND the mask onto the source:
   BitBlt hdcSrc, 0, 0, lWidth, lHeight, hDCMask, 0, 0, vbSrcAnd

   ' Set fore/back color of the output DC to white/black:
   SetBkColor hDCMem, &HFFFFFF
   SetTextColor hDCMem, &H0&
   ' AND the mask onto the output
   BitBlt hDCMem, 0, 0, lWidth, lHeight, hDCMask, 0, 0, vbSrcAnd
   ' XOR source onto output
   BitBlt hDCMem, 0, 0, lWidth, lHeight, hdcSrc, 0, 0, vbSrcPaint

   ' Copy to destination:
   BitBlt hDCDest, x, y, lWidth, lHeight, hDCMem, 0, 0, vbSrcCopy

   
   ' Clear up the output DC & bitmap
   SelectObject hDCMem, hBmpMemOld
   DeleteObject hBmpMem
   DeleteDC hDCMem
   
   ' Clear up the mask DC & bitmap
   SelectObject hDCMask, hBmpMaskOld
   DeleteObject hBmpMask
   DeleteDC hDCMask

End Sub

'Private Function AVIStreamEndTime() As Long
'Dim lSample As Long
'   lSample = AVIStreamStart(ByVal m_pAS) + AVIStreamLength(ByVal m_pAS)
'   lSample = AVIStreamStart(ByVal pAVIStream) + AVIStreamLength(ByVal pAVIStream)
'   AVIStreamEndTime = AVIStreamSampleToTime(ByVal pAVIStream, lSample)
'End Function

'Private Sub pErr(ByVal lErr As Long)
'Dim sMsg As String
'   Select Case lErr
'   Case 1
'      sMsg = "File not found"
'   Case 2
'      sMsg = "Failed to open AVI file"
'   Case 3
'      sMsg = "Failed get Frames"
'   Case 4
'      sMsg = "AVI contains no frames"
'   Case 5
'      sMsg = "Could not read the format of the AVI"
'   Case 6
'      sMsg = "Unsupported AVI format"
'   Case Else
'      sMsg = "Unexpected error " & lErr
'   End Select
'   'Err.Raise ERR_BASE + lErr, App.EXEName & ".cAVIFrameExtract", sMsg
'   MsgBox sMsg, vbCritical
''   Unload
'End Sub

'mzt Private Function FAILED(ByVal hR As Long) As Boolean
'   FAILED = Not (SUCCEEDED(hR))
'End Function

'mzt Private Function SUCCEEDED(ByVal hR As Long) As Boolean
'mzt    SUCCEEDED = ((hR And &H80000000) = 0)
'mzt End Function

'mzt Private Function ApiError(ByVal e As Long) As String
'    Dim s As String, c As Long
'    s = String(256, 0)
'    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
'                      FORMAT_MESSAGE_IGNORE_INSERTS, _
'                      0, e, 0&, s, Len(s), ByVal 0)
'    If c Then ApiError = Left$(s, c)
'End Function

Private Function BitmapToPicture(ByVal hBmp As Long) As IPicture

If (hBmp = 0) Then Exit Function

Dim oNewPic As Picture, tPicConv As PictDesc, IGuid As Guid

' Fill PictDesc structure with necessary parts:
With tPicConv
    .cbSizeofStruct = Len(tPicConv)
    .picType = vbPicTypeBitmap
    .hImage = hBmp
End With

' Fill in IDispatch Interface ID
With IGuid
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
End With

' Create a picture object:
OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic

' Return it:
Set BitmapToPicture = oNewPic


End Function

'mzt Private Function FourCCFromString(ByVal sString As String) As Long
'Dim lRet As Long
'Dim sChar As String
'Dim lChar As Long
'   sChar = Mid$(sString, 1)
'   lRet = lRet Or Asc(sChar)
'   sChar = Mid$(sString, 2)
'   lRet = lRet Or Asc(sChar) * &H100&
'   sChar = Mid$(sString, 3)
'   lRet = lRet Or Asc(sChar) * &H10000
'   sChar = Mid$(sString, 4)
'   lChar = Asc(sChar)
'   lRet = lRet Or (lChar And &H7F&) * &H1000000
'   If (lChar And &H80&) = &H80& Then
'      lRet = lRet Or &H80000000
'   End If
'   FourCCFromString = lRet
'End Function

Private Function FourCCToString(ByVal lFourCC As Long) As String
Dim sRet As String
Dim lUByte As Long
   sRet = Chr(lFourCC And &HFF)
   sRet = sRet & Chr((lFourCC And &HFF00&) \ &H100&)
   sRet = sRet & Chr((lFourCC And &HFF0000) \ &H10000)
   lUByte = (lFourCC And &H7F000000) \ &H1000000
   If (lFourCC And &H80000000) = &H80000000 Then
      lUByte = lUByte Or &H80&
   End If
   sRet = sRet & Chr(lUByte)
   FourCCToString = sRet
End Function

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
' Convert Automation color to Windows color
If OleTranslateColor(oClr, hPal, TranslateColor) Then
    TranslateColor = -1 'CLR_INVALID
End If
End Function

Private Sub Class_Initialize()
   LSet m_tBMIHBlank = m_tBMIH
   AVIFileInit
   m_hDrawDib = DrawDibOpen()
End Sub

Private Sub Class_Terminate()
   Unload
   AVIFileExit
   DrawDibClose m_hDrawDib
End Sub

Public Function AVIStreamIsKeyFrame(l As Long) As Boolean
 '  AVIStreamIsKeyFrame = AVIStreamNearestKeyFrame(pAVIStream, l) = l
   AVIStreamIsKeyFrame = AVIStreamNearestKeyFrame(l) = l
End Function
Public Function AVIStreamNearestKeyFrame(l As Long) As Long
AVIStreamNearestKeyFrame = AVIStreamFindSample(pAVIStream, l, FIND_PREV Or FIND_KEY)
'Debug.Print AVIStreamNearestKeyFrame
End Function


Public Function AVIStreamNearestPrevKeyFrame(l As Long) As Long
'Debug.Print "< " & AVIStreamFindSample(pAVIStream, l - 1, FIND_DIR), FrmMain.Position
'If AVIStreamFindSample(pAVIStream, l - 1, FIND_PREV) < FrmMain.Position Then
'AVIStreamNearestPrevKeyFrame = AVIStreamFindSample(pAVIStream, l - 1, FIND_ANY Or FIND_PREV)
'AVIStreamNearestPrevKeyFrame l
'End If

'Do While AVIStreamFindSample(pAVIStream, l - 1, FIND_DIR) < FrmMain.Position
'AVIStreamNearestPrevKeyFrame = AVIStreamFindSample(pAVIStream, l - 1, FIND_KEY Or FIND_PREV)
'Loop
AVIStreamNearestPrevKeyFrame = AVIStreamFindSample(pAVIStream, l - 1, FIND_PREV Or FIND_KEY)
'AVIStreamNearestPrevKeyFrame = AVIStreamFindSample(pAVIStream, AVIStreamNearestPrevKeyFrame - 1, FIND_KEY Or FIND_PREV)
'AVIStreamNearestPrevKeyFrame = AVIStreamFindSample(pAVIStream, AVIStreamNearestPrevKeyFrame - 1, FIND_KEY Or FIND_PREV)


'AVIStreamNearestPrevKeyFrame = AVIStreamFindSample(pAVIStream, l - 1, FIND_KEY Or FIND_DIR)

'AVIStreamNearestPrevKeyFrame = AVIStreamFindSample(pAVIStream, AVIStreamNearestPrevKeyFrame - 1, FIND_KEY Or FIND_PREV)


'if AVIStreamNearestPrevKeyFrame = -1 Then AVIStreamNearestPrevKeyFrame = 1
If AVIStreamNearestPrevKeyFrame < 1 Then AVIStreamNearestPrevKeyFrame = 1
End Function


Public Function AVIStreamNearestNextKeyFrame(l As Long) As Long
'Debug.Print "> " & AVIStreamFindSample(pAVIStream, l + 1, FIND_DIR), FrmMain.Position
'If AVIStreamFindSample(pAVIStream, l + 1, FIND_DIR) < FrmMain.Position Then
'AVIStreamNearestNextKeyFrame l
'End If

AVIStreamNearestNextKeyFrame = AVIStreamFindSample(pAVIStream, l + 1, FIND_KEY Or FIND_NEXT)
'AVIStreamNearestNextKeyFrame = AVIStreamFindSample(pAVIStream, AVIStreamNearestNextKeyFrame + 1, FIND_KEY Or FIND_NEXT)


'AVIStreamNearestNextKeyFrame = AVIStreamFindSample(pAVIStream, l + 1, FIND_KEY Or FIND_DIR)
'AVIStreamNearestNextKeyFrame = AVIStreamFindSample(pAVIStream, AVIStreamNearestNextKeyFrame + 1, FIND_KEY Or FIND_NEXT)

If AVIStreamNearestNextKeyFrame < 1 Then AVIStreamNearestNextKeyFrame = 1
End Function



