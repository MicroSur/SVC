VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ��� m_cDib

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
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
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal Handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'mzt Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const BI_RGB = 0&
'mzt Private Const BI_RLE4 = 2&
'mzt Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'mzt Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'mzt Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Clipboard functions:
'Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function CloseClipboard Lib "user32" () As Long
'Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'Private Declare Function EmptyClipboard Lib "user32" () As Long
'mzt Private Const CF_BITMAP = 2
'mzt Private Const CF_DIB = 8

' Handle to the current DIBSection:
Private m_hDIb As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_tBI As BITMAPINFO



Public Function CreateDIB( _
        ByVal lhDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection( _
            lhDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)
    CreateDIB = (hDib <> 0)
End Function
Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
Dim lhDC As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBMP As BITMAP
Dim lHWnd As Long
    
    GetObjectAPI picThis.Handle, Len(tBMP), tBMP
    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lHWnd = GetDesktopWindow()
        lhDCDesktop = GetDC(lHWnd)
        If (lhDCDesktop <> 0) Then
            lhDC = CreateCompatibleDC(lhDCDesktop)
            ReleaseDC lHWnd, lhDCDesktop
            If (lhDC <> 0) Then
                lhBmpOld = SelectObject(lhDC, picThis.Handle)
                LoadPictureBlt lhDC
                SelectObject lhDC, lhBmpOld
                DeleteDC lhDC
            End If
        End If
    End If
End Function
Public Function Create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
    
'If lHeight / 8 <> Round(lHeight / 8, 0) Then lHeight = Round(lHeight / 8 + 1, 0) * 8
'If lWidth / 8 <> Round(lWidth / 8, 0) Then lWidth = Round(lWidth / 8 + 1, 0) * 8

        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            Create = True
        Else
            DeleteDC m_hDC
            m_hDC = 0
        End If
    End If
End Function
Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt( _
        ByVal lhDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
End Sub


Public Sub PaintPicture( _
        ByVal lhDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Property Get hdc() As Long
    hdc = m_hDC
End Property
Public Property Get hDib() As Long
    hDib = m_hDIb
End Property
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
Public Sub RandomiseBits( _
        Optional ByVal bGray As Boolean = False _
    )
Dim bDib() As Byte
Dim x As Long, y As Long
Dim lC As Long
Dim tSA As SAFEARRAY2D
Dim xEnd As Long
    
    ' Get the bits in the from DIB section:
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

    ' random:
    Randomize Timer
    
    xEnd = (Width - 1) * 3
    If (bGray) Then
        For y = 0 To m_tBI.bmiHeader.biHeight - 1
            For x = 0 To xEnd Step 3
                lC = Rnd * 255
                bDib(x, y) = lC
                bDib(x + 1, y) = lC
                bDib(x + 2, y) = lC
            Next x
        Next y
    Else
        For x = 0 To xEnd Step 3
            For y = 0 To m_tBI.bmiHeader.biHeight - 1
                bDib(x, y) = 0
                bDib(x + 1, y) = Rnd * 255
                bDib(x + 2, y) = Rnd * 255
            Next y
        Next x
    End If
    
    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDIb <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb
        End If
        DeleteDC m_hDC
    End If
    m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub

Public Function Resample( _
        ByVal lNewHeight As Long, _
        ByVal lNewWidth As Long _
    ) As cDIBSection
Dim cDib As cDIBSection
    Set cDib = New cDIBSection
    If cDib.Create(lNewWidth, lNewHeight) Then
        If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or (lNewHeight <> m_tBI.bmiHeader.biHeight) Then
            ' Change in size, do resample:
            ResampleDib cDib
        Else
            ' No size change so just return a copy:
            cDib.LoadPictureBlt m_hDC
        End If
        Set Resample = cDib
    End If
End Function

Private Function ResampleDib(ByRef cDibTo As cDIBSection) As Boolean
Dim bDibFrom() As Byte
Dim bDibTo() As Byte

Dim tSAFrom As SAFEARRAY2D
Dim tSATo As SAFEARRAY2D

    ' Get the bits in the from DIB section:
    With tSAFrom
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

    ' Get the bits in the to DIB section:
    With tSATo
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibTo.BytesPerScanLine()
        .pvData = cDibTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

Dim xScale As Single
Dim yScale As Single

Dim x As Long, y As Long, xEnd As Long, xOut As Long

Dim fX As Single, fY As Single
Dim ifY As Long, ifX As Long
Dim dX As Single, dy As Single
Dim R As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
Dim G As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
Dim B As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
Dim ir1 As Long, ig1 As Long, ib1 As Long
Dim ir2 As Long, ig2 As Long, ib2 As Long

    xScale = (Width - 1) / cDibTo.Width
    yScale = (Height - 1) / cDibTo.Height
    
    xEnd = cDibTo.Width - 1
        
    For y = 0 To cDibTo.Height - 1
        
        fY = y * yScale
        ifY = Int(fY)
        dy = fY - ifY
        
        For x = 0 To xEnd
            fX = x * xScale
            ifX = Int(fX)
            dX = fX - ifX
            
            ifX = ifX * 3
            ' Interpolate using the four nearest pixels in the source
            b1 = bDibFrom(ifX, ifY): g1 = bDibFrom(ifX + 1, ifY): r1 = bDibFrom(ifX + 2, ifY)
            b2 = bDibFrom(ifX + 3, ifY): g2 = bDibFrom(ifX + 4, ifY): r2 = bDibFrom(ifX + 5, ifY)
            b3 = bDibFrom(ifX, ifY + 1): g3 = bDibFrom(ifX + 1, ifY + 1): r3 = bDibFrom(ifX + 2, ifY + 1)
            b4 = bDibFrom(ifX + 3, ifY + 1): g4 = bDibFrom(ifX + 4, ifY + 1): r4 = bDibFrom(ifX + 5, ifY + 1)
            
            ' Interplate in x direction:
            ir1 = r1 * (1 - dy) + r3 * dy: ig1 = g1 * (1 - dy) + g3 * dy: ib1 = b1 * (1 - dy) + b3 * dy
            ir2 = r2 * (1 - dy) + r4 * dy: ig2 = g2 * (1 - dy) + g4 * dy: ib2 = b2 * (1 - dy) + b4 * dy
            ' Interpolate in y:
            R = ir1 * (1 - dX) + ir2 * dX: G = ig1 * (1 - dX) + ig2 * dX: B = ib1 * (1 - dX) + ib2 * dX
            
            ' Set output:
            If (R < 0) Then R = 0
            If (R > 255) Then R = 255
            If (G < 0) Then G = 0
            If (G > 255) Then G = 255
            If (B < 0) Then B = 0
            If (B > 255) Then
                B = 255
            End If
            xOut = x * 3
            bDibTo(xOut, y) = B
            bDibTo(xOut + 1, y) = G
            bDibTo(xOut + 2, y) = R
            
        Next x
        
    Next y

    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
    CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4


End Function

Private Sub Class_Terminate()
    ClearUp
End Sub

Public Function CreateFromPicture2( _
        ByRef picThis As StdPicture _
    )
Dim lhDC As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBMP As BITMAP
Dim lHWnd As Long
    
    GetObjectAPI picThis.Handle, Len(tBMP), tBMP
    If (Create2(tBMP.bmWidth, tBMP.bmHeight)) Then
        lHWnd = GetDesktopWindow()
        lhDCDesktop = GetDC(lHWnd)
        If (lhDCDesktop <> 0) Then
            lhDC = CreateCompatibleDC(lhDCDesktop)
            ReleaseDC lHWnd, lhDCDesktop
            If (lhDC <> 0) Then
                lhBmpOld = SelectObject(lhDC, picThis.Handle)
                LoadPictureBlt lhDC
                SelectObject lhDC, lhBmpOld
                DeleteDC lhDC
            End If
        End If
    End If
End Function
Public Function Create2( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
    
If lHeight / 8 <> Round(lHeight / 8, 0) Then lHeight = Round(lHeight / 8 + 0.5, 0) * 8
If lWidth / 8 <> Round(lWidth / 8, 0) Then lWidth = Round(lWidth / 8 + 0.5, 0) * 8

        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            Create2 = True
        Else
            DeleteDC m_hDC
            m_hDC = 0
        End If
    End If
End Function


