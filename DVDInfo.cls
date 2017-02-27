VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptoDVDInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************
'*** Flamed DVD Information
'*******************************************

' NOT TESTED!

Public Enum FL_DVD_TRACK_DENSITY
    [0.74 um/track]
    [0.80 um/track]
    [0.615 um/track]
End Enum

Public Enum FL_DVD_LINEAR_DENSITY
    [0.267 um/bit]
    [0.293 um/bit]
    [0.409 to 0.435 um/bit]
    [0.280 to 0.291 um/bit]
    [0.353 um/bit]
End Enum

Public Enum FL_DVD_LAYER_TYPE
    DVD_DATA_EMBOSSED
    DVD_DATA_RECORDED
    DVD_DATA_REWRITABLE
End Enum

Public Enum FL_DVD_TRACK_PATH
    DVD_PARALLEL_TRACK_PATH
    DVD_OPPOSITE_TRACK_PATH
End Enum

Public Enum FL_DVD_MAX_RATE
    [2.52 Mbps]
    [5.04 Mbps]
    [10.08 Mbps]
    DVD_RATE_NOT_SPECIFIED
End Enum

Public Enum FL_DVD_DISK_SIZES
    DVD_120mm
    DVD_80mm
End Enum

Public Enum FL_DVD_BOOKTYPES
    DVD_ROM
    DVD_RAM
    DVD_R
    DVD_RW
    DVD_PLUS_RW
    DVD_PLUS_R
End Enum

Private udeBook As FL_DVD_BOOKTYPES
Private udeSize As FL_DVD_DISK_SIZES
Private udeRate As FL_DVD_MAX_RATE
Private udePath As FL_DVD_TRACK_PATH
Private udeType As FL_DVD_LAYER_TYPE
Private udeLinDens As FL_DVD_LINEAR_DENSITY
Private udeTrkDens As FL_DVD_TRACK_DENSITY

Private intPart As Integer
Private intLayers As Integer

Private lngStartSectorDataArea As Long
Private lngEndSectorDataArea As Long
Private lngEndSectorLayer0 As Long

Private btRegionMask As Byte

Public Property Get BookType() As FL_DVD_BOOKTYPES
BookType = udeBook
End Property

Public Property Get PartVersion() As Integer
PartVersion = intPart
End Property

Public Property Get DiskSize() As FL_DVD_DISK_SIZES
DiskSize = udeSize
End Property

Public Property Get MaximumRate() As FL_DVD_MAX_RATE
MaximumRate = udeRate
End Property

Public Property Get NumLayers() As Integer
NumLayers = intLayers
End Property

Public Property Get TrackPath() As FL_DVD_TRACK_PATH
TrackPath = udePath
End Property

Public Property Get LayerType() As FL_DVD_LAYER_TYPE
LayerType = udeType
End Property

Public Property Get LinearDensity() As FL_DVD_LINEAR_DENSITY
LinearDensity = udeLinDens
End Property

Public Property Get TrackDensity() As FL_DVD_TRACK_DENSITY
TrackDensity = udeTrkDens
End Property

Public Property Get PhysicalStartSectorDataArea() As Long
PhysicalStartSectorDataArea = lngStartSectorDataArea
End Property

Public Property Get PhysicalEndSectorDataArea() As Long
PhysicalEndSectorDataArea = lngEndSectorDataArea
End Property

Public Property Get PhysicalEndSectorLayer0() As Long
PhysicalEndSectorLayer0 = lngEndSectorLayer0
End Property

Public Property Get RegionMask() As Byte
RegionMask = btRegionMask
End Property

Public Function GetInfo(ByVal DrvID As String, ByVal layer As Integer) As Boolean

Dim udtDVD As t_DVD_Phys
Dim udtDVDCpy As t_DVD_CPYINFO
Dim bt As Byte

GetInfo = CDRomReadDVDStructure(DrvID, 0, layer, 0, VarPtr(udtDVD), Len(udtDVD) - 1)
CDRomReadDVDStructure DrvID, 0, layer, 1, VarPtr(udtDVDCpy), Len(udtDVDCpy) - 1

With udtDVD

    ' BOOKTYPE
    bt = cCD.RShift(.BookType And &HF0, 4)


    Select Case bt
    Case 0: udeBook = DVD_ROM
    Case 1: udeBook = DVD_RAM
    Case 2: udeBook = DVD_R
    Case 3: udeBook = DVD_RW
        'sur?
    Case 9: udeBook = DVD_PLUS_RW
    Case 10: udeBook = DVD_PLUS_R
    End Select

    ' BOOK PART VERSION
    '        bt = .BookType And &HF
    '        intPart = bt

    ' DISK SIZE
    '        bt = cCD.RShift(.discsize And &HF0, 4)
    '        Select Case bt
    '            Case 0: udeSize = DVD_120mm
    '            Case 1: udeSize = DVD_80mm
    '        End Select

    ' MAXIMUM RATE
    '        bt = .discsize And &HF
    '        Select Case bt
    '            Case &H0: udeRate = [2.52 Mbps]
    '            Case &H1: udeRate = [5.04 Mbps]
    '            Case &H2: udeRate = [10.08 Mbps]
    '            Case &HF: udeRate = DVD_RATE_NOT_SPECIFIED
    '        End Select

    ' NUMBER OF LAYERS
    '        bt = cCD.RShift(.LayerType And &H60, 5)
    '        Select Case bt
    '            Case 0: intLayers = 1
    '            Case 1: intLayers = 2
    '        End Select

    ' TRACK PATH
    '        Select Case CBool(.LayerType And &H10)
    '            Case True: udePath = DVD_OPPOSITE_TRACK_PATH
    '            Case False: udePath = DVD_PARALLEL_TRACK_PATH
    '        End Select

    ' LAYER READ/WRITE ABILITY
    '        bt = .LayerType And &HF
    '        udeType = 0
    '        If IsBitSet(bt, 0) Then udeType = udeType Or DVD_DATA_EMBOSSED
    '        If IsBitSet(bt, 1) Then udeType = udeType Or DVD_DATA_RECORDED
    '        If IsBitSet(bt, 2) Then udeType = udeType Or DVD_DATA_REWRITABLE

    ' LINEAR DENSITY
    '        bt = cCD.RShift(.TrackDens And &HF0, 4)
    '        Select Case bt
    '            Case 0: udeLinDens = [0.267 um/bit]
    '            Case 1: udeLinDens = [0.293 um/bit]
    '            Case 2: udeLinDens = [0.409 to 0.435 um/bit]
    '            Case 4: udeLinDens = [0.280 to 0.291 um/bit]
    '            Case 8: udeLinDens = [0.353 um/bit]
    '        End Select

    ' TRACK DENSITY
    '        bt = .TrackDens And &HF
    '        Select Case bt
    '            Case 0: udeTrkDens = [0.74 um/track]
    '            Case 1: udeTrkDens = [0.80 um/track]
    '            Case 2: udeTrkDens = [0.615 um/track]
    '        End Select

    '        lngStartSectorDataArea = cCD.LShift(.StartSector(0), 16) Or _
             cCD.LShift(.StartSector(1), 8) Or _
             .StartSector(2)

    '        lngEndSectorDataArea = cCD.LShift(.EndSector(0), 16) Or _
             cCD.LShift(.EndSector(1), 8) Or _
             .EndSector(2)

    '        lngEndSectorLayer0 = cCD.LShift(.EndSectorLayer0(0), 16) Or _
             cCD.LShift(.EndSectorLayer0(1), 8) Or _
             .EndSectorLayer0(2)

End With

'    btRegionMask = udtDVDCpy.regioninfo

End Function
