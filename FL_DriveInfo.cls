VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptoDriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************
'*** Flamed Drive information
'***
'*** reads information about a drive
'****************************************

Public Enum FL_WriteCaps
    WC_CDR = &H1&
    WC_CDRW = &H2&
    WC_DVDR = &H4&
    WC_DVDRAM = &H8&
    WC_TESTMODE = &H10&
    WC_BURNPROOF = &H20&
    WC_TAO = &H40&
    WC_TAO_TEST = &H80&
    WC_SAO = &H100&
    WC_SAO_TEST = &H200&
    WC_RAW_16 = &H400&
    WC_RAW_16_TEST = &H800&
    WC_RAW_96 = &H1000&
    WC_RAW_96_TEST = &H2000&
    WC_DVDPRW = &H4000&
    WC_DVDPR = &H8000&
    WC_DVDRRW = &H10000
    WC_DVDPRDL = &H20000
    WC_MRW = &H40000
End Enum

Public Enum FL_ReadCaps
    RC_CDR = &H1&
    RC_CDRW = &H2&
    RC_DVDR = &H4&
    RC_DVDRAM = &H8&
    RC_DVDROM = &H10&
    RC_C2 = &H20&
    RC_CDDARAW = &H40&
    RC_BARCODE = &H80&
    RC_ISRC = &H100&
    RC_MODE2FORM1 = &H200&
    RC_MODE2FORM2 = &H400&
    RC_MULTISESSION = &H800&
    RC_SUBCHANNELS = &H1000&
    RC_SUBCHANNELS_CORRECTED = &H2000&
    RC_SUBCHANNELS_FROM_LEADIN = &H4000&
    RC_CDTEXT = &H8000&
    RC_DVDPRW = &H10000
    RC_DVDPR = &H20000
    RC_DVDPRDL = &H40000
    RC_MRW = &H80000
End Enum

Public Enum FL_LoadingMech
    LOAD_CADDY                       ' Caddy
    LOAD_TRAY                        ' Tray
    LOAD_POPUP                       ' Popup
    LOAD_CHANGER                     ' Changer
    LOAD_UNKNWN                      ' Unknown
End Enum

Public Enum FL_PhysicalInterfaces
    IF_SCSI                          ' SCSI
    IF_ATAPI                         ' ATAPI
    IF_IEEE                          ' IEEE
    IF_USB                           ' USB
    IF_UNKNWN                        ' unknown
End Enum

Private blnDiscPresent As Boolean
Private blnLocked As Boolean
Private blnClosed As Boolean

Private strVendor As String
Private strProduct As String
Private strRevision As String

Private intHostAdapter As Integer
Private intTarget As Integer

Private blnAnalogAudio As Boolean
Private blnJitterCorrection As Boolean
Private blnLockable As Boolean

Private lngBufferSize As Long

Private lngMaxReadSpeed As Long
Private lngMaxWriteSpeed As Long

Private lngCurReadSpeed As Long
Private lngCurWriteSpeed As Long

Private lngReadCaps As FL_ReadCaps
Private lngWriteCaps As FL_WriteCaps

Private udeLoadingMech As e_LoadingMechanism
Private udeIPhy As e_DrvInterfaces

Private lngIdleTimer As Long
Private lngStandbyTimer As Long
Private lngSpinDownTimer As Long

Public Property Get SpinDownTimerMS() As Long
SpinDownTimerMS = lngSpinDownTimer
End Property

Public Property Get StandbyTimer100MS() As Long
StandbyTimer100MS = lngStandbyTimer
End Property

Public Property Get IdleTimer100MS() As Long
IdleTimer100MS = lngIdleTimer
End Property

Public Property Get PhysicalInterface() As FL_PhysicalInterfaces
PhysicalInterface = udeIPhy
End Property

Public Property Get DriveClosed() As Boolean
DriveClosed = blnClosed
End Property

Public Property Get DriveLocked() As Boolean
DriveLocked = blnLocked
End Property

Public Property Get DiscPresent() As Boolean
DiscPresent = blnDiscPresent
End Property

Public Property Get Vendor() As String
Vendor = strVendor
End Property

Public Property Get Product() As String
Product = strProduct
End Property

Public Property Get Revision() As String
Revision = strRevision
End Property

Public Property Get AnalogAudioPlayback() As Boolean
AnalogAudioPlayback = blnAnalogAudio
End Property

Public Property Get JitterEffectCorrection() As Boolean
JitterEffectCorrection = blnJitterCorrection
End Property

Public Property Get Lockable() As Boolean
Lockable = blnLockable
End Property

Public Property Get BufferSizeKB() As Long
BufferSizeKB = lngBufferSize
End Property

Public Property Get ReadSpeedMax() As Long
ReadSpeedMax = lngMaxReadSpeed
End Property

Public Property Get WriteSpeedMax() As Long
WriteSpeedMax = lngMaxWriteSpeed
End Property

Public Property Get ReadSpeedCur() As Long
ReadSpeedCur = lngCurReadSpeed
End Property

Public Property Get WriteSpeedCur() As Long
WriteSpeedCur = lngCurWriteSpeed
End Property

Public Property Get ReadCapabilities() As FL_ReadCaps
ReadCapabilities = lngReadCaps
End Property

Public Property Get WriteCapabilities() As FL_WriteCaps
WriteCapabilities = lngWriteCaps
End Property

Public Property Get HostAdapter() As Integer
HostAdapter = intHostAdapter
End Property

Public Property Get Target() As Integer
Target = intTarget
End Property

Public Property Get LoadingMechanism() As FL_LoadingMech
LoadingMechanism = udeLoadingMech
End Property

Public Function GetReadSpeeds(ByVal DriveID As String) As Integer()

Dim speeds() As Integer
Dim intSpeeds As Integer

' simply add speeds in 4x steps
intSpeeds = CDRomGetSpeed(DriveID).MaxRSpeed \ 176

ReDim speeds(intSpeeds / 4 - 2) As Integer

' try to leave out the last read speed
' as it gets added in the end
For intSpeeds = 4 To (intSpeeds - 4) Step 4
    speeds(intSpeeds / 4 - 1) = intSpeeds * 177
Next

' add max read speed
ReDim Preserve speeds(UBound(speeds) + 1) As Integer
speeds(UBound(speeds)) = CDRomGetSpeed(DriveID).MaxRSpeed

GetReadSpeeds = speeds

End Function

Public Function GetWriteSpeeds(ByVal DriveID As String) As Integer()
Dim speeds() As Integer
Dim i As Integer

CDRomGetWriteSpeeds DriveID, speeds

For i = LBound(speeds) To UBound(speeds)
    ' last bit indicates CAV rotation
    ' remove it
    speeds(i) = speeds(i) And &H7FFF
Next

GetWriteSpeeds = speeds
End Function

Public Function GetInfo(ByVal DriveID As String) As Boolean

On Error GoTo ErrorHandler

'If Not CDRomTestUnitReady(DriveID) Then Exit Function

Dim manager As New cOptoManager
Dim char As String
Dim udtFtCD As t_Feat_CD_READ
Dim udtFtDVDPRW As t_Feat_DVD_P_RW
Dim udtFtDVDPR As t_Feat_DVD_P_R
Dim udtFtDVDRRW As t_Feat_DVD_R_RW
Dim udtFtDVDPRDL As t_Feat_DVD_P_R_DL
Dim udtFtMRW As t_Feat_MRW

char = manager.DrvID2DrvChr(DriveID)

intHostAdapter = Asc(Left$(cCD.HAIDFromDrvChr(char), 1))
intTarget = Asc(Right$(cCD.HAIDFromDrvChr(char), 1))

strVendor = cCD.CDROMGetName(DriveID, &H1)
strProduct = cCD.CDROMGetName(DriveID, &H2)
strRevision = cCD.CDROMGetName(DriveID, &H4)

blnDiscPresent = CDRomIsDiscPresent(DriveID) > 0
blnClosed = CDRomIsTrayOpen(DriveID) < 1
blnLocked = CDRomIsTrayLocked(DriveID)

lngIdleTimer = CDRomGetIdleTimer(DriveID)
lngStandbyTimer = CDRomGetStandbyTimer(DriveID)

Select Case CDRomGetSpinDown(DriveID)
Case SD_125MS: lngSpinDownTimer = 125&
Case SD_16MIN: lngSpinDownTimer = 16 * 60 * 1000&
Case SD_16SEC: lngSpinDownTimer = 16 * 1000&
Case SD_1MIN: lngSpinDownTimer = 1 * 60 * 1000&
Case SD_1SEC: lngSpinDownTimer = 1 * 1000&
Case SD_250MS: lngSpinDownTimer = 250&
Case SD_2MIN: lngSpinDownTimer = 2 * 60 * 1000&
Case SD_2SEC: lngSpinDownTimer = 2 * 1000&
Case SD_32MIN: lngSpinDownTimer = 32 * 60 * 1000&
Case SD_32SEC: lngSpinDownTimer = 32 * 1000&
Case SD_4MIN: lngSpinDownTimer = 4 * 60 * 1000&
Case SD_4SEC: lngSpinDownTimer = 4 * 1000&
Case SD_500MS: lngSpinDownTimer = 500&
Case SD_8MIN: lngSpinDownTimer = 8 * 60 * 1000&
Case SD_8SEC: lngSpinDownTimer = 8 * 1000&
Case SD_VS: lngSpinDownTimer = -1&
End Select

' Has CD-Text feature?
CDRomGetConfiguration DriveID, &H1E, 2, VarPtr(udtFtCD), Len(udtFtCD) - 1
' Can read/write DVD+RW?
CDRomGetConfiguration DriveID, &H2A, 2, VarPtr(udtFtDVDPRW), Len(udtFtDVDPRW) - 1
' Can read/write DVD+R?
CDRomGetConfiguration DriveID, &H2B, 2, VarPtr(udtFtDVDPR), Len(udtFtDVDPR) - 1
' Can write DVD-R/-RW?
CDRomGetConfiguration DriveID, &H2F, 2, VarPtr(udtFtDVDRRW), Len(udtFtDVDRRW) - 1
' Can read/write DVD+R DL?
CDRomGetConfiguration DriveID, &H3B, 2, VarPtr(udtFtDVDPRDL), Len(udtFtDVDPRDL) - 1
' Can read/write MRW?
CDRomGetConfiguration DriveID, &H28, 2, VarPtr(udtFtMRW), Len(udtFtMRW) - 1

' Drive information
With CDRomGetLWInfo(DriveID)

    blnAnalogAudio = .AnalogAudio
    lngBufferSize = .BufferSize
    blnJitterCorrection = .JitterCorrection
    udeLoadingMech = .LoadingMechanism
    udeIPhy = .Interface
    blnLockable = .LockMedia
    lngMaxReadSpeed = .speeds.MaxRSpeed
    lngMaxWriteSpeed = .speeds.MaxWSpeed
    lngCurReadSpeed = .speeds.CurrRSpeed
    lngCurWriteSpeed = .speeds.CurrWSpeed

    lngReadCaps = 0
    lngWriteCaps = 0

    With .ReadFeatures

        If .BC Then _
                lngReadCaps = lngReadCaps Or RC_BARCODE
        If .C2ErrorPointers Then _
                lngReadCaps = lngReadCaps Or RC_C2
        If .CDDARawRead Then _
                lngReadCaps = lngReadCaps Or RC_CDDARAW
        If .CDR Then _
                lngReadCaps = lngReadCaps Or RC_CDR
        If .CDRW Then _
                lngReadCaps = lngReadCaps Or RC_CDRW
        If .DVDR Then _
                lngReadCaps = lngReadCaps Or RC_DVDR
        If .DVDRAM Then _
                lngReadCaps = lngReadCaps Or RC_DVDRAM
        If .DVDROM Then _
                lngReadCaps = lngReadCaps Or RC_DVDROM
        If udtFtDVDPRW.code(1) = &H2A Then _
                lngReadCaps = lngReadCaps Or RC_DVDPRW
        If udtFtDVDPR.code(1) = &H2B Then _
                lngReadCaps = lngReadCaps Or RC_DVDPR
        If udtFtDVDPRDL.code(1) = &H3B Then _
                lngReadCaps = lngReadCaps Or RC_DVDPRDL
        If udtFtMRW.code(1) = &H28 Then _
                lngReadCaps = lngReadCaps Or RC_MRW
        If .ISRC Then _
                lngReadCaps = lngReadCaps Or RC_ISRC
        If .Mode2Form1 Then _
                lngReadCaps = lngReadCaps Or RC_MODE2FORM1
        If .Mode2Form2 Then _
                lngReadCaps = lngReadCaps Or RC_MODE2FORM2
        If .Multisession Then _
                lngReadCaps = lngReadCaps Or RC_MULTISESSION
        If .subchannels Then _
                lngReadCaps = lngReadCaps Or RC_SUBCHANNELS
        If .SubChannelsCorrected Then _
                lngReadCaps = lngReadCaps Or RC_SUBCHANNELS_CORRECTED
        If .SubChannelsFormLeadIn Then _
                lngReadCaps = lngReadCaps Or RC_SUBCHANNELS_FROM_LEADIN
        If IsBitSet(udtFtCD.CDText, 0) Then _
                lngReadCaps = lngReadCaps Or RC_CDTEXT

    End With

    With .WriteFeatures

        If .BURNProof Then _
                lngWriteCaps = lngWriteCaps Or WC_BURNPROOF
        If .CDR Then _
                lngWriteCaps = lngWriteCaps Or WC_CDR
        If .CDRW Then _
                lngWriteCaps = lngWriteCaps Or WC_CDRW
        If .DVDR Then _
                lngWriteCaps = lngWriteCaps Or WC_DVDR
        If .DVDRAM Then _
                lngWriteCaps = lngWriteCaps Or WC_DVDRAM
        If IsBitSet(udtFtDVDRRW.writeDVDRW, 1) Then _
                lngWriteCaps = lngWriteCaps Or WC_DVDRRW
        If IsBitSet(udtFtDVDPRW.write, 0) Then _
                lngWriteCaps = lngWriteCaps Or WC_DVDPRW
        If IsBitSet(udtFtDVDPR.write, 0) Then _
                lngWriteCaps = lngWriteCaps Or WC_DVDPR
        If IsBitSet(udtFtDVDPRDL.write, 0) Then _
                lngWriteCaps = lngWriteCaps Or WC_DVDPRDL
        If IsBitSet(udtFtMRW.write, 0) Then _
                lngWriteCaps = lngWriteCaps Or WC_MRW
        If .TestMode Then _
                lngWriteCaps = lngWriteCaps Or WC_TESTMODE
        If .WriteModes.TAO Then _
                lngWriteCaps = lngWriteCaps Or WC_TAO
        If .WriteModes.TAOTest Then _
                lngWriteCaps = lngWriteCaps Or WC_TAO_TEST
        If .WriteModes.SAO Then _
                lngWriteCaps = lngWriteCaps Or WC_SAO
        If .WriteModes.SAOTest Then _
                lngWriteCaps = lngWriteCaps Or WC_SAO_TEST
        If .WriteModes.Raw16 Then _
                lngWriteCaps = lngWriteCaps Or WC_RAW_16
        If .WriteModes.Raw16Test Then _
                lngWriteCaps = lngWriteCaps Or WC_RAW_16_TEST
        If .WriteModes.Raw96 Then _
                lngWriteCaps = lngWriteCaps Or WC_RAW_96
        If .WriteModes.Raw96Test Then _
                lngWriteCaps = lngWriteCaps Or WC_RAW_96_TEST

    End With

End With

GetInfo = True

ErrorHandler:

End Function
