Attribute VB_Name = "MOptoInfo"
Option Explicit
Public OptoManagerInited As Boolean 'флаг инициализации манагера

Public cManager        As New cOptoManager
Private cDVDInfo        As New cOptoDVDInfo
Private cInfo           As New cOptoCDInfo

Private Type t_InqDat
    PDT                         As Byte               ' drive type
    PDQ                         As Byte               ' removable drive
    VER                         As Byte               ' MMC Version (zero for ATAPI)
    RDF                         As Byte               ' interface depending field
    DLEN                        As Byte               ' additional len
    rsv1(1)                     As Byte               ' reserved
    Feat                        As Byte               ' ?
    VID(7)                      As Byte               ' vendor
    PID(15)                     As Byte               ' Product
    PVER(3)                     As Byte               ' revision (= Firmware Version)
    FWVER(20)                   As Byte               ' ?
End Type
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Type t_MMC
    PageCode                    As Byte               ' Page Code
    PageLen                     As Byte               ' Page len
    rsvd2(7)                    As Byte               ' reserved
    ReadSupported               As Byte               ' readable formats
    WriteSupported              As Byte               ' writable formats
    misc(3)                     As Byte               ' misc.
    MaxReadSpeed(1)             As Byte               ' max. read speed
    NumVolLevels(1)             As Byte               ' num. volume levels
    BufferSize(1)               As Byte               ' buffer size
    CurrReadSpeed(1)            As Byte               ' curr. read speed
    rsvd                        As Byte               ' reserved
    misc2                       As Byte               ' misc.
    MaxWriteSpeed(1)            As Byte               ' max. write speed
    CurrWriteSpeed(1)           As Byte               ' curr write speed
    RotationControl             As Byte
    CurrWriteSpeedMMC3(1)       As Byte
End Type
' Multimedia Capabilities page write speed descriptor
Public Type t_MMCP_WriteSpeed
    rsvd        As Byte     ' reserved
    rotation    As Byte     ' rotation control
    speed(1)    As Byte     ' speed in kb/s
End Type
Public Type t_Speed
    MaxRSpeed                   As Integer            ' max. read speed
    MaxWSpeed                   As Integer            ' max. write speed
    CurrRSpeed                  As Integer            ' curr. read speed
    CurrWSpeed                  As Integer            ' curr write speed
End Type
Public Type t_ReadFeatures                            ' can read:
    CDR                         As Boolean            '      CD-R
    CDRW                        As Boolean            '      CD-RW
    DVDR                        As Boolean            '      DVD-R
    DVDROM                      As Boolean            '      DVD-ROM
    DVDRAM                      As Boolean            '      DVD-RAM
    subchannels                 As Boolean            '      Sub-Channels
    SubChannelsCorrected        As Boolean            '      Sub-Channels corrected
    SubChannelsFormLeadIn       As Boolean            '      Sub-Channels from Lead-In
    C2ErrorPointers             As Boolean            '      C2 Error Pointers
    ISRC                        As Boolean            '      International Standard Recording Code
    UPC                         As Boolean            '      ?
    BC                          As Boolean            '      Bar Code
    Mode2Form1                  As Boolean            '      Mode 2 Form 1 sectors
    Mode2Form2                  As Boolean            '      Mode 2 Form 2
    Multisession                As Boolean            '      Multi-Session CDs
    CDDARawRead                 As Boolean            '      Audio sectors
End Type

Public Type t_WriteModes
    Raw96                       As Boolean            ' Raw + 96
    Raw16                       As Boolean            ' Raw + 16
    SAO                         As Boolean            ' Session At Once
    TAO                         As Boolean            ' Track At Once
    Raw96Test                   As Boolean            ' Raw + 96 + Test-Mode
    Raw16Test                   As Boolean            ' Raw + 16 + Test-Mode
    SAOTest                     As Boolean            ' Session At Once + Test-Mode
    TAOTest                     As Boolean            ' Track At Once + Test-Mode
End Type

Public Type t_WriteFeatures                           ' can write:
    CDR                         As Boolean            '      CD-R
    CDRW                        As Boolean            '      CD-RW
    DVDR                        As Boolean            '      DVD-R
    DVDRAM                      As Boolean            '      DVD-RAM
    TestMode                    As Boolean            '      Test-Mode
    BURNProof                   As Boolean            '      BURN-Proof
    WriteModes                  As t_WriteModes       ' supported write modes
End Type
Public Type t_DrvInfo
    ReadFeatures                As t_ReadFeatures     ' read features
    WriteFeatures               As t_WriteFeatures    ' write features
    speeds                      As t_Speed            ' speeds
    AnalogAudio                 As Boolean            ' analog audio playback?
    JitterCorrection            As Boolean            ' jitter effect correction?
    BufferSize                  As Long               ' buffer size
    LockMedia                   As Boolean            ' can lock media?
    LoadingMechanism            As e_LoadingMechanism ' loading mechanism
    Interface                   As e_DrvInterfaces    ' drive interface
End Type
'(from CDROM-TOOL [GPL])
Public Enum e_SpinDown
    SD_VS                                             ' vendor specific
    SD_125MS                                          ' 125 ms
    SD_250MS                                          ' 250     "
    SD_500MS                                          ' 500     "
    SD_1SEC                                           '   1 sec
    SD_2SEC                                           '   2     "
    SD_4SEC                                           '   4     "
    SD_8SEC                                           '   8     "
    SD_16SEC                                          '  16     "
    SD_32SEC                                          '  32     "
    SD_1MIN                                           '   1 min
    SD_2MIN                                           '   2     "
    SD_4MIN                                           '   4     "
    SD_8MIN                                           '   8     "
    SD_16MIN                                          '  16     "
    SD_32MIN                                          '  32     "
End Enum
Public Type t_Feat_Hdr
    DataLen(3) As Byte      ' data length
    rsvd(1) As Byte         ' reserved
    curr_profile(1) As Byte ' current profile
End Type
Public Type t_Feat_CD_READ
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version (= 1)
    additional_len As Byte  ' additional length (= 4)
    CDText As Byte          ' can read CD-Text and C2?
    rsvd(2) As Byte         ' reserved
End Type

Public Type t_Feat_DVD_P_RW
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version (= 0)
    additional_len As Byte  ' additional length (= 4)
    write As Byte           ' can write DVD+RW
    close_only As Byte      ' background format
    rsvd(1) As Byte         ' reserved
End Type

'Feature 2Fh: DVD-R/RW
Public Type t_Feat_DVD_R_RW
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version (= 1)
    additional_len As Byte  ' additional length (= 4)
    writeDVDRW As Byte      ' can write DVD-RW
    rsvd(2) As Byte         ' reserved
End Type

'Feature 2Bh: DVD+R
Public Type t_Feat_DVD_P_R
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version
    additional_len As Byte  ' additional length
    write As Byte           ' can write DVD+R
    rsvd(2) As Byte         ' reserved
End Type

' Feature 3Bh: DVD+R Double Layer
Public Type t_Feat_DVD_P_R_DL
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version
    additional_len As Byte  ' additional length
    write As Byte           ' can write DVD+R DL
    rsvd(2) As Byte         ' reserved
End Type

' Feature 28h: Mount Rainer
Public Type t_Feat_MRW
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version
    additional_len As Byte  ' additional length
    write As Byte           ' can write MRW
    rsvd(2) As Byte         ' reserved
End Type


Public Enum e_DrvInterfaces
    IF_SCSI                                           ' SCSI
    IF_ATAPI                                          ' ATAPI
    IF_IEEE                                           ' IEEE
    IF_USB                                            ' USB
    IF_UNKNWN                                         ' unknown
End Enum
Public Enum e_LoadingMechanism
    LOAD_CADDY                                        ' Caddy
    LOAD_TRAY                                         ' Tray
    LOAD_POPUP                                        ' Popup
    LOAD_CHANGER                                      ' Changer
    LOAD_UNKNWN                                       ' Unknown
End Enum
Public Type t_DVD_CPYINFO
    Length(1)                   As Byte
    rsvd(1)                     As Byte
    cpyprotectsystype           As Byte
    regioninfo                  As Byte
End Type
Public Type t_DVD_Phys
    Length(1)                   As Byte               ' data len
    rsvd(1)                     As Byte               ' reserved
    BookType                    As Byte               ' Booktype (DVD-ROM, DVD-RAM, DVD-R, DVD+R, ...)
    discsize                    As Byte               ' DVD size/Max. Rate
    LayerType                   As Byte               ' number of layers/Layer Type
    TrackDens                   As Byte               ' Track density
    zero                        As Byte               ' 0
    StartSector(2)              As Byte               ' physical start sector
    zero2                       As Byte               ' 0
    EndSector(2)                As Byte               ' physical end sector
    zero3                       As Byte               ' 0
    EndSectorLayer0(2)          As Byte               ' physical end sector in layer 0
    bca                         As Byte               ' bca
    mspec(2030)                 As Byte               ' depends on disc
End Type
Public Type t_ReadCap
    Blocks(3)                   As Byte               ' written sectors
    BlockLen(3)                 As Byte               ' sectorsize
End Type
Public Type t_ATIP
    Length(1)                   As Byte               ' data len
    rsvd1(1)                    As Byte               ' reserved
    ITWP                        As Byte               ' ?
    Uru                         As Byte               ' ?
    DiscType                    As Byte               ' CD Type (CD-R/CD-RW)
    rsvd2                       As Byte               ' reserved
    LeadIn_Min                  As Byte               ' Lead-In Start (minutes)
    LeadIn_Sec                  As Byte               ' Lead-In Start (seconds)
    LeadIn_Frm                  As Byte               ' Lead-In Start (frames)
    rsvd3                       As Byte               ' reserved
    LeadOut_Min                 As Byte               ' Lead-Out Start (minutes)
    LeadOut_Sec                 As Byte               ' Lead-Out Start (seconds)
    LeadOut_Frm                 As Byte               ' Lead-Out Start (frames)
    rsvd4                       As Byte               ' reserved
    Rest(12)                    As Byte               ' rest
End Type

Public Type t_RDI
    PageLen(1)                  As Byte               ' Page len
    states                      As Byte               ' misc. data (erasable, ...)
    FirstTrack                  As Byte               ' first track on the disc (experience: 1...)
    NumSessionsLSB              As Byte               ' number of sessions
    FirstTrackLastSessionLSB    As Byte               ' first track in last session
    LastTrackLastSessionLSB     As Byte               ' last track in last session
    misc                        As Byte               ' misc.
    DiscType                    As Byte               ' CD sub type (CD-ROM/CD-I/XA)
    NumSessionsMSB              As Byte               ' number of sessions
    FirstTrackLastSessionMSB    As Byte               ' first track in last session
    LastTrackLastSessionMSB     As Byte               ' last track in last session
    DiscIdentification(3)       As Byte               ' CD ID
    LastSessionLeadInStart(3)   As Byte               ' Lead-In start time (h:m:s:f)
    LastPossibleLeadOutStart(3) As Byte               ' last possible Lead-Out Start (h:m:s:f)
    DBC(6)                      As Byte               ' Disc Bar Code
End Type
Public cCD As New cOptoCDROM                             ' CD-ROM class
Public Declare Sub Sleep Lib "KERNEL32" ( _
    ByVal dwMS As Long _
)
Public Type t_MSFLBA
    M                           As Byte               ' minutes
    s                           As Byte               ' seconds
    F                           As Byte               ' frames
    LBA                         As Long               ' logical block address
End Type
Public Enum e_CDType
    ROMTYPE_CDROM                                     ' CD-ROM
    ROMTYPE_CDR                                       ' CD-R
    ROMTYPE_CDRW                                      ' CD-RW
    ROMTYPE_CDROM_R_RW                                ' CD-ROM, CD-R oder CD-RW
    ROMTYPE_DVD_ROM                                   ' DVD-ROM
    ROMTYPE_DVD_R                                     ' DVD-R
    ROMTYPE_DVD_RW                                    ' DVD-RW
    ROMTYPE_DVD_RAM                                   ' DVD-RAM
    ROMTYPE_DVD_P_R                                   ' DVD+R
    ROMTYPE_DVD_P_RW                                  ' DVD+RW
End Enum
Public Enum e_CD_SubType
    STYPE_CDROMDA                                     ' CD-ROM or CDDA
    STYPE_CDI                                         ' CD-I
    STYPE_XA                                          ' CD-XA
    STYPE_UNKNWN                                      ' unknown
End Enum
Public Enum e_Status
    STAT_EMPTY                                        ' empty
    STAT_INCOMPLETE                                   ' uncomplete
    STAT_COMPLETE                                     ' complete
    STAT_UNKNWN                                       ' unknown
End Enum
Public Type t_CDInfo
    Capacity                    As Long               ' capacity (only CD-R[W])
    LeadIn                      As t_MSFLBA           ' Lead-In Start
    LeadOut                     As t_MSFLBA           ' Lead-Out Start
    DiscStatus                  As e_Status           ' CD Status
    LastSessionStatus           As e_Status           ' last session's status
    CDType                      As e_CDType           ' CD Type
    CDSubType                   As e_CD_SubType       ' Sub Type
    Erasable                    As Boolean            ' erasable?
    Tracks                      As Byte               ' number of tracks
    Sessions                    As Byte               ' number of sessions
    Size                        As Double               ' size of the disc
    Vendor                      As String             ' CD-R(W) vendor
End Type

Public Function CDRomTestUnitReady(ByVal DrvID As String) As Boolean

    Dim cmd(5) As Byte
    Dim i      As Integer

    If cCD.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0) Then
'Debug.Print "Unit ready? = " & True
        CDRomTestUnitReady = True
        Exit Function
    End If

    ' no disc present
    If cCD.LastASC = &H3A Then
        '
        Exit Function

    ' unit is becoming ready,
    ' or not ready to ready change
    ' because medium may have changed,
    ' wait for it
    ElseIf (cCD.LastASC = &H4 And cCD.LastASCQ = &H1) _
    Or (cCD.LastSK = 6 And cCD.LastASC = 40) Then

        ' try 5 times (~5 seconds)
        For i = 1 To 5
            If cCD.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0) Then Exit For
            Sleep 1000
        Next i

        CDRomTestUnitReady = cCD.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

    End If

End Function

'collects some information about the inserted disc
Public Function CDRomGetCDInfo(ByVal strDrv As String) As t_CDInfo

Dim BufAtip As t_ATIP
Dim BufRDI As t_RDI

'mzt     Dim sBuf            As String
'mzt     Dim conf_hdr(512)   As Byte


'get some information
CDRomReadDiscInformation strDrv, VarPtr(BufRDI), Len(BufRDI) - 1
'ccd.
CDRomReadTOC strDrv, 4, True, 0, VarPtr(BufAtip), Len(BufAtip) - 1


With CDRomGetCDInfo

    'Lead-In start time
    .LeadIn.M = BufAtip.LeadIn_Min
    .LeadIn.s = BufAtip.LeadIn_Sec
    .LeadIn.F = BufAtip.LeadIn_Frm
    .LeadIn.LBA = cCD.MSF2LBA(.LeadIn.M, .LeadIn.s, .LeadIn.F)

    'Lead-Out start time
    .LeadOut.M = BufAtip.LeadOut_Min
    .LeadOut.s = BufAtip.LeadOut_Sec
    .LeadOut.F = BufAtip.LeadOut_Frm
    .LeadOut.LBA = cCD.MSF2LBA(.LeadOut.M, .LeadOut.s, .LeadOut.F)

    'CD Status
    If IsBitSet(BufRDI.states, 1) = False And _
       IsBitSet(BufRDI.states, 0) = False Then

        .DiscStatus = STAT_EMPTY

    ElseIf IsBitSet(BufRDI.states, 1) = False And _
           IsBitSet(BufRDI.states, 0) Then

        .DiscStatus = STAT_INCOMPLETE

    ElseIf IsBitSet(BufRDI.states, 1) And _
           IsBitSet(BufRDI.states, 0) = False Then

        .DiscStatus = STAT_COMPLETE

    Else

        .DiscStatus = STAT_UNKNWN

    End If

    'last session status
    If Not IsBitSet(BufRDI.states, 3) And _
       Not IsBitSet(BufRDI.states, 2) Then

        .LastSessionStatus = STAT_EMPTY

    ElseIf Not IsBitSet(BufRDI.states, 3) And _
           IsBitSet(BufRDI.states, 2) Then

        .LastSessionStatus = STAT_INCOMPLETE

    ElseIf IsBitSet(BufRDI.states, 3) And _
           IsBitSet(BufRDI.states, 2) Then

        .LastSessionStatus = STAT_COMPLETE

    Else

        .DiscStatus = STAT_UNKNWN

    End If

    'CD Type
    .CDType = CDRomGetCDType(strDrv)

    'CD Sub-Type
    If BufRDI.DiscType = &H0 Then
        .CDSubType = STYPE_CDROMDA
    ElseIf BufRDI.DiscType = &H10 Then
        .CDSubType = STYPE_CDI
    ElseIf BufRDI.DiscType = &H20 Then
        .CDSubType = STYPE_XA
    Else
        .CDSubType = STYPE_UNKNWN
    End If

    'Erasable?
    .Erasable = IsBitSet(BufRDI.states, 4)

    'CD-R(W) Vendor
    '.Vendor = CDRomGetCDRWVendor(strDrv)

    'Sessions
    '.Sessions = cCD.LShift(BufRDI.NumSessionsMSB, 8) Or _
     BufRDI.NumSessionsLSB

    'Tracks
    '.Tracks = cCD.LShift(BufRDI.LastTrackLastSessionMSB, 8) Or _
     BufRDI.LastTrackLastSessionLSB

    'capacity in bytes (Mode 1)
    .Capacity = .LeadOut.LBA * 2048&
    If .Capacity < 0 Then .Capacity = 0

    'used part of the disc in bytes
    .Size = CDRomGetUsedBytes(strDrv)
End With
End Function

'read disc information
Public Function CDRomReadDiscInformation(ByVal strDrv As String, _
                                         ByVal PtrBuffer As Long, _
                                         ByVal BufferLen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H51                   ' READ DISC INFORMATION Op-Code
    cmd(7) = BufferLen \ &HFF       ' allocation length
    cmd(8) = BufferLen Mod &HFF     ' allocation length

    CDRomReadDiscInformation = cCD.ExecCMD(strDrv, cmd, 10, False, _
                                          SRB_DIR_IN, PtrBuffer, BufferLen, 10)
'Debug.Print "CDRomReadDiscInformation = " & CDRomReadDiscInformation
End Function

'read TOC/PMA/ATIP/CD-TEXT
Public Function CDRomReadTOC(ByVal DrvID As String, ByVal TOC_Format As Integer, _
                             ByVal MSF As Boolean, ByVal Track_Session As Integer, _
                             ByVal PtrBuffer As Long, ByVal BufferLen As Long _
                            ) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H43                           ' READ TOC OpCode
    cmd(1) = IIf(MSF, &H2, 0)               ' MSF or LBA?
    cmd(2) = TOC_Format                     ' Format (TOC, PMA, ATIP, CD-Text)
    cmd(6) = Track_Session                  ' Track/Session
    cmd(7) = BufferLen \ &HFF               ' allocation length
    cmd(8) = BufferLen And &HFF             ' allocation length

    CDRomReadTOC = cCD.ExecCMD(DrvID, cmd, 10, False, _
                              SRB_DIR_IN, PtrBuffer, BufferLen)

'Debug.Print "CDRomReadTOC - " & CDRomReadTOC
End Function
'bit in a byte is set?
Public Function IsBitSet(ByVal InByte As Byte, ByVal Bit As Byte) As Boolean
    IsBitSet = ((InByte And (2 ^ Bit)) > 0)
End Function
'trys to get the type of the inserted CD/DVD
Public Function CDRomGetCDType(ByVal strDrv As String) As e_CDType

    Dim BufRDI          As t_RDI        ' buffer for ReadDiscInformation
    Dim BufAtip         As t_ATIP       ' ATIP for CD-R/CD-RW
    Dim conf_hdr(512)   As Byte         ' configuration header (active profile)

Dim GetCorrectTypeFlag As Boolean

    'read disc information
    CDRomReadDiscInformation strDrv, VarPtr(BufRDI), Len(BufRDI) - 1


    'first try to read the ATIP to exclude CD-R/W
    If Not CDRomReadTOC(strDrv, 4, True, 0, VarPtr(BufAtip), Len(BufAtip) - 1) Then
'If True Then
'???не читает ток , выдает cd-rom (0) хотя просто не получил данные для cdr

        'if the Lead-Out is 255:255.255 MSF, it should be a CD-ROM
        If cCD.MSF2LBA(BufRDI.LastPossibleLeadOutStart(1), _
                      BufRDI.LastPossibleLeadOutStart(2), _
                      BufRDI.LastPossibleLeadOutStart(3)) _
            = cCD.MSF2LBA(255, 255, 255) Then

            'normal CD-ROM
            CDRomGetCDType = ROMTYPE_CDROM
            GetCorrectTypeFlag = True
        Else

            'could be a CD-ROM/R/RW
            CDRomGetCDType = ROMTYPE_CDROM_R_RW
            GetCorrectTypeFlag = True
        End If

    Else

        'ATIP could be read, either CD-R oder CD-RW.
        'but we could get fooled, so check the ATIP data :)

        'valid Lead-In start time?
        If BufAtip.LeadIn_Min > 0 Or _
           BufAtip.LeadIn_Sec > 0 Or _
           BufAtip.LeadIn_Frm > 0 Then

            'valide Lead-Out start time?
            If BufAtip.LeadIn_Min < 255 Or _
               BufAtip.LeadIn_Sec < 255 Or _
               BufAtip.LeadIn_Frm < 255 Then

                If IsBitSet(BufAtip.DiscType, 6) Then
                    'CD-RW
                    CDRomGetCDType = ROMTYPE_CDRW
                    GetCorrectTypeFlag = True
                Else
                    'CD-R
                    CDRomGetCDType = ROMTYPE_CDR
                    GetCorrectTypeFlag = True
                End If

            End If
        End If

    End If

If Not GetCorrectTypeFlag Then 'не узнали
    'is DVD in drive?
    If IsDVD(strDrv) Then
        'seems to be a DVD, determine its type
        CDRomGetCDType = GetDVDBookType(strDrv)
        'didn't work?
        If CDRomGetCDType = -1 Then
            'read the configuration header, we want the active profile
            If CDRomGetConfiguration(strDrv, 0, 2, VarPtr(conf_hdr(0)), UBound(conf_hdr)) Then
                'CD type by the active drive profile
                Select Case (cCD.LShift(conf_hdr(6), 8) Or conf_hdr(7))
                    Case &H8: CDRomGetCDType = ROMTYPE_CDROM       ' CD-ROM
                    Case &H9: CDRomGetCDType = ROMTYPE_CDR         ' CD-R
                    Case &HA: CDRomGetCDType = ROMTYPE_CDRW        ' CD-RW
                    Case &H10: CDRomGetCDType = ROMTYPE_DVD_ROM    ' DVD-ROM
                    Case &H11: CDRomGetCDType = ROMTYPE_DVD_R      ' DVD-R
                    Case &H12: CDRomGetCDType = ROMTYPE_DVD_RAM    ' DVD-RAM
                    Case &H13: CDRomGetCDType = ROMTYPE_DVD_RW     ' DVD-RW
                    Case &H14: CDRomGetCDType = ROMTYPE_DVD_RW     ' DVD-RW
                    Case &H1A: CDRomGetCDType = ROMTYPE_DVD_P_RW   ' DVD+RW
                    Case &H1B: CDRomGetCDType = ROMTYPE_DVD_P_R    ' DVD+R
                    Case Else: CDRomGetCDType = ROMTYPE_CDROM_R_RW ' das ging also mal garnicht...
                End Select
    
            End If

        End If

    End If
End If
End Function

'search for CD-R(W) vendor
'Private Function CDRomGetCDRWVendor(ByVal strDrv As String) As String
'срезано
'end sub

'Warning: the written sectors will be multiplied with 2048.
'         Mode 2 or DA not supported.
Private Function CDRomGetUsedBytes(ByVal strDrv As String) As Double

Dim cap As t_ReadCap
Dim cmd(9) As Byte

cmd(0) = &H25                         ' READ CAPACITY Op-Code

If Not cCD.ExecCMD(strDrv, cmd, 10, False, SRB_DIR_IN, _
        VarPtr(cap), Len(cap), 10) Then

    'failed
    CDRomGetUsedBytes = -1

Else

    'return written sectors
    CDRomGetUsedBytes = CDbl(cCD.LShift(cap.Blocks(0), 24) Or _
            cCD.LShift(cap.Blocks(1), 16) Or _
            cCD.LShift(cap.Blocks(2), 8) Or _
            cap.Blocks(3)) _
            * 2048#

End If
End Function


'simple DVD detection
Public Function IsDVD(ByVal strDrv As String) As Boolean
    Dim dummy(512) As Byte

    'DVD?
    If CDRomReadDVDStructure(strDrv, 0, 0, 0, VarPtr(dummy(0)), UBound(dummy)) Then
        If dummy(0) > 0 Or dummy(1) > 0 Then IsDVD = True
    End If
End Function


'read DVD Book
Private Function GetDVDBookType(ByVal strDrv As String) As e_CDType
Dim physdata As t_DVD_Phys
Dim book As Byte

'get DVD Book
If CDRomReadDVDStructure(strDrv, 0, 0, 0, VarPtr(physdata), _
        Len(physdata) - 1) Then

    '
    With physdata

        If IsBitSet(.BookType, 4) Then book = 1
        If IsBitSet(.BookType, 5) Then book = book Or 2
        If IsBitSet(.BookType, 6) Then book = book Or 4
        If IsBitSet(.BookType, 7) Then book = book Or 8

        Select Case book
        Case 0: GetDVDBookType = ROMTYPE_DVD_ROM           ' DVD-ROM
        Case 1: GetDVDBookType = ROMTYPE_DVD_RAM           ' DVD-RAM
        Case 2: GetDVDBookType = ROMTYPE_DVD_R             ' DVD-R
        Case 3: GetDVDBookType = ROMTYPE_DVD_RW            ' DVD-RW
        Case 9: GetDVDBookType = ROMTYPE_DVD_P_RW          ' DVD+RW
        Case 10: GetDVDBookType = ROMTYPE_DVD_P_R          ' DVD+R
        Case Else: GetDVDBookType = -1                     ' unknown
        End Select

    End With

End If
End Function


'read drive features
Public Function CDRomGetConfiguration(ByVal strDrv As String, ByVal StartFeature As Byte, _
                                      ByVal RT As Byte, ByVal PtrBuffer As Long, _
                                      ByVal buflen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H46                   ' GET CONFIGURATION Op-Code
    cmd(1) = RT                     ' RT Byte
    cmd(2) = StartFeature \ &HFF    ' startfeature
    cmd(3) = StartFeature And &HFF  ' startfeature
    cmd(7) = buflen \ &HFF          ' allocation length
    cmd(8) = buflen And &HFF        ' allocation length

    CDRomGetConfiguration = cCD.ExecCMD(strDrv, cmd, 10, False, _
                                       SRB_DIR_IN, PtrBuffer, buflen + 1, 10)
End Function

'read DVD structure
Public Function CDRomReadDVDStructure(ByVal strDrv As String, ByVal LBA As Long, _
                                      ByVal LayerNr As Byte, ByVal Format As Byte, _
                                      ByVal PtrBuffer As Long, _
                                      ByVal BufferLen As Long) As Boolean

    Dim cmd(11) As Byte

    cmd(0) = &HAD                           ' READ DVD STRUCTURE Op-Code
    cmd(2) = cCD.RShift(LBA, 24) And &HFF    ' LBA LSB
    cmd(3) = cCD.RShift(LBA, 16) And &HFF
    cmd(4) = cCD.RShift(LBA, 8) And &HFF
    cmd(5) = LBA And &HFF                   ' LBA MSB
    cmd(6) = LayerNr                        ' Layer Number
    cmd(7) = Format                         ' Information Format
    cmd(8) = BufferLen \ &HFF               ' allocation length
    cmd(9) = BufferLen And &HFF             ' allocation length

    CDRomReadDVDStructure = cCD.ExecCMD(strDrv, cmd, 12, False, _
                                       SRB_DIR_IN, PtrBuffer, BufferLen, 10)
'Debug.Print "DVD? = " & CDRomReadDVDStructure
End Function

Public Function CDRomGetWriteSpeeds(ByVal DrvID As String, _
                                    speeds() As Integer) As Boolean

    Dim buf(512)        As Byte
    Dim mpage()         As Byte
    Dim udtDescriptor   As t_MMCP_WriteSpeed
    Dim intSize         As Integer
    Dim intDescriptors  As Integer
    Dim i               As Integer

    ' get MMCP
    If Not CDRomModeSense10(DrvID, &H2A, VarPtr(buf(0)), 512, True, True) Then
        Exit Function
    End If

    ' get size of the page
    intSize = cCD.LShift(buf(0), 8) Or buf(1)

    ' set new buffer to grab full page
    ReDim mpage(intSize + 1) As Byte

    ' get the whole MMCP
    If Not CDRomModeSense10(DrvID, &H2A, VarPtr(mpage(0)), intSize + 2, True, True) Then
        Exit Function
    End If

    If intSize > 38 Then

        ' get the number of write speed descriptors
        intDescriptors = (cCD.LShift(mpage(30 + 8), 8) Or _
                          mpage(31 + 8)) \ 4

    End If

    ' write speed descriptors supplied?
    If intDescriptors > 0 Then

        ReDim speeds(intDescriptors) As Integer

        ' save CLV descriptors
        For i = 1 To intDescriptors

            ' get descriptor
            CopyMemory udtDescriptor, mpage(28 + 8 + (i * 4)), 4

            ' save speed (in kbytes/s)
            speeds(i - 1) = cCD.LShift(udtDescriptor.speed(0), 8) Or udtDescriptor.speed(1)

            ' mark CAV descriptors
            If CBool(udtDescriptor.rotation And &H7) Then
                speeds(i - 1) = speeds(i - 1) Or &H8000
            End If

        Next

    Else

        ' No write speed descriptors
        ' supplied with MMCP.
        ' Simply add write speeds in 4x steps:

        intDescriptors = CDRomGetSpeed(DrvID).MaxWSpeed \ 176
        If intDescriptors > 0 Then
            ReDim speeds(intDescriptors / 4 - 1) As Integer

            For i = 4 To (intDescriptors - 4) Step 4
                speeds((i / 4) - 1) = i * 177
            Next

        Else
            ReDim speeds(0) As Integer
        End If

    End If

    If intDescriptors > 0 Then
        ' add max write speed
        speeds(UBound(speeds)) = CDRomGetSpeed(DrvID).MaxWSpeed
    End If

    ' finished
    CDRomGetWriteSpeeds = True

End Function


'set a new read an write speed
'&HFFFF& = max. speed
Public Function CDRomSetCDSpeed(ByVal strDrv As String, _
                                ByVal NewReadSpeed As Long, _
                                ByVal NewWriteSpeed As Long, _
                                ByVal CAV As Boolean) As Boolean

    Dim cmd(11) As Byte

    If NewReadSpeed > &HFFFF& Then NewReadSpeed = &HFFFF&
    If NewWriteSpeed > &HFFFF& Then NewWriteSpeed = &HFFFF&

    cmd(0) = &HBB                           ' SET CD SPEED Op-Code
    cmd(1) = Abs(CAV)                       ' CAV rotation?

    If NewReadSpeed < &HFFFF& Then
        cmd(2) = NewReadSpeed \ &HFF        ' NewReadSpeed MSB
        cmd(3) = NewReadSpeed Mod &HFF      ' NewReadSpeed LSB
    Else
        cmd(2) = &HFF                       ' max read speed MSB
        cmd(3) = &HFF                       ' max read speed LSB
    End If

    If NewWriteSpeed < &HFFFF& Then
        cmd(4) = NewWriteSpeed \ &HFF       ' NewWriteSpeed MSB
        cmd(5) = NewWriteSpeed Mod &HFF     ' NewWriteSpeed LSB
    Else
        cmd(4) = &HFF                       ' max write speed MSB
        cmd(5) = &HFF                       ' max write speed LSB
    End If

    CDRomSetCDSpeed = cCD.ExecCMD(strDrv, cmd, 12, False, SRB_DIR_OUT, 0, 0)
End Function

'load media
Public Function CDRomLoadTray(ByVal DrvID As String) As Boolean

    Dim cmd(6) As Byte

    cmd(0) = &H1B           ' LOUNLOAD OpCode
    cmd(4) = &H3            ' Load Flag

    CDRomLoadTray = cCD.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

End Function

'eject media
Public Function CDRomUnloadTray(ByVal DrvID As String) As Boolean

    Dim cmd(6) As Byte

    cmd(0) = &H1B           ' LOUNLOAD OpCode
    cmd(4) = &H2            ' Unload Flag

    CDRomUnloadTray = cCD.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

End Function
'lock media
Public Function CDRomLockMedia(ByVal DrvID As String) As Boolean

    Dim cmd(5) As Byte

    cmd(0) = &H1E           ' LOCK/UNLOCK OpCode
    cmd(4) = 1              ' Lock Flag

    CDRomLockMedia = cCD.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

End Function
'unlock media
Public Function CDRomUnlockMedia(ByVal DrvID As String) As Boolean

    Dim cmd(5) As Byte

    cmd(0) = &H1E           ' LOCK/UNLOCK OpCode
    cmd(4) = 0              ' remove flags

    CDRomUnlockMedia = cCD.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

End Function
'read read- and writespeeds
Public Function CDRomGetSpeed(ByVal strDrv As String) As t_Speed
'mzt     Dim cmd(9) As Byte
    Dim mmc As t_MMC

    CDRomModeSense10 strDrv, &H2A, VarPtr(mmc), Len(mmc) - 1, True, True

    With CDRomGetSpeed
        .MaxRSpeed = cCD.LShift(mmc.MaxReadSpeed(0), 8) Or mmc.MaxReadSpeed(1)
        .CurrRSpeed = cCD.LShift(mmc.CurrReadSpeed(0), 8) Or mmc.CurrReadSpeed(1)
        .MaxWSpeed = cCD.LShift(mmc.MaxWriteSpeed(0), 8) Or mmc.MaxWriteSpeed(1)

        ' MMC 3/4 write speed?
        .CurrWSpeed = cCD.LShift(mmc.CurrWriteSpeedMMC3(0), 8) Or mmc.CurrWriteSpeedMMC3(1)
        If .CurrWSpeed = 0 Then
            ' no, take the MMC 1/2 one
            .CurrWSpeed = cCD.LShift(mmc.CurrWriteSpeed(0), 8) Or mmc.CurrWriteSpeed(1)
        End If
    End With
End Function

'disc present?
Public Function CDRomIsDiscPresent(ByVal DrvID As String) As Boolean

    Dim media_event_req(8) As Byte

    'get Tray Status
    If Not CDRomGetEventStatusNotification(DrvID, &H10, _
                                           VarPtr(media_event_req(0)), _
                                           UBound(media_event_req)) Then

        CDRomIsDiscPresent = CDRomTestUnitReady(DrvID)
        Exit Function

    End If

    'valid data?
    If media_event_req(0) = 0 And media_event_req(1) = 0 Then
        CDRomIsDiscPresent = CDRomTestUnitReady(DrvID)
        Exit Function
    Else
        'disc present?
        CDRomIsDiscPresent = IsBitSet(media_event_req(5), 1)
    End If

End Function
'Tray open or closed?
Public Function CDRomIsTrayOpen(ByVal DrvID As String) As Long

Dim media_event_req(8) As Byte

If Not CDRomTestUnitReady(DrvID) Then

    If cCD.LastASC = &H3A And cCD.LastASCQ = &H1 Then
        CDRomIsTrayOpen = False
    ElseIf cCD.LastASC = &H3A And cCD.LastASCQ = &H2 Then
        CDRomIsTrayOpen = True

        ' drive doesn't report door status with TUR
    ElseIf cCD.LastASC = &H3A And cCD.LastASCQ = 0 Then

        'get Tray-Status
        If Not CDRomGetEventStatusNotification(DrvID, &H10, _
                                               VarPtr(media_event_req(0)), _
                                               UBound(media_event_req)) Then

            CDRomIsTrayOpen = -1: Exit Function

        End If

        'valid data?
        If media_event_req(0) = 0 And media_event_req(1) = 0 Then
            CDRomIsTrayOpen = -1: Exit Function
        Else
            'Tray open or closed?
            CDRomIsTrayOpen = Abs(IsBitSet(media_event_req(5), 0))
        End If

    End If

Else

    CDRomIsTrayOpen = False

End If

End Function
'Media locked?
Public Function CDRomIsTrayLocked(ByVal DrvID As String) As Boolean
    Dim mmc As t_MMC

    'read MM Capabilities Page
    If Not CDRomModeSense10(DrvID, &H2A, VarPtr(mmc), Len(mmc) - 1, True, True) Then _
        CDRomIsTrayLocked = -1: Exit Function

    'check "Lock State" Bit
    CDRomIsTrayLocked = Abs(IsBitSet(mmc.misc(2), 1))
End Function

Public Function CDRomGetIdleTimer(ByVal strDrvID As String) As Long

    Dim mpage(19) As Byte

    If CDRomModeSense10(strDrvID, &H1A, VarPtr(mpage(0)), UBound(mpage), True, True) Then

        CDRomGetIdleTimer = cCD.LShift(mpage(12), 24) Or _
                            cCD.LShift(mpage(13), 16) Or _
                            cCD.LShift(mpage(14), 8) Or _
                            mpage(15)

    End If

End Function

Public Function CDRomGetStandbyTimer(ByVal strDrvID As String) As Long

Dim mpage(19) As Byte

If CDRomModeSense10(strDrvID, &H1A, VarPtr(mpage(0)), UBound(mpage), True, True) Then

    CDRomGetStandbyTimer = cCD.LShift(mpage(16), 24) Or _
                           cCD.LShift(mpage(17), 16) Or _
                           cCD.LShift(mpage(18), 8) Or _
                           mpage(19)

End If

End Function

'read spin down speed
'from CDROM TOOL (GPL)
Public Function CDRomGetSpinDown(ByVal strDrvID As String) As Integer
Dim mpage(255) As Byte

CDRomModeSense10 strDrvID, &HD, VarPtr(mpage(0)), UBound(mpage), True, True
CDRomGetSpinDown = mpage(11)
End Function
'collection drive information
Public Function CDRomGetLWInfo(ByVal strDrvID As String) As t_DrvInfo

Dim mmc As t_MMC

'read Multimedia Capabilities Mode Page
CDRomModeSense10 strDrvID, &H2A, VarPtr(mmc), Len(mmc) - 1, True, True

'read read features
With CDRomGetLWInfo.ReadFeatures
    .CDR = IsBitSet(mmc.ReadSupported, 0)
    .CDRW = IsBitSet(mmc.ReadSupported, 1)
    .DVDROM = IsBitSet(mmc.ReadSupported, 3)
    .DVDR = IsBitSet(mmc.ReadSupported, 4)
    .DVDRAM = IsBitSet(mmc.ReadSupported, 5)
    .CDDARawRead = IsBitSet(mmc.misc(1), 0)
    .Mode2Form1 = IsBitSet(mmc.misc(0), 4)
    .Mode2Form2 = IsBitSet(mmc.misc(0), 5)
    .Multisession = IsBitSet(mmc.misc(0), 6)
    .ISRC = IsBitSet(mmc.misc(1), 5)
    .UPC = IsBitSet(mmc.misc(1), 6)
    .BC = IsBitSet(mmc.misc(1), 7)
    .subchannels = IsBitSet(mmc.misc(1), 2)
    .SubChannelsCorrected = IsBitSet(mmc.misc(1), 3)
    .SubChannelsFormLeadIn = IsBitSet(mmc.misc(3), 5)
    .C2ErrorPointers = IsBitSet(mmc.misc(1), 4)
End With

'read write features
With CDRomGetLWInfo.WriteFeatures
    .CDR = IsBitSet(mmc.WriteSupported, 0)
    .CDRW = IsBitSet(mmc.WriteSupported, 1)
    .TestMode = IsBitSet(mmc.WriteSupported, 2)
    .DVDR = IsBitSet(mmc.WriteSupported, 4)
    .DVDRAM = IsBitSet(mmc.WriteSupported, 5)
    .BURNProof = IsBitSet(mmc.misc(0), 7)

    If CDRomWriteParams(strDrvID, False, False, 150, 1, 0, 0, False) Then _
       .WriteModes.TAO = True

    If CDRomWriteParams(strDrvID, True, False, 150, 1, 0, 0, False) Then _
       .WriteModes.TAOTest = True

    If CDRomWriteParams(strDrvID, False, False, 150, 2, 0, 8, False) Then _
       .WriteModes.SAO = True

    If CDRomWriteParams(strDrvID, True, False, 150, 2, 0, 8, False) Then _
       .WriteModes.SAOTest = True

    If CDRomWriteParams(strDrvID, False, False, 150, 3, 0, 1, False) Then _
       .WriteModes.Raw16 = True

    If CDRomWriteParams(strDrvID, True, False, 150, 3, 0, 1, False) Then _
       .WriteModes.Raw16Test = True

    If CDRomWriteParams(strDrvID, False, False, 150, 3, 0, 3, False) Then _
       .WriteModes.Raw96 = True

    If CDRomWriteParams(strDrvID, True, False, 150, 3, 0, 3, False) Then _
       .WriteModes.Raw96Test = True
End With

'generic information
With CDRomGetLWInfo
    .speeds = CDRomGetSpeed(strDrvID)
    .Interface = CDRomGetInterface(strDrvID)
    .LockMedia = IsBitSet(mmc.misc(2), 0)
    .AnalogAudio = IsBitSet(mmc.misc(0), 0)
    .JitterCorrection = IsBitSet(mmc.misc(1), 1)
    .BufferSize = cCD.LShift(mmc.BufferSize(0), 8) Or mmc.BufferSize(1)

    If IsBitSet(mmc.misc(2), 5) = False And _
       IsBitSet(mmc.misc(2), 6) = False And _
       IsBitSet(mmc.misc(2), 7) = False Then

        .LoadingMechanism = LOAD_CADDY

    ElseIf IsBitSet(mmc.misc(2), 5) And _
           IsBitSet(mmc.misc(2), 6) = False And _
           IsBitSet(mmc.misc(2), 7) = False Then

        .LoadingMechanism = LOAD_TRAY

    ElseIf IsBitSet(mmc.misc(2), 5) = False And _
           IsBitSet(mmc.misc(2), 6) And _
           IsBitSet(mmc.misc(2), 7) = False Then

        .LoadingMechanism = LOAD_POPUP

    ElseIf IsBitSet(mmc.misc(2), 5) = False And _
           IsBitSet(mmc.misc(2), 6) = False And _
           IsBitSet(mmc.misc(2), 7) Then

        .LoadingMechanism = LOAD_CHANGER

    ElseIf IsBitSet(mmc.misc(2), 5) And _
           IsBitSet(mmc.misc(2), 6) = False And _
           IsBitSet(mmc.misc(2), 7) Then

        .LoadingMechanism = LOAD_CHANGER

    Else

        .LoadingMechanism = LOAD_UNKNWN

    End If
End With

End Function

'read a Mode Page
Public Function CDRomModeSense10(ByVal DrvID As String, ByVal MP As Byte, _
                                 ByVal PtrBuffer As Long, _
                                 ByVal BufferLen As Long, _
                                 Optional ByVal DBD As Boolean, _
                                 Optional ByVal CV As Boolean) As Boolean

Dim cmd(9) As Byte

cmd(0) = &H5A                           ' MODE SENSE 10 OpCode
cmd(1) = Abs(DBD) * &H8                 ' Disable Block Descriptors
cmd(2) = MP Or (Abs(Not CV) * &H80)     ' Mode Page (default values)
cmd(7) = cCD.RShift(BufferLen, 8)        ' allocation length
cmd(8) = BufferLen And &HFF             ' allocation length

CDRomModeSense10 = cCD.ExecCMD(DrvID, cmd, 10, False, _
                               SRB_DIR_IN, PtrBuffer, BufferLen + 1)

End Function
'get Event/Status
Public Function CDRomGetEventStatusNotification(ByVal DrvID As String, _
                                                ByVal Request As Byte, _
                                                ByVal PtrBuffer As Long, _
                                                ByVal BufferLen As Long) As Boolean

Dim cmd(9) As Byte

cmd(0) = &H4A                       ' GET EVENT/STATUS NOTIFICATION OpCode
cmd(1) = 1                          '
cmd(4) = Request                    ' Request Type
cmd(8) = BufferLen                  ' allocation length

CDRomGetEventStatusNotification = cCD.ExecCMD( _
                                  DrvID, cmd, 10, False, SRB_DIR_IN, _
                                  PtrBuffer, BufferLen _
                                             )

End Function

'send new write parameters page
Public Function CDRomWriteParams(ByVal strDrv As String, ByVal TestMode As Boolean, _
                                 ByVal BURNProof As Boolean, ByVal TrackPause As Long, _
                                 ByVal WriteType As Byte, ByVal TrackMode As Byte, _
                                 ByVal DataBlockType As Byte, ByVal Multisession As Boolean _
                                                              ) As Boolean

Dim bufData(60) As Byte
'mzt     Dim PS As Byte
'mzt     Dim i As Integer

'read the page
'CDRomModeSense10 strDrv, &H5, VarPtr(bufData(0)), UBound(bufData)

bufData(1) = 58                               ' length

bufData(8) = &H5                              ' Page Code
bufData(9) = &H32                             ' Page length

'WriteType, Test-Mode and Burn Proof
bufData(10) = WriteType Or _
              Abs(TestMode) * &H10 Or _
              Abs(BURNProof) * &H40

'Track-Mode, Multi-Session
bufData(11) = TrackMode Or Abs(Multisession) * &HC0

'Data-Mode (Mode 1: 2048 Bytes User-Data)
bufData(12) = DataBlockType

'Track Pause: 150 sectors (frames) = 2 seconds
bufData(22) = cCD.RShift(TrackPause, 8) And &HFF
bufData(23) = TrackPause And &HFF

'send new WPP
CDRomWriteParams = CDRomModeSelect10(strDrv, VarPtr(bufData(0)), 60)
End Function
'gets the interface of a drive
Public Function CDRomGetInterface(ByVal strDrv As String) As e_DrvInterfaces
Dim Buffer(15) As Byte, cmd(9) As Byte
Dim inquiry As t_InqDat

'try to read the drive's core feature
CDRomGetConfiguration strDrv, 1, 2, VarPtr(Buffer(0)), UBound(Buffer)

'determine the interface from it
Select Case cCD.LShift(Buffer(12), 24) Or cCD.LShift(Buffer(13), 16) Or _
       cCD.LShift(Buffer(14), 8) Or Buffer(15)

Case 1: CDRomGetInterface = IF_SCSI
Case 2, 7: CDRomGetInterface = IF_ATAPI
Case 3, 4, 6: CDRomGetInterface = IF_IEEE
Case 8: CDRomGetInterface = IF_USB
Case Else: CDRomGetInterface = IF_UNKNWN

End Select

'if it didn't work, try INQUIRY
If CDRomGetInterface = IF_UNKNWN Then

    cmd(0) = &H12                   ' Inquiry OpCode
    cmd(4) = Len(inquiry) - 1       ' allocation length

    If cCD.ExecCMD(strDrv, cmd, 10, False, SRB_DIR_IN, _
                   VarPtr(inquiry), Len(inquiry), 10) Then

        If inquiry.rsv1(0) = 0 Then CDRomGetInterface = IF_ATAPI

    End If

End If
End Function
'send a Mode Page
Public Function CDRomModeSelect10(ByVal DrvID As String, ByVal PtrBuffer As Long, _
                                  ByVal BufferLen As Long) As Boolean

Dim cmd(9) As Byte

cmd(0) = &H55                       ' MODE SELECT10 OpCode
cmd(1) = &H10                       ' PF = 1 (Page Format)
cmd(7) = BufferLen \ &HFF           ' allocation length
cmd(8) = BufferLen Mod &HFF         ' allocation length

CDRomModeSelect10 = cCD.ExecCMD(DrvID, cmd, 10, True, _
                                SRB_DIR_OUT, PtrBuffer, BufferLen)

End Function

Public Function GetOptoInfo(letter As String) As String
'letter v:
Dim strDrvID As String


If Not Opt_GetMediaType Then
    GetOptoInfo = vbNullString
    Exit Function
End If

'If Not cManager.Init() Then cManager.Goodbye: Exit Function
If Not OptoManagerInited Then
    If Not cManager.InitOpto() Then Exit Function  ' true - форсить аспи
    OptoManagerInited = True
End If

strDrvID = vbNullString
'GetOptoInfo = vbNullString
    
    If cManager.IsCDVDDrive(letter) Then 'если дисковод
        strDrvID = cManager.DrvChr2DrvID(letter)
        
        If IsDVD(strDrvID) Then
            'dvd info
            GetOptoInfo = ShowInfoDVD(strDrvID)
        Else
            'cd info
            GetOptoInfo = ShowInfoCD(strDrvID)
        End If
        
    End If
    
'cManager.Goodbye
End Function
Private Function ShowInfoDVD(strDrvID) As String

'    Dim intLayers   As Integer
Dim strBuf As String

If Not cDVDInfo.GetInfo(strDrvID, 0) Then
    If Not cDVDInfo.GetInfo(strDrvID, 1) Then
        'MsgBox "Could not get info for layer 0.", vbExclamation
        ToDebug "Err_infolayer" ' 0/1"
        Exit Function
    End If
End If

'If Not cDVDInfo.GetInfo(strDrvID, 1) Then
'    MsgBox "Could not get info for layer 1.", vbExclamation
''    Exit Sub
'End If

With cDVDInfo
    '        lstNfo.AddItem "Layers: " & intLayers
    'FL_DVD_BOOKTYPES.
    Select Case .BookType
    Case DVD_ROM: strBuf = "DVD-ROM"
    Case DVD_RAM: strBuf = "DVD-RAM"
    Case DVD_R: strBuf = "DVD-R"
    Case DVD_RW: strBuf = "DVD-RW"
    Case DVD_PLUS_R: strBuf = "DVD+R"
    Case DVD_PLUS_RW: strBuf = "DVD+RW"
    End Select
    
    ShowInfoDVD = strBuf
    ToDebug "BookType: " & strBuf
    
'Debug.Print .LayerType

    '        lstNfo.AddItem "Part version: " & .PartVersion

    '        Select Case .DiskSize
    '            Case DVD_120mm: strBuf = "120 mm"
    '            Case DVD_80mm: strBuf = "80 mm"
    '        End Select
    '        lstNfo.AddItem "Disk size: " & strBuf

    '        If CBool(.LayerType And DVD_DATA_EMBOSSED) Then
    '            strBuf = "data embossed "
    '        End If
    '        If CBool(.LayerType And DVD_DATA_RECORDED) Then
    '            strBuf = strBuf & "recordable "
    '        End If
    '        If CBool(.LayerType And DVD_DATA_REWRITABLE) Then
    '            strBuf = strBuf & "rewritable"
    '        End If
    '        lstNfo.AddItem "Layertype: " & strBuf

    '        Select Case .LinearDensity
    '            Case [0.267 um/bit]: strBuf = "0.267 um/bit"
    '            Case [0.280 to 0.291 um/bit]: strBuf = "0.280 to 0.291 um/bit"
    '            Case [0.293 um/bit]: strBuf = "0.293 um/bit"
    '            Case [0.353 um/bit]: strBuf = "0.353 um/bit"
    '            Case [0.409 to 0.435 um/bit]: strBuf = "0.409 to 0.435 um/bit"
    '        End Select
    '        lstNfo.AddItem "Linear density: " & strBuf

    '        Select Case .TrackDensity
    '            Case [0.615 um/track]: strBuf = "0.615 um/track"
    '            Case [0.74 um/track]: strBuf = "0.74 um/track"
    '            Case [0.80 um/track]: strBuf = "0.80 um/track"
    '        End Select
    '        lstNfo.AddItem "Track density: " & strBuf

    '        Select Case .MaximumRate
    '            Case [10.08 Mbps]: strBuf = "10.08 Mbps"
    '            Case [5.04 Mbps]: strBuf = "5.04 Mbps"
    '            Case [2.52 Mbps]: strBuf = "2.52 Mbps"
    '        End Select
    '        lstNfo.AddItem "Maximum Rate: " & strBuf

    '        Select Case .TrackPath
    '            Case DVD_PARALLEL_TRACK_PATH: strBuf = "parallel track path"
    '            Case DVD_OPPOSITE_TRACK_PATH: strBuf = "opposite track path"
    '        End Select
    '        lstNfo.AddItem "Track path: " & strBuf

    '        lstNfo.AddItem "Physical start sector data area: " & .PhysicalStartSectorDataArea
    '        lstNfo.AddItem "Physical end sector data area: " & .PhysicalEndSectorDataArea
    '        lstNfo.AddItem "Physical end sector layer 0: " & .PhysicalEndSectorLayer0

End With

'Set cDVDInfo = Nothing
End Function
Private Function ShowInfoCD(strDrvID) As String

If Not cInfo.GetInfo(strDrvID) Then
    '    MsgBox "Failed to read CD information.", vbExclamation, "Error"
    'попытка взять dvd инфо
    'ShowInfoCD = ShowInfoDVD(strDrvID)
    ToDebug "Err_CdInfoFail"
    Exit Function
End If

With cInfo

    '        lstNfo.AddItem "Capacity: " & (.Capacity \ 1024 ^ 2) & " MB"
    ShowInfoCD = CDTypeToStr(.MediaType)
    ToDebug "Тип CD: " & ShowInfoCD
    'Debug.Print strDrvID
    'Debug.Print ".MediaType: " & CDTypeToStr(.MediaType)
    'Debug.Print ".CDRWType: " & STypeToStr(.CDRWType)
    '       Debug.Print ".CDRWVendor: " & .CDRWVendor
    'Debug.Print ".Erasable: " & .Erasable
    '        lstNfo.AddItem "Last session's state: " & Status2Str(.LastSessionState)
    '        lstNfo.AddItem "Lead-In: " & .LeadInMSF.MSF & " MSF (" & .LeadInMSF.LBA & " LBA)"
    '        lstNfo.AddItem "Last possible Lead-Out start: " & .LeadOutMSF.MSF & " MSF (" & .LeadOutMSF.LBA & " LBA)"
    'Debug.Print ".MediaStatus: " & Status2Str(.MediaStatus)
    '        lstNfo.AddItem "Sessions: " & .Sessions
    '        lstNfo.AddItem "Tracks: " & .Tracks
    '        lstNfo.AddItem "Size: " & (.Size \ 1024 ^ 2) & " MB"
End With

'Set cInfo = Nothing
End Function
Private Function CDTypeToStr(s As e_CDType) As String
Select Case s
Case ROMTYPE_CDR: CDTypeToStr = "CD-R"
Case ROMTYPE_CDROM: CDTypeToStr = "CD-ROM"
Case ROMTYPE_CDROM_R_RW: CDTypeToStr = "CD-ROM/R/RW"
Case ROMTYPE_CDRW: CDTypeToStr = "CD-RW"
Case ROMTYPE_DVD_P_R: CDTypeToStr = "DVD+R"
Case ROMTYPE_DVD_P_RW: CDTypeToStr = "DVD+RW"
Case ROMTYPE_DVD_R: CDTypeToStr = "DVD-R"
Case ROMTYPE_DVD_RAM: CDTypeToStr = "DVD-RAM"
Case ROMTYPE_DVD_ROM: CDTypeToStr = "DVD-ROM"
Case ROMTYPE_DVD_RW: CDTypeToStr = "DVD-RW"
End Select

End Function
'Private Function STypeToStr(s As e_CD_SubType) As String
'    Select Case s
'        Case e_CD_SubType.STYPE_CDI: STypeToStr = "CD-I"
'        Case e_CD_SubType.STYPE_CDROMDA: STypeToStr = "CD-ROM/CDDA"
'        Case e_CD_SubType.STYPE_UNKNWN: STypeToStr = "Unknown"
'        Case e_CD_SubType.STYPE_XA: STypeToStr = "CD-XA"
'    End Select
'End Function

