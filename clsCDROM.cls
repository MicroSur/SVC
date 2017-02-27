VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptoCDROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TODO: ADD LUN SUPPORT!

'                                                   '
'+_________________________________________________+'
'                                                   '
'written @ 1024x768                                 '
'          Courier New 10pt                         '
'          4-Space Indents                          '
''__________________________________________________'
'|                                                  |
'|  clsCDROM - [rm_code] 2005                       |
'|                                                  |
'|  Low-Level CD/DVD-ROM access for VB              |
'|  for Windows 95/98/Me/NT/2000/XP                 |
'|__________________________________________________|
''!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!''
''     >>>>>>>>>>>>>>> WARNING <<<<<<<<<<<<<<<     ''
''                                                 ''
'' This code has low level access to your          ''
'' hardware. I'm not responsible in any way        ''
'' for any damage.                                 ''
''                                                 ''
''                                                 ''
'' But it can't happen much.                       ''
'' Maybe you burn some coasters or if              ''
'' if your ASPI isn't installed correctly          ''
'' your system could crash.                        ''
''                                                 ''
' __ Changelog: ___________________________________+.
'/                                                  |
'| 0.8.5                                            |
'|   - GetDriveHandle updated                       |
'|                                                  |
'| 0.8                                              |
'|   - pasted clsCDECL into clsCDROM                |
'|   - HA:ID functions                              |
'|   - CDRomGetName() added                         |
'|                                                  |
'| 0.7                                              |
'|   - added Drive ID model (faster)                |
'|   - clsCDECL added (by Paul Caton)               |
'|   - heavily improved ASPI and SPTI functions     |
'|   - removed ASPIShim.dll aka SCSI.dll            |
'|                                                  |
'| 0.6                                              |
'|   - dynamic CTL code calculation                 |
'|   - SPTI functions in a class                    |
'|                                                  |
'| 0.5                                              |
'|   - moved SPTI functions to ASPIShim.dll         |
'|   - renamed ASPIShim.dll to SCSI.dll             |
'|                                                  |
'| 0.4                                              |
'|   - ASPI (Win 9x/Me) Support                     |
'|   - ASPIShim.dll By Jon F. Zahornaky             |
'|                                                  |
'| 0.3                                              |
'|   - SP2 Bug fixed                                |
'|                                                  |
'| 0.2                                              |
'|   - improved SPTI functions                      |
'|                                                  |
'| 0.1                                              |
'|   - SPTI support                                 |
'|   - SPTICMD()                                    |
'|                                                  |
'\_________________________________________________+°
''                                                 ''
' __________________________________________________'
'|                                                  \
'| SPTI:                                            |
'| The SCSI Pass-Through Interface was added in     |
'| Windows NT and offers direct access to           |
'| SCSI/ATAPI/USB/Firewire drives.                  |
'| WARNING: Administrator priviledges needed.       |
'|          If you don't have them                  |
'|          the class trys the ASPI way.            |
'|          If there is no ASPI, nothing happens.   |
'\__________________________________________________/
''                                                 ''
' __________________________________________________'
'|                                                  \
'| ASPI:                                            |
'| The Advanced SCSI Programming Interface is       |
'| a driver developed by Adaptec,                   |
'| which works almost similar to the SPTI.          |
'| You don't need admin priviledges here.           |
'| If ASPI should fail, the ASPI is probably not    |
'| installed corretly.                              |
'| In this case you should try to install a new one.|
'| There are a lot of sources in the internet.      |
'\__________________________________________________/
''                                                 ''
' __________________________________________________'
'/                                                  \
'| Credits                                          |
'|                                                  |
'| Jon F. Zahornacky - ASPIShim.dll                 |
'| Alvise    - http://hochfeiler.it/alvise/         |
'| CD Freaks - http://club.cdfreaks.com/            |
'| T-10      - http://t10.org                       |
'| T-13      - http://t13.org                       |
'| CDR DAO   - http://cdrdao.sourceforge.net        |
'| AKRip     - http://akrip.sourceforge.net         |
'| CDRecord  - http://sites.inka.de/~W1752/cdrecord/|
'\__________________________________________________/
''                                                 ''
'                                                   .



Option Explicit                         ' explicit parsing

#Const doDebug = 0

'////////////////////////////////////\
'/////////////////////// functions   //=>>>
'/////////////////////////////////////

Private Declare Function CreateEvent Lib "KERNEL32" Alias "CreateEventA" ( _
        lpEventAttributes As SECURITY_ATTRIBUTES, _
        ByVal bManualReset As Long, _
        ByVal bInitialState As Long, _
        ByVal lpName As String _
        ) As Long

Private Declare Function WaitForSingleObject Lib "KERNEL32" ( _
        ByVal hHandle As Long, ByVal dwMilliseconds As Long _
        ) As Long

Private Declare Function ResetEvent Lib "KERNEL32" ( _
        ByVal hEvent As Long _
        ) As Long

Private Declare Function DeviceIoControl Lib "KERNEL32" ( _
        ByVal hDevice As Long, ByVal dwIoControlCode As Long, _
        ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, _
        ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
        ByRef lpBytesReturned As Long, ByRef lpOverlapped As t_OVERLAPPED _
        ) As Long

Private Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" ( _
        ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        hTemplateFile As Long _
        ) As Long

Private Declare Function CloseHandle Lib "KERNEL32" ( _
        ByVal hObject As Long _
        ) As Long

Private Declare Function GetDriveType Lib "KERNEL32" Alias "GetDriveTypeA" ( _
        ByVal nDrive As String _
        ) As Long

Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" ( _
        lpVersionInformation As OSVERSIONINFOEX _
        ) As Long

'--- by Paul Caton -----------------------------------------------|
Private Declare Sub RtlMoveMemory Lib "KERNEL32" ( _
        Destination As Any, Source As Any, ByVal Length As Long _
        )

'Private Declare Function FreeLibrary Lib "KERNEL32" ( _
 ByVal hLibModule As Long _
 ) As Long

Private Declare Function GetProcAddress Lib "KERNEL32" ( _
        ByVal hModule As Long, ByVal lpProcName As String _
        ) As Long

Private Declare Function LoadLibraryA Lib "KERNEL32" ( _
        ByVal lpLibFileName As String _
        ) As Long

'------------------------------------------------------------------|

'////////////////////////////////////\
'/////////////////////// structures  //=>>>
'/////////////////////////////////////

Private Type SRB
    SRB_Cmd As Byte      ' Command Code
    SRB_Status As Byte      ' Command Status Byte
    SRB_HaID As Byte      ' Host Adapter
    SRB_Flags As Byte      ' ASPI RequestFlags
    SRB_Hdr_Rsvd As Long      ' reserved
End Type

Private Type SRB_HAInquiry
    HA_Hdr As SRB       ' Std. SRB Header
    HA_Count As Byte      ' number of Host Adapters
    HA_Id As Byte      ' SCSI ID of the Host Adapter
    HA_MgrId As String * 16    ' Managername
    HA_Ident As String * 16    ' Hostadaptername
    HA_Unique(15) As Byte      ' param for the HA
    HA_Rsvd As Integer   ' reserved
    HA_Pad(19) As Byte      ' padding
End Type

Private Type SRB_GetDevType
    SRB_Hdr As SRB       ' Std. SRB Header
    SRB_Target As Byte      ' Target
    SRB_Lun As Byte      ' Target Logical Unit Number
    DEV_DeviceType As Byte      ' Devicetype
    DEV_Rsvd1 As Byte      ' reserved
    DEV_Pad(67) As Byte      ' padding
End Type

Private Type SRB_GetDiskInfo
    SRB_Hdr As SRB      ' Std. SRB Header
    SRB_Target As Byte     ' Target
    SRB_Lun As Byte     ' Target Logical Unit Number
    SRB_DriveFlags As Byte     ' drive flags
    SRB_Int13DrvInfo As Byte     ' + 65 = drive char (Ascii)
    SRB_Heads As Byte     ' heads
    SRB_Sectors As Byte     ' sectors
    SRB_Rsvd1(9) As Byte     ' reserved
End Type

Private Type SRB_ExecuteIO
    SRB_Hdr As SRB      ' Std. SRB Header
    SRB_Target As Byte     ' Target
    SRB_Lun As Byte     ' Target Logical Unit Number
    SRB_Rsvd1 As Integer  ' reserved
    SRB_BufLen As Long     ' data allocation length
    SRB_BufPointer As Long     ' Data Buffer Pointer
    SRB_SenseLen As Byte     ' sense buffer allocation length
    SRB_CDBLen As Byte     ' CDB Lдnge
    SRB_HaStat As Byte     ' Host Adapter Status
    SRB_TargStat As Byte     ' Target Status
    SRB_PostProc As Long     ' Post Routine
    SRB_Rsvd2(19) As Byte     ' reserved
    SRB_CDBByte(15) As Byte     ' SCSI CDB
    SRB_SenseData(15) As Byte     ' Request Sense buffer
End Type

Private Type SRB_Abort
    SRB_Hdr As SRB      ' Std. SRB Header
    SRB_SRBAbort As Long     ' Pointer to the SRB to cancel
End Type

Private Type t_SCSI_ADDRESS
    Length As Long     ' data len
    PortNumber As Byte     ' Host Adapter
    PathId As Byte     ' ?
    TargetId As Byte     ' Target
    Lun As Byte     ' Logical Unit Number
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long     ' data len
    lpSecurityDesc As Long     ' Security Descriptor
    hInherit As Long     ' Inherit Handle
End Type

Private Type t_ASPIInfo
    Status As Long     ' ASPI Status
    has As Long     ' number of Host Adapters
End Type

Private Type t_OVERLAPPED
    Internal As Long     '
    InternalHigh As Long     '
    offset As Long     '
    OffsetHigh As Long     '
    hEvent As Long     '
End Type

Private Type t_HAID
    HA As Byte     ' Host Adapter
    ID As Byte     ' Target
End Type

Private Type t_InqDat
    PDT As Byte     ' drive type
    PDQ As Byte     ' removable drive
    VER As Byte     ' MMC Version (zero for ATAPI)
    RDF As Byte     ' interface depending field
    DLEN As Byte     ' additional len
    rsv1(1) As Byte     ' reserved
    Feat As Byte     ' ?
    VID(7) As Byte     ' vendor
    PID(15) As Byte     ' Product
    PVER(3) As Byte     ' revision (= Firmware Version)
    FWVER(20) As Byte     ' ?
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long   ' data len
    dwMajorVersion As Long   ' Windows Major Version
    dwMinorVersion As Long   ' Windows Minor Version
    dwBuildNumber As Long   ' Windows Build Number
    dwPlatformId As Long   ' Windows Platform Identifier
    szCSDVersion As String * 128    ' Windows Version as a string
End Type

Private Type t_SPTD
    Length As Integer    ' data len
    ScsiStatus As Byte    ' SCSI Status
    PathId As Byte    ' Bus Number
    TargetId As Byte    ' Target
    Lun As Byte    ' Logical Unit Number
    CdbLength As Byte    ' CDB (Command Descriptor Block)
    SenseInfoLength As Byte    ' Sense Info len
    DataIn As Byte    ' data direction
    DataTransferLength As Long    ' data transfer length
    TimeOutValue As Long    ' command timeout
    DataBuffer As Long    ' Pointer to the data buffer
    SenseInfoOffset As Long    ' Sense Info Offset
    cdb(15) As Byte    ' Command Descriptor Block
    Fill(2) As Byte    '
End Type

Private Type t_SCSI_PASS_THROUGH_DIRECT_WITH_BUFFER
    spt As t_SPTD     ' SPTD structure
    'Fill(3)         As Byte       '
    SenseBuffer(35) As Byte       ' Debugging-Info from the drive
End Type

'/////////////////////////////////////\
'/////////////////////// enumerations //=>>>
'//////////////////////////////////////

Public Enum e_SPTIDirection
    SCSI_IOCTL_DATA_OUT = 0           ' send data to the drive
    SCSI_IOCTL_DATA_IN = 1            ' retrieve data from the drive
    SCSI_IOCTL_DATA_UNSPECIFIED = 2   ' no transfer
End Enum

Public Enum e_AspiDirection
    SRB_DIR_IN = &H8              ' retrieve data from the drive
    SRB_DIR_OUT = &H10            ' send data to the drive
End Enum

Public Enum e_Inquiry
    Vendor = &H1                  ' return vendor
    Product = &H2                 ' return product id
    Revision = &H4                ' return revision
End Enum

Public Enum e_Interfaces
    ASPI = 1
    SPTI = 2
End Enum

'/////////////////////////////////////\
'/////////////////////// constants    //=>>>
'//////////////////////////////////////

'mzt Private Const WAIT_FAILED           As Long = &HFFFFFFFF '
'mzt Private Const WAIT_OBJECT_0         As Long = &H0        '
'mzt Private Const WAIT_ABANDONED        As Long = &H80       '
Private Const WAIT_TIMEOUT As Long = &H102      '
'mzt Private Const STANDARD_RIGHTS_ALL   As Long = &H1F0000   '

Private Const IOCTL_SCSI_BASE As Long = &H4        '

Private Const METHOD_BUFFERED As Long = &H0        '
'mzt Private Const METHOD_IN_DIRECT      As Long = &H1        '
'mzt Private Const METHOD_OUT_DIRECT     As Long = &H2        '
'mzt Private Const METHOD_NEITHER        As Long = &H3        '

Private Const FILE_ANY_ACCESS As Long = &H0        ' full access
Private Const FILE_READ_ACCESS As Long = &H1        ' read access
Private Const FILE_WRITE_ACCESS As Long = &H2        ' write access

Private Const SC_HA_INQUIRY As Long = &H0        ' Host adapter inquiry
Private Const SC_GET_DEV_TYPE As Long = &H1        ' Get device type
Private Const SC_EXEC_SCSI_CMD As Long = &H2        ' Execute SCSI command
Private Const SC_GET_DISK_INFO As Long = &H6        ' Get Disk Info
Private Const SC_ABORT_SRB As Long = &H3        ' cancel SRB

'mzt Private Const SS_ERR                As Long = &H4        ' SRB failed
Private Const SS_COMP As Long = &H1        ' SRB finished
Private Const SS_PENDING As Long = &H0        ' SRB pending

'mzt Private Const SCSI_GOOD             As Long = &H0
'mzt Private Const SCSI_CHECK_CONDITION  As Long = &H2
'mzt Private Const SCSI_CONDITION_MET    As Long = &H4
'mzt Private Const SCSI_BUSY             As Long = &H8
'mzt Private Const SCSI_INTERMEDIATE     As Long = &H10
'mzt Private Const SCSI_INTER_COND_MET   As Long = &H14
'mzt Private Const SCSI_RESERV_CONFLICT  As Long = &H18
'mzt Private Const SCSI_CMD_TERMINATED   As Long = &H22
'mzt Private Const SCSI_QUEUE_FULL       As Long = &H28

Private Const SENSE_LEN As Long = &HE        ' Sensebuffer length

Private Const SRB_EVENT_NOTIFY As Long = &H40       ' ASPI Event

'mzt Private Const INVALID_HANDLE_VALUE  As Long = -1         ' invalid handle
Private Const OPEN_EXISTING As Long = &H3        ' open existing
Private Const GENERIC_READ As Long = &H80000000    ' generic read rights
Private Const GENERIC_WRITE As Long = &H40000000    ' generic write right
Private Const FILE_SHARE_READ As Long = &H1        ' shared read rights
Private Const FILE_SHARE_WRITE As Long = &H2        ' shared write right
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80       ' ?

Private Const VER_PLATFORM_WIN32_NT As Long = &H2        ' Windows NT 3/4/2000/XP/Longhorn

Private IOCTL_SCSI_PASS_THROUGH_DIRECT As Long           ' DevIoCtl Code
Private IOCTL_SCSI_PASS_THROUGH As Long           ' DevIoCtl Code
Private IOCTL_SCSI_GET_ADDRESS As Long           ' DevIoCtl Code

'ASPI API
Private Const ASPI_INFO As String = "GetASPI32SupportInfo"
Private Const ASPI_CMD As String = "SendASPI32Command"


'--- by Paul Caton ----------------------------------------
Private Const ERR_SRC As String = "clsCDECL"       ' Error source name
Private Const ERR_NUM As Long = vbObjectError      ' cCDECL error number base
Private Const MAX_ARG As Long = 16                 ' Maximum number of parameters, you can change this if required
Private Const PATCH_01 As Long = 15                 ' CDECL patch, CDECL function address
'mzt Private Const PATCH_02      As Long = 10                 ' Callback patch, bas mod function address patch
'mzt Private Const PATCH_03      As Long = 16                 ' Callback patch, stack adjustment patch

Private Const CODE_CDECL As String = "538B5C240C8B0BE305FF348BE2FBE8<fix 01>8B0BC1E10201CC5B8B54240C890231C0C20C00"
Private Const CODE_WRAPPER As String = "E8000000005A8F4219E8<fix 02>81EC<fix 03>E8000000005AFF7205C300000000"

'Parameter block
Private Type tParamBlock
    ParamCount As Long                    ' Number of parameters to be forwarded to the cdecl function
    Params(0 To MAX_ARG - 1) As Long                    ' Array of parameters to be forwarded to the cdecl function
End Type

'-----------------------------------------------------------


'///////////////////////////////////\
'/////////////////////// variables  //=>>>
'////////////////////////////////////

Private blnW2K As Boolean                         ' >= Windows 2000?

Private lngPower2(31) As Long                            ' bit shifting values
Private blnASPI As Boolean                         ' ASPI or SPTI?

Private lngHandles(25) As Long

Private intLastSK As Integer
Private intLastASC As Integer
Private intLastASCQ As Integer

Private Init As Boolean

'--- by Paul Cation -----------------------------
Private bNewDLL As Boolean                   ' Flag to indicate that the loaded DLL has changed
Private pMe As Long                      ' vtable address
Private hMod As Long                      ' DLL module handle
Private nAddr As Long                      ' Cache the previous cdecl function's address
Private nEntry As Long                      ' vtable entry index
Private pCode As Long                      ' Pointer to the CDECL code
Private sLastFunc As String                    ' Cache the previous cdecl function's name
Private Scode() As String                    ' Code buffer string array...
Private pb As tParamBlock               ' Parameter block instance

'Convert the passed string of hex character pairs to bytes stored in an ASCII
'string buffer. If indicated, patch the appropriate vtable entry to point to the byte codes
Private Sub Redirect(ByVal sHexCode As String, ByVal bPatch As Boolean)
Dim i As Long
Dim nLen As Long
Dim s As String

nLen = Len(sHexCode)

For i = 1 To nLen Step 2
    s = s & ChrB$(Val("&H" & Mid$(sHexCode, i, 2)))
Next i

ReDim Preserve Scode(0 To nEntry)
Scode(nEntry) = s

If bPatch Then
    'Patch the vtable entry to point to the code
    Call RtlMoveMemory(ByVal pMe + &H1C + (nEntry * 4), StrPtr(Scode(nEntry)), 4)
End If

nEntry = nEntry + 1                                 'In case another patch is added to the class
End Sub

'This sub is replaced by machine code in sCode(0) at class instance creation...
'IT MUST ONLY be called internally by CallFunc. It can't be made Private as it wouldn't
'then use the vtable. Being the first public method in this class, we know that the vtable
'pointer to this function will be located at [vtable + &H1C]
Public Function z_DO_NOT_CALL(ByVal nAddrParamBlock As Long) As Long
End Function

'by rm_code
Private Function FunctionExists(ByVal sFunction As String) As Boolean
FunctionExists = GetProcAddress(hMod, sFunction) <> 0
End Function

'Purpose:
' Call the named cdecl function with the passed parameters
'
'Arguments:
' sFunction - Name of the cdecl function to call
' ParmLongs - ParamArray of parameters to pass to the named cdecl function
'
'Return:
'  The return value of the named cdecl function
'
Private Function CallFunc(ByVal sFunction As String, ParamArray ParmLongs() As Variant) As Long
Dim i As Long
Dim j As Long

'Check that the DLL is loaded
If hMod = 0 Then

    'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
    Call err.Raise(ERR_NUM + 0, ERR_SRC, "DLL not loaded")
End If

'Check to see if we're calling the same cdecl function as the previous call to CallFunc
If (StrComp(sLastFunc, sFunction) <> 0) Or bNewDLL Then

    'Get the address of the function
    nAddr = GetProcAddress(hMod, sFunction)
    If nAddr = 0 Then

        'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
        Call err.Raise(ERR_NUM + 1, ERR_SRC, "Failed to locate function: " & sFunction)
    End If

    'Patch the code buffer to call the relative address to the cdecl function
    Call RtlMoveMemory(ByVal pCode + PATCH_01, nAddr - pCode - (PATCH_01 + 4), 4)
    bNewDLL = False
    sLastFunc = sFunction
End If

With pb
    j = UBound(ParmLongs)
    If j >= MAX_ARG Then

        'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
        Call err.Raise(ERR_NUM + 2, ERR_SRC, "Too many parameters")
    End If

    'Fill the parameter block
    For i = 0 To j
        .Params(i) = ParmLongs(i)
    Next i

    .ParamCount = i                                         '(j + 1)
End With

CallFunc = z_DO_NOT_CALL(VarPtr(pb))                        'Execute the code buffer passing the address of the parameter block
End Function

'Load the DLL
Private Function DllLoad(ByVal sName As String) As Boolean
hMod = LoadLibraryA(sName)

If hMod <> 0 Then
    DllLoad = True
    'It's remotely possible that the programmer could change the dll and then call a function
    'in the new dll with exactly the same name as the previous CallFunc to the previous DLL. This would
    'defeat the caching scheme and result in the old function in the old dll being called. An unlikely
    'scenario, but stranger things have happened. Soooo, explicitly indicate that we're using a new dll
    bNewDLL = True
End If
End Function

'It's not important to do this, but, if you've finished with a DLL there's no harm in releasing
'its memory. Don't bother at app end... it will be dealt with automatically when the process ends.
'Private Function DllUnload() As Boolean
'  If hMod <> 0 Then
'    DllUnload = (FreeLibrary(hMod) <> 0)
'    hMod = 0
'  End If
'End Function

'-------------------------------------------------

'read CD/DVD-ROM drive name with inquiry
Public Function CDROMGetName(ByVal strDrvID As String, _
        ByVal Flags As e_Inquiry) As String

Dim inq As t_InqDat         ' Inquiry structure
Dim cmd(5) As Byte             ' CDB

cmd(0) = &H12                   ' 0x12 Inquiry (6 Byte)
cmd(4) = Len(inq) - 1           ' allocation length

'execute the command (inquiry needs to be supported by every CD/DVD-ROM)
ExecCMD strDrvID, cmd, 6, True, SRB_DIR_IN, VarPtr(inq), Len(inq)

'return vendor
If Flags And Vendor Then _
        CDROMGetName = StrConv(inq.VID, vbUnicode, LCID)

'return product ID
If Flags And Product Then _
        CDROMGetName = CDROMGetName & StrConv(inq.PID, vbUnicode, LCID)

'return revision
If Flags And Revision Then _
        CDROMGetName = CDROMGetName & StrConv(inq.PVER, vbUnicode, LCID)

End Function

'drive is CD/DVD-ROM?
Public Function IsCDROM(ByVal strDrv As String) As Boolean
IsCDROM = GetDriveType(Left$(strDrv, 1) & ":") = 5
End Function

' last sense key
Public Property Get LastSK() As Integer
LastSK = intLastSK
End Property

' last additional sense code
Public Property Get LastASC() As Integer
LastASC = intLastASC
End Property

' last additional sense code qualifier
Public Property Get LastASCQ() As Integer
LastASCQ = intLastASCQ
End Property

'choose an interface
Public Function ChooseInterface(Optional ByVal ForceASPI As Boolean) As Boolean

'ASPI available?
If ASPIAvailable Then
    'GetASPI32SupportInfo 'закоментарено в оригинале
    ChooseInterface = True
    blnASPI = True
End If

'if only ASPI allowed, cancel
If ForceASPI Then Exit Function

'SPTI available?
If SPTIAvailable Then
    ChooseInterface = True
    blnASPI = False
End If

ToDebug "Use ASPI (else SPTI) = " & blnASPI

'load the ASPI driver
'If blnASPI Then DllLoad "wnaspi32" 'закоментарено в оригинале

End Function

'return the currently used interface
Public Function CurrentInterface() As e_Interfaces

CurrentInterface = Choose(Abs(blnASPI) + 1, e_Interfaces.SPTI, e_Interfaces.ASPI)

End Function

'test supported interfaces
Public Function InterfaceAvailable(ByVal i As e_Interfaces) As Boolean

Select Case i
Case ASPI
    InterfaceAvailable = ASPIAvailable
Case SPTI
    InterfaceAvailable = SPTIAvailable
End Select

End Function

'ASPI installed?
Private Function ASPIAvailable() As Boolean

ASPIAvailable = DllLoad("wnaspi32") And _
        FunctionExists(ASPI_INFO) And _
        FunctionExists(ASPI_CMD)

End Function

'SPTI available?
Private Function SPTIAvailable() As Boolean

Dim pswb As t_SCSI_PASS_THROUGH_DIRECT_WITH_BUFFER, ol As t_OVERLAPPED
Dim Length As Long, returned As Long, i As Long
Dim Status As Long, fh As Long
Dim drv As String

'find the first CD/DVD-ROM drive
For i = 1 To 26
    If GetDriveType(Chr$(i + 64) & ":") = 5 Then
        drv = Chr$(i + 64)
        Exit For
    End If
Next

'get its drive handle
If Not GetDriveHandle(drv, fh) Then Exit Function

With pswb.spt
    '.SenseInfoOffset = Len(pswb.spt) + 4               ' Sense Info Offset
    .SenseInfoLength = UBound(pswb.SenseBuffer) - 4         ' Sense Info size
    .SenseInfoOffset = VarPtr(pswb.SenseBuffer(0)) - VarPtr(pswb)
    .DataIn = SCSI_IOCTL_DATA_IN                           ' data direction
    .TimeOutValue = 10                                     ' Timeout
    .CdbLength = 6                                         ' CDB length
    .Length = 44                                           ' size of the substructure
End With
Length = Len(pswb)                                         ' size of the structure

'SPTD (SCSI Pass-Through Direct, need admin priviledges)
Status = DeviceIoControl(fh, IOCTL_SCSI_PASS_THROUGH_DIRECT, _
        pswb, Length, pswb, Length, _
        returned, ol)

'CloseHandle fh

'success?
SPTIAvailable = Status = 1

End Function

'get a drive's handle
Private Function GetDriveHandle(ByVal drv As String, ByRef fh As Long) As Boolean

Dim Flags As Long, i As Integer

' function already called?
If Not Init Then

    ' no, get the handles for every CD/DVD-ROM

    ' starting from win 2k you also need GENERIC_WRITE
    Flags = GENERIC_READ
    If blnW2K Then Flags = Flags Or GENERIC_WRITE

    For i = 1 To 26

        If GetDriveType(Chr$(i + 64) & ":") = 5 Then

            ' get the handle with CreateFile().
            ' you can access drives with "\\.\X:", where X is the drive's char.
            fh = CreateFile("\\.\" & Chr$(i + 64) & ":", Flags, _
                    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, _
                    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)

            lngHandles(i - 1) = fh

        Else

            ' not a CD/DVD-ROM drive
            lngHandles(i - 1) = -1

        End If

    Next

    ' initialized
    Init = True

End If

' return drive handle
fh = lngHandles(Asc(UCase$(drv)) - 65)

GetDriveHandle = fh <> -1

End Function

'execute a command through the ASPI
Private Function ASPI32Cmd(ByVal DrvID As String, ByRef cmd() As Byte, _
        ByVal CmdLen As Byte, ByVal eDir As e_AspiDirection, _
        ByVal BufPtr As Long, ByVal buflen As Long, _
        Optional ByVal WFE As Boolean = False, _
        Optional ByVal TimeoutSecs As Integer = 5 _
        ) As Boolean

Dim i As Long, hEvent As Long, hTmr As Long, ret As Long
Dim sd As SECURITY_ATTRIBUTES
Dim SRB As SRB_ExecuteIO

With sd
    .nLength = Len(sd)
    .lpSecurityDesc = 0
    .hInherit = 0
End With

'create a new event
hEvent = CreateEvent(sd, True, False, 0)
If hEvent = 0 Then Exit Function

With SRB

    .SRB_Hdr.SRB_Cmd = SC_EXEC_SCSI_CMD                 ' ASPI Command
    .SRB_Hdr.SRB_Flags = eDir Or SRB_EVENT_NOTIFY       ' data dir

    .SRB_Hdr.SRB_HaID = Asc(Left$(DrvID, 1))            ' Host Adapter
    .SRB_Target = Asc(Right$(DrvID, 1))                 ' Target

    .SRB_BufPointer = BufPtr                            ' buffer pointer
    .SRB_BufLen = buflen                                ' buffer len

    .SRB_SenseLen = SENSE_LEN                           ' sense buffer len
    .SRB_PostProc = hEvent                              ' process event

    .SRB_CDBLen = CmdLen                                ' CDB len

End With

For i = 0 To CmdLen - 1
    SRB.SRB_CDBByte(i) = cmd(i)                         ' copy the CDB
Next

'call ASPI32SendCommand()
ret = CallFunc(ASPI_CMD, VarPtr(SRB))

'if the command didn't complete yet...
If ret = SS_PENDING Then

    ' wait for ever...
    If WFE Then

        Do While SRB.SRB_Hdr.SRB_Status = SS_PENDING
            DoEvents
        Loop

        ' or a specific time...
    Else

        'wait TimeoutSecs seconds for it to complete
        hTmr = WaitForSingleObject(hEvent, TimeoutSecs * 1000&)

        'if timed out, reset the event
        If hTmr = WAIT_TIMEOUT Then
            ResetEvent hEvent
        End If

        'if still not complete...
        If SRB.SRB_Hdr.SRB_Status = SS_PENDING Then

            '... cancel the command
            ASPI32Abort SRB, SRB.SRB_Hdr.SRB_HaID
            CloseHandle hEvent

            Exit Function

        End If

    End If

End If

CloseHandle hEvent

'success?
ASPI32Cmd = SRB.SRB_Hdr.SRB_Status = SS_COMP

intLastSK = SRB.SRB_SenseData(2) And &HF
intLastASC = SRB.SRB_SenseData(12)
intLastASCQ = SRB.SRB_SenseData(13)

End Function

'from CDEx
'cancel an ASPI command
Private Function ASPI32Abort(SRB As SRB_ExecuteIO, ByVal HA As Byte)

Dim dwStatus As Long
Dim s As SRB_Abort

s.SRB_Hdr.SRB_Cmd = SC_ABORT_SRB
s.SRB_Hdr.SRB_HaID = HA
s.SRB_Hdr.SRB_Flags = 0
s.SRB_SRBAbort = VarPtr(SRB)

dwStatus = CallFunc(ASPI_CMD, VarPtr(s))

ASPI32Abort = s.SRB_Hdr.SRB_Status = SS_COMP

End Function

'execute a command through the SPTI
Private Function SPTICmd(ByVal drv As String, ByRef cmd() As Byte, _
        ByVal CmdLen As Integer, ByVal WFE As Boolean, _
        ByVal eDir As e_SPTIDirection, ByVal Pointer As Long, _
        ByVal PointerLen As Long, Optional ByVal Timeout As Long = 5 _
        ) As Boolean

Dim pswb As t_SCSI_PASS_THROUGH_DIRECT_WITH_BUFFER, ol As t_OVERLAPPED
Dim Length As Long, returned As Long, i As Long
Dim Status As Long, fh As Long

'get the drive's handle
If Not GetDriveHandle(drv, fh) Then Exit Function

'check the timeout
If WFE Then Timeout = 9999 Else If Timeout = 0 Then Timeout = 10

With pswb.spt

    .Length = 44                                           ' size of the substructure
    .DataIn = eDir                                         ' data dir
    .TimeOutValue = Timeout                                ' Timeout
    .SenseInfoLength = UBound(pswb.SenseBuffer) - 4        ' Sense Info size
    .SenseInfoOffset = Len(pswb.spt) + 4                   ' Sense Info Offset
    '.SenseInfoOffset = VarPtr(pswb.SenseBuffer(0)) - VarPtr(pswb)

    .DataTransferLength = PointerLen                       ' data len
    .DataBuffer = Pointer                                  ' data buffer pointer

    For i = 0 To CmdLen - 1: .cdb(i) = cmd(i): Next i      ' CDB from Array
    .CdbLength = CmdLen                                    ' CDB len

End With

Length = Len(pswb)                                         ' size of the structure

'SPTD
Status = DeviceIoControl(fh, IOCTL_SCSI_PASS_THROUGH_DIRECT, _
        pswb, Length, pswb, Length, _
        returned, ol)

'close the drive's handle
'CloseHandle fh

'success?
SPTICmd = Status = 1 And pswb.spt.ScsiStatus = 0

intLastSK = pswb.SenseBuffer(2) And &HF
intLastASC = pswb.SenseBuffer(12)
intLastASCQ = pswb.SenseBuffer(13)

'###########################################
#If doDebug Then
Debug.Print "###### SPTI CMD ######"
Debug.Print "======================"
Debug.Print "SCSI Status: " & Hex$(pswb.spt.ScsiStatus) & "h"
Debug.Print "CDB: ";
    For i = 0 To pswb.spt.CdbLength - 1
Debug.Print Format(Hex$(pswb.spt.cdb(i)), "00") & " ";
    Next
Debug.Print ""
Debug.Print "Direction: " & pswb.spt.DataIn
Debug.Print "Data len: " & pswb.spt.DataTransferLength
Debug.Print "SK: " & intLastSK
Debug.Print "ASC: " & intLastASC
Debug.Print "ASCQ: " & intLastASCQ
#End If
'###########################################

End Function

'Einen Befehl ьber ein Interface ausfьhren
Public Function ExecCMD(ByVal DrvID As String, ByRef cmd() As Byte, _
        ByVal CmdLen As Integer, ByVal WFE As Boolean, _
        ByVal eDir As e_AspiDirection, ByVal BufPtr As Long, _
        ByVal buflen As Long, Optional ByVal Timeout As Long = 5 _
        ) As Boolean

Dim DataDir As Long

If blnASPI Then

    'ASPI
    ExecCMD = ASPI32Cmd(DrvID, cmd, CmdLen, eDir, BufPtr, buflen, WFE, Timeout)

Else

    'SPTI
    DataDir = Choose(Abs(CBool(eDir = SRB_DIR_IN)) + 1, _
            SCSI_IOCTL_DATA_OUT, SCSI_IOCTL_DATA_IN)

    ExecCMD = SPTICmd(DrvID, cmd, CmdLen, WFE, DataDir, BufPtr, buflen, Timeout)

End If

End Function

'Get ASPI status and installed Host Adapters
Private Function ASPI32Info() As t_ASPIInfo
Dim lRet As Long

'call ASPI32SupportInfo()
lRet = CallFunc(ASPI_INFO)

With ASPI32Info
    .Status = lRet \ &HFF
    .has = lRet And &HFF
End With

End Function

'convert a drive char to a drive ID
Public Function DrvIDFromDrvChr(ByVal strDrv As String) As String

strDrv = Left$(strDrv, 1)

If blnASPI Then
    With ASPIGetHAID(strDrv)
        DrvIDFromDrvChr = Chr$(.HA) & Chr$(.ID)
    End With
Else
    DrvIDFromDrvChr = strDrv
End If

End Function

'convert drive ID to drive char
Public Function DrvChrFromDrvID(ByVal strDrvID As String) As String

If blnASPI Then

    DrvChrFromDrvID = GetDrvChrASPI32( _
            Asc(Left$(strDrvID, 1)), _
            Asc(Right$(strDrvID, 1)) _
            )

Else

    DrvChrFromDrvID = strDrvID

End If

End Function

'HA:ID from drive char
Public Function HAIDFromDrvChr(ByVal strDrv As String) As String

Dim haid As t_HAID

If blnASPI Then
    With ASPIGetHAID(strDrv)
        HAIDFromDrvChr = Chr$(.HA) & Chr$(.ID)
    End With
Else
    If SPTIGetHAID(strDrv, haid) Then
        With haid
            HAIDFromDrvChr = Chr$(.HA) & Chr$(.ID)
        End With
    End If
End If

End Function

'read host adapter and target ID from a drive
Private Function ASPIGetHAID(ByVal sDrv As String) As t_HAID

Dim i As Long, j As Long
Dim bNT As Boolean

Dim devtype As SRB_GetDevType
Dim osver As OSVERSIONINFOEX
Dim inquiry As SRB_HAInquiry


osver.dwOSVersionInfoSize = Len(osver)      ' structure size
GetVersionEx osver                          ' get windows version

'NT?
If osver.dwPlatformId = VER_PLATFORM_WIN32_NT Then bNT = True

'go through all installed host adapters
For i = 0 To ASPI32Info.has

    inquiry.HA_Hdr.SRB_Cmd = SC_HA_INQUIRY      ' Host Adapter Inquiry
    inquiry.HA_Hdr.SRB_HaID = i                 ' Host Adapter Identifier

    'HA Inquiry
    CallFunc ASPI_CMD, VarPtr(inquiry)

    'Host Adapter ready?
    If inquiry.HA_Hdr.SRB_Status = SS_COMP Then

        For j = 0 To 7

            devtype.SRB_Hdr.SRB_Cmd = SC_GET_DEV_TYPE       ' Device Type Inquiry
            devtype.SRB_Hdr.SRB_HaID = i                    ' Host Adapter Identifier
            devtype.SRB_Target = j                          ' Target Identifier

            'Get Device Type
            CallFunc ASPI_CMD, VarPtr(devtype)

            'drive at target?
            If devtype.SRB_Hdr.SRB_Status = SS_COMP Then

                'CD-ROM?
                If devtype.DEV_DeviceType = 5 Then

                    'get the drive char behind HA:ID
                    If bNT Then
                        If GetDrvChrASPINT(i, j) = sDrv Then

                            ASPIGetHAID.HA = i
                            ASPIGetHAID.ID = j
                            Exit Function

                        End If
                    Else
                        If GetDrvChrASPI32(i, j) = sDrv Then

                            ASPIGetHAID.HA = i
                            ASPIGetHAID.ID = j
                            Exit Function

                        End If
                    End If

                End If

            End If

        Next

    End If

Next

End Function

'get drive char from HA:ID (9x/Me)
Private Function GetDrvChrASPI32(ByVal HA As Byte, ByVal ID As Byte) As String

Dim diskinfo As SRB_GetDiskInfo

diskinfo.SRB_Hdr.SRB_Cmd = SC_GET_DISK_INFO     ' ASPI GetDiskInfo command
diskinfo.SRB_Hdr.SRB_Flags = SRB_DIR_IN         ' get data from the drive
diskinfo.SRB_Hdr.SRB_HaID = HA                  ' Host Adapter
diskinfo.SRB_Target = ID                        ' Target

'call ASPI32SendCommand()
CallFunc ASPI_CMD, VarPtr(diskinfo)

'return drive char
GetDrvChrASPI32 = Chr$(diskinfo.SRB_Int13DrvInfo + 65)

End Function

'HA:ID from drive char (NT/2K/XP)
Private Function SPTIGetHAID(ByVal strDrv As String, ByRef haid As t_HAID) As Boolean

Dim returned As Long, Status As Long
Dim fh As Long     'mzt , i      As Long

'mzt     Dim sd          As SECURITY_ATTRIBUTES
Dim pscsiAddr As t_SCSI_ADDRESS
Dim ol As t_OVERLAPPED


'get drive handle
If GetDriveHandle(Left$(strDrv, 1), fh) Then

    'get SCSI address
    pscsiAddr.Length = Len(pscsiAddr)
    Status = DeviceIoControl(fh, IOCTL_SCSI_GET_ADDRESS, _
            pscsiAddr, Len(pscsiAddr), _
            pscsiAddr, Len(pscsiAddr), _
            returned, ol)

    'CloseHandle fh

    'success?
    If Status = 1 Then
        With pscsiAddr
            haid.HA = .PortNumber
            haid.ID = .TargetId
            SPTIGetHAID = True
        End With
        Exit Function
    End If

End If

End Function

'get drive char behind HA:ID (NT/2K/XP)
Private Function GetDrvChrASPINT(ByVal HA As Byte, ByVal ID As Byte) As String

Dim haid As t_HAID
Dim i As Long

'go through all 26 possible drives
For i = 1 To 26

    'CD-ROM?
    If GetDriveType(Chr(i + 64) & ":") = 5 Then

        'compare to the parameters
        If SPTIGetHAID(Chr$(i + 64), haid) Then
            If haid.HA = HA And haid.ID = ID Then
                'found it!
                GetDrvChrASPINT = Chr$(i + 64)
            End If
        End If

    End If

Next

End Function

'calculate control code
Private Function CTL_CODE(ByVal lDevType As Long, ByVal lFunction As Long, _
        ByVal lMethod As Long, ByVal lAccess As Long) As Long

CTL_CODE = LShift(lDevType, 16) Or LShift(lAccess, 14) Or _
        LShift(lFunction, 2) Or lMethod

End Function

Private Sub Class_Initialize()

Dim osver As OSVERSIONINFOEX

osver.dwOSVersionInfoSize = Len(osver)      ' structure size
GetVersionEx osver                          ' get windows version

'is >= Win 2K?
If osver.dwPlatformId = VER_PLATFORM_WIN32_NT Then _
        blnW2K = osver.dwMajorVersion >= 5

'SPTD Control Code
IOCTL_SCSI_PASS_THROUGH_DIRECT = CTL_CODE(IOCTL_SCSI_BASE, &H405, _
        METHOD_BUFFERED, _
        FILE_READ_ACCESS Or FILE_WRITE_ACCESS)

'SPT Control Code
IOCTL_SCSI_PASS_THROUGH = CTL_CODE(IOCTL_SCSI_BASE, &H401, _
        METHOD_BUFFERED, _
        FILE_READ_ACCESS Or FILE_WRITE_ACCESS)

'SGA Control Code
IOCTL_SCSI_GET_ADDRESS = CTL_CODE(IOCTL_SCSI_BASE, &H406, _
        METHOD_BUFFERED, FILE_ANY_ACCESS)


'--- by Paul Caton ---------------------------

'ObjPtr returns the address of me, at that address is the address of the vtable, copy it to pMe.
Call RtlMoveMemory(pMe, ByVal ObjPtr(Me), 4)

'Replace the stub proc (z_DO_NOT_CALL) with machine-code to handle the cdecl function
'stored in sCode(0). Load the wrapper code into sCode(1)
Call Redirect(CODE_CDECL, True)              ' CDECL code
pCode = StrPtr(Scode(0))                     ' Remember the address of the CDECL code

Call Redirect(CODE_WRAPPER, False)           ' Callback wrapper code, vtable not patched... we don't call it
'----------------------------------------------
End Sub

Public Sub Goodbye()
Class_Terminate
End Sub

Private Sub Class_Terminate()
Dim i As Long

For i = 0 To 25
    CloseHandle lngHandles(i)
Next

Init = False

'DllUnload   ' unload ASPI library
End Sub

'is a bit in a byte set?
Public Function IsBitSet(ByVal InByte As Byte, ByVal Bit As Byte) As Boolean
IsBitSet = ((InByte And (2 ^ Bit)) > 0)
End Function

'MSF to LBA
Public Function MSF2LBA(ByVal mins As Long, ByVal secs As Long, _
        ByVal Frames As Long, _
        Optional positive As Boolean) As Long

MSF2LBA = (mins * 60& + secs) * 75& + Frames

If mins < 90 Or positive Then
    MSF2LBA = MSF2LBA - 150
Else
    MSF2LBA = MSF2LBA - 450150
End If

End Function

'LBA to MSF
Public Function LBA2MSF(ByVal LBA As Long) As String

Dim M As Long, s As Long, F As Long, start As Long

start = Choose(Abs(CBool(LBA >= -150)) + 1, 450150, 150)

M = Fix((LBA + start) / (60 * 75))
s = Fix((LBA + start - M * 60 * 75) / 75)
F = Fix(LBA + start - M * 60 * 75 - s * 75)

LBA2MSF = Format(M, "00") & ":" & Format(s, "00") & ":" & Format(F, "00")

End Function

'>> Operator for VB
'by VB-Accelerator
Public Function RShift(ByVal lThis As Long, ByVal lBits As Long) As Long

Static Init As Boolean

If Not Init Then InitShifting: Init = True

If (lBits <= 0) Then
    RShift = lThis
ElseIf (lBits > 63) Then
    Exit Function
ElseIf (lBits > 31) Then
    RShift = 0
Else
    If (lThis And lngPower2(31)) = lngPower2(31) Then
        RShift = (lThis And &H7FFFFFFF) \ lngPower2(lBits) Or lngPower2(31 - lBits)
    Else
        RShift = lThis \ lngPower2(lBits)
    End If
End If

End Function

'<< Operator for VB
'by VB-Accelerator
Public Function LShift(ByVal lThis As Long, ByVal lBits As Long) As Long

Static Init As Boolean

If Not Init Then InitShifting: Init = True

If (lBits <= 0) Then
    LShift = lThis
ElseIf (lBits > 63) Then
    Exit Function
ElseIf (lBits > 31) Then
    LShift = 0
Else
    If (lThis And lngPower2(31 - lBits)) = lngPower2(31 - lBits) Then
        LShift = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits) Or lngPower2(31)
    Else
        LShift = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits)
    End If
End If

End Function

'Shifting values
'by VB-Accelerator
Private Sub InitShifting()
lngPower2(0) = &H1&: lngPower2(1) = &H2&: lngPower2(2) = &H4&
lngPower2(3) = &H8&: lngPower2(4) = &H10&: lngPower2(5) = &H20&
lngPower2(6) = &H40&: lngPower2(7) = &H80&: lngPower2(8) = &H100&
lngPower2(9) = &H200&: lngPower2(10) = &H400&: lngPower2(11) = &H800&
lngPower2(12) = &H1000&: lngPower2(13) = &H2000&: lngPower2(14) = &H4000&
lngPower2(15) = &H8000&: lngPower2(16) = &H10000: lngPower2(17) = &H20000
lngPower2(18) = &H40000: lngPower2(19) = &H80000: lngPower2(20) = &H100000
lngPower2(21) = &H200000: lngPower2(22) = &H400000: lngPower2(23) = &H800000
lngPower2(24) = &H1000000: lngPower2(25) = &H2000000: lngPower2(26) = &H4000000
lngPower2(27) = &H8000000: lngPower2(28) = &H10000000: lngPower2(29) = &H20000000
lngPower2(30) = &H40000000: lngPower2(31) = &H80000000
End Sub
