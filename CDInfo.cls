VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptoCDInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************
'*** Flamed CD-R(W/OM) information
'***
'*** reads information about a CD-R/ROM/RW
'*******************************************
'e_Status
'Public Enum FL_Status
'    STAT_EMPTY                  ' empty
'    STAT_INCOMPLETE             ' uncomplete
'    STAT_COMPLETE               ' complete
'    STAT_UNKNWN                 ' unknown
'End Enum

'e_CD_SubType
'Public Enum FL_CDSubType
'    STYPE_CDROMDA               ' CD-ROM or CDDA
'    STYPE_CDI                   ' CD-I
'    STYPE_XA                    ' CD-XA
'    STYPE_UNKNWN                ' unknown
'End Enum

'e_CDType
'Private Enum FL_CDType
'    ROMTYPE_CDROM               ' CD-ROM
'    ROMTYPE_CDR                 ' CD-R
'    ROMTYPE_CDRW                ' CD-RW
'    ROMTYPE_CDROM_R_RW          ' CD-ROM, CD-R or CD-RW
'    ROMTYPE_DVD_ROM             ' DVD-ROM
'    ROMTYPE_DVD_R               ' DVD-R
'    ROMTYPE_DVD_RW              ' DVD-RW
'    ROMTYPE_DVD_RAM             ' DVD-RAM
'    ROMTYPE_DVD_P_R             ' DVD+R
'    ROMTYPE_DVD_P_RW            ' DVD+RW
'End Enum

Private udeMediaType    As e_CDType 'FL_CDType 'e_CDType
Private udeCDRWType     As e_CD_SubType

Private udeMediaStatus  As e_Status
Private udeLastSessionState As e_Status

Private lngCapacity     As Long
Private lngSize         As Long

Private lngLeadInLBA    As New cOptoMSF
Private lngLeadOutLBA   As New cOptoMSF

Private intSessions     As Integer
Private intTracks       As Integer

Private blnErasable     As Boolean

Private strCDRWVendor   As String

Public Property Get Capacity() As Long
    Capacity = lngCapacity
End Property

Public Property Get Size() As Long
    Size = lngSize
End Property

Public Property Get CDRWVendor() As String
    CDRWVendor = strCDRWVendor
End Property

Public Property Get Erasable() As Boolean
    Erasable = blnErasable
End Property

Public Property Get Tracks() As Integer
    Tracks = intTracks
End Property

Public Property Get Sessions() As Integer
    Sessions = intSessions
End Property

Public Property Get LeadInMSF() As cOptoMSF
    Set LeadInMSF = lngLeadInLBA
End Property

Public Property Get LeadOutMSF() As cOptoMSF
    Set LeadOutMSF = lngLeadOutLBA
End Property

Public Property Get LastSessionState() As e_CDType 'FL_Status
    LastSessionState = udeLastSessionState
End Property

Public Property Get MediaStatus() As e_CDType 'FL_Status
    MediaStatus = udeMediaStatus
End Property

Public Property Get CDRWType() As e_CD_SubType 'FL_CDSubType
    CDRWType = udeCDRWType
End Property

Public Property Get MediaType() As e_CDType 'FL_CDType
    MediaType = udeMediaType
End Property

Public Function GetInfo(ByVal DriveID As String) As Boolean

' drive is ready?
If Not CDRomTestUnitReady(DriveID) Then Exit Function

On Error Resume Next    'ошибка определения cd-dvd - размер

With CDRomGetCDInfo(DriveID)

    udeCDRWType = .CDSubType
    udeMediaType = .CDType
    '        udeMediaStatus = .DiscStatus
    '        udeLastSessionState = .LastSessionStatus

    '        lngCapacity = .Capacity
    lngSize = .Size

    '        lngLeadInLBA.LBA = .LeadIn.LBA
    '        lngLeadOutLBA.LBA = .LeadOut.LBA

    '        intSessions = .Sessions
    '        intTracks = .Tracks

    '        blnErasable = .Erasable
    '        strCDRWVendor = .Vendor

End With

If err.Number = 0 Then GetInfo = True
err.Clear
End Function

