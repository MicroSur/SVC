Attribute VB_Name = "VStrP"
Option Explicit

'function ifoOpen(const name: PChar; const fio_flags: Cardinal): Cardinal; cdecl; external vs_DllName;
Public Declare Function ifoOpen Lib "vstrip.dll" Alias "_ifoOpen@8" (ByVal name As String, ByVal fio_flags As Long) As Long
'function ifoClose(const ifo: Cardinal): Boolean; cdecl; external vs_DllName;
Public Declare Function ifoClose Lib "vstrip.dll" Alias "_ifoClose@4" (ByVal ifo As Long) As Boolean

Public Const fio_USE_ASPI = 4
'Audio
'function ifoGetNumAudio(const ifo: Cardinal): Integer; cdecl; external vs_DllName;
Public Declare Function ifoGetNumAudio Lib "vstrip.dll" Alias "_ifoGetNumAudio@4" (ByVal ifo As Long) As Long
Public Declare Function ifoGetAudioDesc Lib "vstrip.dll" Alias "_ifoGetAudioDesc@8" (ByVal ifo As Long, ByVal audio_idx As Long) As Long

'Lang (from Audio)
Public Declare Function ifoGetLangDesc Lib "vstrip.dll" Alias "_ifoGetLangDesc@8" (ByVal ifo As Long, ByVal audio_idx As Long) As Long

'SubTitle
Public Declare Function ifoGetNumSubPic Lib "vstrip.dll" Alias "_ifoGetNumSubPic@4" (ByVal ifo As Long) As Long
Public Declare Function ifoGetSubPicDesc Lib "vstrip.dll" Alias "_ifoGetSubPicDesc@8" (ByVal ifo As Long, ByVal subp_idx As Long) As Long

'Video
Public Declare Function ifoGetVideoDesc Lib "vstrip.dll" Alias "_ifoGetVideoDesc@4" (ByVal ifo As Long) As Long

'
'function ifoGetPGCIInfo(const ifo: Cardinal; const title: Cardinal; var time_out: t_vs_time): Integer; cdecl; external vs_DllName;
't_vs_time = array[0..3] of Byte; //hh.mm.ss.ms
Public Declare Function ifoGetPGCIInfo Lib "vstrip.dll" Alias "_ifoGetPGCIInfo@12" (ByVal ifo As Long, ByVal title As Long, time_out As Byte) As Long
'function ifoGetNumPGCI(const ifo: Cardinal): Integer; cdecl; external vs_DllName;
Public Declare Function ifoGetNumPGCI Lib "vstrip.dll" Alias "_ifoGetNumPGCI@4" (ByVal ifo As Long) As Long

'String
Public Declare Function lstrlenA Lib "KERNEL32" (ByVal ptr As Any) As Long
Public Declare Function lstrcpyA Lib "KERNEL32" (ByVal RetVal As String, ByVal ptr As Long) As Long
'Public Declare Function lstrlenW Lib "kernel32" (ByVal Ptr As Long) As Long
'Public Declare Function lstrcpyW Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long


'Public Function bstr(ByVal lpszA As Long) As String
'    bstr = String$(lstrlenA(ByVal lpszA), 0)
'    Call lstrcpyA(ByVal bstr, ByVal lpszA)
'End Function

'Public Function bstrW(lpszString As Long) As String
'  Dim lpszStr1 As String, lpszStr2 As String, nRet As Long
'  lpszStr1 = String(1000, "*")
'  nRet = lstrcpyW(lpszStr1, lpszString)
'  lpszStr2 = (StrConv(lpszStr1, vbFromUnicode, LCID))
'  bstrW = Left$(lpszStr2, InStr(lpszStr2, Chr$(0)) - 1)
'End Function
