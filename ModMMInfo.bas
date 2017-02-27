Attribute VB_Name = "MMInfo"
Option Explicit

#If MEDIAINFO_NO_ENUMS Then
#Else

Public Enum MediaInfo_stream_C
  MediaInfo_Stream_General
  MediaInfo_Stream_Video
  MediaInfo_Stream_Audio
  MediaInfo_Stream_Text
  MediaInfo_Stream_Chapters
  MediaInfo_Stream_Image
  MediaInfo_Stream_Max
End Enum

Public Enum MediaInfo_info_C
  MediaInfo_Info_Name
  MediaInfo_Info_Text
  MediaInfo_Info_Measure
  MediaInfo_Info_Options
  MediaInfo_Info_Name_Text
  MediaInfo_Info_Measure_Text
  MediaInfo_Info_Info
  MediaInfo_Info_HowTo
  MediaInfo_Info_Max
End Enum

Public Enum MediaInfo_infooptions_C
  MediaInfo_InfoOption_ShowInInform
  MediaInfo_InfoOption_Support
  MediaInfo_InfoOption_ShowInSupported
  MediaInfo_InfoOption_TypeOfValue
  MediaInfo_InfoOption_Max
End Enum

Public Enum MediaInfo_informoptions_C
  MediaInfo_InformOption_Nothing
  MediaInfo_InformOption_Custom
  MediaInfo_InformOption_HTML
  MediaInfo_InformOption_Max
End Enum

#End If

Public Declare Sub MediaInfo_Close Lib "MediaInfo.dll" (ByVal Handle As Long)
'Public Declare Sub MediaInfo_Delete Lib "MediaInfo.dll" (ByVal Handle As Long)
Public Declare Function MediaInfo_Count_Get Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal StreamKind As MediaInfo_stream_C, ByVal StreamNumber As Long) As Long
Public Declare Function MediaInfo_Get Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal StreamKind As MediaInfo_stream_C, ByVal StreamNumber As Long, ByVal Parameter As Long, ByVal InfoKind As MediaInfo_info_C, ByVal SearchKind As MediaInfo_info_C) As Long
'Public Declare Function MediaInfo_GetI Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal StreamKind As MediaInfo_stream_C, ByVal StreamNumber As Long, ByVal Parameter As Long, ByVal InfoKind As MediaInfo_info_C) As Long
'Public Declare Function MediaInfo_Inform Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal Options As MediaInfo_informoptions_C) As Long
Public Declare Function MediaInfo_New Lib "MediaInfo.dll" () As Long
'Public Declare Function MediaInfo_New_Quick Lib "MediaInfo.dll" (ByVal file As Long, ByVal Options As Long) As Long
Public Declare Function MediaInfo_Open Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal file As Long) As Long
'Public Declare Function MediaInfo_Open_Buffer Lib "MediaInfo.dll" (ByVal Handle As Long, Begin As Any, ByVal Begin_Size As Long, End_ As Any, ByVal End_Size As Long) As Long
Public Declare Function MediaInfo_Option Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal Option_ As Long, ByVal Value As Long) As Long
'Public Declare Function MediaInfo_Save Lib "MediaInfo.dll" (ByVal Handle As Long) As Long
'Public Declare Function MediaInfo_Set Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal ToSet As Long, ByVal StreamKind As MediaInfo_stream_C, ByVal StreamNumber As Long, ByVal Parameter As Long, ByVal OldParameter As Long) As Long
'Public Declare Function MediaInfo_SetI Lib "MediaInfo.dll" (ByVal Handle As Long, ByVal ToSet As Long, ByVal StreamKind As MediaInfo_stream_C, ByVal StreamNumber As Long, ByVal Parameter As Long, ByVal OldParameter As Long) As Long
'Public Declare Function MediaInfo_State_Get Lib "MediaInfo.dll" (ByVal Handle As Long) As Long

Public Declare Function lstrlenA Lib "kernel32" (ByVal ptr As Any) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal ptr As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal pStr As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bLen As Long)

Public Function bstr(ptr As Long) As String
' convert a C wchar* to a Visual Basic string
  Dim l As Long
  l = lstrlenW(ptr)
  bstr = String$(l, vbNullChar)
  RtlMoveMemory ByVal StrPtr(bstr), ByVal ptr, l * 2
End Function


Public Function IFObstr(ByVal lpszA As Long) As String
    IFObstr = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal IFObstr, ByVal lpszA)
End Function

'Public Function bstrW(lpszString As Long) As String
'  Dim lpszStr1 As String, lpszStr2 As String, nRet As Long
'  lpszStr1 = String(1000, "*")
'  nRet = lstrcpyW(lpszStr1, lpszString)
'  lpszStr2 = (StrConv(lpszStr1, vbFromUnicode, LCID))
'  bstrW = Left$(lpszStr2, InStr(lpszStr2, Chr$(0)) - 1)
'End Function

Public Function FormatTime(ByVal inTime As Long) As String
'обратная функция time2sec
Dim TimeLeft As Long, ThisTime As Long

Const OneMinute As Long = 60
Const OneHour As Long = OneMinute ^ 2
Const OneDay As Long = OneHour * 24

TimeLeft = inTime

If (TimeLeft >= OneDay) Then
    ThisTime = inTime \ OneDay
    FormatTime = CStr(ThisTime) & ":"
    TimeLeft = inTime Mod OneDay
End If

If (TimeLeft >= OneHour) Then
    ThisTime = TimeLeft \ OneHour
    FormatTime = FormatTime & Format$(ThisTime, "00") & ":"     'FormatTime & (IIf(FormatTime <> vbNullString, Format$(ThisTime, "00"), ThisTime)) & ":"
    TimeLeft = TimeLeft Mod OneHour
Else
    FormatTime = FormatTime & "00:"
End If

FormatTime = FormatTime & (Format(TimeLeft \ OneMinute, "00") & ":" & Format$(TimeLeft Mod OneMinute, "00"))
'FormatTime = Format$(FormatTime, "##:##:##")
End Function
Public Function Chnnls(n As String) As String

Select Case n
Case 1
    Chnnls = "Mono"
Case 2
    Chnnls = "Stereo"
Case Else
    Chnnls = n
    If LCase$(right$(Chnnls, 2)) <> LCase$("ch") Then Chnnls = n & "ch"
End Select

End Function

Public Function MyCodec(s As String) As String
'для красоты отображения
'см Public Function strMPEG() As String
Select Case UCase$(s)
Case "MPV1", "MPEG-1V"
    MyCodec = "MPEG1"
Case "MPV2", "MPEG-2V"
    MyCodec = "MPEG2"
Case Else
    MyCodec = s
End Select

End Function
