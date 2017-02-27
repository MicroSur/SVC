Attribute VB_Name = "INI"
Option Explicit
Public Declare Function GetPrivateProfileStringByKeyName& Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Public Declare Function GetPrivateProfileStringKeys& Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
'Public Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)

' This first line is the declaration from win32api.txt

Public Declare Function WritePrivateProfileStringByKeyName& Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Public Declare Function WritePrivateProfileStringToDeleteKey& Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
'Public Declare Function WritePrivateProfileStringToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)
'Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Function VBGetPrivateProfileString(section As String, Key As String, file As String, Optional ByRef SameStr As String) As String

Dim KeyValue As String
Dim characters As Long

'KeyValue = String$(1024, 0)
KeyValue = AllocString_ADV(1024)

characters = GetPrivateProfileStringByKeyName(section, Key, "=", KeyValue, 1024, file)
'    Debug.Print Asc(KeyValue)
If (AscW(KeyValue) = 0) Or (Left$(KeyValue, 1) = "=") Then     ' не нашли
    VBGetPrivateProfileString = SameStr
    Exit Function
End If

If characters > 0 Then
    KeyValue = Left$(KeyValue, characters)
End If

VBGetPrivateProfileString = KeyValue

End Function


'Public Function GetSectionNames(filename As String, SectionNames As Variant) As Integer
'    'GetSectionNames Return Number of Section in file
'    'SectionNames return all section names
'
'    Dim characters As Long
'    Dim SectionList As String
'    Dim ArrSection() As String
'    Dim i As Integer
'    Dim NullOffset%
'
'    SectionList = String$(128, 0)
'
'    ' Retrieve the list of keys in the section
'    characters = GetPrivateProfileStringSections(0, 0, "", SectionList, 127, filename)
'
'    ' Load sections into Arrey
'    i = 0
'    Do
'        NullOffset% = InStr(SectionList, vbNullChar)
'        If NullOffset% > 1 Then
'            ReDim Preserve ArrSection(i)
'            ArrSection(i) = Mid$(SectionList, 1, NullOffset% - 1)
'            SectionList$ = Mid$(SectionList, NullOffset% + 1)
'            i = i + 1
'        End If
'   Loop While NullOffset% > 1
'    GetSectionNames = i - 1
'    SectionNames = ArrSection
'
'
'End Function

Public Function GetKeyNames(SectionName As String, filename As String, KeyNames As Variant) As Integer
'GetKeyNames Return Number of key in section или -1
'KeyNames Return list of keyNames in section массив

Dim characters As Long
Dim KeyList As String
Dim ArrName() As String

Dim i As Integer

KeyList = String$(128, 0)
' Retrieve the list of keys in the section

characters = GetPrivateProfileStringKeys(SectionName, 0, "", KeyList, 127, filename)

' Load Keys into Arrey
Dim NullOffset%
i = 0
Do
    NullOffset% = InStr(KeyList, vbNullChar)
    If NullOffset% > 1 Then
        ReDim Preserve ArrName(i)
        ArrName(i) = Mid$(KeyList, 1, NullOffset% - 1)
        KeyList$ = Mid$(KeyList, NullOffset% + 1)
        i = i + 1
    End If
Loop While NullOffset% > 1
GetKeyNames = i - 1
KeyNames = ArrName

End Function

Public Function DeleteKey(KeyName As String, SectionName As String, filename As String) As Long
    'Return 0 if Deletion not sucsesful или нечего удалять (не нашел)
    ' Delete the selected key
DeleteKey = WritePrivateProfileStringToDeleteKey(SectionName, KeyName, 0, filename)
End Function

Public Function WriteKey(SectionName As String, KeyName As String, KeyValue As String, filename As String) As Long
If Len(KeyValue) = 0 Then KeyValue = vbNullString
WriteKey = WritePrivateProfileStringByKeyName(SectionName, KeyName, KeyValue, filename)
End Function

'Public Function WriteSection(SectionName As String, filename As String) As Long
'    WriteSection = WritePrivateProfileSection(SectionName, "", filename)
'End Function
'
'Public Function DeleteSection(SectionName, filename) As Long
'    DeleteSection = WritePrivateProfileStringToDeleteSection(SectionName, 0&, 0&, filename)
'End Function
    
Public Sub SaveInterface()
'пишем измененные параметры интерфейса, например при переходе по табам баз или выходе
Dim WFD As WIN32_FIND_DATA
Dim ret As Long
Dim i As Integer


If Not INIFileFlagRW Then Exit Sub

On Error Resume Next

'Debug.Print "SaveInterface в " & INIFILE

'                                              Current Base ini
iniFileName = App.Path
If Right$(iniFileName, 1) <> "\" Then iniFileName = iniFileName & "\"
iniFileName = iniFileName & INIFILE
ret = FindFirstFile(iniFileName, WFD)
If ret < 0 Then MakeINI INIFILE
FindClose ret

'ToDebug "Сохранение интерфейса в " & INIFILE

'                                                                  LV
WriteKey "LIST", "LVWidth%", CStr(LVWidth), iniFileName
WriteKey "LIST", "TVWidth", CStr(TVWidth), iniFileName '            tv
WriteKey "LIST", "ScrShotWidth%", CStr(SplitLVD), iniFileName '            ss

For i = 1 To lvHeaderIndexPole
WriteKey "LIST", "C" & i, FrmMain.ListView.ColumnHeaders(i).Width, iniFileName
WriteKey "LIST", "P" & i, FrmMain.ListView.ColumnHeaders(i).Position, iniFileName
Next i

WriteKey "LIST", "LVSortColl", CStr(LVSortColl), iniFileName
WriteKey "LIST", "LVSortOrder", CStr(LVSortOrder), iniFileName

'                                                           Обложка
WriteKey "COVER", "txt_Stan_L", CStr(cov_stan.l), iniFileName
WriteKey "COVER", "txt_Stan_T", CStr(cov_stan.t), iniFileName
WriteKey "COVER", "txt_Stan_W", CStr(cov_stan.w), iniFileName
WriteKey "COVER", "txt_Stan_H", CStr(cov_stan.H), iniFileName
WriteKey "COVER", "txt_Conv_L", CStr(cov_conv.l), iniFileName
WriteKey "COVER", "txt_Conv_T", CStr(cov_conv.t), iniFileName
WriteKey "COVER", "txt_Conv_W", CStr(cov_conv.w), iniFileName
WriteKey "COVER", "txt_Conv_H", CStr(cov_conv.H), iniFileName
WriteKey "COVER", "txt_Dvd_L", CStr(cov_dvd.l), iniFileName
WriteKey "COVER", "txt_Dvd_T", CStr(cov_dvd.t), iniFileName
WriteKey "COVER", "txt_Dvd_W", CStr(cov_dvd.w), iniFileName
WriteKey "COVER", "txt_Dvd_H", CStr(cov_dvd.H), iniFileName
WriteKey "COVER", "txt_List_L", CStr(cov_list.l), iniFileName
WriteKey "COVER", "txt_List_T", CStr(cov_list.t), iniFileName
WriteKey "COVER", "txt_List_W", CStr(cov_list.w), iniFileName
WriteKey "COVER", "txt_List_H", CStr(cov_list.H), iniFileName
'If err.Number = 0 Then
'ToDebug "...ok"
'Else
'ToDebug "...error: " & err.Description
'End If

Call err.Clear
End Sub


