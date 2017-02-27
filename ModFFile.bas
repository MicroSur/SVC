Attribute VB_Name = "ModFFile"
Option Explicit
Public PlayMovieFolderFlag As Boolean 'открывать ли папку с фильмом или играть его

'========================================================
' API declarations for the file searching operations
'========================================================
'Public Const FILE_ATTRIBUTE_NORMAL = &H80
'Public Const FILE_ATTRIBUTE_HIDDEN = &H2
'Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_READONLY = &H1

Private Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * 260
  cAlternate        As String * 14
End Type

Public Declare Function FindFirstFile _
    Lib "KERNEL32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile _
    Lib "KERNEL32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "KERNEL32" (ByVal hFindFile As Long) As Long

Public StopSearching As Boolean


Public Sub searchForFile(ByVal startPath As String, ByVal match As String)
'меняет FindFilePath   - как путь к match(файлу)
Dim fPath As String, fname As String, fPathName As String
Dim hfind As Long, nameLen As Integer, matchLen As Integer
Dim WFD As WIN32_FIND_DATA
Dim found As Boolean

Const Dot1 = "."
Const Dot2 = ".."

    FindFilePath = vbNullString
    fPath = LCase$(startPath)
    If Right$(fPath, 1) <> "\" Then fPath = fPath & "\"
    
    matchLen = Len(match)
    match = LCase$(match)
    
    'The first API call is to FindFirstFile.
    '  Note that we get all files with a "*"
    '  and not specify just the file extension
    '  because we need to get the directories too.
    hfind = FindFirstFile(fPath & "*", WFD)
    found = (hfind > 0)
    
    Do While found
    
    DoEvents
         If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then
'         If GetKeyState(vbKeyEscape) < 0 Then
            StopSearching = True
         End If
    
        fname = TrimNull(WFD.cFileName)
        fname = LCase$(fname)
        nameLen = Len(fname)
        fPathName = fPath & fname
'        If fname = "." Or fname = ".." Then
        If fname = Dot1 Or fname = Dot2 Then
         
        ElseIf WFD.dwFileAttributes And _
          FILE_ATTRIBUTE_DIRECTORY Then
          
FrmMain.TextItemHid = fPathName

            searchForFile fPathName, match
            
        ElseIf matchLen = nameLen Then
        If InStrB(fname, match) <> 0 Then
'Debug.Print fPathName
         FindFilePath = fPathName
        Else
         'DoEvents
'         If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then
'         If GetKeyState(vbKeyEscape) < 0 Then
'            StopSearching = True
'         End If
        
        End If

            'Don't do anything if found is too short
'        ElseIf LCase$(Right$(fName, matchLen)) _
'          = match Then
'            'We have an extension match
'            AddAnItem fPathName
'        Else
'            DoEvents
        End If
        
        If FindFilePath <> vbNullString Then Exit Do
        'Subsequent API calls are to FindNextFile.
        If StopSearching Then
            Exit Sub
        Else
            found = FindNextFile(hfind, WFD)
        End If
    Loop
    
    'Then close the findfile operation
    FindClose hfind
    FrmMain.TextItemHid = FindFilePath

End Sub
Public Sub FindFiles(ByVal startPath As String, ByVal ext_match As String, SubFolders As Integer)
'дополняем LstFinded, список автопоиска
Dim fPath As String, fname As String, fPathName As String, fExt As String
Dim hfind As Long    ', nameLen As Integer ', matchLen As Integer
Dim WFD As WIN32_FIND_DATA
Dim found As Boolean
Dim ExtArr() As String
Dim i As Integer

ExtArr = Split(ext_match)    'по пробелу

FindFilePath = vbNullString
fPath = LCase$(startPath)
If Right$(fPath, 1) <> "\" Then fPath = fPath & "\"

ext_match = LCase$(ext_match)

hfind = FindFirstFile(fPath & "*", WFD)
found = (hfind > 0)

Do While found
    DoEvents
    If frmAuto.LstFinded.ListCount > 32700 Then FindClose hfind: Exit Sub

    If GetKeyState(vbKeyEscape) < 0 Then FindClose hfind: Exit Sub
    'If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then FindClose hfind: Exit Sub

    fname = TrimNull(WFD.cFileName)
    'fName = LCase$(fName)
    fExt = LCase$(getExtFromFile(fname))
    fPathName = fPath & fname

    If fname = "." Or fname = ".." Then

    ElseIf WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
        If SubFolders = 1 Then
            'рекурсия
            FindFiles fPathName, ext_match, 1
        End If
    ElseIf InStrB(ext_match, fExt) <> 0 Then
        For i = 0 To UBound(ExtArr)
            If LCase$(fExt) = LCase$(ExtArr(i)) Then
                frmAuto.LstFinded.AddItem fPathName
                Exit For
            End If
        Next i
    ElseIf Len(ext_match) = 0 Then
        'добавлять все
        frmAuto.LstFinded.AddItem fPathName
        'Else
        'DoEvents
    End If


    'If FindFilePath <> vbNullString Then Exit Do
    'Subsequent API calls are to FindNextFile.
    found = FindNextFile(hfind, WFD)
Loop

'Then close the findfile operation
FindClose hfind

End Sub
Public Function GetFirstFileByExt(fPath As String, ext As String) As String
'выдать первый попавшийся файл с данным расширением в данной папке
Dim hfind As Long
Dim fname As String

Dim WFD As WIN32_FIND_DATA

hfind = FindFirstFile(fPath & "*." & ext, WFD)
If hfind > 0 Then
fname = TrimNull(WFD.cFileName)
GetFirstFileByExt = fPath & fname
End If
End Function
Public Function getExtFromFile(F As String) As String
Dim en As Integer
en = InStrRev(F, ".")
If en > 0 Then
    getExtFromFile = Right$(F, Len(F) - en)
End If
End Function
Public Function EraseExtFromFile(F As String) As String
Dim en As Integer
en = InStrRev(F, ".")
If en > 0 Then
    EraseExtFromFile = Left$(F, en - 1)
End If
End Function

Public Function TrimNull(ByVal Item As String) As String
    Dim pos As Integer

    pos = InStr(Item, vbNullChar)
    If pos = 1 Then
        Item = vbNullString
    ElseIf pos > 1 Then
        Item = Left$(Item, pos - 1)
    End If

    TrimNull = Item
End Function
Public Sub PlayMovie(mn As String)
Dim strPath As String
Dim a() As String
Dim i As Integer
'Dim WFD As WIN32_FIND_DATA
Dim ret As Long
Dim tmp As String

strPath = Space$(255)
StopSearching = False

If LenB(mn) = 0 Then GoTo selvideo

'есть ли файл по абс пути в поле файл?
If FileExists(mn) Then
    FindFilePath = mn
Else
    'оставить только название файла
    If InStr(mn, ":") > 0 Then mn = GetNameFromPathAndName(mn)

    'поделить пути в настройках
    If Tokenize04(ComboCDHid_Text, a(), ";,", False) > -1 Then
        FindFilePath = vbNullString
        'сначала поискать по абсолютному пути, склеив Path и mn
        For i = 0 To UBound(a)
            If Right$(a(i), 1) <> "\" Then a(i) = a(i) & "\"
            tmp = a(i) & mn
            If FileExists(tmp) Then
                FindFilePath = tmp
                Exit For
            End If
        Next i
        'Debug.Print FindFilePath
        'потом поискать по всем путям и подпапкам

        If Len(FindFilePath) = 0 Then
            For i = 0 To UBound(a)
                searchForFile a(i), mn
                If Len(FindFilePath) > 0 Then
                    'FindFilePath = tmp
                    Exit For
                End If
            Next i
        End If
        'Debug.Print FindFilePath


        If FindFilePath <> vbNullString Then
            If PlayMovieFolderFlag Then
                'показать папку и выйти
                tmp = GetPathFromPathAndName(FindFilePath)
                ret = ShellExecute(GetDesktopWindow(), "open", tmp, vbNull, vbNull, 1)
                PlayMovieFolderFlag = False
                Exit Sub
            Else
                'найти плеер
                ret = FindExecutable(FindFilePath, "", strPath)
                ToDebug "медиаплеер: " & ModFFile.TrimNull(strPath)
            End If
        End If

    End If    'Tokenize04

    If FindFilePath = vbNullString Then

selvideo:
        'не нашли ни в пути, ни на CD - попросить указать

        'не идем дальше пока нажат Esc
        Do While GetAsyncKeyState(vbKeyEscape) And &H1 = &H1: DoEvents: Loop

        FindFilePath = pLoadDialog(NamesStore(12), mn)
        If LenB(FindFilePath) = 0 Then
            Exit Sub
        Else
            ret = FindExecutable(FindFilePath, "", strPath)
        End If
    End If    'FindFilePath = vbNullString

    Select Case ret
    Case 31
        myMsgBox msgsvc(1), vbInformation, , FrmMain.hwnd
        Exit Sub
    Case 2
        myMsgBox msgsvc(2), vbInformation, , FrmMain.hwnd
        Exit Sub
    Case Is <> 42
        myMsgBox msgsvc(3), vbInformation, , FrmMain.hwnd
        Exit Sub
    End Select

End If    ' нашли по пути в поле файл или текущая папка случайно та, что надо

On Error Resume Next

If PlayMovieFolderFlag Then
    'открыть папку и выйти
    tmp = GetPathFromPathAndName(FindFilePath)
    ret = ShellExecute(GetDesktopWindow(), "open", tmp, vbNull, vbNull, 1)
    PlayMovieFolderFlag = False
    Exit Sub
Else
    'проиграть файл
    ret = ShellExecute(GetDesktopWindow(), "open", FindFilePath, vbNull, vbNull, 1)
    'Private Const SW_SHOWNORMAL As Long = 1
    Call err.Clear
    'If ret = 42 Then FrmMain.WindowState = vbMinimized
    FrmMain.Timer2.Enabled = False
    
End If

End Sub



