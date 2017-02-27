Attribute VB_Name = "modRS"
Option Explicit
Public Function RestoreBasePos() As Boolean
'������������ �����. ���. � ����
'���� ��� GetCoverSpisok, ��� ������ ���������� ���
'����� �� ������ ������� ��������, ���� ���������� ���� �������� (������ �����... �������...)
If rs Is Nothing Then Exit Function
If FrmMain.ListView.ListItems.Count < 1 Then Exit Function
If CurSearch > FrmMain.ListView.ListItems.Count Then Exit Function
If CurSearch < 1 Then CurSearch = 1

With FrmMain.ListView
'If Not MultiSel Then
    'Set ListView.SelectedItem = ListView.ListItems(CurSearch)
    'RSGoto ListView.SelectedItem.Key
    RSGoto .ListItems(CurSearch).Key '�� ������� ���������
    
    If .Visible Then
        'ListView.SetFocus
        .ListItems(CurSearch).EnsureVisible
    End If

    If Val(.SelectedItem.Key) = rs("Key") Then
        RestoreBasePos = True
    Else
        ToDebug "ERR_RestoreB: " & CurSearch
    End If
'Else
'End If
End With
End Function
Public Function OpenRS() As Boolean
ToDebug "Open table"
On Error Resume Next

Set rs = DB.OpenRecordset("Storage", dbOpenTable)
If err.Number = 0 Then
    rs.Index = "Key"
    'ToDebug "Key=" & rs.Index

ElseIf err.Number = 3011 Then
    ToDebug "Error: � ���� ��� ������ �������: Storage"
    myMsgBox msgsvc(20), vbCritical, , FrmMain.hwnd:  '�� ���� ������� ����.\�����? �� SVC ����? ���� ������� � MS Access?
    Set FrmMain.TabLVHid.SelectedItem = FrmMain.TabLVHid.Tabs(oldTabLVInd)
    OpenRS = False
    Exit Function
Else     'error
    ToDebug err.Number & " - " & err.Description
    myMsgBox msgsvc(26), vbCritical, , FrmMain.hwnd  '� ����� �������� ������ ���������?\�������� ������ �� ����?
    '������� �� ��������� ���������� ����
    FrmMain.FrameView.Caption = FrameViewCaption & " 0 )"
    Set FrmMain.TabLVHid.SelectedItem = FrmMain.TabLVHid.Tabs(oldTabLVInd)
    OpenRS = False
    Exit Function
End If

OpenRS = True
End Function
Public Function OpenDB() As Boolean

Dim WFD As WIN32_FIND_DATA
Dim ret As Long
Dim temp As String
Dim NotSVCBase As Integer    '������� ���� �� ����

Dim fld As DAO.Field
'Const dbText As Integer = 10
'Const dbChar = 18
Dim tdt As DAO.TableDef
'Dim cfFlag As Boolean

'DoEvents

BaseReadOnly = False: BaseReadOnlyU = False
OpenDB = False
NoDBFlag = True

ret = FindFirstFile(bdname, WFD)
If LenB(bdname) = 0 Or ret < 0 Then
'If Not FileExists(bdname) Then
    ToDebug "Err_NoDBFile!"
    myMsgBox msgsvc(5), vbInformation, , FrmMain.hwnd    '��� ����� � �����
    bdname = vbNullString
End If
FindClose ret

If Len(bdname) = 0 Then Exit Function

If WFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
    'ToDebug "������� ���� - ������ ������"
    'myMsgBox msgsvc(24), vbInformation, , Me.hwnd '������ �� ������ - ����� ����� �������� ������
    BaseReadOnly = True
End If


GetExtensionFromFileName bdname, temp
temp = temp & ".ldb"
Call KillLdb(temp)
'ret = FindFirstFile(temp, WFD)
If FileExists(temp) And Not BaseReadOnly Then
    ToDebug "���� ���� " & temp & " - � ����� ��������!"
    'myMsgBox msgsvc(22), vbInformation
    BaseReadOnlyU = True
    'BaseReadOnly = True
End If
'FindClose ret

On Error Resume Next
'Dim ROFlag As Boolean
'If BaseReadOnly Then ROFlag = True 'Or BaseReadOnlyU Then

ToDebug "Open DB (RO=" & BaseReadOnly & " ROU=" & BaseReadOnlyU & ")"

Set DB = DBEngine.OpenDatabase(bdname, BaseReadOnly, BaseReadOnly)
If err.Number = 3051 Then
    Set DB = DBEngine.OpenDatabase(bdname, True, True)    '��� ����
    BaseReadOnly = True
End If
'Set DB = DBEngine.OpenDatabase(bdname, True, True)
'Debug.Print err.Number
'3031 - �� ��� ������
'3043 - ���� ���� ����

If err.Number = 3031 Then     '���� ������
    ToDebug "DB������?"
    'If Len(pwd) = 0 Then
    SetTimer FrmMain.hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
    pwd = myInputBox(NamesStore(5) & vbCrLf & bdname)       ', vbNullString)
    'End If
    Set DB = DBEngine.OpenDatabase(bdname, False, BaseReadOnly, "MS Access;PWD=" & pwd)
    'tdfRegionOne.Attributes = dbAttachSavePWD
End If
If err.Number = 3043 Then ToDebug "������ ����� ��� ����."
err.Clear

'-------------------------------------------------��������� ������� ���� �������

If Not OpenRS Then Exit Function

'�������� � �����
NotSVCBase = SVCBaseFielsCount    ' 24 '����, ����� ���������� 1, ���� ���������� �� 0 - ���� ������ ��� �� ����
Set tdt = DB.TableDefs("Storage")
For Each fld In tdt.Fields
    If fld.OrdinalPosition = 0 Then dbFirstField = fld.name
    Select Case LCase$(fld.name)
        Case "moviename"
            dbMovieNameInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1   '
        Case "label"
            dbLabelInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1       '
        Case "genre"
            dbGenreInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1       '
        Case "year"
            dbYearInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1        '
        Case "country"
            dbCountryInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1     '
        Case "director"
            dbDirectorInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1    '
        Case "acter"
            dbActerInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1       '
        Case "time"
            dbTimeInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1        '
        Case "resolution"
            dbResolutionInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1  '
        Case "audio"
            dbAudioInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1       '
        Case "fps"
            dbFpsInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1         '
        Case "filelen"
            dbFileLenInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1     '
        Case "cdn"
            dbCDNInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1         '
        Case "video"
            dbVideoInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "filename"
            dbFileNameInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "debtor"
            dbDebtorInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "sndisk"
            dbsnDiskInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "other"
            dbOtherInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "annotation"
            dbAnnotationInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "checked"
            dbCheckedInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1

        Case "snapshot1"
            dbSnapShot1Ind = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "snapshot2"
            dbSnapShot2Ind = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "snapshot3"
            dbSnapShot3Ind = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "frontface"
            dbFrontFaceInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1

        Case "key"
            dbKeyInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1

        Case "subtitle"
            dbSubTitleInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "coverpath"
            dbCoverPathInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "movieurl"
            dbMovieURLInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "rating"
            dbRatingInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "mediatype"
            dbMediaTypeInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1
        Case "language"
            dbLanguageInd = fld.OrdinalPosition: NotSVCBase = NotSVCBase - 1

    End Select
Next

Set tdt = Nothing
Set fld = Nothing

If NotSVCBase = 0 Then    '���� �������
    FrmMain.Caption = "SurVideoCatalog " & bdname
    OpenDB = True
    NoDBFlag = False
    LastBaseIsGood = True
Else


    ToDebug "Error: ���� ������ ��� �� SVC ����. �������������� ���������� Convert2SVC"
    myMsgBox msgsvc(32), vbCritical, , FrmMain.hwnd
    DoEvents
    
If LastBaseIsGood Then
    LastBaseIsGood = False
    InitFlag = True
    Set FrmMain.TabLVHid.SelectedItem = FrmMain.TabLVHid.Tabs(oldTabLVInd)
End If

    Exit Function

End If

Exit Function
err:

End Function


Public Sub OpenActDB()
Dim WFD As WIN32_FIND_DATA
Dim ret As Long
Dim fn As Integer
Dim a() As Byte
Dim temp As String

abdname = App.Path & "\people.mdb"
ret = FindFirstFile(abdname, WFD)
If WFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
    ToDebug "ABD RO"
    'myMsgBox msgsvc(25), vbInformation, , Me.hwnd '���� � ����� ������� ������ ������ �� ������
    BaseAReadOnly = True
End If

If ret < 0 Then    '������� �������
    ToDebug "��� ���� �������, �������..."
    'myMsgBox msgsvc(19), vbInformation, , Me.hwnd
    a() = LoadResData("PEOPLE", "CUSTOM")
    fn = FreeFile
    Open abdname For Binary Access Write As fn
    Put #fn, , a()
    Close #fn
End If
FindClose ret

GetExtensionFromFileName abdname, temp
temp = temp + ".ldb"
Call KillLdb(temp)
ret = FindFirstFile(temp, WFD)
If ret > 0 Then
    ToDebug "ABD RO(2)"
    'myMsgBox msgsvc(23), vbInformation, , Me.hwnd '���� � ����� ������� ������������ ������ ����������
    BaseAReadOnlyU = True
End If
FindClose ret

On Error Resume Next

If BaseAReadOnly Then    'Or BaseAReadOnlyU Then
    ToDebug "Open RO ADB"
    Set ADB = DBEngine.OpenDatabase(abdname, False, True)
Else
    ToDebug "Open ADB"
    Set ADB = DBEngine.OpenDatabase(abdname, False, False)
End If
If (err.Number = 3051) Or (err.Number = 3050) Then
    Set ADB = DBEngine.OpenDatabase(abdname, True, True)     '��� ����
    BaseAReadOnly = True
End If

err.Clear

If abdname <> vbNullString Then

    ToDebug "Open Acter"
    Set ars = ADB.OpenRecordset("Acter", dbOpenTable)
    If err.Number = 0 Then
        ars.Index = "KeyAct"
        'ToDebug "Key=" & ars.Index
    Else    'error
        ToDebug "Err_Exit: " & msgsvc(27)
        myMsgBox msgsvc(27), vbCritical, , FrmMain.hwnd: End         '� ����� ������� �������� ������ ���������. �����.
    End If
End If

actInitFlag = False

Call err.Clear
On Error GoTo 0
End Sub
Public Function OpenDBdd() As Boolean

Dim WFD As WIN32_FIND_DATA
Dim ret As Long
Dim temp As String, tmp As String

'Dim cfFlag As Boolean

OpenDBdd = False


tmp = LstBases_List(MouseOverTabLV)
ret = FindFirstFile(tmp, WFD)

If Len(tmp) = 0 Or ret < 0 Then
ToDebug "Err_NoDB2DragDrop"
    myMsgBox msgsvc(5) & " " & tmp, vbInformation, , FrmMain.hwnd '��� ����� � �����
    Exit Function
End If
FindClose ret

If WFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
ToDebug "DragDrop RO!"
    myMsgBox msgsvc(24) & " " & tmp, vbInformation, , FrmMain.hwnd '������ �� ������
    Exit Function
End If

GetExtensionFromFileName tmp, temp
temp = temp + ".ldb"
Call KillLdb(temp)
ret = FindFirstFile(temp, WFD)
If ret > 0 Then
ToDebug "Exists " & temp
    myMsgBox msgsvc(22), vbInformation, , FrmMain.hwnd
    Exit Function
End If
FindClose ret

On Error Resume Next

ToDebug "Open DragDrop DB (" & tmp & ")"
    Set DBdd = DBEngine.OpenDatabase(tmp, False, False)
    'Debug.Print err.Number
    '3031 - �� ��� ������
    '3043 - ���� ���� ����
    If err.Number = 3031 Then '���� ������
        ToDebug "DragDrop password?"
        'If Len(pwd) = 0 Then
         SetTimer FrmMain.hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
         pwd = myInputBox(NamesStore(5) & vbCrLf & tmp)  ', vbNullString)
        'End If
        Set DBdd = DBEngine.OpenDatabase(tmp, False, False, "MS Access;PWD=" & pwd)
    End If
    If err.Number = 3043 Then ToDebug "������ ����� ��� ����."

'************************************* �������� ����

'-------------------------------------------------��������� ������� ���� �������

ToDebug "Open DragDrop table"
Set rsdd = DBdd.OpenRecordset("Storage", dbOpenTable)
    If err.Number = 0 Then
        rsdd.Index = "key"
        
    ElseIf err.Number = 3011 Then
     ToDebug "Err: no table in DragDrop"
     myMsgBox msgsvc(20), vbCritical, , FrmMain.hwnd:   '�� ���� ������� ����.\�����? �� SVC ����? ���� ������� � MS Access?
     Exit Function
    Else 'error
     ToDebug err.Number & " - " & err.Description
     'ToDebug "���� ��� DragDrop ������? �������� ������?"
     myMsgBox msgsvc(26), vbCritical, , FrmMain.hwnd '� ����� �������� ������ ���������?\�������� ������ �� ����?
     Exit Function
    End If

OpenDBdd = True
End Function



Public Function GetPutJoin() As Boolean
'����������� ����� ����� ����
'�����������:
' �����
' ����� ������
' ����� ������

Dim tmpL As Long, tmp As String, tmpl2 As Long
Dim tmps As String
Dim tmpf As Long
Dim curKey As String    ' for RS
Dim curKeyJoin As String    ' - ������ ����� �������� ����
Dim CurLVKey As String    '��������������� LV
Dim NoSumTime As Boolean    '�� ����������� �����

'Dim i As Long

'��� ��������� GPDJ
On Error Resume Next

curKeyJoin = rs("Key")    '������� ����� �������� � ������ �������� ����������
rsJoin.MoveFirst
rsJoin.FindFirst "[Key] = " & curKeyJoin    'Val(ListView.SelectedItem.Key)
If rsJoin.NoMatch Then curKeyJoin = Val(CheckRowsKey(1))    '���� ������ ���, �� ��������

rs.AddNew
rsJoin.MoveFirst

NoSumTime = False

Do While Not rsJoin.EOF

    If rsJoin("Key") = curKeyJoin Then
        '������ ������ ���� ��������� (������ ����� ��� ����)
        rs.Fields("MovieName") = rsJoin.Fields("MovieName")
        rs.Fields("Label") = rsJoin.Fields("Label")
        rs.Fields("Genre") = rsJoin.Fields("Genre")
        rs.Fields("Year") = rsJoin.Fields("Year")
        rs.Fields("Country") = rsJoin.Fields("Country")
        rs.Fields("Director") = rsJoin.Fields("Director")
        rs.Fields("Acter") = rsJoin.Fields("Acter")
        rs.Fields("Time") = rsJoin.Fields("Time")    ' - ����� �����
        rs.Fields("Resolution") = rsJoin.Fields("Resolution")
        rs.Fields("Audio") = rsJoin.Fields("Audio")
        rs.Fields("FPS") = rsJoin.Fields("FPS")
        rs.Fields("FileLen") = Val(rsJoin.Fields("FileLen"))    ' - ����� �����
        rs.Fields("CDN") = rsJoin.Fields("CDN")
        rs.Fields("MediaType") = rsJoin.Fields("MediaType")
        rs.Fields("Video") = rsJoin.Fields("Video")
        rs.Fields("SubTitle") = rsJoin.Fields("SubTitle")
        rs.Fields("Language") = rsJoin.Fields("Language")
        rs.Fields("Rating") = rsJoin.Fields("Rating")
        rs.Fields("FileName") = rsJoin.Fields("FileName")    '- ����� �����
        rs.Fields("Debtor") = rsJoin.Fields("Debtor")
        rs.Fields("snDisk") = rsJoin.Fields("snDisk")
        rs.Fields("Other") = rsJoin.Fields("Other")
        rs.Fields("CoverPath") = rsJoin.Fields("CoverPath")
        rs.Fields("MovieURL") = rsJoin.Fields("MovieURL")
        rs.Fields("Annotation") = rsJoin.Fields("Annotation")
        'rs.Fields("Checked") = rsJoin.Fields("Checked") '�� ����, � �� ������

        rs.Fields("SnapShot1") = rsJoin.Fields("SnapShot1")
        rs.Fields("SnapShot2") = rsJoin.Fields("SnapShot2")
        rs.Fields("SnapShot3") = rsJoin.Fields("SnapShot3")
        rs.Fields("FrontFace") = rsJoin.Fields("FrontFace")
    End If

    '���� ����
    '����������� �����
    tmpl2 = Time2sec(rsJoin.Fields("Time"))
    If tmpl2 = -1 Or NoSumTime Then
        tmp = tmp & ", " & rsJoin.Fields("Time")
        NoSumTime = True
    Else
        tmpL = tmpL + tmpl2
    End If

    '��������� ����� ������
    If Len(rsJoin.Fields("FileName")) <> 0 Then
        If Len(tmps) = 0 Then
            tmps = rsJoin.Fields("FileName")
        Else
            tmps = tmps & " | " & rsJoin.Fields("FileName")
        End If
    End If
    '����������� ����� ������
    tmpf = tmpf + Val(rsJoin.Fields("FileLen"))

    '����� ��������� ����������

    rsJoin.MoveNext
Loop

'�����
If NoSumTime Then
    If Len(tmp) > 2 Then
        tmp = right$(tmp, Len(tmp) - 2)
        If tmpL > 0 Then tmp = tmp & ", " & FormatTime(tmpL)

        If rs.Fields("Time").Type = 10 Then
            rs.Fields("Time") = left$(tmp, 255)    'txt
        Else
            rs.Fields("Time") = tmp    'memo
        End If

    End If
Else
    rs.Fields("Time") = FormatTime(tmpL)
End If

'����� ������
If rs.Fields("FileName").Type = 10 Then
    rs.Fields("FileName") = left$(tmps, 255)    'txt
Else
    rs.Fields("FileName") = tmps    'memo
End If

'����� ������
rs.Fields("FileLen") = tmpf


curKey = rs("Key")    ' ������� ���� ����� ������ , �� �������
rs.Update
''''''''''''''''''''''''''''
'�������� � LV
With FrmMain
    RSGoto curKey
    .ListView.Sorted = False
    ReDim Preserve lvItemLoaded(.ListView.ListItems.Count + 1)    ' 1
    Add2LV .ListView.ListItems.Count, .ListView.ListItems.Count + 1    '2
    '��������

    CurLVKey = curKey & """"
    Set .ListView.SelectedItem = .ListView.ListItems(CurLVKey)
    CurSearch = .ListView.SelectedItem.Index
End With

'���� ���� ���������� - ���������� ��
If Opt_SortLVAfterEdit Then
    If LVSortColl > 0 Then LVSOrt (LVSortColl)
    If LVSortColl = -1 Then SortByCheck 0, True
End If

'''''''''''''''''''''''''''''''''''

If err.Number <> 0 Then
    ToDebug "Err GPDJ: " & err.Description
Else
    GetPutJoin = True
End If

err.Clear
End Function

Public Function GetGroupFieldName(ind As Integer) As String
'��������� ������ � �������, ���������� �������� ���� �� ����
If rs Is Nothing Then Exit Function

If ind > 0 Then GetGroupFieldName = rs(ind - 1).name
'If ind = -1 Then GetGroupFieldName = rs(dbCheckedInd).name

End Function
Public Function RSGoto(k As String) As Boolean
'������� � ������ ���� �� ����� "k"
'true - ���� ��������� �������

Dim kl As Long
kl = Val(k)

Select Case rs.Type
Case dbOpenTable
    rs.Seek "=", kl
    If rs.NoMatch Then
        Debug.Print "�� ������ ����"
        Exit Function
    End If
'Case dbOpenDynamic
'Case dbOpenSnapshot
Case dbOpenDynaset, dbOpenSnapshot
    rs.FindFirst "[Key] = " & kl
    If rs.NoMatch Then Debug.Print "�� ������ ����": Exit Function
'Case dbOpenForwardOnly
Case Else
    '�� �� �����, � �� �������
    '����� exit sub
   Debug.Print "�������� ��� ����": Exit Function
'    rs.MoveFirst
'    rs.Move ListView.SelectedItem.SubItems(lvIndexPole)
End Select

'��� � �����
RSGoto = True
End Function

Public Function RSGotoDD(kl As Long) As Boolean
'������� � ������ ���� �� ����� "kl"
'true - ���� ��������� �������

'Dim kl As Long
'kl = Val(k)

Select Case rsdd.Type
Case dbOpenTable
    rsdd.Seek "=", kl
    If rsdd.NoMatch Then
        Debug.Print "�� ������ ����"
        Exit Function
    End If
'Case dbOpenDynamic
'Case dbOpenSnapshot
Case dbOpenDynaset, dbOpenSnapshot
    rsdd.FindFirst "[Key] = " & kl
    If rsdd.NoMatch Then Debug.Print "�� ������ ����": Exit Function
'Case dbOpenForwardOnly
Case Else
    '�� �� �����, � �� �������
    '����� exit sub
   Debug.Print "�������� ��� ����": Exit Function
'    rs.MoveFirst
'    rs.Move ListView.SelectedItem.SubItems(lvIndexPole)
End Select

'��� � �����
RSGotoDD = True
End Function
Public Function RSGotoAct(k As String) As Boolean
'������� � ������ ���� �� ����� "k"
'true - ���� ��������� �������

Dim kl As Long
kl = Val(k)

Select Case ars.Type

Case dbOpenTable
    ars.Seek "=", kl
    If ars.NoMatch Then
        Debug.Print "�� ������ ����"
        Exit Function
    End If
    
Case dbOpenDynaset, dbOpenSnapshot
    ars.FindFirst "[Key] = " & kl
    If ars.NoMatch Then Debug.Print "�� ������ ����": Exit Function

'Case dbOpenForwardOnly
'Case dbOpenDynamic
'Case dbOpenSnapshot
Case Else
    '�� �� �����, � �� �������
    '����� exit
    '    ars.MoveFirst
    '    ars.Move ListView.SelectedItem.SubItems(lvIndexPole)
   Debug.Print "�������� ��� ����"
   Exit Function
End Select

'��� � �����
RSGotoAct = True
End Function
Public Sub JoinMovies(ch As Boolean)
If BaseReadOnly Or BaseReadOnlyU Then
'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
Exit Sub
End If

Dim ret As VbMsgBoxResult
Dim JoinSuccess As Boolean

'Dim LVSelected As String ' ����� ���������� ����� ������ LV - ( 10,11,)
Dim SQLstrGet As String, SQLstrDel As String

ToDebug "JoinM:" & CheckCount

If ch Then
    'checked
    If CheckCount < 2 Then Exit Sub    'If rsJoin.Fields.Count > 1 Then
    '��������� �� ����� �����������
    SQLstrGet = "SELECT * FROM Storage Where Checked = '1'"
    SQLstrDel = "DELETE FROM Storage Where Checked = '1'"

    ret = myMsgBox(msgsvc(41), vbYesNoCancel, , FrmMain.hwnd)
    Select Case ret
        Case vbNo    '������� join, �� �������
            Set rsJoin = DB.OpenRecordset(SQLstrGet)
            JoinSuccess = GetPutJoin
        Case vbYes    '������� join � ������� ������
            Set rsJoin = DB.OpenRecordset(SQLstrGet)
            JoinSuccess = GetPutJoin
            If JoinSuccess Then
                DB.Execute (SQLstrDel)
                DelLVItems True
            End If
        Case Else    'cancel
            ToDebug "...������"
            Exit Sub
    End Select

Else    'selected
'�� ����� �������
    If SelCount < 2 Then Exit Sub
    '�� DelMovies
End If

Set rsJoin = Nothing

End Sub

Public Sub LV_AllItemsCheck()
If rs Is Nothing Then Exit Sub
If rs.RecordCount < 1 Then Exit Sub
If BaseReadOnly Or BaseReadOnlyU Then
    'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
    Exit Sub
End If

rs.MoveFirst
Do While Not rs.EOF
    rs.Edit
    rs.Fields(dbCheckedInd) = "1"
    rs.Update
    rs.MoveNext
Loop

End Sub
Public Sub LV_AllItemsUnCheck()
If rs.RecordCount < 1 Then Exit Sub
If rs Is Nothing Then Exit Sub
If rs Is Nothing Then Exit Sub
If rs.RecordCount < 1 Then Exit Sub
If BaseReadOnly Or BaseReadOnlyU Then
    'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
    Exit Sub
End If

rs.MoveFirst
Do While Not rs.EOF
    rs.Edit
    rs.Fields(dbCheckedInd) = vbNullString
    rs.Update
    rs.MoveNext
Loop

End Sub

Public Sub NoBaseClear()
'������ frameview - � ������ ������� ������� ����
With FrmMain

    .Timer2.Enabled = False
    .Caption = "SurVideoCatalog"
    .ListView.ListItems.Clear
    .UCLV.Clear
    .Image0.Cls
    .TextVAnnot.Text = vbNullString
    .PicFaceV.Cls

End With
End Sub
Public Sub NoListClear()
With FrmMain

    .ListView.ListItems.Clear    'ListView.Sorted = False
    .TextItemHid.Text = vbNullString
    .TextVAnnot.Text = vbNullString
    .UCLV.Clear
    '? tvGroup.ListItems.Clear
    Set .PicFaceV = Nothing
    SelCount = 0
    CheckCount = 0

End With
End Sub

Public Function SearchSNinbase(ftext As String) As Boolean
'���� �� ����� �������� � ����
Dim strSQL As String
Dim rstemp As DAO.Recordset

On Error Resume Next

strSQL = "SELECT Label FROM Storage Where snDisk = '" & ftext & "'"
Set rstemp = DB.OpenRecordset(strSQL)
If rstemp.RecordCount > 0 Then
 rstemp.MoveFirst
 If Not IsNull(rstemp(0)) Then
        SearchSNinbase = True
        SameCDLabel = rstemp(0)
 End If
End If

Set rstemp = Nothing
End Function

Public Sub GetAFields()
'TextActName = CheckANoNull("Name")
'TextActBio = CheckANoNull("Bio")

With FrmMain
If ars.Fields("Name") <> vbNullString Then
    .TextActName.Text = ars.Fields("Name")
Else
    .TextActName.Text = vbNullString
End If

'Call SendMessage(TextActBio.hwnd, WM_SETREDRAW, False, ByVal 0&)

If ars.Fields("Bio") <> vbNullString Then
    .TextActBio.Text = ars.Fields("Bio")
Else
    .TextActBio.Text = vbNullString
End If

'Call SendMessage(TextActBio.hwnd, WM_SETREDRAW, True, ByVal 0&)
End With
End Sub
Public Sub PutAFields()
With FrmMain
If Len(.TextActName.Text) > 255 Then .TextActName.Text = left$(.TextActName.Text, 255)
ars.Fields("Name") = .TextActName.Text
ars.Fields("Bio") = .TextActBio.Text
End With
End Sub

Public Function PutActName(an As String) As Boolean

'�������� ������ � ������ an � ���� �������
Dim akey As String '���� �� ������� (�����)
On Error GoTo err

If BaseAReadOnly Then myMsgBox msgsvc(25), vbInformation, , FrmMain.hwnd: Exit Function
If Not LVActerFilled Then FrmMain.FillActListView

ars.AddNew
'� ����
If Len(an) > 255 Then an = left$(an, 255)
ars.Fields("Name") = an

akey = ars("Key") & """" '  1
ars.Update '                2

FrmMain.LVActer.Sorted = False
    '�������� � ALV
    LastIndAct = FrmMain.LVActer.ListItems.Count + 1
    FrmMain.LVActer.ListItems.Add(, akey, an).ListSubItems.Add 1, , LastIndAct
FrmMain.LVActer.Sorted = True
'ToActFromLV = LastIndAct 'Val(akey)
PutActName = True
Exit Function

err:
PutActName = False
ToDebug "Err_PutAN " & err.Description
End Function
