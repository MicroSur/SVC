Attribute VB_Name = "ModExcel"
Option Explicit
'Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Function IsExcelLoad() As Boolean
On Error Resume Next
Dim ExcelApp As Object
Set ExcelApp = GetObject(, "Excel.Application")
If err.Number <> 0 Then
    ToDebug "Err_NoExcel1:" & err.Description
    err.Clear
    IsExcelLoad = False
Else
    IsExcelLoad = True
    Set ExcelApp = Nothing
End If
End Function

Private Function GetExcelObject() As Object


If IsExcelLoad Then
    Set GetExcelObject = GetObject(, "Excel.Application")
Else
    On Error Resume Next
    Set GetExcelObject = CreateObject("Excel.Application")
    If err.Number <> 0 Then
        ToDebug "Err_NoExcel2:" & err.Description
        err.Clear
        Set GetExcelObject = Nothing
        Exit Function
    End If
End If
End Function

Function NewWorkbook(ExcelApp As Object) As Object
'� ����� ������� ����� ��������� ������ ���� ������� ����
ExcelApp.SheetsInNewWorkbook = 1

ExcelApp.Workbooks.Add
Set NewWorkbook = ExcelApp.ActiveWorkbook
End Function
Function OpenWorkbook(ExcelApp As Object, F As String) As Object
ExcelApp.Workbooks.Open filename:=F
Set OpenWorkbook = ExcelApp.ActiveWorkbook
End Function


'Public Sub ToExcel(objRst As Recordset)
'Dim ExcelApp As Object, ExcelWB As Object    'mzt , ExcelWS As Object
'Dim i As Long
'Dim j As Long
'Dim k As Integer
''Dim objRst As Recordset
'Dim vData As Variant
'Dim strFName As String
'
'Set ExcelApp = GetExcelObject()
'
'
'ExcelApp.Visible = False
'Set ExcelWB = NewWorkbook(ExcelApp)
''Set objRst = objRst.OpenRecordset("")
'If objRst.RecordCount <> 0 Then
'    objRst.MoveLast
'    objRst.MoveFirst
'    vData = objRst.GetRows(objRst.RecordCount)
'    If IsArray(vData) Then
'
'        For i = 0 To UBound(vData, 2)
'            k = 0
'            For j = 0 To 16    'UBound(vData, 1)
'                If j <> 5 Then
'                    k = k + 1
'                    strFName = ExcelApp.Columns(k).Address
'                    strFName = right$(strFName, Len(strFName) - (InStr(1, strFName, ":") + 1))
'                    'Debug.Print vData(j, i)
'
'                    If Not IsNull(vData(j, i)) Then
'                        ExcelApp.Sheets(1).Range(strFName & i + 1) = CStr(vData(j, i))
'                    Else
'                        ExcelApp.Sheets(1).Range(strFName & i + 1) = vbNullString
'                    End If
'
'                    'ExcelApp.columnWidth = 1000
'                    If j = 17 Then
'                        ExcelApp.Sheets(1).Range(strFName & i + 1).Font.Size = 8
'                        'Selection.Font.Bold = True
'                        ExcelApp.Sheets(1).Range(strFName & i + 1).Select
'                        ExcelApp.selection.ColumnWidth = 200
'                    End If
'                End If    '5
'            Next j
'        Next i
'        ExcelApp.Columns("A:AY").EntireColumn.AutoFit
'        ExcelApp.Columns("A:AY").EntireRow.AutoFit
'
'    End If    'IsArray(vData)
'    ExcelApp.Visible = True
'
'End If    'objRst.RecordCount <> 0
'
''?ExcelWB.guit
''?ExcelApp.quit
'
'End Sub

'Private Function LastCell(ws As Worksheet) As Range
'  Dim LastRow&, LastCol%
' Use : MsgBox LastCell(Sheet1).Row
' Error-handling is here in case there is not any
' data in the worksheet
'  On Error Resume Next
'  With ws
' Find the last real row
'    LastRow& = .Cells.Find(What:="*", _
     '      SearchDirection:=xlPrevious, _
     '      SearchOrder:=xlByRows).Row
' Find the last real column
'    LastCol% = .Cells.Find(What:="*", _
     '      SearchDirection:=xlPrevious, _
     '      SearchOrder:=xlByColumns).Column
'  End With
' Finally, initialize a Range object variable for
' the last populated row.
'  Set LastCell = ws.Cells(LastRow&, LastCol%)
'End Function

Public Sub ToExcelQuick(objRst As Recordset, PoleNames As String)
Dim ExcelApp As Object
'Dim ExcelApp As Excel.Application

Dim ExcelWB As Object
Dim strFName As String
Dim ar() As String '������ �������� �����
Dim i As Integer

On Error GoTo err
'Set ExcelApp = New Excel.Application
Set ExcelApp = GetExcelObject()



ExcelApp.Visible = True
'excel_sheet.Cells(row, 1) = rs!DataPoint

Set ExcelWB = NewWorkbook(ExcelApp)
'Set ExcelWB = OpenWorkbook(ExcelApp, "c:\catalogs\test.xls")

'��������� ������ ������ ��� ��������
ExcelApp.ActiveWindow.SplitRow = 1
ExcelApp.ActiveWindow.FreezePanes = True

'ExcelApp.Selection.columnWidth = 20
'ExcelApp.Sheets(1).Range.ColumnWidth = 30

ar = Split(PoleNames, ",")

'�������� �����
For i = 0 To UBound(ar)
strFName = ExcelApp.Columns(i + 1).Address '$A:$A
strFName = right$(strFName, Len(strFName) - (InStr(1, strFName, ":") + 1)) ' A B ...

ExcelApp.Sheets(1).Range(strFName & 1).ColumnWidth = 20
ExcelApp.Sheets(1).Range(strFName & 1).Font.Bold = True
ExcelApp.Sheets(1).Range(strFName & 1).Interior.color = &HC0FFFF

'������ ����� <, > ">"
If right$(ar(i), 2) = " >" Or right$(ar(i), 2) = " <" Then ar(i) = left$(ar(i), Len(ar(i)) - 2)

ExcelApp.Sheets(1).Range(strFName & 1).Value = ar(i)
Next i

'����� ���������

'
ExcelApp.Sheets(1).cells.VerticalAlignment = 1 '�� �������� ����
'ExcelApp.Sheets(1).Cells.HorizontalAlignment = 5 '� �����������
'ExcelApp.Sheets(1).cells.WrapText = True
'ExcelApp.Columns(strFName).EntireColumn.AutoFit
'ExcelApp.Columns(strFName).EntireRow.AutoFit

''��������� ������ ������ ��������� - ��������� ������
'If LstExport_Arr(dbAnnotationInd) Then
''ExcelApp.Sheets(1).Range(strFName & 1).ColumnWidth = 100
'ExcelApp.Sheets(1).Range(strFName & 1).ColumnWidth = 100
'End If

'������
'If objRst.RecordCount <> 0 Then
'ExcelApp.Sheets(1).Range("A2").CopyFromRecordset objRst '911 �������� �� ������, ����� ������
'End If
'�������� ��������
With ExcelApp.Sheets(1).QueryTables.Add(Connection:=objRst, _
Destination:=ExcelApp.Sheets(1).Range("A2"))
''.CommandText = oSQL
.name = "qSurVideoCatalog"
.FieldNames = False
.RowNumbers = False
.FillAdjacentFormulas = False
'.PreserveFormatting = True
'.RefreshOnFileOpen = False
.BackgroundQuery = True
.RefreshStyle = 0 'xlInsertDeleteCells
.SavePassword = False
.SaveData = True
.AdjustColumnWidth = False 'True
.RefreshPeriod = 0
.PreserveColumnInfo = True
.Refresh BackgroundQuery:=False
.EnableRefresh = False 'this is NOT equiv to unchecking "Save query definition."

End With
'delete the query table. this is equivalent to unchecking "Save query definition."
ExcelApp.Sheets(1).QueryTables.Item("qSurVideoCatalog").Delete



'ExcelApp.Columns("A:AY").EntireColumn.AutoFit
'ExcelApp.Columns("A:AY").EntireRow.AutoFit
'ExcelWB.guit
'ExcelApp.quit

Set ExcelWB = Nothing
Set ExcelApp = Nothing

Exit Sub
err:
Set ExcelWB = Nothing
Set ExcelApp = Nothing
Debug.Print "Err_QuickExcel:" & err.Description

End Sub

Public Sub Export2Excel(ch As Boolean)
Dim LVSelected As String    ' ����� ���������� ����� ������ LV - ( 10,11,)
Dim statement As String
'Dim tmp As String
Dim expFields As String    ' MovieName,Label,Genre - �� ����
Dim rsExcel As DAO.Recordset
Dim j As Long, i As Long
Dim pArr() As Integer
Dim localNames As String    '��������,�����,����
Dim ProcessAnnotFlag As Boolean
Dim IP As Integer    '���-�� ����� �������� � �������
Dim tmp As String

On Error Resume Next

'��������� ������ ������������   ������ � ������ - �������
'������ ������� - �������, �������� - ������ ���� ������
ReDim pArr(FrmMain.ListView.ColumnHeaders.Count)
For j = 0 To LstExport_ListCount    '   '��� 0-24, ����� ������ �� ����������
    pArr(FrmMain.ListView.ColumnHeaders(j + 1).Position) = j
Next j

'����
For j = 1 To UBound(pArr)    '1-25 ������ �� ����������
    If LstExport_Arr(pArr(j)) Then
        If pArr(j) <> dbAnnotationInd Then
            expFields = expFields & "," & rs.Fields(pArr(j)).name
            IP = IP + 1
            '�������
            localNames = localNames & "," & TranslatedFieldsNames(pArr(j))
        End If
    End If
Next j

' + ���������
If LstExport_Arr(dbAnnotationInd) Then
    expFields = expFields & "," & rs.Fields(dbAnnotationInd).name
    localNames = localNames & "," & TranslatedFieldsNames(dbAnnotationInd)
    IP = IP + 1
    ProcessAnnotFlag = True
End If


If Len(expFields) = 0 Then
    '�� ������� ���� ��������
    ToDebug "No Exp2Exl Fields"
    Exit Sub
Else
    expFields = right$(expFields, Len(expFields) - 1)    '- ����� ,
    localNames = right$(localNames, Len(localNames) - 1)
End If

'expFields = "Annotation"

If ch Then
    'checked
    statement = "SELECT " & expFields & " FROM Storage Where Checked = '1'"
Else
    'selected
    LVSelected = vbNullString
    Dim LVArr() As String    ' ������ ������ lv - SelRowsKey -1 � ��� "
    ReDim LVArr(UBound(SelRowsKey) - 1)
    For i = 0 To UBound(LVArr)
        LVArr(i) = Val(SelRowsKey(i + 1))
    Next i
    LVSelected = Join(LVArr, ",")

    statement = "SELECT " & expFields & " FROM Storage WHERE Key IN (" & LVSelected & ")"
End If


On Error GoTo err

'������ ����� �� ��������
    Set rsExcel = DB.OpenRecordset(statement)
    If Not (rsExcel.BOF And rsExcel.EOF) Then
        rsExcel.MoveLast: rsExcel.MoveFirst
        
        If ProcessAnnotFlag Then
        For i = 1 To rsExcel.RecordCount
            tmp = CheckNoNullValMyRS(IP - 1, rsExcel)
            If InStr(tmp, vbCrLf) Then
                rsExcel.Edit
                rsExcel(IP - 1) = Replace(tmp, vbCrLf, "")
                rsExcel.Update
                rsExcel.MoveNext
            End If
        Next i
        rsExcel.MoveFirst
        End If
    End If
    
ToExcelQuick rsExcel, localNames

ToDebug "Exp2Ex_done"

err:
Set rsExcel = Nothing
If err Then
    MsgBox msgsvc(46), vbExclamation
    ToDebug "Err_Exp2Ex:" & err.Description
End If

End Sub
Public Sub Export2HTML_ABCD(Marker As Boolean)
'If Marker Then   ���������� ����� ���������� v
Dim FooterHTML As String
Dim BodyHTML As String
Dim HeaderHTML As String
Dim s As String
Dim tmpArr() As String    '����� ���� �����
Dim ABCDLine() As String    '����� 1 - ������ ������ �� ������ �������
Dim doFlagSS1 As Boolean, doFlagSS2 As Boolean, doFlagSS3 As Boolean, doFlagSS0 As Boolean
Dim strPath As String

Dim htmlDir As String
Dim CoverDir As String '��� ����
Dim SShotDir As String
Dim o_CoverDir As String '��� ����
Dim o_SShotDir As String

Dim htmlFile As String
Dim htmlFileFirst As String    '��� ������� �����

Dim iFile As Integer    '����� ����� html
Dim j As Long, i As Long    ', k As Long
Dim temp As String    ', temp2 As String

Dim rsRecordCount As Integer    '���-�� ����� ��� ������
Dim AllLinkStr() As String    '������� ������ �������� � �������, ���� ����
Dim AllCharStr() As String    '������� ������ ��������
Dim strAllLink As String    '��� ������ ������
Dim rsTmp As DAO.Recordset
Dim sSQL As String
Dim arrPagesCount() As Long    '������� ������ ����� ���� � ����
Dim ABCLinks() As String    'tokenize 1 ��������
Dim ubABC As Long    'ubound ABCLinks()
Dim crFlag As Boolean    '���� �� ���������� ������ (\)

On Error GoTo err
'On Error GoTo 0

'��������� ������ �����
iFile = FreeFile
Open App.Path & "\Templates\" & CurrentHtmlTemplate For Binary As #iFile
's = Space$(LOF(iFile))
s = AllocString_ADV(LOF(iFile))
Get #iFile, , s
Close #iFile

'� ����� �����?
temp = BrowseForFolderByPath(FixPath(lastHTMLfolderPath), NamesStore(4))
If Len(temp) = 0 Then FrmMain.Timer2.Enabled = True: Exit Sub
lastHTMLfolderPath = temp '�������� ��� �������

On Error Resume Next
'�����
If Opt_ExpUseFolders Then
    If Len(Opt_ExpFolder1) <> 0 Then
        ReplaceFNStr Opt_ExpFolder1
        htmlDir = lastHTMLfolderPath & "\" & Opt_ExpFolder1
        MkDir htmlDir
    Else
        htmlDir = lastHTMLfolderPath
    End If
    If Len(Opt_ExpFolder2) <> 0 Then
        ReplaceFNStr Opt_ExpFolder2
        CoverDir = lastHTMLfolderPath & "\" & Opt_ExpFolder2 & "\"
        MkDir CoverDir
        If Len(Opt_ExpFolder1) = 0 Then o_CoverDir = Opt_ExpFolder2 & "/" Else o_CoverDir = "../" & Opt_ExpFolder2 & "/"
    Else
        CoverDir = lastHTMLfolderPath & "\"
        If Len(Opt_ExpFolder1) = 0 Then o_CoverDir = vbNullString Else o_CoverDir = "../"
    End If
    If Len(Opt_ExpFolder3) <> 0 Then
        ReplaceFNStr Opt_ExpFolder3
        SShotDir = lastHTMLfolderPath & "\" & Opt_ExpFolder3 & "\"
        MkDir SShotDir
        If Len(Opt_ExpFolder1) = 0 Then o_SShotDir = Opt_ExpFolder3 & "/" Else o_SShotDir = "../" & Opt_ExpFolder3 & "/"
    Else
        SShotDir = lastHTMLfolderPath & "\"
        If Len(Opt_ExpFolder1) = 0 Then o_SShotDir = vbNullString Else o_SShotDir = "../"
    End If
Else
    '� ����
    htmlDir = lastHTMLfolderPath
    CoverDir = lastHTMLfolderPath & "\"
    SShotDir = lastHTMLfolderPath & "\"
    o_CoverDir = vbNullString
    o_SShotDir = vbNullString
End If
'If err Then Debug.Print "��"
err.Clear
On Error GoTo err

'����������� background � nopicture, �����
temp = App.Path & "\Templates\nopicture.jpg"
If FileExists(temp) Then FileCopy temp, htmlDir & "/nopicture.jpg"
temp = App.Path & "\Templates\background.jpg"
If FileExists(temp) Then FileCopy temp, htmlDir & "/background.jpg"
temp = App.Path & "\Templates\styles.css"
If FileExists(temp) Then FileCopy temp, htmlDir & "/styles.css"

'If Marker Then    '����������   ' ��� �� ����� �� �� �������
'    rsRecordCount = SelCount
'Else    '          ���������� (v)
'    rsRecordCount = CheckCount
'End If
'������ ����� ������
's = Replace(s, "$TOTAL$", rsRecordCount, , , vbTextCompare)
s = Replace(s, "$OWNER$", GetPCUserName, , , vbTextCompare)
s = Replace(s, "$DATE$", Date, , , vbTextCompare)
s = Replace(s, "$TIME$", Format$(Time, "hh:mm"), , , vbTextCompare)
s = Replace(s, "$SVCBASENAME$", FrmMain.Caption, , , vbTextCompare)

'����� ��� ��������
If InStrB(s, "$SNAPSHOT1$") > 0 Then doFlagSS1 = True
If InStrB(s, "$SNAPSHOT2$") > 0 Then doFlagSS2 = True
If InStrB(s, "$SNAPSHOT3$") > 0 Then doFlagSS3 = True
If InStrB(s, "$COVER$") > 0 Then doFlagSS0 = True

'����� �� �����, ���� � �����
tmpArr() = Split(s, "$SVC_BODY$")
If UBound(tmpArr) <> 2 Then myMsgBox msgsvc(35), vbCritical, , FrmMain.hwnd: Screen.MousePointer = vbNormal: Exit Sub    ' �� ��� ��������

'�� ������ ������ ������ ������ ABCDLine(1)
ABCDLine() = Split(tmpArr(0), "$SVC_ABCD_LINE$")
If UBound(ABCDLine) <> 2 Then myMsgBox msgsvc(35), vbCritical, , FrmMain.hwnd: Screen.MousePointer = vbNormal: Exit Sub    ' �� ��� ��������

'����� ���� �� ����� �� ������� (����� ���� Z\ - ���������� ������ ����� Z)

ubABC = Tokenize04(ABCDLine(1), ABCLinks(), " ", False)
If ubABC > -1 Then
    ReDim AllLinkStr(ubABC): ReDim arrPagesCount(ubABC): ReDim AllCharStr(ubABC)
Else
    ToDebug "Exp2ABC: no links"
    'Debug.Print "��� ������"
End If

DoEvents
Screen.MousePointer = vbHourglass

'��������
FrmMain.PBar.Max = ubABC: FrmMain.PBar.Value = 0: FrmMain.PBar.ZOrder 0
'��������� AllLinkStr �������, ��������� � ��������� ���-��� arrPagesCount ������ �����
For i = 0 To ubABC
    crFlag = False
    If InStr(ABCLinks(i), "\") > 0 Then
        crFlag = True        '����� ����� ������� ������
    End If

    If ABCLinks(i) <> "[0-9]" Then
        AllCharStr(i) = Chr$(Asc(ABCLinks(i)))    '�������� 1 ���
    Else
        AllCharStr(i) = ABCLinks(i)    '"[0-9]"
    End If

    If Marker Then    '����������

        '���������� ������ ���������� ������
        Dim LVArr() As String
        Dim LVSelected As String
        ReDim LVArr(UBound(SelRowsKey) - 1)
        For j = 0 To UBound(LVArr)
            LVArr(j) = Val(SelRowsKey(j + 1))
        Next j
        LVSelected = Join(LVArr, ",")
        Erase LVArr

        sSQL = "SELECT COUNT(MovieName) FROM STORAGE WHERE ((MovieName Like '" & AllCharStr(i) & "*') AND (Key IN (" & LVSelected & ")))"

    Else    '          ���������� (v)

        sSQL = "SELECT COUNT(MovieName) FROM STORAGE WHERE ((MovieName Like '" & AllCharStr(i) & "*') AND (Checked = '1'))"
    End If
    Set rsTmp = DB.OpenRecordset(sSQL)
    '��������� ������ ���-��� ��������� � ���� � ������ �����������/�����������
    If Not (rsTmp.BOF And rsTmp.EOF) Then arrPagesCount(i) = rsTmp(0)

    '������� ���� ���� ���� ������
    If arrPagesCount(i) > 0 Then
        AllLinkStr(i) = "<a href='svc_" & AllCharStr(i) & ".htm'>" & AllCharStr(i) & "</a>"
    Else
        AllLinkStr(i) = AllCharStr(i)
    End If
    If crFlag Then AllLinkStr(i) = AllLinkStr(i) & "<br>"        '������� ������

    FrmMain.PBar.Value = i    '������� ��������
Next i
Erase ABCLinks: Erase ABCDLine
For i = 0 To ubABC
    strAllLink = strAllLink & "&nbsp;" & AllLinkStr(i)    '��� ������ ������
    rsRecordCount = rsRecordCount + arrPagesCount(i)    '� ������� ������� �������� ��� ��������� �����
Next i

HeaderHTML = tmpArr(0): FooterHTML = tmpArr(2)
HeaderHTML = Replace(HeaderHTML, "$TOTAL$", rsRecordCount, , , vbTextCompare)    '����� ����� �������
FooterHTML = Replace(FooterHTML, "$TOTAL$", rsRecordCount, , , vbTextCompare)

HeaderHTML = Replace(HeaderHTML, "$PAGELINE$", strAllLink, , , vbTextCompare)
FooterHTML = Replace(FooterHTML, "$PAGELINE$", strAllLink, , , vbTextCompare)

'���������� � ������������ ����� (������ � �������� ���� ������)
For i = 0 To ubABC
    If arrPagesCount(i) > 0 Then
        htmlFileFirst = htmlDir & "\svc_" & AllCharStr(i) & ".htm"
        Exit For
    End If
Next i

'��������
FrmMain.PBar.Max = ubABC: FrmMain.PBar.Value = 0: FrmMain.PBar.ZOrder 0

'����� �������
For i = 0 To ubABC
    If arrPagesCount(i) > 0 Then

        If Marker Then    '����������
            sSQL = "SELECT * FROM STORAGE WHERE ((MovieName Like '" & AllCharStr(i) & "*') AND (Key IN (" & LVSelected & "))) Order by MovieName Asc"
        Else    '          ���������� (v)
            sSQL = "SELECT * FROM STORAGE WHERE ((MovieName Like '" & AllCharStr(i) & "*') AND (Checked = '1')) Order by MovieName Asc"
        End If
        'sSQL = "SELECT * FROM STORAGE WHERE MovieName Like '" & AllCharStr(i) & "*' Order by MovieName Asc"
        'Debug.Print sSQL

        Set rsTmp = DB.OpenRecordset(sSQL)

        If Not (rsTmp.BOF And rsTmp.EOF) Then
            rsTmp.MoveLast: rsTmp.MoveFirst

            iFile = FreeFile    '������� � ��������� ����
            htmlFile = htmlDir & "\svc_" & AllCharStr(i) & ".htm"   '���������� ���� ��� ������ �����

            Open htmlFile For Output As #iFile
            '�����
            Print #iFile, Replace(HeaderHTML, "$NUMBER$", arrPagesCount(i), , , vbTextCompare)

            'body
            For j = 1 To arrPagesCount(i)
                BodyHTML = tmpArr(1)
                Call Export2HTML_BODY(BodyHTML, j, False, CoverDir, SShotDir, o_CoverDir, o_SShotDir, rsTmp)
                Print #iFile, BodyHTML
                rsTmp.MoveNext
            Next j

            '�����
            Print #iFile, Replace(FooterHTML, "$NUMBER$", arrPagesCount(i), , , vbTextCompare)
            Close #iFile

        End If    'If Not (rsTMP.BOF And rsTMP.EOF) Then
    End If    'If arrPagesCount(i) <> 0

    FrmMain.PBar.Value = i    '������� ��������
Next i

Screen.MousePointer = vbNormal
FrmMain.TextItemHid.ZOrder 0
'�� ������� RestoreBasePos

'������� ��������
If Len(htmlFileFirst) > 0 Then
    strPath = Space$(255)
    temp = FindExecutable(htmlFileFirst, "", strPath)
    Select Case temp
    Case 31
        myMsgBox msgsvc(17), vbInformation, , FrmMain.hwnd
        Screen.MousePointer = vbNormal
        Exit Sub
    End Select
    temp = ShellExecute(GetDesktopWindow(), "open", htmlFileFirst, vbNull, vbNull, 1)
End If

Set rsTmp = Nothing
ToDebug "ABCHTM_done"
Exit Sub

err:
On Error Resume Next
Debug.Print "Err_ABCHTML"
ToDebug "Error_ABCHTML"
Resume Next
End Sub
Public Sub Export2HTML(Marker As Boolean)
'��� �������
' + �������� ��������� ������ JSFlag
' + ������������ �����������������, ������ ������ 1
' + �������� �������� �������� ������� �������� � ���� ���� svc1.htm (svc_����������)
' + � ������ ������ �� ������ ���� " '
' + ������ " �� \"
' - ����� ��� ����������� �������?
' �� ������ �������� err

If rs Is Nothing Then Exit Sub

'�������
Dim FooterHTML As String
Dim BodyHTML As String
Dim HeaderHTML As String
Dim s As String
Dim tmpArr() As String
Dim jpgHtmlArr(3) As String    '�� nm (0123) ������ ����� �������� ��� html
Dim Skey As String    '����� � ����������� �� j
Dim JSFlag As Boolean    '���� �������� � � html � JS

'''
Dim strPath As String

Dim htmlDir As String
Dim CoverDir As String '��� ����
Dim SShotDir As String
Dim o_CoverDir As String '��� ����
Dim o_SShotDir As String

Dim htmlFile As String
Dim iFile As Integer    '����� ����� html
Dim JPGFile As String
Dim ret As Long
Dim ubnd As Long
Dim j As Integer, i As Integer, M As Long
Dim temp As String    ', temp2 As String
Dim doflag As Boolean

Dim lPtr As Long
Dim PicSize As Long
Dim b() As Byte
Dim jFile As String    ' ��� ����� jpeg
Dim LFile As Integer    '��������� ���� ��� ��������
Dim img As ImageFile
Dim vec As Vector
Dim PicExt As String

'��������������
Dim PagesCount As Integer
Dim RecsOnPage As Integer
Dim rcount As Integer    ' ������� �������
Dim htmlFileFirst As String    '��� ������� �����
Dim StartStr As String
Dim Ostatok As Integer    '������� �� ���� ��������
Dim TechCount As Integer    '������ ������ ����� ��� ������� ��������. ������ � 1
Dim rsRecordCount As Integer    '���-�� ����� ��� ������
Dim AllLinkStr() As String    '������ ������ �� ��. ��������
Dim LinkStr() As String    '���� ����
'''
Dim nm As Integer    '�������� �������� � ����� ����� ��������
'Dim strFoldPath As String, strFoldName As String
Dim anno As String    '�������� �� <br>
Dim tmpField As String    '�������� ���� ����


On Error GoTo ErrHandler
'��� ������� �� ������...
If Len(CurrentHtmlTemplate) = 0 Then myMsgBox msgsvc(34), vbCritical, , FrmMain.hwnd: Exit Sub    '���������� ������
FrmMain.Timer2.Enabled = False

'��������� ������ (�� ���. ������ ������-������ �������)
iFile = FreeFile
Open App.Path & "\Templates\" & CurrentHtmlTemplate For Binary As #iFile
's = Space$(LOF(iFile))
s = AllocString_ADV(LOF(iFile))
Get #iFile, , s
Close #iFile

'���� ����-������ - ������� � ������ ������������
If InStr(1, s, "$SVC.ABCD$", vbTextCompare) Then
    On Error GoTo 0
    Call Export2HTML_ABCD(Marker)
    Exit Sub
End If
'�������� ������, ����� �� ������ ����� ��� ������������� � JS
If InStr(1, s, "$SVC.JS.Array$", vbTextCompare) Then JSFlag = True

'� ����� �����?
temp = BrowseForFolderByPath(FixPath(lastHTMLfolderPath), NamesStore(4))
If Len(temp) = 0 Then FrmMain.Timer2.Enabled = True: Exit Sub
lastHTMLfolderPath = temp

On Error Resume Next
'����� '� ��� ������� ��� ������� ��� ���� ����� ����� /
If Opt_ExpUseFolders Then
    If Len(Opt_ExpFolder1) <> 0 Then
        ReplaceFNStr Opt_ExpFolder1
        htmlDir = lastHTMLfolderPath & "\" & Opt_ExpFolder1
        MkDir htmlDir
    Else
        htmlDir = lastHTMLfolderPath
    End If
    If Len(Opt_ExpFolder2) <> 0 Then
        ReplaceFNStr Opt_ExpFolder2
        CoverDir = lastHTMLfolderPath & "\" & Opt_ExpFolder2 & "\"
        MkDir CoverDir
        If Len(Opt_ExpFolder1) = 0 Then o_CoverDir = Opt_ExpFolder2 & "/" Else o_CoverDir = "../" & Opt_ExpFolder2 & "/"
    Else
        CoverDir = lastHTMLfolderPath & "\"
        If Len(Opt_ExpFolder1) = 0 Then o_CoverDir = vbNullString Else o_CoverDir = "../"
    End If
    If Len(Opt_ExpFolder3) <> 0 Then
        ReplaceFNStr Opt_ExpFolder3
        SShotDir = lastHTMLfolderPath & "\" & Opt_ExpFolder3 & "\"
        MkDir SShotDir
        If Len(Opt_ExpFolder1) = 0 Then o_SShotDir = Opt_ExpFolder3 & "/" Else o_SShotDir = "../" & Opt_ExpFolder3 & "/"
    Else
        SShotDir = lastHTMLfolderPath & "\"
        If Len(Opt_ExpFolder1) = 0 Then o_SShotDir = vbNullString Else o_SShotDir = "../"
    End If
Else
    '� ����
    htmlDir = lastHTMLfolderPath
    CoverDir = lastHTMLfolderPath & "\"
    SShotDir = lastHTMLfolderPath & "\"
    o_CoverDir = vbNullString
    o_SShotDir = vbNullString
End If
'If err Then Debug.Print "��"
err.Clear
On Error GoTo ErrHandler


If Not JSFlag Then
    '������ ���� �� �����
    '������� �� ���������
    If TxtNnOnPage_Text > 0 Then
        RecsOnPage = Val(TxtNnOnPage_Text)
    Else
        RecsOnPage = 30
    End If

    If Marker Then    '����������
        PagesCount = UBound(SelRows) / RecsOnPage
        Ostatok = UBound(SelRows) - PagesCount * RecsOnPage
        rsRecordCount = UBound(SelRows)
    Else    '          ���������� (v)
        PagesCount = UBound(CheckRows) / RecsOnPage
        Ostatok = UBound(CheckRows) - PagesCount * RecsOnPage
        rsRecordCount = UBound(CheckRows)
    End If

    If Ostatok > 0 Then PagesCount = PagesCount + 1

    '���������� ������ � �������� �������
    ReDim LinkStr(PagesCount)
    ReDim AllLinkStr(PagesCount)

    For i = 1 To PagesCount
        htmlFile = "svc" & i & ".htm"    '�������������, ��� ������
        LinkStr(i) = "<a href=""" & htmlFile & """><b>" & i & "</b></a>"
    Next i
    For i = 1 To PagesCount
        For j = 1 To PagesCount
            If j <> i Then
                AllLinkStr(i) = AllLinkStr(i) & LinkStr(j) & ", "
            Else
                AllLinkStr(i) = AllLinkStr(i) & j & ", "
            End If
        Next j
        AllLinkStr(i) = left$(AllLinkStr(i), Len(AllLinkStr(i)) - 2)
        StartStr = i * RecsOnPage - RecsOnPage + 1

        If (i = PagesCount) And (Ostatok > 0) Then AllLinkStr(i) = AllLinkStr(i) & "<br>" & StartStr & "-" & StartStr + Ostatok - 1 & " / " & rsRecordCount
        If (i = PagesCount) And (Ostatok <= 0) Then AllLinkStr(i) = AllLinkStr(i) & "<br>" & StartStr & "-" & StartStr + Ostatok + RecsOnPage - 1 & " / " & rsRecordCount
        If (i <> PagesCount) Then AllLinkStr(i) = AllLinkStr(i) & "<br>" & StartStr & "-" & StartStr + RecsOnPage - 1 & " / " & rsRecordCount

    Next i

Else
    'Java, ���� ��������
    PagesCount = 1
    '���������� ������ � �������� �������
    ReDim LinkStr(PagesCount)
    ReDim AllLinkStr(PagesCount)
    If Marker Then    '����������
        RecsOnPage = SelCount
    Else    '          ���������� (v)
        RecsOnPage = CheckCount
    End If

End If    'jsflag

'����������� background � nopicture, �����
temp = App.Path & "\Templates\nopicture.jpg"
If FileExists(temp) Then FileCopy temp, htmlDir & "/nopicture.jpg"
temp = App.Path & "\Templates\background.jpg"
If FileExists(temp) Then FileCopy temp, htmlDir & "/background.jpg"
temp = App.Path & "\Templates\styles.css"
If FileExists(temp) Then FileCopy temp, htmlDir & "/styles.css"

DoEvents
Screen.MousePointer = vbHourglass

rcount = 0: TechCount = 1

'���������� � ������������ �����
If PagesCount = 1 Then
    htmlFileFirst = htmlDir & "/index.htm"
Else
    htmlFileFirst = htmlDir & "/svc1.htm"
End If

'������ ����� ������
s = Replace(s, "$TOTAL$", rsRecordCount, , , vbTextCompare)
s = Replace(s, "$OWNER$", GetPCUserName, , , vbTextCompare)
s = Replace(s, "$DATE$", Date, , , vbTextCompare)
s = Replace(s, "$TIME$", Format$(Time, "hh:mm"), , , vbTextCompare)
s = Replace(s, "$SVCBASENAME$", FrmMain.Caption, , , vbTextCompare)


    '����� �� �����, ���� � �����
    tmpArr() = Split(s, "$SVC_BODY$")
    If UBound(tmpArr) <> 2 Then myMsgBox msgsvc(35), vbCritical, , FrmMain.hwnd: Screen.MousePointer = vbNormal: Exit Sub    ' �� ��� ��������

    
'                                                                   ������������ ����

For i = 1 To PagesCount
    '������ i ������
    'sTmp = s
    HeaderHTML = tmpArr(0): FooterHTML = tmpArr(2)
    If PagesCount > 1 Then
        HeaderHTML = Replace(HeaderHTML, "$PAGELINE$", AllLinkStr(i), , , vbTextCompare)
        FooterHTML = Replace(FooterHTML, "$PAGELINE$", AllLinkStr(i), , , vbTextCompare)

    Else
        HeaderHTML = Replace(HeaderHTML, "$PAGELINE$", vbNullString, , , vbTextCompare)
        FooterHTML = Replace(FooterHTML, "$PAGELINE$", vbNullString, , , vbTextCompare)
    End If
    HeaderHTML = Replace(HeaderHTML, "$PAGENUMBER$", i, , , vbTextCompare)
    FooterHTML = Replace(FooterHTML, "$PAGENUMBER$", i, , , vbTextCompare)

    htmlFile = htmlDir & "/svc" & i & ".htm"    '���������� ��� �����

    '������� ����� html
    iFile = FreeFile
    Open htmlFile For Output As #iFile
    '�����
    Print #iFile, HeaderHTML

    '����� �� ������ ��������
    If Marker Then
        ubnd = UBound(SelRows)
    Else
        ubnd = UBound(CheckRows)
    End If

    FrmMain.PBar.ZOrder 0: FrmMain.PBar.Max = ubnd

    '                                                                   ���� �� ���-�� �� ��������
    For M = TechCount To ubnd
        FrmMain.PBar.Value = M

        If Marker Then
            If MultiSel Then
                RSGoto SelRowsKey(M)
            End If
        Else
            RSGoto CheckRowsKey(M)
        End If

        '�� ������ �� ����. ��������
        If rcount = RecsOnPage Then rcount = 0: TechCount = M: Exit For

        BodyHTML = tmpArr(1)
        
        Call Export2HTML_BODY(BodyHTML, M, JSFlag, CoverDir, SShotDir, o_CoverDir, o_SShotDir, rs)

        Print #iFile, BodyHTML

        rcount = rcount + 1
    Next M    'm = TechCount To UBound(...Rows)

    Print #iFile, FooterHTML
    Close #iFile

Next i    'For i = 1 To PagesCount

FrmMain.TextItemHid.ZOrder 0
RestoreBasePos
Screen.MousePointer = vbNormal

'������� ��������
If Len(htmlFileFirst) > 0 Then
strPath = Space$(255)
temp = FindExecutable(htmlFileFirst, "", strPath)
Select Case temp
Case 31
    myMsgBox msgsvc(17), vbInformation, , FrmMain.hwnd
    Screen.MousePointer = vbNormal
    Exit Sub
End Select
temp = ShellExecute(GetDesktopWindow(), "open", htmlFileFirst, vbNull, vbNull, 1)
End If

ToDebug "Exp2HTML_done"
Exit Sub

ErrHandler:
Close #iFile
Close #LFile
'Debug.Print "exp2html " & err.Description
ToDebug "Error_Exp2HTML"
Screen.MousePointer = vbNormal
End Sub


Public Sub Export2HTML_BODY(ByRef BHTML, ByRef CurNum As Long, ByRef JSFlg As Boolean, ByRef cov_Dir As String, ByRef ss_Dir As String, ByRef o_cov_Dir As String, ByRef o_ss_Dir As String, ByRef rs2proc As DAO.Recordset)
Dim jpgHtmlArr(3) As String    '�� nm (0123) ������ ����� �������� ��� html
Dim j As Integer
Dim doflag As Boolean
Dim doFlagSS1 As Boolean, doFlagSS2 As Boolean, doFlagSS3 As Boolean, doFlagSS0 As Boolean
Dim a() As String    'tokenize
Dim jFile As String    ' ��� ����� jpeg
Dim tmpField As String
Dim PicSize As Long
Dim b() As Byte
Dim LFile As Integer    '��������� ���� ��� ��������
Dim img As ImageFile
Dim vec As Vector
Dim PicExt As String
Dim nm As Integer    '�������� �������� � ����� ����� ��������
Dim JPGFile As String
Dim Skey As String    '����� � ����������� �� j
Dim picDir As String '�����
Dim o_picDir As String

On Error GoTo err '��� img

'����� ��� ��������
If InStrB(BHTML, "$SNAPSHOT1$") > 0 Then doFlagSS1 = True
If InStrB(BHTML, "$SNAPSHOT2$") > 0 Then doFlagSS2 = True
If InStrB(BHTML, "$SNAPSHOT3$") > 0 Then doFlagSS3 = True
If InStrB(BHTML, "$COVER$") > 0 Then doFlagSS0 = True

'Erase jpgHtmlArr
For j = dbSnapShot1Ind To dbFrontFaceInd
    doflag = False
    If (j = dbSnapShot1Ind) And doFlagSS1 Then doflag = True: nm = 1: picDir = ss_Dir: o_picDir = o_ss_Dir
    If (j = dbSnapShot2Ind) And doFlagSS2 Then doflag = True: nm = 2: picDir = ss_Dir: o_picDir = o_ss_Dir
    If (j = dbSnapShot3Ind) And doFlagSS3 Then doflag = True: nm = 3: picDir = ss_Dir: o_picDir = o_ss_Dir
    If (j = dbFrontFaceInd) And doFlagSS0 Then doflag = True: nm = 0: picDir = cov_Dir: o_picDir = o_cov_Dir

    If doflag Then
        '����� jpeg ������
        jFile = vbNullString

        '��� ����� ����� ��������
        Select Case Opt_HtmlJpgName
        Case 0    'filename
            tmpField = CheckNoNullValMyRS(dbFileNameInd, rs2proc)
            If Len(tmpField) <> 0 Then
                If Tokenize04(tmpField, a(), "|", False) > -1 Then
                '����� ������ ����
                    GetExtensionFromFileName GetNameFromPathAndName(a(0)), jFile
                End If
            End If
        Case 1    'title
            tmpField = CheckNoNullValMyRS(dbMovieNameInd, rs2proc)
            If Len(tmpField) <> 0 Then
                GetExtensionFromFileName GetNameFromPathAndName(tmpField), jFile
                ReplaceFNStr jFile
            End If
        Case 2    '������
            jFile = GetRNDFile
        End Select

        If Len(jFile) = 0 Then jFile = GetRNDFile


        '������� �������� �� ����
        PicSize = rs2proc.Fields(j).FieldSize
        If PicSize <> 0& Then

            ReDim b(PicSize - 1)
            b() = rs2proc.Fields(j).GetChunk(0, PicSize)

            '����������, ��� �� ���� � ���� ���������� PicExt
            Set vec = New Vector
            vec.BinaryData = b

            Set img = vec.ImageFile
            Set vec = Nothing

            If Not img Is Nothing Then

                Select Case img.FormatID
                Case wiaFormatBMP: PicExt = ".bmp"
                Case wiaFormatJPEG: PicExt = ".jpg"
                Case wiaFormatGIF: PicExt = ".gif"
                Case wiaFormatPNG: PicExt = ".png"
                Case wiaFormatTIFF: PicExt = ".tif"
                Case Else: PicExt = "pic"    '��� ��������, �� ������� ����������
                End Select

'                Select Case j '����
'                Case dbSnapShot1Ind: nm = 1: picDir = ss_Dir: o_picDir = o_ss_Dir
'                Case dbSnapShot2Ind: nm = 2: picDir = ss_Dir: o_picDir = o_ss_Dir
'                Case dbSnapShot3Ind: nm = 3: picDir = ss_Dir: o_picDir = o_ss_Dir
'                Case dbFrontFaceInd: nm = 0: picDir = cov_Dir: o_picDir = o_cov_Dir
'                End Select

                jFile = jFile & "_" & nm & PicExt    '���� ����� �������� � ����������
                JPGFile = picDir & jFile     '������ �������� � �����
                
                If FileExists(JPGFile) Then
                    '�������� ��� ����, ���� ���������
                    jFile = GetRNDFile & PicExt
                    JPGFile = picDir & jFile
                End If

                jpgHtmlArr(nm) = o_picDir & jFile     '� ������ ��� ������ (���)


                '�������� � ����, ������ ��������
                LFile = FreeFile
                Open JPGFile For Binary As #LFile
                Put #LFile, 1, b()
                Close #LFile

            End If    'If Not img Is Nothing
        Else    'no pic
        End If    'PicSize

    End If    'doflag
Next j

'                                                           ����

For j = 0 To lvIndexPole

    Select Case j
    Case dbMovieNameInd: Skey = "$TITLE$"
    Case dbLabelInd: Skey = "$LABEL$"
    Case dbGenreInd: Skey = "$GENRE$"
    Case dbYearInd: Skey = "$YEAR$"
    Case dbCountryInd: Skey = "$COUNTRY$"
    Case dbDirectorInd: Skey = "$DIRECTOR$"
    Case dbActerInd: Skey = "$ACTORS$"
    Case dbTimeInd: Skey = "$LENGTH$"
    Case dbResolutionInd: Skey = "$RESOLUTION$"
    Case dbAudioInd: Skey = "$AUDIO$"
    Case dbFpsInd: Skey = "$FRAMERATE$"
    Case dbFileLenInd: Skey = "$FILESIZE$"
    Case dbCDNInd: Skey = "$DISKS$"
    Case dbVideoInd: Skey = "$VIDEO$"
    Case dbFileNameInd: Skey = "$FILENAME$"
    Case dbDebtorInd: Skey = "$DEBTOR$"
    Case dbsnDiskInd: Skey = "$DISKSERIAL$"
    Case dbOtherInd: Skey = "$COMMENTS$"
    Case dbAnnotationInd: Skey = "$DESCRIPTION$"
    Case dbSubTitleInd: Skey = "$SUBTITLE$"
    Case dbCoverPathInd: Skey = "$COVERPATH$"
    Case dbMovieURLInd: Skey = "$URLMOVIE$"
    Case dbRatingInd: Skey = "$RATING$"
    Case dbMediaTypeInd: Skey = "$MEDIA$"
    Case dbLanguageInd: Skey = "$LANGUAGE$"
    End Select

    '������ �����
    tmpField = CheckNoNullValMyRS(j, rs2proc)
    If Len(tmpField) <> 0 Then
        If JSFlg Then
            ReplaceJSStr tmpField    '������ ���������� ��� JS � �����
        End If

        If j = dbAnnotationInd Then tmpField = Replace(tmpField, vbCrLf, "<BR>")  '������ ������� ������ �� ��
        BHTML = Replace(BHTML, Skey, tmpField, , , vbTextCompare)

    Else
        BHTML = Replace(BHTML, Skey, vbNullString, , , vbTextCompare)
    End If
Next j

'� ��������
If Len(jpgHtmlArr(1)) = 0 Then
    BHTML = Replace(BHTML, "$SNAPSHOT1$", "nopicture.jpg", , , vbTextCompare)
Else
    BHTML = Replace(BHTML, "$SNAPSHOT1$", jpgHtmlArr(1), , , vbTextCompare)
End If
If Len(jpgHtmlArr(2)) = 0 Then
    BHTML = Replace(BHTML, "$SNAPSHOT2$", "nopicture.jpg", , , vbTextCompare)
Else
    BHTML = Replace(BHTML, "$SNAPSHOT2$", jpgHtmlArr(2), , , vbTextCompare)
End If
If Len(jpgHtmlArr(3)) = 0 Then
    BHTML = Replace(BHTML, "$SNAPSHOT3$", "nopicture.jpg", , , vbTextCompare)
Else
    BHTML = Replace(BHTML, "$SNAPSHOT3$", jpgHtmlArr(3), , , vbTextCompare)
End If
If Len(jpgHtmlArr(0)) = 0 Then
    BHTML = Replace(BHTML, "$COVER$", "nopicture.jpg", , , vbTextCompare)
Else
    BHTML = Replace(BHTML, "$COVER$", jpgHtmlArr(0), , , vbTextCompare)
End If


'���������� �����
BHTML = Replace(BHTML, "$NUMBER$", CurNum, , , vbTextCompare)

Exit Sub
err:
'���� ���������� ������ � ����������
Debug.Print "Err_BODYHTML: " & tmpField
On Error Resume Next
Resume Next

End Sub
