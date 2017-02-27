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
'В новой рабочей книге создавать только один рабочий лист
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
Dim ar() As String 'массив названий полей
Dim i As Integer

On Error GoTo err
'Set ExcelApp = New Excel.Application
Set ExcelApp = GetExcelObject()



ExcelApp.Visible = True
'excel_sheet.Cells(row, 1) = rs!DataPoint

Set ExcelWB = NewWorkbook(ExcelApp)
'Set ExcelWB = OpenWorkbook(ExcelApp, "c:\catalogs\test.xls")

'закрепить первую строку для названий
ExcelApp.ActiveWindow.SplitRow = 1
ExcelApp.ActiveWindow.FreezePanes = True

'ExcelApp.Selection.columnWidth = 20
'ExcelApp.Sheets(1).Range.ColumnWidth = 30

ar = Split(PoleNames, ",")

'названия полей
For i = 0 To UBound(ar)
strFName = ExcelApp.Columns(i + 1).Address '$A:$A
strFName = right$(strFName, Len(strFName) - (InStr(1, strFName, ":") + 1)) ' A B ...

ExcelApp.Sheets(1).Range(strFName & 1).ColumnWidth = 20
ExcelApp.Sheets(1).Range(strFName & 1).Font.Bold = True
ExcelApp.Sheets(1).Range(strFName & 1).Interior.color = &HC0FFFF

'убрать метку <, > ">"
If right$(ar(i), 2) = " >" Or right$(ar(i), 2) = " <" Then ar(i) = left$(ar(i), Len(ar(i)) - 2)

ExcelApp.Sheets(1).Range(strFName & 1).Value = ar(i)
Next i

'Общие настройки

'
ExcelApp.Sheets(1).cells.VerticalAlignment = 1 'по верхнему краю
'ExcelApp.Sheets(1).Cells.HorizontalAlignment = 5 'с заполнением
'ExcelApp.Sheets(1).cells.WrapText = True
'ExcelApp.Columns(strFName).EntireColumn.AutoFit
'ExcelApp.Columns(strFName).EntireRow.AutoFit

''поправить формат ячейки аннотации - последняя всегда
'If LstExport_Arr(dbAnnotationInd) Then
''ExcelApp.Sheets(1).Range(strFName & 1).ColumnWidth = 100
'ExcelApp.Sheets(1).Range(strFName & 1).ColumnWidth = 100
'End If

'Данные
'If objRst.RecordCount <> 0 Then
'ExcelApp.Sheets(1).Range("A2").CopyFromRecordset objRst '911 символов не больше, иначе ошибка
'End If
'создадим запросом
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
Dim LVSelected As String    ' ключи помеченных строк списка LV - ( 10,11,)
Dim statement As String
'Dim tmp As String
Dim expFields As String    ' MovieName,Label,Genre - по базе
Dim rsExcel As DAO.Recordset
Dim j As Long, i As Long
Dim pArr() As Integer
Dim localNames As String    'Название,Метка,Жанр
Dim ProcessAnnotFlag As Boolean
Dim IP As Integer    'кол-во полей попавших в экспорт
Dim tmp As String

On Error Resume Next

'заполнить массис соответствия   индекс в списке - позиция
'индекс массива - позиция, значение - индекс поля списка
ReDim pArr(FrmMain.ListView.ColumnHeaders.Count)
For j = 0 To LstExport_ListCount    '   'все 0-24, потом индекс не обработаем
    pArr(FrmMain.ListView.ColumnHeaders(j + 1).Position) = j
Next j

'поля
For j = 1 To UBound(pArr)    '1-25 индекс не обработаем
    If LstExport_Arr(pArr(j)) Then
        If pArr(j) <> dbAnnotationInd Then
            expFields = expFields & "," & rs.Fields(pArr(j)).name
            IP = IP + 1
            'перевод
            localNames = localNames & "," & TranslatedFieldsNames(pArr(j))
        End If
    End If
Next j

' + аннотация
If LstExport_Arr(dbAnnotationInd) Then
    expFields = expFields & "," & rs.Fields(dbAnnotationInd).name
    localNames = localNames & "," & TranslatedFieldsNames(dbAnnotationInd)
    IP = IP + 1
    ProcessAnnotFlag = True
End If


If Len(expFields) = 0 Then
    'не выбраны поля экспорта
    ToDebug "No Exp2Exl Fields"
    Exit Sub
Else
    expFields = right$(expFields, Len(expFields) - 1)    '- левую ,
    localNames = right$(localNames, Len(localNames) - 1)
End If

'expFields = "Annotation"

If ch Then
    'checked
    statement = "SELECT " & expFields & " FROM Storage Where Checked = '1'"
Else
    'selected
    LVSelected = vbNullString
    Dim LVArr() As String    ' массив ключей lv - SelRowsKey -1 и без "
    ReDim LVArr(UBound(SelRowsKey) - 1)
    For i = 0 To UBound(LVArr)
        LVArr(i) = Val(SelRowsKey(i + 1))
    Next i
    LVSelected = Join(LVArr, ",")

    statement = "SELECT " & expFields & " FROM Storage WHERE Key IN (" & LVSelected & ")"
End If


On Error GoTo err

'убрать ентер из описания
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
'If Marker Then   выделенные иначе помеченные v
Dim FooterHTML As String
Dim BodyHTML As String
Dim HeaderHTML As String
Dim s As String
Dim tmpArr() As String    'хедер боди футер
Dim ABCDLine() As String    'нужен 1 - строка ссылок из хедера шаблона
Dim doFlagSS1 As Boolean, doFlagSS2 As Boolean, doFlagSS3 As Boolean, doFlagSS0 As Boolean
Dim strPath As String

Dim htmlDir As String
Dim CoverDir As String 'абс путь
Dim SShotDir As String
Dim o_CoverDir As String 'отн путь
Dim o_SShotDir As String

Dim htmlFile As String
Dim htmlFileFirst As String    'имя первого файла

Dim iFile As Integer    'номер файла html
Dim j As Long, i As Long    ', k As Long
Dim temp As String    ', temp2 As String

Dim rsRecordCount As Integer    'кол-во строк для вывода
Dim AllLinkStr() As String    'таблица нужных символов с линками, если надо
Dim AllCharStr() As String    'таблица нужных символов
Dim strAllLink As String    'вся строка ссылок
Dim rsTmp As DAO.Recordset
Dim sSQL As String
Dim arrPagesCount() As Long    'сколько каждой буквы есть в базе
Dim ABCLinks() As String    'tokenize 1 содержит
Dim ubABC As Long    'ubound ABCLinks()
Dim crFlag As Boolean    'надо ли переводить строку (\)

On Error GoTo err
'On Error GoTo 0

'прочитать шаблон снова
iFile = FreeFile
Open App.Path & "\Templates\" & CurrentHtmlTemplate For Binary As #iFile
's = Space$(LOF(iFile))
s = AllocString_ADV(LOF(iFile))
Get #iFile, , s
Close #iFile

'в какую папку?
temp = BrowseForFolderByPath(FixPath(lastHTMLfolderPath), NamesStore(4))
If Len(temp) = 0 Then FrmMain.Timer2.Enabled = True: Exit Sub
lastHTMLfolderPath = temp 'получаем без палочки

On Error Resume Next
'папки
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
    'в кучу
    htmlDir = lastHTMLfolderPath
    CoverDir = lastHTMLfolderPath & "\"
    SShotDir = lastHTMLfolderPath & "\"
    o_CoverDir = vbNullString
    o_SShotDir = vbNullString
End If
'If err Then Debug.Print "пр"
err.Clear
On Error GoTo err

'скопировать background и nopicture, стили
temp = App.Path & "\Templates\nopicture.jpg"
If FileExists(temp) Then FileCopy temp, htmlDir & "/nopicture.jpg"
temp = App.Path & "\Templates\background.jpg"
If FileExists(temp) Then FileCopy temp, htmlDir & "/background.jpg"
temp = App.Path & "\Templates\styles.css"
If FileExists(temp) Then FileCopy temp, htmlDir & "/styles.css"

'If Marker Then    'выделенные   ' нет мы можеи их не увидеть
'    rsRecordCount = SelCount
'Else    '          помеченные (v)
'    rsRecordCount = CheckCount
'End If
'замена общих ключей
's = Replace(s, "$TOTAL$", rsRecordCount, , , vbTextCompare)
s = Replace(s, "$OWNER$", GetPCUserName, , , vbTextCompare)
s = Replace(s, "$DATE$", Date, , , vbTextCompare)
s = Replace(s, "$TIME$", Format$(Time, "hh:mm"), , , vbTextCompare)
s = Replace(s, "$SVCBASENAME$", FrmMain.Caption, , , vbTextCompare)

'флаги для картинок
If InStrB(s, "$SNAPSHOT1$") > 0 Then doFlagSS1 = True
If InStrB(s, "$SNAPSHOT2$") > 0 Then doFlagSS2 = True
If InStrB(s, "$SNAPSHOT3$") > 0 Then doFlagSS3 = True
If InStrB(s, "$COVER$") > 0 Then doFlagSS0 = True

'делим на хедер, тело и футер
tmpArr() = Split(s, "$SVC_BODY$")
If UBound(tmpArr) <> 2 Then myMsgBox msgsvc(35), vbCritical, , FrmMain.hwnd: Screen.MousePointer = vbNormal: Exit Sub    ' не наш темплейт

'из хедера читаем строку ссылок ABCDLine(1)
ABCDLine() = Split(tmpArr(0), "$SVC_ABCD_LINE$")
If UBound(ABCDLine) <> 2 Then myMsgBox msgsvc(35), vbCritical, , FrmMain.hwnd: Screen.MousePointer = vbNormal: Exit Sub    ' не наш темплейт

'делим лайн на буквы по пробелу (будут типа Z\ - переводить строку после Z)

ubABC = Tokenize04(ABCDLine(1), ABCLinks(), " ", False)
If ubABC > -1 Then
    ReDim AllLinkStr(ubABC): ReDim arrPagesCount(ubABC): ReDim AllCharStr(ubABC)
Else
    ToDebug "Exp2ABC: no links"
    'Debug.Print "нет ссылок"
End If

DoEvents
Screen.MousePointer = vbHourglass

'прогресс
FrmMain.PBar.Max = ubABC: FrmMain.PBar.Value = 0: FrmMain.PBar.ZOrder 0
'заполнить AllLinkStr буквами, запросить и заполнить кол-вом arrPagesCount каждой буквы
For i = 0 To ubABC
    crFlag = False
    If InStr(ABCLinks(i), "\") > 0 Then
        crFlag = True        'после этого перевод строки
    End If

    If ABCLinks(i) <> "[0-9]" Then
        AllCharStr(i) = Chr$(Asc(ABCLinks(i)))    'оставить 1 чар
    Else
        AllCharStr(i) = ABCLinks(i)    '"[0-9]"
    End If

    If Marker Then    'выделенные

        'подготовим строку вылеленных ключей
        Dim LVArr() As String
        Dim LVSelected As String
        ReDim LVArr(UBound(SelRowsKey) - 1)
        For j = 0 To UBound(LVArr)
            LVArr(j) = Val(SelRowsKey(j + 1))
        Next j
        LVSelected = Join(LVArr, ",")
        Erase LVArr

        sSQL = "SELECT COUNT(MovieName) FROM STORAGE WHERE ((MovieName Like '" & AllCharStr(i) & "*') AND (Key IN (" & LVSelected & ")))"

    Else    '          помеченные (v)

        sSQL = "SELECT COUNT(MovieName) FROM STORAGE WHERE ((MovieName Like '" & AllCharStr(i) & "*') AND (Checked = '1'))"
    End If
    Set rsTmp = DB.OpenRecordset(sSQL)
    'заполнить массив кол-вом имеющихся в базе с учетом помеченного/выделенного
    If Not (rsTmp.BOF And rsTmp.EOF) Then arrPagesCount(i) = rsTmp(0)

    'сделать линк если есть данные
    If arrPagesCount(i) > 0 Then
        AllLinkStr(i) = "<a href='svc_" & AllCharStr(i) & ".htm'>" & AllCharStr(i) & "</a>"
    Else
        AllLinkStr(i) = AllCharStr(i)
    End If
    If crFlag Then AllLinkStr(i) = AllLinkStr(i) & "<br>"        'перевод строки

    FrmMain.PBar.Value = i    'двинуть прогресс
Next i
Erase ABCLinks: Erase ABCDLine
For i = 0 To ubABC
    strAllLink = strAllLink & "&nbsp;" & AllLinkStr(i)    'вся строка линков
    rsRecordCount = rsRecordCount + arrPagesCount(i)    'и сколько фильмов подходят под выбранные буквы
Next i

HeaderHTML = tmpArr(0): FooterHTML = tmpArr(2)
HeaderHTML = Replace(HeaderHTML, "$TOTAL$", rsRecordCount, , , vbTextCompare)    'всего можно увидеть
FooterHTML = Replace(FooterHTML, "$TOTAL$", rsRecordCount, , , vbTextCompare)

HeaderHTML = Replace(HeaderHTML, "$PAGELINE$", strAllLink, , , vbTextCompare)
FooterHTML = Replace(FooterHTML, "$PAGELINE$", strAllLink, , , vbTextCompare)

'стартуемый в обозревателе файлы (первый у которого есть фильмы)
For i = 0 To ubABC
    If arrPagesCount(i) > 0 Then
        htmlFileFirst = htmlDir & "\svc_" & AllCharStr(i) & ".htm"
        Exit For
    End If
Next i

'прогресс
FrmMain.PBar.Max = ubABC: FrmMain.PBar.Value = 0: FrmMain.PBar.ZOrder 0

'вывод страниц
For i = 0 To ubABC
    If arrPagesCount(i) > 0 Then

        If Marker Then    'выделенные
            sSQL = "SELECT * FROM STORAGE WHERE ((MovieName Like '" & AllCharStr(i) & "*') AND (Key IN (" & LVSelected & "))) Order by MovieName Asc"
        Else    '          помеченные (v)
            sSQL = "SELECT * FROM STORAGE WHERE ((MovieName Like '" & AllCharStr(i) & "*') AND (Checked = '1')) Order by MovieName Asc"
        End If
        'sSQL = "SELECT * FROM STORAGE WHERE MovieName Like '" & AllCharStr(i) & "*' Order by MovieName Asc"
        'Debug.Print sSQL

        Set rsTmp = DB.OpenRecordset(sSQL)

        If Not (rsTmp.BOF And rsTmp.EOF) Then
            rsTmp.MoveLast: rsTmp.MoveFirst

            iFile = FreeFile    'создать и заполнить файл
            htmlFile = htmlDir & "\svc_" & AllCharStr(i) & ".htm"   'абсолютный путь для нового файла

            Open htmlFile For Output As #iFile
            'хедер
            Print #iFile, Replace(HeaderHTML, "$NUMBER$", arrPagesCount(i), , , vbTextCompare)

            'body
            For j = 1 To arrPagesCount(i)
                BodyHTML = tmpArr(1)
                Call Export2HTML_BODY(BodyHTML, j, False, CoverDir, SShotDir, o_CoverDir, o_SShotDir, rsTmp)
                Print #iFile, BodyHTML
                rsTmp.MoveNext
            Next j

            'футер
            Print #iFile, Replace(FooterHTML, "$NUMBER$", arrPagesCount(i), , , vbTextCompare)
            Close #iFile

        End If    'If Not (rsTMP.BOF And rsTMP.EOF) Then
    End If    'If arrPagesCount(i) <> 0

    FrmMain.PBar.Value = i    'двинуть прогресс
Next i

Screen.MousePointer = vbNormal
FrmMain.TextItemHid.ZOrder 0
'не двигали RestoreBasePos

'открыть страницу
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
'для фреймов
' + узнавать фреймовый шаблон JSFlag
' + Игнорировать многостраничность, делать только 1
' + пытаться засунуть создание фреймов скриптом в один файл svc1.htm (svc_интерактив)
' + В ИМЕНАХ ФАЙЛОв НЕ ДОЛЖНО БЫТЬ " '
' + менять " на \"
' - какие еще запрещенные символы?
' Не забыть включить err

If rs Is Nothing Then Exit Sub

'шаблоны
Dim FooterHTML As String
Dim BodyHTML As String
Dim HeaderHTML As String
Dim s As String
Dim tmpArr() As String
Dim jpgHtmlArr(3) As String    'от nm (0123) массив путей картинок для html
Dim Skey As String    'ключи в зависимости от j
Dim JSFlag As Boolean    'флаг экспорта в в html с JS

'''
Dim strPath As String

Dim htmlDir As String
Dim CoverDir As String 'абс путь
Dim SShotDir As String
Dim o_CoverDir As String 'отн путь
Dim o_SShotDir As String

Dim htmlFile As String
Dim iFile As Integer    'номер файла html
Dim JPGFile As String
Dim ret As Long
Dim ubnd As Long
Dim j As Integer, i As Integer, M As Long
Dim temp As String    ', temp2 As String
Dim doflag As Boolean

Dim lPtr As Long
Dim PicSize As Long
Dim b() As Byte
Dim jFile As String    ' имя файла jpeg
Dim LFile As Integer    'свободный файл для картинки
Dim img As ImageFile
Dim vec As Vector
Dim PicExt As String

'постраничность
Dim PagesCount As Integer
Dim RecsOnPage As Integer
Dim rcount As Integer    ' счетчик записей
Dim htmlFileFirst As String    'имя первого файла
Dim StartStr As String
Dim Ostatok As Integer    'записей на посл странице
Dim TechCount As Integer    'индекс начала цикла для текущей страницы. начало с 1
Dim rsRecordCount As Integer    'кол-во строк для вывода
Dim AllLinkStr() As String    'Группа линков на др. страницы
Dim LinkStr() As String    'один линк
'''
Dim nm As Integer    'цифровая прибавка в имени файла картинки
'Dim strFoldPath As String, strFoldName As String
Dim anno As String    'заменить на <br>
Dim tmpField As String    'значение поля базы


On Error GoTo ErrHandler
'без шаблона не начнем...
If Len(CurrentHtmlTemplate) = 0 Then myMsgBox msgsvc(34), vbCritical, , FrmMain.hwnd: Exit Sub    'Определите шаблон
FrmMain.Timer2.Enabled = False

'прочитать шаблон (до исп. чтения информ-флагов шаблона)
iFile = FreeFile
Open App.Path & "\Templates\" & CurrentHtmlTemplate For Binary As #iFile
's = Space$(LOF(iFile))
s = AllocString_ADV(LOF(iFile))
Get #iFile, , s
Close #iFile

'если абвг-шаблон - перейти в другую подпрограмму
If InStr(1, s, "$SVC.ABCD$", vbTextCompare) Then
    On Error GoTo 0
    Call Export2HTML_ABCD(Marker)
    Exit Sub
End If
'спросить шаблон, нужно ли менять текст для совместимости с JS
If InStr(1, s, "$SVC.JS.Array$", vbTextCompare) Then JSFlag = True

'в какую папку?
temp = BrowseForFolderByPath(FixPath(lastHTMLfolderPath), NamesStore(4))
If Len(temp) = 0 Then FrmMain.Timer2.Enabled = True: Exit Sub
lastHTMLfolderPath = temp

On Error Resume Next
'папки 'в отн ссылках для страниц явы надо такую черту /
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
    'в кучу
    htmlDir = lastHTMLfolderPath
    CoverDir = lastHTMLfolderPath & "\"
    SShotDir = lastHTMLfolderPath & "\"
    o_CoverDir = vbNullString
    o_SShotDir = vbNullString
End If
'If err Then Debug.Print "пр"
err.Clear
On Error GoTo ErrHandler


If Not JSFlag Then
    'только если не фрейм
    'поделим по страницам
    If TxtNnOnPage_Text > 0 Then
        RecsOnPage = Val(TxtNnOnPage_Text)
    Else
        RecsOnPage = 30
    End If

    If Marker Then    'выделенные
        PagesCount = UBound(SelRows) / RecsOnPage
        Ostatok = UBound(SelRows) - PagesCount * RecsOnPage
        rsRecordCount = UBound(SelRows)
    Else    '          помеченные (v)
        PagesCount = UBound(CheckRows) / RecsOnPage
        Ostatok = UBound(CheckRows) - PagesCount * RecsOnPage
        rsRecordCount = UBound(CheckRows)
    End If

    If Ostatok > 0 Then PagesCount = PagesCount + 1

    'подготовка строки с номерами страниц
    ReDim LinkStr(PagesCount)
    ReDim AllLinkStr(PagesCount)

    For i = 1 To PagesCount
        htmlFile = "svc" & i & ".htm"    'относительный, для ссылки
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
    'Java, одна страница
    PagesCount = 1
    'подготовка строки с номерами страниц
    ReDim LinkStr(PagesCount)
    ReDim AllLinkStr(PagesCount)
    If Marker Then    'выделенные
        RecsOnPage = SelCount
    Else    '          помеченные (v)
        RecsOnPage = CheckCount
    End If

End If    'jsflag

'скопировать background и nopicture, стили
temp = App.Path & "\Templates\nopicture.jpg"
If FileExists(temp) Then FileCopy temp, htmlDir & "/nopicture.jpg"
temp = App.Path & "\Templates\background.jpg"
If FileExists(temp) Then FileCopy temp, htmlDir & "/background.jpg"
temp = App.Path & "\Templates\styles.css"
If FileExists(temp) Then FileCopy temp, htmlDir & "/styles.css"

DoEvents
Screen.MousePointer = vbHourglass

rcount = 0: TechCount = 1

'стартуемые в обозревателе файлы
If PagesCount = 1 Then
    htmlFileFirst = htmlDir & "/index.htm"
Else
    htmlFileFirst = htmlDir & "/svc1.htm"
End If

'замена общих ключей
s = Replace(s, "$TOTAL$", rsRecordCount, , , vbTextCompare)
s = Replace(s, "$OWNER$", GetPCUserName, , , vbTextCompare)
s = Replace(s, "$DATE$", Date, , , vbTextCompare)
s = Replace(s, "$TIME$", Format$(Time, "hh:mm"), , , vbTextCompare)
s = Replace(s, "$SVCBASENAME$", FrmMain.Caption, , , vbTextCompare)


    'делим на хедер, тело и футер
    tmpArr() = Split(s, "$SVC_BODY$")
    If UBound(tmpArr) <> 2 Then myMsgBox msgsvc(35), vbCritical, , FrmMain.hwnd: Screen.MousePointer = vbNormal: Exit Sub    ' не наш темплейт

    
'                                                                   Постраничный цикл

For i = 1 To PagesCount
    'замена i ключей
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

    htmlFile = htmlDir & "/svc" & i & ".htm"    'абсолютный для файла

    'создать новый html
    iFile = FreeFile
    Open htmlFile For Output As #iFile
    'хедер
    Print #iFile, HeaderHTML

    'полей на каждую страницу
    If Marker Then
        ubnd = UBound(SelRows)
    Else
        ubnd = UBound(CheckRows)
    End If

    FrmMain.PBar.ZOrder 0: FrmMain.PBar.Max = ubnd

    '                                                                   цикл по кол-ву на странице
    For M = TechCount To ubnd
        FrmMain.PBar.Value = M

        If Marker Then
            If MultiSel Then
                RSGoto SelRowsKey(M)
            End If
        Else
            RSGoto CheckRowsKey(M)
        End If

        'не начать ли след. страницу
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

'открыть страницу
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
Dim jpgHtmlArr(3) As String    'от nm (0123) массив путей картинок для html
Dim j As Integer
Dim doflag As Boolean
Dim doFlagSS1 As Boolean, doFlagSS2 As Boolean, doFlagSS3 As Boolean, doFlagSS0 As Boolean
Dim a() As String    'tokenize
Dim jFile As String    ' имя файла jpeg
Dim tmpField As String
Dim PicSize As Long
Dim b() As Byte
Dim LFile As Integer    'свободный файл для картинки
Dim img As ImageFile
Dim vec As Vector
Dim PicExt As String
Dim nm As Integer    'цифровая прибавка в имени файла картинки
Dim JPGFile As String
Dim Skey As String    'ключи в зависимости от j
Dim picDir As String 'темпы
Dim o_picDir As String

On Error GoTo err 'для img

'флаги для картинок
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
        'имена jpeg файлов
        jFile = vbNullString

        'как брать имена картинок
        Select Case Opt_HtmlJpgName
        Case 0    'filename
            tmpField = CheckNoNullValMyRS(dbFileNameInd, rs2proc)
            If Len(tmpField) <> 0 Then
                If Tokenize04(tmpField, a(), "|", False) > -1 Then
                'взять первый файл
                    GetExtensionFromFileName GetNameFromPathAndName(a(0)), jFile
                End If
            End If
        Case 1    'title
            tmpField = CheckNoNullValMyRS(dbMovieNameInd, rs2proc)
            If Len(tmpField) <> 0 Then
                GetExtensionFromFileName GetNameFromPathAndName(tmpField), jFile
                ReplaceFNStr jFile
            End If
        Case 2    'рандом
            jFile = GetRNDFile
        End Select

        If Len(jFile) = 0 Then jFile = GetRNDFile


        'достать картинку из базы
        PicSize = rs2proc.Fields(j).FieldSize
        If PicSize <> 0& Then

            ReDim b(PicSize - 1)
            b() = rs2proc.Fields(j).GetChunk(0, PicSize)

            'определить, что за файл и дать расширение PicExt
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
                Case Else: PicExt = "pic"    'для проверки, не рабочее расширение
                End Select

'                Select Case j 'выше
'                Case dbSnapShot1Ind: nm = 1: picDir = ss_Dir: o_picDir = o_ss_Dir
'                Case dbSnapShot2Ind: nm = 2: picDir = ss_Dir: o_picDir = o_ss_Dir
'                Case dbSnapShot3Ind: nm = 3: picDir = ss_Dir: o_picDir = o_ss_Dir
'                Case dbFrontFaceInd: nm = 0: picDir = cov_Dir: o_picDir = o_cov_Dir
'                End Select

                jFile = jFile & "_" & nm & PicExt    'дать номер картинки и расширение
                JPGFile = picDir & jFile     'полное название с путем
                
                If FileExists(JPGFile) Then
                    'название уже есть, дать случайное
                    jFile = GetRNDFile & PicExt
                    JPGFile = picDir & jFile
                End If

                jpgHtmlArr(nm) = o_picDir & jFile     'в массив для ссылки (отн)


                'положить в файл, писать напрямую
                LFile = FreeFile
                Open JPGFile For Binary As #LFile
                Put #LFile, 1, b()
                Close #LFile

            End If    'If Not img Is Nothing
        Else    'no pic
        End If    'PicSize

    End If    'doflag
Next j

'                                                           поля

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

    'замена полей
    tmpField = CheckNoNullValMyRS(j, rs2proc)
    If Len(tmpField) <> 0 Then
        If JSFlg Then
            ReplaceJSStr tmpField    'менять запрещенки для JS в полях
        End If

        If j = dbAnnotationInd Then tmpField = Replace(tmpField, vbCrLf, "<BR>")  'менять перевод строки на бр
        BHTML = Replace(BHTML, Skey, tmpField, , , vbTextCompare)

    Else
        BHTML = Replace(BHTML, Skey, vbNullString, , , vbTextCompare)
    End If
Next j

'и картинок
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


'порядковый номер
BHTML = Replace(BHTML, "$NUMBER$", CurNum, , , vbTextCompare)

Exit Sub
err:
'типа обозначить ошибку и продолжить
Debug.Print "Err_BODYHTML: " & tmpField
On Error Resume Next
Resume Next

End Sub
