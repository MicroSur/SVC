Attribute VB_Name = "modSQL"
Option Explicit
Public Function FilterLikeString(fStr As String, sStr As String) As String
'ищет вхождение в строке, вхождение - должно быть отдельным словом, (не как instr)
'fStr где искать
'sStr что искать

'1 Драма
'2 Мелодрама

'найти Драма, должны получить только первую строку

FilterLikeString = "((" & fStr & _
" Like '*[!0-9A-я]" & sStr & "[!0-9:A-я]*') Or (" & fStr & _
" Like '" & sStr & "[!0-9:A-я]*') Or (" & fStr & _
" Like '*[!0-9A-я]" & sStr & "') Or (" & fStr & _
" Like '" & sStr & "'))"

'FilterLikeString = "((" & fStr & " Like '*" & sStr & "*'))"


'с inStr ?
'instr(1,

If Len(sStr) = 0 Then FilterLikeString = "(" & fStr & " = '')"

End Function
Public Function GetGroupNum(s As String) As String
'пощитать запросом кол-во потрошенных значений
'FilterLikeString оччен долог

Dim strSQL As String
Dim strLike As String
'mzt Dim tmpL As Long
Dim rsTV As DAO.Recordset

'Exit Function
On Error Resume Next

SQLCompatible s

If s <> "Null" Then
    strLike = FilterLikeString(GroupField, s)
Else
    strLike = GroupField & " Is Null"
End If

'strSQL = "Select Count(" & GroupField & ") From Storage WHERE " & strLike    'Group By " & GroupField
strSQL = "Select Count(*) From Storage WHERE " & strLike    'Group By " & GroupField

Set rsTV = DB.OpenRecordset(strSQL)

If Not (rsTV.BOF And rsTV.EOF) Then    'If rsTV.RecordCount > 0 Then
    'rsTV.MoveFirst
    GetGroupNum = rsTV(0)
End If

'rsTV.MoveLast: rsTV.MoveFirst
'GetGroupNum = rsTV.RecordCount
Set rsTV = Nothing
End Function

Public Sub FilterItemsSQL(IsTyped As Boolean, Optional sNot As String)

'''!!!  Обязательно OnError !!!!!!!! для ошибок неправильного шаблона

'принять IsTyped чтобы по false не считать текстовый запрос, а перейти к картинкам
Dim strSQL As String
'Dim DoFlag As Boolean
Dim SelectString As String
Dim s(11) As String    ' =.cbs(i) для краткости наверно?
Dim i As Integer    ', n As Integer, k As Integer
Dim a As Integer, o As Integer

'Dim iAND As Integer    'индекс, попавшийся со знаком AND
'Dim iFill As Integer    'индекс, первый заполненный
Dim aOr() As Integer    'индексы всех OR
Dim aAnd() As Integer    'индексы всех AND
Dim aEmp() As Integer    'индексы незаполненных пользователем

ReDim aOr(0): ReDim aAnd(0): ReDim aEmp(0)

SelectString = "SELECT * FROM Storage WHERE ("

Screen.MousePointer = vbHourglass
With FrmFilter

    If IsTyped Then    'что-то вписано в фильтр

        'iAND = -1: iFill = -1    'для начала
        For i = 0 To cbTotal - 1
            s(i) = .cbs(i)

            If Len(s(i)) <> 0 Then
                'дублировать спецсимвол
                SQLCompatible s(i)
                'найти все Or и AND в массивы (пойдут с 1)
                If .chAO(i).Value = vbChecked Then
                    ReDim Preserve aAnd(UBound(aAnd) + 1)
                    aAnd(UBound(aAnd)) = i
                Else
                    ReDim Preserve aOr(UBound(aOr) + 1)
                    aOr(UBound(aOr)) = i
                End If
                '                'найти первый (step -1) со знаком AND
                '                If .chAO(i).Value = vbChecked Then iAND = i
                '                'найти первый (step -1) заполненный
                '                iFill = i
            Else    'пустышки
                ReDim Preserve aEmp(UBound(aEmp) + 1)
                aEmp(UBound(aEmp)) = i
            End If
        Next i

        'с любой частью поля (общий запрос)

        'или жанр = 1
        'и   жанр = 2
        'и страна = 3
        'и рейтинг= 4

        '1 Идея сгруппировать все по полям (жанр отдельно, год отдельно)
        'внутри групп расставить знаки



        '2 идея определить в каком массиве одиночка и поставить его в конец со знаком
        ' и определять все Or в первую группу

        If (UBound(aAnd) = 1) And (UBound(aOr) = 1) Then
            'если по одному с разными знаками ставить их через AND
            a = aAnd(1): o = aOr(1)
            strSQL = "(" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ") AND (" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"
            'посерить галочку у or
            '.chAO(o).Value = vbGrayed


            'If (UBound(aAnd) = 1) And (UBound(aOr) = 1) Then
            ''если по одному с разными знаками ставить их через OR
            'a = aAnd(1): o = aOr(1)
            'strSQL = "(" & IPN(.cbl(a).ListIndex) &  CheckForShablon(s(a)) & ") OR (" & IPN(.cbl(o).ListIndex) &  CheckForShablon(s(o)) & ")"

        ElseIf (UBound(aAnd) = 1) And (UBound(aOr) = 0) Then
            'если только 1 and и 0 or
            a = aAnd(1)
            strSQL = "(" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ")"

        ElseIf (UBound(aAnd) = 0) And (UBound(aOr) = 1) Then
'одно значение без галочек
'если только 1 or и 0 and
            o = aOr(1)
            'strSQL = "(" & IPN(.cbl(o).ListIndex) & CheckForShablon(s(o), o) & ")"
            strSQL = "(" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"
'Debug.Print strSQL
        ElseIf (UBound(aAnd) > 1) And (UBound(aOr) = 0) Then
            'если много and и 0 or
            'добавить первый and
            a = aAnd(1)
            strSQL = "(" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ")"
            'добить остальными AND
            For i = 2 To UBound(aAnd)
                a = aAnd(i)
                strSQL = strSQL & " AND (" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ")"
            Next i

        ElseIf (UBound(aAnd) = 0) And (UBound(aOr) > 1) Then
            'если много or и 0 and
            'взять первый Or
            o = aOr(1)
            strSQL = "(" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"
            'и остальные or
            For i = 2 To UBound(aOr)
                o = aOr(i)
                strSQL = strSQL & " OR (" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"
            Next i

        ElseIf (UBound(aAnd) = 1) And (UBound(aOr) > 1) Then
            'если только 1 and и много or
            'сначала  1 and  потом and группу Or
            'добавить первый and
            a = aAnd(1)
            strSQL = "(" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ")"
            'добавить первый Or
            o = aOr(1)
            strSQL = strSQL & " AND ((" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"
            'и остальные or
            For i = 2 To UBound(aOr)
                o = aOr(i)
                strSQL = strSQL & " OR (" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"
            Next i
            'закрыть группу or
            strSQL = strSQL & ")"


            'ElseIf (UBound(aAnd) = 1) And (UBound(aOr) > 1) Then
            ''если только 1 and и много or
            ''сначала группа Or потом 1 and
            ''взять первый Or
            'o = aOr(1)
            'strSQL = "((" & IPN(.cbl(o).ListIndex) &  CheckForShablon(s(o)) & ")"
            ''и остальные or
            'For i = 2 To UBound(aOr)
            'o = aOr(i)
            'strSQL = strSQL & " OR (" & IPN(.cbl(o).ListIndex) &  CheckForShablon(s(o)) & ")"
            'Next i
            ''закрыть группу or
            'strSQL = strSQL & ")"
            ''добавить 1 AND
            'a = aAnd(1)
            'strSQL = strSQL & " AND (" & IPN(.cbl(a).ListIndex) &  CheckForShablon(s(a)) & ")"

        ElseIf (UBound(aAnd) > 1) And (UBound(aOr) = 1) Then
            'если только 1 or и много and
            'сначала все And потом and 1 Or
            'добавить первый and
            a = aAnd(1)
            strSQL = "(" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ")"
            'добить остальными AND
            For i = 2 To UBound(aAnd)
                a = aAnd(i)
                strSQL = strSQL & " AND (" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ")"
            Next i
            'добавит 1 Or
            o = aOr(1)
            strSQL = strSQL & " AND (" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"

            'ElseIf (UBound(aAnd) > 1) And (UBound(aOr) = 1) Then
            ''если только 1 or и много and
            ''сначала все And потом 1 Or
            ''добавить первый and
            'a = aAnd(1)
            'strSQL = "(" & IPN(.cbl(a).ListIndex) &  CheckForShablon(s(a)) & ")"
            ''добить остальными AND
            'For i = 2 To UBound(aAnd)
            'a = aAnd(i)
            'strSQL = strSQL & " AND (" & IPN(.cbl(a).ListIndex) &  CheckForShablon(s(a)) & ")"
            'Next i
            ''добавит 1 Or
            'o = aOr(1)
            'strSQL = strSQL & " OR (" & IPN(.cbl(o).ListIndex) &  CheckForShablon(s(o)) & ")"

        ElseIf (UBound(aAnd) > 1) And (UBound(aOr) > 1) Then
            'если много or и много and
            'сначала все and потом and группу Or
            'добавить первый and
            a = aAnd(1)
            strSQL = "(" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ")"
            'добить остальными AND
            For i = 2 To UBound(aAnd)
                a = aAnd(i)
                strSQL = strSQL & " AND (" & IPN(.cbl(a).ListIndex, s(a)) & CheckForShablon(s(a), a) & ")"
            Next i
            'добавить первый Or
            o = aOr(1)
            strSQL = strSQL & " AND ((" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"
            'и остальные or
            For i = 2 To UBound(aOr)
                o = aOr(i)
                strSQL = strSQL & " OR (" & IPN(.cbl(o).ListIndex, s(o)) & CheckForShablon(s(o), o) & ")"
            Next i
            'закрыть группу or
            strSQL = strSQL & ")"


            'ElseIf (UBound(aAnd) > 1) And (UBound(aOr) > 1) Then
            ''если много or и много and
            ''сначала все and потом все Or
            ''добавить первый and
            'a = aAnd(1)
            'strSQL = "(" & IPN(.cbl(a).ListIndex) &  CheckForShablon(s(a)) & ")"
            ''добить остальными AND
            'For i = 2 To UBound(aAnd)
            'a = aAnd(i)
            'strSQL = strSQL & " AND (" & IPN(.cbl(a).ListIndex) &  CheckForShablon(s(a)) & ")"
            'Next i
            ''добить всеми or
            'For i = 1 To UBound(aOr)
            'o = aOr(i)
            'strSQL = strSQL & " OR (" & IPN(.cbl(o).ListIndex) &  CheckForShablon(s(o)) & ")"
            'Next i

            'ElseIf (UBound(aAnd) > 1) And (UBound(aOr) > 1) Then
            ''если много or и много and
            ''сначала группа Or потом все and
            ''взять первый or
            'o = aOr(1)
            'strSQL = "((" & IPN(.cbl(o).ListIndex) &  CheckForShablon(s(o)) & ")"
            ''и остальные or
            'For i = 2 To UBound(aOr)
            'o = aOr(i)
            'strSQL = strSQL & " OR (" & IPN(.cbl(o).ListIndex) &  CheckForShablon(s(o)) & ")"
            'Next i
            ''закрыть группу or
            'strSQL = strSQL & ")"
            ''добить всеми AND
            'For i = 1 To UBound(aAnd)
            'a = aAnd(i)
            'strSQL = strSQL & " AND (" & IPN(.cbl(a).ListIndex) &  CheckForShablon(s(a)) & ")"
            'Next i
        End If

        'конец общего запроса


        'If .ChFiltWhole.Value = vbUnchecked Then
        ' If .chFiltStart.Value = vbUnchecked Then
        ' 'с любой частью поля (уже сделано)
        '
        '
        ''[!0-9:A-я]
        ''рейтинк от 7 до 10  (от 7 до 9 и 1)     [7-91]*
        ''. / * : ! # &
        '
        'Else    'с Начала поля (для быстрого фильтра по первой букве)
        ''заменить '* на ' в общем запросе
        ''strSQL = Replace(strSQL, "'*", "'")
        ''уже в CheckForShablon
        'End If    'chFiltStart
        '
        'Else
        ''поле целиком
        ''заменить ('* на ') и (*' на ') в общем запросе
        ''шаблончики * отрезаются, но можно ?
        ''strSQL = Replace(strSQL, "'*", "'")
        ''strSQL = Replace(strSQL, "*'", "'")
        ''уже в CheckForShablon
        '
        'End If

    End If    'IsTyped

    If .ChFiltSShots.Value = vbChecked Then
        'если ИСКЛЮЧИТЬ
        If sNot = "NOT" Then
            'strSQL = strSQL & " Or ((SnapShot1 <> '') Or (SnapShot2 <> '') Or (SnapShot3 <> ''))"
        Else
            'strSQL = strSQL & " And ((SnapShot1 <> '') Or (SnapShot2 <> '') Or (SnapShot3 <> ''))"
            strSQL = strSQL & " And ((SnapShot1 Is Not Null) Or (SnapShot2 Is Not Null) Or (SnapShot3 Is Not Null))"
        End If
    End If
    If .ChFiltCover.Value = vbChecked Then
        'если ИСКЛЮЧИТЬ
        If sNot = "NOT" Then
            'strSQL = strSQL & " And (FrontFace <> '')"
            'strSQL = strSQL & " And (False)" 'если текстовые пустые чтобы не отрезать пустой нот
        Else
            'strSQL = strSQL & " And (FrontFace <> '')"
            strSQL = strSQL & " And (FrontFace Is Not Null)"
        End If
    End If
    If Left$(strSQL, 5) = " And " Then
        strSQL = Right$(strSQL, Len(strSQL) - 5)
    End If
    If Left$(strSQL, 4) = " Or " Then
        strSQL = Right$(strSQL, Len(strSQL) - 4)
    End If

    'если ИСКЛЮЧИТЬ
    If sNot = "NOT" Then

        ''и инвертировать
        'strSQL = "Not (" & strSQL & ")"

        'добавить еще все заполненные Or Is Null
        For i = 1 To UBound(aAnd)
            a = aAnd(i)
            strSQL = strSQL & " Or (" & IPN(.cbl(a).ListIndex, s(a)) & " Is Null)"
        Next i
        For i = 1 To UBound(aOr)
            o = aOr(i)
            strSQL = strSQL & " Or (" & IPN(.cbl(o).ListIndex, s(a)) & " Is Null)"
        Next i

        ''и инвертировать
        strSQL = "Not (" & strSQL & ")"



        If .ChFiltSShots.Value = vbChecked Then
            strSQL = strSQL & " and (SnapShot1 Is Null) and (SnapShot2 Is Null) and (SnapShot3 Is Null)"
        End If

        If .ChFiltCover.Value = vbChecked Then
            strSQL = strSQL & " and (FrontFace Is Null)"
        End If

        If LCase$(Left$(strSQL, 11)) = "not () and " Then strSQL = Right$(strSQL, Len(strSQL) - 11)

        '''и инвертировать
        'strSQL = "Not (" & strSQL & ")"

    End If    'not

End With

LastSQLFilterString = "(" & strSQL & ")"    '1
strSQL = SelectString & strSQL    '2

If GroupedFlag Then
    'добавить запрос от группировки
    strSQL = strSQL & " AND " & LastSQLGroupString & ")"
Else
    strSQL = strSQL & ")"
End If

'strSQL = "SELECT * FROM Storage WHERE ((Genre Like '*Приключения*') and ((Genre Like '*Комедия*') OR (Country Like '*США*') OR (Rating Like '*7*')))"
'strSQL = "SELECT * FROM Storage WHERE (Not (((SnapShot1 <> '') Or (SnapShot2 <> '') Or (SnapShot3 <> ''))) or SnapShot1 Is Null)"
'strSQL = "SELECT * FROM Storage WHERE ((Genre Like '*Приключения*') OR (Country Like '*США*') AND (Genre Like '*Комедия*') AND (Rating Like '[7-91]*'))"
'strSQL = "SELECT * FROM Storage WHERE Not ( (Year Like '*2006*') or (Year Is Null) ) and (FrontFace Is Null)"
'strSQL = "SELECT * FROM Storage WHERE (Not ( (Year Like '*2006*') Or (FrontFace <> '') ) or (Year Is Null) and (FrontFace Is Null) )"

'strSQL = "SELECT * FROM Storage WHERE ((Genre Like '*Боевик*') AND (Genre not Like '*Комедия*'))"
'strSQL = "SELECT * FROM Storage WHERE ((cvar(Label) > cvar('9,1')))"

'Debug.Print "FIS: " & strSQL

On Error GoTo err
Set rs = DB.OpenRecordset(strSQL)
FilteredFlag = True    'флаг неполного показа
FrmMain.FillListView
FrmMain.ComFilter.BackColor = &HC0C0FF

Screen.MousePointer = vbNormal
Exit Sub

err:
Screen.MousePointer = vbNormal
MsgBox msgsvc(46), vbExclamation
ToDebug "Err_fisq" ' & err.Description

End Sub



Public Function IPN(i As Integer, s As String) As String
'возвращает имя поля по индексу выпадающего листа в фильтре
'i c 0
Select Case i
Case 0
If Len(GetMathFilter(s)) Then
    IPN = "Val(MovieName)" 'есть знак - число
Else
    IPN = "MovieName"
End If

Case 1
If Len(GetMathFilter(s)) Then
    IPN = "Val(Label)" 'есть знак - число
Else
    IPN = "Label"
End If

Case 2
If Len(GetMathFilter(s)) Then
    IPN = "Val(Genre)" 'есть знак - число
Else
    IPN = "Genre"
End If

Case 3
If Len(GetMathFilter(s)) Then
    IPN = "Val(Year)" 'есть знак - число
Else
    IPN = "Year"
End If

Case 4
If Len(GetMathFilter(s)) Then
    IPN = "Val(Country)" 'есть знак - число
Else
    IPN = "Country"
End If

Case 5
If Len(GetMathFilter(s)) Then
    IPN = "Val(Director)" 'есть знак - число
Else
    IPN = "Director"
End If

Case 6
If Len(GetMathFilter(s)) Then
    IPN = "Val(Acter)" 'есть знак - число
Else
    IPN = "Acter"
End If

Case 7
If Len(GetMathFilter(s)) Then
    IPN = "Val(Time)" 'есть знак - число
Else
    IPN = "Time"
End If

Case 8
If Len(GetMathFilter(s)) Then
    IPN = "Val(Resolution)" 'есть знак - число
Else
    IPN = "Resolution"
End If

Case 9
If Len(GetMathFilter(s)) Then
    IPN = "Val(Audio)" 'есть знак - число
Else
    IPN = "Audio"
End If

Case 10
If Len(GetMathFilter(s)) Then
    IPN = "Val(FPS)" 'есть знак - число
Else
    IPN = "FPS"
End If

Case 11
If Len(GetMathFilter(s)) Then
    IPN = "Val(FileLen)" 'есть знак - число
Else
    IPN = "FileLen"
End If

Case 12
If Len(GetMathFilter(s)) Then
    IPN = "Val(CDN)" 'есть знак - число
Else
    IPN = "CDN"
End If

Case 13
If Len(GetMathFilter(s)) Then
    IPN = "Val(MediaType)" 'есть знак - число
Else
    IPN = "MediaType"
End If

Case 14
If Len(GetMathFilter(s)) Then
    IPN = "Val(Video)" 'есть знак - число
Else
    IPN = "Video"
End If

Case 15
If Len(GetMathFilter(s)) Then
    IPN = "Val(SubTitle)" 'есть знак - число
Else
    IPN = "SubTitle"
End If

Case 16
If Len(GetMathFilter(s)) Then
    IPN = "Val(Language)" 'есть знак - число
Else
    IPN = "Language"
End If

Case 17
If Len(GetMathFilter(s)) Then
    IPN = "Val(Rating)" 'есть знак - число
Else
    IPN = "Rating"
End If

Case 18
If Len(GetMathFilter(s)) Then
    IPN = "Val(FileName)" 'есть знак - число
Else
    IPN = "FileName"
End If

Case 19
If Len(GetMathFilter(s)) Then
    IPN = "Val(Debtor)" 'есть знак - число
Else
    IPN = "Debtor"
End If

Case 20
If Len(GetMathFilter(s)) Then
    IPN = "Val(snDisk)" 'есть знак - число
Else
    IPN = "snDisk"
End If

Case 21
If Len(GetMathFilter(s)) Then
    IPN = "Val(Other)" 'есть знак - число
Else
    IPN = "Other"
End If

Case 22
If Len(GetMathFilter(s)) Then
    IPN = "Val(CoverPath)" 'есть знак - число
Else
    IPN = "CoverPath"
End If

Case 23
If Len(GetMathFilter(s)) Then
    IPN = "Val(MovieURL)" 'есть знак - число
Else
    IPN = "MovieURL"
End If

Case 24
If Len(GetMathFilter(s)) Then
    IPN = "Val(Annotation)" 'есть знак - число
Else
    IPN = "Annotation"
End If

End Select
End Function
Private Function GetMathFilter(s As String) As String
Dim znak As String
'есть ли мат знак в начале строки поиска?
If Left$(s, 1) = ">" Then znak = ">"
If Left$(s, 1) = "<" Then znak = "<"
If Left$(s, 1) = "=" Then znak = "="
If Left$(s, 2) = ">=" Then
    znak = ">="
ElseIf Left$(s, 2) = "<=" Then
    znak = "<="
ElseIf Left$(s, 2) = "<>" Then
    znak = "<>"
End If
GetMathFilter = znak
End Function
Private Function CheckForShablon(s As String, i As Integer) As String
's значение комбика
'i - номер текущего комбика
Dim n As Long
Dim tmp As String ', tmp2 As String
Dim znak As String
Dim LNL As String    ' " Like " " Not Like "

On Error GoTo err

'если ищем нумерик, поставить вместо like указанный знак и выйти
'есть ли мат знак в начале строки поиска?
znak = GetMathFilter(s)

If Len(znak) > 0 Then
    'вырезать заданный знак
    tmp = Trim$(Replace(s, znak, vbNullString))
End If

If FrmFilter.chNot(i).Value = vbChecked Then
    'Not
    Select Case znak
    Case ">": znak = "<="
    Case "<": znak = ">="
    Case "=": znak = "<>"
    Case ">=": znak = "<"
    Case "<=": znak = ">"
    Case "<>": znak = "="
    End Select

    LNL = " Not Like "
Else
    LNL = " Like "
End If
tmp = Replace2Regional(tmp)
If IsNumeric(tmp) Then
' точка для sql
tmp = Replace(tmp, SeparadorDecimal, ".")

    'поставить знак и выйти
    'CheckForShablon = " " & znak & " '" & tmp & "'"
    CheckForShablon = " " & znak & " " & tmp   'с числом
    'CheckForShablon = " " & znak & " '" & tmp & "'"   'с числом
    Exit Function
End If


'проверка, помещать ли для фильтра строку в звезды
'1 поле целиком. не передавать в звездах, оставить как есть и выйти
If FrmFilter.ChFiltWhole.Value = vbChecked Then CheckForShablon = LNL & "'" & s & "'": Exit Function

'2 в начале поля, звезду в конец и выйти
If FrmFilter.chFiltStart.Value = vbChecked Then CheckForShablon = LNL & "'" & s & "*'": Exit Function

'3 проверить, есть ли sql шаблон в строке, если есть не передавать в звездах
If InStr(s, "*") > 0 Then CheckForShablon = LNL & "'" & s & "'": Exit Function
If InStr(s, "?") > 0 Then CheckForShablon = LNL & "'" & s & "'": Exit Function
n = InStr(s, "[")
If n > 0 Then
    If InStr(n, s, "]") Then
        CheckForShablon = LNL & "'" & s & "'"
    Else
        CheckForShablon = LNL & "'*" & s & "*'"
    End If
Else
    CheckForShablon = LNL & "'*" & s & "*'"
End If

'3 вернуть в звездах
'CheckForShablon =  LNL & "'*" & s & "*'"
Exit Function
err:
Debug.Print err.Description
End Function


Public Sub SQLCompatible(ByRef s As String)
'менять '# на ?
If Len(s) <> 0 Then
   ' If InStr(s, "?") > 0 Then s = Replace(s, "?", "[?]") '1
   ' If InStr(s, "'") > 0 Then s = Replace(s, "'", "?")
   ' If InStr(s, "#") > 0 Then s = Replace(s, "#", "?")
   If InStr(s, "'") > 0 Then s = Replace(s, "'", "''")
End If

End Sub
Public Sub ActOtherFilters(Index As Integer)
'    Set ars = ADB.OpenRecordset("Select * From Acter Where " & tmp)
'Dim strSQL As String

Select Case Index
'Case 0    'all
'    'Set ars = ADB.OpenRecordset("Select * From Acter")
'    Set ars = ADB.OpenRecordset("Acter", dbOpenTable)
'    ars.Index = "KeyAct"
'    FilterActFlag = False
'    ToDebug "Все актеры"
'
'Case 1    'with foto
'    Set ars = ADB.OpenRecordset("Select * From Acter Where Face <> ''")
'
'   ' Set ars = ADB.OpenRecordset("Select * From Acter Where Name not Like '*[a-Z]*'")
'   ' Set ars = ADB.OpenRecordset("Select * From Acter Where Name not Like '*[а-Я]*'")
'    ' Set ars = ADB.OpenRecordset("Select * From Acter Where Name not Like '* *'")
''Set ars = ADB.OpenRecordset("Select * From Acter Where ((Name not Like '*[a-Z]*') and (Name not Like '*[а-Я]*'))")
'
'    FilterActFlag = True
'    ToDebug "Актеры с фото"
'Case 2    'w/out foto
'    Set ars = ADB.OpenRecordset("Select * From Acter Where Face Is Null")
'    FilterActFlag = True
'    ToDebug "Актеры без фото"
'
'Case 3
'    Set ars = ADB.OpenRecordset("Select * From Acter Where Name In (Select Name From Acter Group By Name HAVING Count(Name) > 1)")
'    FilterActFlag = True
'    ToDebug "Актеры дубликаты"
'
'    'work Set ars = ADB.OpenRecordset("Select * From Acter Where Left(Name,10) In (Select Left(Name,10) From Acter Group By Left(Name,10) HAVING Count(Left(Name,10)) > 1)")
'
'
'    'Dim tmp As String
'    'Dim rstemp As DAO.Recordset
'    'Dim i As Integer
'    ''tmp = "Select Left(Name,10) From Acter Group By Left(Name,10)" ' HAVING Count(Left(Name,5)) > 1"
'    'tmp = "Select Left(Name,8) From Acter" ' Group By Name HAVING Count(Left(Name,1)) > 1"
'    'tmp = "Select * From Acter Where Left(Name,8) in (Select Left(Name,8) From Acter Group By Left(Name,8) HAVING Count(Left(Name,8)) > 1)"
'    'Debug.Print tmp
'    'Set rstemp = ADB.OpenRecordset(tmp)
'    'rstemp.MoveFirst: For i = 0 To 10: Debug.Print rstemp(0): rstemp.MoveNext: Next i
'    ''tmp = "Select * From Acter Where Name Like (Select Left$(Name,5) From Acter Group By Name HAVING Count(Name) > 1)"
'    ''Debug.Print tmp
'    ''Set ars = ADB.OpenRecordset(tmp)
'    ''Set ars = ADB.OpenRecordset("Select * From Acter Where Name In (Select Left$(Name,5) From Acter Group By Name HAVING Count(Name) > 1)")
'    ''Set ars = ADB.OpenRecordset("Select * From Acter Where Left$(Name,5) In (Select Name From Acter Group By Name HAVING Count(Name) > 1)")


Case 4    'правая кнопка на разделенных именах
    'выборка по именам
    Dim i As Integer
    Dim tmp As String
    If FrmMain.ListBActHid.SelCount < 1 Then Exit Sub
ToDebug "АФ по именам"
    For i = 0 To FrmMain.ListBActHid.ListCount - 1
        If FrmMain.ListBActHid.Selected(i) Then
            tmp = tmp & " And InStr(name, '" & FrmMain.ListBActHid.List(i) & "') > 0"
        End If
    Next i
    tmp = Right$(tmp, Len(tmp) - 5)    '- левый AND
    Set ars = ADB.OpenRecordset("Select * From Acter Where " & tmp)

    FilterActFlag = True
    'снять точки
'    For i = 0 To 3
'        OptActOnlyFotoHid(i).Value = False
'    Next i

Case 5 'показ всех персон из меню карточки фильма
ToDebug "АФ по персонам фильма"

 FilterActFlag = True
     'снять точки
'    For i = 0 To 3
'        OptActOnlyFotoHid(i).Value = False
'    Next i

End Select

ArsProcess


'If FilterActFlag Then
''OptActOnlyFotoHid(0).BackColor = &HFF&
'FrameActer.ForeColor = &HC0&      '&HFFFF&
'Else
''OptActOnlyFotoHid(0).BackColor = &H8000000F
'FrameActer.ForeColor = &H80000012
'End If
'
'
''встать на 1
'If ars.RecordCount > 0 Then CurAct = 1
'
'If ars.RecordCount = 0 Then
'    'чистка окон актера
'    PicActFoto.Height = 0: PicActFoto.Width = 0    'убирает прокрутку
'    LVActer.ListItems.Clear
'    Set PicActFoto.Picture = Nothing
'    TextActName.Text = vbNullString
'    TextActBio.Text = vbNullString
'    FrameActer.Caption = FrameActerCaption & "0)"
'    ListBActHid.Clear
'    ComActEdit.Enabled = False
'    ComActDel.Enabled = False
'
'    Exit Sub
'End If
'
'ComActEdit.Enabled = True
'ComActDel.Enabled = True
'
''читать в список
'FillActListView
'
''кликнуть
'LVActClick
'
'If LVActer.ListItems.Count > 0 Then LVActer.SelectedItem.EnsureVisible

End Sub


Public Sub FilterMovieWithPers(s As String)
's типа: Director Like 'Актер1' And Director Like 'Актер2' , уже с заменой ''
'вызов из окна актеров () как mFiltAct_Click из карточки
'оставить в лв только фильмы с этим актером
'искать в актерах и режиссерах
'фильтрация  будет отменена, группировку оставим - бестолку - отменяет группировка наш фильтр
'                   при клике показать окно фильмов

'If Len(sPerson) = 0 Then Exit Sub

Dim strSQL As String

Screen.MousePointer = vbHourglass

LastSQLPersonString = "(" & s & ")" 'FiltPersonFlag
strSQL = "SELECT * FROM Storage WHERE " & s

If GroupedFlag Then
    'добавить запрос от группировки
    strSQL = strSQL & " AND (" & LastSQLGroupString & ")"
'Else
'    strSQL = strSQL & ")"
End If

'Debug.Print "mFA:" & strSQL

On Error GoTo err
Set rs = DB.OpenRecordset(strSQL)

FrmMain.FillListView

FrmMain.LActMarkCount.Caption = FrmMain.LActMarkCountCaption & " " & rs.RecordCount
'отменять-то надо If rs.RecordCount > 0 Then
    'свой флаг как mFiltAct_Click
    FiltPersonFlag = True
    FrmMain.ComFilter.BackColor = &HC0C0FF
'End If

Screen.MousePointer = vbNormal
Exit Sub

err:
Screen.MousePointer = vbNormal
ToDebug "Err_mFiltA" '& err.Description
MsgBox msgsvc(46), vbExclamation ': ToDebug err.Description

End Sub
Private Function GetGroupLikeStr(s As String) As String
'Dim tmp As String
Dim strLike As String

SQLCompatible s

'''''''''''''''
'не потрошить дб аналог FillTVGroup
'Case "MovieName", "Label", "Year", "Resolution", "FPS", "CDN", "MediaType", "Audio",
' "Video", "Debtor", "Other", "CoverPath", "MovieURL", "Rating", "FileLen"


If s <> "Null" Then ' filltvgroup
    Select Case GroupField
    Case "MovieName", "Label", "Year", "Resolution", "FPS", "CDN", "MediaType", "Audio", "Video", "Debtor", "Other", "CoverPath", "MovieURL", "Rating" ', "FileLen"
        If Len(s) = 0 Then
            strLike = "(" & GroupField & " = '')"
        Else
            strLike = "(" & GroupField & " = '" & s & "')"
        End If
    Case "FileLen" 'числовой, равенство без кавычек
            strLike = "(" & GroupField & " = " & s & ")"
    Case Else    'FilterLike
        strLike = FilterLikeString(GroupField, s)
    End Select

Else    ' Null
    strLike = "(" & GroupField & " Is Null)"
End If

GetGroupLikeStr = strLike
End Function
Public Sub TVCLICK()
'Группировка, клик по строке
Dim strSQL As String
Dim tmp As String
Dim strLike As String

On Error GoTo errn

If FrmMain.tvGroup.ListItems.Count < 1 Then Exit Sub    'не кликать на пустоту
If Len(GroupField) = 0 Then Exit Sub
If FrmMain.tvGroup.SelectedItem Is Nothing Then Exit Sub    '1

GroupedFlag = True
FrmMain.Timer2.Enabled = False

'много нельзя - длинный запрос ... складывать лайки всех выделенных
'Dim Itm As ListItem
'For Each Itm In FrmMain.tvGroup.ListItems
'    If Itm.Selected Then
        'там каждая GetGroupLikeStr в скобочках
'        strLike = GetGroupLikeStr(Itm.Text) & " Or " & strLike
'    End If
'Next

tmp = FrmMain.tvGroup.ListItems(FrmMain.tvGroup.SelectedItem.Index)
strLike = GetGroupLikeStr(tmp)

If Len(strLike) = 0 Then
    'ничего не выделено
    Exit Sub
Else
    'убрать последний ор (от цикла, если есть - сейчас нет)
    'strLike = Left$(strLike, Len(strLike) - 4)
End If
'запрос от группировки в скобочки
LastSQLGroupString = "(" & strLike & ")"


strSQL = "Select * From Storage Where " & LastSQLGroupString

If FilteredFlag And Len(LastSQLFilterString) <> 0 Then
    'добавить запрос от фильтра
    strSQL = strSQL & " AND " & LastSQLFilterString
End If
If FiltPersonFlag And Len(LastSQLPersonString) <> 0 Then
    'добавить запрос от фильтра по актерам
    strSQL = strSQL & " AND " & LastSQLPersonString
End If

'Debug.Print "ГР: " & strSQL

Set rs = DB.OpenRecordset(strSQL)    ', dbOpenSnapshot)

FrmMain.FillListView

Exit Sub

errn:
ToDebug "Err_TVCl " & Len(strSQL)
'Debug.Print strSQL
MsgBox msgsvc(46), vbExclamation
End Sub

Public Sub FillTVGroup()
'заполнение списка групп

Dim i As Long, j As Long
Dim strSQL As String
Dim SQL1 As String
Dim rsTV As DAO.Recordset
'mzt Dim TempPole As String
Dim pFlag As Boolean    'потрошить или нет
Dim PustoFlag As Boolean 'были ли пустые значения
Dim NullFlag As Boolean 'были ли null значения
Dim rsArr() As String 'массив одиночных значений (и потрошенные)
Dim R() As String
'Dim TokNums As Integer
Dim ArrFlag As Boolean 'флаг начала добавления пунктов в массив rsArr

'сабкласс lv
'ModLVSubClass.UnAttach FrameView.hWnd

On Error Resume Next
If rs Is Nothing Then Exit Sub

'tvGroup.Visible = False
FrmMain.tvGroup.ListItems.Clear
'снимать сортировку, чтоб не мешала Order By
FrmMain.tvGroup.Sorted = False


If Len(GroupField) = 0 Then    '            отмена группировки
    If FilteredFlag And Len(LastSQLFilterString) <> 0 Then    'применить запрос от фильтра
        strSQL = "Select * From Storage Where " & LastSQLFilterString
        Set rs = DB.OpenRecordset(strSQL)
        GroupedFlag = False
        FrmMain.FillListView
        Exit Sub
    ElseIf FiltPersonFlag And Len(LastSQLPersonString) <> 0 Then 'была фильтрация по актеру, применить
        strSQL = "Select * From Storage Where " & LastSQLPersonString
        Set rs = DB.OpenRecordset(strSQL)
        GroupedFlag = False
        FrmMain.FillListView
        Exit Sub

    Else    'вернуть все
        'tvGroup.Visible = True
        FrmMain.tvGroup.Refresh ' - пропадает хедер
        
        If GroupedFlag Then
            strSQL = "Select * From Storage"
            Set rs = DB.OpenRecordset(strSQL)
            GroupedFlag = False
            FrmMain.FillListView
            Exit Sub
        Else
            Exit Sub
        End If
    End If


Else

    'Не разделять Не потрошить некоторые поля (Название, Метка, Год, Формат, к.c, Дисков, Носитель, Видео, Должник, Прим, URL об, URL фильма
    'тоже в TVCLICK
    
    Select Case GroupField
        'не потрошить
        Case "MovieName", "Label", "Year", "Resolution", "FPS", "CDN", "MediaType", "Audio", "Video", "Debtor", "Other", "CoverPath", "MovieURL", "Rating", "FileLen"
            pFlag = False
        Case Else
            pFlag = True    'потрошить
    End Select

'ToDebug "Splitting: " & GroupField & " = " & pFlag

    If pFlag Then
    'нет каунта и, соответственно, нет группировки
        SQL1 = "Select " & GroupField & " From Storage"
        If FilteredFlag And Len(LastSQLFilterString) <> 0 Then    'добавить запрос от фильтра
        
            '!Группировка режет данные до 255 символов
            strSQL = SQL1 & " WHERE " & LastSQLFilterString    '& " Group By " & GroupFieldт
        Else
            strSQL = SQL1    '& " Group By " & GroupField
        End If
        
    Else
        'SQL1 = "Select " & GroupField & ", Count(" & GroupField & ") From Storage"
        SQL1 = "Select " & GroupField & ", Count(*) From Storage" 'Count(*) чтоб Null считать
        
        If FilteredFlag And Len(LastSQLFilterString) <> 0 Then    'добавить запрос от фильтра
            'strSQL = SQL1 & " WHERE " & LastSQLFilterString & " Group By " & GroupField
            
            strSQL = SQL1 & " WHERE " & LastSQLFilterString & " Group By " & GroupField & " Order By " & GroupField '& " Desc"
        Else
        
            strSQL = SQL1 & " Group By " & GroupField & " Order By " & GroupField  '& " Desc"
        
'если точно поле - текст и надо отсортировать как номера, то " Order By Val(" & GroupField & ")" - хрен, если Null
'strSQL = SQL1 & " Group By " & GroupField & " Order By IIf(Label Is Null, Label, Val(Label))"
'strSQL = SQL1 & " Group By " & GroupField & " Order By Int(Label)"
'strSQL = "Select Label, Count(Label) From Storage Group By Label Order By Label"
        End If
    End If
End If    'If Len(GroupField) <> 0

'Debug.Print "FillTV: " & strSQL

Set rsTV = DB.OpenRecordset(strSQL)
If err Then
    Debug.Print "Err FillTVGroup " & err.Description
    ToDebug "Err_FillTVGroup: " & err.Description
    Exit Sub
End If

Screen.MousePointer = vbHourglass
On Error GoTo 0
'Debug.Print rsTV(0).Type

With FrmMain

ReDim rsArr(0)    'заполнение с 0
ArrFlag = False

If Not (rsTV.BOF And rsTV.EOF) Then    'If rsTV.RecordCount > 0 Then
    rsTV.MoveLast: rsTV.MoveFirst

    'прогрессбар
    .PBar.min = 0: .PBar.Max = rsTV.RecordCount
    'TextItemHid.ZOrder 0
    .PBar.ZOrder 0
    .PBar.Value = 0
    
    For i = 1 To rsTV.RecordCount    'Do While Not rsTV.EOF '
        .PBar.Value = i
        If IsNull(rsTV(0)) Then    'проверка первого на null, второй, если есть, - сумма - не null
            'сумма Null только при Count(*)
            If pFlag Then
                NullFlag = True
                'tvGroup.ListItems.Add Text:="Null"
            Else
                .tvGroup.ListItems.Add(Text:="Null").ListSubItems.Add Text:=rsTV(1)
            End If
        Else
            'If rsTV(0) = vbNullString Then
            'tvGroup.ListItems.Add(Text:=vbNullString).ListSubItems.Add Text:=rsTV(1)
            'Else
            'tvGroup.ListItems.Add(Text:=rsTV(0)).ListSubItems.Add Text:=rsTV(1)
            
            If pFlag Then
                'собрать воедино найденные поля для FillGroupArray
                'TempPole = TempPole & rsTV(0) & ","
                'проба собрать все в массив
                
                If Len(rsTV(0)) = 0 Then
                    PustoFlag = True 'была пустышка, потом добавим одну в список
                Else
                    If Tokenize04(rsTV(0), R(), ",;", False) > -1 Then ' False! Пустышек быть не должно. (жанр1, жанр2,)
                        For j = 0 To UBound(R)
                        
                           If ArrFlag Then ReDim Preserve rsArr(UBound(rsArr) + 1)
                            rsArr(UBound(rsArr)) = R(j)
                            ArrFlag = True
                            
                        Next j
                    End If
                End If
            
            Else
                'не потрошили, положить два значения в лист
                .tvGroup.ListItems.Add(Text:=rsTV(0)).ListSubItems.Add Text:=rsTV(1)
            End If
            'End If
        End If
        rsTV.MoveNext
    Next i    'Loop
Else
    'rs пуст
    pFlag = False ' чтоб не показывать пустышку
End If

Set rsTV = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''
'ReDim Preserve rsArr(UBound(rsArr) - 1) 'избыток баунда после цикла - хсним

If pFlag Then
    If UBound(rsArr) > 0 Then

        TriQuickSortString rsArr    'sorts your string array

        remdups rsArr    'removes all duplicates
        
        If PustoFlag Then ReDim Preserve rsArr(UBound(rsArr) + 1): rsArr(UBound(rsArr)) = vbNullString 'добавить пустышку
        If NullFlag Then ReDim Preserve rsArr(UBound(rsArr) + 1): rsArr(UBound(rsArr)) = "Null" 'добавить

        .PBar.min = 0
        If UBound(rsArr) > 0 Then .PBar.Max = UBound(rsArr) Else .PBar.Max = 1
        'TextItemHid.ZOrder 0
        .PBar.ZOrder 0

        For i = 0 To UBound(rsArr)
            .PBar.Value = i
            If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
            .tvGroup.ListItems.Add(i + 1, Text:=rsArr(i)).ListSubItems.Add 'Text:=GetGroupNum(rsArr(I))
        Next i

    Else
           .tvGroup.ListItems.Add(1, Text:=rsArr(0)).ListSubItems.Add 'нужен
            If PustoFlag And (Len(rsArr(0)) <> 0) Then .tvGroup.ListItems.Add(1, Text:=vbNullString).ListSubItems.Add

    End If    'If UBound(rsArr) > 1 Then
End If    'If pFlag Then


.TextItemHid.ZOrder 0
'''''''''''''''''''''''''''''''''''''''''''''''
'примечания:
'при потрошении все сортировано строками
'список становится сортированным (tvGroup.Sorted = True), при сортировке по кол-ву и остается сортированным

.tvGroup.SortKey = 0: .tvGroup.SortOrder = lvwAscending    ': .tvGroup.Sorted = True
.tvGroup.Visible = True
If .tvGroup.ListItems.Count > 0 Then
    Set .tvGroup.SelectedItem = .tvGroup.ListItems(1)
    TVCLICK
    .tvGroup.ColumnHeaders(1).Text = GroupColumnHeader & "<" & .tvGroup.ListItems.Count & ">"    'сколько уникальных
End If

Screen.MousePointer = vbNormal
End With
End Sub

