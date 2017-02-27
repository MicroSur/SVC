Attribute VB_Name = "modLV"
Option Explicit

Public Const nHistory As Integer = 24
Public arrHistoryKeys(24) As String 'movietitle/ключи истории кликнутых фильмов (меню mHist(0-14))
'Public arrHistoryKeys(24) As Long 'movietitle/ключи истории кликнутых фильмов (меню mHist(0-14))
Public arrHistoryTitles(24) As String

Public LVManualClickFlag As Boolean


Public Sub subActFiltCancel()
On Error GoTo err
If Not FilterActFlag Then
    If frmActFiltFlag Then Unload frmActFilt
    Exit Sub
End If

Set ars = ADB.OpenRecordset("Acter", dbOpenTable)
ars.Index = "KeyAct"
FilterActFlag = False
ToDebug "Все актеры"

ArsProcess
Exit Sub

err:
Debug.Print "Err_SAFCancel " & err.Description
End Sub


Public Sub ArsProcess()
'покраска, заполнение списка актеров после запроса
On Error GoTo err

With FrmMain

'встать на 1
If ars.RecordCount > 0 Then CurAct = 1

If ars.RecordCount = 0 Then
    'чистка окон актера
    .PicActFoto.Height = 0: .PicActFoto.Width = 0    'убирает прокрутку
    .LVActer.ListItems.Clear
    Set .PicActFoto.Picture = Nothing
    .TextActName.Text = vbNullString
    .TextActBio.Text = vbNullString
    .FrameActer.Caption = FrameActerCaption & "0)"
    .ListBActHid.Clear
    .ComActEdit.Enabled = False
    .ComActDel.Enabled = False
Else
    .ComActEdit.Enabled = True
    .ComActDel.Enabled = True
    'читать в список
    .FillActListView
    'кликнуть
    .LVActClick
    If .LVActer.ListItems.Count > 0 Then .LVActer.SelectedItem.EnsureVisible
End If

If FilterActFlag Then
    'OptActOnlyFotoHid(0).BackColor = &HFF&
    .FrameActer.ForeColor = &HC0&      '&HFFFF&
    .comActFilt.BackColor = &HC0C0FF
    
    .mPutThisActer.Enabled = False
    
Else
    'OptActOnlyFotoHid(0).BackColor = &H8000000F
    .FrameActer.ForeColor = &H80000012
    .comActFilt.BackColor = &HFFFFFF
End If

End With
Exit Sub

err:
Debug.Print "err_arsproc " & err.Description
End Sub
Public Sub FillLvSubs(ind As Long)

On Error Resume Next 'for null обязательно

'LockWindowUpdate ListView.hwnd
'''саб поля
With FrmMain.ListView.ListItems(ind)
    'CheckNoNullVal(dbLabelInd)

'.SmallIcon = 3

        .SubItems(dbLabelInd) = rs(dbLabelInd)
        .SubItems(dbGenreInd) = rs(dbGenreInd)
        .SubItems(dbYearInd) = rs(dbYearInd)
        .SubItems(dbCountryInd) = rs(dbCountryInd)
        .SubItems(dbDirectorInd) = rs(dbDirectorInd)
        .SubItems(dbActerInd) = rs(dbActerInd)
        .SubItems(dbTimeInd) = rs(dbTimeInd)
        .SubItems(dbResolutionInd) = rs(dbResolutionInd)
        .SubItems(dbFpsInd) = Replace2Regional(rs(dbFpsInd))
        .SubItems(dbVideoInd) = rs(dbVideoInd)
        .SubItems(dbAudioInd) = rs(dbAudioInd)
        .SubItems(dbFileLenInd) = rs(dbFileLenInd)
        .SubItems(dbCDNInd) = rs(dbCDNInd)
        .SubItems(dbFileNameInd) = rs(dbFileNameInd)
        .SubItems(dbDebtorInd) = rs(dbDebtorInd)

    If Opt_Debtors_Colorize Then    'пометка цветом если есть должник
        If .SubItems(dbDebtorInd) <> vbNullString Then 'раскрасить
            .ForeColor = .ForeColor Xor &H4080&
        End If
    End If

        .SubItems(dbsnDiskInd) = rs(dbsnDiskInd)
        .SubItems(dbOtherInd) = rs(dbOtherInd)
        .SubItems(dbSubTitleInd) = rs(dbSubTitleInd)
        .SubItems(dbCoverPathInd) = rs(dbCoverPathInd)
        .SubItems(dbMovieURLInd) = rs(dbMovieURLInd)
        .SubItems(dbMediaTypeInd) = rs(dbMediaTypeInd)
        .SubItems(dbRatingInd) = Replace2Regional(rs(dbRatingInd))
        .SubItems(dbLanguageInd) = rs(dbLanguageInd)

End With
'''''''''''end поля
'LockWindowUpdate 0
lvItemLoaded(ind) = True 'загружено не только название
err.Clear

Exit Sub
'err:
'If err.Number <> 0 Then ToDebug "Error, FillLvSubs, " & err.Description
End Sub


Public Sub LVDragDropMulti(ByRef lvList As ListView, ByVal X As Single, ByVal Y As Single)

    Dim objDrag As ListItem
    Dim objDrop As ListItem
    Dim objNew As ListItem
    Dim objSub As ListSubItem
    Dim intIndex As Integer
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim intSelected As Integer
    Dim arrItems() As ListItem
    
    'Retrieve the original items
    Set objDrop = lvList.HitTest(X, Y)
    Set objDrag = lvList.SelectedItem
    If (objDrop Is Nothing) Or (objDrag Is Nothing) Then
        Set lvList.DropHighlight = Nothing
        Set objDrop = Nothing
        Set objDrag = Nothing
        Exit Sub
    End If
    
    'Retrieve the drop position
    intIndex = objDrop.Index
    intCount = lvList.ListItems.Count
    intSelected = 0
    'Remove the drop highlighting
    Set lvList.DropHighlight = Nothing

    'Loop through and retrieve the selected items
    For intLoop = 1 To intCount
        If lvList.ListItems(intLoop).Selected Then
            intSelected = intSelected + 1
            ReDim Preserve arrItems(1 To intSelected) As ListItem
            Set arrItems(intSelected) = lvList.ListItems(intLoop)
        End If
    Next
    'Loop through in reverse and remove the selected items
    'Going in reverse prevents index shifting
    For intLoop = UBound(arrItems) To LBound(arrItems) Step -1
        lvList.ListItems.Remove arrItems(intLoop).Index
    Next
    'Loop through again and add the items back
    'Going in reverse keeps the items in order
    For intLoop = UBound(arrItems) To LBound(arrItems) Step -1
        Set objDrag = arrItems(intLoop)
        'Add it back into the dropped position
        Set objNew = lvList.ListItems.Add(intIndex, objDrag.Key, objDrag.Text, objDrag.Icon, objDrag.SmallIcon)
        'Copy the original subitems to the new item
        If objDrag.ListSubItems.Count > 0 Then
            For Each objSub In objDrag.ListSubItems
                objNew.ListSubItems.Add objSub.Index, objSub.Key, objSub.Text, objSub.ReportIcon, objSub.ToolTipText
            Next
        End If
        objNew.Selected = True
    Next
    
    'Destroy all objects
    ReDim arrItems(1)
    Set arrItems(1) = Nothing
    Set objNew = Nothing
    Set objDrag = Nothing
    Set objDrop = Nothing

End Sub

Public Sub LVDragDropSingle(ByRef lvList As ListView) ', ByVal x As Single, ByVal y As Single)
' переделано для multi
Dim itmX As ListItem
Dim objDrag As ListItem
Dim objNew As ListItem
'mzt Dim objSub As ListSubItem
'mzt Dim intIndex As Integer

'Retrieve the original items
'Set objDrop = lvList.HitTest(x, y)
Set objDrag = lvList.SelectedItem
If objDrag Is Nothing Then
    Set lvList.DropHighlight = Nothing
    Set objDrag = Nothing
    Exit Sub
End If

'Retrieve the drop position
'intIndex = objDrop.Index

'Remove the dragged item
'''lvList.ListItems.Remove objDrag.Index

'открыть базу куда дропать
If OpenDBdd Then
    Screen.MousePointer = vbHourglass
    For Each itmX In FrmMain.ListView.ListItems

        If itmX.Selected Then
            'If MultiSel Then 'флаг терялся кудато
                RSGoto itmX.Key
            'End If
            
            'запись в базу dd
            'rsdd.AddNew 'в GetPutDD
            GetPutDD
            'rsdd.Update 'в GetPutDD

        End If
    Next
    
    Screen.MousePointer = vbNormal

    ' закрыть базу куда дропать
    rsdd.Close: DBdd.Close
    Set rsdd = Nothing: Set DBdd = Nothing

End If    'open

'Destroy all objects
Set objNew = Nothing
Set objDrag = Nothing
'Set objDrop = Nothing
Set lvList.DropHighlight = Nothing

RestoreBasePos

End Sub
Public Sub GetPutDD()
Dim i As Integer

'положим поля открытой базы в базу драг-дропа
'Dim tmp As String
'Dim strSQL As String
'Dim rstemp As DAO.Recordset
'Dim FoundInDD As Boolean
'Dim ddKey As Long

On Error Resume Next

'если найдено такоеже название, обновить пустые поля фильма данными из копируемой базы
'работает, но как-то не нужно
'tmp = CheckNoNull("MovieName")
'If Len(tmp) > 0 Then
'    'есть ли такая запись в дд базе
'    strSQL = "Select MovieName, Key From Storage Where MovieName = '" & tmp & "'"
'    Set rstemp = DBdd.OpenRecordset(strSQL)
'    If rstemp.RecordCount > 0 Then
'        rstemp.MoveFirst 'только первое вхождение
'        If Not IsNull(rstemp(0)) Then
'            FoundInDD = True
'            ddKey = rstemp(1)
'        End If
'    End If
'End If

'If FoundInDD Then
''позиционироваться на ключе и отредактировать запись в дд базе
'    If RSGotoDD(ddKey) Then
'        rsdd.Edit
'
'        If rsdd("Label").FieldSize = 0 Then rsdd("Label") = CheckNoNull("Label")
'        If rsdd("Genre").FieldSize = 0 Then rsdd("Genre") = CheckNoNull("Genre")
'        If rsdd("Year").FieldSize = 0 Then rsdd("Year") = CheckNoNull("Year")
'        If rsdd("Country").FieldSize = 0 Then rsdd("Country") = CheckNoNull("Country")
'        If rsdd("Director").FieldSize = 0 Then rsdd("Director") = CheckNoNull("Director")
'        If rsdd("Acter").FieldSize = 0 Then rsdd("Acter") = CheckNoNull("Acter")
'        If rsdd("Time").FieldSize = 0 Then rsdd("Time") = CheckNoNull("Time")
'        If rsdd("Resolution").FieldSize = 0 Then rsdd("Resolution") = CheckNoNull("Resolution")
'        If rsdd("Audio").FieldSize = 0 Then rsdd("Audio") = CheckNoNull("Audio")
'        If rsdd("FPS").FieldSize = 0 Then rsdd("FPS") = CheckNoNull("FPS")
'        If rsdd("FileLen").FieldSize = 0 Then rsdd("FileLen") = Val(CheckNoNull("FileLen"))
'        If rsdd("CDN").FieldSize = 0 Then rsdd("CDN") = CheckNoNull("CDN")
'        If rsdd("MediaType").FieldSize = 0 Then rsdd("MediaType") = CheckNoNull("MediaType")
'        If rsdd("Video").FieldSize = 0 Then rsdd("Video") = CheckNoNull("Video")
'        If rsdd("SubTitle").FieldSize = 0 Then rsdd("SubTitle") = CheckNoNull("SubTitle")
'        If rsdd("Language").FieldSize = 0 Then rsdd("Language") = CheckNoNull("Language")
'        If rsdd("Rating").FieldSize = 0 Then rsdd("Rating") = CheckNoNull("Rating")
'        If rsdd("FileName").FieldSize = 0 Then rsdd("FileName") = CheckNoNull("FileName")
'        If rsdd("Debtor").FieldSize = 0 Then rsdd("Debtor") = CheckNoNull("Debtor")
'        If rsdd("snDisk").FieldSize = 0 Then rsdd("snDisk") = CheckNoNull("snDisk")
'        If rsdd("Other").FieldSize = 0 Then rsdd("Other") = CheckNoNull("Other")
'        If rsdd("CoverPath").FieldSize = 0 Then rsdd("CoverPath") = CheckNoNull("CoverPath")
'        If rsdd("MovieURL").FieldSize = 0 Then rsdd("MovieURL") = CheckNoNull("MovieURL")
'        If rsdd("Annotation").FieldSize = 0 Then rsdd("Annotation") = CheckNoNull("Annotation")
'        'If rsdd("Checked").FieldSize = 0 Then rsdd("Checked") = CheckNoNull("Checked")
'        If rsdd("SnapShot1").FieldSize = 0 Then rsdd("SnapShot1") = rs("SnapShot1")
'        If rsdd("SnapShot2").FieldSize = 0 Then rsdd("SnapShot2") = rs("SnapShot2")
'        If rsdd("SnapShot3").FieldSize = 0 Then rsdd("SnapShot3") = rs("SnapShot3")
'        If rsdd("FrontFace").FieldSize = 0 Then rsdd("FrontFace") = rs("FrontFace")
'
'    End If
'Else

'просто добавить новую запись
rsdd.AddNew

For i = 0 To rs.Fields.Count - 1

    Select Case LCase$(rs(i).name)
    Case "key", "checked"
    Case Else
        rsdd(i) = rs(i)
    End Select

Next i

'End If 'FoundInDD

rsdd.Update
'err.Clear
End Sub

Public Sub LVSOrt(i As Integer)
'берет с 1
'работает с -1

Dim ind As Long

ind = i - 1
'ToDebug "SortBaseField: " & ind


Select Case ind
Case dbTimeInd  'time
    SortByDates ind
Case dbYearInd, dbFileLenInd, dbCDNInd, dbRatingInd, lvIndexPole
    SortByNumber ind, FrmMain.ListView

Case dbLabelInd
    If Opt_SortLabelAsNum Then
        SortByNumber ind, FrmMain.ListView
    Else
        SortByString ind
    End If

Case Else
    SortByString ind
End Select

End Sub


Public Sub SortByCheck(lngIndex As Integer, Optional NoChangeOrder As Boolean)
'Sort Numerically
Dim l As Long
Dim strData() As String

ToDebug "SortCheck."

With FrmMain.ListView
    If FirstActivateFlag Then .SortOrder = LVSortOrder
    If Not NoChangeOrder Then .SortOrder = (.SortOrder + 1) Mod 2
    'LVSortOrder = ListView.SortOrder

    'strFormat = String(30, "0") & "." & String(30, "0")

    ' Loop through the values in this column. Re-format the values so as they
    ' can be sorted alphabetically, having already stored their visible
    ' values in the tag, along with the tag's original value

    With .ListItems
        For l = 1 To .Count
            With .Item(l)
                .Tag = .Text & vbNullChar & .Tag
                .Text = .Checked
            End With
        Next l
    End With

    'Sort the list alphabetically by this column
    '.SortOrder = (.SortOrder + 1) Mod 2
    .SortKey = lngIndex
    .Sorted = True

    ' Restore the previous values to the 'cells' in this
    ' column of the list from the tags, and also restore
    ' the tags to their original values

    With .ListItems
        For l = 1 To .Count
            With .Item(l)
                strData = Split(.Tag, vbNullChar)
                .Text = strData(0)
                .Tag = strData(1)
            End With
        Next l
    End With

End With
End Sub

Public Sub SortByDates(lngIndex As Long)
' Sort by date.
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String

With FrmMain.ListView
            strFormat = "YYYYMMDDHhNnSs"
        
            ' Loop through the values in this column. Re-format
            ' the dates so as they can be sorted alphabetically,
            ' having already stored their visible values in the
            ' tag, along with the tag's original value
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & vbNullChar & .Tag
                            If IsDate(.Text) Then
                                .Text = Format$(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = vbNullString
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & vbNullChar & .Tag
                            If IsDate(.Text) Then
                                .Text = Format$(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = vbNullString
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Sort the list alphabetically by this column
            
'            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = lngIndex
            .Sorted = True
            
            ' Restore the previous values to the 'cells' in this
            ' column of the list from the tags, and also restore
            ' the tags to their original values
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, vbNullChar)
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, vbNullChar)
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With '.ListItems

End With 'listview
End Sub

Public Sub SortByNumber(lngIndex As Long, LV As Object)

           ' Sort Numerically
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String

On Error GoTo Errr



If LV.ListItems.Count < 1 Then Exit Sub
'Должны быть все сабы добавлены
'If .ListItems.Item(1).ListSubItems.Count < 1 Then Exit Sub ' (lngIndex)

With LV
            strFormat = String(30, "0") & "." & String(30, "0")
        
            ' Loop through the values in this column. Re-format the values so as they
            ' can be sorted alphabetically, having already stored their visible
            ' values in the tag, along with the tag's original value
        
            With .ListItems
                If (lngIndex > 0) Then 'сабы
                    For l = 1 To .Count
'If .Item(l).ListSubItems.Count > 0 Then
'Debug.Print tvGroup.ListItems.Item(l).ListSubItems(lngIndex)
'Debug.Print tvGroup.ListItems.Count

                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & vbNullChar & .Tag
                            If IsNumeric(.Text) Then
                            
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format$(CDbl(.Text), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format$(0 - CDbl(.Text), _
                                        strFormat))
                                End If
                            Else
                                .Text = vbNullString
                            End If
                        End With
'End If
                    Next l
                Else 'айтемы
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & vbNullChar & .Tag
                            If IsNumeric(.Text) Then
                            
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format$(CDbl(.Text), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format$(0 - CDbl(.Text), _
                                        strFormat))
                                End If
                            Else
                                .Text = vbNullString
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Sort the list alphabetically by this column
            
'            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = lngIndex
            .Sorted = True
            
            ' Restore the previous values to the 'cells' in this
            ' column of the list from the tags, and also restore
            ' the tags to their original values
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
'If .Item(l).ListSubItems.Count > 0 Then
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, vbNullChar)
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
'End If
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, vbNullChar)
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
'.Sorted = False хорошо если только числа
'.Sorted = False

End With
Exit Sub
Errr:
'Debug.Print err.Description
End Sub

Public Sub SortByString(lngIndex As Long)
'sort how listview can
With FrmMain.ListView
'            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = lngIndex
            .Sorted = True
End With
End Sub
'****************************************************************
' InvNumber
' Function used to enable negative numbers to be sorted
' alphabetically by switching the characters
'----------------------------------------------------------------

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function
Public Sub ReloadLVHeaders(hi As Integer)
'hi - сколько полей
Dim i As Integer
'hi = 1
'Очистка хедеров списка
With FrmMain.ListView
    For i = .ColumnHeaders.Count To 1 Step -1: .ColumnHeaders.Remove i: Next i
    'Хедеры списка
    For i = 1 To hi: .ColumnHeaders.Add: Next
End With
End Sub
Public Sub DelLVItems(ch As Boolean)
'удалить из списка
Dim i As Long

'ListView.Visible = False
ToDebug "DelFromList"

With FrmMain.ListView
    If ch Then
        If CheckCount = .ListItems.Count Then
            .ListItems.Clear
        Else
            For i = .ListItems.Count To 1 Step -1
                If .ListItems(i).Checked Then
                    .ListItems.Remove i
                End If
            Next i
        End If
    Else    'sel
        If SelCount = .ListItems.Count Then
            .ListItems.Clear
        Else
            For i = .ListItems.Count To 1 Step -1
                If .ListItems(i).Selected Then
                    .ListItems.Remove i
                End If
            Next i
        End If

    End If

    'ListView.Visible = True

    'переписать все индекс поля в LV ? а если сортировано...
    .Sorted = False
    For i = 1 To .ListItems.Count
        .ListItems(i).SubItems(lvIndexPole) = i - 1
    Next i

    'если была сортировка - произвести ее
    If Opt_SortLVAfterEdit Then
        If LVSortColl > 0 Then LVSOrt (LVSortColl)
        If LVSortColl = -1 Then SortByCheck 0
    End If

End With

FrmMain.LVCLICK
End Sub


Public Sub DelMovies(ch As Boolean)
'удалить из базы
If BaseReadOnly Or BaseReadOnlyU Then
    'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
    Exit Sub
End If

Dim ret As VbMsgBoxResult
Dim i As Long
Dim SQLstrDel As String
Dim strSQL As String
Dim SelectString As String
Dim tOld As Boolean

FrmMain.Timer2.Enabled = False

If ch Then
    'checked
    ret = myMsgBox(msgsvc(45) & vbCrLf & vbCrLf & "(" & CheckCount & ")", vbYesNo, , FrmMain.hwnd)
    Select Case ret

    Case vbYes        'удалить записи
        rs.Close
        'If CheckCount = rsRecordCount Then ' нельзя изза фильтров
        'SQLstrDel = "DELETE * FROM Storage"
        SQLstrDel = "DELETE FROM Storage Where Checked = '1'"
        DB.Execute (SQLstrDel)
        ToDebug "DelChecked"
    Case Else        'no
        Exit Sub
    End Select

Else    'selected

    ret = myMsgBox(msgsvc(44) & vbCrLf & vbCrLf & "(" & SelCount & ")", vbYesNo, , FrmMain.hwnd)
    Select Case ret
    Case vbYes        'удалить записи

        rs.Close
        'If SelCount = rsRecordCount Then ' нельзя изза фильтров
        'SQLstrDel = "DELETE * FROM Storage"
        For i = 0 To UBound(SelRowsKey) - 1    'для каждого ключа , долго, но надо не ошибится
            SQLstrDel = "DELETE FROM Storage Where Key = " & Val(SelRowsKey(i + 1))
            DB.Execute (SQLstrDel)
        Next i
        ToDebug "DelSelected"
    Case Else        'no
        Exit Sub
    End Select

End If

OpenRS

If ch Then
    DelLVItems True    'del selected in lv
Else
    DelLVItems False    'del selected
End If

'и применить фильтры
If FilteredFlag Or GroupedFlag Then

    SelectString = "SELECT * FROM Storage WHERE ("

    If FilteredFlag And GroupedFlag Then
        strSQL = SelectString & LastSQLFilterString & " AND " & LastSQLGroupString & ")"
    Else
        If FilteredFlag Then    'повторить запрос-группировку
            strSQL = SelectString & LastSQLFilterString & ")"
        ElseIf GroupedFlag Then
            strSQL = SelectString & LastSQLGroupString & ")"
        End If
    End If

    'Debug.Print "FAD: " & strSQL

    On Error GoTo err
    Screen.MousePointer = vbHourglass
    Set rs = DB.OpenRecordset(strSQL)
    'FilteredFlag = True    'флаг неполного показа
    FrmMain.FillListView
    'FrmMain.ComFilter.BackColor = &HC0C0FF

    Screen.MousePointer = vbNormal

End If

'LVCLICK in DelLVItems
FrmMain.FrameView.Caption = FrameViewCaption & " " & FrmMain.ListView.ListItems.Count & " )"    'нестыкуется с группами

Exit Sub

err:
Screen.MousePointer = vbNormal
ToDebug "Err_DelMo:" & err.Description
End Sub

Public Sub lvwAutoSizeColumns(lvw As MSComctlLib.ListView, WithHeaders As Boolean)
Dim c As MSComctlLib.ColumnHeader
For Each c In lvw.ColumnHeaders
'PostMessage lvw.hWnd, LVM_FIRST + 30, C.Index - 1, -1
If WithHeaders Then
PostMessage lvw.hwnd, LVM_SETCOLUMNWIDTH, c.Index - 1, ByVal LVSCW_AUTOSIZE_USEHEADER
Else
PostMessage lvw.hwnd, LVM_SETCOLUMNWIDTH, c.Index - 1, -1
End If
Next
lvw.Refresh
End Sub
Public Function GotoLV(lvk As String) As Long
'возвращать индекс lv, по ключу
'lvk с кавычками на конце
Dim i As Long
With FrmMain.ListView
    GotoLV = -1
    For i = 1 To .ListItems.Count
        If .ListItems(i).Key = lvk Then
            GotoLV = i
            Exit For
        End If
    Next i
End With
End Function
Public Function GotoLVLong(lvk As Long) As Long
'возвращать индекс lv, по ключу
Dim i As Long
With FrmMain.ListView
    GotoLVLong = -1
    For i = 1 To .ListItems.Count
        If Val(.ListItems(i).Key) = lvk Then
            GotoLVLong = i
            Exit For
        End If
    Next i
End With
End Function
Public Sub GotoLVAct(lvk As String)
'Встаем на задунную ключем строку списка
Dim i As Long
With FrmMain.LVActer
    CurAct = 1
    For i = 1 To .ListItems.Count
        If .ListItems(i).Key = lvk Then
            CurAct = i
            Exit For
        End If
    Next i
    Set .SelectedItem = .ListItems(CurAct)
End With
End Sub
Public Sub MakeDupCurrent()
'ctrl+w
'сделать дубликат текущей строки
'ListView.SelectedItem.Key = 91"

If BaseReadOnly Or BaseReadOnlyU Then
    'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
    Exit Sub
End If

Dim strSQL As String
Dim rsTmp As DAO.Recordset
Dim i As Integer
Dim curKey As String

On Error GoTo err
If FrmMain.ListView.SelectedItem Is Nothing Then Exit Sub

With FrmMain.ListView
    strSQL = "Select * From Storage Where Key = " & Val(.SelectedItem.Key)
    Set rsTmp = DB.OpenRecordset(strSQL)

    If rsTmp.RecordCount = 1 Then

        rs.AddNew

        curKey = rs("Key")
        ToDebug "MDCur_Key=" & curKey

        For i = 0 To rs.Fields.Count - 1

            If LCase$(rs(i).name) = "key" Then
                'Debug.Print "s"
            Else
                rs(i) = rsTmp(i)
            End If

        Next i
        rs.Update
    End If

    'добавить строку в LV
    RSGoto curKey

    .Sorted = False

    ReDim Preserve lvItemLoaded(.ListItems.Count + 1)    ' 1
    Add2LV .ListItems.Count, .ListItems.Count + 1    '2

    CurLVKey = rs("Key") & """"
    CurSearch = GotoLV(CurLVKey)
    'пометить
    If .ListItems.Count > 0 Then
        Set .SelectedItem = .ListItems(CurSearch)
    End If

    'если была сортировка - произвести ее
    If Opt_SortLVAfterEdit Then
        If LVSortColl > 0 Then LVSOrt (LVSortColl)
        If LVSortColl = -1 Then SortByCheck 0, True
    End If

    DoEvents
    If FrmMain.FrameView.Visible Then .SelectedItem.EnsureVisible
    FrmMain.FrameView.Caption = FrameViewCaption & " " & .ListItems.Count & " )"


    Set rsTmp = Nothing
    ToDebug "CtrlW ok"

End With
Exit Sub
err:
Set rsTmp = Nothing
ToDebug "Ctrl+w " & err.Description
End Sub

Public Function Add2LV(indPole As Long, indLV As Long) As Boolean
If rs.RecordCount < 1 Then Exit Function
'indPole - значение поля индекс
'indLV - Куда по индексу LV
On Error GoTo err
'Call SendMessage(FrmMain.ListView.hwnd, WM_SETREDRAW, False, ByVal 0&)
With FrmMain.ListView
    .ListItems.Add indLV, rs("Key") & """", CheckNoNullVal(0)
    'Debug.Print "Добавили CurSearch =" & CurSearch
    'Debug.Print "всего " & ListView.ListItems.Count
    'Index
    .ListItems(indLV).SubItems(lvIndexPole) = indPole
    .ListItems(indLV).Checked = Val(CheckNoNullVal(dbCheckedInd))
    'Subs
    If Opt_LoadOnlyTitles = False Then    'не только названия
        FillLvSubs indLV
    End If
    ToDebug "A2LV_Ok"
    Add2LV = True

End With
'Call SendMessage(FrmMain.ListView.hwnd, WM_SETREDRAW, True, ByVal 0&)
Exit Function

err:
ToDebug "Err_Add2LV, " & err.Description & ", Key=" & rs("Key")
Debug.Print "Err_Add2LV, " & err.Description
'myMsgBox msgsvc(36), vbCritical, , Me.hwnd
Add2LV = False
End Function

Public Function EditLV(indLV As Long) As Boolean

If rs.RecordCount < 1 Then Exit Function

err.Clear
On Error GoTo err

With FrmMain.ListView

    .ListItems(indLV).Text = CheckNoNullVal(0)
    'Subs
    If Opt_LoadOnlyTitles = False Then    'не только названия
        FillLvSubs indLV
    End If

    ToDebug " LV_updated"
    EditLV = True

End With
Exit Function

err:
ToDebug "Err_EditLV, " & err.Description & ", Key=" & rs("Key")
'myMsgBox msgsvc(36), vbCritical, , Me.hwnd
EditLV = False
End Function
Public Function MenuActSelect(s As String) As Boolean
'поиск  в списке базы актеров
'вызывается после выделения

Dim word As String
Dim itmX As ListItem
Dim fotoFoundFlag As Boolean
Dim troetoch As String

If Len(s) > 40 Then Exit Function    'актер не больше 40 символов

If Len(s) > 3 Then
    'Пометить актеров, если они есть в базе
    FrmMain.mnuShowThisActer.Enabled = False

    'заполнить список актеров
    If FrmMain.LVActer.ListItems.Count < 1 Then
        If abdname <> vbNullString Then
            If Not LVActerFilled Then FrmMain.FillActListView
        Else
            Exit Function
        End If
    End If
    If FrmMain.LVActer.ListItems.Count < 1 Then Exit Function

    word = Trim$(s)
    FrmMain.mnuShowThisActer.Caption = word        ' иначе бы вышли давно Left$(word, actlen)

    UCLVShowPersonFlag = False
    If Not FilterActFlag Then FrmMain.mPutThisActer.Enabled = True    'можно добавлять его в базу
    For Each itmX In FrmMain.LVActer.ListItems
        If InStr(1, itmX.Text, word, vbTextCompare) <> 0 Then
            'нашли персону
            UCLVShowPersonFlag = True
            FrmMain.mnuShowThisActer.Enabled = True
            FrmMain.mPutThisActer.Enabled = False    'не надо добавлять его в базу
            'CurAct = itmX.Index
            ToActFromLV = itmX.Index


            'Debug.Print "Up " & itmX.Key
            'запросить картинку по ключу?
            RSGotoAct itmX.Key
            'PicActFoto в uclv
            'фотку актера
            If Opt_UCLV_Vis And Opt_UCLVPic_Vis Then
                If GetPic(FrmMain.PicTempHid(1), 2, "Face") Then
                    'ResizeWIA FrmMain.PicTempHid(1), FrmMain.UCLV.Controls("picUCLV").ScaleWidth, FrmMain.UCLV.Controls("picUCLV").ScaleHeight, aratio:=True
                    ResizeWIA FrmMain.PicTempHid(1), FrmMain.UCLV.Controls("picUCLV").ScaleHeight, FrmMain.UCLV.Controls("picUCLV").ScaleHeight, aratio:=True
                    FrmMain.UCLV.Controls("picUCLV").Width = FrmMain.PicTempHid(1).Width
                    FrmMain.UCLV.Controls("picUCLV").Picture = FrmMain.PicTempHid(1)
                    fotoFoundFlag = True
                Else    'нет картинки
                    fotoFoundFlag = False
                End If
                'FrmMain.UCLV.Correct_tBIO ' в PutCoverUCLV

                'био
                FrmMain.UCLV.Controls("tBIO").Text = vbNullString
                If ars.Fields("Name") <> vbNullString Then
                    FrmMain.UCLV.Controls("tBIO").Text = ars.Fields("Name") & vbCrLf
                End If
                If ars.Fields("Bio") <> vbNullString Then
                    If Len(ars.Fields("Bio")) > 250 Then troetoch = "..." Else troetoch = vbNullString
                    FrmMain.UCLV.Controls("tBIO").Text = FrmMain.UCLV.Controls("tBIO").Text & left$(ars.Fields("Bio"), 250) & troetoch
                    FrmMain.UCLV.Controls("tBIO").Visible = True
                Else
                    FrmMain.UCLV.Controls("tBIO").Text = vbNullString
                End If

            End If

            Exit For

        End If
    Next

    MenuActSelect = fotoFoundFlag


End If        'TextItemHid.SelLength > 1

End Function
Public Sub PutCoverUCLV(fotoFound As Boolean)
If Not fotoFound And Opt_UCLVPic_Vis Then    'не нашли актера
'ковер
    If GetPic(FrmMain.PicTempHid(1), 1, "FrontFace") Then
        'ResizeWIA FrmMain.PicTempHid(1), FrmMain.UCLV.Controls("picUCLV").ScaleWidth, FrmMain.UCLV.Controls("picUCLV").ScaleHeight, aratio:=True
        ResizeWIA FrmMain.PicTempHid(1), FrmMain.UCLV.Controls("picUCLV").ScaleHeight, FrmMain.UCLV.Controls("picUCLV").ScaleHeight, aratio:=True

        FrmMain.UCLV.Controls("picUCLV").Width = FrmMain.PicTempHid(1).Width
        FrmMain.UCLV.Controls("picUCLV").Picture = FrmMain.PicTempHid(1).Picture
    Else
        FrmMain.UCLV.Controls("picUCLV").Picture = Nothing
    End If

    If Not UCLVShowPersonFlag Then FrmMain.UCLV.Controls("tBIO").Visible = False
End If
FrmMain.UCLV.Correct_tBIO
End Sub
Public Sub MNU_DOLG(SelFlag As Boolean)
'false = Check
If rs Is Nothing Then Exit Sub
If BaseReadOnly Or BaseReadOnlyU Then
    'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
    Exit Sub
End If

Dim Dolg As String
Dim itmX As ListItem
Dim itmX_CheckSel As Boolean    'true если работаем с Selected
Dim Debt() As String
Dim i As Integer, j As Integer
Dim DefVal As String

On Error GoTo err


DefVal = CheckNoNull("Debtor")

If SelFlag Then    'Selected
    ToDebug "должник_sel"
    Dolg = myInputBox(FrmMain.ListView.ColumnHeaders(dbDebtorInd + 1).Text & vbCrLf & FrmMain.mnuLVSelected.Caption, , FrmMain.hwnd, sDefault:=DefVal)
Else    'checked
    ToDebug "должник_ch"
    Dolg = myInputBox(FrmMain.ListView.ColumnHeaders(dbDebtorInd + 1).Text & vbCrLf & FrmMain.mnuLVChecked.Caption, , FrmMain.hwnd, sDefault:=DefVal)
End If

If StrPtr(Dolg) = 0 Then Exit Sub    'cancel

'+ дата
If Len(Dolg) <> 0 Then Dolg = Dolg & " (" & Date & ")"

With FrmMain
    i = 0
    ReDim Debt(0) As String
    For Each itmX In .ListView.ListItems

        itmX.ForeColor = LVFontColor    ' убрать цвет должников

        If SelFlag Then
            itmX_CheckSel = itmX.Selected
        Else
            itmX_CheckSel = itmX.Checked
        End If

        ' заполнять массив меток диска
        If itmX_CheckSel Then
            i = i + 1
            ReDim Preserve Debt(i)
            Debt(i) = itmX.SubItems(dbLabelInd)
        End If

        If itmX_CheckSel Then

            RSGoto itmX.Key

            rs.Edit
            rs.Fields(dbDebtorInd) = Dolg
            rs.Update

            'Debug.Print itmX.SubItems(16)
            itmX.SubItems(dbDebtorInd) = Dolg    '17-1=16
            'Debug.Print rs.Fields(0)
        End If
    Next

    If Opt_LoanAllSameLabels Then
        ' заново для общих меток дисков
        For Each itmX In .ListView.ListItems
            ' есть ли еще фильмы с такими метками
            itmX_CheckSel = False
            For j = 1 To UBound(Debt)
                If Len(itmX.SubItems(dbLabelInd)) <> 0 Then    'не пустые метки
                    If Debt(j) = itmX.SubItems(dbLabelInd) Then
                        itmX_CheckSel = True
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next j


            If itmX_CheckSel Then

                'rs.MoveFirst: rs.Move itmX.ListSubItems(lvIndexPole)
                'RSMOVE itmX.ListSubItems(lvIndexPole), "MNU_DOLG"
                RSGoto itmX.Key

                rs.Edit
                rs.Fields(dbDebtorInd) = Dolg
                rs.Update

                'Debug.Print itmX.SubItems(16)
                itmX.SubItems(dbDebtorInd) = Dolg    '17-1=16
                'Debug.Print rs.Fields(0)
            End If

        Next
    End If

    'пометка цветом если есть должник
    If Opt_Debtors_Colorize Then
        For Each itmX In .ListView.ListItems
            If Len(itmX.SubItems(dbDebtorInd)) = 0 Then
                itmX.ForeColor = LVFontColor
            Else
                itmX.ForeColor = itmX.ForeColor Xor &H4080&
            End If
        Next
    End If

    'восстановить текущ. поз. в базе
    RestoreBasePos

End With
Exit Sub

err:
ToDebug "Err_mdlg:" & err.Description
End Sub

Public Sub MNU_Label(SelFlag As Boolean)
'false = Check
If rs Is Nothing Then Exit Sub
If BaseReadOnly Or BaseReadOnlyU Then
    'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
    Exit Sub
End If

Dim NewLabel As String
Dim itmX As ListItem
Dim itmX_CheckSel As Boolean    'true если работаем с Selected
'Dim Debt() As String
Dim i As Integer, j As Integer
Dim DefVal As String

On Error GoTo err


DefVal = CheckNoNull("Label")

If SelFlag Then    'Selected
    ToDebug "Изменение меток выделенных"
    NewLabel = myInputBox(FrmMain.ListView.ColumnHeaders(dbLabelInd + 1).Text & vbCrLf & FrmMain.mnuLVSelected.Caption, , FrmMain.hwnd, sDefault:=DefVal)
Else    'checked
    ToDebug "Изменение меток помеченных"
    NewLabel = myInputBox(FrmMain.ListView.ColumnHeaders(dbLabelInd + 1).Text & vbCrLf & FrmMain.mnuLVChecked.Caption, , FrmMain.hwnd, sDefault:=DefVal)
End If

If StrPtr(NewLabel) = 0 Then Exit Sub    'cancel

With FrmMain
    i = 0
    For Each itmX In .ListView.ListItems
        If SelFlag Then
            itmX_CheckSel = itmX.Selected
        Else
            itmX_CheckSel = itmX.Checked
        End If

        If itmX_CheckSel Then

            'rs.MoveFirst: rs.Move itmX.ListSubItems(lvIndexPole)
            'RSMOVE itmX.ListSubItems(lvIndexPole), "MNU_Label"
            RSGoto itmX.Key

            rs.Edit
            rs.Fields(dbLabelInd) = NewLabel
            rs.Update

            'Debug.Print itmX.SubItems(16)
            'вписать в ячейку
            itmX.SubItems(dbLabelInd) = NewLabel
            'Debug.Print rs.Fields(0)
        End If
    Next

    'восстановить текущ. поз. в базе
    RestoreBasePos

End With
Exit Sub

err:
ToDebug "Err_mlab:" & err.Description
End Sub


Public Function ChangeLVHSortMark(s As String) As String
Dim ColHead As ColumnHeader
'убрать метку <, > ">"
With FrmMain
    For Each ColHead In .ListView.ColumnHeaders
        If right$(ColHead.Text, 2) = " >" Or right$(ColHead.Text, 2) = " <" Then ColHead.Text = left$(ColHead.Text, Len(ColHead.Text) - 2)
    Next
    If right$(s, 2) = " >" Or right$(s, 2) = " <" Then s = left$(s, Len(s) - 2)

    If LVSortColl > 0 Then
        'установить
        If .ListView.SortOrder = lvwAscending Then
            ChangeLVHSortMark = s & " >"
        Else
            ChangeLVHSortMark = s & " <"
        End If
    End If
End With
End Function


Public Function FindNextLV(ind As Integer, ftext As String) As Integer
'для поиска в главном окне
Dim itmX As ListItem
Dim itmRet As ListItem
Dim i As Long    ', j As Integer
Dim temp As String

If FrmMain.ListView.SelectedItem.Index = FrmMain.ListView.ListItems.Count Then FrmMain.ComNext.Enabled = False: Exit Function
If FrmMain.ListView.SelectedItem.Index < CurSearch Then CurSearch = FrmMain.ListView.SelectedItem.Index

With FrmMain
    For Each itmX In .ListView.ListItems
        i = i + 1
        If ind = 0 Then
            temp = itmX.Text
        Else
            temp = itmX.SubItems(ind)
        End If

        If i > CurSearch Then
            If InStr(1, temp, ftext, vbTextCompare) <> 0 Then
                Set itmRet = itmX
                .ListView.SelectedItem = .ListView.ListItems.Item(i)
                If .ListView.Visible Then .ListView.SetFocus
                LV_EnsureVisible .ListView, i
                ' и пометить если надо
                If .ChMarkFindHid Then .ListView.ListItems(i).Checked = True
                RSGoto .ListView.SelectedItem.Key
                If .ListView.Visible Then
                    .LVCLICK
                    Exit For
                End If
                FindNextLV = FindNextLV + 1
            End If

        End If      'i >
    Next        'For Each

    If i = .ListView.ListItems.Count Then .ComNext.Enabled = False

    CurSearch = i
    Set itmRet = Nothing
End With
End Function

Public Sub OpenNewDataBase()
'из клика по табам выбора баз
FrmMain.Timer2.Enabled = False

'прячем открытые формы
Unload frmSR    'надо для его активейта

ToDebug "Загрузка " & bdname

GroupedFlag = False: FilteredFlag = False    'флаг неполного показа
FiltPersonFlag = False: FiltValidationFlag = False 'персы метка.серийник

Set rs = Nothing: Set DB = Nothing
' ReDim CheckRows(0): ReDim SelRows(0): ReDim lvItemLoaded(0)
' ReDim CheckRowsKey(0): ReDim SelRowsKey(0)

FrmMain.Image0.Cls
'     'nopic
'    Image0.Move 0, 0, FrameImageHid.Width, FrameImageHid.Height
'    If ImageList.ListImages.Count >= 3 Then
'    Image0.PaintPicture ImageList.ListImages(LastImageListInd).Picture, 0, 0, Image0.Width, Image0.Height
'    End If

'сохраняет интерфейс, но не из опций
If Not frmOptFlag Then
    If Opt_AutoSaveOpt Then SaveInterface
End If


'                                                 OpenDB

' в ридИни - ListView.Visible = false при изменении размеров полей
Call NoListClear

If Not OpenDB Then
    'LockWindowUpdate 0
    FrmMain.ListView.Visible = True
    InitFlag = True    'перечитать
    If oldTabLVInd > 0 Then    'при старте и отсутствии базы = 0
        FrmMain.TabLVHid.Tabs(oldTabLVInd).Selected = True
    End If
    Exit Sub
End If

frmEditor.ComDel.Enabled = True

If (oldTabLVInd <> FrmMain.TabLVHid.SelectedItem.Index) Or optReadIniFlag Then
    'не перечитывать ини
    FrmMain.ReadINI GetNameFromPathAndName(bdname)
    NoSetColorFlag = False
    'галки в тудебаг
'    ToDebug "  SortOnStart=" & Opt_SortOnStart
'    ToDebug "  sortcol=" & LVSortColl
'    ToDebug "  CDdrive=" & ComboCDHid_Text
'    ToDebug "  QJPG=" & QJPG
'    ToDebug "  UseAspect=" & Opt_UseAspect
'    ToDebug "  LoanAllSameLabels=" & Opt_LoanAllSameLabels
'    ToDebug "  LoadOnlyTitle=" & Opt_LoadOnlyTitles
    ToDebug "  FreeDVDFilters=" & Opt_UseOurMpegFilters
'    ToDebug "  SaveBigPix=" & Opt_PicRealRes
    ToDebug "  UseProxy=" & Opt_InetUseProxy
    ToDebug "  DS_AVI=" & Opt_AviDirectShow
    ToDebug "  GetVolumeInfo=" & Opt_GetVolumeInfo
    ToDebug "  GetMediaType=" & Opt_GetMediaType
'    ToDebug "  NoAutoStartOpto=" & Opt_QueryCancelAutoPlay
    ToDebug "  InetGetPicUseTempFile=" & Opt_InetGetPicUseTempFile
'    ToDebug "  LVEDIT=" & Opt_LVEDIT
'    ToDebug "  FileWithPath=" & Opt_FileWithPath

End If

oldTabLVInd = FrmMain.TabLVHid.SelectedItem.Index
If rs.RecordCount > 0 Then rs.MoveFirst: rs.MoveLast


'LockWindowUpdate 0 'ListView.Visible = True ': UCLV.Visible = True
Screen.MousePointer = vbHourglass    'выключим в FillListView
'                                              Установка цветов
setForeColor
'                                              Заполнение списка
FrmMain.FillListView
'                                              было помечено
If Len(FrmMain.TextItemHid.Text) = 0 Then
    If CheckCount > 0 Then
        FrmMain.TextItemHid.Text = NamesStore(2) & Chr$(32) & CheckCount
    Else
        FrmMain.TextItemHid.Text = vbNullString
    End If
End If

'If Opt_Group_Vis Then FillTVGroup 'применить группировку
'почистить заранее, а то видно и чистится на глазах
FrmMain.tvGroup.ListItems.Clear: FrmMain.tvGroup.ColumnHeaders(1).Text = "<" & NamesStore(8) & ">"
If Opt_Group_Vis Then FrmMain.mGroup_Click 0    'очистить

If Not (FrmMain.ListView.SelectedItem Is Nothing) Then
    'ListView.SelectedItem.EnsureVisible
    LV_EnsureVisible FrmMain.ListView, FrmMain.ListView.SelectedItem.Index
End If

'предупредить о ридонли
If BaseReadOnly Or BaseReadOnlyU Then
    myMsgBox msgsvc(24), vbInformation ', , FrmMain.hwnd
End If

InitFlag = False

End Sub

Public Sub StoreHistory(newTitle As String, newKey As String)
'последний сверху списка (0)
'mHist(0-14)
'arrHistory()
'newKey с кавычкой в конце
Dim arrTmpKeys(nHistory) As String
Dim arrTmpTitles(nHistory) As String
Dim i As Integer
Dim ExitFlag As Boolean

For i = 0 To nHistory
    If newKey = arrHistoryKeys(i) Then ExitFlag = True: Exit For
Next i
If ExitFlag Then Exit Sub

'спускаем все вниз, нижний пропадает
For i = 0 To nHistory - 1
    arrTmpKeys(i + 1) = arrHistoryKeys(i)
    arrTmpTitles(i + 1) = arrHistoryTitles(i)
Next i
arrTmpKeys(0) = newKey
arrTmpTitles(0) = newTitle

'присваиваем обратно
For i = 0 To nHistory
    arrHistoryKeys(i) = arrTmpKeys(i)
    arrHistoryTitles(i) = arrTmpTitles(i)
Next i


End Sub


