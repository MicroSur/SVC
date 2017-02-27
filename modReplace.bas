Attribute VB_Name = "modReplace"
Option Explicit

Public Function InsertTextInDB(ind As Integer, lvInd As Long, insText As String, Optional LookIn As LV_AllSelCheck = AllLVRows, Optional InsertIn As BeginEnd = sBegin) As Boolean
'���������� ���� �������������� ��� ��� (��� ��������)
'������� � ������� ������, �� ������� �����
'ind - ���� ����
'LVInd -������ � lv

Dim temp As String '������ ����
Dim GoGoGo As Boolean

With FrmMain.ListView

GoGoGo = False
Select Case LookIn
Case AllLVRows
    GoGoGo = True
Case CheckedLVRows
    If .ListItems(lvInd).Checked Then GoGoGo = True
Case SelectedLVRows
    If .ListItems(lvInd).Selected Then GoGoGo = True
End Select

If GoGoGo Then

    '��������������� ����, �������� �������� ����
    RSGoto .ListItems(lvInd).Key
    If rs(ind).Value <> vbNullString Then temp = rs(ind) Else temp = vbNullString
    
    Select Case InsertIn

    Case sBegin    '� ������
        temp = insText & temp
    Case sEnd    ' � �����
        temp = temp & insText
    End Select

    '�������� � ����
    rs.Edit

    On Error GoTo err
    rs(ind) = temp
    rs.Update

    InsertTextInDB = True '���������
    
    '������� � ������
    If ind <> lvIndexPole Then
    If ind = 0 Then
        .ListItems(lvInd).Text = temp
    Else
        .ListItems(lvInd).SubItems(ind) = temp
    End If
    End If
    '���������� ������ lv - �����
    'EditLV lvInd

End If    'go

End With
Exit Function

err:
ToDebug "Err_itid: " & err.Description
rs.CancelUpdate

End Function
Public Sub ReplaceInDB(ind As Integer, lvInd As Long, ftext As String, RText As String, CompareMethod As VbCompareMethod, Optional LookIn As LV_AllSelCheck = AllLVRows, Optional StartIn As AnyWholeFirst = Search_Anywhere)
'������ � ������� ������, �� ������� �����
'ind - ���� ����
'LVInd -������ � lv

Dim temp As String
Dim GoGoGo As Boolean
With FrmMain.ListView

    GoGoGo = False
    Select Case LookIn
    Case AllLVRows
        GoGoGo = True
    Case CheckedLVRows
        If .ListItems(lvInd).Checked Then GoGoGo = True
    Case SelectedLVRows
        If .ListItems(lvInd).Selected Then GoGoGo = True
    End Select

    If GoGoGo Then
    
    '��������������� ����, �������� �������� ����
    RSGoto .ListItems(lvInd).Key
    If rs(ind).Value <> vbNullString Then temp = rs(ind) Else temp = vbNullString
    
    Select Case StartIn
    
    Case Search_Anywhere ' � ����� ����� ����
        
        If (Len(temp) = 0) And (Len(ftext) = 0) Then
            temp = RText
        Else
            temp = Replace(temp, ftext, RText, Compare:=CompareMethod)
        End If

    Case Search_StartWith ' � ������ ����
        Select Case CompareMethod
            Case vbTextCompare
                If LCase$(left$(temp, Len(ftext))) = LCase$(ftext) Then
                    temp = RText & right$(temp, Len(temp) - Len(ftext))
                End If
            Case vbBinaryCompare
                If left$(temp, Len(ftext)) = ftext Then
                    temp = RText & right$(temp, Len(temp) - Len(ftext))
                End If

        End Select
      
    Case Search_WholeField '���� �������
        Select Case CompareMethod
            Case vbTextCompare
                If LCase$(temp) = LCase$(ftext) Then
                    temp = RText
                End If
            Case vbBinaryCompare
                If temp = ftext Then
                    temp = RText
                End If
        End Select
        
    Case Search_Shablon '���� ������� � ��������� *?[!0-9�-�]
        If temp Like ftext Then temp = RText
' ���� ���� ������� ����
'      -----------------
' ����*����
'

    Case Search_EndWith '� ����� ����
        Select Case CompareMethod
            Case vbTextCompare
                If LCase$(right$(temp, Len(ftext))) = LCase$(ftext) Then
                 temp = left$(temp, Len(temp) - Len(ftext)) & RText
                End If
            Case vbBinaryCompare
                If right$(temp, Len(ftext)) = ftext Then
                 temp = left$(temp, Len(temp) - Len(ftext)) & RText
                End If
        End Select
    
    End Select

'�������� � ����
rs.Edit

On Error GoTo err
rs(ind) = temp
rs.Update

    '������� � ������
    If ind <> lvIndexPole Then
    If ind = 0 Then
        .ListItems(lvInd).Text = temp
    Else
        .ListItems(lvInd).SubItems(ind) = temp
    End If
    End If
    '���������� ������ lv - �����
    'EditLV lvInd

End If 'go

End With
Exit Sub

err:
ToDebug "Err_rid: " & err.Description
rs.CancelUpdate

End Sub

Public Function SearchLV(ind As Integer, ftext As String, CompareMethod As VbCompareMethod, Optional SearchIn As LV_AllSelCheck = AllLVRows) As Integer
' ind - ������� � LV
Dim itmX As ListItem
Dim itmRet As ListItem
Dim i As Long    ', j As Integer
Dim temp As String
Dim GoGoGo As Boolean

With FrmMain
    For Each itmX In .ListView.ListItems
        i = i + 1

        GoGoGo = False
        Select Case SearchIn
        Case AllLVRows
            GoGoGo = True
        Case CheckedLVRows
            If itmX.Checked Then GoGoGo = True
        Case SelectedLVRows
            If itmX.Selected Then GoGoGo = True
        End Select

        If GoGoGo Then

            If ind = 0 Then
                temp = itmX.Text
            Else
                temp = itmX.SubItems(ind)
            End If

            If InStr(1, temp, ftext, CompareMethod) <> 0 Then
                Set itmRet = itmX
                .ListView.SelectedItem = .ListView.ListItems.Item(i)
                If .ListView.Visible = True Then .ListView.SetFocus

                LV_EnsureVisible FrmMain.ListView, i
                ' � �������� ���� ����
                If .ChMarkFindHid Then .ListView.ListItems(i).Checked = True
                RSGoto .ListView.SelectedItem.Key

                If .ListView.Visible = True Then .LVCLICK
                SearchLV = 1
                Exit For
            End If

        End If    'go
    Next        'For Each
    CurSearch = i
    Set itmRet = Nothing
End With
End Function

Public Sub SearchNextDB(ind As Integer, ftext As String, StartFrom As Integer, WithCurrent As Boolean, CompareMethod As VbCompareMethod, Optional LookIn As LV_AllSelCheck = AllLVRows, Optional StartIn As AnyWholeFirst = Search_Anywhere)
'first = true - �������� �� ������ LV �� ������� �����
Dim temp As String
Dim ret As Long
'Dim i As Integer
Dim nxt As Integer
Dim GoGoGo As Boolean

nxt = StartFrom 'ListView.SelectedItem.Index
If WithCurrent Then nxt = StartFrom - 1 'ListView.SelectedItem.Index - 1


Do 'While Not rs.EOF
    nxt = nxt + 1
    If FrmMain.ListView.ListItems.Count < nxt Then FrmMain.ComNext.Enabled = False: Exit Sub
With FrmMain
    GoGoGo = False
    Select Case LookIn
    Case AllLVRows
        GoGoGo = True
    Case CheckedLVRows
        If .ListView.ListItems(nxt).Checked Then GoGoGo = True
    Case SelectedLVRows
        If .ListView.ListItems(nxt).Selected Then GoGoGo = True
    End Select

    If GoGoGo Then

    RSGoto .ListView.ListItems(nxt).Key
    If rs(ind).Value <> vbNullString Then temp = rs(ind) Else temp = vbNullString
    
    Select Case StartIn
    
    Case Search_Anywhere ' � ����� ����� ����
        If (Len(temp) = 0) And (Len(ftext) = 0) Then
            ret = 1
        Else
            ret = InStr(1, temp, ftext, CompareMethod)
        End If
    Case Search_StartWith ' � ������ ����
        ret = 0
        Select Case CompareMethod
            Case vbTextCompare
                If LCase$(left$(temp, Len(ftext))) = LCase$(ftext) Then ret = 1
            Case vbBinaryCompare
                If left$(temp, Len(ftext)) = ftext Then ret = 1
        End Select
      
    Case Search_WholeField '���� �������
        ret = 0
        Select Case CompareMethod
            Case vbTextCompare
                If LCase$(temp) = LCase$(ftext) Then ret = 1
            Case vbBinaryCompare
                If temp = ftext Then ret = 1
        End Select
        
    Case Search_Shablon '���� ������� � ��������� *?[!]
        ret = 0
        If temp Like ftext Then ret = 1

    Case Search_EndWith '� ����� ����
        ret = 0
        Select Case CompareMethod
            Case vbTextCompare
                If LCase$(right$(temp, Len(ftext))) = LCase$(ftext) Then ret = Len(temp) - Len(ftext) + 1
            Case vbBinaryCompare
                If right$(temp, Len(ftext)) = ftext Then ret = Len(temp) - Len(ftext) + 1
        End Select
    
    End Select
     
    If ret > 0 Then

        Set .ListView.SelectedItem = .ListView.ListItems.Item(nxt)
        .ListView.SetFocus
        .ListView.ListItems(.ListView.SelectedItem.Index).EnsureVisible
        ' � �������� ���� ����
        If .ChMarkFindHid Then .ListView.ListItems(.ListView.SelectedItem.Index).Checked = True

        .LVCLICK
        
If StartIn <> Search_Shablon Then
       
        If ind = dbAnnotationInd Then
            .ComShowAn_Click
            .TextVAnnot.SetFocus
            .TextVAnnot.SelStart = ret - 1
            .TextVAnnot.SelLength = Len(ftext)
            'TextVAnnot.SelText = FText
        Else
            .TextItemHid = temp
            .TextItemHid.SetFocus
            .TextItemHid.SelStart = ret - 1
            .TextItemHid.SelLength = Len(ftext)
        End If
Else
    If ind = dbAnnotationInd Then
    Else
        .TextItemHid = temp
        .TextItemHid.SetFocus
        .TextItemHid.SelStart = 0
        .TextItemHid.SelLength = Len(temp)
    End If
End If 'StartIn <> Search_Shablon

        Exit Do '�����, �����
    End If 'ret > 0
End If 'go

End With
Loop

End Sub
Public Sub ConvertInDB(ind As Integer, lvInd As Long, ConvMethod As HowConvert, sDelimiter As String, Optional LookIn As LV_AllSelCheck = AllLVRows)
'�������������� � ������� ������
'ind - ���� ����
'LVInd -������ � lv

Dim temp As String    '������ ����
Dim lenTemp As Long    '����� ������ ����
Dim GoGoGo As Boolean

GoGoGo = False
Select Case LookIn
Case AllLVRows
    GoGoGo = True
Case CheckedLVRows
    If FrmMain.ListView.ListItems(lvInd).Checked Then GoGoGo = True
Case SelectedLVRows
    If FrmMain.ListView.ListItems(lvInd).Selected Then GoGoGo = True
End Select

If GoGoGo Then

    '��������������� ����, �������� �������� ����
    RSGoto FrmMain.ListView.ListItems(lvInd).Key
    If rs(ind).Value <> vbNullString Then temp = rs(ind) Else temp = vbNullString
    lenTemp = Len(temp)
    If lenTemp = 0 Then Exit Sub

    With FrmMain

        Select Case ConvMethod
        Case LCaseAll
            temp = LCase$(temp)
        Case UCaseAll
            temp = UCase$(temp)
        Case UcaseFirst
            temp = UCase$(left$(temp, 1)) & LCase$(right$(temp, Len(temp) - 1))    '����� ������ ������
        Case UCaseWord
            temp = UcaseCharAfterDelimiter(temp, sDelimiter)
        End Select

        '�������� � ����
        rs.Edit

        On Error GoTo err
        rs(ind) = temp
        rs.Update

        '������� � ������
        If ind <> lvIndexPole Then
            If ind = 0 Then
                .ListView.ListItems(lvInd).Text = temp
            Else
                .ListView.ListItems(lvInd).SubItems(ind) = temp
            End If
        End If
        '���������� ������ lv - �����
        'EditLV lvInd
    End With
End If    'go
Exit Sub

err:
ToDebug "Err_cid: " & err.Description
rs.CancelUpdate


End Sub
Public Sub DeleteInDB(ind As Integer, lvInd As Long, DelLen As Long, DelStart As Long, Optional LookIn As LV_AllSelCheck = AllLVRows, Optional CountFrom As BeginEnd = sBegin)
'�������� � ������� ������, �� ������� �����
'ind - ���� ����
'LVInd -������ � lv

Dim temp As String    '������ ����
Dim lenTemp As Long    '����� ������ ����
Dim GoGoGo As Boolean

GoGoGo = False
Select Case LookIn
Case AllLVRows
    GoGoGo = True
Case CheckedLVRows
    If FrmMain.ListView.ListItems(lvInd).Checked Then GoGoGo = True
Case SelectedLVRows
    If FrmMain.ListView.ListItems(lvInd).Selected Then GoGoGo = True
End Select

If GoGoGo Then

    '��������������� ����, �������� �������� ����
    RSGoto FrmMain.ListView.ListItems(lvInd).Key
    If rs(ind).Value <> vbNullString Then temp = rs(ind) Else temp = vbNullString
    lenTemp = Len(temp)
    If lenTemp = 0 Then Exit Sub

    With FrmMain
        Select Case CountFrom
        Case sBegin    '� ������
            temp = sys_StrDel(temp, DelStart, DelLen)
        Case sEnd    '� �����
            temp = sys_StrDelRev(temp, DelStart, DelLen)
        End Select

        '�������� � ����
        rs.Edit

        On Error GoTo err
        rs(ind) = temp
        rs.Update

        '������� � ������
        If ind <> lvIndexPole Then
            If ind = 0 Then
                .ListView.ListItems(lvInd).Text = temp
            Else
                .ListView.ListItems(lvInd).SubItems(ind) = temp
            End If
        End If
        '���������� ������ lv - �����
        'EditLV lvInd
    End With
End If    'go
Exit Sub

err:
ToDebug "Err_did: " & err.Description
rs.CancelUpdate

End Sub
Public Sub DeletePixInDB(ind As Integer, lvInd As Long, Optional LookIn As LV_AllSelCheck = AllLVRows)
'�������� � ������� ������, �� ������� �����
'ind - ���� ���� ��������
'LVInd -������ � lv

'Dim temp As String '������ ����
'Dim lenTemp As Long '����� ������ ����
Dim GoGoGo As Boolean

GoGoGo = False
Select Case LookIn
Case AllLVRows
    GoGoGo = True
Case CheckedLVRows
    If FrmMain.ListView.ListItems(lvInd).Checked Then GoGoGo = True
Case SelectedLVRows
    If FrmMain.ListView.ListItems(lvInd).Selected Then GoGoGo = True
End Select

If GoGoGo Then

    '��������������� ����, �������� �������� ����
    RSGoto FrmMain.ListView.ListItems(lvInd).Key
    
    '�������� � ����
    rs.Edit
    On Error GoTo err
    rs(ind) = vbNullString
    rs.Update

End If    'go
Exit Sub

err:
ToDebug "Err_dpid: " & err.Description
rs.CancelUpdate

End Sub

