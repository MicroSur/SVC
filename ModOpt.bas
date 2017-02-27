Attribute VB_Name = "ModOpt"
Option Explicit

Public ExportDelim As String ' разделитель экспортируемых полей
' соответствия былых объектов и переменных
Public optsaved As Boolean ' true = изменялись, надобы сохранить
Public delBaseFlag As Boolean 'исключили только что базу
Public optReadIniFlag As Boolean ' опции просят прочитать ини
Public TabSOptLast As Integer 'последний таб

'Public NoCheckLstOptFlag As Boolean - FillLstProgFlag 'не выполнять код при айтем чеке у LstOpt

''''''''''' ОПЦИИ
'переделано на тв Public Const LstOptNum = 16 ' скоко пунктов опций
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Lst Export
'LstExport.ListCount = LstExport_ListCount
Public Const LstExport_ListCount = 24 '19       +текстовые поля + аннотация                       'LstExport_ListCount
Public LstExport_Arr(24) As Boolean ' с 0

' Opt_GoInCatalog - Переходить в окно просмотра
Public Opt_PicRealRes As Boolean
Public Opt_UseOurMpegFilters As Boolean
Public Opt_LoadOnlyTitles As Boolean
Public Opt_NoSlideShow As Boolean
' флаг показа UCLV инфы
Public Opt_UCLV_Vis As Boolean
Public Opt_UCLVPic_Vis As Boolean
' раскрашивать ли должников
Public Opt_Debtors_Colorize As Boolean
' Запоминать настройки
Public Opt_AutoSaveOpt As Boolean
' Масштабировать кадры
Public Opt_UseAspect As Boolean
'LstOpt.Selected (9)
Public Opt_SortOnStart As Boolean '- сортировать список?
Public Opt_Group_Vis As Boolean 'ini-GroupWindow показывать ли Группировку
Public Opt_LoanAllSameLabels As Boolean 'ini-LoanAllSameLabels помечать одноименные метки как отданные должнику

Public Opt_ShowColNames As Boolean 'показывать названия колонок в списке обложки и буфере
Public Opt_CenterShowPic As Boolean 'всегда центровать 1:1 окно

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''List Bases
'LstBases.ListIndex = LstBases_ListIndex - индекс выбранной в опциях базы -1 0 1 ...
Public LstBases_ListIndex As Integer
'LstBases.ListCount = LstBases_ListCount = UBound(LstBases_List)+1
Public LstBases_ListCount As Integer
'LstBases.List(j) = LstBases_List(j) -  таблица с путями к файлам баз
Public LstBases_List() As String '- c 1


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Язык
'ComboLangHid = ComboLang() - таблица языков - c 1
Public ComboLang() As String
Public LangCount As Integer
Public LastLanguage As String '- русский english


'ComboCDHid = ComboCDHid_Text -путь к файлам и буква CD
Public ComboCDHid_Text As String

Public Opt_HtmlJpgName As Integer '0 - filename, 1 - title, 2 - рандом

'CombTemplate.Text = CurrentHtmlTemplate
Public CurrentHtmlTemplate As String
'TxtNnOnPage.text = TxtNnOnPage - кол-во фильмов на страницу
Public TxtNnOnPage_Text As String

'TabLVHid.SelectedItem.Index = CurrentBaseIndex - номер выбранной базы
Public CurrentBaseIndex As Integer 'c 1

Public VMSameColor As Boolean ' использовать для VertMenu цвета LV
Public StripedLV As Boolean 'полосатый lv
Public NoLVSelFrame As Boolean 'true нет рамки выделения
' Сетка в списке фильмов
Public Opt_ShowLVGrid As Boolean

'''''''''''''''INET
Public Opt_InetSecureFlag As Boolean
Public Opt_InetProxyServerPort As String
Public Opt_InetUserName As String
Public Opt_InetPassword As String
Public Opt_InetUseProxy As Integer '0 - no, 1 - IE , 2 - My

'AutoAdd
'запоминать на текущую сессию
Public ch_chSubFolders As Integer
Public ch_chAviHid As Integer
Public ch_chDSHid As Integer
Public ch_chShots As Integer
Public ch_chNoMess As Integer
Public ch_cAutoClose As Integer
Public ch_cEjectMedia As Integer

Public ch_chAutoFiles0  As Integer 'cAddCoverExt in ini
Public ch_chAutoFiles1  As Integer 'cAddCoverAny
Public ch_chAutoFiles2  As Integer 'cAddTXTDescr
Public ch_chAutoFiles3  As Integer 'менять шаблон обложка
Public ch_chAutoFiles4  As Integer 'менять шаблон аннот
'Public ch_chAutoFiles5  As Integer 'нет дублей по имени файла


Public nScrollPos As Long 'позиция скролла LV

Public Opt_GetMediaType As Boolean ' true - читать инфу о типе носителя
Public Opt_GetVolumeInfo As Boolean 'true - читать cd serial и метку
Public Opt_QueryCancelAutoPlay As Boolean 'true - не давать авто запуск
Public Opt_AviDirectShow As Boolean 'false - Обрабатывать AVI фильтрами DirectShow (как MPEG1/2)
Public Opt_LVEDIT As Boolean 'редактировать в гриде
Public Opt_FileWithPath As Boolean 'писать в базу полный путь к файлам фильма

Public Opt_SortLVAfterEdit As Boolean 'сортировать список автоматически после редактирования строк (нет)
Public Opt_SortLabelAsNum As Boolean 'сортировать метку в списке как число (иначе текст) (нет)
Public Opt_PutOtherInAnnot As Boolean 'помещать примечания в конец описания в окне фильмов (нет)



Public Opt_InetGetPicUseTempFile As Boolean 'брать картинку из инета через темп файл, иначе потоком. по умолч = false

Public Opt_ExpUseFolders As Boolean 'chExpFolders
Public Opt_ExpFolder1 As String 'вложенные папки для html
Public Opt_ExpFolder2 As String 'Обложек
Public Opt_ExpFolder3 As String 'кадров


Public Function getExpDelim(s As String) As String
'"Enter" -> vbcrlf

getExpDelim = Replace(s, "enter", vbCrLf, Compare:=vbTextCompare)
getExpDelim = Replace(getExpDelim, "tab", vbTab, Compare:=vbTextCompare)
'If Right$(getExpDelim, 1) <> " " Then getExpDelim = getExpDelim & " "
'Select Case LCase$(s)
'Case "enter": getExpDelim = vbCrLf
'Case "tab": getExpDelim = vbTab
'Case Else:: getExpDelim = s
'End Select

End Function
Public Function putExpDelim(s As String) As String
'vbcrlf -> "Enter"
putExpDelim = Replace(s, vbCrLf, "Enter")
putExpDelim = Replace(putExpDelim, vbTab, "Tab")

'Select Case s
'Case vbCrLf: putExpDelim = "Enter"
'Case vbTab: putExpDelim = "Tab"
'Case Else: putExpDelim = s
'End Select

End Function
'Public Sub setForeColor2()
'Dim c As Control
'On Error Resume Next
'For Each c In FrmMain.Controls
'c.BackColor = LVBackColor
'c.ForeColor = LVFontColor
'Next
'End Sub

Public Sub setForeColor()
'setForeColorOpt
On Error Resume Next

If Not NoSetColorFlag Then    'в начале да

    With FrmMain
        'ToDebug "Установка цветов и шрифтов..."
        'Me.BackColor = LVBackColor
        'PicSplitLVDHid.BackColor = LVBackColor
        
        Set .Font = FontListView
        Set .txtEdit.Font = FontListView
        Set .LstFiles.Font = FontListView
        
        Set .ListView.Font = FontListView
        'ListView.Font.Bold = FontListView.Bold
        'ListView.Font.Charset = FontListView.Charset
        'ListView.Font.Italic = FontListView.Italic
        'ListView.Font.name = FontListView.name
        'ListView.Font.Size = FontListView.Size
        'ListView.Font.Strikethrough = FontListView.Strikethrough
        'ListView.Font.Underline = FontListView.Underline
        'ListView.Font.weight = FontListView.weight

        .ListView.BackColor = LVBackColor
        .ListView.ForeColor = LVFontColor

        .tvGroup.BackColor = LVBackColor
        .tvGroup.ForeColor = LVFontColor
        'Set tvGroup.Font = ListView.Font
        Set .tvGroup.Font = FontListView

        .tvGroup.ColumnHeaders(1).Width = TVWidth - .tvGroup.ColumnHeaders(2).Width - 260    ' +Fr Up
        'аннотация
        'Set TextVAnnot.Font = FontListView
        'Set TextVAnnot.Font = ListView.Font
        Set .TextVAnnot.Font = FontListView

        .TextVAnnot.BackColor = LVBackColor
        .TextVAnnot.ForeColor = LVFontColor

'        .LstFiles.BackColor = LVBackColor
'        .LstFiles.ForeColor = LVFontColor

        .TextItemHid.BackColor = LVBackColor
        .TextItemHid.ForeColor = LVFontColor

        'Set CombFind.Font = FontListView - появляется выделение
        .CombFind.ForeColor = LVFontColor
        .CombFind.BackColor = LVBackColor
        .TextFind.ForeColor = LVFontColor
        .TextFind.BackColor = LVBackColor

        .PicFaceV.BackColor = LVBackColor
        .picScrollBoxV.BackColor = LVBackColor

        '                                                   Acter
        .TextActName.BackColor = LVBackColor
        .TextActBio.BackColor = LVBackColor

        'Set LVActer.Font = FontListView
        'Set LVActer.Font = ListView.Font
        Set .LVActer.Font = FontListView

        .PicActFotoScroll.BackColor = LVBackColor
        .PicActFoto.BackColor = LVBackColor

        .LVActer.BackColor = LVBackColor
        .LVActer.ForeColor = LVFontColor
        'Set TextActName.Font = ListView.Font 'FontListView
        Set .TextActName.Font = FontListView
        .TextActName.ForeColor = LVFontColor

        If Not .FrameActer.Visible Then .TextActBio = vbNullString
        'Set TextActBio.Font = ListView.Font 'FontListView
        Set .TextActBio.Font = FontListView
        .TextActBio.Font.Charset = FontListView.Charset
        .TextActBio.ForeColor = LVFontColor

        'Set TextSearchLVActTypeHid.Font = ListView.Font 'FontListView
        .TextSearchLVActTypeHid.ForeColor = LVFontColor
        .TextSearchLVActTypeHid.BackColor = LVBackColor
        'Set .ListBActHid.Font = ListView.Font 'FontListView
        .ListBActHid.ForeColor = LVFontColor
        .ListBActHid.BackColor = LVBackColor

        .UCLV.BackColor = LVBackColor
        .UCLV.ForeColor = LVFontColor
        'Set .UCLV.Font = ListView.Font 'FontListView
        Set .UCLV.Font = FontListView
                '.UCLV.Controls("lvaddon_cScroll").VBGColor = New_BackColor
        
        If FrmMain.Visible Then .UCLV.Refresh    'resize

        SendMessage .PBar.hwnd, &H2001, 0, ByVal LVBackColor    'RGB(255, 255, 100) 'PBar Forecolor
        SendMessage .PBar.hwnd, &H409, 0, ByVal RGB(100, 100, 100)    'PBar Backcolor
        'SendMessage .PBar.hwnd, &H409, 0, ByVal LVFontColor Xor RGB(100, 100, 100)   'PBar Backcolor

        If VMSameColor Then
            .VerticalMenu.Controls("PicVM").BackColor = LVBackColor
            .VerticalMenu.Controls("PicVM").ForeColor = LVFontColor
            .VerticalMenu.SetClolor
        Else
            .VerticalMenu.Controls("PicVM").BackColor = &H8000000C
            .VerticalMenu.Controls("PicVM").ForeColor = &H80000005
            .VerticalMenu.SetClolor
        End If

'FrmMain.FrFindViewHid.BackColor = LVBackColor
'FrmMain.PicSplitLVDHid.BackColor = LVBackColor
'FrmMain.FrameSearch.BackColor = LVBackColor
'.FrActButtons.BackColor = LVBackColor
'.FrActLeft.BackColor = LVBackColor
'.FrameFoto.BackColor = LVBackColor
'.FrameActer.BackColor = LVBackColor
'.FrActSelect.BackColor = LVBackColor

        .Refresh


        ''сабкласс lv
        If Not DebugMode Then
            
            If StripedLV Or NoLVSelFrame Or Opt_LVEDIT Then    'сабклассить
                ModLVSubClass.UnAttach .FrameView.hwnd
                ModLVSubClass.Attach .FrameView.hwnd
                'глючит если сабклассим цвет хайлайта при мультивыделении и гриде, при мульти и редакторы на выходе
                'LVHighlight.BackGround = LVHighLightLong '&HF0CAA6 'GetSysColor(vbInactiveTitleBarText And Not &H80000000) ' &HF0CAA6    ''&HFF8080    'RGB(0, 128, 0)
                'LVHighlight.ForeGround = .ListView.ForeColor 'RGB(0, 255, 0)
                'ModLVSubClass.SetHighLightColour LVHighlight
                'ModLVSubClass.UseCustomHighLight True 'False 'True

                If NoLVSelFrame Then    'нет рамки выделения
                    ModLVSubClass.NoHighLightFrame True
                Else
                    ModLVSubClass.NoHighLightFrame False
                End If

                LVItemColor.ForeGround = .ListView.ForeColor    'RGB(255, 255, 0)
                If StripedLV Then
                    Dim g_R As Long, g_G As Long, g_B As Long
                    Dim H As Single, s As Single, l As Single
                    Dim R As Long, g As Long, b As Long
                    Dim lv_light As OLE_COLOR
                    Dim lv_color As OLE_COLOR

                    lv_color = .ListView.BackColor

                    If lv_color < 0 Then lv_color = GetSysColor(lv_color And Not &H80000000)

                    'Convert Long value to R,G,B values:
                    g_R = (lv_color And &HFF&)
                    g_G = (lv_color And &HFF00&) \ &H100&
                    g_B = (lv_color And &HFF0000) \ &H10000

                    'Get H,S,L values:
                    RGBToHSL g_R, g_G, g_B, H, s, l

                    'Get 3DColor values (on L)
                    HSLToRGB H, s, l + (1 - l) / 4, R, g, b
                    lv_light = RGB(R, g, b)
                    'HSLToRGB H, s, l + (1 - l) / 2, R, G, B
                    '    g_HighLight = RGB(R, G, B)
                    'HSLToRGB H, s, l / 1.5, R, G, B
                    '    g_Shadow = RGB(R, G, B)
                    'HSLToRGB H, s, l / 3.5, R, G, B
                    '    g_DarkShadow = RGB(R, G, B)

                    LVItemColor.BackGround = lv_light    'RGB(g_R, g_G, g_B)  'lv_light '&HE4E2D9 '&HE4E2CA     'RGB(255, 255, 0)
                    ModLVSubClass.SetCustomColour LVItemColor
                    ModLVSubClass.UseAlternatingColour True    'цвета строк черезстрочно
                Else
                    ModLVSubClass.UseAlternatingColour False
                End If    'StripedLV
            Else
                ModLVSubClass.UnAttach .FrameView.hwnd
                If .ListView.Visible Then FrmMain.ListView.Refresh
                If .tvGroup.Visible Then FrmMain.tvGroup.Refresh

            End If    'If StripedLV Or NoLVSelFrame
        End If    'If Not DebugMode
        ''''''''''


    End With

    NoSetColorFlag = True
End If    'flag
End Sub

Public Sub AddTabsLV()
'создать кнопки подключенных баз
Dim i As Integer
Dim temp As String, temp2 As String

For i = FrmMain.TabLVHid.Tabs.Count To 1 Step -1
    FrmMain.TabLVHid.Tabs.Remove i
Next i

If LstBases_ListCount = 0 Then 'нет баз
    InitFlag = True
    Exit Sub
End If

For i = 1 To LstBases_ListCount
    temp = GetNameFromPathAndName(LstBases_List(i))
    GetExtensionFromFileName temp, temp2
    FrmMain.TabLVHid.Tabs.Add i, , temp2 ', 1
Next
End Sub
