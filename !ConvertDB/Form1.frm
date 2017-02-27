VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert2SVC"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frAMC 
      Caption         =   "AntMovieCatalog 3.5"
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   8835
      Begin VB.TextBox tPicLocate 
         Height          =   285
         Left            =   60
         TabIndex        =   26
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lPicLocate 
         Caption         =   "Путь к обложкам. Если не указано - искать в amc базе"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   240
         Width           =   5715
      End
   End
   Begin VB.Frame frCSV 
      Caption         =   "Comma Separated Values"
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   5880
      Width           =   8835
      Begin VB.TextBox tCsvDelim 
         Height          =   285
         Left            =   60
         MaxLength       =   1
         TabIndex        =   25
         Text            =   ","
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lCsvDelim 
         Caption         =   "Разделитель  полей в CSV файле"
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   5715
      End
   End
   Begin VB.CommandButton ComOpenSVC 
      Caption         =   "Актеры"
      Height          =   405
      Index           =   1
      Left            =   4080
      TabIndex        =   35
      ToolTipText     =   "Открыть базу people.mdb"
      Top             =   120
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   4440
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   1
   End
   Begin VB.CheckBox chCSV 
      Caption         =   "Пометки"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4860
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox ListItemSVC 
      Height          =   2790
      ItemData        =   "Form1.frx":030A
      Left            =   2760
      List            =   "Form1.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   2595
   End
   Begin VB.ComboBox cbCsvSvc 
      Height          =   315
      Left            =   5400
      TabIndex        =   32
      Top             =   120
      Visible         =   0   'False
      Width           =   3555
   End
   Begin MSComctlLib.ListView lvCSV 
      Height          =   4215
      Left            =   7860
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton ComReadMe 
      Caption         =   "Почитать файл помощи"
      Height          =   375
      Left            =   2820
      TabIndex        =   21
      Top             =   4860
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   6480
      Width           =   8835
      Begin VB.CheckBox ChNoDup 
         Caption         =   "Не импортировать дубликаты :"
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   120
         Width           =   2835
      End
      Begin VB.TextBox TxtLSplit 
         Height          =   315
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Text            =   "Form1.frx":030E
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox TxtRSplit 
         Height          =   315
         Left            =   1020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "Form1.frx":0311
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Разделители для объединенных полей"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   480
         Width           =   5715
      End
      Begin VB.Label LDupFields 
         Caption         =   "Выберите связь (для CSV - заголовок нужного поля)"
         Height          =   225
         Left            =   3000
         TabIndex        =   20
         Top             =   165
         Width           =   5655
      End
   End
   Begin VB.CommandButton ComOld2NewSVC 
      Caption         =   "Связать похожие"
      Height          =   315
      Left            =   2820
      TabIndex        =   17
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton ComCheck 
      Caption         =   "Проверка данных"
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   2535
   End
   Begin VB.ListBox LstRSplit 
      Height          =   1815
      Left            =   9840
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ListBox LstLSplit 
      Height          =   1815
      Left            =   9120
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ListBox ListItemToDoView 
      Height          =   2790
      ItemData        =   "Form1.frx":0313
      Left            =   5400
      List            =   "Form1.frx":0315
      TabIndex        =   5
      Top             =   1080
      Width           =   3555
   End
   Begin VB.TextBox TextSVC 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   345
      Left            =   7320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   1635
   End
   Begin VB.TextBox TextIn 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   345
      Left            =   5400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   1635
   End
   Begin VB.CommandButton ComDelLink 
      Caption         =   "Отменить связь"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5700
      TabIndex        =   7
      Top             =   4380
      Width           =   3015
   End
   Begin VB.CommandButton ComOpenSVC 
      Caption         =   "Фильмы"
      Height          =   405
      Index           =   0
      Left            =   2820
      TabIndex        =   1
      ToolTipText     =   "Открыть базу фильмов SurVideoCatalog"
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton ComOpenIn 
      Caption         =   "Открыть источник "
      Height          =   405
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Открыть базу-источник"
      Top             =   120
      Width           =   2595
   End
   Begin VB.CommandButton ComImport 
      Caption         =   "Импорт"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5700
      TabIndex        =   8
      Top             =   4860
      Width           =   3015
   End
   Begin VB.ListBox ListItemToDo 
      Height          =   2595
      ItemData        =   "Form1.frx":0317
      Left            =   9180
      List            =   "Form1.frx":0319
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox ListItemIn 
      Height          =   2790
      ItemData        =   "Form1.frx":031B
      Left            =   120
      List            =   "Form1.frx":031D
      TabIndex        =   3
      Top             =   1080
      Width           =   2595
   End
   Begin VB.CommandButton ComAddLink 
      Caption         =   "Добавить связь"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5700
      TabIndex        =   6
      Top             =   3960
      Width           =   3015
   End
   Begin VB.ComboBox CmbTableIn 
      Height          =   315
      ItemData        =   "Form1.frx":031F
      Left            =   120
      List            =   "Form1.frx":0321
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   630
      Width           =   2595
   End
   Begin VB.Image ImgTest 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   240
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   1035
   End
   Begin VB.Label LblType 
      Alignment       =   1  'Right Justify
      Caption         =   "Типы данных : "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "в"
      Height          =   255
      Left            =   7140
      TabIndex        =   12
      Top             =   720
      Width           =   255
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu mCheckAll 
         Caption         =   "Пометить все"
      End
      Begin VB.Menu mCheckNone 
         Caption         =   "Снять пометки"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Last
',LCID

' ? парси XML с помощью SAX из той же MSXML. Может XSLT справится?
' + Не давать выбрать поле дупов если это картинка
' ? дупы в csv & amc
' + конвертер amc
' + конвертер CSV
' + поиск синонимов для автосвязей (screen40)
' + Не импортировтаь пустую связь в скобках ()
' - # в источнике не проходят проверку дубликата
' - собирать из нескольких источников, находящихся в разных нодах
'
' ! поставить On Error в сабах импорта
'
'? не дуп никогда  = 'Калигула 2 - нерасказанная история (                                                                                                                                                                                                                  L'' (Ital'

Private MyTableName As String 'Storage или Acter
Private rsIndex As String 'Key KeyAct -кажется проверка в холостую

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private CSVDupField As Integer 'поле проверки для csv дупов

Private FromDupSub As Boolean ' флаг, что вызов происходит из саба поисков дупов

Private TotalCsvFields As Long 'c 1
Private pArr() As Long    'заполнять с 1
Private arrProcess() As Boolean 'обработано ли поле (csv)

Private arrAMC() As Long 'с 0,1 . двумерный (33,число записей) - хранит позиции начала полей
Private AMCTotalRec As Long
Private amcNamesArr(33) As String
Private amcTypesArr(33) As Integer

Private DupInd As Integer 'индекс в списке связей выделенной связи для дубликатов
Private Dobavleno As Long 'сколько записей импортировано
Private VsegoBilo As Long 'сколько записей на входе

Private newbdname As String
Private bdname As String

'mdb
Private DB As DAO.Database
Private rs As DAO.Recordset
Private newDB As DAO.Database
Private inRs As DAO.Recordset

'xml
Private XML_tmp As Object
Private XML_Child As Object
Private XML_Attr As Object

Private IsXMLflag As Boolean
Private IsMDBflag As Boolean
Private IsAMCflag As Boolean
Private IsCSVflag As Boolean

Private SVCPoleName As String

'гориз. скролл листбокса
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Private LastSVCInd  As Integer ' куда кликали последний раз

'Private Declare Function SendMessage Lib "user32" _
  Alias "SendMessageA" (ByValhwnd As Long, _
  ByValwMsg As Long, ByValwParam As Long, _
  ByVallParam As String) As Long
Private Const LB_SELECTSTRING = &H18C



Private Sub chCSV_Click()
Dim Itm As MSComctlLib.ListItem

If chCSV.Value = vbChecked Then
'пометить все
For Each Itm In lvCSV.ListItems: Itm.Checked = True: Next
Else
For Each Itm In lvCSV.ListItems: Itm.Checked = False: Next
End If
End Sub

Private Sub ChNoDup_Click()

If IsCSVflag Then
 'CSV
 If ChNoDup.Value = vbUnchecked Then LDupFields.ForeColor = &H80000012: Exit Sub
    If CSVDupField = 0 Then

        LDupFields.ForeColor = &HFF&
        LDupFields.Caption = "Кликните на заголовок нужного поля"

    Else
    If Len(lvCSV.ColumnHeaders(CSVDupField)) = 0 Then
        LDupFields.ForeColor = &HFF&
        LDupFields.Caption = "Кликните на заголовок нужного поля"
Else
        LDupFields = lvCSV.ColumnHeaders(CSVDupField)
        LDupFields.ForeColor = &H80000012
        End If
    End If
Else
    If ChNoDup.Value = vbUnchecked Then
        LDupFields.ForeColor = &H80000012
    Else
        If ListItemToDoView.ListIndex < 0 Then
            ChNoDup.Value = vbUnchecked
            LDupFields.ForeColor = &HFF&
            LDupFields.Caption = "выберите связь"
        Else
            LDupFields = ListItemToDoView.List(ListItemToDoView.ListIndex)
            LDupFields.ForeColor = &H80000012
        End If
    End If
End If
End Sub

Private Sub CmbTableIn_Click()
ListItemIn.Clear

If IsXMLflag Then
    ProcessChild CmbTableIn.Text
End If

If IsMDBflag Then
    OpenInRs
'Debug.Print CmbTableIn.ListIndex
End If


End Sub
Private Sub OpenInRs()
On Error GoTo err
Dim tmp As String
'!после помещения в ListItemIn имя поля уже не равно реальному (с умляутами), с переменными все нормально
'список не сортирован
Dim i As Integer
Set inRs = newDB.OpenRecordset(CmbTableIn.Text, dbOpenTable)
For i = 0 To inRs.Fields.Count - 1
    If inRs.Fields(i).Name <> vbNullString Then


        'Debug.Print inRs(tmp)
        'Debug.Print tmp = inRs.Fields(i).Name
        'ListItemIn.Font.Charset = 238
        
        
        ListItemIn.AddItem inRs.Fields(i).Name

        'Попытка сразу взять значение поля по имени
       ' tmp = inRs(ListItemIn.List(i))    '- ошибки с умляутами
If IsNull(inRs(ListItemIn.List(i))) Then
End If

        ', inRs.Fields(i).OrdinalPosition - 1
        'ListItemIn.ItemData(i) = inRs.Fields(i).Type ' не работает с ундо
        'ListItemIn.ItemData(i) = inRs.Fields(i).OrdinalPosition

    Else
        ListItemIn.AddItem vbNullString         'чисто для проверки?
    End If
Next i

Exit Sub

err:
Select Case err.Number
Case 3265
    MsgBox "Не прошла проверка обращения к полю по имени: " & ListItemIn.List(i) & vbCrLf & "Возможно в именах полей присутствуют умляуты." & vbCrLf & "В данной версии возможны ошибки при работе с этой таблицей.", vbCritical
    Resume Next
Case Else
    If err.Number <> 0 Then MsgBox err.Number & vbCrLf & err.Description, vbCritical
End Select

End Sub
Private Sub ComAddLinkClick()
Dim LstIndIn As Integer, LstIndSVC As Integer, LstIndToDo As Integer
'Dim i As Integer
Dim a() As String

If ListItemIn.SelCount = 0 Then Exit Sub


If ListItemSVC.SelCount = 0 Then
    If ListItemToDoView.SelCount > 0 Then
    'Объединение (Изменение пары)
    'текущие листиндексы
    LstIndIn = ListItemIn.ListIndex
    LstIndToDo = ListItemToDoView.ListIndex
    
    'Разделить пару
    a = Split(ListItemToDo.List(LstIndToDo), " > ", 2, vbTextCompare)
    'изменить пару в туду, вставив новый |In|
    ListItemToDo.List(LstIndToDo) = a(0) & " | " & ListItemIn.Text & " | " & " > " & a(1)
    'то же для View
    a = Split(ListItemToDoView.List(LstIndToDo), " > ", 2, vbTextCompare)
    ListItemToDoView.List(LstIndToDo) = a(0) & TxtLSplit & ListItemIn.Text & TxtRSplit & " > " & a(1)
    
    'изменить листы с разделителями
    LstLSplit.List(LstIndToDo) = LstLSplit.List(LstIndToDo) & "$0$" & TxtLSplit.Text
    LstRSplit.List(LstIndToDo) = LstRSplit.List(LstIndToDo) & "$0$" & TxtRSplit.Text

    'убрать из списков-источников
    ListItemIn.RemoveItem LstIndIn
        Exit Sub
    Else
        Exit Sub
    End If
End If

'Добавление новой пары
'текущие листиндексы
LstIndIn = ListItemIn.ListIndex
LstIndSVC = ListItemSVC.ListIndex

'вставить пару в туду
ListItemToDo.AddItem ListItemIn.Text & " > " & ListItemSVC.Text
'ListItemToDo.ItemData(ListItemToDoView.ListCount - 1) = CmbTableIn.ListIndex

ListItemToDoView.AddItem ListItemIn.Text & " > " & ListItemSVC.Text
ListItemToDoView.ItemData(ListItemToDoView.ListCount - 1) = rs.Fields(ListItemSVC.List(ListItemSVC.ListIndex)).Type

'Заполнить листы разделителями
LstLSplit.AddItem TxtLSplit.Text
LstRSplit.AddItem TxtRSplit.Text

'убрать из списков-источников
ListItemIn.RemoveItem LstIndIn
ListItemSVC.RemoveItem LstIndSVC


SetListboxScrollbar ListItemToDo
ComDelLink.Enabled = True: ComImport.Enabled = True: ComAddLink.Enabled = False
CmbTableIn.Enabled = False

End Sub

Private Sub CmbTableIn_GotFocus()
'If IsCSVflag Then CmbTableIn.Text = vbNullString

End Sub

Private Sub ComAddLink_Click()
ComAddLinkClick
End Sub



Private Sub ComCheck_Click()
Dim tmp As String

If ListItemIn.Text = vbNullString Then Exit Sub

Set ImgTest = Nothing


If IsXMLflag Then
    Dim XML_1 As Object
    Dim XML_2 As Object
    Dim XML_a As Object
    Dim IsNode As Boolean
    'Добываем значения по имени ноды(аттра) для проверки

    Set XML_1 = XML_Root.getElementsByTagName(ListItemIn.Text)
    'Debug.Print XML_1.nodeType
    For Each XML_2 In XML_1
        MsgBox "Node: " & XML_2.nodeName & "=" & XML_2.nodeTypedValue
        'Debug.Print "NodeName>>> " & XML_2.NodeName
        'Debug.Print "nodeTypedValue>>> " & XML_2.nodeTypedValue
        IsNode = True
        Exit For    '1 разик
    Next

    If Not IsNode Then

        Set XML_1 = XML_Root.getElementsByTagName(CmbTableIn.Text)

        For Each XML_2 In XML_1
            'Debug.Print "NodeType>>> " & XML_2.NodeType
            If XML_2.Attributes.length > 0 Then
                Set XML_a = XML_2.Attributes
                MsgBox "Attr: " & ListItemIn.Text & "=" & XML_a.getNamedItem(ListItemIn.Text).nodeValue
                'Debug.Print XML_a.getNamedItem(ListItemIn.Text).nodeValue
                Exit For    '1
            End If
        Next

    End If
End If

If IsMDBflag Then
    On Error Resume Next
    If inRs.BOF And inRs.EOF Then    'ваще пустая
        MsgBox "В таблице нет записей."
        Exit Sub
    End If
    inRs.MoveFirst
    If IsNull(inRs(ListItemIn.Text)) Then
        If err = 3265 Then 'умляут
            MsgBox "Ошибка работы с полем.", vbCritical
        Else
            MsgBox "Не заполнено."
        End If
    ElseIf CStr(inRs(ListItemIn.Text)) = vbNullString Then
        MsgBox "Пусто."
    Else
        If inRs(ListItemIn.List(ListItemIn.ListIndex)).Type = 11 Then
            ImgTest.Picture = GetMDB_Pic
        Else
            MsgBox inRs(ListItemIn.Text)
        End If
    End If
    err.Clear: On Error GoTo 0
End If

If IsAMCflag Then
    Select Case GetAmcTypeFromName(ListItemIn.List(ListItemIn.ListIndex))
    Case 1    'integer
        tmp = GetAMC_Int(GetAmcIndFromName(ListItemIn.List(ListItemIn.ListIndex)), 1)
        If Len(tmp) <> 0 Then MsgBox tmp Else MsgBox "Пусто"
    Case 2    'boolean
        tmp = GetAMC_Bool(GetAmcIndFromName(ListItemIn.List(ListItemIn.ListIndex)), 1)
        If Len(tmp) <> 0 Then MsgBox tmp Else MsgBox "Пусто"
    Case 3    'string
        tmp = GetAMC_Str(GetAmcIndFromName(ListItemIn.List(ListItemIn.ListIndex)), 1)
        If Len(tmp) <> 0 Then MsgBox tmp Else MsgBox "Пусто"
    Case 4    'data
        ImgTest.Picture = GetAMC_PicShow(GetAmcIndFromName(ListItemIn.List(ListItemIn.ListIndex)), 1)
    End Select


End If

End Sub
Private Function GetAMC_Str(col As Integer, row As Long) As String
Dim N As Long, tmpl As Long
N = arrAMC(col, row)
'длина
tmpl = GetInt(N)
N = N + 4
'string
GetAMC_Str = GetStr(N, tmpl)
End Function
Private Function GetAMC_Int(col As Integer, row As Long) As Long
Dim N As Long ', tmpl As Long
N = arrAMC(col, row)
If col = 2 Then    'Rating
    GetAMC_Int = GetInt(N) / 10
Else
    GetAMC_Int = GetInt(N)
End If
End Function
Private Function GetAMC_Bool(col As Integer, row As Long) As Integer
Dim N As Long ', tmpl As Long
N = arrAMC(col, row)
GetAMC_Bool = GetBoolean(N)
End Function
Private Function GetAMC_PicShow(col As Integer, row As Long) As StdPicture
Dim PicSize As Long
Dim PicName As String
Dim img As ImageFile
Dim vec As Vector
Dim pb() As Byte
Dim N As Long
'Dim l As Long

On Error GoTo err

'name
'N = arrAMC(col - 2, row)
PicName = GetAMC_Str(col - 2, row)

'size
N = arrAMC(col - 1, row)
PicSize = GetInt(N)

'Picture
If PicSize > 0 Then
    'картинку из базы
    N = arrAMC(col, row)
    ReDim pb(PicSize - 1)
    Get ff, N, pb
    Set img = New ImageFile
    Set vec = New Vector
    vec.BinaryData = pb
    Set img = vec.ImageFile
    If Not img Is Nothing Then
        Set GetAMC_PicShow = img.ARGBData.Picture(img.Width, img.Height)
    End If
Else
    If Len(tPicLocate) > 0 Then
        'обрезать путь к картинке и добавить пользовательский путь
        If Right$(tPicLocate, 1) <> "\" Then tPicLocate = tPicLocate & "\"
        PicName = Right$(PicName, Len(PicName) - InStrRev(PicName, "\"))
        PicName = tPicLocate & PicName
    End If
    If Len(PicName) <> 0 And Len(Dir(PicName)) <> 0 Then
        Set img = New ImageFile
        img.LoadFile PicName
        If Not img Is Nothing Then
            Set GetAMC_PicShow = img.ARGBData.Picture(img.Width, img.Height)
        End If
    End If

End If

Exit Function
err:
MsgBox err.Description, vbCritical

End Function
Private Function GetAMC_Pic(col As Integer, row As Long) As Variant
Dim PicSize As Long
Dim PicName As String
Dim img As ImageFile
'Dim vec As Vector
Dim pb() As Byte
Dim N As Long
'Dim l As Long

On Error GoTo err

'name
'N = arrAMC(col - 2, row)
PicName = GetAMC_Str(col - 2, row)

'size
N = arrAMC(col - 1, row)
PicSize = GetInt(N)

'Picture
If PicSize > 0 Then
    'картинку из базы
    N = arrAMC(col, row)
    ReDim pb(PicSize - 1)
    Get ff, N, pb
    Set img = New ImageFile
'    Set vec = New Vector
'    vec.BinaryData = pb
        GetAMC_Pic = pb 'img.ARGBData.Picture(img.Width, img.Height)
Else
    If Len(tPicLocate) > 0 Then
        'обрезать путь к картинке и добавить пользовательский путь
        If Right$(tPicLocate, 1) <> "\" Then tPicLocate = tPicLocate & "\"
        PicName = Right$(PicName, Len(PicName) - InStrRev(PicName, "\"))
        PicName = tPicLocate & PicName
    End If
    If Len(PicName) <> 0 And Len(Dir(PicName)) <> 0 Then
        Set img = New ImageFile
        img.LoadFile PicName
        If Not img Is Nothing Then
            GetAMC_Pic = img.ARGBData.BinaryData '.Picture(img.Width, img.Height)
        End If
    End If

End If

Exit Function
err:
MsgBox err.Description, vbCritical

End Function

Private Function GetMDB_Pic() As StdPicture
Dim PicSize As Long
Dim img As ImageFile
Dim vec As Vector
Dim pb() As Byte

On Error GoTo err

'size
PicSize = inRs.Fields(ListItemIn.List(ListItemIn.ListIndex)).FieldSize
If PicSize = 0& Then Exit Function

ReDim pb(PicSize - 1)
pb() = inRs.Fields(ListItemIn.List(ListItemIn.ListIndex)).GetChunk(0, PicSize)

'Picture
Set img = New ImageFile
Set vec = New Vector
vec.BinaryData = pb
Set img = vec.ImageFile
If Not img Is Nothing Then
    Set GetMDB_Pic = img.ARGBData.Picture(img.Width, img.Height)
End If

Exit Function
err:
MsgBox "Ошибка при получении изображения:" & vbCrLf & err.Description, vbCritical

End Function
Private Sub ComDelLink_Click()
Dim i As Integer ', j As Integer

i = ListItemToDoView.ListIndex 'где мы
If i = -1 Then Exit Sub

DelLink i
End Sub
Private Sub DelAllLinks()
Dim i As Integer ', j As Integer

If ListItemToDoView.ListCount < 1 Then Exit Sub

For i = ListItemToDoView.ListCount - 1 To 0 Step -1
DelLink i
Next i

End Sub

Private Sub DelLink(nl As Integer)
'Dim i As Integer
Dim j As Integer
Dim a() As String 'масcив для In и SVC
Dim b() As String 'массив для объединенных входящих полей (IN)

'делим пару
a = Split(ListItemToDo.List(nl), " > ", 2, vbTextCompare)
b = Split(a(0), " | ", -1, vbTextCompare)

'вставить обратно пары
If UBound(b) <> 0 Then
    For j = 0 To UBound(b) - 1
    If Len(b(j)) <> 0 Then
        ListItemIn.AddItem b(j)
    End If
    Next j
Else
ListItemIn.AddItem a(0)
End If

ListItemSVC.AddItem a(1)

'убрать из туду
ListItemToDo.RemoveItem nl
ListItemToDoView.RemoveItem nl
'убрать из листы с разделителями
LstLSplit.RemoveItem nl
LstRSplit.RemoveItem nl

If ListItemToDo.ListCount < 1 Then 'нет связей
ComImport.Enabled = False 'запретить импорт
ComDelLink.Enabled = False 'запретить удалять
CmbTableIn.Enabled = True 'разрешить менять таблицу
End If
End Sub
Private Sub ComImport_Click()
Dim i As Integer
Set ImgTest = Nothing

DupInd = -1
'ищем идекс выбранной для дубликатов связи
For i = 0 To ListItemToDoView.ListCount - 1
    If LDupFields.Caption = ListItemToDoView.List(i) Then DupInd = i: Exit For
Next i

'xml
If IsXMLflag Then ImportFromXML

'mdb
If IsMDBflag Then ImportFromMDB

'amc
If IsAMCflag Then ImportFromAMC

'csv
If IsCSVflag Then ImportFromCSV

'rS.Close: DB.Close 'svc

'If errtype Then MsgBox "Были несовпадения типов ячеек баз." & vbCrLf & "Данные между ними не скопированы.", vbInformation, "2SVC"
MsgBox "Добавлено записей: " & Dobavleno & vbCrLf & "из: " & VsegoBilo, vbInformation, "2SVC"

End Sub



Private Sub ImportFromAMC()
Dim i As Integer, j As Long
Dim tmp As Variant
Dim ret As VbMsgBoxResult

On Error GoTo err

If rs.BOF And rs.EOF Then    'ваще пустая
Else
    rs.MoveLast    'добавить в конец svc базы
End If

Dobavleno = 0
VsegoBilo = AMCTotalRec
If VsegoBilo < 1 Then MsgBox "База-источник пуста.": Exit Sub

Screen.MousePointer = vbHourglass

PBar.Max = VsegoBilo
PBar.Value = 0
PBar.Visible = True

For j = 1 To VsegoBilo
PBar.Value = PBar.Value + 1
    
    If Not DupsFound_mdb(j) Then

        rs.AddNew
        For i = 0 To ListItemToDo.ListCount - 1  'по кол-ву туду
            ' получаем текущие данные из in базы
            tmp = GetCurrentInAMC(i, j)      ' там получаем и SVCPoleName

            If LCase$(SVCPoleName) <> rsIndex Then
                Select Case rs.Fields(SVCPoleName).Type
                Case 1: rs.Fields(SVCPoleName) = CBool(tmp)    'в "Логика"
                    'Case 2: Type2Str = "?"
                    'Case 3: Type2Str = "?"
                Case 4: rs.Fields(SVCPoleName) = StrToDbl(CStr((tmp)))     ' "Число"
                Case 5: rs.Fields(SVCPoleName) = Left$(tmp, 255)    ' "Деньги"
                    'Case 6: Type2Str = "?"
                    'Case 7: Type2Str = "?"
                Case 8: rs.Fields(SVCPoleName) = Left$(tmp, 255)    '"Дата"
                    'Case 9: Type2Str = "?"
                Case 10
                    If SVCPoleName = "Checked" Then
                        'rS.Fields(SVCPoleName) = IIf(tmp, "1", "0") ' "Чек"
                        If tmp <> vbNullString Then rs.Fields(SVCPoleName) = "1" Else rs.Fields(SVCPoleName) = vbNullString
                    Else
                        rs.Fields(SVCPoleName) = Left$(tmp, 255)   '    "Текст"
                    End If

                Case 11: rs.Fields(SVCPoleName) = tmp         ' "Объект OLE"
                Case 12: rs.Fields(SVCPoleName) = tmp    '            "MEMO"
                Case Else: rs.Fields(SVCPoleName) = Left$(tmp, 255)
                End Select
            End If    'не ключ
        Next i

        rs.Update
        Dobavleno = Dobavleno + 1
    End If    'dups

Next j

PBar.Visible = False
Screen.MousePointer = vbNormal

Exit Sub

err:
If err.Number <> 0 Then
    ret = MsgBox("ImportFromAMC" & vbCrLf & err.Description, vbOKCancel)
    If ret = vbCancel Then
        Screen.MousePointer = vbNormal
        PBar.Visible = False
        Exit Sub
    Else
        Resume Next
    End If
End If

End Sub
Private Function GetCurrentInAMC(ind As Integer, row As Long) As Variant
'ind c 1
Dim j As Integer
Dim i As Integer
Dim a() As String
Dim b() As String    'массив для объединенных входящих полей (IN)
Dim lSplit As String, rSplit As String
Dim tmpv As Variant

'делим пополам вход-выход
a = Split(ListItemToDo.List(ind), " > ", 2, vbTextCompare)
'делим входы
b = Split(a(0), " | ", -1, vbTextCompare)

i = 1
If UBound(b) <> 0 Then
    'если много входов
    For j = 0 To UBound(b) - 1
        If Len(b(j)) <> 0 Then
            'Получить данные объединяемых полей


            If j = 0 Then
                'первый in
                Select Case GetAmcTypeFromName(b(0))
                Case 1        'integer
                    GetCurrentInAMC = GetAMC_Int(GetAmcIndFromName(b(0)), row)          ' c 0,1
                Case 2        'boolean
                    GetCurrentInAMC = GetAMC_Bool(GetAmcIndFromName(b(0)), row)
                Case 3        'string
                    GetCurrentInAMC = GetAMC_Str(GetAmcIndFromName(b(0)), row)
                    '                    Case 4    'data
                    '                        GetCurrentInAMC = GetAMC_Pic(GetAmcIndFromName(b(0)), row)
                End Select

            Else
                'следующие in
                'получить сплиттеры конкретного ind в списке и позиции j
                lSplit = vbNullString: rSplit = vbNullString
                GetSplitters ind, i, lSplit, rSplit

                Select Case GetAmcTypeFromName(b(j))
                Case 1        'integer
                    tmpv = GetAMC_Int(GetAmcIndFromName(b(j)), row)
                Case 2        'boolean
                    tmpv = GetAMC_Bool(GetAmcIndFromName(b(j)), row)
                Case 3        'string
                    tmpv = GetAMC_Str(GetAmcIndFromName(b(j)), row)
                    '                    Case 4    'data
                    '                        tmpv = GetAMC_Pic(GetAmcIndFromName(b(j)), row)
                End Select
If Len(tmpv) <> 0 Then GetCurrentInAMC = GetCurrentInAMC & lSplit & tmpv & rSplit

                i = i + 1
            End If
        End If 'Len(b(j)) <> 0
    Next j
'GetCurrentInAMC = Trim$(GetCurrentInAMC)
If Left$(GetCurrentInAMC, 2) = lSplit Then
If Right$(GetCurrentInAMC, 1) = rSplit Then
GetCurrentInAMC = Right$(GetCurrentInAMC, Len(GetCurrentInAMC) - 2)
GetCurrentInAMC = Left$(GetCurrentInAMC, Len(GetCurrentInAMC) - 1)
End If
End If

Else
    'если один вход
    Select Case GetAmcTypeFromName(a(0))
    Case 1        'integer
        GetCurrentInAMC = GetAMC_Int(GetAmcIndFromName(a(0)), row)       ' c 0,1
    Case 2        'boolean
        GetCurrentInAMC = GetAMC_Bool(GetAmcIndFromName(a(0)), row)
    Case 3        'string
        GetCurrentInAMC = GetAMC_Str(GetAmcIndFromName(a(0)), row)
    Case 4        'data
        GetCurrentInAMC = GetAMC_Pic(GetAmcIndFromName(a(0)), row)
    End Select

End If

'Debug.Print GetCurrentInAMC
'выход
SVCPoleName = a(1)

End Function

Private Function GetAmcTypeFromName(s As String) As Integer
Dim i As Integer

For i = 0 To UBound(amcNamesArr) '- 1
    If amcNamesArr(i) = s Then
        GetAmcTypeFromName = amcTypesArr(i)
        Exit For
    End If
Next
End Function
Private Function GetAmcIndFromName(s As String) As Integer
Dim i As Integer
For i = 0 To UBound(amcNamesArr) '- 1
    If amcNamesArr(i) = s Then
        GetAmcIndFromName = i
        Exit For
    End If
Next
End Function
Private Function GetCurrentInMDB(ind As Integer) As String

Dim j As Integer
Dim i As Integer
Dim a() As String
Dim b() As String 'массив для объединенных входящих полей (IN)
Dim lSplit As String, rSplit As String

On Error Resume Next

'делим пополам вход-выход
a = Split(ListItemToDo.List(ind), " > ", 2, vbTextCompare)
'делим входы
b = Split(a(0), " | ", -1, vbTextCompare)

i = 1
If UBound(b) <> 0 Then
'если много входов
    For j = 0 To UBound(b) - 1
    If Len(b(j)) <> 0 Then
    'Получить данные объединяемых полей
    
        If Not IsNull(inRs.Fields(b(j)).Value) Then
            
            If j = 0 Then
                'первый in
                GetCurrentInMDB = inRs.Fields(b(0))
            Else
                'следующие in
                'получить сплиттеры конкретного ind в списке и позиции j
                lSplit = vbNullString: rSplit = vbNullString
                GetSplitters ind, i, lSplit, rSplit
                
                If Not IsNull(inRs.Fields(b(j))) Then
                If Len(inRs.Fields(b(j))) <> 0 Then GetCurrentInMDB = GetCurrentInMDB & lSplit & inRs.Fields(b(j)) & rSplit
                End If

                i = i + 1
            End If
            
        End If
  
    End If
    Next j

'почистить скобки
If Left$(GetCurrentInMDB, 2) = lSplit Then
If Right$(GetCurrentInMDB, 1) = rSplit Then
GetCurrentInMDB = Right$(GetCurrentInMDB, Len(GetCurrentInMDB) - 2)
GetCurrentInMDB = Left$(GetCurrentInMDB, Len(GetCurrentInMDB) - 1)
End If
End If

Else
'если один вход
                If Not IsNull(inRs.Fields(a(0))) Then
                GetCurrentInMDB = inRs.Fields(a(0))
                Else
                GetCurrentInMDB = vbNullString
                End If
End If

'выход
SVCPoleName = a(1)

err.Clear


End Function
Private Function GetCurrentInCSV(i As Integer, j As Long) As String
'i ind
'j row
Dim tmp As String
Dim p As Long

If pArr(i) = 1 Then     'item
    If lvCSV.ListItems(j).Text <> vbNullString Then tmp = lvCSV.ListItems(j).Text
Else                    'сабы
    If lvCSV.ListItems(j).SubItems(pArr(i) - 1) <> vbNullString Then tmp = lvCSV.ListItems(j).SubItems(pArr(i) - 1)
End If

If Not FromDupSub Then arrProcess(i) = True 'не помечать. если проверяем дупы
If i < TotalCsvFields Then
    For p = i + 1 To TotalCsvFields    'проверить еще одноименные
    If Len(lvCSV.ColumnHeaders(pArr(p)).Text) <> 0 Then
        If lvCSV.ColumnHeaders(pArr(p)).Text = lvCSV.ColumnHeaders(pArr(i)).Text Then
            If Not FromDupSub Then arrProcess(p) = True

            If pArr(p) = 1 Then 'item
                If lvCSV.ListItems(j).Text <> vbNullString Then tmp = tmp & TxtLSplit & lvCSV.ListItems(j).Text & TxtRSplit
            Else                 'сабы
                If lvCSV.ListItems(j).SubItems(pArr(p) - 1) <> vbNullString Then tmp = tmp & TxtLSplit & lvCSV.ListItems(j).SubItems(pArr(p) - 1) & TxtRSplit
            End If

        End If
    End If
    Next p
End If


'почистить скобки
If Left$(tmp, 2) = TxtLSplit Then
    If Right$(tmp, 1) = TxtRSplit Then
        tmp = Right$(tmp, Len(tmp) - 2)
        tmp = Left$(tmp, Len(tmp) - 1)
    End If
End If
SVCPoleName = lvCSV.ColumnHeaders(pArr(i)).Text
GetCurrentInCSV = tmp
End Function
Private Function GetCurrentInXML(ind As Integer) As String

Dim j As Integer
Dim i As Integer

Dim a() As String
Dim b() As String 'массив для объединенных входящих полей (IN)
Dim lSplit As String, rSplit As String


'делим пополам вход-выход
a = Split(ListItemToDo.List(ind), " > ", 2, vbTextCompare)
'делим входы
b = Split(a(0), " | ", -1, vbTextCompare)

i = 1
If UBound(b) <> 0 Then
'если много входов
    For j = 0 To UBound(b) - 1
    If Len(b(j)) <> 0 Then
    'Получить данные объединяемых полей
    
          
            If j = 0 Then
                'первый in
                GetCurrentInXML = GetXMLNamedValue(b(0))
            Else
                'следующие in
                'получить сплиттеры конкретного ind в списке и позиции j
                lSplit = vbNullString: rSplit = vbNullString
                GetSplitters ind, i, lSplit, rSplit
                'соединить
                If Len(GetXMLNamedValue(b(j))) <> 0 Then GetCurrentInXML = GetCurrentInXML & lSplit & GetXMLNamedValue(b(j)) & rSplit
                i = i + 1
            End If
            
  
    End If
    Next j
'почистить скобки
If Left$(GetCurrentInXML, 2) = lSplit Then
If Right$(GetCurrentInXML, 1) = rSplit Then
GetCurrentInXML = Right$(GetCurrentInXML, Len(GetCurrentInXML) - 2)
GetCurrentInXML = Left$(GetCurrentInXML, Len(GetCurrentInXML) - 1)
End If
End If

Else
'если один вход
    
    GetCurrentInXML = GetXMLNamedValue(a(0))
End If

'out
SVCPoleName = a(1)

End Function
Private Function GetXMLNamedValue(sName As String) As String
Dim XML_1 As Object
Dim XML_a As Object

On Error Resume Next    ' нет ноды или аттра

If ItIsNode(sName) Then
    'node
    For Each XML_1 In XML_tmp.childNodes
        If XML_1.nodeName = sName Then GetXMLNamedValue = XML_1.nodeTypedValue
    Next
Else
    'attr
    If XML_tmp.Attributes.length > 0 Then
        Set XML_a = XML_tmp.Attributes
        GetXMLNamedValue = XML_a.getNamedItem(sName).nodeValue
    End If

End If

err.Clear

End Function
Private Function ItIsNode(sName As String) As Boolean
Dim XML_1 As Object
Dim XML_2 As Object

Set XML_1 = XML_tmp.childNodes
For Each XML_2 In XML_1
    If XML_2.nodeName = sName Then ItIsNode = True: Exit For
Next

End Function

Public Sub ImportFromMDB()
Dim i As Integer
Dim tmp As String
Dim ret As VbMsgBoxResult

On Error GoTo err

If rs.BOF And rs.EOF Then    'ваще пустая
Else
    rs.MoveLast    'добавить в конец svc базы
End If

Dobavleno = 0
VsegoBilo = 0

If inRs.BOF And inRs.EOF Then    'ваще пустая
    MsgBox "База-источник пуста."
    Exit Sub
Else
    inRs.MoveLast
    inRs.MoveFirst
End If

VsegoBilo = inRs.RecordCount

Screen.MousePointer = vbHourglass

PBar.Max = VsegoBilo
PBar.Value = 0
PBar.Visible = True

Do While Not inRs.EOF
PBar.Value = PBar.Value + 1

    If Not DupsFound_mdb Then


        rs.AddNew
        For i = 0 To ListItemToDo.ListCount - 1    'по кол-ву туду
            ' получаем текущие данные из in базы
            tmp = GetCurrentInMDB(i)    ' там получаем и SVCPoleName

            If LCase$(SVCPoleName) <> rsIndex Then 'игнор ключу
                Select Case rs.Fields(SVCPoleName).Type
                Case 1: rs.Fields(SVCPoleName) = CBool(tmp)    'в "Логика"
                    'Case 2: Type2Str = "?"
                    'Case 3: Type2Str = "?"
                Case 4: rs.Fields(SVCPoleName) = StrToDbl(tmp)    ' "Число"
                Case 5: rs.Fields(SVCPoleName) = Left$(tmp, 255)    ' "Деньги"
                    'Case 6: Type2Str = "?"
                    'Case 7: Type2Str = "?"
                Case 8: rs.Fields(SVCPoleName) = Left$(tmp, 255)    '"Дата"
                    'Case 9: Type2Str = "?"
                Case 10
                    If SVCPoleName = "Checked" Then
                        'rS.Fields(SVCPoleName) = IIf(tmp, "1", "0") ' "Чек"
                        If tmp <> vbNullString Then rs.Fields(SVCPoleName) = "1" Else rs.Fields(SVCPoleName) = vbNullString
                    Else
                        rs.Fields(SVCPoleName) = Left$(tmp, 255)    '    "Текст"
                    End If

                Case 11: rs.Fields(SVCPoleName) = tmp         ' "Объект OLE"
                Case 12: rs.Fields(SVCPoleName) = tmp    '            "MEMO"
                Case Else: rs.Fields(SVCPoleName) = Left$(tmp, 255)
                End Select
            End If    'не ключ
        Next i

        rs.Update
        Dobavleno = Dobavleno + 1
    End If    'dups
    inRs.MoveNext

Loop

PBar.Visible = False
Screen.MousePointer = vbNormal

Exit Sub

err:
If err.Number <> 0 Then
    ret = MsgBox("ImportFromMDB" & vbCrLf & err.Description, vbOKCancel)
    If ret = vbCancel Then
        Screen.MousePointer = vbNormal
        PBar.Visible = False
        Exit Sub
    Else
        Resume Next
    End If
End If


End Sub
Public Sub ImportFromXML()
Dim i As Integer ', j As Integer
Dim tmp As String
'Dim ret As VbMsgBoxResult

Screen.MousePointer = vbHourglass

'On Error GoTo err
On Error Resume Next

If rs.BOF And rs.EOF Then    'ваще пустая
Else
    rs.MoveLast    'добавить в конец svc базы
End If

Dobavleno = 0
VsegoBilo = 0


'цикл по нодам
Set XML_Child = XML_Root.getElementsByTagName(CmbTableIn.Text)

PBar.Max = XML_Child.length
PBar.Value = 0
PBar.Visible = True

For Each XML_tmp In XML_Child
    VsegoBilo = VsegoBilo + 1
    PBar.Value = VsegoBilo
    
    If Not DupsFound_mdb Then

        rs.AddNew
        'цикл по кол-ву туду
        For i = 0 To ListItemToDo.ListCount - 1

            ' получаем текущие данные из in базы
            tmp = GetCurrentInXML(i)    ' там получаем и SVCPoleName

            If LCase$(SVCPoleName) <> rsIndex Then
                Select Case rs.Fields(SVCPoleName).Type
                Case 1: rs.Fields(SVCPoleName) = CBool(tmp)    ' "Логика"
                    'Case 2: Type2Str = "?"
                    'Case 3: Type2Str = "?"
                Case 4: rs.Fields(SVCPoleName) = StrToDbl(tmp)    ' "Число"
                Case 5: rs.Fields(SVCPoleName) = Left$(tmp, 255)    ' "Деньги"
                    'Case 6: Type2Str = "?"
                    'Case 7: Type2Str = "?"
                Case 8: rs.Fields(SVCPoleName) = Left$(tmp, 255)    '"Дата"
                    'Case 9: Type2Str = "?"
                Case 10
                    If SVCPoleName = "Checked" Then
                        'rS.Fields(SVCPoleName) = IIf(tmp, "1", "0") ' "Чек"
                        If tmp <> vbNullString Then rs.Fields(SVCPoleName) = "1" Else rs.Fields(SVCPoleName) = vbNullString
                    Else
                        rs.Fields(SVCPoleName) = Left$(tmp, 255)    ' "Текст"
                    End If
                Case 11: rs.Fields(SVCPoleName) = tmp    ' "Объект OLE"
                Case 12: rs.Fields(SVCPoleName) = tmp    ' "MEMO"
                
                Case Else: rs.Fields(SVCPoleName) = Left$(tmp, 255) 'не правильно это, лучше все типы применить
                End Select
            End If    'не ключевое

        Next i

        rs.Update
        Dobavleno = Dobavleno + 1
    End If    'dup
Next

PBar.Visible = False

err:
If err.Number <> 0 Then
    Debug.Print "ImportFromXML" & vbCrLf & err.Description
End If
Screen.MousePointer = vbNormal
End Sub

Private Sub ComOld2NewSVC_Click()
Dim i As Integer, j As Integer
Dim sinonims() As String
Dim SinonimsNum As Integer
Dim added As Boolean

If ListItemIn.ListCount < 1 Or ListItemSVC.ListCount < 1 Then Exit Sub

Do While ListItemIn.ListCount - i > 0
    ListItemIn.Selected(i) = True
    added = False

    SinonimsNum = GetSinonims(sinonims, ListItemIn.List(i))

    For j = 0 To SinonimsNum

        'пометить такой-же
        'SendMessage ListItemSVC.hwnd, LB_SELECTSTRING, -1, ByVal ListItemIn.List(i)
        SendMessage ListItemSVC.hwnd, LB_SELECTSTRING, -1, ByVal sinonims(j)

        If ListItemSVC.ListIndex > -1 Then
            'если помечен, то кликнуть и добавить пару
            ListItemSVC_Click
            ComAddLinkClick
            added = True
            Exit For
        End If
    Next j

    If Not added Then i = i + 1

Loop

End Sub
Private Function GetSinonims(arr() As String, word As String) As Integer
'возврат колва синонимов -1
'синонимы в массиве

'сначала (case) ищем в in базе

GetSinonims = -1
Select Case LCase$(word)

Case "acter", "actors", "starring", "actor"
arr = Split("acter,actors,starring,actor", ",")
GetSinonims = UBound(arr)

Case "annotation", "description"
arr = Split("annotation,description", ",")
GetSinonims = UBound(arr)

Case "audio", "audioformat", "audioinfo"
arr = Split("audio,audiobitrate,audioformat,audioinfo", ",")
GetSinonims = UBound(arr)

Case "cdn", "disks", "mediacount"
arr = Split("cdn,disks,mediacount", ",")
GetSinonims = UBound(arr)

Case "checked"
arr = Split("checked", ",")
GetSinonims = UBound(arr)

Case "country", "studio"
arr = Split("country,studio", ",")
GetSinonims = UBound(arr)

Case "coverpath", "coverurl"
arr = Split("coverpath,coverurl", ",")
GetSinonims = UBound(arr)

Case "debtor", "loaner", "borrower"
arr = Split("debtor,loaner,borrower", ",")
GetSinonims = UBound(arr)

Case "director"
arr = Split("director", ",")
GetSinonims = UBound(arr)

Case "filelen", "size", "filesize"
arr = Split("filelen,size,filesize", ",")
GetSinonims = UBound(arr)

Case "filename", "file", "files"
arr = Split("filename,file,files", ",")
GetSinonims = UBound(arr)

Case "fps", "framerate"
arr = Split("fps,framerate", ",")
GetSinonims = UBound(arr)

Case "frontface", "cover", "screen40", "picture"
arr = Split("frontface,cover,screen", ",")
GetSinonims = UBound(arr)

Case "genre", "category"
arr = Split("genre,category", ",")
GetSinonims = UBound(arr)

Case "label", "media", "medialabel", "cdboxid"
arr = Split("label,media,medialabel,cdboxid", ",")
GetSinonims = UBound(arr)

Case "language", "languages"
arr = Split("language,languages", ",")
GetSinonims = UBound(arr)

Case "mediatype"
arr = Split("mediatype", ",")
GetSinonims = UBound(arr)

Case "moviename", "title", "movietitle", "name", "translatedtitle"
arr = Split("moviename,title,movietitle,name,translatedtitle", ",")
GetSinonims = UBound(arr)

Case "movieurl", "url"
arr = Split("movieurl,url", ",")
GetSinonims = UBound(arr)

Case "other", "comments"
arr = Split("other,comments", ",")
GetSinonims = UBound(arr)

Case "rating", "score"
arr = Split("rating,score", ",")
GetSinonims = UBound(arr)

Case "resolution"
arr = Split("resolution", ",")
GetSinonims = UBound(arr)

Case "snapshot", "screen"
arr = Split("snapshot,screen", ",")
GetSinonims = UBound(arr)

Case "sndisk", "serial", "cdserial"
arr = Split("sndisk,serial,cdserial", ",")
GetSinonims = UBound(arr)

Case "subtitle", "subtitles"
arr = Split("subtitle,subtitles", ",")
GetSinonims = UBound(arr)

Case "time", "length"
arr = Split("time,length", ",")
GetSinonims = UBound(arr)

Case "video", "videoformat", "videoinfo"
arr = Split("video,videoformat,videoinfo", ",")
GetSinonims = UBound(arr)

Case "year"
arr = Split("year", ",")
GetSinonims = UBound(arr)

Case "snapshot1", "snapshot2", "snapshot3"
arr = Split("snapshot1,snapshot2,snapshot3", ",")
GetSinonims = UBound(arr)

Case "name"
arr = Split("name", ",")
GetSinonims = UBound(arr)

Case "bio"
arr = Split("bio", ",")
GetSinonims = UBound(arr)

Case "face"
arr = Split("face", ",")
GetSinonims = UBound(arr)

End Select

End Function
Private Sub ComOpenIn_Click()
   Dim cd As New cCommonDialog
   Dim sFile As String
   Dim tdData As DAO.TableDef


If (cd.VBGetOpenFileName( _
      sFile, _
      Filter:="MS Access|*.mdb;*.amm|XML files|*.xml|AntMovieCatalog (v3.5)|*.amc|CSV Files (*.csv, *.txt)|*.csv;*.txt|All Files (*.*)|*.*|All Supported |*.mdb;*.xml;*.amm;*.amc;*.csv", _
      FilterIndex:=6, _
      DefaultExt:="mdb", _
      DlgTitle:="Открыть базу-источник", _
      Owner:=Me.hwnd)) Then
      newbdname = sFile
End If
Me.Refresh

If sFile <> vbNullString Then

'чистка
DelAllLinks 'убить все связи
ClearObj
ListItemIn.Clear


Set ImgTest = Nothing
Set inRs = Nothing
Set newDB = Nothing




IsXMLflag = False: IsMDBflag = False: IsAMCflag = False: IsCSVflag = False
CmbTableIn.Enabled = True
ComOld2NewSVC.Enabled = True
ListItemIn.Enabled = True
CmbTableIn.Clear
lvCSV.Visible = False
cbCsvSvc.Visible = False
CmbTableIn.Enabled = True
ImgTest.Visible = True
ComImport.Enabled = False
chCSV.Visible = False


'cd.FilterIndex

Select Case UCase$(Right$(newbdname, 3))
Case "XML"
    IsXMLflag = True
    DoEvents
    Screen.MousePointer = vbHourglass
        ProcessRoot newbdname
    Screen.MousePointer = vbNormal
    CmbTableIn.Text = "Выбрать ноду ..."
'    ComCheck.Enabled = True
    
Case "MDB", "AMM"
    IsMDBflag = True
    ComAddLink.Enabled = False
        Set newDB = DBEngine.OpenDatabase(newbdname, False, True)
        For Each tdData In newDB.TableDefs
            If tdData.Attributes = 0 Then
                CmbTableIn.AddItem tdData.Name
            End If
        Next tdData
    CmbTableIn.Text = "Выбрать таблицу ..."
'    ComCheck.Enabled = False

Case "AMC"
    IsAMCflag = True
    tPicLocate.Enabled = True
    CmbTableIn.Enabled = False
    
    AddAMCFields
    
Case "CSV", "TXT"
    IsCSVflag = True

    lvCSV.ListItems.Clear
    lvCSV.ColumnHeaders.Clear
    lvCSV.ZOrder 0
    lvCSV.Visible = True
    cbCsvSvc.Visible = True
    ImgTest.Visible = False
    chCSV.Visible = True
If Not (rs Is Nothing) Then ComImport.Enabled = True
    GetCsvFieldsToList
    'FillComboWithSvcFieldsNames
    chCSV_Click
    
End Select
End If 'bdname <> vbNullString


Set cd = Nothing
Exit Sub
err:
End Sub
Private Sub GetCsvFieldsToList()
Dim l As String
Dim i As Integer, j As Integer
Dim v() As String

Dim initLV As Boolean    'нужны хедеры

On Error GoTo err
Screen.MousePointer = vbHourglass

ff = FreeFile
Open newbdname For Input Lock Write As #ff
initLV = True

LockWindowUpdate lvCSV.hwnd

Do While Not EOF(ff)
'DoEvents нет
    Line Input #ff, l
    
    'l = Replace(l, tCsvDelim & """""" & tCsvDelim, tCsvDelim & tCsvDelim)
    l = Replace(l, """""", "")

    TotalCsvFields = ParseCSV01(l, tCsvDelim, v)

    If initLV Then
        initLV = False
        For j = 1 To TotalCsvFields
            lvCSV.ColumnHeaders.Add j, , , 1500
        Next
    End If

    i = i + 1
    'If i > 12 Then Exit Do

    lvCSV.ListItems.Add i, , v(0)
    For j = 1 To lvCSV.ColumnHeaders.Count - 1
        lvCSV.ListItems(i).SubItems(j) = v(j)
    Next j

Loop

LockWindowUpdate 0
Screen.MousePointer = vbNormal
SetListboxScrollbar ListItemIn

Exit Sub
err:
LockWindowUpdate 0
Screen.MousePointer = vbNormal

End Sub
Private Sub ImportFromCSV()
Dim j As Long, i As Integer
Dim FieldAssigned As Boolean
Dim ret As VbMsgBoxResult

On Error GoTo err

If rs.BOF And rs.EOF Then    'ваще пустая
Else
    rs.MoveLast    'добавить в конец svc базы
End If

Dobavleno = 0: VsegoBilo = 0
'If VsegoBilo < 1 Then MsgBox "База-источник пуста.": Exit Sub
'проверить. задано ли хоть 1 поле
For i = 1 To lvCSV.ColumnHeaders.Count
    If lvCSV.ColumnHeaders(i).Text <> vbNullString Then
        FieldAssigned = True
        Exit For
    End If
Next i

'если не задано выйти
If Not FieldAssigned Then
    MsgBox "Дайте названия нужным колонкам.", vbCritical
    Exit Sub
End If

Screen.MousePointer = vbHourglass

'заполнить массис соответствия   индекс в списке - позиция
'индекс массива - позиция, значение - индекс поля списка

ReDim pArr(TotalCsvFields)
For i = 1 To TotalCsvFields
    pArr(lvCSV.ColumnHeaders(i).Position) = i
Next i

ReDim arrProcess(TotalCsvFields)    'true - если уже поле обработано

Dim tmp As String

Dim Itm As ListItem
For Each Itm In lvCSV.ListItems
    With Itm
        If .Checked Then
            VsegoBilo = VsegoBilo + 1
        End If
    End With
Next

For j = 1 To VsegoBilo
    'по строкам списка

If lvCSV.ListItems(j).Checked Then 'помечен
    If Not DupsFound_mdb(j) Then
        
        rs.AddNew
        For i = 1 To TotalCsvFields    'по колонкам
            If Not arrProcess(i) Then
            
'            If Len(lvCSV.ColumnHeaders(i)) <> 0 Then 'если задано имя
            If Len(lvCSV.ColumnHeaders(pArr(i))) <> 0 Then 'если задано имя
            
            tmp = GetCurrentInCSV(i, j)

                'записать в базу
                SVCPoleName = lvCSV.ColumnHeaders(pArr(i)).Text

                'If LCase$(SVCPoleName) <> rSIndex Then 'игнор ключу
                Select Case rs.Fields(SVCPoleName).Type
                Case 1: rs.Fields(SVCPoleName) = CBool(tmp)    'в "Логика"
                    'Case 2: Type2Str = "?"
                    'Case 3: Type2Str = "?"
                Case 4: rs.Fields(SVCPoleName) = StrToDbl(tmp)    ' "Число"
                Case 5: rs.Fields(SVCPoleName) = Left$(tmp, 255)    ' "Деньги"
                    'Case 6: Type2Str = "?"
                    'Case 7: Type2Str = "?"
                Case 8: rs.Fields(SVCPoleName) = Left$(tmp, 255)    '"Дата"
                    'Case 9: Type2Str = "?"
                Case 10
                    If SVCPoleName = "Checked" Then
                        'rS.Fields(SVCPoleName) = IIf(tmp, "1", "0") ' "Чек"
                        If tmp <> vbNullString Then rs.Fields(SVCPoleName) = "1" Else rs.Fields(SVCPoleName) = vbNullString
                    Else
                        rs.Fields(SVCPoleName) = Left$(tmp, 255)    '    "Текст"
                    End If

                Case 11: rs.Fields(SVCPoleName) = tmp         ' "Объект OLE"
                Case 12: rs.Fields(SVCPoleName) = tmp    '            "MEMO"
                Case Else: rs.Fields(SVCPoleName) = Left$(tmp, 255)
                End Select
                '            End If    'не ключ
            End If    'если задано имя
            End If    'Not arrProcess(i)
            'Debug.Print lvCSV.ColumnHeaders(pArr(j)).Text
        Next i    'след колонка
        ReDim arrProcess(TotalCsvFields)    'почистить
        Dobavleno = Dobavleno + 1
        rs.Update
    End If 'dups
    End If    'Checked

Next j

Screen.MousePointer = vbNormal
Exit Sub

err:
If err.Number <> 0 Then
    ret = MsgBox("ImportFromCSV" & vbCrLf & err.Description, vbOKCancel)
    If ret = vbCancel Then
        Screen.MousePointer = vbNormal
        PBar.Visible = False
        Exit Sub
    Else
        Resume Next
    End If
End If

End Sub
Private Sub AddAMCFields()

Dim tmpl As Long
'Dim tmps As String
'Dim tmpi As Integer
Dim MNumber As Long
ff = FreeFile
Dim N As Long
Dim PicSize As Long
Dim i As Long
Dim HasRecords As Boolean

ReDim arrAMC(32, 0)
AMCTotalRec = 0

On Error GoTo err

Screen.MousePointer = vbHourglass

Open newbdname For Binary Lock Write As ff
N = 82  'пропуск хедера
Do While Not EOF(ff)
i = i + 1
'number
AMCTotalRec = MNumber
MNumber = GetInt(N)
If MNumber < 1 Then Exit Do ' нет записей

ReDim Preserve arrAMC(32, i) 'MNumber)

arrAMC(0, i) = N: N = N + 4
'Дата занесения - число от 14.09.1752 ?
arrAMC(1, i) = N: N = N + 4
'Рейтинг
arrAMC(2, i) = N: N = N + 4
'год
arrAMC(3, i) = N: N = N + 4
'Length мин
arrAMC(4, i) = N: N = N + 4
'VideoBitrate
arrAMC(5, i) = N: N = N + 4
'AudioBitrate:
arrAMC(6, i) = N: N = N + 4
'Disks:
arrAMC(7, i) = N: N = N + 4
'Expotr
arrAMC(8, i) = N: N = N + 1

'Label
tmpl = GetInt(N)
arrAMC(9, i) = N: N = N + 4 + tmpl
'MediaType
tmpl = GetInt(N): arrAMC(10, i) = N: N = N + 4 + tmpl
'Source
tmpl = GetInt(N): arrAMC(11, i) = N: N = N + 4 + tmpl
'Borrower
tmpl = GetInt(N): arrAMC(12, i) = N: N = N + 4 + tmpl
'OriginalTitle
tmpl = GetInt(N): arrAMC(13, i) = N: N = N + 4 + tmpl
'TranslatedTitle
tmpl = GetInt(N): arrAMC(14, i) = N: N = N + 4 + tmpl
'Director
tmpl = GetInt(N): arrAMC(15, i) = N: N = N + 4 + tmpl
'Producer
tmpl = GetInt(N): arrAMC(16, i) = N: N = N + 4 + tmpl
'Country
tmpl = GetInt(N): arrAMC(17, i) = N: N = N + 4 + tmpl
'Category
tmpl = GetInt(N): arrAMC(18, i) = N: N = N + 4 + tmpl
'Actors
tmpl = GetInt(N): arrAMC(19, i) = N: N = N + 4 + tmpl
'URL
tmpl = GetInt(N): arrAMC(20, i) = N: N = N + 4 + tmpl
'Description
tmpl = GetInt(N): arrAMC(21, i) = N: N = N + 4 + tmpl
'Comments
tmpl = GetInt(N): arrAMC(22, i) = N: N = N + 4 + tmpl
'VideoFormat
tmpl = GetInt(N): arrAMC(23, i) = N: N = N + 4 + tmpl
'AudioFormat
tmpl = GetInt(N): arrAMC(24, i) = N: N = N + 4 + tmpl
'Resolution
tmpl = GetInt(N): arrAMC(25, i) = N: N = N + 4 + tmpl
'Framerate
tmpl = GetInt(N): arrAMC(26, i) = N: N = N + 4 + tmpl
'Languages
tmpl = GetInt(N): arrAMC(27, i) = N: N = N + 4 + tmpl
'Subtitles
tmpl = GetInt(N): arrAMC(28, i) = N: N = N + 4 + tmpl
'Size
tmpl = GetInt(N): arrAMC(29, i) = N: N = N + 4 + tmpl
'PictureName
tmpl = GetInt(N): arrAMC(30, i) = N: N = N + 4 + tmpl

'PictureSize
PicSize = GetInt(N): arrAMC(31, i) = N: N = N + 4
'Picture
arrAMC(32, i) = N: N = N + PicSize

HasRecords = True
Loop

If Not HasRecords Then Screen.MousePointer = vbNormal: Exit Sub 'нет записей

'
ListItemIn.AddItem "Number", 0
    amcNamesArr(0) = "Number": amcTypesArr(0) = 1 'Число
ListItemIn.AddItem "Date", 1
    amcNamesArr(1) = "Date": amcTypesArr(1) = 1 'Число
ListItemIn.AddItem "Rating", 2
    amcNamesArr(2) = "Rating": amcTypesArr(2) = 1 'Число
ListItemIn.AddItem "Year", 3
    amcNamesArr(3) = "Year": amcTypesArr(3) = 1 'Число
ListItemIn.AddItem "Length", 4
    amcNamesArr(4) = "Length": amcTypesArr(4) = 1 'Число
ListItemIn.AddItem "VideoBitrate", 5
    amcNamesArr(5) = "VideoBitrate": amcTypesArr(5) = 1 'Число
ListItemIn.AddItem "AudioBitrate", 6
    amcNamesArr(6) = "AudioBitrate": amcTypesArr(6) = 1 'Число
ListItemIn.AddItem "Disks", 7
    amcNamesArr(7) = "Disks": amcTypesArr(7) = 1 'Число
    
ListItemIn.AddItem "Export", 8
    amcNamesArr(8) = "Export": amcTypesArr(8) = 2 'Логика
    
ListItemIn.AddItem "Media", 9
    amcNamesArr(9) = "Media": amcTypesArr(9) = 3 'Строка
ListItemIn.AddItem "MediaType", 10
    amcNamesArr(10) = "MediaType": amcTypesArr(10) = 3 'Строка
ListItemIn.AddItem "Source", 11
    amcNamesArr(11) = "Source": amcTypesArr(11) = 3 'Строка
ListItemIn.AddItem "Borrower", 12
    amcNamesArr(12) = "Borrower": amcTypesArr(12) = 3 'Строка
ListItemIn.AddItem "OriginalTitle", 13
    amcNamesArr(13) = "OriginalTitle": amcTypesArr(13) = 3 'Строка
ListItemIn.AddItem "TranslatedTitle", 14
    amcNamesArr(14) = "TranslatedTitle": amcTypesArr(14) = 3 'Строка
ListItemIn.AddItem "Director", 15
    amcNamesArr(15) = "Director": amcTypesArr(15) = 3 'Строка
ListItemIn.AddItem "Producer", 16
    amcNamesArr(16) = "Producer": amcTypesArr(16) = 3 'Строка
ListItemIn.AddItem "Country", 17
    amcNamesArr(17) = "Country": amcTypesArr(17) = 3 'Строка
ListItemIn.AddItem "Category", 18
    amcNamesArr(18) = "Category": amcTypesArr(18) = 3 'Строка
ListItemIn.AddItem "Actors", 19
    amcNamesArr(19) = "Actors": amcTypesArr(19) = 3 'Строка
ListItemIn.AddItem "URL", 20
    amcNamesArr(20) = "URL": amcTypesArr(20) = 3 'Строка
ListItemIn.AddItem "Description", 21
    amcNamesArr(21) = "Description": amcTypesArr(21) = 3 'Строка
ListItemIn.AddItem "Comments", 22
    amcNamesArr(22) = "Comments": amcTypesArr(22) = 3 'Строка
ListItemIn.AddItem "VideoFormat", 23
    amcNamesArr(23) = "VideoFormat": amcTypesArr(23) = 3 'Строка
ListItemIn.AddItem "AudioFormat", 24
    amcNamesArr(24) = "AudioFormat": amcTypesArr(24) = 3 'Строка
ListItemIn.AddItem "Resolution", 25
    amcNamesArr(25) = "Resolution": amcTypesArr(25) = 3 'Строка
ListItemIn.AddItem "Framerate", 26
    amcNamesArr(26) = "Framerate": amcTypesArr(26) = 3 'Строка
ListItemIn.AddItem "Languages", 27
    amcNamesArr(27) = "Languages": amcTypesArr(27) = 3 'Строка
ListItemIn.AddItem "Subtitles", 28
    amcNamesArr(28) = "Subtitles": amcTypesArr(28) = 3 'Строка
ListItemIn.AddItem "Size", 29
    amcNamesArr(29) = "Size": amcTypesArr(29) = 3 'Строка
ListItemIn.AddItem "PictureName", 30
    amcNamesArr(30) = "PictureName": amcTypesArr(30) = 3 'Строка
    
ListItemIn.AddItem "PictureSize", 31
    amcNamesArr(31) = "PictureSize": amcTypesArr(31) = 1 'Число
    
ListItemIn.AddItem "Picture", 32
    amcNamesArr(32) = "Picture": amcTypesArr(32) = 4 'Данные

Screen.MousePointer = vbNormal
Exit Sub
err:
Screen.MousePointer = vbNormal
MsgBox err.Description
End Sub

Private Sub OpenMyBase()
Dim cd As New cCommonDialog
Dim sFile As String
'Dim tdData As DAO.TableDef
Dim i As Integer

If (cd.VBGetOpenFileName( _
    sFile, _
    Filter:="MDB Files (*.mdb)|*.mdb|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    DlgTitle:="Открыть базу SurVideoCatalog", _
    HideReadOnly:=True, _
    DefaultExt:="mdb", _
    Owner:=Me.hwnd)) Then
    bdname = sFile
End If
Me.SetFocus

If sFile <> vbNullString Then

    DelAllLinks 'убить все связи
    'почистка
    'ClearObj - много лишнего чистится
    ListItemSVC.Clear
    ListItemToDoView.Clear
    ListItemToDo.Clear
    Set rs = Nothing
    Set DB = Nothing

    ComAddLink.Enabled = False
    Set DB = DBEngine.OpenDatabase(bdname, False, False)

    On Error Resume Next
    Set rs = DB.OpenRecordset(MyTableName, dbOpenTable)
    If err.Number <> 0 Then
        MsgBox bdname & " уже открыта или не SVC-база!"
        DB.Close
        Exit Sub
    End If
    err.Clear
    On Error GoTo 0

    'CSV
    If IsCSVflag Then ComImport.Enabled = True: Exit Sub

    'в список
    For i = 0 To rs.Fields.Count - 1
        If rs.Fields(i).Name <> vbNullString Then
            ListItemSVC.AddItem rs.Fields(i).Name
            ' ListItemSVC.ItemData(i) = rS.Fields(i).Type 'не работает с ундо
        Else
            ListItemSVC.AddItem vbNullString
        End If
    Next i
End If    'bdname <> vbNullString

'для ундо
ReDim CollDBSVC(ListItemSVC.ListCount)
ReDim iUndoSVC(ListItemSVC.ListCount)

End Sub

Private Sub ComOpenSVC_Click(Index As Integer)
Select Case Index
Case 0
    MyTableName = "Storage"
    rsIndex = "key" 'маленькими
Case 1
    MyTableName = "Acter"
    rsIndex = "key" '"keyact"
End Select

OpenMyBase

FillComboWithSvcFieldsNames

End Sub

Private Sub ComReadMe_Click()
Shell "notepad.exe " & App.Path & "\converter.txt", vbNormalFocus
End Sub

Private Sub Form_Load()

ReDim CollDBIn(0)
ReDim CollDBSVC(0)
'Me.ScaleMode = 3 'pixel для дроп комбо
Dim X As Long, Y As Long, w As Long, h As Long
X = ScaleX(cbCsvSvc.Left, vbTwips, vbPixels)
Y = ScaleY(cbCsvSvc.Top, vbTwips, vbPixels)
w = ScaleY(cbCsvSvc.Width, vbTwips, vbPixels)
SetWindowPos cbCsvSvc.hwnd, 0, X, Y, w, 500, 0

LCID = GetSystemDefaultLCID

End Sub

Private Sub FillComboWithSvcFieldsNames()
Dim i As Integer
If rs Is Nothing Then Exit Sub

cbCsvSvc.Clear
For i = 0 To rs.Fields.Count - 1

Select Case rs.Fields(i).Name
Case "key", "Face", "SnapShot1", "SnapShot2", "SnapShot3", "FrontFace"
Case Else
    cbCsvSvc.AddItem rs.Fields(i).Name
End Select

Next i

cbCsvSvc.Text = "Доступные поля:"

End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub 'спрятан, а то Me.ScaleWidth = 0
lvCSV.Move 120, 600, Me.Width - 300, 4200
'lvCSV.Move CmbTableIn.Left, CmbTableIn.Top - 3, Me.ScaleWidth - 15, Me.ScaleHeight - 245


End Sub

Private Sub Form_Unload(Cancel As Integer)
ClearObj
End Sub

Private Sub ListItemin_Click()
Dim a() As String
'Dim tmp As String

'Debug.Print "1click"
If ListItemIn.ListIndex < 0 Then Exit Sub
If ListItemSVC.SelCount Then ComAddLink.Enabled = True

'xml
If IsXMLflag Then
    TextIn = "XML"
    Exit Sub
End If

If IsMDBflag Then

'проверить обращение по имени поля (мб умляуты) - не сделано
On Error Resume Next 'нет rs
TextIn = "?" 'по умолчанию
    TextIn = Type2Str(inRs.Fields(ListItemIn.List(ListItemIn.ListIndex)).Type)

    If ListItemSVC.SelCount Then
        'in out
        ChTypes inRs.Fields(ListItemIn.List(ListItemIn.ListIndex)).Type, rs.Fields(ListItemSVC.List(ListItemSVC.ListIndex)).Type
    ElseIf ListItemToDoView.SelCount Then
        'in view
        a = Split(ListItemToDo.List(ListItemIn.ListIndex), " > ", 2, vbTextCompare)
        If UBound(a) > 0 Then
        ChTypes inRs.Fields(ListItemIn.List(ListItemIn.ListIndex)).Type, rs.Fields(a(1)).Type
        End If
    Else
        ChTypes 1, 1 'почернить
    End If

End If

If IsAMCflag Then
Select Case GetAmcTypeFromName(ListItemIn.List(ListItemIn.ListIndex))
Case 1
TextIn = "Число"
Case 2
TextIn = "Логика"
Case 3
TextIn = "Строка"
Case 4
TextIn = "Двоичные данные"
End Select
End If


err.Clear
End Sub

Private Sub ListItemIn_DblClick()
'Dim res As Long
'поискать то-же в svc и кликнуть
'TextTemp = ListItemIn.List(ListItemIn.ListIndex)
SendMessage ListItemSVC.hwnd, LB_SELECTSTRING, -1, ByVal ListItemIn.List(ListItemIn.ListIndex) 'TextTemp.Text
If ListItemSVC.ListIndex > -1 Then ListItemSVC_Click
End Sub

Private Sub ListItemIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
If ComAddLink.Enabled Then ComAddLink_Click

End If

End Sub

Private Sub ListItemSVC_Click()
If ListItemSVC.ListIndex = -1 Then Exit Sub

If ListItemToDoView.ListIndex > -1 Then ListItemToDoView.Selected(ListItemToDoView.ListIndex) = False


If ListItemIn.SelCount Then ComAddLink.Enabled = True

If ListItemSVC.ListIndex > -1 Then
    TextSVC = Type2Str(rs.Fields(ListItemSVC.List(ListItemSVC.ListIndex)).Type)
On Error Resume Next 'нет rs
'If Not (inRs Is Nothing) Then
    ChTypes inRs.Fields(ListItemIn.List(ListItemIn.ListIndex)).Type, rs.Fields(ListItemSVC.List(ListItemSVC.ListIndex)).Type
End If

'описание полей?
'If rS.Fields(ListItemSVC.List(ListItemSVC.ListIndex)).Properties(23) <> "109" Then
'ListItemSVC.ToolTipText = rS.Fields(ListItemSVC.List(ListItemSVC.ListIndex)).Properties(23)
'Else
'ListItemSVC.ToolTipText = vbNullString
'End If

err.Clear
End Sub

Private Function Type2Str(t As Integer) As String
Select Case t
Case 1: Type2Str = "Логика"
Case 2: Type2Str = "Байт"
Case 3: Type2Str = "Число, Integer"
Case 4: Type2Str = "Число, Long"
Case 5: Type2Str = "Деньги"
Case 6: Type2Str = "Число, Single"
Case 7: Type2Str = "Число, Double"
Case 8: Type2Str = "Дата"
Case 9: Type2Str = "Бинарное"
Case 10: Type2Str = "Текст"
Case 11: Type2Str = "Бинарное, Long"
Case 12: Type2Str = "MEMO"
'Case 13: Type2Str = "MEMO"
'Case 14: Type2Str = "MEMO"
Case 15: Type2Str = "GUID"
Case 16: Type2Str = "Число, BigInt"
Case 17: Type2Str = "Расширяемое, VarBinary"
Case 18: Type2Str = "Знак"
Case 19: Type2Str = "Числовое"
Case 20: Type2Str = "Десятичное"
'Case 21: Type2Str =
Case 22: Type2Str = "Время"
Case 23: Type2Str = "Время, TimeStamp"

Case Else: Type2Str = "?"
End Select

End Function
Private Function ChTypes(fIn As Integer, fOut As Integer) As Boolean
Select Case fOut
Case 1 'в логику
    Select Case fIn
        Case 1: ChTypes = True 'логику
        Case 4: ChTypes = False 'число
        Case 5: ChTypes = False 'деньги
        Case 8: ChTypes = False 'дата
        Case 10: ChTypes = False 'текст
        Case 11: ChTypes = False 'OLE
        Case 12: ChTypes = False 'Memo
        Case Else: ChTypes = False '?
    End Select
Case 4 'в число
    Select Case fIn
        Case 1: ChTypes = True 'логику
        Case 4: ChTypes = True 'число
        Case 5: ChTypes = False 'деньги
        Case 8: ChTypes = False 'дата
        Case 10: ChTypes = True 'текст
        Case 11: ChTypes = False 'OLE
        Case 12: ChTypes = True 'Memo
        Case Else: ChTypes = False '?
    End Select
Case 5 'в деньги
    Select Case fIn
        Case 1: ChTypes = False 'логику
        Case 4: ChTypes = False 'число
        Case 5: ChTypes = True 'деньги
        Case 8: ChTypes = False 'дата
        Case 10: ChTypes = True 'текст
        Case 11: ChTypes = False 'OLE
        Case 12: ChTypes = True 'Memo
        Case Else: ChTypes = False '?
    End Select
Case 8 'в дату
    Select Case fIn
        Case 1: ChTypes = False 'логику
        Case 4: ChTypes = False 'число
        Case 5: ChTypes = False 'деньги
        Case 8: ChTypes = True 'дата
        Case 10: ChTypes = True 'текст
        Case 11: ChTypes = False 'OLE
        Case 12: ChTypes = True 'Memo
        Case Else: ChTypes = False '?
    End Select
Case 10 'в текст
    Select Case fIn
        Case 1: ChTypes = True 'логику
        Case 4: ChTypes = True 'число
        Case 5: ChTypes = True 'деньги
        Case 8: ChTypes = True 'дата
        Case 10: ChTypes = True 'текст
        Case 11: ChTypes = False 'OLE
        Case 12: ChTypes = True 'Memo
        Case Else: ChTypes = False '?
    End Select
Case 11 'в OLE
    Select Case fIn
        Case 1: ChTypes = False 'логику
        Case 4: ChTypes = False 'число
        Case 5: ChTypes = False 'деньги
        Case 8: ChTypes = False 'дата
        Case 10: ChTypes = False 'текст
        Case 11: ChTypes = True 'OLE
        Case 12: ChTypes = False 'Memo
        Case Else: ChTypes = False '?
    End Select
Case 12 'в MEMO
    Select Case fIn
        Case 1: ChTypes = True 'логику
        Case 4: ChTypes = True 'число
        Case 5: ChTypes = True 'деньги
        Case 8: ChTypes = True 'дата
        Case 10: ChTypes = True 'текст
        Case 11: ChTypes = False 'OLE
        Case 12: ChTypes = True 'Memo
        Case Else: ChTypes = False '?
    End Select
Case Else 'в ?
        ChTypes = False

End Select

If ChTypes Then
LblType.ForeColor = 0 'black
Else
LblType.ForeColor = &HFF& 'red
'ComAddLink.Enabled = False
End If

End Function
Public Sub SetListboxScrollbar(lb As ListBox)
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

    For i = 0 To lb.ListCount - 1
        new_len = 10 + ScaleX(TextWidth(lb.List(i)), ScaleMode, vbPixels)
        If max_len < new_len Then max_len = new_len
    Next i

    SendMessage lb.hwnd, _
        LB_SETHORIZONTALEXTENT, _
        max_len, 0
End Sub

Private Sub ListItemSVC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
If ComAddLink.Enabled Then ComAddLink_Click

End If
End Sub

Private Sub ListItemToDoView_Click()
If ListItemToDoView.ListIndex = -1 Then Exit Sub

If ListItemToDoView.SelCount Then ComAddLink.Enabled = True


If ListItemSVC.ListIndex > -1 Then ListItemSVC.Selected(ListItemSVC.ListIndex) = False

If ListItemToDoView.ListIndex > -1 Then
    TextSVC = Type2Str(ListItemToDoView.ItemData(ListItemToDoView.ListIndex))
On Error Resume Next 'нет rs или ListItemIn.ListIndex
    ChTypes inRs.Fields(ListItemIn.List(ListItemIn.ListIndex)).Type, ListItemToDoView.ItemData(ListItemToDoView.ListIndex)
End If


End Sub
Private Sub GetSplitters(lstInd As Integer, pos As Integer, ByRef lS As String, ByRef rs As String)
Dim s1() As String
Dim s2() As String

s1 = Split(LstLSplit.List(lstInd), "$0$", -1, vbTextCompare)
s2 = Split(LstRSplit.List(lstInd), "$0$", -1, vbTextCompare)

If UBound(s1) >= pos Then lS = s1(pos)
If UBound(s2) >= pos Then rs = s2(pos)

End Sub

Private Sub ClearObj()
'ListItemIn.Clear
'ListItemSVC.Clear
ListItemToDoView.Clear
ListItemToDo.Clear
LstLSplit.Clear
LstRSplit.Clear

'mdb
'Set rS = Nothing
'Set DB = Nothing
'Set inRs = Nothing
'Set newDB = Nothing

'xml
Set XML_tmp = Nothing
Set XML_Child = Nothing
Set XML_Attr = Nothing

tPicLocate.Enabled = False

Reset

End Sub
Private Function StrToDbl(StringNumber As String) As Double
On Error Resume Next
StrToDbl = CDbl(StringNumber)
If err Then StrToDbl = Val(StringNumber)
err.Clear
End Function

Private Function SQLCompatible(ByRef s As String) As Boolean
'менять '# на ?
'If InStr(s, "?") > 0 Then s = Replace(s, "?", "[?]"): SQLCompatible = True '1
'If InStr(s, "'") > 0 Then s = Replace(s, "'", "?"): SQLCompatible = True
'If InStr(s, "#") > 0 Then s = Replace(s, "#", "?"): SQLCompatible = True

If InStr(s, "'") > 0 Then s = Replace(s, "'", "''") ': SQLCompatible = True


End Function

Private Function DupsFound_mdb(Optional row As Long) As Boolean
'pIn - содержимое поля, для стравнения на дубликаты из базы источника
'Dim i As Integer
Dim pIn As String
Dim sEqLike As String

'On Error Resume Next

DupsFound_mdb = False
'проверки на вшивость
If ChNoDup.Value = vbUnchecked Then Exit Function
If Not IsCSVflag Then
If DupInd < 0 Then Exit Function
Else
If CSVDupField < 1 Then Exit Function
End If
' взять данные
FromDupSub = True
If IsXMLflag Then pIn = GetCurrentInXML(DupInd) ' там получаем и SVCPoleName
If IsMDBflag Then pIn = GetCurrentInMDB(DupInd)
If IsAMCflag Then pIn = GetCurrentInAMC(DupInd, row)
If IsCSVflag Then pIn = GetCurrentInCSV(CSVDupField, row)
FromDupSub = False

Dim rsTmp As DAO.Recordset
Dim strSQL As String

'для чисел
Select Case rs(SVCPoleName).Type
Case 2, 3, 4 'число
    pIn = " = " & pIn

Case 12 'memo
    pIn = " = '" & pIn & "'"

Case 10 ' text
    sEqLike = " = '"
    If SQLCompatible(pIn) Then sEqLike = " Like '"
    pIn = sEqLike & Left$(pIn, 255) & "'"

Case Else 'нельзя по этим полям дупы искать
 ChNoDup.Value = vbUnchecked
 Exit Function

End Select


'если счетчик найденных совпадений pIn и SVCPoleName >0 то были совпадения
strSQL = "Select Count(" & SVCPoleName & ") From " & MyTableName & " Where " & SVCPoleName & pIn


Set rsTmp = DB.OpenRecordset(strSQL)

If Not (rsTmp.BOF And rsTmp.EOF) Then
    If rsTmp.RecordCount > 0 Then
        If rsTmp(0) <> 0 Then
        DupsFound_mdb = True
        'Debug.Print "дуп"
Else
Debug.Print pIn
        End If

    End If
End If


Set rsTmp = Nothing
err.Clear
End Function

Private Sub lvCSV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next

'если нет галки дупов
If ChNoDup.Value = vbUnchecked Then
    If cbCsvSvc.Text = cbCsvSvc.List(cbCsvSvc.ListIndex) Then 'не брать текст фразы-помощи
        'установить заголовок
        ColumnHeader.Text = cbCsvSvc
        'перейти к следующему в списке, если не пусто
        If cbCsvSvc <> vbNullString Then
        If cbCsvSvc.ListIndex = cbCsvSvc.ListCount - 1 Then
        cbCsvSvc.ListIndex = 0
        Else
        cbCsvSvc.ListIndex = cbCsvSvc.ListIndex + 1
        End If
        End If
    End If
Else
'установка поля дупов
    CSVDupField = ColumnHeader.Index
    If Len(lvCSV.ColumnHeaders(CSVDupField)) = 0 Then
        LDupFields.ForeColor = &HFF&
        LDupFields.Caption = "Кликните на заголовок нужного поля"
        CSVDupField = 0
    Else
        LDupFields = lvCSV.ColumnHeaders(CSVDupField)
        LDupFields.ForeColor = &H80000012
    End If
End If

End Sub

Private Sub lvCSV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Me.PopupMenu mnuPop

End If
End Sub

Private Sub mCheckAll_Click()
chCSV.Value = vbChecked
End Sub

Private Sub mCheckNone_Click()
chCSV.Value = vbUnchecked
End Sub
