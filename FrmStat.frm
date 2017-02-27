VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStat 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Statistics"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   Icon            =   "FrmStat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView LVStat 
      Height          =   3915
      Left            =   1905
      TabIndex        =   4
      Top             =   75
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6906
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Param"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   176
      EndProperty
   End
   Begin VB.ListBox ListStat 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   2760
      ItemData        =   "FrmStat.frx":000C
      Left            =   75
      List            =   "FrmStat.frx":003A
      TabIndex        =   3
      Top             =   1230
      Width           =   1755
   End
   Begin VB.Frame FrApplyTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Apply to"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   75
      TabIndex        =   5
      Top             =   60
      Width           =   1755
      Begin VB.OptionButton OpbApplyTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Selected"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   1455
      End
      Begin VB.OptionButton OpbApplyTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OpbApplyTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Checked"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'+ Локализация
'+ Число дисков по сумме носителей

' General:
'1+  Всего фильмов
'2+  Всего дисков
'       Sum (уникальная метка * кол-во носителей в ней)
'3+  Общий размер файлов

        'For i = 1 To FrmMain.ListView.ListItems.Count
        'k = k + Val(FrmMain.ListView.ListItems(i).SubItems(dbFileLenInd)) / oneMillion
        'Next i
        'Get_TotalFileLen = k

'4?  Обшая длительность
'5+  Всего обложек
'6+  Общий размер обложек
'7+  Всего кадров
'8+  Общий размер кадров
'9+  Фильмов у должников
'10+  дисков у должников

Private NStoreStat(13) As String 'от 0, доп. фразы
Private Const ListStatNum = 13 ' от 0, пунктов в списке выбора статистики
Private rsStat As DAO.Recordset
Private ApplyTo As Integer ' что в OpbApplyTo
Private Const OneMillion = 1000000
Private LVSelected As String ' ключи помеченных строк списка LV - ( 10,11,)
Private statBDName As String 'хранит bdname для исключения лишних операций, если не менялась база
Private SBNumFlag As Boolean 'true, если надо 1 поле сортировать как числа

Private SelNums As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub ListStat_Click()
Dim TotalDisks As Long
Dim TotalFileLen As Double
Dim i As Long    ', j As Long

Dim TotalDisksLabel As String
Dim TotalDisksSN As String

Dim strSQL As String, temp As String

On Error GoTo err

FrmMain.Timer2.Enabled = False
LVStat.ListItems.Clear
If FrmMain.ListView.ListItems.Count < 1 Then Exit Sub    'Unload Me
Screen.MousePointer = vbHourglass
LVStat.Sorted = False


If ApplyTo = 2 Then
    If (bdname <> statBDName) Or LVSelectChanged Then    'Or (SelNums <> UBound(SelRowsKey)) Then
        LVSelectChanged = False
        Dim LVArr() As String
        ReDim LVArr(UBound(SelRowsKey) - 1)
        For i = 0 To UBound(LVArr)
            LVArr(i) = Val(SelRowsKey(i + 1))
        Next i
        LVSelected = Join(LVArr, ",")
        Erase LVArr
        '''''
        SelNums = UBound(SelRowsKey)
        statBDName = bdname
    End If
End If

Select Case ListStat.ListIndex
Case 0    'General
    SBNumFlag = False

    LVStat.Sorted = False
    LVStat.ListItems.Add 1, , NStoreStat(0)
    LVStat.ListItems(1).SubItems(1) = Get_TotalTitles

    LVStat.ListItems.Add 2, , NStoreStat(1)
    TotalFileLen = Get_TotalFileLen
    LVStat.ListItems(2).SubItems(1) = TotalFileLen

    LVStat.ListItems.Add 3, , NStoreStat(2)    'с учетом меток
    TotalDisksLabel = Get_TotalDisks
    LVStat.ListItems(3).SubItems(1) = TotalDisksLabel

    'кол-во дисков по серийнику
    LVStat.ListItems.Add 4, , NStoreStat(10)
    TotalDisksSN = Get_TotalDisksSerial
    LVStat.ListItems(4).SubItems(1) = TotalDisksSN
    If TotalDisksSN = TotalDisksLabel Then
        LVStat.ListItems(4).SubItems(2) = ":)"
    Else
        LVStat.ListItems(4).SubItems(2) = ":("    '"<> " & TotalDisksLabel
    End If

    'кол-во пустых меток
    LVStat.ListItems.Add 5, , NStoreStat(11)
    LVStat.ListItems(5).SubItems(1) = Get_TotalEmpty("Label")

    'кол-во пустых серийников
    LVStat.ListItems.Add 6, , NStoreStat(12)
    LVStat.ListItems(6).SubItems(1) = Get_TotalEmpty("SNDisk")

    LVStat.ListItems.Add 7, , NStoreStat(3)
    TotalDisks = Get_TotalDisksSum
    LVStat.ListItems(7).SubItems(1) = TotalDisks

    '              'если общий размер / кол-во дисков мал - то наверно размер в базе в больших единицах , в мб а не в кб
    '              If TotalDisks > 0 Then
    '                If TotalFileLen / TotalDisks > 0.01 Then
    '                    LVStat.ListItems(2).SubItems(2) = "GB"
    '                Else
    '                    LVStat.ListItems(2).SubItems(2) = "TB"
    '                End If
    '              Else
    '                LVStat.ListItems(2).SubItems(2) = "GB"
    '              End If
    '              If TotalFileLen = 0 Then LVStat.ListItems(2).SubItems(2) = "GB"
    LVStat.ListItems(2).SubItems(2) = "GB"

    LVStat.ListItems.Add 8, , NStoreStat(4)
    LVStat.ListItems(8).SubItems(1) = Get_TotalCovers

    LVStat.ListItems.Add 9, , NStoreStat(5)
    LVStat.ListItems(9).SubItems(1) = Get_CoversLen
    LVStat.ListItems(9).SubItems(2) = "MB"

    LVStat.ListItems.Add 10, , NStoreStat(6)
    'долгоооо
    LVStat.ListItems(10).SubItems(1) = Get_TotalShots

    LVStat.ListItems.Add 11, , NStoreStat(7)
    'долгоооо
    LVStat.ListItems(11).SubItems(1) = Get_ShotsLen
    LVStat.ListItems(11).SubItems(2) = "MB"

    LVStat.ListItems.Add 12, , NStoreStat(8)
    LVStat.ListItems(12).SubItems(1) = Get_DebtMovies

    LVStat.ListItems.Add 13, , NStoreStat(9)
    LVStat.ListItems(13).SubItems(1) = Get_DebtDisks


Case 1    '                                                    Media
    SBNumFlag = False
    strSQL = "Select MediaType, Count(MediaType) From Storage Group By MediaType"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select MediaType, Count(MediaType) From Storage WHERE Checked = '1' Group By MediaType"
    Case 2    'select
        strSQL = "Select MediaType, Count(MediaType) From Storage WHERE Key IN (" & LVSelected & ") Group By MediaType"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount

Case 2    '                                                        Video
    SBNumFlag = False
    strSQL = "Select Video From Storage"
    Select Case ApplyTo
    Case 1    'check
        strSQL = strSQL & " WHERE Checked = '1'"
    Case 2    'select
        strSQL = strSQL & " WHERE Key IN (" & LVSelected & ")"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    TrimAndGroup " "

Case 3    '                                                      Audio
    SBNumFlag = False
    strSQL = "Select Audio From Storage"
    Select Case ApplyTo
    Case 1    'check
        strSQL = strSQL & " WHERE Checked = '1'"
    Case 2    'select
        strSQL = strSQL & " WHERE Key IN (" & LVSelected & ")"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    TrimAndGroup " ("

Case 4    '                                                      Framerate
    SBNumFlag = True
    strSQL = "Select FPS, Count(FPS) From Storage Group By FPS"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select FPS, Count(FPS) From Storage WHERE Checked = '1' Group By FPS"
    Case 2    'select
        strSQL = "Select FPS, Count(FPS) From Storage WHERE Key IN (" & LVSelected & ") Group By FPS"
    End Select

    Set rsStat = DB.OpenRecordset(strSQL)
    If Not (rsStat.BOF And rsStat.EOF) Then
        'If rsStat.RecordCount > 0 Then
        rsStat.MoveLast: rsStat.MoveFirst
        For i = 1 To rsStat.RecordCount
            If IsNull(rsStat(0)) Or IsNull(rsStat(1)) Then
            Else
                If (rsStat(0) = vbNullString) Or (rsStat(1) = vbNullString) Then
                Else
                    temp = Replace(rsStat(0), ",", ".")
                    LVStat.ListItems.Add(, , Replace2Regional(rsStat(0))).ListSubItems.Add 1, , rsStat(1)
                    'LVStat.ListItems.Add(, , rsStat(0)).ListSubItems.Add 1, , rsStat(1)

                    Select Case left$(temp, 2)
                    Case "25": LVStat.ListItems(LVStat.ListItems.Count).SubItems(2) = "PAL"
                    Case "29": LVStat.ListItems(LVStat.ListItems.Count).SubItems(2) = "NTSC"
                    Case "23": LVStat.ListItems(LVStat.ListItems.Count).SubItems(2) = "FILM"
                    End Select
                End If
            End If
            rsStat.MoveNext
        Next i
    End If
    LVStat.SortOrder = lvwDescending
    SortByNumber 1, LVStat

Case 5    '                                                Format
    SBNumFlag = False
    strSQL = "Select Resolution, Count(Resolution) From Storage Group By Resolution"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select Resolution, Count(Resolution) From Storage WHERE Checked = '1' Group By Resolution"
    Case 2    'select
        strSQL = "Select Resolution, Count(Resolution) From Storage WHERE Key IN (" & LVSelected & ") Group By Resolution"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount

Case 6    '                                                Страна
    SBNumFlag = False
    strSQL = "Select Country, Count(Country) From Storage Group By Country"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select Country, Count(Country) From Storage WHERE Checked = '1' Group By Country"
    Case 2    'select
        strSQL = "Select Country, Count(Country) From Storage WHERE Key IN (" & LVSelected & ") Group By Country"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount

Case 7    '                                               Жанр
    SBNumFlag = False
    strSQL = "Select Genre, Count(Genre) From Storage Group By Genre"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select Genre, Count(Genre) From Storage WHERE Checked = '1' Group By Genre"
    Case 2    'select
        strSQL = "Select Genre, Count(Genre) From Storage WHERE Key IN (" & LVSelected & ") Group By Genre"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount

Case 8    '                                                      Lang
    SBNumFlag = False
    strSQL = "Select Language, Count(Language) From Storage Group By Language"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select Language, Count(Language) From Storage WHERE Checked = '1' Group By Language"
    Case 2    'select
        strSQL = "Select Language, Count(Language) From Storage WHERE Key IN (" & LVSelected & ") Group By Language"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount

Case 9    '                                                      Subs
    SBNumFlag = False
    strSQL = "Select SubTitle, Count(SubTitle) From Storage Group By SubTitle"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select SubTitle, Count(SubTitle) From Storage WHERE Checked = '1' Group By SubTitle"
    Case 2    'select
        strSQL = "Select SubTitle, Count(SubTitle) From Storage WHERE Key IN (" & LVSelected & ") Group By SubTitle"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount

Case 10    '                                                      Rating
    SBNumFlag = True
    strSQL = "Select Rating, Count(Rating) From Storage Group By Rating"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select Rating, Count(Rating) From Storage WHERE Checked = '1' Group By Rating"
    Case 2    'select
        strSQL = "Select Rating, Count(Rating) From Storage WHERE Key IN (" & LVSelected & ") Group By Rating"
    End Select

    Set rsStat = DB.OpenRecordset(strSQL)
    If Not (rsStat.BOF And rsStat.EOF) Then
        'If rsStat.RecordCount > 0 Then
        rsStat.MoveLast: rsStat.MoveFirst
        For i = 1 To rsStat.RecordCount
            If IsNull(rsStat(0)) Or IsNull(rsStat(1)) Then
            Else
                If (rsStat(0) = vbNullString) Or (rsStat(1) = vbNullString) Then
                Else
                    LVStat.ListItems.Add(, , Replace2Regional(rsStat(0))).ListSubItems.Add 1, , rsStat(1)
                    'LVStat.ListItems.Add(, , rsStat(0)).ListSubItems.Add 1, , rsStat(1)
                End If
            End If
            rsStat.MoveNext
        Next i
    End If

    LVStat.SortOrder = lvwDescending
    SortByNumber 1, LVStat

Case 11    '                                                     Должник
    SBNumFlag = False
    strSQL = "Select Debtor, Count(Debtor) From Storage Group By Debtor"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select Debtor, Count(Debtor) From Storage WHERE Checked = '1' Group By Debtor"
    Case 2    'select
        strSQL = "Select Debtor, Count(Debtor) From Storage WHERE Key IN (" & LVSelected & ") Group By Debtor"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount

Case 12    '                                                     примеч
    SBNumFlag = False
    strSQL = "Select Other, Count(Other) From Storage Group By Other"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select Other, Count(Other) From Storage WHERE Checked = '1' Group By Other"
    Case 2    'select
        strSQL = "Select Other, Count(Other) From Storage WHERE Key IN (" & LVSelected & ") Group By Other"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount

Case 13    '                                                    год
    SBNumFlag = True
    strSQL = "Select Year, Count(Year) From Storage Group By Year"
    Select Case ApplyTo
    Case 1    'check
        strSQL = "Select Year, Count(Year) From Storage WHERE Checked = '1' Group By Year"
    Case 2    'select
        strSQL = "Select Year, Count(Year) From Storage WHERE Key IN (" & LVSelected & ") Group By Year"
    End Select
    Set rsStat = DB.OpenRecordset(strSQL)
    FillGroupCount
End Select

err:
Set rsStat = Nothing

Screen.MousePointer = vbNormal
If err Then MsgBox msgsvc(46), vbExclamation

End Sub
Private Sub Form_Load()
LVStat.ColumnHeaders(3).Width = LVStat.Width - LVStat.ColumnHeaders(1).Width - LVStat.ColumnHeaders(2).Width - 300
ApplyTo = 0
SetColorStat
GetLangStat
End Sub
Public Sub SetColorStat()


FrApplyTo.ForeColor = LVFontColor
FrApplyTo.BackColor = LVBackColor

OpbApplyTo(0).ForeColor = LVFontColor
OpbApplyTo(0).BackColor = LVBackColor
OpbApplyTo(1).ForeColor = LVFontColor
OpbApplyTo(1).BackColor = LVBackColor
OpbApplyTo(2).ForeColor = LVFontColor
OpbApplyTo(2).BackColor = LVBackColor

FrmStat.ForeColor = LVFontColor
FrmStat.BackColor = LVBackColor
ListStat.ForeColor = LVFontColor
ListStat.BackColor = LVBackColor
LVStat.ForeColor = LVFontColor
LVStat.BackColor = LVBackColor
End Sub
Private Sub Form_Resize()
''Background
If lngBrush <> 0 Then
GetClientRect hwnd, rctMain
FillRect hdc, rctMain, lngBrush
End If
End Sub

Private Sub GetLangStat()
Dim Contrl As Control
Dim i As Integer
'Dim temp As String

On Error Resume Next

If Dir(lngFileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = vbNullString Or Len(lngFileName) = 0 Then
Call myMsgBox("Не найден файл локализации! Исправьте параметр LastLang в global.ini" & vbCrLf & "Language file not found: " & vbCrLf & lngFileName, vbCritical, , Me.hwnd)
Exit Sub
End If

'ToDebug "Чтение файла локализации: " & lngFileName

For Each Contrl In FrmStat.Controls

'If TypeOf Contrl Is Label Then '                           Label
'Contrl.Caption = ReadLangStat(Contrl.name & ".Caption")
'End If

If TypeOf Contrl Is Frame Then '                           Frame
Contrl.Caption = ReadLangStat(Contrl.name & ".Caption", Contrl.Caption)
End If

If TypeOf Contrl Is OptionButton Then '                     OptionButton
For i = 0 To 2
OpbApplyTo(i).Caption = ReadLangStat(Contrl.name & i & ".Caption", OpbApplyTo(i).Caption)
Next
End If

'If TypeOf Contrl Is CheckBox Then '                         CheckBox
'Contrl.Caption = ReadLangStat(Contrl.name & ".Caption")
'End If

If TypeOf Contrl Is ListView Then '                        ListView
    If Contrl.name = "LVStat" Then
        For i = 1 To LVStat.ColumnHeaders.Count
            Contrl.ColumnHeaders(i).Text = ReadLangStat(Contrl.name & ".CH" & i, Contrl.ColumnHeaders(i).Text)
        Next i
    End If
End If

'                                                           Lst
For i = 0 To ListStatNum
 ListStat.List(i) = ReadLangStat("ListStat" & i, ListStat.List(i))
Next i

Next 'Contrl

'                                                       NamesStoreStat()
For i = 0 To UBound(NStoreStat)
NStoreStat(i) = ReadLangStat("NStoreStat" & i)
Next i

Me.Caption = "SurVideoCatalog - " & ReadLang("VerticalMenu.MenuItemCaption7")
Me.Icon = FrmMain.Icon
End Sub

Private Function Get_TotalTitles() As Integer
Dim strSQL As String
Select Case ApplyTo
Case 0 'all
'Get_TotalTitles = rs.RecordCount 'FrmMain.ListView.ListItems.Count
strSQL = "SELECT Count(key) FROM Storage"      'всего фильмов
Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.RecordCount = 1 Then
 rsStat.MoveFirst
 If IsNull(rsStat(0)) Then
  Get_TotalTitles = 0
 Else
  Get_TotalTitles = Val(rsStat(0))
 End If
End If

Case 1 'check
    Get_TotalTitles = CheckCount
Case 2 'select
    Get_TotalTitles = SelCount
End Select

Set rsStat = Nothing
End Function

Private Function Get_TotalDisks() As Long
Dim strSQL As String ', temp As String

Select Case ApplyTo
Case 0 'all
 strSQL = "SELECT MAX(VAL(IIf(CDN Is Null, 0, CDN))) AS TOTAL FROM Storage Group By(Label)"      'всего дисков
 strSQL = "SELECT SUM(TOTAL) FROM (" & strSQL & ")"
Case 1 'check
 strSQL = "SELECT MAX(VAL(IIf(CDN Is Null, 0, CDN))) AS TOTAL FROM Storage WHERE Checked = '1' Group By(Label)"      'всего дисков
 strSQL = "SELECT SUM(TOTAL) FROM (" & strSQL & ")"
Case 2 'select
 strSQL = "SELECT MAX(VAL(IIf(CDN Is Null, 0, CDN))) AS TOTAL FROM Storage WHERE Key IN (" & LVSelected & ") Group By(Label)"      'всего дисков
 strSQL = "SELECT SUM(TOTAL) FROM (" & strSQL & ")"
End Select

'Debug.Print strSQL
Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.RecordCount = 1 Then
 rsStat.MoveFirst
 If IsNull(rsStat(0)) Then
  Get_TotalDisks = 0
 Else
  Get_TotalDisks = Val(rsStat(0))
 End If
End If

Set rsStat = Nothing
End Function
Private Function Get_DebtDisks() As Integer
Dim strSQL As String ', temp As String

Select Case ApplyTo
Case 0 'all
 strSQL = "SELECT MAX(VAL(IIf(CDN Is Null, 0, CDN))) AS TOTAL FROM Storage WHERE Debtor <> """" Group By(Label)"
 strSQL = "SELECT SUM(TOTAL) FROM (" & strSQL & ")"
Case 1 'check
 strSQL = "SELECT MAX(VAL(IIf(CDN Is Null, 0, CDN))) AS TOTAL FROM Storage WHERE Debtor <> """" And Checked = '1' Group By(Label)"
 strSQL = "SELECT SUM(TOTAL) FROM (" & strSQL & ")"
Case 2 'select
 strSQL = "SELECT MAX(VAL(IIf(CDN Is Null, 0, CDN))) AS TOTAL FROM Storage WHERE Debtor <> """" And Key IN (" & LVSelected & ") Group By(Label)"
 strSQL = "SELECT SUM(TOTAL) FROM (" & strSQL & ")"
End Select

'Debug.Print strSQL
Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.RecordCount = 1 Then
 rsStat.MoveFirst
 If IsNull(rsStat(0)) Then
  Get_DebtDisks = 0
 Else
  Get_DebtDisks = Val(rsStat(0))
 End If
End If

Set rsStat = Nothing
End Function
Private Function Get_TotalFileLen() As Double
Dim strSQL As String
'в гигабайтах

strSQL = "SELECT SUM(FileLen) FROM Storage"

Select Case ApplyTo
Case 1 'check
 strSQL = strSQL & " WHERE Checked = '1'"
Case 2 'select
 strSQL = strSQL & " WHERE Key IN (" & LVSelected & ")"
End Select

Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.Fields.Count = 1 Then
 rsStat.MoveFirst
 If IsNull(rsStat(0)) Then
  Get_TotalFileLen = 0
 Else
  Get_TotalFileLen = rsStat(0) / OneMillion
 End If
End If

Set rsStat = Nothing
End Function
Private Function Get_TotalDisksSum() As Double
Dim strSQL As String

strSQL = "SELECT SUM(VAL(IIf(CDN Is Null, 0, CDN))) FROM Storage"

Select Case ApplyTo
Case 1 'check
 strSQL = strSQL & " WHERE Checked = '1'"
Case 2 'select
 strSQL = strSQL & " WHERE Key IN (" & LVSelected & ")"
End Select

Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.Fields.Count = 1 Then
 rsStat.MoveFirst
 If IsNull(rsStat(0)) Then
  Get_TotalDisksSum = 0
 Else
  Get_TotalDisksSum = rsStat(0)
 End If
End If

Set rsStat = Nothing
End Function
Private Function Get_TotalCovers() As Integer
Dim strSQL As String

strSQL = "SELECT Count(FrontFace) FROM Storage"  'Where FrontFace <> """" "
'strSQL = "SELECT Count(FrontFace<>'') FROM Storage"

Select Case ApplyTo
Case 1 'check
 strSQL = strSQL & " Where Checked = '1'"
Case 2 'select
 strSQL = strSQL & " Where Key IN (" & LVSelected & ")"
End Select

'Debug.Print strSQL
Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.Fields.Count = 1 Then
 If IsNull(rsStat(0)) Then
  Get_TotalCovers = 0
 Else
  Get_TotalCovers = rsStat(0)
 End If
End If

Set rsStat = Nothing
End Function
Private Function Get_DebtMovies() As Integer
Dim strSQL As String

strSQL = "SELECT Count(Debtor) FROM Storage WHERE Debtor <> """""

Select Case ApplyTo
Case 1 'check
 strSQL = strSQL & " And Checked = '1'"
Case 2 'select
 strSQL = strSQL & " And Key IN (" & LVSelected & ")"
End Select

'Debug.Print strSQL
Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.Fields.Count = 1 Then
 If IsNull(rsStat(0)) Then
  Get_DebtMovies = 0
 Else
  Get_DebtMovies = rsStat(0)
 End If
End If

Set rsStat = Nothing
End Function

Private Function Get_TotalShots() As Integer
Dim strSQL As String

'strSQL = "SELECT Count(SnapShot1<>'') + Count(SnapShot2<>'') + Count(SnapShot3<>'') FROM Storage"
strSQL = "SELECT Count(SnapShot1) + Count(SnapShot2) + Count(SnapShot3) FROM Storage"

Select Case ApplyTo
Case 1 'check
 strSQL = strSQL & " WHERE Checked = '1'"
Case 2 'select
 strSQL = strSQL & " WHERE Key IN (" & LVSelected & ")"
End Select

'Debug.Print strSQL
Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.Fields.Count = 1 Then
 If IsNull(rsStat(0)) Then
  Get_TotalShots = 0
 Else
  Get_TotalShots = rsStat(0)
 End If
End If

Set rsStat = Nothing
End Function

Private Function Get_ShotsLen() As Double
Dim strSQL As String

strSQL = "SELECT SUM(IIf(Len(SnapShot1) Is Null,0,Len(SnapShot1)) + IIf(Len(SnapShot2) Is Null,0,Len(SnapShot2)) + IIf(Len(SnapShot3) Is Null,0,Len(SnapShot3))) FROM Storage"

Select Case ApplyTo
Case 1 'check
 strSQL = strSQL & " Where Checked = '1'"
Case 2 'select
 strSQL = strSQL & " Where Key IN (" & LVSelected & ")"
End Select

'Debug.Print strSQL
Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.Fields.Count = 1 Then
 If IsNull(rsStat(0)) Then
  Get_ShotsLen = 0
 Else
  Get_ShotsLen = rsStat(0) / OneMillion * 2
  End If
End If

Set rsStat = Nothing
End Function
Private Function Get_CoversLen() As Single
Dim strSQL As String

strSQL = "SELECT SUM(Len(FrontFace)) FROM Storage"

Select Case ApplyTo
Case 1 'check
 strSQL = strSQL & " Where Checked = '1'"
Case 2 'select
 strSQL = strSQL & " Where Key IN (" & LVSelected & ")"
End Select

'Debug.Print strSQL
Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.Fields.Count = 1 Then
 If IsNull(rsStat(0)) Then
  Get_CoversLen = 0
 Else
  Get_CoversLen = rsStat(0) / OneMillion * 2
  End If
End If

Set rsStat = Nothing
End Function
Private Function Get_TotalEmpty(F As String) As Long
'считаем кол-во незаполненных полей f
Dim strSQL As String

Select Case ApplyTo
Case 0 'all
 strSQL = "SELECT Count(" & F & ") FROM Storage Where ((" & F & " ='') or (" & F & " Is Null))"
Case 1 'check
 strSQL = "SELECT Count(" & F & ") FROM Storage Where (((" & F & " ='') or (" & F & " Is Null)) And (Checked = '1'))"
Case 2 'select
 strSQL = "SELECT Count(" & F & ") FROM Storage Where (((" & F & " ='') or (" & F & " Is Null)) And (Key IN (" & LVSelected & ")))"

End Select

Set rsStat = DB.OpenRecordset(strSQL)
If rsStat.Fields.Count = 1 Then
 If IsNull(rsStat(0)) Then
  Get_TotalEmpty = 0
 Else
  Get_TotalEmpty = rsStat(0)
 End If
End If

Set rsStat = Nothing
End Function

Private Function Get_TotalDisksSerial() As Long
'Добыть, Потрошить по , и вернуть кол-во
Dim strSQL As String

Select Case ApplyTo
Case 0 'all
 strSQL = "SELECT snDisk FROM Storage"
Case 1 'check
 strSQL = "SELECT snDisk FROM Storage WHERE Checked = '1'"
Case 2 'select
 strSQL = "SELECT snDisk FROM Storage WHERE Key IN (" & LVSelected & ")"
End Select

'Debug.Print strSQL
Set rsStat = DB.OpenRecordset(strSQL)
If Not (rsStat.BOF And rsStat.EOF) Then
 rsStat.MoveLast: rsStat.MoveFirst
 If rsStat.RecordCount = 0 Then
  Get_TotalDisksSerial = 0
 Else
  'Потрошить по , и вернуть кол-во
  Get_TotalDisksSerial = GetSeparNums
  'Get_TotalDisksSerial = Val(rsStat(0))
 End If
Else
  Get_TotalDisksSerial = 0
End If

Set rsStat = Nothing
End Function

Private Function GetSeparNums() As Long
'потрошить данные рекордсета в массив rsArr, вернуть кол-во
Dim j As Long
Dim R() As String
Dim ArrFlag As Boolean
Dim rsArr() As String

ReDim rsArr(0)
Do While Not rsStat.EOF
    If IsNull(rsStat(0)) Then    'пустышка раз
    ElseIf Len(rsStat(0)) = 0 Then    'пустышка два
    Else
        If Tokenize04(rsStat(0), R(), ",", False) > -1 Then              ' False! Пустышек быть не должно.
            For j = 0 To UBound(R)
                If ArrFlag Then ReDim Preserve rsArr(UBound(rsArr) + 1)            'пропустить первую
                rsArr(UBound(rsArr)) = R(j)
                ArrFlag = True
            Next j
        End If
    End If
    rsStat.MoveNext
Loop

TriQuickSortString rsArr    'sorts your string array
remdups rsArr    'removes all duplicates

If (UBound(rsArr) = 0) And (Not ArrFlag) Then
GetSeparNums = 0
Else
GetSeparNums = UBound(rsArr) + 1
End If
End Function


Private Sub LVStat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

LVStat.SortOrder = (LVStat.SortOrder + 1) Mod 2

Select Case ColumnHeader.Index
Case 1
 If SBNumFlag Then
  SortByNumber 0, LVStat
 Else
  LVStat.SortKey = 0: LVStat.Sorted = True
 End If
Case 2
 SortByNumber 1, LVStat
Case Else
 LVStat.SortKey = 2: LVStat.Sorted = True
End Select
End Sub

Private Sub OpbApplyTo_Click(Index As Integer)
ApplyTo = Index

If ListStat.ListIndex > -1 Then ListStat_Click
End Sub

Private Sub TrimAndGroup(ch As String)
'ch - использовать подстроку до начала этой строки
Dim i As Integer, j As Integer, M As Integer, p As Integer
Dim a() As String
Dim s() As String
Dim n() As Integer
Dim temp As String

If Not (rsStat.BOF And rsStat.EOF) Then
 rsStat.MoveLast: rsStat.MoveFirst
 ReDim a(rsStat.RecordCount)
 For i = 1 To rsStat.RecordCount
  If IsNull(rsStat(0)) Then
  Else
   If rsStat(0) = vbNullString Then
   Else
    temp = UCase$(LTrim$(rsStat(0)))
    j = InStr(temp, ch)
    If j > 0 Then
     a(i) = left$(temp, j - 1)
    Else
     a(i) = temp
    End If
   End If
  End If
  rsStat.MoveNext
 Next i
End If

ReDim s(0): ReDim n(0): p = 0
For i = 1 To rsStat.RecordCount
 If a(i) <> vbNullString Then
  M = 1 ' сумма группированных
  p = p + 1 ' счетчик уникальных строк
  ReDim Preserve s(p) ' таблица уникальных строк
  ReDim Preserve n(p) ' таблица сумм
  s(p) = a(i)
  n(p) = M
  a(i) = vbNullString
  
  For j = 1 To rsStat.RecordCount
   If a(j) <> vbNullString Then
    If s(p) = a(j) Then
     M = M + 1
     n(p) = M
     a(j) = vbNullString
    End If
   End If
  Next j
 End If
Next i

For i = 1 To UBound(s)
LVStat.ListItems.Add(i, , s(i)).ListSubItems.Add 1, , n(i)
Next i
LVStat.SortOrder = lvwDescending
SortByNumber 1, LVStat

End Sub

Private Sub FillGroupCount()
Dim i As Long

If Not (rsStat.BOF And rsStat.EOF) Then
'If rsStat.RecordCount > 0 Then
 rsStat.MoveLast: rsStat.MoveFirst
 For i = 1 To rsStat.RecordCount
  If IsNull(rsStat(0)) Or IsNull(rsStat(1)) Then
  Else
   If (rsStat(0) = vbNullString) Or (rsStat(1) = vbNullString) Then
   Else
    LVStat.ListItems.Add(, , rsStat(0)).ListSubItems.Add 1, , rsStat(1)
   End If
  End If
  rsStat.MoveNext
 Next i
End If

LVStat.SortOrder = lvwDescending
SortByNumber 1, LVStat
End Sub
