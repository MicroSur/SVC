VERSION 5.00
Begin VB.Form FrmFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SurVideoCatalog"
   ClientHeight    =   5670
   ClientLeft      =   4815
   ClientTop       =   3420
   ClientWidth     =   5895
   Icon            =   "FrmFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chNotAll 
      Height          =   195
      Left            =   5580
      TabIndex        =   54
      Top             =   3720
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   9
      Left            =   5580
      TabIndex        =   53
      Top             =   3360
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   8
      Left            =   5580
      TabIndex        =   52
      Top             =   3000
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   7
      Left            =   5580
      TabIndex        =   51
      Top             =   2640
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   6
      Left            =   5580
      TabIndex        =   50
      Top             =   2280
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   5
      Left            =   5580
      TabIndex        =   49
      Top             =   1920
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   4
      Left            =   5580
      TabIndex        =   48
      Top             =   1560
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   3
      Left            =   5580
      TabIndex        =   47
      Top             =   1200
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   2
      Left            =   5580
      TabIndex        =   46
      Top             =   840
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   1
      Left            =   5580
      TabIndex        =   45
      Top             =   480
      Width           =   195
   End
   Begin VB.CheckBox chNot 
      Height          =   195
      Index           =   0
      Left            =   5580
      TabIndex        =   44
      Top             =   120
      Width           =   195
   End
   Begin VB.CheckBox chAOAll 
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   3720
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   33
      Top             =   3360
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   32
      Top             =   3000
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   31
      Top             =   2640
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   30
      Top             =   2280
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   28
      Top             =   1560
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   480
      Width           =   195
   End
   Begin VB.CheckBox chAO 
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   195
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   0
      ItemData        =   "FrmFilter.frx":000C
      Left            =   1680
      List            =   "FrmFilter.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   60
      Width           =   3795
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   9
      ItemData        =   "FrmFilter.frx":0010
      Left            =   1680
      List            =   "FrmFilter.frx":0012
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   3300
      Width           =   3795
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   8
      ItemData        =   "FrmFilter.frx":0014
      Left            =   1680
      List            =   "FrmFilter.frx":0016
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   2940
      Width           =   3795
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   7
      ItemData        =   "FrmFilter.frx":0018
      Left            =   1680
      List            =   "FrmFilter.frx":001A
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   2580
      Width           =   3795
   End
   Begin VB.CheckBox ChFiltWhole 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   5280
      MaskColor       =   &H8000000F&
      TabIndex        =   37
      Top             =   4080
      Width           =   195
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   1
      ItemData        =   "FrmFilter.frx":001C
      Left            =   1680
      List            =   "FrmFilter.frx":001E
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   420
      Width           =   3795
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      CausesValidation=   0   'False
      Height          =   315
      Index           =   6
      ItemData        =   "FrmFilter.frx":0020
      Left            =   1680
      List            =   "FrmFilter.frx":0022
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   2220
      Width           =   3795
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   4
      ItemData        =   "FrmFilter.frx":0024
      Left            =   1680
      List            =   "FrmFilter.frx":0026
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   1500
      Width           =   3795
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   5
      ItemData        =   "FrmFilter.frx":0028
      Left            =   1680
      List            =   "FrmFilter.frx":002A
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   1860
      Width           =   3795
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   3
      ItemData        =   "FrmFilter.frx":002C
      Left            =   1680
      List            =   "FrmFilter.frx":002E
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   1140
      Width           =   3795
   End
   Begin VB.ComboBox cbs 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   2
      ItemData        =   "FrmFilter.frx":0030
      Left            =   1680
      List            =   "FrmFilter.frx":0032
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   780
      Width           =   3795
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   9
      ItemData        =   "FrmFilter.frx":0034
      Left            =   420
      List            =   "FrmFilter.frx":0036
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3300
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   8
      ItemData        =   "FrmFilter.frx":0038
      Left            =   420
      List            =   "FrmFilter.frx":003A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2940
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   7
      ItemData        =   "FrmFilter.frx":003C
      Left            =   420
      List            =   "FrmFilter.frx":003E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2580
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   6
      ItemData        =   "FrmFilter.frx":0040
      Left            =   420
      List            =   "FrmFilter.frx":0042
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2220
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   5
      ItemData        =   "FrmFilter.frx":0044
      Left            =   420
      List            =   "FrmFilter.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1860
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   4
      ItemData        =   "FrmFilter.frx":0048
      Left            =   420
      List            =   "FrmFilter.frx":004A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1500
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   3
      ItemData        =   "FrmFilter.frx":004C
      Left            =   420
      List            =   "FrmFilter.frx":004E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1140
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   2
      ItemData        =   "FrmFilter.frx":0050
      Left            =   420
      List            =   "FrmFilter.frx":0052
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   780
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   1
      ItemData        =   "FrmFilter.frx":0054
      Left            =   420
      List            =   "FrmFilter.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1755
   End
   Begin VB.ComboBox cbl 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   0
      ItemData        =   "FrmFilter.frx":0058
      Left            =   420
      List            =   "FrmFilter.frx":005A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   1755
   End
   Begin VB.CheckBox chFiltStart 
      Height          =   195
      Left            =   5280
      TabIndex        =   38
      Top             =   4320
      Width           =   195
   End
   Begin VB.CheckBox ChFiltCover 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   360
      MaskColor       =   &H8000000F&
      TabIndex        =   36
      Top             =   4320
      Width           =   195
   End
   Begin VB.CheckBox ChFiltSShots 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   360
      MaskColor       =   &H8000000F&
      TabIndex        =   35
      Top             =   4080
      Width           =   195
   End
   Begin SurVideoCatalog.XpB cExcludeFilter 
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   4680
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   661
      Caption         =   "Exclude"
      ButtonStyle     =   3
      Picture         =   "FrmFilter.frx":005C
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB CommClearFilter 
      Height          =   375
      Left            =   300
      TabIndex        =   21
      Top             =   5160
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   661
      Caption         =   "Clear"
      ButtonStyle     =   3
      Picture         =   "FrmFilter.frx":05F6
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB CommShowAll 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   5160
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   661
      Caption         =   "Undo"
      ButtonStyle     =   3
      Picture         =   "FrmFilter.frx":1008
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB CommApplyFilter 
      Default         =   -1  'True
      Height          =   375
      Left            =   300
      TabIndex        =   20
      Top             =   4680
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   661
      Caption         =   "Include"
      ButtonStyle     =   3
      Picture         =   "FrmFilter.frx":1A1A
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin VB.Label lFiltNot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "not"
      Height          =   195
      Left            =   4080
      TabIndex        =   55
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblFiltAndOr 
      BackStyle       =   0  'Transparent
      Caption         =   "and/or"
      Height          =   255
      Left            =   420
      TabIndex        =   43
      Top             =   3720
      Width           =   3435
   End
   Begin VB.Label lChFiltWhole 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Whole"
      Height          =   195
      Left            =   2880
      TabIndex        =   39
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lChFiltSShots 
      BackStyle       =   0  'Transparent
      Caption         =   "With screer shots"
      Height          =   195
      Left            =   660
      TabIndex        =   40
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lchFiltStart 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      Height          =   195
      Left            =   2820
      TabIndex        =   42
      Top             =   4320
      Width           =   2355
   End
   Begin VB.Label lChFiltCover 
      BackStyle       =   0  'Transparent
      Caption         =   "With cover"
      Height          =   195
      Left            =   660
      TabIndex        =   41
      Top             =   4320
      Width           =   1695
   End
End
Attribute VB_Name = "FrmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Private Const CB_SHOWDROPDOWN = &H14F
'
'
'Private Sub cbl_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'SendMessage cbl(Index).hWnd, CB_SHOWDROPDOWN, 1, ByVal 0
'End Sub



Private Sub cbl_Click(Index As Integer)
Dim tmp As String

tmp = IPN(cbl(Index).ListIndex, vbNullString) 'vbNullString тут нужно чистые названия полей, без val

Select Case LCase$(tmp)
Case "annotation"
Case "moviename", "label", "year", "time", "resolution", "fps", "cdn", "mediatype", "rating", "debtor", "other", "coverpath", "movieurl"
    FillCombosNoParse tmp, cbs(Index)
Case Else
    FillCombosParse tmp, cbs(Index)
End Select
End Sub

Private Sub cbs_Change(Index As Integer)
If Len(cbs(Index)) = 0 Then
cbs(Index).BackColor = &H8000000F
Else
cbs(Index).BackColor = &H80000005
End If

End Sub

Private Sub cbs_Click(Index As Integer)
If Len(cbs(Index)) = 0 Then
cbs(Index).BackColor = &H8000000F
Else
cbs(Index).BackColor = &H80000005
End If
End Sub

Private Sub cExcludeFilter_Click()
''InitFlag = True
Dim IsTyped As Boolean 'что-то вписано
Dim IsChecked As Boolean ' что-то помечено
Dim Contrl As Control

For Each Contrl In FrmFilter.Controls
 If TypeOf Contrl Is ComboBox Then
    If Contrl.style <> 2 Then 'не листы
        If LenB(Contrl.Text) <> 0 Then IsTyped = True
    End If
 End If
Next
If ChFiltSShots.Value = vbChecked Then IsChecked = True
If ChFiltCover.Value = vbChecked Then IsChecked = True


If IsTyped Or IsChecked Then
    FilterAddTypedItems
    'Call FrmMain.DelFilterItems
    'Call FrmMain.ExclFilterItemsSQL(IsTyped)
    Call FilterItemsSQL(IsTyped, "NOT")
    'Call FrmMain.FillFilter
    'FrmMain.FrameView.Caption = FrameViewCaption & " " & FrmMain.ListView.ListItems.Count & " )" 'нестыкуется с группами
End If
End Sub

Private Sub chAOAll_Click()
Dim i As Integer
For i = 0 To cbTotal - 1
chAO(i).Value = chAOAll.Value
Next i
End Sub

Private Sub chFiltStart_Click()
If chFiltStart.Value = vbChecked Then
    'ChFiltWhole.Enabled = False
    ChFiltWhole.Value = vbUnchecked
Else
    ChFiltWhole.Enabled = True
End If
End Sub

Private Sub ChFiltWhole_Click()
If ChFiltWhole.Value = vbChecked Then
    'chFiltStart.Enabled = False
    chFiltStart.Value = vbUnchecked
Else
    chFiltStart.Enabled = True
End If
End Sub

Private Sub chNotAll_Click()
Dim i As Integer
For i = 0 To cbTotal - 1
chNot(i).Value = chNotAll.Value
Next i
End Sub

Private Sub CommApplyFilter_Click()
''InitFlag = True
Dim IsTyped As Boolean
Dim IsChecked As Boolean
Dim Contrl As Control

For Each Contrl In FrmFilter.Controls
 If TypeOf Contrl Is ComboBox Then
    If Contrl.style <> 2 Then 'не листы
        If LenB(Contrl.Text) <> 0 Then IsTyped = True
    End If
 End If
Next
If ChFiltSShots.Value = vbChecked Then IsChecked = True
If ChFiltCover.Value = vbChecked Then IsChecked = True

If IsChecked Or IsTyped Then
    FilterAddTypedItems 'вбить в комбики введенное пользователем
    'Call FrmMain.DelFilterItems
    Call FilterItemsSQL(IsTyped)
    'Call FrmMain.FillFilter
    'FrmMain.FrameView.Caption = FrameViewCaption & " " & FrmMain.ListView.ListItems.Count & " )"
End If
End Sub

Private Sub FilterAddTypedItems()
'вбить в комбики введенное пользователем
Dim i As Integer
On Error Resume Next

For i = 0 To cbTotal - 1
If Len(cbs(i).Text) <> 0 Then
   If SearchCBO(cbs(i), cbs(i).Text, False) < 0 Then
    SendMessage cbs(i).hwnd, CB_ADDSTRING, 0, ByVal cbs(i).Text
   End If
End If
Next i

End Sub
Private Sub CommClearFilter_Click()
Dim Contrl As Control

FilterAddTypedItems 'вбить в комбики введенное пользователем

For Each Contrl In FrmFilter.Controls
 If TypeOf Contrl Is ComboBox Then
    If Contrl.style <> 2 Then Contrl.Text = vbNullString
 End If
Next
ChFiltSShots.Value = vbUnchecked
ChFiltCover.Value = vbUnchecked
ChFiltWhole.Value = vbUnchecked
chFiltStart.Value = vbUnchecked
End Sub

Private Sub CommShowAll_Click()
'отмена группировки, аналог ComFilter_ShiftClick
Dim strSQL As String
On Error Resume Next ' надо если после фильтров нет помеченных полей

If FiltPersonFlag Or FiltValidationFlag Or FilteredFlag Then
'заходим
Else
    Unload FrmFilter
    Exit Sub
End If

LastInd = FrmMain.ListView.SelectedItem.SubItems(lvIndexPole)

If GroupedFlag And Len(LastSQLGroupString) <> 0 Then
'применить запрос от группировки
strSQL = "Select * From Storage Where " & LastSQLGroupString
Else
'вернуть все
strSQL = "Select * From Storage"
End If

Set rs = DB.OpenRecordset(strSQL)
'rs.MoveFirst: rs.MoveLast
'ReDim lvItemLoaded(rs.RecordCount)

FiltPersonFlag = False 'снять флаг фильтрации по актеру
FiltValidationFlag = False 'снять флаг проверки метка.серийник
FilteredFlag = False 'флаг неполного показа (фильтрация) 1

'If GroupedFlag Then
'    FrmMain.FillListView 'FrmMain.FillTVGroup '2
'    'FrmMain.TVCLICK
'Else
FrmMain.FillListView
'End If

FillFilter

FrmMain.ComFilter.BackColor = &HFFFFFF
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 112 'F1
    If FrmMain.ChBTT.Value = 0 Then FrmMain.ChBTT.Value = 1 Else FrmMain.ChBTT.Value = 0
    Me.SetFocus
End Select
End Sub

Private Sub Form_Load()
Dim Contrl As Control
Dim i As Integer, j As Integer
Dim X As Long, Y As Long, w As Long
Dim tmp As String
Dim sItemText As String

'prepare the receiving buffer
sItemText = Space$(512)
'sItemText = AllocString_ADV(512)

'Dim FieldNameLocal(25) As String

'заполнить названиями комболисты
For j = 1 To lvIndexPole '24
    tmp = FrmMain.ListView.ColumnHeaders(j).Text
    If right$(tmp, 2) = " >" Or right$(tmp, 2) = " <" Then tmp = left$(tmp, Len(tmp) - 2)
    cbl(0).AddItem tmp
Next j
cbl(0).AddItem frmEditor.LFilm(9).Caption

For j = 0 To lvIndexPole '24
    ' get the item text
    SendMessage cbl(0).hwnd, CB_GETLBTEXT, j, ByVal sItemText
    ' get the item data
    'itmData = SendMessage(Source.hWnd, CB_GETITEMDATA, Index, ByVal 0&)
    For i = 1 To cbTotal - 1
        ' add the item text to the target list
        SendMessage cbl(i).hwnd, CB_ADDSTRING, 0&, ByVal sItemText
        ' add the item data to the target list
        'SendMessage Target.hWnd, CB_SETITEMDATA, Index, ByVal itmData
    Next i
Next j

For i = 0 To cbTotal - 1
    'высота комбиков
    X = ScaleX(cbl(i).left, vbTwips, vbPixels)
    Y = ScaleY(cbl(i).top, vbTwips, vbPixels)
    w = ScaleY(cbl(i).Width, vbTwips, vbPixels)
    SetWindowPos cbl(i).hwnd, 0, X, Y, w, 500, SWP_NOZORDER
Next i

'вывести нужные названия
SendMessage cbl(0).hwnd, CB_SETCURSEL, 0, 0    'MovieName по Function IPN
SendMessage cbl(1).hwnd, CB_SETCURSEL, 2, 0    'Genre
SendMessage cbl(2).hwnd, CB_SETCURSEL, 4, 0    'Country
SendMessage cbl(3).hwnd, CB_SETCURSEL, 3, 0    'Year
SendMessage cbl(4).hwnd, CB_SETCURSEL, 5, 0    'Director
SendMessage cbl(5).hwnd, CB_SETCURSEL, 6, 0    'Acter
SendMessage cbl(6).hwnd, CB_SETCURSEL, 17, 0    'Rating
SendMessage cbl(7).hwnd, CB_SETCURSEL, 18, 0    'FileName
SendMessage cbl(8).hwnd, CB_SETCURSEL, 1, 0    'Label
SendMessage cbl(9).hwnd, CB_SETCURSEL, 19, 0    'Debtor

For Each Contrl In FrmFilter.Controls
    If TypeOf Contrl Is ComboBox Then Contrl.Font.Charset = 204

    If TypeOf Contrl Is Label Then    '                           Label
        Contrl.Caption = ReadLangFilt(Contrl.name & ".Caption", Contrl.Caption)
    End If
    If TypeOf Contrl Is XpB Then
        Contrl.Caption = ReadLangFilt(Contrl.name & ".Caption", Contrl.Caption)
        '        Contrl.ToolTipText = ReadLangFilt(Contrl.name & ".ToolTip")
        Contrl.pInitialize
    End If
    If TypeOf Contrl Is CheckBox Then    '                         CheckBox
        Contrl.Caption = ReadLangFilt(Contrl.name & ".Caption", Contrl.Caption)
    End If

Next

Me.Caption = "SurVideoCatalog - " & FrmMain.ComFilter.Caption
Me.Icon = FrmMain.Icon

frmFilterFlag = True
End Sub

Private Sub Form_Resize()
'Background
If lngBrush <> 0 Then
    GetClientRect hwnd, rctMain
    FillRect hdc, rctMain, lngBrush
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmFilterFlag = False

If Not ExitSVC Then
    Cancel = True
    'прятать
    Me.Hide
    FrmMain.Timer2.Enabled = True
Else
 frmFilterFlag = False
End If
End Sub
Public Sub FillCombosNoParse(fldName As String, cmbox As ComboBox)
'заполняет комбики значениями из полей базы
'fldName - поле базы
Dim tmp As String
Dim i As Long
Dim strSQL As String
Dim rsTV As DAO.Recordset
Dim rsArr() As String

On Error Resume Next

Screen.MousePointer = vbHourglass
tmp = cmbox.Text: Clear cmbox    '= cmbox.Clear

strSQL = "Select " & fldName & " From Storage"

Set rsTV = DB.OpenRecordset(strSQL)
If err Then
    ToDebug "Err_FCNP: " & strSQL
    Screen.MousePointer = vbNormal
    Exit Sub
End If
On Error GoTo 0

ReDim rsArr(0)    'заполнение с 0
If Not (rsTV.BOF And rsTV.EOF) Then    'If rsTV.RecordCount > 0 Then
    rsTV.MoveLast: rsTV.MoveFirst

    For i = 1 To rsTV.RecordCount
        If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
        If IsNull(rsTV(0)) Then
        Else
            '            If Tokenize04(rsTV(0),  R(), ",;", True) > -1 Then
            '                For j = 0 To UBound(R)
            rsArr(UBound(rsArr)) = rsTV(0)
            ReDim Preserve rsArr(UBound(rsArr) + 1)
            'If Len(R(j)) = 0 Then PustoFlag = True
            '                Next j
        End If
        rsTV.MoveNext
    Next i
End If
If UBound(rsArr) > 0 Then
    TriQuickSortString rsArr        'sorts your string array
    remdups rsArr                   'removes dups
    For i = 0 To UBound(rsArr)
        If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
        '        If GetKeyState(vbKeyEscape) <0 Then Exit For
        'заполнить комбо без пустышек
        If Len(rsArr(i)) <> 0 Then SendMessage cmbox.hwnd, CB_ADDSTRING, 0, ByVal rsArr(i)
    Next i
End If

cmbox.Text = tmp

Set rsTV = Nothing
Screen.MousePointer = vbNormal
End Sub
Public Sub FillCombosParse(fldName As String, cmbox As ComboBox)
'заполняет комбики потрошенными значениями
'fldName - поле базы
Dim tmp As String
'mzt Dim temp As String
Dim i As Long, j As Long
'mzt Dim itmX As ListItem
Dim R() As String
Dim strSQL As String
Dim rsTV As DAO.Recordset
Dim rsArr() As String
Dim sDelim As String

On Error Resume Next

Screen.MousePointer = vbHourglass
tmp = cmbox.Text: Clear cmbox '= cmbox.Clear

strSQL = "Select " & fldName & " From Storage"

Set rsTV = DB.OpenRecordset(strSQL)
If err Then
    ToDebug "Err_FCP: " & strSQL
    Screen.MousePointer = vbNormal
    Exit Sub
End If
On Error GoTo 0

If LCase$(fldName) = "filename" Then
sDelim = "|"
Else
sDelim = ",;"
End If

ReDim rsArr(0)    'заполнение с 0
If Not (rsTV.BOF And rsTV.EOF) Then    'If rsTV.RecordCount > 0 Then
    rsTV.MoveLast: rsTV.MoveFirst

    For i = 1 To rsTV.RecordCount
        If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
'        If GetKeyState(vbKeyEscape) <0 Then Exit For

        If IsNull(rsTV(0)) Then
        Else

            If Tokenize04(rsTV(0), R(), sDelim, True) > -1 Then
                For j = 0 To UBound(R)
                    rsArr(UBound(rsArr)) = R(j)
                    ReDim Preserve rsArr(UBound(rsArr) + 1)
                    'If Len(R(j)) = 0 Then PustoFlag = True
                Next j
            End If
        End If
        rsTV.MoveNext
    Next i
End If
If UBound(rsArr) > 0 Then
    TriQuickSortString rsArr        'sorts your string array
    remdups rsArr                   'removes dups
    For i = 0 To UBound(rsArr)
        If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
'        If GetKeyState(vbKeyEscape) <0 Then Exit For
        'заполнить комбо без пустышек
        If Len(rsArr(i)) <> 0 Then SendMessage cmbox.hwnd, CB_ADDSTRING, 0, ByVal rsArr(i)
    Next i
End If

cmbox.Text = tmp

Set rsTV = Nothing
Screen.MousePointer = vbNormal
End Sub

Private Sub lChFiltCover_Click()
If ChFiltCover.Value = vbChecked Then
    ChFiltCover.Value = vbUnchecked
Else
    ChFiltCover.Value = vbChecked
End If

End Sub

Private Sub lChFiltSShots_Click()
If ChFiltSShots.Value = vbChecked Then
    ChFiltSShots.Value = vbUnchecked
Else
    ChFiltSShots.Value = vbChecked
End If
End Sub

Private Sub lChFiltWhole_Click()
If ChFiltWhole.Value = vbChecked Then
    ChFiltWhole.Value = vbUnchecked
Else
    ChFiltWhole.Value = vbChecked
End If
End Sub

'Private Sub LRes_Click()
'FillCombosParse "Director", ComboRes
'End Sub

Public Sub FillFilter()
'заполнение списков окна фильтра
Dim Contrl As Control

With FrmFilter
'    .FillCombosParse "Genre", .ComboGenre
'    .FillCombosParse "Country", .ComboCountry
'    .FillCombosParse "Debtor", .ComboDebtor
'    .FillCombosParse "Label", .cmbLabel
'    .FillCombosParse "MediaType", .ComboNCD    'таки не номер
'    '.FillCombosParse "FPS", .ComboFPS
'    .FillCombosParse "Resolution", .ComboFormat

    For Each Contrl In .Controls    'если пусто - просто пустые выпадающие списочки
        If TypeOf Contrl Is ComboBox Then
            If Contrl.style <> 2 Then 'не листы
                If Contrl.ListCount < 1 Then
                    Contrl.AddItem vbNullString
                End If
            End If
        End If
    Next

End With

End Sub

