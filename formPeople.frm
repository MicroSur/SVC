VERSION 5.00
Begin VB.Form FrmPeople 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " SVC: персоны"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chDupNoPic 
      Caption         =   "но добавить, если в базе нет фото и фото получено с сайта"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3480
      Value           =   1  'Checked
      Width           =   5895
   End
   Begin VB.ComboBox comboURL 
      Height          =   315
      ItemData        =   "formPeople.frx":0000
      Left            =   600
      List            =   "formPeople.frx":000D
      TabIndex        =   16
      Text            =   "http://www.kinopoisk.ru/level/4/people/"
      Top             =   120
      Width           =   4695
   End
   Begin VB.CheckBox chNoDup 
      Caption         =   "Без дубликатов по имени"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3180
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin VB.CommandButton ComFind 
      Caption         =   "найти"
      Default         =   -1  'True
      Height          =   315
      Left            =   9660
      TabIndex        =   14
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox TxtSearchHTML 
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Top             =   8520
      Width           =   9495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5550
      TabIndex        =   12
      Text            =   "1000"
      Top             =   3000
      Width           =   675
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Добавлять в базу актеров"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2940
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   4350
      ItemData        =   "formPeople.frx":0095
      Left            =   60
      List            =   "formPeople.frx":0097
      TabIndex        =   8
      Top             =   3960
      Width           =   10455
   End
   Begin VB.TextBox TxtTo 
      Height          =   315
      Left            =   5550
      TabIndex        =   7
      Text            =   "2"
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox TxtFrom 
      Height          =   315
      Left            =   4830
      TabIndex        =   6
      Text            =   "1"
      Top             =   2520
      Width           =   675
   End
   Begin VB.PictureBox PicFaceA 
      Height          =   3675
      Left            =   6300
      ScaleHeight     =   3615
      ScaleWidth      =   3135
      TabIndex        =   5
      Top             =   120
      Width           =   3195
   End
   Begin VB.TextBox TxtBIO 
      Height          =   1395
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1020
      Width           =   5595
   End
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   600
      MaxLength       =   255
      TabIndex        =   1
      Top             =   540
      Width           =   5595
   End
   Begin VB.TextBox TxtInd 
      Height          =   345
      Left            =   5340
      TabIndex        =   9
      Text            =   "1"
      Top             =   120
      Width           =   855
   End
   Begin SurVideoCatalog.XpB ComGetCur 
      Height          =   315
      Left            =   150
      TabIndex        =   18
      Top             =   2520
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      Caption         =   "Получить текущую"
      ButtonStyle     =   3
      PictureWidth    =   0
      PictureHeight   =   0
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB ComGetAuto 
      Height          =   315
      Left            =   2490
      TabIndex        =   19
      Top             =   2520
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      Caption         =   "Получить от и до"
      ButtonStyle     =   3
      PictureWidth    =   0
      PictureHeight   =   0
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin VB.Label Label3 
      Caption         =   "БИО"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Имя"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "задержка (мс)"
      Height          =   255
      Left            =   4230
      TabIndex        =   11
      Top             =   3060
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private objScript As New ClsScript 'public



Private Sub ComFind_Click()
Dim i As Integer
Dim start As Integer

If List1.ListIndex > -1 Then
start = List1.ListIndex + 1
End If

For i = start To List1.ListCount - 1
If InStr(1, List1.List(i), TxtSearchHTML.Text, vbTextCompare) > 0 Then
List1.Selected(i) = True
Exit Sub
End If
Next i
End Sub

Private Sub ComGetAuto_Click()
Dim i As Long

For i = TxtFrom To TxtTo
    Do
     DoEvents
     'If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
     If GetKeyState(vbKeyEscape) < 0 Then Exit For

      If IsNumeric(Text1) Then
        Text1.Text = CInt(Text1.Text)
        Sleep CInt(Text1.Text)
      Else
        Text1.Text = 1000
        Sleep Text1.Text
      End If
    
      If ComGetCur.Enabled = True Then Exit Do
    Loop
    'http://dvd.home-video.ru/cgi-bin/show.cgi?t=2&id=1
    TxtInd.Text = i
    ComGetCur_Click
Next
End Sub
' Set the list box's horizontal extent
'добавление горизонтальной прокрутки listbox, вещь местная для формы
Public Sub SetListboxScrollbar(lB As ListBox)
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

For i = 0 To lB.ListCount - 1
 new_len = 10 + ScaleX(TextWidth(lB.List(i)), ScaleMode, vbPixels)
 If max_len < new_len Then max_len = new_len
Next i

SendMessage lB.hwnd, LB_SETHORIZONTALEXTENT, max_len, 0
End Sub
Private Sub ComGetCur_Click()
Dim temp As String
Dim AName As String
Dim ABIO As String
'Dim APic As StdPicture
Dim i As Integer

AutoNoMessFlag = True

ComGetCur.Enabled = False

Select Case comboURL.Text
Case "http://dvd.home-video.ru/cgi-bin/show.cgi?t=2&id="
    temp = comboURL.Text & TxtInd.Text
Case "http://www.world-art.ru/people.php?id="
    temp = comboURL.Text & TxtInd.Text
Case "http://www.kinopoisk.ru/level/4/people/"
    temp = comboURL.Text & TxtInd.Text & "/view_bio/ok"
End Select

'DoEvents


PageText = OpenURLProxy(temp, "txt")
PageArray() = Split(PageText, vbLf)

TxtName.Text = vbNullString
TxtBIO.Text = vbNullString
TxtName.Refresh
TxtBIO.Refresh
Set PicFaceA = Nothing

'List1.Visible = False
'List1.Clear
'For i = 0 To UBound(PageArray)
'    List1.AddItem i & " |" & PageArray(i), i
'    SetListboxScrollbar List1 'тормоз
'Next
'List1.Visible = True

Dim ExitFlag As Boolean

Select Case comboURL.Text
Case "http://dvd.home-video.ru/cgi-bin/show.cgi?t=2&id="
    If Not AnalyzePage_DVDHome(AName, ABIO) Then ExitFlag = True
Case "http://www.world-art.ru/people.php?id="
    If Not AnalyzePage_WorldArt(AName, ABIO) Then ExitFlag = True
Case "http://www.kinopoisk.ru/level/4/people/"
    If Not AnalyzePage_KinoPoisk(AName, ABIO) Then ExitFlag = True
End Select
If ExitFlag Then
TxtBIO.Text = " НЕТ ДАННЫХ"
ComGetCur.Enabled = True
Exit Sub
End If

'если нет ни био (меньше режиссер сценарист звукооператор) ни картинки не писать
If (Len(ABIO) < 15) And (PicFaceA.Picture = 0) Then
    ComGetCur.Enabled = True
    TxtBIO.Text = " МАЛО ДАННЫХ"
    Exit Sub
End If

'поля
TxtName.Text = AName
TxtName.Refresh

ABIO = BlockDelete(ABIO, "&#", ";")    'убить спецсимволы по шаблону от и до
TxtBIO.Text = ABIO
TxtBIO.Refresh

TxtInd.Refresh

'DoEvents
'в базу
If Check1.Value = vbChecked Then
    '    addActFlag = True

    If ars.EditMode <> 0 Then ars.CancelUpdate
    ars.AddNew
    If PicFaceA.Picture <> 0 Then
        Pic2JPG PicFaceA, 2, "Face"    '2
    End If
    If Len(AName) > 255 Then AName = Left$(AName, 255)
    ars.Fields("Name") = AName
    ars.Fields("Bio") = ABIO
    ars.Update
End If

ComGetCur.Enabled = True
End Sub
Public Function AnalyzePage_DVDHome(ByRef n As String, ByRef b As String) As Boolean
Dim Line As String, Value As String
Dim LineNr As Integer
Dim BeginPos As Long, EndPos As Long
Dim temp As String

sReferer = "http://dvd.home-video.ru/"

'имя
LineNr = objScript.FindLine("<TITLE>", 0)    'с нуля
If LineNr > -1 Then
    Line = PageArray(LineNr)
    BeginPos = InStr(1, Line, "<TITLE>", vbTextCompare)
    If BeginPos > 0 Then BeginPos = BeginPos + 7
    EndPos = InStr(BeginPos, Line, "</TITLE>", vbTextCompare)
    If (EndPos > BeginPos) And (BeginPos > 0) Then
        Value = Mid$(Line, BeginPos, EndPos - BeginPos)
        n = objScript.HTML2TEXT(Value)
        If InStr(1, n, "ERROR", vbTextCompare) > 0 Then AnalyzePage_DVDHome = False: Exit Function
    End If
Else
    'нет имени
    AnalyzePage_DVDHome = False: Exit Function
End If

'Картинка
LineNr = objScript.FindLine("<H1>" & n, LineNr)
If LineNr > -1 Then
    LineNr = LineNr + 2    'по идее строка с картинкой (<p> в конце)
    Line = PageArray(LineNr)
    BeginPos = InStr(1, Line, "src=", vbTextCompare)
    If BeginPos > 0 Then
        BeginPos = BeginPos + 5

        EndPos = InStr(BeginPos, Line, ".jpg", vbTextCompare)
        If (EndPos > BeginPos) And (BeginPos > 0) Then
            Value = Mid$(Line, BeginPos, EndPos - BeginPos + 4)
            'Value = BaseAddress & Value
            getPeopleFlag = True
            OpenURLProxy Value, "pic"
        End If
    End If    'нет картинки
End If

'bio
'LineNr = FindLine("<TITLE>", 0)
LineNr = LineNr + 1    'на след строке - Профессия
'If InStr(BeginPos, Line, "</TD>", vbTextCompare) < 1 Then 'это и конец
If LineNr > -1 Then
    Line = PageArray(LineNr)
    BeginPos = 1    'InStr(1, Line, "<TITLE>", vbTextCompare)
    EndPos = Len(Line)    'InStr(BeginPos, Line, "</TD>", vbTextCompare)
    If (EndPos > BeginPos) And (BeginPos > 0) Then
        Value = Mid$(Line, BeginPos, EndPos - BeginPos)
        If Len(Value) > 0 Then b = objScript.HTML2TEXT(Value) & vbCrLf
    End If
End If
LineNr = LineNr + 1    'на след строке - Дата, Место
If LineNr > -1 Then
    Line = PageArray(LineNr)
    BeginPos = 1    'InStr(1, Line, "<TITLE>", vbTextCompare)
    EndPos = Len(Line)
    If (EndPos > BeginPos) And (BeginPos > 0) Then
        Value = Mid$(Line, BeginPos, EndPos - BeginPos)
        If Len(Value) > 0 Then b = b & objScript.HTML2TEXT(Value)
    End If
End If
'Else 'вернуть
'LineNr = LineNr - 1
'End If

Do
    LineNr = LineNr + 1    'на след строке - Био
    If LineNr > -1 Then
        Line = PageArray(LineNr)
        BeginPos = 1
        EndPos = InStr(BeginPos, Line, "</TD>", vbTextCompare)
        If (EndPos > BeginPos) And (BeginPos > 0) Then
            Value = Mid$(Line, BeginPos, EndPos - BeginPos)
            'b = b & vbCrLf & HTML2TEXT(Value)
            Exit Do
        Else
            temp = objScript.HTML2TEXT(Line)
            If Len(temp) > 0 Then b = b & " " & temp

        End If
    End If
Loop

'дупы
If DupsFound(n) Then
    AnalyzePage_DVDHome = False
    Exit Function 'нашли дуп
End If

AnalyzePage_DVDHome = True
End Function
Private Function DupsFound(s As String) As Boolean
'поискать дупы в базе актеров
Dim tmp As String, strSQL As String
Dim rsTmp As DAO.Recordset



If chNoDup.Value = vbChecked Then
    'менять "'" на "''"
    If InStr(s, "'") > 0 Then
        tmp = Replace(s, "'", "''")
        tmp = " Like '" & tmp & "'"
    Else
        tmp = " = '" & s & "'"
    End If
    
    If chDupNoPic.Value = vbUnchecked Then
        strSQL = "Select Count(Name) From Acter Where Name" & tmp
    Else
    'добавлять, если в базе нет фото
    If PicFaceA.Picture <> 0 Then
        'и если фото получено с инета
        strSQL = "Select Count(Name) From Acter Where ( (Name" & tmp & ") And (Face <> '') )"
    Else
        'иначе только по имени
        strSQL = "Select Count(Name) From Acter Where Name" & tmp
    End If
    
    End If
    
    'Debug.Print strSQL
    On Error Resume Next
    
    Set rsTmp = ADB.OpenRecordset(strSQL)
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        If rsTmp.RecordCount > 0 Then
            If rsTmp(0) <> 0 Then
            'нашли имя или (имя и есть картинка)
     
                TxtBIO.Text = " ДУБЛИКАТ"
                DupsFound = True

            End If
        End If
    End If
    
    Set rsTmp = Nothing
End If
End Function
Public Function AnalyzePage_KinoPoisk(ByRef n As String, ByRef b As String) As Boolean
Dim Line As String, Value As String, block As String
Dim LineNr As Integer    ', i As Integer
Dim BeginPos As Long ', EndPos As Long
'Dim temp As String
Dim tmp As String, btmp As String
Dim RusName As String, LatName As String
'Dim strSQL As String

sReferer = "http://www.kinopoisk.ru/"

'имя
LineNr = objScript.FindLine("<H1 class=", 0)
If LineNr > -1 Then
 Line = PageArray(LineNr)
 RusName = objScript.HTML2TEXT(Line) 'русское имя
 BeginPos = LineNr
End If

LineNr = objScript.FindLine(">имя (лат.)<", LineNr + 1)
If LineNr > -1 Then
 Line = PageArray(LineNr)
 Value = objScript.HTML2TEXT(Line)
 LatName = Replace(Value, "имя (лат.) ", "")
 If LatName = "-" Then LatName = ""
End If

If (Len(RusName) <> 0) And (Len(LatName) <> 0) Then
    n = RusName & " (" & LatName & ")"
ElseIf Len(RusName) <> 0 Then
    n = RusName
ElseIf Len(LatName) <> 0 Then
    n = LatName
Else
    'нет имени
    AnalyzePage_KinoPoisk = False: Exit Function
End If


'био
LineNr = objScript.FindLine(">карьера<", BeginPos)
If LineNr > -1 Then
    Line = PageArray(LineNr)
    Value = objScript.HTML2TEXT(Line)
    tmp = Replace(Value, "карьера ", "")
    If tmp <> "-" Then btmp = tmp & vbCrLf
End If
LineNr = objScript.FindLine(">рост<", LineNr + 1)
If LineNr > -1 Then
    Line = PageArray(LineNr)
    tmp = objScript.HTML2TEXT(Line)
    If tmp <> "-" Then btmp = btmp & tmp & vbCrLf
End If
LineNr = objScript.FindLine(">дата рождения<", LineNr + 1)
If LineNr > -1 Then
    LineNr = objScript.FindLine("<td><a href=""/level/10", LineNr + 1)
    If LineNr > -1 Then
        Line = PageArray(LineNr)
        tmp = objScript.HTML2TEXT(Line)
        If tmp <> "-" Then btmp = btmp & tmp & vbCrLf
    End If
End If
LineNr = objScript.FindLine(">место рождения<", LineNr + 1)
If LineNr > -1 Then
    Line = PageArray(LineNr)
    Value = objScript.HTML2TEXT(Line)
    tmp = Replace(Value, "место рождения ", "")
    If tmp <> "-" Then btmp = btmp & tmp & vbCrLf
End If
LineNr = objScript.FindLine(">супруг", LineNr + 1)
If LineNr > -1 Then
    Line = PageArray(LineNr)
    tmp = objScript.HTML2TEXT(Line)
    If tmp <> "-" Then btmp = btmp & tmp & vbCrLf
End If
block = objScript.GetBlockFrom(">Биография<", "<td colspan=3 height=13>")
If Len(block) <> 0 Then
 Value = objScript.HTML2TEXT(block)
 'tmp = Trim$(Right$(Value, Len(Value) - 9))
 tmp = Right$(Value, Len(Value) - 9)
 btmp = btmp & vbCrLf & tmp
End If

b = btmp
'Картинка
getPeopleFlag = True
Value = "http://www.kinopoisk.ru/images/actor/" & TxtInd.Text & ".jpg"
OpenURLProxy Value, "pic"

If DupsFound(n) Then
    AnalyzePage_KinoPoisk = False
    Exit Function 'нашли дуп
End If

AnalyzePage_KinoPoisk = True

End Function
Public Function AnalyzePage_WorldArt(ByRef n As String, ByRef b As String) As Boolean
Dim Line As String, Value As String, block As String
Dim LineNr As Integer    ', i As Integer
Dim BeginPos As Long, EndPos As Long

sReferer = "http://www.world-art.ru/"
'имя
LineNr = objScript.FindLine("<TITLE>", 0)    'с нуля
If LineNr > -1 Then
    Line = PageArray(LineNr)
    BeginPos = InStr(1, Line, "<title>World Art | Персоны | ", vbTextCompare)
    If BeginPos > 0 Then
    BeginPos = BeginPos + 29
    EndPos = InStr(BeginPos, Line, "</title>", vbTextCompare)
    If EndPos > BeginPos Then
        Value = Mid$(Line, BeginPos, EndPos - BeginPos)
        n = objScript.HTML2TEXT(Value)
        If InStr(1, n, "ERROR", vbTextCompare) > 0 Then AnalyzePage_WorldArt = False: Exit Function
        n = Replace(n, "[", "(")
        n = Replace(n, "]", ")")
    End If
    End If
Else
    'нет имени
    AnalyzePage_WorldArt = False: Exit Function
End If

'Картинка
LineNr = objScript.FindLine("img/people/", LineNr)
If LineNr > -1 Then
    Line = PageArray(LineNr)
    BeginPos = InStr(1, Line, "img/people/", vbTextCompare)
    If BeginPos > 0 Then
        EndPos = InStr(BeginPos, Line, ".jpg", vbTextCompare)
        If EndPos > BeginPos Then
            Value = Mid$(Line, BeginPos, EndPos - BeginPos + 4)
            Value = "http://www.world-art.ru/" & Value
            'Value = BaseAddress & Value
            getPeopleFlag = True
'Value = "http://www.world-art.ru/image_convert.php?id=1&type=people&pack=10000"
            OpenURLProxy Value, "pic"
        End If
    End If    'нет картинки
End If

'bio
block = objScript.GetBlockFrom(">Дата рождения:<", "<td width='100%' height=1")
If Len(block) <> 0 Then
b = objScript.HTML2TEXT(block)
End If

'дупы
If DupsFound(n) Then
    AnalyzePage_WorldArt = False
    Exit Function 'нашли дуп
End If

AnalyzePage_WorldArt = True
End Function
Private Sub Form_Activate()
frmPeopleFlag = True
End Sub


Private Sub Form_Load()
Me.Icon = FrmMain.Icon
End Sub

Private Sub Form_Resize()
'''Background
'If lngBrush <> 0 Then
'GetClientRect hwnd, rctMain
'FillRect hdc, rctMain, lngBrush
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objScript = Nothing
frmPeopleFlag = False
End Sub

