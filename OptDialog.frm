VERSION 5.00
Begin VB.Form OptDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   9825
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   17205
   Icon            =   "OptDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   17205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ComOptSave 
      Caption         =   "Сохранить"
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4140
      Width           =   2655
   End
   Begin VB.Frame FrFont 
      Caption         =   "Шрифты"
      Height          =   4485
      Left            =   2760
      TabIndex        =   38
      Top             =   60
      Width           =   6795
      Begin VB.TextBox TextFontH 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   1380
         Width           =   4545
      End
      Begin VB.TextBox TextFontV 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   660
         Width           =   4545
      End
      Begin VB.CommandButton ComFontH 
         Caption         =   "Выбрать"
         Height          =   255
         Left            =   4800
         TabIndex        =   45
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton ComFontV 
         Caption         =   "Выбрать"
         Height          =   255
         Left            =   4800
         TabIndex        =   44
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton ComFontLV 
         Caption         =   "Выбрать"
         Height          =   255
         Left            =   4800
         TabIndex        =   43
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TextFontLV 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   2160
         Width           =   4560
      End
      Begin VB.CommandButton ComColorPick 
         Caption         =   "Цвет фона"
         Height          =   255
         Left            =   5340
         TabIndex        =   41
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CommandButton ComCoverVertFillColor 
         Caption         =   "Цвет фона"
         Height          =   255
         Left            =   5340
         TabIndex        =   40
         Top             =   780
         Width           =   1095
      End
      Begin VB.CommandButton ComCoverHorFillColor 
         Caption         =   "Цвет фона"
         Height          =   255
         Left            =   5340
         TabIndex        =   39
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label LblFontVHid 
         Caption         =   "Label20"
         Height          =   195
         Left            =   3960
         TabIndex        =   53
         Top             =   1140
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label LblFontHHid 
         Caption         =   "Label20"
         Height          =   195
         Left            =   3840
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label LHFont 
         Caption         =   "Горизонтальный шрифт обложки"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label LVFont 
         Caption         =   "Вертикальный шрифт обложки"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label LLVFont 
         Caption         =   "Шрифт списка фильмов"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1800
         Width           =   2880
      End
      Begin VB.Label LblFontLVHid 
         Caption         =   "Label20"
         Height          =   195
         Left            =   3840
         TabIndex        =   48
         Top             =   1980
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.Frame FrExport 
      Caption         =   "Экспорт"
      Height          =   4815
      Left            =   9900
      TabIndex        =   30
      Top             =   60
      Width           =   6795
      Begin VB.OptionButton OptHtml 
         Caption         =   "Случайно"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   540
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptHtml 
         Caption         =   "Имя файла"
         Height          =   315
         Index           =   0
         Left            =   1860
         TabIndex        =   35
         Top             =   540
         Width           =   1635
      End
      Begin VB.OptionButton OptHtml 
         Caption         =   "Название"
         Height          =   315
         Index           =   1
         Left            =   3600
         TabIndex        =   34
         Top             =   540
         Width           =   2055
      End
      Begin VB.CommandButton ComMarkAll 
         Caption         =   "Все"
         Height          =   315
         Left            =   2400
         TabIndex        =   33
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton ComUnmarkAll 
         Caption         =   "Ничего"
         Height          =   315
         Left            =   4620
         TabIndex        =   32
         Top             =   3120
         Width           =   1995
      End
      Begin VB.ListBox LstExport 
         BackColor       =   &H00C0FFFF&
         Columns         =   5
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   31
         Top             =   1020
         Width           =   6495
      End
      Begin VB.Label LblOptHtml 
         Caption         =   "имена для html"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   300
         Width           =   4515
      End
   End
   Begin VB.Frame FrGlobal 
      Caption         =   "Общие"
      Height          =   4755
      Left            =   9900
      TabIndex        =   14
      Top             =   4920
      Width           =   6795
      Begin VB.CommandButton ComOptPath 
         Caption         =   "Путь"
         Height          =   315
         Left            =   6120
         TabIndex        =   29
         Top             =   2460
         Width           =   555
      End
      Begin VB.ComboBox ComboCDHid 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Text            =   "D:\"
         Top             =   2460
         Width           =   5910
      End
      Begin VB.CheckBox CheckLoadLastBD 
         Caption         =   "Открывать окно каталога"
         Height          =   285
         Left            =   180
         TabIndex        =   23
         Top             =   300
         Value           =   1  'Checked
         Width           =   6180
      End
      Begin VB.TextBox TextQJPGHid 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   180
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "80"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CheckBox CheckCDAutorun 
         Caption         =   "no CD"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6420
         TabIndex        =   21
         Top             =   -60
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox CheckSavBigPix 
         Caption         =   "Кадры с реальным разрешением"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   600
         Value           =   1  'Checked
         Width           =   6435
      End
      Begin VB.CommandButton ComLangOk 
         Caption         =   "Поменять"
         Height          =   315
         Left            =   3420
         TabIndex        =   19
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox ComboLangHid 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "OptDialog.frx":08CA
         Left            =   120
         List            =   "OptDialog.frx":08CC
         TabIndex        =   18
         Text            =   "Русский"
         Top             =   3360
         Width           =   3030
      End
      Begin VB.ComboBox cboPrinterHid 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "OptDialog.frx":08CE
         Left            =   120
         List            =   "OptDialog.frx":08D0
         TabIndex        =   17
         Text            =   "Null"
         Top             =   4200
         Width           =   6495
      End
      Begin VB.CheckBox ChDSFilt 
         Caption         =   "DirectShow фильтр программы"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   900
         Value           =   1  'Checked
         Width           =   6375
      End
      Begin VB.CheckBox ChOnlyTitle 
         Caption         =   "Только названия"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   1200
         Width           =   6435
      End
      Begin VB.Label LCD 
         Caption         =   "Буква CD"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2100
         Width           =   6375
      End
      Begin VB.Label LQJPG 
         Caption         =   "JPEG (0-100)"
         Height          =   285
         Left            =   720
         TabIndex        =   27
         Top             =   1620
         Width           =   5835
      End
      Begin VB.Label LPrinter 
         Caption         =   "Принтер"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3900
         Width           =   6435
      End
      Begin VB.Label lblLang 
         Caption         =   "Язык"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   2895
      End
   End
   Begin VB.Frame FrameBD 
      Caption         =   "Базы"
      Height          =   4785
      Left            =   2820
      TabIndex        =   1
      Top             =   4800
      Width           =   6795
      Begin VB.CommandButton ComNewBD 
         BackColor       =   &H00C0C0FF&
         Caption         =   "New BD"
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton ComOpenBD 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Add BD"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton ComCompact 
         Caption         =   "Сжать БД"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3300
         Width           =   2055
      End
      Begin VB.CommandButton ComCompactA 
         Caption         =   "Сжать базу актеров"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   4020
         Width           =   2055
      End
      Begin VB.CommandButton ComPassword 
         Caption         =   "Пароль"
         Height          =   255
         Left            =   5580
         TabIndex        =   4
         Top             =   4380
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox LstBases 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   1785
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   6495
      End
      Begin VB.CommandButton ComDelBD 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Del BD"
         Height          =   375
         Left            =   2340
         TabIndex        =   2
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label LabelCurrBDHid 
         Caption         =   ">"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   2220
         Width           =   6450
      End
      Begin VB.Label LBDSize 
         Caption         =   "Размер БД (Kб):"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   3420
         Width           =   1575
      End
      Begin VB.Label LBDSizeHid 
         Caption         =   "0"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   3420
         Width           =   1035
      End
      Begin VB.Label LABDSizeHid 
         Caption         =   "0"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   4140
         Width           =   1095
      End
      Begin VB.Label LABDSize 
         Caption         =   "Размер AБД (Kб):"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   4140
         Width           =   1575
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3915
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   6906
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "OptDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'гориз. скролл листбокса
Private Const LB_SETHORIZONTALEXTENT = &H194



Private Sub ComUpDownHid_Click()

End Sub

Private Sub cboPrinterHid_LostFocus()
'   cmdSelect.Enabled = (cboPrinterHid.ListIndex >= 0)
If cboPrinterHid <> "Null" Then
    If FrmMain.SelectPrinter(cboPrinterHid) Then
        myMsgBox msgsvc(14), vbOKOnly, , Me.hwnd
        cboPrinterHid = "Null"
        Exit Sub
    Else
        FrmMain.SelectPrinter cboPrinterHid
        optsaved = False
    End If
End If

End Sub

Private Sub ChDSFilt_Click()
optsaved = False
End Sub

Private Sub CheckLoadLastBD_Click()
optsaved = False
End Sub

Private Sub CheckSavBigPix_Click()
optsaved = False
End Sub

Private Sub ChOnlyTitle_Click()
optsaved = False
End Sub

Private Sub ComboCDHid_Change()
optsaved = False
End Sub
Private Sub ComboCDHid_DropDown()

Call Drives
optsaved = False
End Sub

Private Sub ComboLangHid_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub ComboLangHid_LostFocus()
'get language from ini
Dim LangCount As Integer
Dim temp As String
Dim i As Integer

'ComboLangHid.Clear

'LangCount = Int(Val(VBGetPrivateProfileString("Language", "LCount", iniFileName)))
LangCount = Int(Val(VBGetPrivateProfileString("Language", "LCount", iniGlobalFileName)))

If LangCount < 1 Then LangCount = 1
For i = 1 To LangCount
temp = VBGetPrivateProfileString("Language", "L" & i, iniGlobalFileName)


If temp = OptDialog.ComboLangHid Then lngFileName = App.Path + "\" + _
VBGetPrivateProfileString("Language", "L" & i & "File", iniGlobalFileName)
Next i
optsaved = False
End Sub

Private Sub ComColorPick_Click()
Dim c As Long
Dim cd As New cCommonDialog
Dim Ret As Long

cd.CustomColor(0) = LVBackColor

Ret = cd.VBChooseColor(c, True, False, False, Me.hwnd)
If Ret = 0 Then Exit Sub
LVBackColor = c
Set cd = Nothing
NoSetColorFlag = False
FrmMain.setForecolor
optsaved = False
End Sub

Private Sub ComCompact_Click()
Dim tempDBname As String

If BaseReadOnly Then myMsgBox msgsvc(24), vbInformation, , Me.hwnd: Exit Sub
If BaseReadOnlyU Then myMsgBox msgsvc(22), vbInformation, , Me.hwnd: Exit Sub

ToDebug "Сжатие базы фильмов ..."
'rs.Close: DB.Close
Set rs = Nothing
Set DB = Nothing
Screen.MousePointer = vbHourglass

tempDBname = App.Path + "\svcmdb.tmp"

On Error Resume Next 'GoTo err
If Dir(bdname) <> vbNullString Then
    If Dir(tempDBname) <> vbNullString Then Kill tempDBname
    
        DBEngine.CompactDatabase bdname, tempDBname
        If err.Number = 3031 Then 'пассворд
            SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
            pwd = myInputBox(ComPassword.Caption & vbCrLf & bdname)
            On Error GoTo err
            DBEngine.CompactDatabase bdname, tempDBname, , , ";PWD=" & pwd
        End If

    If Dir(tempDBname) <> vbNullString Then Kill bdname
    If Dir(bdname) = vbNullString Then Name tempDBname As bdname
    If Dir(bdname) = vbNullString Then FileCopy tempDBname, bdname: ToDebug "... ok"
End If

LBDSizeHid = FileLen(bdname) / 1024
'Set DB = DBEngine.OpenDatabase(bdname, False)
'Set rs = DB.OpenRecordset("Storage", dbOpenTable)


CurSearch = 0 'потом будет 1
InitFlag = True
Screen.MousePointer = vbNormal

Exit Sub
err:
Screen.MousePointer = vbNormal
MsgBox err.Description, vbCritical
ToDebug "... " & err.Description

End Sub



Private Sub ComCompactA_Click()
Dim tempDBname As String
If BaseAReadOnly Then myMsgBox msgsvc(25), vbInformation, , Me.hwnd: Exit Sub
If BaseAReadOnlyU Then myMsgBox msgsvc(23), vbInformation, , Me.hwnd: Exit Sub

ToDebug "Сжатие базы актеров ..."
'ars.Close: ADB.Close
Set ars = Nothing
Set ADB = Nothing

Screen.MousePointer = vbHourglass

tempDBname = App.Path + "\pmdb.tmp"

On Error GoTo err
If Dir(abdname) <> vbNullString Then
    If Dir(tempDBname) <> vbNullString Then Kill tempDBname
    DBEngine.CompactDatabase abdname, tempDBname
    If Dir(tempDBname) <> vbNullString Then Kill abdname
    If Dir(abdname) = vbNullString Then Name tempDBname As abdname
    If Dir(abdname) = vbNullString Then FileCopy tempDBname, abdname: ToDebug "... ok"
End If

LABDSizeHid.Caption = FileLen(abdname) / 1024
Set ADB = DBEngine.OpenDatabase(abdname, False)
Set ars = ADB.OpenRecordset("Acter", dbOpenTable)

Screen.MousePointer = vbNormal

Exit Sub
err:
Screen.MousePointer = vbNormal
MsgBox err.Description, vbCritical
ToDebug "... " & err.Description

End Sub



Private Sub ComCoverHorFillColor_Click()
Dim c As Long
Dim cd As New cCommonDialog
Dim Ret As Long

cd.CustomColor(0) = CoverHorBackColor
Ret = cd.VBChooseColor(c, True, False, False, Me.hwnd)
If Ret = 0 Then Exit Sub
CoverHorBackColor = c

Set cd = Nothing

'Call setForecolor
optsaved = False
End Sub

Private Sub ComCoverVertFillColor_Click()
Dim c As Long
Dim cd As New cCommonDialog
Dim Ret As Long

cd.CustomColor(0) = CoverVertBackColor
Ret = cd.VBChooseColor(c, True, False, False, Me.hwnd)
If Ret = 0 Then Exit Sub
CoverVertBackColor = c
'PicCoverPaper.Line (35, 145)-(172, 262), c, BF

Set cd = Nothing

'Call setForecolor
optsaved = False
End Sub

Private Sub ComDelBD_Click()
'Dim i As Integer
'For i = 0 To LstBases.ListCount - 1
'    If LstBases.Selected(i) Then
If LstBases.ListIndex > -1 Then LstBases.RemoveItem LstBases.ListIndex
'        Exit For
'    End If
'Next i

If LstBases.ListCount > 0 Then
    bdname = LstBases.List(0)
    LstBases.Selected(0) = True
    LabelCurrBDHid.Caption = bdname
    LBDSizeHid.Caption = FileLen(bdname) / 1024
    ComCompact.Enabled = True
Else
    bdname = vbNullString
    LabelCurrBDHid.Caption = bdname
    LBDSizeHid.Caption = vbNullString  'FileLen(bdname) / 1024
    ComCompact.Enabled = False
End If

optsaved = False
FrmMain.AddTabsLV
InitFlag = True
End Sub

Private Sub ComFontH_Click()
   Dim cd As New cCommonDialog
      Dim sFnt As StdFont
      Dim temp As String
      
Set sFnt = LblFontHHid.Font
If HFontColor = 0 Then HFontColor = 1

 If (cd.VBChooseFont(sFnt, , Me.hwnd, HFontColor)) Then
 
 LblFontHHid.Font = sFnt
 
temp = " " & LblFontHHid.Font.Size
If LblFontHHid.Font.Bold Then temp = temp + " Bold"
If LblFontHHid.Font.Italic Then temp = temp + " Italic"
TextFontH.Text = LblFontHHid.Font.name & temp

 End If
 
optsaved = False

Set cd = Nothing
Set sFnt = Nothing
End Sub

Private Sub ComFontLV_Click()
   Dim cd As New cCommonDialog
      Dim sFnt As StdFont
      Dim temp As String
      
Set sFnt = LblFontLVHid.Font
If LVFontColor = 0 Then LVFontColor = 1

 If (cd.VBChooseFont(sFnt, , Me.hwnd, LVFontColor)) Then
 
 LblFontLVHid.Font = sFnt
 
temp = " " & LblFontLVHid.Font.Size
If LblFontLVHid.Font.Bold Then temp = temp + " Bold"
If LblFontLVHid.Font.Italic Then temp = temp + " Italic"
TextFontLV.Text = LblFontLVHid.Font.name & temp
'Set TextFontLV.Font = LblFontLVHid.Font

 End If
 
NoSetColorFlag = False
FrmMain.setForecolor

optsaved = False

Set cd = Nothing
Set sFnt = Nothing
End Sub

Private Sub ComFontV_Click()
   Dim cd As New cCommonDialog
      Dim sFnt As StdFont
            Dim temp As String
Set sFnt = LblFontVHid.Font

If VFontColor = 0 Then VFontColor = 1

 If (cd.VBChooseFont(sFnt, , Me.hwnd, VFontColor, , , CF_NoOemFonts Or CF_ScalableOnly)) Then
 
 LblFontVHid.Font = sFnt
 
'font
temp = " " & LblFontVHid.Font.Size
If LblFontVHid.Font.Bold Then temp = temp + " Bold"
If LblFontVHid.Font.Italic Then temp = temp + " Italic"
TextFontV.Text = LblFontVHid.Font.name & temp
 End If
 
optsaved = False

Set cd = Nothing
Set sFnt = Nothing
End Sub


Public Sub ComLangOk_Click()
Dim Contrl As Control
Dim i As Integer
Dim temp As String

If Dir(lngFileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = vbNullString Then
Call myMsgBox(msgsvc(18), vbOKOnly, , Me.hwnd)
Exit Sub
End If

If LastLangFile = lngFileName Then Exit Sub 'не перечитывать
Screen.MousePointer = vbHourglass

ToDebug "Чтение файла локализации: " & lngFileName

LockWindowUpdate Me.hwnd
'On Error Resume Next
'err = 0



For Each Contrl In FrmMain.Controls
'If (TypeOf Contrl Is ComboBox _
Or TypeOf Contrl Is ListView _
Or TypeOf Contrl Is Label _
Or TypeOf Contrl Is Frame _
Or TypeOf Contrl Is Slider _
Or TypeOf Contrl Is CommandButton _
Or TypeOf Contrl Is PictureBox _
Or TypeOf Contrl Is Image _
Or TypeOf Contrl Is VerticalMenu _
) Then
'msgbox
'popup menu
'checkbox

If Right$(Contrl.name, 3) <> "Hid" Then

If Contrl.name = "VerticalMenu" Then '                      VerticalMenu
'Contrl.MenuCaption = ReadLang(Contrl.name & ".MenuCaption")
'For i = 1 To 8
'Contrl.MenuItemCur = i
Contrl.lbl_1 = ReadLang(Contrl.name & ".MenuItemCaption" & 1)
Contrl.lbl_2 = ReadLang(Contrl.name & ".MenuItemCaption" & 2)
Contrl.lbl_3 = ReadLang(Contrl.name & ".MenuItemCaption" & 3)
Contrl.lbl_4 = ReadLang(Contrl.name & ".MenuItemCaption" & 4)
Contrl.lbl_5 = ReadLang(Contrl.name & ".MenuItemCaption" & 5)
Contrl.lbl_6 = ReadLang(Contrl.name & ".MenuItemCaption" & 6)
Contrl.lbl_7 = ReadLang(Contrl.name & ".MenuItemCaption" & 7)
Contrl.lbl_8 = ReadLang(Contrl.name & ".MenuItemCaption" & 8)
'Next i
End If 'Contrl Is VerticalMenu

If TypeOf Contrl Is Label Then '                           Label
Contrl.Caption = ReadLang(Contrl.name & ".Caption")
End If

If TypeOf Contrl Is Frame Then '                           Frame
Contrl.Caption = ReadLang(Contrl.name & ".Caption")
End If

If TypeOf Contrl Is ListView Then '                        ListView
    If Contrl.name = "ListView" Then
        For i = 1 To FrmMain.ListView.ColumnHeaders.Count
            Contrl.ColumnHeaders(i).Text = ReadLang(Contrl.name & ".CH" & i)
        Next i
    End If
End If

If TypeOf Contrl Is CommandButton Then '                    CommandButton
Contrl.Caption = ReadLang(Contrl.name & ".Caption")
Contrl.ToolTipText = ReadLang(Contrl.name & ".ToolTip")
End If

If TypeOf Contrl Is ComboBox Then '                         Combo
    If Contrl.name = "ComboGenre" Then
    FrmMain.ComboGenre.Clear
        For i = 1 To 22
            FrmMain.ComboGenre.AddItem (ReadLang("ComboGenre.Item" & i))
        Next i
    End If
    If Contrl.name = "CombFind" Then
    FrmMain.CombFind.Clear
        For i = 0 To 8
        FrmMain.CombFind.AddItem (ReadLang("CombFind.Item" & i))
        FrmMain.CombFind.ItemData(i) = i
        Next i
    FrmMain.CombFind.Text = FrmMain.CombFind.List(0): FrmMain.CombFind.SelLength = 0: FrmMain.CombFind.SelText = " "
    
    
    
    End If
    If Contrl.name = "ComboCountry" Then
    FrmMain.ComboCountry.Clear
        For i = 1 To 20
            FrmMain.ComboCountry.AddItem (ReadLang("ComboCountry.Item" & i))
        Next i
    End If
    If Contrl.name = "ComboSites" Then
    FrmMain.ComboSites.Clear
        For i = 1 To 6
        temp = ReadLang("ComboSites.Item" & i)
        If temp <> vbNullString Then FrmMain.ComboSites.AddItem temp
        Next i
    FrmMain.ComboSites.Text = FrmMain.ComboSites.List(0)
    End If

End If

If TypeOf Contrl Is OptionButton Then '                     Option
For i = OptHtml.LBound To OptHtml.UBound
OptHtml(i).Caption = ReadLang(Contrl.name & i & ".Caption")
Next

End If

If TypeOf Contrl Is Menu Then '                             Menu
Contrl.Caption = ReadLang(Contrl.name & ".Caption")
End If

If TypeOf Contrl Is CheckBox Then '                         CheckBox
Contrl.Caption = ReadLang(Contrl.name & ".Caption")
End If

'TabStripCover.Tabs(1).Caption
If TypeOf Contrl Is TabStrip Then '                        TabStripCover
Select Case Contrl.name
Case "TabStripCover"
Contrl.Tabs(1).Caption = ReadLang(Contrl.name & ".Tabs(1).Caption")
Contrl.Tabs(2).Caption = ReadLang(Contrl.name & ".Tabs(2).Caption")
Contrl.Tabs(3).Caption = ReadLang(Contrl.name & ".Tabs(3).Caption")
Contrl.Tabs(4).Caption = ReadLang(Contrl.name & ".Tabs(4).Caption")
Case "TabStrAdEd"
Contrl.Tabs(1).Caption = ReadLang(Contrl.name & ".Tabs(1).Caption")
Contrl.Tabs(2).Caption = ReadLang(Contrl.name & ".Tabs(2).Caption")
End Select

End If


End If '(TypeOf Contrl
Next


'                                                           lstExport
'For i = 0 To LstExport.ListCount - 1
LstExport.List(0) = FrmMain.ListView.ColumnHeaders(1).Text 'назв
LstExport.List(1) = FrmMain.ListView.ColumnHeaders(2).Text 'метка
LstExport.List(2) = FrmMain.ListView.ColumnHeaders(3).Text 'жанр
LstExport.List(3) = FrmMain.ListView.ColumnHeaders(4).Text 'год
LstExport.List(4) = FrmMain.ListView.ColumnHeaders(5).Text 'Произв.
LstExport.List(5) = FrmMain.ListView.ColumnHeaders(6).Text 'Реж
LstExport.List(6) = FrmMain.ListView.ColumnHeaders(7).Text 'Роль
LstExport.List(7) = FrmMain.ListView.ColumnHeaders(8).Text 'время
LstExport.List(8) = FrmMain.ListView.ColumnHeaders(9).Text 'формат
LstExport.List(9) = FrmMain.ListView.ColumnHeaders(10).Text 'звук
LstExport.List(10) = FrmMain.ListView.ColumnHeaders(11).Text 'кс
LstExport.List(11) = FrmMain.ListView.ColumnHeaders(12).Text 'разм
LstExport.List(12) = FrmMain.ListView.ColumnHeaders(13).Text 'nn
LstExport.List(13) = FrmMain.ListView.ColumnHeaders(14).Text 'видео
LstExport.List(14) = FrmMain.ListView.ColumnHeaders(15).Text 'имя ф
LstExport.List(15) = FrmMain.ListView.ColumnHeaders(16).Text 'долг
LstExport.List(16) = FrmMain.ListView.ColumnHeaders(17).Text 'сер
LstExport.List(17) = FrmMain.ListView.ColumnHeaders(18).Text 'прим

LstExport.List(18) = FrmMain.LAnnot.Caption                    'аннот
LstExport.List(19) = FrmMain.LblCover.Caption                  'облож
LstExport.List(20) = FrmMain.LblScrShot.Caption & " 1"
LstExport.List(21) = FrmMain.LblScrShot.Caption & " 2"
LstExport.List(22) = FrmMain.LblScrShot.Caption & " 3"
'Next

'                                                           msgbox
For i = 1 To UBound(msgsvc)
msgsvc(i) = Change2lfcr(ReadLang("msg" & i))
Next i

'edit/add
FrmMain.ComOpenHid.Caption = FrmMain.ComOpen.Caption
FrmMain.ComAddHid.Caption = FrmMain.ComAdd.Caption
FrmMain.ComDelHid.Caption = FrmMain.ComDel.Caption
FrmMain.ComSaveRecHid.Caption = FrmMain.ComSaveRec.Caption
FrmMain.ComCancelHid.Caption = FrmMain.ComCancel.Caption

'
FrameViewCaption = FrmMain.FrameView.Caption
FrameActerCaption = FrmMain.FrameActer.Caption
LActMarkCountCaption = FrmMain.LActMarkCount.Caption

LockWindowUpdate 0
LastLangFile = lngFileName
LstExport.ListIndex = LstExport.ListCount - 1: LstExport.TopIndex = 0 'чтоб не видно выделения

Screen.MousePointer = vbNormal
End Sub



Public Function ReadLang(Itm As String) As String
ReadLang = VBGetPrivateProfileString("Language", Itm, lngFileName)
End Function

Private Sub ComMarkAll_Click()
Dim i As Integer
For i = 0 To OptDialog.LstExport.ListCount - 1
OptDialog.LstExport.Selected(i) = True
Next

End Sub


Private Sub ComNewBD_Click()
Dim a() As Byte
Dim fn As Integer
   Dim cd As New cCommonDialog
   Dim sFile As String
  
ToDebug "Создать новую базу фильмов ..."

   If (cd.VBGetSaveFileName( _
      sFile, _
      Filter:="MDB (*.mdb)|*.mdb|All Files (*.*)|*.*", _
      FilterIndex:=1, _
      DefaultExt:="mdb", _
      Owner:=Me.hwnd)) Then
   End If
   
a() = LoadResData(101, "CUSTOM")

fn = FreeFile
If sFile <> vbNullString Then

If rs Is Nothing Then GoTo openf
    rs.Close
    DB.Close
    Set rs = Nothing
    Set DB = Nothing
    
Set cd = Nothing

openf:
   
Open sFile For Binary Access Write As fn
     Put #fn, , a()
Close #fn

ToDebug "... ok"


End If
End Sub


Private Sub ComOpenBD_Click()
   Dim cd As New cCommonDialog
   Dim sFile As String
'Dim oldname As String
'oldname = bdname
Dim i As Integer
Dim temp As String


ToDebug "Добавить другую базу фильмов ..."

   If (cd.VBGetOpenFileName( _
      sFile, _
      Filter:="MDB Files (*.mdb)|*.mdb|All Files (*.*)|*.*", _
      FilterIndex:=1, _
      DefaultExt:="mdb", _
      Owner:=Me.hwnd)) Then
      temp = sFile
   End If
   Me.SetFocus

If (temp <> vbNullString) And (sFile <> vbNullString) Then

For i = 0 To LstBases.ListCount - 1
    If LstBases.List(i) = temp Then
        ToDebug temp & "... уже есть в списке"
    Exit Sub 'повтор
    End If
Next i

Set rs = Nothing
Set DB = Nothing
FrmMain.Caption = "SurVideoCatalog"
bdname = temp
If Not FrmMain.OpenDB Then pwd = vbNullString: Exit Sub

'add to list
LstBases.AddItem temp: SetListboxScrollbar LstBases
LstBases.Selected(LstBases.ListCount - 1) = True
FrmMain.AddTabsLV

InitFlag = True
NoDBFlag = False

FrmMain.ReadINI GetNameFromPathAndName(bdname)
NoSetColorFlag = False
'Call setForecolor

    LabelCurrBDHid.Caption = bdname
    LBDSizeHid.Caption = FileLen(bdname) / 1024
    ComCompact.Enabled = True
    


ToDebug "... ok"
ComOptSave.Enabled = True

optsaved = False
End If 'null ret
Set cd = Nothing
End Sub


' Set the list box's horizontal extent so it
' can display its longest entry. This routine
' assumes the form is using the same font as
' the list box.
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
Public Sub ComOptSave_Click()
Dim WFD As WIN32_FIND_DATA
Dim Ret As Long
Dim i As Integer

If Not INIFileFlagRW Then ComOptSave.Enabled = False: Exit Sub

Screen.MousePointer = vbHourglass
On Error Resume Next

' Write the new key to Global.ini
'check ini
iniFileName = App.Path
If Right$(iniFileName, 1) <> "\" Then iniFileName = iniFileName & "\"
iniFileName = iniFileName & "global.ini"
Ret = FindFirstFile(iniFileName, WFD)
If Ret < 0 Then FrmMain.MakeINI "global.ini"
FindClose Ret

WriteKey "GLOBAL", "LoadLastBD", CheckLoadLastBD.Value, iniFileName
WriteKey "Language", "LastLang", ComboLangHid, iniFileName

WriteKey "GLOBAL", "BDCount", LstBases.ListCount, iniFileName
For i = 0 To LstBases.ListCount - 1
WriteKey "GLOBAL", "BDname" & i + 1, LstBases.List(i), iniFileName
Next i


'WriteKey "GLOBAL", "LoadTech", CheckShowTech.Value, iniFileName
WriteKey "GLOBAL", "QJPG", TextQJPGHid, iniFileName
WriteKey "GLOBAL", "SaveBigPix", CheckSavBigPix.Value, iniFileName

' Write the new key to current ini
iniFileName = App.Path
If Right$(iniFileName, 1) <> "\" Then iniFileName = iniFileName & "\"
iniFileName = iniFileName & INIFILE
Ret = FindFirstFile(iniFileName, WFD)
If Ret < 0 Then FrmMain.MakeINI INIFILE
FindClose Ret

WriteKey "CD", "CDdrive", ComboCDHid, iniFileName
WriteKey "PRINTER", "CurrentPrinter", cboPrinterHid, iniFileName
' DS filter checkbox
WriteKey "GLOBAL", "FreeDVDFilters", ChDSFilt, iniFileName

'LV

For i = 1 To lvHeaderIndexPole
WriteKey "LIST", "C" & i, FrmMain.ListView.ColumnHeaders(i).Width, iniFileName
WriteKey "LIST", "P" & i, FrmMain.ListView.ColumnHeaders(i).Position, iniFileName
Next i

'Export
For i = 0 To LstExport.ListCount - 1
WriteKey "EXPORT", "L" & i, LstExport.Selected(i), iniFileName
Next
For i = OptHtml.LBound To OptHtml.UBound
DeleteKey "OptHtml" & i & ".Caption", "EXPORT", iniFileName
If OptHtml(i).Value = True Then
WriteKey "EXPORT", "OptHtml" & i & ".Caption", True, iniFileName
End If
Next


'Font
WriteKey "FONT", "VFontName", LblFontVHid.Font.name, iniFileName
WriteKey "FONT", "VFontSize", LblFontVHid.Font.Size, iniFileName
WriteKey "FONT", "VFontBold", LblFontVHid.Font.Bold, iniFileName
WriteKey "FONT", "VFontItalic", LblFontVHid.Font.Italic, iniFileName
WriteKey "FONT", "VFontColor", Str$(VFontColor), iniFileName

WriteKey "FONT", "HFontName", LblFontHHid.Font.name, iniFileName
WriteKey "FONT", "HFontSize", LblFontHHid.Font.Size, iniFileName
WriteKey "FONT", "HFontBold", LblFontHHid.Font.Bold, iniFileName
WriteKey "FONT", "HFontItalic", LblFontHHid.Font.Italic, iniFileName
WriteKey "FONT", "HFontColor", Str$(HFontColor), iniFileName

WriteKey "FONT", "LVFontName", LblFontLVHid.Font.name, iniFileName
WriteKey "FONT", "LVFontSize", LblFontLVHid.Font.Size, iniFileName
WriteKey "FONT", "LVFontBold", LblFontLVHid.Font.Bold, iniFileName
WriteKey "FONT", "LVFontItalic", LblFontLVHid.Font.Italic, iniFileName
WriteKey "FONT", "LVFontColor", Str$(LVFontColor), iniFileName

WriteKey "FONT", "LVBackColor", Str$(LVBackColor), iniFileName
WriteKey "FONT", "CoverHorBackColor", Str$(CoverHorBackColor), iniFileName
WriteKey "FONT", "CoverVertBackColor", Str$(CoverVertBackColor), iniFileName




WriteKey "GLOBAL", "LoadLastBD", CheckLoadLastBD.Value, iniFileName
'WriteKey "GLOBAL", "LastBD", bdname, iniFileName
'WriteKey "GLOBAL", "LoadTech", CheckShowTech.Value, iniFileName
WriteKey "GLOBAL", "QJPG", TextQJPGHid, iniFileName
WriteKey "GLOBAL", "SaveBigPix", CheckSavBigPix.Value, iniFileName
WriteKey "GLOBAL", "LVLoadOnlyTitle", ChOnlyTitle.Value, iniFileName

optsaved = True
Screen.MousePointer = vbNormal
Call err.Clear
End Sub


Private Sub ComUnmarkAll_Click()
Dim i As Integer
For i = 0 To OptDialog.LstExport.ListCount - 1
OptDialog.LstExport.Selected(i) = False
Next

End Sub


Private Sub Form_Load()
'OptDialFlag = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'OptDialFlag = False
End Sub

Private Sub LstBases_Click()
If Len(LstBases.Text) = 0 Then Exit Sub
'If FrmMain.FrameView.Visible Then If oldTabLVInd = LstBases.ListIndex + 1 Then Exit Sub ' не отрабатывать при кликах на табах LV
FrmMain.Timer2.Enabled = False

InitFlag = True
bdname = LstBases.Text
FrmMain.ReadINI GetNameFromPathAndName(bdname)
'LstExport.ListIndex = LstExport.ListCount - 1: LstExport.TopIndex = 0 'чтоб не видно выделения

LabelCurrBDHid.Caption = bdname
LBDSizeHid.Caption = FileLen(bdname) / 1024
ComCompact.Enabled = True

optsaved = True
End Sub

Private Sub LstExport_ItemCheck(Item As Integer)
optsaved = False
End Sub

Private Sub OptHtml_Click(Index As Integer)
optsaved = False
End Sub

Private Sub TextQJPGHid_LostFocus()
If (Val(TextQJPGHid) < 1) Or (Val(TextQJPGHid) > 100) Then TextQJPGHid = 80
QJPG = TextQJPGHid
optsaved = False
End Sub

