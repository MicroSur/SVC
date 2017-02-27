VERSION 5.00
Begin VB.Form frmSR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SR"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SurVideoCatalog.XpB comReplace 
      Height          =   375
      Index           =   1
      Left            =   5820
      TabIndex        =   12
      Top             =   2160
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      Caption         =   "Replace All"
      ButtonStyle     =   3
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB comReplace 
      Height          =   375
      Index           =   0
      Left            =   5820
      TabIndex        =   11
      Top             =   1680
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      Caption         =   "Replace"
      ButtonStyle     =   3
      Picture         =   "frmSR.frx":0000
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB comSRFind 
      Height          =   375
      Index           =   1
      Left            =   5820
      TabIndex        =   10
      Top             =   1020
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      Caption         =   "Find Next"
      ButtonStyle     =   3
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB comSRFind 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5820
      TabIndex        =   9
      Top             =   540
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      Caption         =   "Search"
      ButtonStyle     =   3
      Picture         =   "frmSR.frx":059A
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin VB.Frame frDelete 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   180
      TabIndex        =   28
      Top             =   420
      Width           =   7275
      Begin VB.OptionButton optDelWhat 
         Caption         =   "3"
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   47
         Top             =   660
         Width           =   1815
      End
      Begin VB.OptionButton optDelWhat 
         Caption         =   "2"
         Height          =   195
         Index           =   2
         Left            =   5400
         TabIndex        =   46
         Top             =   420
         Width           =   1815
      End
      Begin VB.OptionButton optDelWhat 
         Caption         =   "shot 1"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   45
         Top             =   180
         Width           =   1815
      End
      Begin VB.OptionButton optDelWhat 
         Caption         =   "cover"
         Height          =   195
         Index           =   4
         Left            =   5400
         TabIndex        =   44
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton optDelWhat 
         Height          =   315
         Index           =   0
         Left            =   4680
         TabIndex        =   43
         Top             =   120
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.ComboBox cDelFrom 
         Height          =   315
         Left            =   1320
         TabIndex        =   33
         Top             =   840
         Width           =   3795
      End
      Begin VB.TextBox txtDelStart 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3600
         TabIndex        =   32
         Text            =   "1"
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox txtDelLen 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3600
         TabIndex        =   30
         Text            =   "0"
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblDelFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   60
         TabIndex        =   34
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lblDelPos 
         Alignment       =   1  'Right Justify
         Caption         =   "DelStart"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   480
         Width           =   3435
      End
      Begin VB.Label lSRDel 
         Alignment       =   1  'Right Justify
         Caption         =   "DelLen"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   3315
      End
   End
   Begin VB.Frame frInsert 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   120
      TabIndex        =   20
      Top             =   420
      Width           =   7275
      Begin VB.TextBox txtIns 
         Height          =   285
         Left            =   2100
         TabIndex        =   26
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtIncr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5100
         TabIndex        =   25
         Text            =   "1"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtNumIns 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2100
         TabIndex        =   24
         Text            =   "001 "
         Top             =   120
         Width           =   1155
      End
      Begin VB.OptionButton optIns 
         Alignment       =   1  'Right Justify
         Caption         =   "text"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optIns 
         Alignment       =   1  'Right Justify
         Caption         =   "number"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox CInsWhere 
         Height          =   315
         Left            =   1380
         TabIndex        =   27
         Top             =   840
         Width           =   3795
      End
      Begin VB.Label lPlus 
         Alignment       =   1  'Right Justify
         Caption         =   "+ "
         Height          =   255
         Left            =   3420
         TabIndex        =   35
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label LInsWhere 
         Caption         =   "Where:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   900
         Width           =   1215
      End
   End
   Begin VB.Frame frConvert 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   36
      Top             =   480
      Width           =   7275
      Begin VB.OptionButton optChCase 
         Caption         =   "U Case First"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   42
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox txtDelim 
         Height          =   315
         Left            =   3360
         TabIndex        =   40
         Text            =   " ([""/'"
         Top             =   420
         Width           =   1695
      End
      Begin VB.OptionButton optChCase 
         Caption         =   "P Case"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   39
         Top             =   120
         Width           =   3495
      End
      Begin VB.OptionButton optChCase 
         Caption         =   "U Case"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   38
         Top             =   420
         Width           =   2835
      End
      Begin VB.OptionButton optChCase 
         Caption         =   "L Case"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   37
         Top             =   120
         Value           =   -1  'True
         Width           =   2835
      End
      Begin VB.Label lDelim 
         Caption         =   "Delim"
         Height          =   255
         Left            =   5220
         TabIndex        =   41
         Top             =   480
         Width           =   1995
      End
   End
   Begin VB.Frame FrSR 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   180
      TabIndex        =   14
      Top             =   420
      Width           =   5595
      Begin MSComctlLib.ImageList ImLstSR 
         Left            =   4620
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSR.frx":0B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSR.frx":10CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSR.frx":1668
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSR.frx":1C02
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSR.frx":2614
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox ChMultiSel 
         Caption         =   "Multiselect"
         Height          =   195
         Left            =   3240
         TabIndex        =   8
         Top             =   1860
         Width           =   2415
      End
      Begin VB.OptionButton OpbApplyTo 
         Caption         =   "Checked"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   6
         Top             =   2100
         Width           =   1875
      End
      Begin VB.CheckBox ChCaseSens 
         Caption         =   "CaseSens."
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   1620
         Width           =   2475
      End
      Begin VB.OptionButton OpbApplyTo 
         Caption         =   "Selected"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   1860
         Width           =   1815
      End
      Begin VB.OptionButton OpbApplyTo 
         Caption         =   "Visible"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   1620
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.ComboBox CReplStr 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   4155
      End
      Begin VB.ComboBox CWhere 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   3795
      End
      Begin VB.ComboBox CSearchIn 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   3795
      End
      Begin VB.ComboBox CSearchStr 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   120
         Width           =   4155
      End
      Begin VB.Label LLookIn 
         Caption         =   "Look In:"
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label LWhere 
         Caption         =   "Where:"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label LSearchIn 
         Caption         =   "Field"
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label LReplStr 
         Caption         =   "Replace"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label LSearchStr 
         Caption         =   "Search"
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip TabStripSR 
      Height          =   2835
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   5001
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      TabStyle        =   1
      ImageList       =   "ImLstSR"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Replace"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Insert"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Convert"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'- вызов только при видимом списке, закрывать если не так

Option Explicit

'Private ReplaceFlag As Boolean
Private TabInd As Integer

Private NStoreSR(0) As String 'от 0, доп. фразы


Private Sub ChMultiSel_Click()
If ChMultiSel.Value = vbChecked Then
OpbApplyTo(1).Enabled = True
Else
If OpbApplyTo(1).Value = True Then OpbApplyTo(0).Value = True
OpbApplyTo(1).Enabled = False
OpbApplyTo(1).Value = False
End If
End Sub



Private Sub comReplace_Click(Index As Integer)

Dim FindText As String
Dim ReplText As String

Dim ind As Integer    'поле базы db...ind
Dim lvInd As Long    'строка в lv
Dim CompareMethod As VbCompareMethod
Dim LookIn As LV_AllSelCheck
Dim SearchIn As AnyWholeFirst
Dim Itm As ListItem

Dim InsertIn As BeginEnd    'SearchIn for insert
Dim CountFrom As BeginEnd    'count from for delete
'LookIn As LV_AllSelCheck

Dim n As Long    'n = Val(txtIncr)
Dim X As Long    'x=x+n
Dim ins As String    'ins = sys_InsNumsIncr

Dim ConvertMethod As HowConvert

On Error GoTo err

If CSearchIn.ListIndex < 0 Then 'не выбрано поле
    If optDelWhat(0).Value = True Then 'и не выбрано удаление картинок
        Exit Sub
    End If
Else
    ind = CSearchIn.ItemData(CSearchIn.ListIndex) 'не для картинок
End If

If myMsgBox(msgsvc(50), vbOKCancel, , frmSR.hwnd) <> vbOK Then Exit Sub
DoEvents

FrmMain.Timer2.Enabled = False

'LookIn где в списке (для любого действия)
If OpbApplyTo(0).Value = True Then LookIn = AllLVRows
If OpbApplyTo(1).Value = True Then LookIn = SelectedLVRows
If OpbApplyTo(2).Value = True Then LookIn = CheckedLVRows

Select Case TabInd

Case 3  '                                         вставка

    'InsertIn куда вставлять в строке
    Select Case CInsWhere.ListIndex
    Case 0
        InsertIn = sBegin
    Case 1
        InsertIn = sEnd
    End Select

    If optIns(0).Value = True Then    '               число

        Screen.MousePointer = vbHourglass
        'LockWindowUpdate FrmMain.ListView.hWnd
        n = Int(Replace2Regional(txtIncr))
        ins = sys_InsNumsIncr(txtNumIns, X)
        For Each Itm In FrmMain.ListView.ListItems
            If InsertTextInDB(ind, Itm.Index, ins, LookIn, InsertIn) Then
                'увеличить добавку
                X = X + n
                ins = sys_InsNumsIncr(txtNumIns, X)
            End If
        Next
        'LockWindowUpdate 0
        Screen.MousePointer = vbNormal

    Else    '                                         текст
        If Len(txtIns) = 0 Then Exit Sub
    End If

    'вставить везде
    Screen.MousePointer = vbHourglass
    'LockWindowUpdate FrmMain.ListView.hWnd
    For Each Itm In FrmMain.ListView.ListItems
        InsertTextInDB ind, Itm.Index, txtIns, LookIn, InsertIn
    Next
    'LockWindowUpdate 0
    Screen.MousePointer = vbNormal

Case 4    '                                         удаление

If optDelWhat(0).Value = True Then
    'проверить txtDelLen, txtDelStart
    If Not IsNumeric(txtDelLen) Then Exit Sub
    If Not IsNumeric(txtDelStart) Then Exit Sub

    'от чего считать
    Select Case cDelFrom.ListIndex
    Case 0
        CountFrom = sBegin
    Case 1
        CountFrom = sEnd
    End Select

    Screen.MousePointer = vbHourglass
    For Each Itm In FrmMain.ListView.ListItems
        DeleteInDB ind, Itm.Index, txtDelLen, txtDelStart, LookIn, CountFrom
    Next
    Screen.MousePointer = vbNormal
ElseIf optDelWhat(1).Value = True Then 'удалить кадр 1
    Screen.MousePointer = vbHourglass
    For Each Itm In FrmMain.ListView.ListItems
        DeletePixInDB dbSnapShot1Ind, Itm.Index, LookIn
    Next
    Screen.MousePointer = vbNormal
ElseIf optDelWhat(2).Value = True Then 'удалить кадр 2
    Screen.MousePointer = vbHourglass
    For Each Itm In FrmMain.ListView.ListItems
        DeletePixInDB dbSnapShot2Ind, Itm.Index, LookIn
    Next
    Screen.MousePointer = vbNormal
ElseIf optDelWhat(3).Value = True Then 'удалить кадр 3
    Screen.MousePointer = vbHourglass
    For Each Itm In FrmMain.ListView.ListItems
        DeletePixInDB dbSnapShot3Ind, Itm.Index, LookIn
    Next
    Screen.MousePointer = vbNormal
ElseIf optDelWhat(4).Value = True Then 'удалить cover
    Screen.MousePointer = vbHourglass
    For Each Itm In FrmMain.ListView.ListItems
        DeletePixInDB dbFrontFaceInd, Itm.Index, LookIn
    Next
    Screen.MousePointer = vbNormal
End If


Case 5 '                                           Изменение строки

For n = 0 To 3
If optChCase(n).Value = True Then ConvertMethod = n
Next n

    Screen.MousePointer = vbHourglass
    'LockWindowUpdate FrmMain.ListView.hWnd
    For Each Itm In FrmMain.ListView.ListItems
        ConvertInDB ind, Itm.Index, ConvertMethod, txtDelim.Text, LookIn
    Next
    'LockWindowUpdate 0
    Screen.MousePointer = vbNormal

Case 2    '                                        замена в строке

    FindText = CSearchStr.Text    'что ищем
    ReplText = CReplStr.Text    ' на что менять

    'чувствительность к регистру
    If ChCaseSens.Value = 0 Then
        CompareMethod = vbTextCompare
    ElseIf ChCaseSens.Value = 1 Then
        CompareMethod = vbBinaryCompare
    End If

    'SearchIn где в строке
    Select Case CWhere.ListIndex
    Case 0
        SearchIn = Search_Anywhere
    Case 1
        SearchIn = Search_WholeField
    Case 2
        SearchIn = Search_StartWith
    Case 3
        SearchIn = Search_EndWith
    Case 4
        SearchIn = Search_Shablon
    End Select

    Select Case Index
    Case 0    'заменить текущую
        lvInd = FrmMain.ListView.SelectedItem.Index
        ReplaceInDB ind, lvInd, FindText, ReplText, CompareMethod, LookIn, SearchIn

    Case 1    'заменить все
        Screen.MousePointer = vbHourglass
        'LockWindowUpdate FrmMain.ListView.hWnd
        For Each Itm In FrmMain.ListView.ListItems
            ReplaceInDB ind, Itm.Index, FindText, ReplText, CompareMethod, LookIn, SearchIn
        Next
        'LockWindowUpdate 0
        Screen.MousePointer = vbNormal
    End Select

    'запомнить вписанное
    FilterAddTypedItems

End Select

'RestoreBasePos
'If ind = lvIndexPole Then
FrmMain.LVCLICK 'для описания и карточки

If FrmMain.FrameView.Visible Then FrmMain.Timer2.Enabled = True
Exit Sub

err:
'LockWindowUpdate 0
Screen.MousePointer = vbNormal
ToDebug "err_repl: " & err.Description
End Sub


Private Sub comSRFind_Click(Index As Integer)
Dim FindText As String
Dim ind As Integer
Dim CompareMethod As VbCompareMethod
Dim LookIn As LV_AllSelCheck
Dim SearchIn As AnyWholeFirst

FrmMain.Timer2.Enabled = False

If CSearchIn.ListIndex < 0 Then Exit Sub

FindText = CSearchStr.Text

If ChMultiSel.Value = vbUnchecked Then FrmMain.ListView.MultiSelect = False

ind = CSearchIn.ItemData(CSearchIn.ListIndex)

'чувствительность к регистру
If ChCaseSens.Value = 0 Then
    CompareMethod = vbTextCompare
ElseIf ChCaseSens.Value = 1 Then
    CompareMethod = vbBinaryCompare
End If

'где в списке
If OpbApplyTo(0).Value = True Then LookIn = AllLVRows
If OpbApplyTo(1).Value = True Then LookIn = SelectedLVRows
If OpbApplyTo(2).Value = True Then LookIn = CheckedLVRows

'где в строке
Select Case CWhere.ListIndex
Case 0
    SearchIn = Search_Anywhere
Case 1
    SearchIn = Search_WholeField
Case 2
    SearchIn = Search_StartWith
Case 3
    SearchIn = Search_EndWith
Case 4
    SearchIn = Search_Shablon
End Select

'на первую строку
'Set FrmMain.ListView.SelectedItem = FrmMain.ListView.ListItems.Item(1)

If Index = 0 Then 'Find
SearchNextDB ind, FindText, 1, True, CompareMethod, LookIn, SearchIn
Else 'Find Next
SearchNextDB ind, FindText, FrmMain.ListView.SelectedItem.Index, False, CompareMethod, LookIn, SearchIn
End If

RestoreBasePos
'FrmMain.LVCLICK

FrmMain.ListView.MultiSelect = True
If FrmMain.FrameView.Visible Then FrmMain.Timer2.Enabled = True

FilterAddTypedItems 'запомнить вписанное
End Sub

Private Sub Form_Activate()
If BaseReadOnly Or BaseReadOnlyU Or (FrmMain.ListView.ListItems.Count < 1) Then
TabStripSR.Enabled = False
TabStripSR.Tabs(1).Selected = True

Else
TabStripSR.Enabled = True
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 112 'F1
    If FrmMain.ChBTT.Value = 0 Then FrmMain.ChBTT.Value = 1 Else FrmMain.ChBTT.Value = 0
    Me.SetFocus
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
frmSRFlag = True
TabStripSR_Click

CWhere.Clear
CWhere.AddItem "Any": CWhere.AddItem "Whole"
CWhere.AddItem "Start": CWhere.AddItem "End"
CWhere.AddItem "Wildcards"

CInsWhere.Clear
CInsWhere.AddItem "begin": CInsWhere.AddItem "end"

cDelFrom.Clear
cDelFrom.AddItem "begin": cDelFrom.AddItem "end"

GetLangRS

ForceTextBoxNumeric txtIncr, True
ForceTextBoxNumeric txtDelLen, True
ForceTextBoxNumeric txtDelStart, True

FillComboWithSvcFieldsNames

ChMultiSel_Click
CSearchStr.SetFocus

End Sub
Private Sub FillComboWithSvcFieldsNames()
Dim i As Integer

CSearchIn.Clear
'CSearchIn.AddItem "All", 0
'CSearchIn.ItemData(0) = -1
For i = 0 To UBound(TranslatedFieldsNames)
CSearchIn.AddItem TranslatedFieldsNames(i), i
CSearchIn.ItemData(i) = i
Next i
'CSearchIn.ListIndex = 0


End Sub
Private Sub Form_Resize()

frInsert.Visible = False
frDelete.Visible = False
frConvert.Visible = False
LReplStr.Visible = False
CReplStr.Visible = False
comReplace(0).Visible = False
comReplace(1).Visible = False

Select Case TabInd
Case 1 'search

Case 2 'replace
frInsert.Visible = False
frDelete.Visible = False
LReplStr.Visible = True
CReplStr.Visible = True
comReplace(0).Visible = True
comReplace(1).Visible = True

Case 3 'insert
frInsert.ZOrder 0
frInsert.Visible = True
'frDelete.Visible = False
'comReplace(0).Visible = False
comReplace(1).Visible = True

Case 4 'delete
frInsert.ZOrder 0: frDelete.ZOrder 0
frInsert.Visible = True
frDelete.Visible = True
'comReplace(0).Visible = False
comReplace(1).Visible = True

Case 5 'convert
frConvert.ZOrder 0
frConvert.Visible = True
comReplace(1).Visible = True

End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

'в целях безопасности
CSearchIn = vbNullString
optDelWhat(0).Value = True
''''

If ExitSVC Then
    frmSRFlag = False
Else
    frmSR.Hide
    Cancel = True
End If
End Sub

Private Sub TabStripSR_Click()
On Error Resume Next
TabInd = TabStripSR.SelectedItem.Index

Form_Resize
CSearchStr.SetFocus

End Sub
Private Sub GetLangRS()
Dim Contrl As Control
Dim i As Integer
'Dim temp As String

On Error Resume Next
DoEvents

If Dir(lngFileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = vbNullString Or Len(lngFileName) = 0 Then
    Call myMsgBox("Не найден файл локализации! Исправьте параметр LastLang в global.ini" & vbCrLf & "Language file not found: " & vbCrLf & lngFileName, vbCritical, , Me.hwnd)
    Exit Sub
End If

For Each Contrl In frmSR.Controls

    If TypeOf Contrl Is TabStrip Then    '                        TabStrip
        Contrl.Tabs(1).Caption = ReadLangSR(Contrl.name & ".Tabs(1).Caption", Contrl.Tabs(1).Caption)
        Contrl.Tabs(2).Caption = ReadLangSR(Contrl.name & ".Tabs(2).Caption", Contrl.Tabs(2).Caption)
        Contrl.Tabs(3).Caption = ReadLangSR(Contrl.name & ".Tabs(3).Caption", Contrl.Tabs(3).Caption)
        Contrl.Tabs(4).Caption = ReadLangSR(Contrl.name & ".Tabs(4).Caption", Contrl.Tabs(4).Caption)
        Contrl.Tabs(5).Caption = ReadLangSR(Contrl.name & ".Tabs(5).Caption", Contrl.Tabs(5).Caption)
    End If    '(TypeOf Contrl


    If TypeOf Contrl Is Label Then    '                           Label
        Contrl.Caption = ReadLangSR(Contrl.name & ".Caption")
    End If

    If TypeOf Contrl Is XpB Then    '                              XPB
        If (Contrl.name = "comSRFind") Or (Contrl.name = "comReplace") Then
            Contrl.Caption = ReadLangSR(Contrl.name & Contrl.Index & ".Caption", Contrl.Caption)
        Else
            Contrl.Caption = ReadLangSR(Contrl.name & ".Caption", Contrl.Caption)
        End If
        Contrl.pInitialize
    End If

    'If TypeOf Contrl Is Frame Then '                           Frame
    'Contrl.Caption = ReadLangSR(Contrl.name & ".Caption", Contrl.Caption)
    'End If

    If TypeOf Contrl Is OptionButton Then    '                     OptionButton
        If Contrl.name = "OpbApplyTo" Then
            For i = 0 To 2
                OpbApplyTo(i).Caption = ReadLangSR(Contrl.name & i & ".Caption", OpbApplyTo(i).Caption)
            Next
        ElseIf Contrl.name = "optIns" Then
            For i = 0 To 1
                optIns(i).Caption = ReadLangSR(Contrl.name & i & ".Caption", optIns(i).Caption)
            Next
        ElseIf Contrl.name = "optChCase" Then
            For i = 0 To 3
                optChCase(i).Caption = ReadLangSR(Contrl.name & i & ".Caption", optChCase(i).Caption)
            Next
        ElseIf Contrl.name = "optDelWhat" Then
            For i = 1 To 4
                optDelWhat(i).Caption = ReadLangSR(Contrl.name & i & ".Caption", optDelWhat(i).Caption)
            Next
        End If
    End If

    If TypeOf Contrl Is CheckBox Then    '                         CheckBox
        Contrl.Caption = ReadLangSR(Contrl.name & ".Caption")
    End If


    '                                                           Lst
    If TypeOf Contrl Is ComboBox Then
        If Contrl.name = "CWhere" Then
            For i = 0 To CWhere.ListCount - 1
                CWhere.List(i) = ReadLangSR("CWhere" & i, CWhere.List(i))
            Next i
            CWhere.ListIndex = 0

        ElseIf Contrl.name = "CInsWhere" Then
            For i = 0 To CInsWhere.ListCount - 1
                CInsWhere.List(i) = ReadLangSR("CInsWhere" & i, CInsWhere.List(i))
            Next i
        CInsWhere.ListIndex = 0

        ElseIf Contrl.name = "cDelFrom" Then
            For i = 0 To cDelFrom.ListCount - 1
             cDelFrom.List(i) = ReadLangSR("cDelFrom" & i, cDelFrom.List(i))
          Next i
        cDelFrom.ListIndex = 0
        End If
    End If

Next    'Contrl

'                                                       NamesStoreSR()
For i = 0 To UBound(NStoreSR)
    NStoreSR(i) = ReadLangSR("NStoreSR" & i)
Next i



Me.Caption = "SurVideoCatalog - " & NStoreSR(0)
Me.Icon = FrmMain.Icon
End Sub
Private Sub FilterAddTypedItems()
'вбить в комбики введенное пользователем
On Error Resume Next
If Len(CSearchStr.Text) <> 0 Then
   If SearchCBO(CSearchStr, CSearchStr.Text, False) < 0 Then
    SendMessage CSearchStr.hwnd, CB_ADDSTRING, 0, ByVal CSearchStr.Text
   End If
End If

If Len(CReplStr.Text) <> 0 Then
   If SearchCBO(CReplStr, CReplStr.Text, False) < 0 Then
    SendMessage CReplStr.hwnd, CB_ADDSTRING, 0, ByVal CReplStr.Text
   End If
End If

End Sub

