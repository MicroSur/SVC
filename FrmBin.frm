VERSION 5.00
Begin VB.Form FrmBin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SVCatalog: DragDrop"
   ClientHeight    =   4995
   ClientLeft      =   2910
   ClientTop       =   2145
   ClientWidth     =   3555
   Icon            =   "FrmBin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chAll 
      Height          =   195
      Left            =   3300
      TabIndex        =   25
      ToolTipText     =   "All"
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   10
      Left            =   3300
      TabIndex        =   24
      ToolTipText     =   "Picture"
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   9
      Left            =   3300
      TabIndex        =   23
      ToolTipText     =   "Annotation"
      Top             =   2820
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   8
      Left            =   3300
      TabIndex        =   22
      Top             =   2520
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   7
      Left            =   3300
      TabIndex        =   21
      Top             =   2220
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   6
      Left            =   3300
      TabIndex        =   20
      Top             =   1920
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   5
      Left            =   3300
      TabIndex        =   19
      Top             =   1620
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   4
      Left            =   3300
      TabIndex        =   18
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   3
      Left            =   3300
      TabIndex        =   15
      Top             =   420
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   2
      Left            =   3300
      TabIndex        =   17
      Top             =   1020
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   1
      Left            =   3300
      TabIndex        =   16
      Top             =   720
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox ch 
      Height          =   195
      Index           =   0
      Left            =   3300
      TabIndex        =   14
      Top             =   120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox TextRate 
      Height          =   285
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   2460
      Width           =   2295
   End
   Begin VB.TextBox TextSubt 
      Height          =   285
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox TextLang 
      Height          =   285
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   1860
      Width           =   2295
   End
   Begin VB.PictureBox PicFrontFace 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   915
      Left            =   1800
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   855
      ScaleWidth      =   1395
      TabIndex        =   33
      Top             =   2820
      Width           =   1455
      Begin VB.CommandButton ComFrontFace 
         Height          =   315
         Index           =   0
         Left            =   1080
         MousePointer    =   1  'Arrow
         Picture         =   "FrmBin.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
   End
   Begin SurVideoCatalog.XpB ComClearBin 
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   4620
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   556
      Caption         =   "clear"
      ButtonStyle     =   3
      Picture         =   "FrmBin.frx":0596
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB ComPastFromBin 
      Height          =   375
      Left            =   60
      TabIndex        =   11
      Top             =   3780
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   661
      Caption         =   "Out"
      ButtonStyle     =   3
      Picture         =   "FrmBin.frx":0B30
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin VB.TextBox TextCountry 
      Height          =   285
      Left            =   960
      MaxLength       =   50
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox TextMName 
      Height          =   285
      Left            =   960
      MaxLength       =   255
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   60
      Width           =   2295
   End
   Begin VB.TextBox TextAnnotation 
      Height          =   675
      Left            =   60
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3060
      Width           =   1725
   End
   Begin VB.TextBox TextAuthor 
      Height          =   285
      Left            =   960
      MaxLength       =   100
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   1260
      Width           =   2295
   End
   Begin VB.TextBox TextRole 
      Height          =   285
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox TextYear 
      Height          =   285
      Left            =   960
      MaxLength       =   20
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox TextGenre 
      Height          =   285
      Left            =   960
      MaxLength       =   100
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   660
      Width           =   2295
   End
   Begin SurVideoCatalog.XpB comPastToBin 
      Height          =   375
      Left            =   60
      TabIndex        =   12
      Top             =   4200
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   661
      Caption         =   "In"
      ButtonStyle     =   3
      Picture         =   "FrmBin.frx":10CA
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   36
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Subt"
      Height          =   255
      Index           =   11
      Left            =   60
      TabIndex        =   35
      Top             =   2220
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Lang"
      Height          =   255
      Index           =   10
      Left            =   60
      TabIndex        =   34
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Descr"
      Height          =   255
      Index           =   9
      Left            =   60
      TabIndex        =   32
      Top             =   2820
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Cat"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   31
      Top             =   720
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   30
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Dir"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   29
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Prod"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   28
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Act"
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   27
      Top             =   1620
      Width           =   855
   End
   Begin VB.Label LFilm 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   255
      Index           =   6
      Left            =   60
      TabIndex        =   26
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "FrmBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'mzt Private Declare Sub ReleaseCapture Lib "user32" ()
'mzt Private Const WM_NCLBUTTONDOWN = &HA1
'mzt Private Const HTCAPTION = 2
Private IsTyping As Boolean 'ввод с клавы


Private Sub chAll_Click()
Dim i As Integer
'chAll.Value = Not chAll.Value

Select Case chAll.Value
Case vbChecked
For i = 0 To 10: ch(i).Value = vbChecked: Next i
Case vbUnchecked
For i = 0 To 10: ch(i).Value = vbUnchecked: Next i
End Select
End Sub

Private Sub ComClearBin_Click()
Dim Contrl As Control
For Each Contrl In Me.Controls
 If TypeOf Contrl Is TextBox Then Contrl.Text = vbNullString
Next
Set PicFrontFace = Nothing
End Sub

Private Sub ComFrontFace_Click(Index As Integer)
If Clipboard.GetFormat(vbCFDIB) Then PicFrontFace.Picture = Clipboard.GetData
End Sub

Private Sub ComPastFromBin_Click()
Dim Contrl As Control

If Not frmEditorFlag Then Exit Sub

'почистить
TextMName.Text = sTrimChars(TextMName.Text, vbNewLine)
TextGenre.Text = sTrimChars(TextGenre.Text, vbNewLine)
TextCountry.Text = sTrimChars(TextCountry.Text, vbNewLine)
TextYear.Text = sTrimChars(TextYear.Text, vbNewLine)
TextAuthor.Text = sTrimChars(TextAuthor.Text, vbNewLine)
TextRole.Text = sTrimChars(TextRole.Text, vbNewLine)
TextLang.Text = sTrimChars(TextLang.Text, vbNewLine)
TextSubt.Text = sTrimChars(TextSubt.Text, vbNewLine)
TextRate.Text = sTrimChars(TextRate.Text, vbNewLine)

'вписать в редактор
With frmEditor
'бывшие в редакторе данные не сохраняются

For Each Contrl In FrmBin.Controls
If TypeOf Contrl Is TextBox Then
 If Contrl.Text <> vbNullString Then
 'Первая буква > заглавная
 'Contrl.Text = StrConv(Contrl.Text, vbProperCase, LCID)
 
  Select Case Contrl.name
  Case "TextMName": If ch(0).Value = vbChecked Then .TextMName.Text = Contrl.Text
  Case "TextGenre": If ch(1).Value = vbChecked Then .TextGenre.Text = Contrl.Text
  Case "TextCountry": If ch(2).Value = vbChecked Then .TextCountry.Text = Contrl.Text
  Case "TextYear": If ch(3).Value = vbChecked Then .TextYear.Text = Contrl.Text
  Case "TextAuthor": If ch(4).Value = vbChecked Then .TextAuthor.Text = Contrl.Text
  Case "TextRole": If ch(5).Value = vbChecked Then .TextRole.Text = Contrl.Text
   Case "TextLang": If ch(6).Value = vbChecked Then .TextLang.Text = Contrl.Text
   Case "TextSubt": If ch(7).Value = vbChecked Then .TextSubt.Text = Contrl.Text
   Case "TextRate": If ch(8).Value = vbChecked Then .TextRate.Text = Contrl.Text

  Case "TextAnnotation": If ch(9).Value = vbChecked Then .TextAnnotation.Text = Contrl.Text
  End Select
 
 End If
End If
Next

If ch(10).Value = vbChecked Then
If PicFrontFace.Picture <> 0 Then
    Set .PicFrontFace = Nothing
    Set .picCanvas = Nothing
    NoPicFrontFaceFlag = False
    DrDroFlag = True
    .PicFrontFace.Picture = PicFrontFace.Picture
    .ImgPrCov.Picture = .PicFrontFace.Picture
    DrawCoverEdit
    
    SaveCoverFlag = True
End If
End If

'On Error Resume Next
'.WindowState = FrmMainState
End With


End Sub

Private Sub comPastToBin_Click()
Dim Contrl As Control

If Not frmEditorFlag Then Exit Sub

'вписать из редактора
With frmEditor
'бывшие данные не сохраняются
For Each Contrl In FrmBin.Controls
If TypeOf Contrl Is TextBox Then
' If Contrl.Text <> vbNullString Then
  Select Case Contrl.name
  Case "TextMName": If ch(0).Value = vbChecked Then Contrl.Text = .TextMName.Text
  Case "TextGenre": If ch(1).Value = vbChecked Then Contrl.Text = .TextGenre.Text
  Case "TextCountry": If ch(2).Value = vbChecked Then Contrl.Text = .TextCountry.Text
  Case "TextYear": If ch(3).Value = vbChecked Then Contrl.Text = .TextYear.Text
  Case "TextAuthor": If ch(4).Value = vbChecked Then Contrl.Text = .TextAuthor.Text
  Case "TextRole": If ch(5).Value = vbChecked Then Contrl.Text = .TextRole.Text
   Case "TextLang": If ch(6).Value = vbChecked Then Contrl.Text = .TextLang.Text
   Case "TextSubt": If ch(7).Value = vbChecked Then Contrl.Text = .TextSubt.Text
   Case "TextRate": If ch(8).Value = vbChecked Then Contrl.Text = .TextRate.Text

  Case "TextAnnotation":  If ch(9).Value = vbChecked Then Contrl.Text = .TextAnnotation.Text
  End Select
 
' End If
End If
Next

If ch(10).Value = vbChecked Then
Set PicFrontFace = Nothing
If .PicFrontFace.Picture <> 0 Then PicFrontFace.Picture = .PicFrontFace.Picture
End If

'On Error Resume Next
'.WindowState = FrmMainState
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim Contrl As Control

Me.Top = Screen.Height / 2 - Height / 2
Me.Left = Screen.Width - Me.Width - 400

For Each Contrl In FrmBin.Controls
    If TypeOf Contrl Is TextBox Then
        Contrl.Font.Charset = 204
    End If
Next

For Each Contrl In FrmBin.Controls

    If TypeOf Contrl Is Label Then        '                           Label
        If Contrl.name = "LTech" Then
        '    LTech(Contrl.Index).Caption = ReadLang("LTech(" & Contrl.Index & ").Caption", LTech(Contrl.Index).Caption)
        ElseIf Contrl.name = "LFilm" Then
            LFilm(Contrl.Index).Caption = ReadLang("LFilm(" & Contrl.Index & ").Caption", LFilm(Contrl.Index).Caption)
        Else
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
        End If
    End If

    If TypeOf Contrl Is XpB Then    '                    XPB
        Contrl.Caption = ReadLang(Contrl.name & ".Caption")
        '  Contrl.ToolTipText = ReadLang(Contrl.name & ".ToolTip")
        Contrl.pInitialize
    End If
Next

PicFrontFace.Print NamesStore(0)

Me.Icon = FrmMain.Icon

MakeTopMost FrmBin.hwnd
frmBinFlag = True
End Sub

Private Sub Form_Resize()
'Background
If lngBrush <> 0 Then
GetClientRect hwnd, rctMain
FillRect hdc, rctMain, lngBrush
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmBinFlag = False
End Sub

Private Sub PicFrontFace_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
PicFrontFace.Picture = Data.GetData(vbCFDIB)
End Sub

Private Sub PicFrontFace_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
' A drop is OK only if bitmap data is available.
If Data.GetFormat(vbCFBitmap) Or Data.GetFormat(vbCFDIB) Then
 Effect = vbDropEffectCopy
Else
 Effect = vbDropEffectNone
End If
End Sub

Private Sub TextAnnotation_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
If Data.GetFormat(1) Then
 If TextAnnotation.Text <> vbNullString Then plus = TextAnnotation.Text & vbNewLine & vbNewLine
 TextAnnotation.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextAnnotation_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
End Sub

Private Sub TextAuthor_Change()
If Not IsTyping Then TextAuthor.Text = sTrimChars(TextAuthor.Text, vbNewLine)
End Sub

Private Sub TextAuthor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextAuthor_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
IsTyping = False
TextAuthor.Text = sTrimChars(TextAuthor.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextAuthor.Text <> vbNullString Then plus = TextAuthor.Text & ", "
 TextAuthor.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextAuthor_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
End Sub

Private Sub TextCountry_Change()
If Not IsTyping Then TextCountry.Text = sTrimChars(TextCountry.Text, vbNewLine)
End Sub

Private Sub TextCountry_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextCountry_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
IsTyping = False
TextCountry.Text = sTrimChars(TextCountry.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextCountry.Text <> vbNullString Then plus = TextCountry.Text & ", "
 TextCountry.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextCountry_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
End Sub

Private Sub TextGenre_Change()
If Not IsTyping Then TextGenre.Text = sTrimChars(TextGenre.Text, vbNewLine)
End Sub

Private Sub TextGenre_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextGenre_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
Dim tmp As String

IsTyping = False
TextGenre.Text = sTrimChars(TextGenre.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextGenre.Text <> vbNullString Then plus = TextGenre.Text & ", "
 tmp = Replace(Data.GetData(1), " /", ",")
 TextGenre.Text = plus & tmp
End If
End Sub

Private Sub TextGenre_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
End Sub

Private Sub TextLang_Change()
If Not IsTyping Then TextLang.Text = sTrimChars(TextLang.Text, vbNewLine)
End Sub

Private Sub TextLang_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextLang_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
IsTyping = False
TextLang.Text = sTrimChars(TextLang.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextLang.Text <> vbNullString Then plus = TextLang.Text & ", "
 TextLang.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextLang_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone

End Sub

Private Sub TextMName_Change()
If Not IsTyping Then TextMName.Text = sTrimChars(TextMName.Text, vbNewLine)
End Sub

Private Sub TextMName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextMName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
IsTyping = False
TextMName.Text = sTrimChars(TextMName.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextMName.Text <> vbNullString Then plus = TextMName.Text & ", "
 TextMName.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextMName_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
End Sub

Private Sub TextRate_Change()
If Not IsTyping Then TextRate.Text = sTrimChars(TextRate.Text, vbNewLine)
End Sub

Private Sub TextRate_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextRate_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
IsTyping = False
TextRate.Text = sTrimChars(TextRate.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextRate.Text <> vbNullString Then plus = TextRate.Text & ", "
 TextRate.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextRate_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
End Sub

Private Sub TextRole_Change()
If Not IsTyping Then TextRole.Text = sTrimChars(TextRole.Text, vbNewLine)
End Sub

Private Sub TextRole_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextRole_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
IsTyping = False
TextRole.Text = sTrimChars(TextRole.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextRole.Text <> vbNullString Then plus = TextRole.Text & ", "
 TextRole.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextRole_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
End Sub

Private Sub TextSubt_Change()
If Not IsTyping Then TextSubt.Text = sTrimChars(TextSubt.Text, vbNewLine)
End Sub

Private Sub TextSubt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextSubt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
IsTyping = False
TextSubt.Text = sTrimChars(TextSubt.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextSubt.Text <> vbNullString Then plus = TextSubt.Text & ", "
 TextSubt.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextSubt_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone

End Sub

Private Sub TextYear_Change()
If Not IsTyping Then TextYear.Text = sTrimChars(TextYear.Text, vbNewLine)
End Sub

Private Sub TextYear_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
IsTyping = True
End Sub

Private Sub TextYear_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim plus As String
IsTyping = False
TextYear.Text = sTrimChars(TextYear.Text, vbNewLine)
If Data.GetFormat(1) Then
 If TextYear.Text <> vbNullString Then plus = TextYear.Text & ", "
 TextYear.Text = plus & Data.GetData(1)
End If
End Sub

Private Sub TextYear_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
End Sub
