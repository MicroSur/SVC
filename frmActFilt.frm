VERSION 5.00
Begin VB.Form frmActFilt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chAF 
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   13
      Top             =   960
      Width           =   195
   End
   Begin VB.CheckBox chAF 
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   12
      Top             =   720
      Width           =   195
   End
   Begin VB.CheckBox chAD 
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   2100
      Width           =   195
   End
   Begin VB.CheckBox chAF 
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Top             =   1500
      Width           =   195
   End
   Begin VB.CheckBox chAF 
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   5
      Top             =   1260
      Width           =   195
   End
   Begin VB.CheckBox chAF 
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   420
      Width           =   195
   End
   Begin VB.CheckBox chAF 
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   195
   End
   Begin SurVideoCatalog.XpB cAFApply 
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2580
      Width           =   1635
      _extentx        =   2884
      _extenty        =   661
      showfocusrect   =   0
      caption         =   "Ok"
      buttonstyle     =   3
      picturewidth    =   0
      pictureheight   =   0
      xpcolor_pressed =   15116940
      xpcolor_hover   =   4692449
   End
   Begin SurVideoCatalog.XpB cAFCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2580
      Width           =   1635
      _extentx        =   2884
      _extenty        =   661
      showfocusrect   =   0
      caption         =   "No"
      buttonstyle     =   3
      picturewidth    =   0
      pictureheight   =   0
      xpcolor_pressed =   15116940
      xpcolor_hover   =   4692449
   End
   Begin VB.Label lNoBio 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lWBio 
      BackStyle       =   0  'Transparent
      Caption         =   "+Bio"
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   720
      Width           =   3015
   End
   Begin VB.Shape sh1 
      BorderColor     =   &H8000000C&
      Height          =   435
      Index           =   1
      Left            =   60
      Top             =   1980
      Width           =   3555
   End
   Begin VB.Shape sh1 
      BorderColor     =   &H8000000C&
      Height          =   1815
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   3555
   End
   Begin VB.Label lDups 
      BackStyle       =   0  'Transparent
      Caption         =   "Dups"
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   2100
      Width           =   3015
   End
   Begin VB.Label lLat 
      BackStyle       =   0  'Transparent
      Caption         =   "Lat"
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   1500
      Width           =   3015
   End
   Begin VB.Label lRus 
      BackStyle       =   0  'Transparent
      Caption         =   "Rus"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   1260
      Width           =   3015
   End
   Begin VB.Label lNoFoto 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   420
      Width           =   3015
   End
   Begin VB.Label lWFoto 
      BackStyle       =   0  'Transparent
      Caption         =   "+ Foto"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   180
      Width           =   3015
   End
End
Attribute VB_Name = "frmActFilt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ManualClickFlag As Boolean 'мануальный клик на галочку

Private Sub cAFApply_Click()
'подготовить запрос по сумме требований
Dim strSQL As String
'Dim tmp As String

On Error GoTo err

strSQL = "Select * From Acter Where "

If chAF(0).Value = vbChecked Then '+ Foto
    strSQL = strSQL & "(Not (Face Is Null)) And "
    ToDebug " Перс с фото"
    FilterActFlag = True
End If
If chAF(1).Value = vbChecked Then '- Foto
    strSQL = strSQL & "(Face Is Null) And "
    ToDebug " Перс без фото"
    FilterActFlag = True
End If


If chAF(4).Value = vbChecked Then '+ BIO
    strSQL = strSQL & "(Not (BIO Is Null)) And "
    ToDebug " Перс с БИО"
    FilterActFlag = True
End If
If chAF(5).Value = vbChecked Then '- BIO
    strSQL = strSQL & "(BIO Is Null) And "
    ToDebug " Перс без БИО"
    FilterActFlag = True
End If


If chAF(2).Value = vbChecked Then 'Rus
    strSQL = strSQL & "(Name not Like '*[a-Z]*') And "
    ToDebug " Перс rus"
    FilterActFlag = True
End If

If chAF(3).Value = vbChecked Then 'lat
    strSQL = strSQL & "(Name not Like '*[а-Я]*') And "
    ToDebug " Перс lat"
    FilterActFlag = True
End If

''''
If chAD(0).Value = vbChecked Then 'дуп
    'Set ars = ADB.OpenRecordset("Select * From Acter Where Name In (Select Name From Acter Group By Name HAVING Count(Name) > 1)")
    strSQL = strSQL & "(Name In (Select Name From Acter Group By Name HAVING Count(Name) > 1)) And "
    ToDebug " Актеры dups"
    FilterActFlag = True
End If

'убрать последний and
If Right$(strSQL, 5) = " And " Then strSQL = Left$(strSQL, Len(strSQL) - 5)

'Debug.Print strSQL
Set ars = ADB.OpenRecordset(strSQL)

ArsProcess
Exit Sub

err:

End Sub

Private Sub cAFCancel_Click()
subActFiltCancel
End Sub


Private Sub chAD_Click(Index As Integer)
Dim i As Integer
If ManualClickFlag Then
    ManualClickFlag = False
    For i = 0 To chAF.Count - 1
        chAF(i).Value = vbUnchecked
    Next i
End If
End Sub

Private Sub chAD_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
ManualClickFlag = True
End Sub

Private Sub chAD_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ManualClickFlag = True
End Sub

Private Sub chAF_Click(Index As Integer)
'типа это парные опции ' 0-1 4-5 2-3
If ManualClickFlag Then
ManualClickFlag = False

chAD(0).Value = vbUnchecked

Select Case Index
Case 0
    chAF(1).Value = vbUnchecked
Case 1
    chAF(0).Value = vbUnchecked
Case 2
    chAF(3).Value = vbUnchecked
Case 3
    chAF(2).Value = vbUnchecked
Case 4
    chAF(5).Value = vbUnchecked
Case 5
    chAF(4).Value = vbUnchecked

End Select
End If
End Sub

Private Sub chAF_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
ManualClickFlag = True
End Sub

Private Sub chAF_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ManualClickFlag = True
End Sub

Private Sub Form_Load()
Dim Contrl As Control
Dim i As Integer

frmActFilt.Caption = ReadLangActFilt("ActFilt.Caption", frmActFilt.Caption)

For Each Contrl In frmActFilt.Controls
    If TypeOf Contrl Is Label Then        '                           Label
            Contrl.Caption = ReadLangActFilt(Contrl.name & ".Caption", Contrl.Caption)
    End If

    If TypeOf Contrl Is XpB Then    '                    XPB
        Contrl.Caption = ReadLangActFilt(Contrl.name & ".Caption")
        '  Contrl.ToolTipText = ReadLangActFilt(Contrl.name & ".ToolTip")
        Contrl.pInitialize
    End If

Next

Me.Icon = FrmMain.Icon

'расставить галки
For i = 0 To UBound(arr_chAF)
chAF(i).Value = arr_chAF(i)
Next i
For i = 0 To UBound(arr_chAD)
chAD(i).Value = arr_chAD(i)
Next i

frmActFiltFlag = True
End Sub

Private Sub Form_Resize()
'Background
If lngBrush <> 0 Then
GetClientRect hwnd, rctMain
FillRect hdc, rctMain, lngBrush
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'расставить галки
Dim i As Integer

For i = 0 To UBound(arr_chAF)
arr_chAF(i) = chAF(i).Value
Next i
For i = 0 To UBound(arr_chAD)
arr_chAD(i) = chAD(i).Value
Next i

frmActFiltFlag = False
End Sub
