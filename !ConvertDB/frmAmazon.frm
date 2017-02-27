VERSION 5.00
Begin VB.Form FormAmazon 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SVC - Amazon Movie Covers Searcher"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "frmAmazon.frx":0000
   LinkTopic       =   "FormAmazon"
   MaxButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8235
      Left            =   120
      ScaleHeight     =   8175
      ScaleWidth      =   8355
      TabIndex        =   5
      Top             =   780
      Width           =   8415
   End
   Begin VB.CommandButton ComPicNext 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7500
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
   Begin VB.CommandButton ComPicPrev 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5820
      TabIndex        =   2
      Top             =   180
      Width           =   735
   End
   Begin VB.CommandButton ComGetPic 
      Caption         =   "Запрос"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Terminator"
      Top             =   180
      Width           =   3795
   End
   Begin VB.Label LabPicCount 
      Alignment       =   2  'Center
      Caption         =   "0/0"
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   300
      Width           =   855
   End
   Begin VB.Menu PicMenu 
      Caption         =   "PicMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "Сохранить как..."
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Скопировать в буфер обмена"
      End
   End
End
Attribute VB_Name = "FormAmazon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CurrentLPUA As Integer 'текущий номер ячейки LargePixURLArr
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Event ErrorDownload(FromPathName As String, ToPathName As String)
Public Event DownloadComplete(FromPathName As String, ToPathName As String)


Private Sub ComGetPic_Click()
Dim i As Integer
Screen.MousePointer = vbHourglass

ComPicNext.Enabled = False: ComPicPrev.Enabled = False
LabPicCount = "0/0"

DoEvents
Call AmazonQuery(UrlEncode(Text1.Text))

If UBound(LargePixURLArr) > 0 Then
CurrentLPUA = 1
'WebBrowser1.Navigate LargePixURLArr(CurrentLPUA)
'Debug.Print CurrentLPUA, LargePixURLArr(CurrentLPUA)
Set Picture1 = URL2Pic(LargePixURLArr(CurrentLPUA))
If UBound(LargePixURLArr) > 1 Then ComPicNext.Enabled = True
LabPicCount = CurrentLPUA & "/" & UBound(LargePixURLArr)
End If

Screen.MousePointer = vbNormal
End Sub



Private Sub ComPicPrev_Click()
If CurrentLPUA - 1 > LBound(LargePixURLArr) Then
CurrentLPUA = CurrentLPUA - 1

ComPicNext.Enabled = True
If CurrentLPUA - 1 = LBound(LargePixURLArr) Then ComPicPrev.Enabled = False
LabPicCount = CurrentLPUA & "/" & UBound(LargePixURLArr)

'WebBrowser1.Navigate LargePixURLArr(CurrentLPUA)
'Debug.Print CurrentLPUA, LargePixURLArr(CurrentLPUA)
Set Picture1 = URL2Pic(LargePixURLArr(CurrentLPUA))
Else
ComPicPrev.Enabled = False
End If

End Sub

Private Sub ComPicNext_Click()
'Dim tmp As Integer
If CurrentLPUA + 1 < UBound(LargePixURLArr) + 1 Then
CurrentLPUA = CurrentLPUA + 1

ComPicPrev.Enabled = True
If CurrentLPUA + 1 = UBound(LargePixURLArr) + 1 Then ComPicNext.Enabled = False
LabPicCount = CurrentLPUA & "/" & UBound(LargePixURLArr)

'WebBrowser1.Navigate LargePixURLArr(CurrentLPUA)
'Debug.Print CurrentLPUA, LargePixURLArr(CurrentLPUA)
Set Picture1 = URL2Pic(LargePixURLArr(CurrentLPUA))
Else
ComPicNext.Enabled = False
End If

End Sub

Private Sub Form_Load()
'WebBrowser1.Navigate "about: Поиск обложки к фильму на сайте Amazon.com: Впишите название фильма (по английски), нажмите кнопку Запрос. Используйте стрелки для выбота подходящей обложки. Скопируйте обложку (пункт по клику правой кнопки мыши на картинке), вставьте в Sur Video Catalog. Enjoy!"
Picture1.Print "Поиск обложки к фильму на сайте Amazon.com: "
Picture1.Print "Впишите название фильма (по английски), "
Picture1.Print "нажмите кнопку Запрос. "
Picture1.Print "Используйте стрелки для выбота подходящей обложки. "
Picture1.Print "Скопируйте обложку (пункт по клику правой кнопки мыши на картинке),"
Picture1.Print "вставьте в Sur Video Catalog. Enjoy!"
End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
Picture1.Picture = Picture1.Image
Clipboard.SetData Picture1.Picture ', vbCFBitmap
End Sub
Public Function DownloadFile(FromPathName As String, ToPathName As String)
If URLDownloadToFile(0, FromPathName, ToPathName, 0, 0) = 0 Then
DownloadFile = True
RaiseEvent DownloadComplete(FromPathName, ToPathName)
Else
DownloadFile = False
RaiseEvent ErrorDownload(FromPathName, ToPathName)
End If
End Function

Private Sub mnuSave_Click()
Dim temp As String
temp = pSaveDialog
If Len(temp) = 0 Then Exit Sub
Call DownloadFile(LargePixURLArr(CurrentLPUA), temp)

End Sub
Public Function ReverseString(ByVal InputString As String) As String
Dim lLen As Long, lCtr As Long
Dim sChar As String
Dim sAns As String
lLen = Len(InputString)
For lCtr = lLen To 1 Step -1
sChar = Mid(InputString, lCtr, 1)
sAns = sAns & sChar
Next
ReverseString = sAns
End Function
Public Function pSaveDialog() As String
   Dim cd As New cCommonDialog
   Dim dext As String
   Dim sFile As String
   
sFile = LargePixURLArr(CurrentLPUA)
dext = GetExt(sFile)
sFile = sFile & "." & dext

   If (cd.VBGetSaveFileName( _
      sFile, _
      Filter:="All|*.*", _
      FilterIndex:=1, _
      DefaultExt:=dext, _
      Owner:=Me.hWnd)) Then
      pSaveDialog = sFile
   End If
Set cd = Nothing
End Function
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
    If Picture1.Picture <> 0 Then
        PopupMenu PicMenu
    End If
    End If
    
End Sub


Public Function GetExt(ByRef sName As String) As String
Dim temp As String
Dim i As Integer

temp = ReverseString(sName)
i = InStr(1, temp, ".")
If i > 0 Then
    GetExt = ReverseString(Left$(temp, i - 1))
    temp = Right$(temp, Len(temp) - i)
    i = InStr(1, temp, "/")
    If i > 0 Then
    sName = ReverseString(Left$(temp, i - 1))
    End If
End If

End Function


