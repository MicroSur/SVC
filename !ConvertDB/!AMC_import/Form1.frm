VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   4500
      ScaleHeight     =   1515
      ScaleWidth      =   2475
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox tCurRec 
      Height          =   315
      Left            =   300
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1200
      Width           =   675
   End
   Begin VB.TextBox tTotalRec 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2940
      TabIndex        =   5
      Text            =   "total"
      Top             =   1200
      Width           =   1155
   End
   Begin VB.ListBox lData 
      Height          =   5715
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6915
   End
   Begin VB.CommandButton cNext 
      Caption         =   ">"
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cPrev 
      Caption         =   "<"
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cReadBase 
      Caption         =   "Читать"
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   660
      Width           =   3915
   End
   Begin VB.TextBox tBasePath 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "база"
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cReadBase_Click()

Dim tmps As String
Dim tmpl As Long

Dim N As Long
Dim File As String

File = App.Path & "\test.amc"

Open File For Binary As #1
N = 82  'пропуск хедера
Do While Not EOF(1)

'number
Dim MNumber As Long
MNumber = GetInt(N)
If MNumber < 1 Then Exit Do ' нет записей
lData.AddItem "Номер" & vbTab & vbTab & MNumber
N = Seek(1)
'Дата занесения - число от 14.09.1752 ?
lData.AddItem "Дата" & vbTab & vbTab & GetInt(N)
N = Seek(1)
'Рейтинг
tmps = GetInt(N) / 10
lData.AddItem "Рейтинг" & vbTab & vbTab & tmps
N = Seek(1)
'год
lData.AddItem "Год" & vbTab & vbTab & GetInt(N)
N = Seek(1)
'Length мин
lData.AddItem "Length" & vbTab & vbTab & GetInt(N)
N = Seek(1)
'VideoBitrate
lData.AddItem "VideoBitrate" & vbTab & GetInt(N)
N = Seek(1)
'AudioBitrate:
lData.AddItem "AudioBitrate" & vbTab & GetInt(N)
N = Seek(1)
'Disks:
lData.AddItem "Disks" & vbTab & vbTab & GetInt(N)
N = Seek(1)
'Expotr
lData.AddItem "Export" & vbTab & vbTab & GetBoolean(N)
N = Seek(1)
'Label
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Label" & vbTab & vbTab & tmps
N = Seek(1)

'MediaType
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "MediaType" & vbTab & tmps
N = Seek(1)

'Source
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Source" & vbTab & vbTab & tmps
N = Seek(1)

'Borrower
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Borrower" & vbTab & vbTab & tmps
N = Seek(1)

'OriginalTitle
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "OriginalTitle" & vbTab & tmps
N = Seek(1)

'TranslatedTitle
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "TranslatedTitle" & vbTab & tmps
N = Seek(1)

'Director
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Director" & vbTab & vbTab & tmps
N = Seek(1)

'Producer
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Producer" & vbTab & vbTab & tmps
N = Seek(1)

'Country
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Country" & vbTab & vbTab & tmps
N = Seek(1)

'Category
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Category" & vbTab & vbTab & tmps
N = Seek(1)

'Actors
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Actors" & vbTab & vbTab & tmps
N = Seek(1)

'URL
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "URL" & vbTab & vbTab & tmps
N = Seek(1)

'Description
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Description" & vbTab & tmps
N = Seek(1)

'Comments
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Comments" & vbTab & tmps
N = Seek(1)

'VideoFormat
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "VideoFormat" & vbTab & tmps
N = Seek(1)

'AudioFormat
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "AudioFormat" & vbTab & tmps
N = Seek(1)

'Resolution
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Resolution" & vbTab & tmps
N = Seek(1)

'Framerate
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Framerate" & vbTab & tmps
N = Seek(1)

'Languages
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Languages" & vbTab & tmps
N = Seek(1)

'Subtitles
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Subtitles" & vbTab & vbTab & tmps
N = Seek(1)

'Size
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
tmps = GetStr(N, tmpl)
lData.AddItem "Size" & vbTab & vbTab & tmps
N = Seek(1)

'PictureName
Dim PicName As String

'.png','.jpg','.gif' или ссылка
    'длина
tmpl = GetInt(N)
N = Seek(1)
    'string
PicName = GetStr(N, tmpl)
lData.AddItem "PictureName" & vbTab & PicName
N = Seek(1)

'PictureSize
Dim PicSize As Long
PicSize = GetInt(N)
'если =0 то PictureName - ссылка на файл
lData.AddItem "Size" & vbTab & vbTab & PicSize
N = Seek(1)

'Picture
Dim img As ImageFile
Dim vec As Vector
Dim pb() As Byte
Set Picture1.Picture = Nothing
If PicSize > 0 Then
 ReDim pb(PicSize - 1)
 Get #1, N, pb
 Set img = New ImageFile
 Set vec = New Vector
 vec.BinaryData = pb
 Set img = vec.ImageFile
 If Not img Is Nothing Then
    Set Picture1.Picture = img.ARGBData.Picture(img.Width, img.Height)
 End If
Else
'потом взять только файл, если путь (настройки) указан
 If Len(Dir(PicName)) <> 0 Then
  Set img = New ImageFile
  img.LoadFile PicName
  If Not img Is Nothing Then
     Set Picture1.Picture = img.ARGBData.Picture(img.Width, img.Height)
  End If
 End If
End If

N = Seek(1)

'Exit Do

tTotalRec = MNumber
Loop
Close #1
End Sub

Private Function GetInt(nn As Long) As Long
Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte

Get #1, nn, b1
Get #1, , b2
Get #1, , b3
Get #1, , b4

On Error GoTo err

GetInt = b1 + Val(b2) * 256 + Val(b3) * 65535 + Val(b4) * 16777215

Exit Function
err:
GetInt = 0
End Function
Private Function GetBoolean(nn As Long) As Integer
Dim b As Byte
Get #1, nn, b
GetBoolean = CInt(b)
End Function
Private Function GetStr(nn As Long, ln As Long) As String
If ln < 1 Then Exit Function
Dim b() As Byte
ReDim b(ln - 1)
Get #1, nn, b
GetStr = StrConv(b, vbUnicode)
End Function
