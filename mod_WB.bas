Attribute VB_Name = "mod_WB"
Option Explicit
'Private WithEvents Document As MSHTML.HTMLDocument 'правый клик
Public TempPicPath As String

'Public Sub Init_WB()
'FrmMain.WBBR.Navigate2 "about:blank" 'in load
'Set WBBRDoc = FrmMain.WBBR.Document 'после навигейта
'End Sub
Private Function WritePicFromBase(dfield As String) As Boolean

Dim b() As Byte
Dim PicSize As Long
Dim fname As String
Dim LFile As Integer

On Error GoTo err


LFile = FreeFile

PicSize = rs.Fields(dfield).FieldSize
If PicSize > 0 Then
    ReDim b(PicSize - 1)
    b() = rs.Fields(dfield).GetChunk(0, PicSize)


    Select Case dfield
    Case "snapshot1"
        fname = "ScreenShot1.jpg"
    Case "snapshot1"
        fname = "ScreenShot2.jpg"
    Case "snapshot1"
        fname = "ScreenShot3.jpg"
    Case "frontface"
        fname = "CoverFront.jpg"
    End Select

    fname = TempPicPath & fname

    Open fname For Binary As LFile
    Put #LFile, 1, b()
    Close LFile

    'Print #m_lFile, StrConv(m_oRS(0).Value, vbUnicode)
    WritePicFromBase = True

End If    'PicSize > 0
Exit Function
err:
MsgBox err.Description
End Function

Public Sub ShowRowInWB()
'показать в браузере значения текущей строки
Dim sHeader As String
Dim sBody As String
Dim sFooter As String
Dim tmp As String
Dim HTMLText As String
Dim i As Integer
Dim ifile As Integer
Dim value As String
Dim LFile As Integer
Dim fldCaption As String 'переведенное название поля

On Error GoTo err
LFile = FreeFile

'WBBR.Navigate2 "about:blank" 'in load
'грузить HTMLText из файла шаблона и править по ходу

sHeader = "<html><head><title>db picture test protocol</title></head><body>"
sFooter = "</body></html>"


For i = 0 To rs.Fields.Count - 1
'DoEvents
    'fldCaption = rs.Fields(i).Properties("Caption") 'не долго ли
    tmp = vbNullString    'обнулить темп
    Select Case LCase$(rs(i).name)
    Case "moviename"
        value = CheckNoNullVal(i)
        If Len(value) <> 0 Then
            tmp = "<br>" & "moviename" & ": " & value
        End If
    Case "label"
        value = CheckNoNullVal(i)
        If Len(value) <> 0 Then
            tmp = "<br>" & "label" & ": " & value
        End If

    Case "snapshot1"
        If WritePicFromBase("snapshot1") Then
            tmp = "<br><img src='" & TempPicPath & "ScreenShot1.jpg' width=200>"
        End If
    Case "snapshot1"
        If WritePicFromBase("snapshot1") Then
            tmp = "<br><img src='" & TempPicPath & "ScreenShot2.jpg' width=200>"
        End If
    Case "snapshot1"
        If WritePicFromBase("snapshot1") Then
            tmp = "<br><img src='" & TempPicPath & "ScreenShot3.jpg' width=200>"
        End If
    Case "frontface"
        If WritePicFromBase("frontface") Then
            tmp = "<br><img src='" & TempPicPath & "CoverFront.jpg' width=200>"
        End If
    End Select
    If Len(tmp) <> 0 Then sBody = sBody & vbCrLf & tmp
Next i

HTMLText = sHeader & vbCrLf & sBody & vbCrLf & sFooter

''для теста положить html в файл
'fname = TempPicPath & "sgctemp.htm"
'Open fname For Output As lFile
'Print #lFile, HTMLText
'Close lFile

'показать в HTMLText WB

With FrmMain
'LockWindowUpdate .WBBR.hWnd
.WBBR.Silent = True
.WBBR.Document.Script.Document.Clear
.WBBR.Document.Script.Document.write HTMLText
.WBBR.Document.Script.Document.Close
'LockWindowUpdate 0
End With

Exit Sub
err:
MsgBox err.Description
End Sub

