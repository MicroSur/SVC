Attribute VB_Name = "ModListBox"
Option Explicit
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                     (ByVal hWnd As Long, ByVal wMsg As Long, _
                      ByVal wParam As Long, lParam As Long) As Long

Public Function ListRowCalc(lstTemp As Control, ByVal y As Single) As Integer
Const LB_GETITEMHEIGHT = &H1A1
'Determines the height of each item in ListBox control in pixels
Dim ItemHeight As Integer
ItemHeight = SendMessage(lstTemp.hWnd, LB_GETITEMHEIGHT, 0, 0)
ListRowCalc = min(((y / Screen.TwipsPerPixelY) \ ItemHeight) + _
                  lstTemp.TopIndex, lstTemp.ListCount - 1)
End Function

Private Function min(x As Integer, y As Integer) As Integer
If x > y Then min = y Else min = x
End Function

Public Function ListRowMove(lstTemp As ListBox, ByVal OldRow As Integer, ByVal NewRow As Integer) As String
'переставляли мышкой строки в лб
Dim SaveList As String, i As Integer
Dim SelectedRow As String

On Error Resume Next

If OldRow = NewRow Then Exit Function
SaveList = lstTemp.List(OldRow)

'запомнить текст помеченного
For i = 0 To lstTemp.ListCount - 1
If lstTemp.Selected(i) = True Then SelectedRow = lstTemp.List(i): Exit For
Next i

If OldRow > NewRow Then
    For i = OldRow To NewRow + 1 Step -1
        lstTemp.List(i) = lstTemp.List(i - 1)
    Next i
Else
    For i = OldRow To NewRow - 1
        lstTemp.List(i) = lstTemp.List(i + 1)
    Next i
End If

lstTemp.List(NewRow) = SaveList

'пометить бывший помеченный (по содержанию)
If Len(SelectedRow) <> 0 Then
For i = 0 To lstTemp.ListCount - 1
If lstTemp.List(i) = SelectedRow Then
    lstTemp.Selected(i) = True
    ListRowMove = SelectedRow
    Exit For
End If
Next i
End If

End Function


