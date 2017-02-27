Attribute VB_Name = "modSendMes"
Option Explicit

'combobox константы
Public Const CB_ERR As Long = -1
Public Const CB_ADDSTRING As Long = &H143
Private Const CB_RESETCONTENT As Long = &H14B
Public Const CB_SETITEMDATA As Long = &H151
Public Const CB_SETCURSEL As Long = &H14E
Public Const CB_GETCOUNT = &H146
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETLBTEXT = &H148


'lv
Public Const LVIF_TEXT = &H1
Public Const LVIF_IMAGE = &H2
Public Const LVIF_PARAM = &H4
Public Const LVIF_STATE = &H8
Public Const LVIF_INDENT = &H10

Public Const LVM_FIRST = &H1000&
Public Const LVM_INSERTITEMA = (LVM_FIRST + 7)
Public Const LVM_INSERTITEMW = (LVM_FIRST + 77)
#If UNICODE Then
Public Const LVM_INSERTITEM = LVM_INSERTITEMW
#Else
Public Const LVM_INSERTITEM = LVM_INSERTITEMA
#End If
Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)

Private Const LVM_SCROLL = (LVM_FIRST + 20)
Private Const LVM_GETTOPINDEX        As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE    As Long = (LVM_FIRST + 40)
Private Const LVM_SETITEMSTATE       As Long = (LVM_FIRST + 43)
Private Const LVIS_FOCUSED           As Long = &H1
Private Const LVIS_SELECTED          As Long = &H2
'Private Const LVIF_STATE             As Long = &H8



#If UNICODE Then
Public Type LVITEM
    mask As Long
    iItem As Long
    iSubitem As Long
    State As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
' #if (_WIN32_IE >= =&H0300)
    iIndent As Long
' #end If
End Type
#Else
Public Type LVITEM
    mask As Long
    iItem As Long
    iSubitem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
' #if (_WIN32_IE >= =&H0300)
    iIndent As Long
' #end If
'#if (_WIN32_WINNT >= 0x501)
    iGroupId As Long
    cColumns As Long '; // tile view columns
    puColumns As Long
'#End If
End Type
#End If

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Function LV_EnsureVisible(ByRef ctlListview As Object, ByVal iIndex As Long) As Boolean
On Error GoTo Hell
 
Dim LV As LVITEM
Dim lvItemsPerPage As Long
Dim lvNeededItems As Long
Dim lvCurrentTopIndex As Long

ActNotManualClick = True
With ctlListview
    ' Since this is a multi-select list, we want to unselect all items before selecting the current track.
    ' Start at the beginning of the playlist and look for the TrackID.

    With LV
        .mask = LVIF_STATE
        .State = False
        .stateMask = LVIS_SELECTED
    End With
    Call SendMessage(.hwnd, LVM_SETITEMSTATE, -1, LV)      ' Poof

    ' Select and set the focus rectangle on the currently playing Track, and bring it to the viewable area.
    With LV
        .mask = LVIF_STATE
        .State = True
        .stateMask = LVIS_SELECTED Or LVIS_FOCUSED
    End With
    Call SendMessage(.hwnd, LVM_SETITEMSTATE, iIndex - 1, LV)      ' Listview index is 0-based in the API world

    ' Determine if desired index + number of items in view will exceed total items in the control
    lvCurrentTopIndex = SendMessage(.hwnd, LVM_GETTOPINDEX, 0&, ByVal 0&)
    lvCurrentTopIndex = lvCurrentTopIndex + 1     'заголовок
    lvItemsPerPage = SendMessage(.hwnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&) - 1

    ' Do we even need to scroll? Not if the selected track is already in view
    If (lvCurrentTopIndex > iIndex) Or (iIndex > lvCurrentTopIndex + lvItemsPerPage) Then

        ' Is 'x' above or below target index?
        If lvCurrentTopIndex > iIndex Then      ' Going UP
            If iIndex > 1 Then     '3
                .ListItems((iIndex - 1)).EnsureVisible       ' Drops the highlighted item down a few so it's not hidden
                ' behind the Column header.
            Else
                .ListItems((iIndex)).EnsureVisible
            End If

        Else     ' Going DOWN
            ' Are there sufficient items to set to the topindex
            If (iIndex + lvItemsPerPage) > .ListItems.Count Then

                ' Can't be set to the top as the control has insufficient
                ' items, so just scroll to the end of listview
                .ListItems(.ListItems.Count).EnsureVisible

            Else

                ' It is below, and since a listview always moves the item just into view,
                ' have it instead move to the top by faking item we want to 'EnsureVisible'
                ' the item lvItemsPerPage -1(or -3) below the actual index of interest.
                If iIndex > 1 Then
                    .ListItems((iIndex + lvItemsPerPage) - 1).EnsureVisible     '3
                Else
                    .ListItems((iIndex + lvItemsPerPage) - 1).EnsureVisible
                End If
            End If
        End If
    End If

ActNotManualClick = False
LV_EnsureVisible = True

End With
Exit Function

Hell:
If err.Number <> 0 Then
    ToDebug "LV_EVis: " & err.Description
End If
End Function

Public Function ListViewScroll(lvw As ListView, ByVal dX As Long, ByVal dy As Long)
    SendMessage lvw.hwnd, LVM_SCROLL, dX, ByVal dy
End Function
Public Sub AddItem(cmb As ComboBox, Text As String, Optional ItemData As Long)
   Dim l As Long
  
   l = SendMessage(cmb.hwnd, CB_ADDSTRING, 0, ByVal Text)
   If l <> CB_ERR Then
      SendMessage cmb.hwnd, CB_SETITEMDATA, l, ByVal ItemData
   End If
End Sub

Public Sub Clear(cmb As ComboBox)
   SendMessage cmb.hwnd, CB_RESETCONTENT, 0, 0&
End Sub


'Public Function lvAddItem( _
'      ByVal m_hWnd As Long, _
'      ByVal sText As String, _
'      Optional ByVal lIndex As Long = 1, _
'      Optional ByVal iIcon As Long = -1, _
'      Optional ByVal iIndent As Long = 0, _
'      Optional ByVal lItemData As Long = 0 _
'   ) As Boolean
'Dim tLV As LVITEM
'Dim lR As Long
'Dim lOrigCount As Long
'   lOrigCount = fCount(m_hWnd)
'   tLV.pszText = sText & vbNullChar
'   tLV.cchTextMax = Len(sText) + 1
'   tLV.iImage = iIcon
'   tLV.iIndent = iIndent
'   tLV.lParam = lItemData
'   tLV.iItem = lIndex - 1
'   tLV.mask = LVIF_TEXT Or LVIF_IMAGE Or LVIF_PARAM Or LVIF_INDENT
'   lR = SendMessage(m_hWnd, LVM_INSERTITEM, 0, tLV)
'   lvAddItem = Not (fCount(m_hWnd) = lOrigCount)
'End Function

Public Function fCount(ByVal m_hWnd As Long) As Long
   fCount = SendMessageLong(m_hWnd, LVM_GETITEMCOUNT, 0, 0)
End Function
