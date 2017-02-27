Attribute VB_Name = "ModXML"
Option Explicit
Public MSXML As New MSXML2.DOMDocument 'Object 'Set MSXML = CreateObject("MSXML2.DOMDocument")
Public XML_Root As Object

Public Sub ProcessRoot(sFile As String)
'Dim i As Integer, j As Integer, k As Integer
 
'Set MSXML = CreateObject("MSXML2.DOMDocument")
'MSXML.createCDATASection
MSXML.async = False

If MSXML.Load(sFile) Then
  ' Get root of document
  Set XML_Root = MSXML.documentElement
'Debug.Print MSXML.Document.Charset
xmlRecurse XML_Root, vbNullString

End If
End Sub

Private Sub xmlRecurse(objXML As Object, strChild As String)
 Dim XML_tmp As Object, XML_tmp2 As Object
 Dim XML_Child As Object
 'Dim AddFlag As Boolean
 Dim i As Integer, j As Integer

On Error Resume Next
'Debug.Print objXML.ChildNodes.Length

For Each XML_Child In objXML.childNodes
'frmMain.PBar.Value = frmMain.PBar.Value + 1

If err Then MsgBox Error.Description, vbInformation: Exit For

'Debug.Print XML_Child.NodeName
'If Left$(XML_Child.NodeName, 1) <> "#" Then
If XML_Child.nodeType = 1 Then 'NODE_ELEMENT

If Len(strChild) <> 0 Then


 If XML_Child.nodeName = strChild Then
  Set XML_tmp = XML_Child.Attributes
  'атрибуты
  For j = 0 To XML_Child.Attributes.length - 1
   frmMain.ListItemIn.AddItem XML_tmp.Item(j).Name
   'Debug.Print XML_tmp.getNamedItem(XML_tmp.Item(j).Name).nodeValue
  Next j
  
 'или ноды
   For Each XML_tmp2 In XML_Child.childNodes
    If XML_tmp2.hasChildNodes Then
     frmMain.ListItemIn.AddItem XML_tmp2.nodeName
     'Debug.Print XML_tmp2.NodeName
    End If
   Next
  
Exit For ' 1 раз
End If
    xmlRecurse XML_Child, strChild

Else
'заполнение списка нод

  '  Debug.Print XML_Child.NodeName
    'Debug.Print vbTab & XML_Child.nodeTypedValue
    
   ' If Len(XML_Child.nodeTypedValue) = 0 Then
   '  If XML_Child.NodeName <> "#text" Then
   ' If XML_Child.hasChildNodes Or XML_Child.Attributes.Length > 0 Then
    
    frmMain.CmbTableIn.AddItem XML_Child.nodeName

    xmlRecurse XML_Child, vbNullString
    
End If

End If
Next


If Len(strChild) <> 0 Then

For i = 1 To frmMain.ListItemIn.ListCount - 1
Do
If UCase(frmMain.ListItemIn.List(i - 1)) = UCase(frmMain.ListItemIn.List(i)) Then
If Len(frmMain.ListItemIn.List(i)) = 0 Then Exit For
frmMain.ListItemIn.RemoveItem i
Else
    Exit Do
End If
Loop
Next i

Else

For i = 1 To frmMain.CmbTableIn.ListCount - 1
Do
If UCase(frmMain.CmbTableIn.List(i - 1)) = UCase(frmMain.CmbTableIn.List(i)) Then
If Len(frmMain.CmbTableIn.List(i)) = 0 Then Exit For
frmMain.CmbTableIn.RemoveItem i
Else
    Exit Do
End If
Loop
Next i

End If

    'Set XML_tmp = XML_Child.getElementsByTagName("Movie")
'        Debug.Print XML_Child.Attributes.getNamedItem("Date").nodeValue
        'Debug.Print XML_Child.Item(1).nodeValue '.getElementsByTagName("Number").nodeValue
        'For Each XML_tmp2 In XML_tmp
        'Debug.Print XML_tmp2.Item(0).nodeTypedValue
        'Next        'objNodeItems.Item(0).nodeTypedValue
        

'frmMain.PBar.Visible = False
End Sub

Public Sub ProcessChild(sChild As String)
  
xmlRecurse XML_Root, sChild

End Sub
