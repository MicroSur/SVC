Attribute VB_Name = "ModAMC"
Option Explicit
'for AMC

Public ff As Integer

Public Function GetInt(nn As Long) As Long
Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte

Get ff, nn, b1
Get ff, , b2
Get ff, , b3
Get ff, , b4

On Error GoTo err

GetInt = b1 + Val(b2) * 256 + Val(b3) * 65535 + Val(b4) * 16777215

Exit Function
err:
GetInt = 0
End Function
Public Function GetBoolean(nn As Long) As Integer
Dim b As Byte
Get ff, nn, b
GetBoolean = CInt(b)
End Function
Public Function GetStr(nn As Long, ln As Long) As String
If ln < 1 Then Exit Function
Dim b() As Byte
ReDim b(ln - 1)
Get ff, nn, b
GetStr = StrConv(b, vbUnicode, LCID)
End Function
