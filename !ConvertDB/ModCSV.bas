Attribute VB_Name = "ModCSV"
Option Explicit
Public Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long

Public Function ParseCSV01(ByRef Expression As String, ByRef Separator As String, asValues() As String) As Long
' by Donald, donald@xbeat.net, 20020603, rev 20020701
  
  Const lAscSpace     As Long = 32   ' Asc(" ")
  Const lAscQuote     As Long = 34   ' Asc("""")
  'Const lAscSeparator As Long = 44   ' Asc(","), comma
  Dim lAscSeparator As Long
  lAscSeparator = Asc(Separator)
  
  Const lValueNone    As Long = 0 ' states of the parser
  Const lValuePlain   As Long = 1
  Const lValueQuoted  As Long = 2
  
  ' BUFFERREDIM is ideally exactly the number of values in Expression (minus 1)
  ' so: if you know what to expect, fine-tune here
  Const BUFFERREDIM   As Long = 64
  Dim ubValues        As Long
  Dim cntValues       As Long
  
  Dim abExpression() As Byte
  Dim lCharCode As Long
  Dim posStart As Long
  Dim posEnd As Long
  Dim cntTrim As Long
  Dim lState As Long
  Dim i As Long
  
  If LenB(Expression) > 0 Then
    
    abExpression = Expression         ' to byte array
    ubValues = -1 + BUFFERREDIM
    ReDim Preserve asValues(ubValues) ' init array (Preserve is faster)
    
    For i = 0 To UBound(abExpression) Step 2
      
      ' 1. unicode char has 16 bits, but 32 bit Longs process faster
      ' 2. add lower and upper byte: ignoring the upper byte can lead to misinterpretations
      lCharCode = abExpression(i) Or (&H100 * abExpression(i + 1))
      
      Select Case lCharCode
      
      Case lAscSpace
        If lState = lValuePlain Then
          ' at non-quoted value: trim 2 unicode bytes for each space
          cntTrim = cntTrim + 2
        End If
      
      Case lAscSeparator
        If lState = lValueNone Then
          ' ends zero-length value
          If cntValues > ubValues Then
            ubValues = ubValues + BUFFERREDIM
            ReDim Preserve asValues(ubValues)
          End If
          asValues(cntValues) = ""
          cntValues = cntValues + 1
          posStart = i + 2
        ElseIf lState = lValuePlain Then
          ' ends non-quoted value
          lState = lValueNone
          posEnd = i - cntTrim
          If cntValues > ubValues Then
            ubValues = ubValues + BUFFERREDIM
            ReDim Preserve asValues(ubValues)
          End If
          asValues(cntValues) = MidB$(Expression, posStart + 1, posEnd - posStart)
          cntValues = cntValues + 1
          posStart = i + 2
          cntTrim = 0
        End If
      
      Case lAscQuote
        If lState = lValueNone Then
          ' starts quoted value
          lState = lValueQuoted
          ' trims the opening quote
          posStart = i + 2
        ElseIf lState = lValueQuoted Then
          ' ends quoted value, or is a quote within
          lState = lValuePlain
          ' trims the closing quote
          cntTrim = 2
        End If
      
      Case Else
        If lState = lValueNone Then
          ' starts non-quoted value
          lState = lValuePlain
          posStart = i
        End If
        ' reset trimming
        cntTrim = 0
      
      End Select
    
    Next
    
    ' remainder
    posEnd = i - cntTrim
    If cntValues <> ubValues Then
      ReDim Preserve asValues(cntValues)
    End If
    asValues(cntValues) = MidB$(Expression, posStart + 1, posEnd - posStart)
    ParseCSV01 = cntValues + 1
  
  Else
    ' (Expression = "")
    ' return single-element array containing a zero-length string
    'ReDim asValues(0)
    ParseCSV01 = 0 '1
  
  End If

End Function


