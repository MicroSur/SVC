VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'‘ункции доступные скрипту (SVC.функци€)
'¬озвращать в скрипт Variant

Public Function PageArr(ln As Variant) As Variant
'доступ из скрипта к массиву PageArray
On Error Resume Next

If UBound(PageArray) >= CInt(ln) Then
PageArr = CVar(PageArray(Abs(CInt(ln))))
Else
PageArr = CVar(vbNullString)
End If

End Function
Public Function FindLine(Pattern As Variant, StartAt As Variant) As Variant
'¬озвращает номер строки в массиве текущей страницы PageArray, содержащей Pattern.
'Ќачинает поиск с StartAt
Dim i As Integer
On Error Resume Next

FindLine = -1
If StartAt < 0 Then StartAt = 0
For i = StartAt To UBound(PageArray)
    If InStr(PageArray(i), Pattern) <> 0 Then
        FindLine = CVar(i)
        Exit For
    End If
Next
End Function

Public Function HTML2TEXT(strIn As Variant) As Variant
Dim InTag  As Boolean
Dim ch As String * 1
Dim so As String
Dim idx As Long, idx2 As Long
Dim tmp As String
Dim si As String

On Error Resume Next

si = Trim$(strIn)
If Len(si) = 0 Then HTML2TEXT = "": Exit Function
so = String(Len(si), " ")

For idx = 1 To Len(si)
    ch = Mid$(si, idx, 1)
    If ch = "<" Then InTag = True: ch = vbNullString
    If ch = ">" Then InTag = False: ch = vbNullString
    
    If Not (InTag) Then
        idx2 = idx2 + 1
        Mid$(so, idx2, 1) = ch
    End If
    
Next

tmp = Left$(so, idx2)
'tmp = Trim$(tmp)

'tmp = Replace(tmp, " ,", ",", , , vbTextCompare)
Do While InStr(1, tmp, " ,")
tmp = Replace(tmp, " ,", ",", , , vbTextCompare)
Loop

tmp = Replace(tmp, "&nbsp;", " ", , , vbTextCompare) 'часто

''''''''''''''''                                            &&&&&&&&&&&
If InStr(tmp, "&") Then
'ASCII Entities
tmp = Replace(tmp, "&quot;", """", , , vbTextCompare)
tmp = Replace(tmp, "&apos;", "'", , , vbTextCompare)
tmp = Replace(tmp, "&amp;", "&", , , vbTextCompare)
tmp = Replace(tmp, "&lt;", "<", , , vbTextCompare)
tmp = Replace(tmp, "&gt;", ">", , , vbTextCompare)

'
tmp = Replace(tmp, "&#34;", """", , , vbTextCompare)
tmp = Replace(tmp, "&#38;", "&", , , vbTextCompare)
tmp = Replace(tmp, "&#39;", "'", , , vbTextCompare)
tmp = Replace(tmp, "&#149;", ",", , , vbTextCompare)
tmp = Replace(tmp, "&#151;", "-", , , vbTextCompare)

'ISO 8859-1 Symbol Entities
'tmp = Replace(tmp, "&nbsp;", " ", , , vbTextCompare) 'в начало
    tmp = Replace(tmp, "&#160;", " ", , , vbTextCompare)
tmp = Replace(tmp, "&iexcl;", "i", , , vbTextCompare) 'inverted exclamation mark
    tmp = Replace(tmp, "&#161;", "i", , , vbTextCompare)
tmp = Replace(tmp, "&curren;", "§", , , vbTextCompare) 'currency
    tmp = Replace(tmp, "&#164;", "§", , , vbTextCompare)
tmp = Replace(tmp, "&cent;", "c", , , vbTextCompare) 'cent
    tmp = Replace(tmp, "&#162;", "c", , , vbTextCompare)
tmp = Replace(tmp, "&pound;", "f", , , vbTextCompare) 'pound
    tmp = Replace(tmp, "&#163;", "f", , , vbTextCompare)
tmp = Replace(tmp, "&yen;", "y", , , vbTextCompare) 'yen
    tmp = Replace(tmp, "&#165;", "y", , , vbTextCompare)
tmp = Replace(tmp, "&brvbar;", "¶", , , vbTextCompare) 'broken vertical bar
    tmp = Replace(tmp, "&#166;", "¶", , , vbTextCompare)
tmp = Replace(tmp, "&sect;", "І", , , vbTextCompare) 'section
    tmp = Replace(tmp, "&#167;", "І", , , vbTextCompare)
tmp = Replace(tmp, "&uml;", " ", , , vbTextCompare) 'spacing diaeresis я
    tmp = Replace(tmp, "&#168;", " ", , , vbTextCompare)
tmp = Replace(tmp, "&copy;", "©", , , vbTextCompare) 'copyright
    tmp = Replace(tmp, "&#169;", "©", , , vbTextCompare)
tmp = Replace(tmp, "&ordf;", "a", , , vbTextCompare) 'feminine ordinal indicator
    tmp = Replace(tmp, "&#170;", "a", , , vbTextCompare)
tmp = Replace(tmp, "&laquo;", "Ђ", , , vbTextCompare) 'angle quotation mark (left) {{
    tmp = Replace(tmp, "&#171;", "Ђ", , , vbTextCompare)
tmp = Replace(tmp, "&not;", "ђ", , , vbTextCompare) 'negation
    tmp = Replace(tmp, "&#172;", "ђ", , , vbTextCompare)
tmp = Replace(tmp, "&shy;", " ", , , vbTextCompare) 'soft hyphen
    tmp = Replace(tmp, "&#173;", " ", , , vbTextCompare)
tmp = Replace(tmp, "&reg;", "Ѓ", , , vbTextCompare) 'registered trademark
    tmp = Replace(tmp, "&#174;", "Ѓ", , , vbTextCompare)
tmp = Replace(tmp, "&trade;", "Щ", , , vbTextCompare) 'trademark
    tmp = Replace(tmp, "&#8482;", "Щ", , , vbTextCompare)
tmp = Replace(tmp, "&macr;", "_", , , vbTextCompare) 'spacing macron
    tmp = Replace(tmp, "&#175;", "_", , , vbTextCompare)
tmp = Replace(tmp, "&deg;", "∞", , , vbTextCompare) 'degree
    tmp = Replace(tmp, "&#176;", "∞", , , vbTextCompare)
tmp = Replace(tmp, "&plusmn;", "±", , , vbTextCompare) 'plus-or-minus
    tmp = Replace(tmp, "&#177;", "±", , , vbTextCompare)
tmp = Replace(tmp, "&sup2;", "^2", , , vbTextCompare) 'superscript 2
    tmp = Replace(tmp, "&#178;", "^2", , , vbTextCompare)
tmp = Replace(tmp, "&sup3;", "^3", , , vbTextCompare) 'superscript 3
    tmp = Replace(tmp, "&#179;", "^3", , , vbTextCompare)
tmp = Replace(tmp, "&acute;", "'", , , vbTextCompare) 'spacing acute
    tmp = Replace(tmp, "&#180;", "'", , , vbTextCompare)
tmp = Replace(tmp, "&micro;", "µ", , , vbTextCompare) 'micro
    tmp = Replace(tmp, "&#181;", "µ", , , vbTextCompare)
tmp = Replace(tmp, "&para;", "ґ", , , vbTextCompare) 'paragraph
    tmp = Replace(tmp, "&#182;", "ґ", , , vbTextCompare)
tmp = Replace(tmp, "&middot;", "Ј", , , vbTextCompare) 'middle dot
    tmp = Replace(tmp, "&#183;", "Ј", , , vbTextCompare)
tmp = Replace(tmp, "&cedil;", " ", , , vbTextCompare) 'spacing cedilla €
    tmp = Replace(tmp, "&#184;", " ", , , vbTextCompare)
tmp = Replace(tmp, "&sup1;", "^1", , , vbTextCompare) 'superscript 1
    tmp = Replace(tmp, "&#185;", "^1", , , vbTextCompare)
tmp = Replace(tmp, "&ordm;", "^0", , , vbTextCompare) 'masculine ordinal indicator
    tmp = Replace(tmp, "&#186;", "^0", , , vbTextCompare)
tmp = Replace(tmp, "&raquo;", "ї", , , vbTextCompare) 'angle quotation mark (right)
    tmp = Replace(tmp, "&#187;", "ї", , , vbTextCompare)
tmp = Replace(tmp, "&frac14;", "1/4", , , vbTextCompare) 'fraction 1/4
    tmp = Replace(tmp, "&#188;", "1/4", , , vbTextCompare)
tmp = Replace(tmp, "&frac12;", "1/2", , , vbTextCompare) 'fraction 1/2
    tmp = Replace(tmp, "&#189;", "1/2", , , vbTextCompare)
tmp = Replace(tmp, "&frac34;", "3/4", , , vbTextCompare) 'fraction 3/4
    tmp = Replace(tmp, "&#190;", "3/4", , , vbTextCompare)
tmp = Replace(tmp, "&iquest;", "?", , , vbTextCompare) 'inverted question mark
    tmp = Replace(tmp, "&#191;", "?", , , vbTextCompare)
tmp = Replace(tmp, "&times;", "4", , , vbTextCompare) 'multiplication „
    tmp = Replace(tmp, "&#215;", "4", , , vbTextCompare)

'ISO 8859-1 Character Entities
tmp = Replace(tmp, "&Agrave;", "A", , , vbTextCompare) '
    tmp = Replace(tmp, "&#192;", "A", , , vbTextCompare)
tmp = Replace(tmp, "&Aacute;", "A", , , vbTextCompare) '
    tmp = Replace(tmp, "&#193;", "A", , , vbTextCompare)
tmp = Replace(tmp, "&Acirc;", "A", , , vbTextCompare) '
    tmp = Replace(tmp, "&#194;", "A", , , vbTextCompare)
tmp = Replace(tmp, "&Atilde;", "A", , , vbTextCompare) '
    tmp = Replace(tmp, "&#195;", "A", , , vbTextCompare)
tmp = Replace(tmp, "&Auml;", "A", , , vbTextCompare) '
    tmp = Replace(tmp, "&#196;", "A", , , vbTextCompare)
tmp = Replace(tmp, "&Aring;", "A", , , vbTextCompare) '
    tmp = Replace(tmp, "&#197;", "A", , , vbTextCompare)
    
tmp = Replace(tmp, "&AElig;", "AE", , , vbTextCompare) '
    tmp = Replace(tmp, "&#198;", "AE", , , vbTextCompare)
    
tmp = Replace(tmp, "&Ccedil;", "C", , , vbTextCompare) '
    tmp = Replace(tmp, "&#199;", "C", , , vbTextCompare)
    
tmp = Replace(tmp, "&Egrave;", "E", , , vbTextCompare) '
    tmp = Replace(tmp, "&#200;", "E", , , vbTextCompare)
tmp = Replace(tmp, "&Eacute;", "E", , , vbTextCompare) '
    tmp = Replace(tmp, "&#201;", "E", , , vbTextCompare)
tmp = Replace(tmp, "&Ecirc;", "E", , , vbTextCompare) '
    tmp = Replace(tmp, "&#202;", "E", , , vbTextCompare)
tmp = Replace(tmp, "&Euml;", "E", , , vbTextCompare) '
    tmp = Replace(tmp, "&#203;", "E", , , vbTextCompare)
    
tmp = Replace(tmp, "&Igrave;", "I", , , vbTextCompare) '
    tmp = Replace(tmp, "&#204;", "I", , , vbTextCompare)
tmp = Replace(tmp, "&Iacute;", "I", , , vbTextCompare) '
    tmp = Replace(tmp, "&#205;", "I", , , vbTextCompare)
tmp = Replace(tmp, "&Icirc;", "I", , , vbTextCompare) '
    tmp = Replace(tmp, "&#206;", "I", , , vbTextCompare)
tmp = Replace(tmp, "&Iuml;", "I", , , vbTextCompare) '
    tmp = Replace(tmp, "&#207;", "I", , , vbTextCompare)
    
tmp = Replace(tmp, "&ETH;", "ETH", , , vbTextCompare) '
    tmp = Replace(tmp, "&#208;", "ETH", , , vbTextCompare)
tmp = Replace(tmp, "&Ntilde;", "N", , , vbTextCompare) '
    tmp = Replace(tmp, "&#209;", "N", , , vbTextCompare)
    
tmp = Replace(tmp, "&Ograve;", "O", , , vbTextCompare) '
    tmp = Replace(tmp, "&#210;", "O", , , vbTextCompare)
tmp = Replace(tmp, "&Oacute;", "O", , , vbTextCompare) '
    tmp = Replace(tmp, "&#211;", "O", , , vbTextCompare)
tmp = Replace(tmp, "&Ocirc;", "O", , , vbTextCompare) '
    tmp = Replace(tmp, "&#212;", "O", , , vbTextCompare)
tmp = Replace(tmp, "&Otilde;", "O", , , vbTextCompare) '
    tmp = Replace(tmp, "&#213;", "O", , , vbTextCompare)
tmp = Replace(tmp, "&Ouml;", "O", , , vbTextCompare) '
    tmp = Replace(tmp, "&#214;", "O", , , vbTextCompare)
tmp = Replace(tmp, "&Oslash;", "O", , , vbTextCompare) '
    tmp = Replace(tmp, "&#216;", "O", , , vbTextCompare)
    
tmp = Replace(tmp, "&Ugrave;", "U", , , vbTextCompare) '
    tmp = Replace(tmp, "&#217;", "U", , , vbTextCompare)
tmp = Replace(tmp, "&Uacute;", "U", , , vbTextCompare) '
    tmp = Replace(tmp, "&#218;", "U", , , vbTextCompare)
tmp = Replace(tmp, "&Ucirc;", "U", , , vbTextCompare) '
    tmp = Replace(tmp, "&#219;", "U", , , vbTextCompare)
tmp = Replace(tmp, "&Uuml;", "U", , , vbTextCompare) '
    tmp = Replace(tmp, "&#220;", "U", , , vbTextCompare)
    
tmp = Replace(tmp, "&Yacute;", "Y", , , vbTextCompare) '
    tmp = Replace(tmp, "&#221;", "Y", , , vbTextCompare)
tmp = Replace(tmp, "&THORN;", "THORN", , , vbTextCompare) '
    tmp = Replace(tmp, "&#222;", "THORN", , , vbTextCompare)
    
tmp = Replace(tmp, "&szlig;", "s", , , vbTextCompare) '
    tmp = Replace(tmp, "&#223;", "s", , , vbTextCompare)
    
tmp = Replace(tmp, "&agrave;", "a", , , vbTextCompare) '
    tmp = Replace(tmp, "&#224;", "a", , , vbTextCompare)
tmp = Replace(tmp, "&aacute;", "a", , , vbTextCompare) '
    tmp = Replace(tmp, "&#225;", "a", , , vbTextCompare)
tmp = Replace(tmp, "&acirc;", "a", , , vbTextCompare) '
    tmp = Replace(tmp, "&#226;", "a", , , vbTextCompare)
tmp = Replace(tmp, "&atilde;", "a", , , vbTextCompare) '
    tmp = Replace(tmp, "&#227;", "a", , , vbTextCompare)
tmp = Replace(tmp, "&auml;", "a", , , vbTextCompare) '
    tmp = Replace(tmp, "&#228;", "a", , , vbTextCompare)
tmp = Replace(tmp, "&aring;", "a", , , vbTextCompare) '
    tmp = Replace(tmp, "&#229;", "a", , , vbTextCompare)
    
tmp = Replace(tmp, "&aelig;", "ae", , , vbTextCompare) '
    tmp = Replace(tmp, "&#230;", "ae", , , vbTextCompare)
    
tmp = Replace(tmp, "&ccedil;", "c", , , vbTextCompare) '
    tmp = Replace(tmp, "&#231;", "c", , , vbTextCompare)
    
tmp = Replace(tmp, "&egrave;", "e", , , vbTextCompare) '
    tmp = Replace(tmp, "&#232;", "e", , , vbTextCompare)
tmp = Replace(tmp, "&eacute;", "e", , , vbTextCompare) '
    tmp = Replace(tmp, "&#233;", "e", , , vbTextCompare)
tmp = Replace(tmp, "&ecirc;", "e", , , vbTextCompare) '
    tmp = Replace(tmp, "&#234;", "e", , , vbTextCompare)
tmp = Replace(tmp, "&euml;", "e", , , vbTextCompare) '
    tmp = Replace(tmp, "&#235;", "e", , , vbTextCompare)
    
tmp = Replace(tmp, "&igrave;", "i", , , vbTextCompare)  '
    tmp = Replace(tmp, "&#236;", "i", , , vbTextCompare)
tmp = Replace(tmp, "&iacute;", "i", , , vbTextCompare) '
    tmp = Replace(tmp, "&#237;", "i", , , vbTextCompare)
tmp = Replace(tmp, "&icirc;", "i", , , vbTextCompare) '
    tmp = Replace(tmp, "&#238;", "i", , , vbTextCompare)
tmp = Replace(tmp, "&iuml;", "i", , , vbTextCompare) '
    tmp = Replace(tmp, "&#239;", "i", , , vbTextCompare)
    
tmp = Replace(tmp, "&eth;", "eth", , , vbTextCompare) '
    tmp = Replace(tmp, "&#240;", "eth", , , vbTextCompare)
tmp = Replace(tmp, "&ntilde;", "n", , , vbTextCompare) '
    tmp = Replace(tmp, "&#241;", "n", , , vbTextCompare)
    
tmp = Replace(tmp, "&ograve;", "o", , , vbTextCompare) '
    tmp = Replace(tmp, "&#242;", "o", , , vbTextCompare)
tmp = Replace(tmp, "&oacute;", "o", , , vbTextCompare) '
    tmp = Replace(tmp, "&#243;", "o", , , vbTextCompare)
tmp = Replace(tmp, "&ocirc;", "o", , , vbTextCompare) '
    tmp = Replace(tmp, "&#244;", "o", , , vbTextCompare)
tmp = Replace(tmp, "&otilde;", "o", , , vbTextCompare) '
    tmp = Replace(tmp, "&#245;", "o", , , vbTextCompare)
tmp = Replace(tmp, "&ouml;", "o", , , vbTextCompare) '
    tmp = Replace(tmp, "&#246;", "o", , , vbTextCompare)
tmp = Replace(tmp, "&oslash;", "o", , , vbTextCompare) '
    tmp = Replace(tmp, "&#248;", "o", , , vbTextCompare)
    
tmp = Replace(tmp, "&ugrave;", "u", , , vbTextCompare) '
    tmp = Replace(tmp, "&#249;", "u", , , vbTextCompare)
tmp = Replace(tmp, "&uacute;", "u", , , vbTextCompare) '
    tmp = Replace(tmp, "&#250;", "u", , , vbTextCompare)
tmp = Replace(tmp, "&ucirc;", "u", , , vbTextCompare) '
    tmp = Replace(tmp, "&#251;", "u", , , vbTextCompare)
tmp = Replace(tmp, "&uuml;", "u", , , vbTextCompare) '
    tmp = Replace(tmp, "&#252;", "u", , , vbTextCompare)
    
tmp = Replace(tmp, "&yacute;", "y", , , vbTextCompare) '
    tmp = Replace(tmp, "&#253;", "y", , , vbTextCompare)
tmp = Replace(tmp, "&thorn;", "thorn", , , vbTextCompare) '
    tmp = Replace(tmp, "&#254;", "thorn", , , vbTextCompare)
tmp = Replace(tmp, "&yuml;", "y", , , vbTextCompare) '
    tmp = Replace(tmp, "&#255;", "y", , , vbTextCompare)

'Some Other Entities supported by HTML
tmp = Replace(tmp, "&OElig;", "OE", , , vbTextCompare) '
    tmp = Replace(tmp, "&#338;", "OE", , , vbTextCompare)
tmp = Replace(tmp, "&oelig;", "oe", , , vbTextCompare) '
    tmp = Replace(tmp, "&#339;", "oe", , , vbTextCompare)
tmp = Replace(tmp, "&Scaron;", "S", , , vbTextCompare) '
    tmp = Replace(tmp, "&#352;", "S", , , vbTextCompare)
tmp = Replace(tmp, "&scaron;", "s", , , vbTextCompare) '
    tmp = Replace(tmp, "&#353;", "s", , , vbTextCompare)
tmp = Replace(tmp, "&Yuml;", "Y", , , vbTextCompare) '
    tmp = Replace(tmp, "&#376;", "Y", , , vbTextCompare)
tmp = Replace(tmp, "&circ;", "^", , , vbTextCompare) '
    tmp = Replace(tmp, "&#710;", "^", , , vbTextCompare)
tmp = Replace(tmp, "&tilde;", "~", , , vbTextCompare) '
    tmp = Replace(tmp, "&#732;", "~", , , vbTextCompare)
tmp = Replace(tmp, "&ensp;", " ", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8194;", " ", , , vbTextCompare)
tmp = Replace(tmp, "&emsp;", " ", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8195;", " ", , , vbTextCompare)
tmp = Replace(tmp, "&thinsp;", " ", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8201;", " ", , , vbTextCompare)
    
tmp = Replace(tmp, "&zwnj;", "", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8204;", "", , , vbTextCompare)
tmp = Replace(tmp, "&zwj;", "", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8205;", "", , , vbTextCompare)
tmp = Replace(tmp, "&lrm;", "", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8206;", "", , , vbTextCompare)
tmp = Replace(tmp, "&rlm;", "", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8207;", "", , , vbTextCompare)
tmp = Replace(tmp, "&ndash;", "-", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8211;", "-", , , vbTextCompare)
tmp = Replace(tmp, "&mdash;", "Ч", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8212;", "Ч", , , vbTextCompare)
    
tmp = Replace(tmp, "&lsquo;", "С", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8216;", "С", , , vbTextCompare)
tmp = Replace(tmp, "&rsquo;", "Т", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8217;", "Т", , , vbTextCompare)
    
tmp = Replace(tmp, "&sbquo;", "В", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8218;", "В", , , vbTextCompare)
    
tmp = Replace(tmp, "&ldquo;", "У", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8220;", "У", , , vbTextCompare)
tmp = Replace(tmp, "&rdquo;", "Ф", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8221;", "Ф", , , vbTextCompare)
    
tmp = Replace(tmp, "&bdquo;", "Д", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8222;", "Д", , , vbTextCompare)
    
tmp = Replace(tmp, "&dagger;", "Ж", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8224;", "Ж", , , vbTextCompare)
tmp = Replace(tmp, "&Dagger;", "З", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8225;", "З", , , vbTextCompare)
    
tmp = Replace(tmp, "&hellip;", "Е", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8230;", "Е", , , vbTextCompare)
tmp = Replace(tmp, "&permil;", "Й", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8240;", "Й", , , vbTextCompare)
    
tmp = Replace(tmp, "&lsaquo;", "Л", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8249;", "Л", , , vbTextCompare)
tmp = Replace(tmp, "&rsaquo;", "Ы", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8250;", "Ы", , , vbTextCompare)
tmp = Replace(tmp, "&euro;", "И", , , vbTextCompare) '
    tmp = Replace(tmp, "&#8364;", "И", , , vbTextCompare)

End If ' есть ли &
''''''''''''''''                                            &&&&&&&&&&&

Do While InStr(1, tmp, vbTab)
tmp = Replace(tmp, vbTab, vbNullString)
Loop

Do While InStr(1, tmp, "  ")
tmp = Replace(tmp, "  ", " ")
Loop

Do While InStr(1, tmp, vbCr)
tmp = Replace(tmp, vbCr, " ")
Loop

Do While InStr(1, tmp, vbLf)
tmp = Replace(tmp, vbLf, " ")
Loop

If Len(tmp) = 1 Then
tmp = Replace(tmp, vbCr, "")
tmp = Replace(tmp, vbLf, "")
End If

If Right$(tmp, 1) = vbCr Then tmp = Left$(tmp, Len(tmp) - 1)

'убрать зап€тые вначале и конце
tmp = Trim$(tmp)
Do While Left$(tmp, 1) = ","
    tmp = Right$(tmp, Len(tmp) - 1)
    tmp = LTrim$(tmp)
Loop
Do While Right$(tmp, 1) = ","
    tmp = Left$(tmp, Len(tmp) - 1)
    tmp = RTrim$(tmp)
Loop


HTML2TEXT = tmp
End Function

'Public Function GetBlockFrom(StartAt As Variant, Optional FinAt As Variant, Optional BeginPosition As Variant) As Variant
Public Function GetBlockFrom(StartAt As Variant, Optional FinAt As Variant) As Variant
'¬озвращает PageText, выреза€ все до подстроки StartAt.
'короче блок текста от и до
'≈сли нет StartAt - равно PageText
Dim StartPos As Long, FinPos As Long ', BeginPos As Long
On Error Resume Next

If StartAt = "" Then
    GetBlockFrom = CVar(PageText): Exit Function
Else

'это символы, надо строки
'    If BeginPosition = "" Then
'        BeginPos = 1
'    Else
'    If IsNumeric(BeginPosition) Then
'        BeginPos = Val(BeginPosition)
'    Else
'        BeginPos = 1
'    End If
'    If BeginPos < 1 Then BeginPos = 1
'    End If
'    StartPos = InStr(BeginPos, PageText, StartAt)

    StartPos = InStr(PageText, StartAt)
    If StartPos > 0 Then
        If FinAt = "" Then
            GetBlockFrom = Right(PageText, Len(PageText) - StartPos + 1)
        Else
            FinPos = InStr(StartPos, PageText, FinAt)
            If FinAt > StartPos Then GetBlockFrom = Mid(PageText, StartPos, FinPos - StartPos)
        End If
    End If
End If
End Function

Public Function UcaseFirst(s As Variant) As Variant
Dim temp As String
Dim i As Integer
On Error Resume Next

s = LTrim(LCase(s))
temp = UCase$(Left$(s, 1)) & Mid$(s, 2)
i = InStr(1, temp, "/")
temp = Left$(temp, i) & UCase$(Mid$(temp, i + 1, 1)) & Mid$(temp, i + 2)
i = InStr(1, temp, "(")
temp = Left$(temp, i) & UCase$(Mid$(temp, i + 1, 1)) & Mid$(temp, i + 2)
UcaseFirst = CVar(temp)
End Function
Public Function ConvStr(s As Variant, c As Variant) As Variant
Dim temp As String
'Dim i As Integer
Dim cc As VbStrConv

On Error Resume Next

cc = CInt(c)
temp = CStr(s)
temp = LTrim$(temp)
temp = StrConv(temp, cc, LCID)
ConvStr = CVar(temp)
End Function
Public Sub AnalyzePage(u As Variant)
Dim i As Integer
'анализ другой страницы
On Error GoTo ErrorHandler
'On Error Resume Next

ToDebug "»нет поиск - редирект"
PageText = OpenURLProxy(u, "txt")
PageArray() = Split(PageText, vbLf)

'Egg
'If frmPeopleFlag Then
'    FrmPeople.List1.Clear
'    FrmPeople.List1.Visible = False
'    For i = LBound(PageArray) To UBound(PageArray)
'        FrmPeople.List1.AddItem i & " |" & PageArray(i)
'    Next i
'    FrmPeople.SetListboxScrollbar FrmPeople.List1
'    FrmPeople.List1.Visible = True
'End If

'скрипт main
SC.Run "AnalyzePage"
If SC.Error.Number <> 0 Then GoTo ErrorHandler

With frmEditor
    .lbInetMovieList.Clear
    For i = 0 To UBound(SC.CodeObject.MTitles)
        .lbInetMovieList.AddItem SC.CodeObject.MTitles(i): FrmMain.SetListboxScrollbar .lbInetMovieList, FrmMain
        .lbInetMovieList.ItemData(i) = SC.CodeObject.MData(i)
    Next i

    .ComInetFind.Enabled = True
    Erase PageArray
End With

Exit Sub
ErrorHandler:

If SC.Error.Number <> 0 Then
    MsgBox "Script Error : " & SC.Error.Number _
           & ": " & SC.Error.Description & " строка " & SC.Error.Line _
           & " колонка " & SC.Error.Column, vbCritical

    Set SC = Nothing
Else
    If err.Number <> 0 Then
        Select Case err.Number
        Case 9
            MsgBox "Nothing found. But may be script errors"
        Case Else
            MsgBox err.Description, vbCritical
        End Select
    End If
End If

End Sub
Public Sub ChangePage(u As Variant)
Dim i As Integer

'On Error GoTo ErrorHandler
On Error Resume Next

ToDebug "»нтернет поиск - смена страницы."
PageText = OpenURLProxy(u, "txt")
PageArray() = Split(PageText, vbLf)

'Egg
'If frmPeopleFlag Then
'    FrmPeople.List1.Clear
'    FrmPeople.List1.Visible = False
'    For i = LBound(PageArray) To UBound(PageArray)
'        FrmPeople.List1.AddItem i & " |" & PageArray(i)
'    Next i
'    FrmPeople.SetListboxScrollbar FrmPeople.List1
'    FrmPeople.List1.Visible = True
'End If

Exit Sub
End Sub
Public Sub PutToSVC()
On Error Resume Next
SetFromScript
End Sub

Public Function myVal(s As Variant)
'оставл€ем только первое встретившеес€ в строке число
Dim i As Long, l As Long
Dim ss As String
ss = Replace(s, "&", "")
ss = Replace(ss, "%", "")
ss = Replace(ss, ",", ".")    'дл€ Val и Isnumeric

l = Len(ss)
For i = 1 To l
    If IsNumeric(Mid$(ss, i, 1)) Then
        myVal = CVar(Val(Right$(ss, l - i + 1)))
        Exit For
    End If
Next i
End Function
