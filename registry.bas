Attribute VB_Name = "Reg"
'Íàéòè ïðîêñè â ðååñòðå

'Dieser Source stammt von http://www.ActiveVB.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Sollten Sie Fehler entdecken oder Fragen haben, dann
'mailen Sie mir bitte unter: Reinecke@ActiveVB.de
'Ansonsten viel Spaß und Erfolg mit diesem Source !
'**************************************************************

' Registry Deklarationen
Option Explicit

Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal _
        lpSubKey As String, ByVal ulOptions As Long, ByVal _
        samDesired As Long, phkResult As Long) As Long
        
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
        
Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Any) As Long
        
Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal _
        lpSubKey As String, ByVal Reserved As Long, ByVal _
        lpClass As String, ByVal dwOptions As Long, ByVal _
        samDesired As Long, ByVal lpSecurityAttributes As Any, _
        phkResult As Long, lpdwDisposition As Long) As Long
        
Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
        
Declare Function RegSetValueEx_String Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
        
Declare Function RegSetValueEx_DWord Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, lpData As Long, ByVal cbData As Long) As Long
        
Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
        "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
        
Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
        "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const HKEY_PERFORMANCE_DATA = &H80000004
Global Const HKEY_CURRENT_CONFIG = &H80000005
Global Const HKEY_DYN_DATA = &H80000006

Global Const KEY_QUERY_VALUE = &H1
Global Const KEY_SET_VALUE = &H2
Global Const KEY_CReatE_SUB_KEY = &H4
Global Const KEY_ENUMERATE_SUB_KEYS = &H8
Global Const KEY_NOTIFY = &H10
Global Const KEY_CReatE_LINK = &H20
Global Const KEY_READ = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
                 
Global Const KEY_ALL_ACCESS = _
             KEY_QUERY_VALUE Or _
             KEY_SET_VALUE Or _
             KEY_CReatE_SUB_KEY Or _
             KEY_ENUMERATE_SUB_KEYS Or _
             KEY_NOTIFY Or _
             KEY_CReatE_LINK
                       
Global Const ERROR_SUCCESS = 0&

Global Const REG_NONE = 0
Global Const REG_SZ = 1
Global Const REG_EXPAND_SZ = 2
Global Const REG_BINARY = 3
Global Const REG_DWORD = 4
Global Const REG_DWORD_LITTLE_ENDIAN = 4
Global Const REG_DWORD_BIG_ENDIAN = 5
Global Const REG_LINK = 6
Global Const REG_MULTI_SZ = 7

Global Const REG_OPTION_NON_VOLATILE = &H0

Function RegKeyExist(ByVal Root As Long, Key$) As Long

Dim result As Long

Dim hKey As Long
'Prüfen ob ein Schlüssel existiert
result = RegOpenKeyEx(Root, Key$, 0, KEY_READ, hKey)
If result = ERROR_SUCCESS Then
    Call RegCloseKey(hKey)
End If
RegKeyExist = result

End Function

'Function RegKeyCreate(ByVal Root As Long, Newkey$) As Long
'
'Dim result&, hKey&, Back&
''Neuen Schlüssel erstellen
'result = RegCreateKeyEx(Root, Newkey$, 0, vbNullString, _
'        REG_OPTION_NON_VOLATILE, _
'        KEY_ALL_ACCESS, 0&, hKey, Back)
'If result = ERROR_SUCCESS Then
'    result = RegFlushKey(hKey)
'    If result = ERROR_SUCCESS Then
'        Call RegCloseKey(hKey)
'    End If
'    RegKeyCreate = Back
'End If
'End Function

'Private Function RegKeyDelete(Root&, Key$) As Long
'  'Schlüssel erstellen
'  RegKeyDelete = RegDeleteKey(Root, Key)
'End Function

'Private Function RegFieldDelete(ByVal Root As Long, ByVal Key$, Field$) As Long
'    Dim result As Long
'    Dim hKey As Long
'    'Feld löschen
'    result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, hKey)
'    If result = ERROR_SUCCESS Then
'        result = RegDeleteValue(hKey, Field)
'        result = RegCloseKey(hKey)
'    End If
'    RegFieldDelete = result
'End Function

'Function RegValueSet(ByVal Root As Long, Key$, Field$, Value As Variant) As Long
'Dim result&, hKey&, s$, l&
''Wert in ein Feld der Registry schreiben
'result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, hKey)
'If result = ERROR_SUCCESS Then
'    Select Case VarType(Value)
'    Case vbInteger, vbLong
'        l = CLng(Value)
'        result = RegSetValueEx_DWord(hKey, Field$, 0, REG_DWORD, l, 4)
'    Case vbString
'        s = CStr(Value)
'        result = RegSetValueEx_String(hKey, Field$, 0, REG_SZ, s, Len(s) + 1)
'    End Select
'    result = RegCloseKey(hKey)
'End If
'RegValueSet = result
'End Function

Function RegValueGet(ByVal Root As Long, Key$, Field$, Value As Variant) As Long
' return value is passed back in variable 'value'
' function return is error value
Dim result&, hKey&, dwType&, Lng&, Buffer$, l&, pos
'Wert aus einem Feld der Registry auslesen
result = RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey)
' Reg  Open creates a handle (similar to brush or font handle)
' field$ determines the parameter to be read
If result = ERROR_SUCCESS Then
    result = RegQueryValueEx(hKey, Field$, 0&, dwType, ByVal 0&, l)
    ' l receives the length
    ' dwType receives 1 in case of string
    ' result is error value
    ' seems setting 0& instead of buffer is used as a dummy just to
    ' determine length before actual reading
    ' Now the value can actually be read (what a Krampf)
    If result = ERROR_SUCCESS Then
        Select Case dwType
        Case REG_SZ
            Buffer = Space$(l + 1)
            result = RegQueryValueEx(hKey, Field$, 0&, dwType, ByVal Buffer, l)
            If result = ERROR_SUCCESS Then
                pos = InStr(1, Buffer$, Chr$(0), 1)     ' this is just for safety
                If pos Then
                    Buffer$ = Left$(Buffer$, pos - 1)
                End If
                Value = Buffer$
            End If
            
        Case REG_DWORD, REG_BINARY
            result = RegQueryValueEx(hKey, Field$, 0&, dwType, Lng, l)
            If result = ERROR_SUCCESS Then
                Value = Lng
            End If
        End Select
    End If
End If

If result = ERROR_SUCCESS Then
    result = RegCloseKey(hKey)
End If

RegValueGet = result
End Function


'Function NextChar(Text$, char$) As String
'    Dim pos As Long
'    pos = InStr(1, Text$, char$)
'    If pos = 0 Then
'      NextChar = Text$
'      Text$ = ""
'    Else
'      NextChar = Left$(Text$, pos - 1)
'      Text$ = Mid$(Text, pos + Len(char$))
'    End If
'End Function



