Attribute VB_Name = "ModBrowseFolder"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'common to both methods
Public Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type
Private Enum BrowseFlagsEnum
  'bifBrowseForComputer = &H1000&
  'bifBrowseForPrinter = &H2000&
  'bifBrowseIncludeFiles = &H4000&
  'bifBrowseIncludeURLs = &H80&
  'bifShareable = &H8000&
  'bifDontGoBelowDomain = &H2&
  bifEditBox = &H10&
  'bifReturnFSAncestors = &H8&
  bifReturnOnlyFSDirs = &H1&
  bifStatusText = &H4&
  bifUseNewUI = &H40&
  'bifValidate = &H20&
End Enum

Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
'Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
    
Public Const MAX_PATH = 260
Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1

'Constants ending in 'A' are for Win95 ANSI
'calls; those ending in 'W' are the wide Unicode
'calls for NT.

'Sets the status text to the null-terminated
'string specified by the lParam parameter.
'wParam is ignored and should be set to 0.
Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

'If the lParam  parameter is non-zero, enables the
'OK button, or disables it if lParam is zero.
'(docs erroneously said wParam!)
'wParam is ignored and should be set to 0.
Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)

'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SELECTIONCHANGED
'message.
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
   

'specific to the PIDL method
'Undocumented call for the example. IShellFolder's
'ParseDisplayName member function should be used instead.
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
'specific to the STRING method
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const lPtr = (LMEM_FIXED Or LMEM_ZEROINIT)

'windows-defined type OSVERSIONINFO
'Public Type OSVERSIONINFO
'  OSVSize         As Long
'  dwVerMajor      As Long
'  dwVerMinor      As Long
'  dwBuildNumber   As Long
'  PlatformID      As Long
'  szCSDVersion    As String * 128
'End Type
'Public Const VER_PLATFORM_WIN32_NT = 2
'Public Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long
  
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'опред версии виндовс (winver)
'http://bbs.vbstreets.ru/viewtopic.php?t=19955&highlight=getversionex
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_NT_WORKSTATION = 1
Private Const VER_NT_DOMAIN_CONTROLLER = 2
Private Const VER_NT_SERVER = 3
'Private Const VER_SERVER_NT = &H80000000
'Private Const VER_WORKSTATION_NT = &H40000000
'Private Const VER_SUITE_SMALLBUSINESS = &H1&
Private Const VER_SUITE_ENTERPRISE = &H2&
'Private Const VER_SUITE_BACKOFFICE = &H4&
'Private Const VER_SUITE_COMMUNICATIONS = &H8&
'Private Const VER_SUITE_TERMINAL = &H10&
'Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20&
'Private Const VER_SUITE_EMBEDDEDNT = &H40&
Private Const VER_SUITE_DATACENTER = &H80&
'Private Const VER_SUITE_SINGLEUSERTS = &H100&
Private Const VER_SUITE_PERSONAL = &H200&
'Private Const VER_SUITE_BLADE = &H400&
'  OSVSize         As Long
'  dwVerMajor      As Long
'  dwVerMinor      As Long
'  dwBuildNumber   As Long
'  PlatformID      As Long
'  szCSDVersion    As String * 128
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Integer
    wMinorVersion As Byte '9x-only
    wMajorVersion As Byte '9x-only
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type
Private Type OSVERSIONINFOEX
    osvi As OSVERSIONINFO
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
'''''''''''''''''''''''''''''''''''''


Public Function winver() As String

On Error Resume Next

Dim osvix As OSVERSIONINFOEX
osvix.osvi.dwOSVersionInfoSize = Len(osvix)
If 0 = GetVersionEx(osvix.osvi) Then
    osvix.osvi.dwOSVersionInfoSize = Len(osvix.osvi)
    If 0 = GetVersionEx(osvix.osvi) Then
        winver = "(unknown)"
        Exit Function
    End If
End If

Select Case osvix.osvi.dwPlatformId
Case VER_PLATFORM_WIN32_NT:
    Select Case osvix.osvi.dwMajorVersion
    Case 3, 4: winver = "Windows NT"
    Case 5:
        Select Case osvix.osvi.dwMinorVersion
        Case 0: winver = "Windows 2000"
        Case 1: winver = "Windows XP"
        Case 2: winver = "Windows 2003"
        Case Else:
            winver = "(Windows NT 5.x)"
        End Select
    Case 6: winver = "Windows Vista"
    Case Else: winver = "(Windows NT)"
    End Select

    If Len(osvix) = osvix.osvi.dwOSVersionInfoSize Then
        Select Case osvix.wProductType
        Case VER_NT_SERVER, VER_NT_DOMAIN_CONTROLLER:
            If osvix.wSuiteMask And VER_SUITE_DATACENTER Then
                winver = winver & " Datacenter"
            ElseIf osvix.wSuiteMask And VER_SUITE_ENTERPRISE Then
                winver = winver & " Advanced Server"
            Else
                winver = winver & " Server"
            End If
        Case VER_NT_WORKSTATION:
            If osvix.osvi.dwMajorVersion < 5 Then
                winver = winver & " Workstation"
            ElseIf osvix.wSuiteMask And VER_SUITE_PERSONAL Then
                winver = winver & " Home"
            Else
                winver = winver & " Professional"
            End If
        Case Else:
            'no append
        End Select
    End If
Case VER_PLATFORM_WIN32_WINDOWS:
    Select Case osvix.osvi.dwMajorVersion
    Case 4:
        Select Case osvix.osvi.dwMinorVersion
        Case 0: winver = "Windows 95"
        Case 10: winver = "Windows 98"
        Case 90: winver = "Windows Me"
        Case Else: winver = "(Windows 9x)"
        End Select
    Case Else: winver = "(Windows)"
    End Select
Case Else:
    winver = "(unknown)"
End Select

winver = winver & ", ver. " & osvix.osvi.dwMajorVersion & "." & osvix.osvi.dwMinorVersion & "." & osvix.osvi.dwBuildNumber
If Len(osvix) = osvix.osvi.dwOSVersionInfoSize Then
    If osvix.wServicePackMajor Then
        winver = winver & " SP" & osvix.wServicePackMajor
        If osvix.wServicePackMinor Then
            winver = winver & "." & osvix.wServicePackMinor
        End If
    End If
End If
End Function



Public Function BrowseCallbackProcStr(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

'Callback for the Browse STRING method.

'On initialization, set the dialog's
'pre-selected folder from the pointer
'to the path allocated as bi.lParam,
'passed back to the callback as lpData param.

Select Case uMsg
Case BFFM_INITIALIZED
    Call SendMessage(hwnd, BFFM_SETSELECTIONA, 1&, ByVal lpData)
End Select
End Function


'Public Function BrowseCallbackProc(ByVal hWnd As Long, _
 '                                   ByVal uMsg As Long, _
 '                                   ByVal lParam As Long, _
 '                                   ByVal lpData As Long) As Long
'
'  'Callback for the Browse PIDL method.
'
'  'On initialization, set the dialog's
'  'pre-selected folder using the pidl
'  'set as the bi.lParam, and passed back
'  'to the callback as lpData param.
'   Select Case uMsg
'      Case BFFM_INITIALIZED
'         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
          '                          0&, ByVal lpData)
'         Case Else:
'   End Select
'End Function


Public Function FARPROC(pfn As Long) As Long
'A dummy procedure that receives and returns
'the value of the AddressOf operator.

'This workaround is needed as you can't
'assign AddressOf directly to a member of a
'user-defined type, but you can assign it
'to another long and use that instead!
FARPROC = pfn
End Function

Public Function BrowseForFolderByPath(sSelPath As String, Nazvanie As String, Optional hWndOwner As Long) As String

Dim bi As BROWSEINFO
Dim pidl As Long
Dim lpSelPath As Long
Dim spath As String * MAX_PATH

With bi
    If hWndOwner = 0 Then
        .hOwner = FrmMain.hwnd
    Else
        .hOwner = hWndOwner
    End If
    .iImage = 0
    .pidlRoot = 0
    .pszDisplayName = String$(MAX_PATH, Chr$(0))
    If Len(Nazvanie) > 0 Then .lpszTitle = Nazvanie
    .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)

    'выбрать заданную папку
    If Len(sSelPath) > 0 Then
        If DirExists(sSelPath) Then
            lpSelPath = LocalAlloc(lPtr, Len(sSelPath) + 1)
            CopyMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath) + 1
            .lParam = lpSelPath
        End If
    End If

    '.lParam = 0

    'двинутый
    .ulFlags = bifReturnOnlyFSDirs Or bifUseNewUI    'Or bifEditBox
    'нет, грузятся формы If .hOwner = frmAuto.hwnd Or .hOwner = FrmOptions.hwnd Then
    If frmAutoFlag Then
        If .hOwner = frmAuto.hwnd Then .ulFlags = bifReturnOnlyFSDirs       'простой
    End If
    If frmOptFlag Then
        If .hOwner = FrmOptions.hwnd Then .ulFlags = bifReturnOnlyFSDirs        'простой
    End If
End With

pidl = SHBrowseForFolder(bi)

If pidl Then
    If SHGetPathFromIDList(pidl, spath) Then
        BrowseForFolderByPath = left$(spath, InStr(spath, vbNullChar) - 1)
    Else
        BrowseForFolderByPath = vbNullString
    End If

    Call CoTaskMemFree(pidl)

Else
    BrowseForFolderByPath = vbNullString
End If

Call LocalFree(lpSelPath)

End Function


'Public Function BrowseForFolderByPIDL(sSelPath As String) As String
'
'Dim BI As BROWSEINFO
'Dim pidl As Long
'Dim spath As String * MAX_PATH
'
'With BI
'    .hOwner = FrmMain.hWnd
'    .pidlRoot = 0
'
'    .lpszTitle = NamesStore(4)   'sur
'    .lParam = GetPIDLFromPath(sSelPath)
'    .ulFlags = bifUseNewUI    'bifReturnOnlyFSDirs 'Or bifEditBox
'
'    .lpfn = FARPROC(AddressOf BrowseCallbackProc)
'End With
'
'pidl = SHBrowseForFolder(BI)
'
'If pidl Then
'    If SHGetPathFromIDList(pidl, spath) Then
'        BrowseForFolderByPIDL = left$(spath, InStr(spath, vbNullChar) - 1)
'    Else
'        BrowseForFolderByPIDL = vbNullString
'    End If
'
'    'free the pidl from SHBrowseForFolder call
'    Call CoTaskMemFree(pidl)
'Else
'    BrowseForFolderByPIDL = vbNullString
'End If
'
''free the pidl (lparam) from GetPIDLFromPath call
'Call CoTaskMemFree(BI.lParam)
'
'End Function

'Private Function GetPIDLFromPath(spath As String) As Long
''return the pidl to the path supplied by calling the
''undocumented API #162 (our name for this undocumented
''function is "SHSimpleIDListFromPath").
''This function is necessary as, unlike documented APIs,
''the API is not implemented in 'A' or 'W' versions.
'
'If IsWinNT() Then
'    GetPIDLFromPath = SHSimpleIDListFromPath(StrConv(spath, vbUnicode, LCID))
'Else
'    GetPIDLFromPath = SHSimpleIDListFromPath(spath)
'End If
'
'End Function

'Private Function IsWinNT() As Boolean
'   #If Win32 Then
'      Dim OSV As OSVERSIONINFO
'      'OSV.OSVSize = Len(OSV)
'      OSV.dwOSVersionInfoSize = Len(OSV)
'     'API returns 1 if a successful call
'      If GetVersionEx(OSV) = 1 Then
'
'        'PlatformId contains a value representing
'        'the OS; if VER_PLATFORM_WIN32_NT,
'        'return true
'         'IsWinNT = OSV.PlatformID = VER_PLATFORM_WIN32_NT
'         IsWinNT = OSV.dwPlatformId = VER_PLATFORM_WIN32_NT
'
'      End If
'   #End If
'End Function


Public Function IsValidDrive(spath As String) As Boolean

Dim buff As String
Dim nBuffsize As Long

'Call the API with a buffer size of 0.
'The call fails, and the required size
'is returned as the result.
nBuffsize = GetLogicalDriveStrings(0&, buff)

'pad a buffer to hold the results
buff = Space$(nBuffsize)
nBuffsize = Len(buff)

'and call again
If GetLogicalDriveStrings(nBuffsize, buff) Then

    'if the drive letter passed is in
    'the returned logical drive string,
    'return True.
    IsValidDrive = InStr(1, buff, spath, vbTextCompare) > 0

End If

End Function


Public Function FixPath(spath As String) As String

  'The Browse callback requires the path string
  'in a specific format - trailing slash if a
  'drive only, or minus a trailing slash if a
  'file system path. This routine assures the
  'string is formatted correctly.
  '
  'In addition, because the calls to LocalAlloc
  'requires a valid path for the call to succeed,
  'the path defaults to C:\ if the passed string
  'is empty.
  
  'Test 1: check for empty string. Since
  'we're setting it we can assure it is
  'formatted correctly, so can bail.
   If Len(spath) = 0 Then
      FixPath = "C:\"
      Exit Function
   End If
   
  'Test 2: is path a valid drive?
  'If this far we did not set the path,
  'so need further tests. Here we ensure
  'the path is properly terminated with
  'a trailing slash as needed.
  '
  'Drives alone require the trailing slash;
  'file system paths must have it removed.
   If IsValidDrive(spath) Then
      
     'IsValidDrive only determines if the
     'path provided is contained in
     'GetLogicalDriveStrings. Since
     'IsValidDrive() will return True
     'if either C: or C:\ is passed, we
     'need to ensure the string is formatted
     'with the trailing slash.
      FixPath = QualifyPath(spath)
   Else
     'The string passed was not a drive, so
     'assume it's a path and ensure it does
     'not have a trailing space.
      FixPath = UnqualifyPath(spath)
   End If
   
End Function


Private Function QualifyPath(spath As String) As String
If Len(spath) > 0 Then
    If right$(spath, 1) <> "\" Then
        QualifyPath = spath & "\"
    Else
        QualifyPath = spath
    End If
Else
    QualifyPath = vbNullString
End If
End Function


Private Function UnqualifyPath(spath As String) As String
'Qualifying a path involves assuring that its format
'is valid, including a trailing slash, ready for a
'filename. Since SHBrowseForFolder will not pre-select
'the path if it contains the trailing slash, it must be
'removed, hence 'unqualifying' the path.
If Len(spath) > 0 Then
    If right$(spath, 1) = "\" Then
        UnqualifyPath = left$(spath, Len(spath) - 1)
        Exit Function
    End If
End If
UnqualifyPath = spath

End Function

Private Function DirExists(sdir As String) As Boolean
On Error Resume Next
DirExists = Len(Dir(sdir, vbDirectory))
End Function
