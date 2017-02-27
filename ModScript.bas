Attribute VB_Name = "ModScript"
Option Explicit
'script
Public SC As Object
Public objScript As New ClsScript
Public POST_flag As Boolean 'тру если надо метод ПОСТ, работает не при прямом моединении

'' flags for InternetOpen():
'с асинхом трудности
'http://support.microsoft.com/kb/189850
'Public Const INTERNET_FLAG_ASYNC = &H10000000              ' this request is asynchronous (where supported)
'Public Const INTERNET_NO_CALLBACK = 0

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''wininet.bas
' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
'Public Const scUserAgent = "http sample"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000

' Security constants
Public Const INTERNET_OPTION_SECURITY_FLAGS = 31
Public Const SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
Public Const INTERNET_FLAG_SECURE = &H800000


' Sends the specified request to the HTTP server.
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_STATUS_DENIED = 401
Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

Private hInternetSession As Long
Private hInternetConnect As Long
Private hHttpOpenRequest As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''end wininet.bas

Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef _
    lpSFlags As Long, ByVal dwReserved As Long) As Long
Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8



Public Const FLAG_ICC_FORCE_CONNECTION = &H1
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" _
    (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long


Public Const NETWORK_ALIVE_AOL = &H4
Public Const NETWORK_ALIVE_LAN = &H1
Public Const NETWORK_ALIVE_WAN = &H2
Public Type QOCINFO
    dwSize As Long
    dwFlags As Long
    dwInSpeed As Long 'in bytes/second
    dwOutSpeed As Long 'in bytes/second
End Type
Public Declare Function IsDestinationReachable Lib "SENSAPI.DLL" Alias "IsDestinationReachableA" (ByVal lpszDestination As String, ByRef lpQOCInfo As QOCINFO) As Long



'Public Const scUserAgent = "Internet Explore 5.x"
Public Const scUserAgent = "mozilla /4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
Public sReferer As String 'а откуда вы? хттп+домен/

'Public Const INTERNET_FLAG_RELOAD = &H80000000

'Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
'Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
'Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long


Public Declare Function InternetSetStatusCallback Lib "wininet.dll" Alias "InternetSetStatusCallbackA" (ByVal hInternet As Long, ByVal CallbackFunctionAddress As Long) As Long
Public Declare Function StringFromPointer Lib "KERNEL32" Alias "lstrcpy" (ByVal lpDestinationStr As String, ByVal lpStringPointer As Long) As Long
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByVal lpszBuffer As String, ByRef lpdwBufferLength As Long) As Long 'BOOL

Public Declare Function InternetAttemptConnect Lib "wininet" (ByVal dwReserved As Long) As Long
'Public Const FLAG_ICC_FORCE_CONNECTION = &H1
'Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" _
    (ByVal lpszUrl As String)
'Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
    (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Long
     
'Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
'Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
'Public Const INTERNET_OPTION_SEND_TIMEOUT = 5
Public Const INTERNET_OPTION_CONNECT_RETRIES = 3

'Public Const INTERNET_SERVICE_HTTP = 3
'Public Const INTERNET_DEFAULT_HTTP_PORT = 80

'Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
'
Public PageArray() As String  'содержимое страницы по строкам
Public PageText As String  'содержимое страницы целиком
Public BaseAddress As String
'Public URLTitleArr() As String  'хранит URLs названий найденных фильмов (lbInetMovieList)
Public url As String  ' текущая обрабатываемая страница

'PictureFromByteStream
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Public Function LoadFile(ByVal FileName As String) As Byte()
'    Dim FileNo As Integer, b() As Byte
'    On Error GoTo Err_Init
'    If Dir(FileName, vbNormal Or vbArchive) = vbnullstring Then
'        Exit Function
'    End If
'    FileNo = FreeFile
'    Open FileName For Binary Access Read As #FileNo
'    ReDim b(0 To LOF(FileNo) - 1)
'    Get #FileNo, , b
'    Close #FileNo
'    LoadFile = b
'    Exit Function
'Err_Init:
'    MsgBox err.Number & " - " & err.Description
'End Function

Public Function PictureFromByteStream(b() As Byte) As IPicture
    Dim LowerBound As Long
    Dim ByteCount  As Long
    Dim hMem  As Long
    Dim lpMem  As Long
    Dim IID_IPicture(15)
    Dim istm As stdole.IUnknown

    On Error GoTo Err_Init
    If UBound(b, 1) < 0 Then
        Exit Function
    End If
    
    LowerBound = LBound(b)
    ByteCount = (UBound(b) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, b(LowerBound), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                  Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                End If
            End If
        End If
    End If
    
    Exit Function
    
Err_Init:
    If err.Number = 9 Then
        'Uninitialized array
        ToDebug "Err_PFBS: empty byte array"
        'You must pass a non-empty byte array to this function!"
    Else
        ToDebug "Err_PFBS: " & err.Number & " - " & err.Description
    End If
End Function


'Public Sub ClearTextFields()
'On Error Resume Next
'
'With FrmMain
'
'.TextGenre = vbNullString
'.TextCountry = vbNullString
'.TextYear = vbNullString
'.TextAuthor = vbNullString
'.TextRole = vbNullString
'.TextAnnotation = vbNullString
'
'
'End With
'End Sub

'Public Function UrlEncode(ByVal urlText As String) As String
'Dim i As Long
'Dim ansi() As Byte
'Dim ascii As Integer
'Dim encText As String
'
'ansi = StrConv(urlText, vbFromUnicode, LCID)
'
'encText = vbNullString
'For i = 0 To UBound(ansi)
'    ascii = ansi(i)
'    Select Case ascii
'    Case 48 To 57, 65 To 90, 97 To 122
'        encText = encText & Chr(ascii)
'    Case 32
'        encText = encText & "+"
'    Case Else
'        If ascii < 16 Then
'            encText = encText & "%0" & Hex(ascii)
'        Else
'            encText = encText & "%" & Hex(ascii)
'        End If
'    End Select
'Next i
'UrlEncode = encText
'End Function

'Public Function CheckInet(u As String) As Boolean
'Dim ret As QOCINFO
''mzt Dim ErrDesc As String
'Dim R As Boolean
'ret.dwSize = Len(ret)
'R = IsDestinationReachable(u, ret)
''GetLastErr_Msg , , , ErrDesc
''Debug.Print ErrDesc
'If R = 0 Then
'    myMsgBox msgsvc(33), vbCritical
'    ToDebug "Нет связи с хостом."
'    CheckInet = False
'Else
'    CheckInet = True
'    ToDebug "Скорость приема: " & Format$(ret.dwInSpeed / 1024, "#.0") + " Kb/s,"
'    ToDebug "Скорость передачи: " & Format$(ret.dwOutSpeed / 1024, "#.0") + " Kb/s."
'End If
'End Function

Public Function HasSpecialChar(ByVal strText As String) As Boolean
    HasSpecialChar = Not (Len(strText) = LenB(StrConv(strText, vbFromUnicode, LCID)))
End Function
Public Function RightStrConv(s As String) As Boolean
'true если преобразование выполняется верно
If StrConv(StrConv(s, vbFromUnicode, LCID), vbUnicode, LCID) = s Then RightStrConv = True
End Function

Public Function OpenURLProxy(ByVal sUrl As String, param As String) As String
'param - "txt" или "pic"
'если pic - сразу в нужную картинку (по флагу)

Dim bDoLoop As Boolean
Dim sReadBuffer As String * 2048    'важно определить размер буфера
Dim lNumberOfBytesRead As Long
Dim sBuffer As String
Dim b() As Byte
'Dim LocaleId As Long

'On Error Resume Next

If frmEditor.FrameAddEdit.Visible Then frmEditor.ComInetFind.Enabled = False

InetOpen

If SendRequest(sUrl) Then
    Screen.MousePointer = vbHourglass
    bDoLoop = True
    'Debug.Print "InternetReadFile"
    While bDoLoop
        sReadBuffer = vbNullString
        bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    'Debug.Print "Ready"

    On Error Resume Next
                    
    If Len(sBuffer) <> 0 Then
        If param = "txt" Then    'возврат страницы
            OpenURLProxy = sBuffer
        Else    '                 получили pic
        
'Debug.Print HasSpecialChar(sBuffer)
        
        ToDebug "strconv_ok = " & RightStrConv(sBuffer)
        
            b = StrConv(sBuffer, vbFromUnicode, LCID)

            'Opt_InetGetPicUseTempFile = True 'в опции
            If Opt_InetGetPicUseTempFile Then    'через темп файл
                Dim ff As Integer
                Dim strTempPath As String
                strTempPath = Space$(260)
                GetTempPath 260, strTempPath
                If InStr(strTempPath, vbNullChar) > 0 Then strTempPath = Left$(strTempPath, InStr(strTempPath, vbNullChar) - 1)
                strTempPath = strTempPath & "svc_temp.bmp"
                ff = FreeFile
                Open strTempPath For Binary As #ff
                Put #ff, , b
                Close #ff
                'Set URLPicture = LoadPicture(strTempPath)
            End If

            With frmEditor
                If getPeopleFlag Then    'для актеров (frmpeople)
                    If Opt_InetGetPicUseTempFile Then
                        Set FrmPeople.PicFaceA.Picture = LoadPicture(strTempPath)    'через темп файл
                    Else
                        Set FrmPeople.PicFaceA.Picture = PictureFromByteStream(b)    ' из потока
                    End If
                    getPeopleFlag = False
                Else

                    Set .ImgPrCov = Nothing

                    If Opt_InetGetPicUseTempFile Then
                        Set .ImgPrCov.Picture = LoadPicture(strTempPath)    'через темп файл
                    Else
                        Set .ImgPrCov.Picture = PictureFromByteStream(b)    ' из потока
                    End If
                    If .ImgPrCov.Picture <> 0 Then
                        Set .PicFrontFace = Nothing
                        Set .picCanvas = Nothing
                        NoPicFrontFaceFlag = False
                        DrDroFlag = True
                        .PicFrontFace.Picture = .ImgPrCov.Picture
                        DrawCoverEdit
                        Mark2Save
                        SaveCoverFlag = True
                        ToDebug "InetGetCover = True"
                    Else
                        Set .ImgPrCov = Nothing: Set .PicFrontFace = Nothing: Set .picCanvas = Nothing
                        NoPicFrontFaceFlag = True
                        Mark2Save
                        SaveCoverFlag = True
                        ToDebug "InetGetCover = False"
                    End If
                End If
            End With


        End If    'txt or pic
    End If    ' sbuffer
End If    'SendRequest

err.Clear

InetClose

Screen.MousePointer = vbDefault
frmEditor.ComInetFind.Enabled = True
End Function

'Sub MyCallBack( _
'      ByVal hInternet As Long, _
'      ByVal dwContext As Long, _
'      ByVal dwInternetStatus As Long, _
'      ByVal pbStatusInformation As Long, _
'      ByVal dwStatusInformationLength As Long)
'
'      ' Callback routine implementation.
'End Sub

Private Sub InetOpen()
'Debug.Print "InternetOpen"
On Error Resume Next

Select Case Opt_InetUseProxy
Case 0 'no
hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
Case 1 'IE
hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
Case 2 'My
hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PROXY, Opt_InetProxyServerPort, "<local>", 0)
'hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PROXY, Opt_InetProxyServerPort, "<local>", INTERNET_FLAG_ASYNC)
'InternetSetStatusCallback hInternetSession, AddressOf MyCallBack
End Select

If CBool(hInternetSession) Then
'    Debug.Print "Ready"
ToDebug "InetOpen: Ok"
Else
'    Debug.Print "InternetOpen failed."
ToDebug "InetOpen: Error"
End If
End Sub
Private Sub InetClose()
On Error Resume Next
InternetCloseHandle (hHttpOpenRequest)
InternetCloseHandle (hInternetSession)
InternetCloseHandle (hInternetConnect)
End Sub

Private Function SendRequest(TxtURL As String) As Boolean
ToDebug "Входной URL = " & TxtURL

If Opt_InetUseProxy = 0 Then    'direct
    'InternetOpenUrl - тут только GET
    'и никакого реферера ? :(
    hHttpOpenRequest = InternetOpenUrl(hInternetSession, TxtURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    If hHttpOpenRequest Then
        SendRequest = True
        ToDebug "DirectSendRequest = True"
    Else
        ToDebug "DirectSendRequest = False"
    End If
    Exit Function
End If

'''''''''''''''''''''''''''''''''прокси''''''''
Dim iRetVal As Integer
Dim sBuffer As String * 1024
Dim lBufferLen As Long
Dim vDllVersion As tWinInetDLLVersion
Dim sStatus As String
Dim sOptionBuffer As String
Dim lOptionBufferLen As Long
Dim SecFlag As Long
Dim dwSecFlag As Long
Dim dwPort As Long

Dim tmp As String
'Dim tempCap As String 'ComInetFind.Caption

Dim dwTimeOut As Long
dwTimeOut = 20000    ' time out

Screen.MousePointer = vbHourglass


TxtURL = CheckUrlRev(TxtURL)
'Debug.Print "URL без http = " & TxtURL

lBufferLen = Len(sBuffer)
If CBool(hInternetSession) Then
    'Debug.Print "InternetQueryOption"
    InternetQueryOption hInternetSession, INTERNET_OPTION_VERSION, vDllVersion, Len(vDllVersion)
    'Debug.Print "lblMajor = " & vDllVersion.lMajorVersion
    'Debug.Print "lblMinor = " & vDllVersion.lMinorVersion


    'If checkSecure.Value = 1 Then
    If Opt_InetSecureFlag Then
        'Debug.Print "Establishing secure connection" & " "
        dwPort = INTERNET_DEFAULT_HTTPS_PORT
        'Debug.Print "Setting security flags" & " "
        SecFlag = INTERNET_FLAG_SECURE Or _
                  INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
                  INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
    Else
        dwPort = INTERNET_DEFAULT_HTTP_PORT
        SecFlag = 0
    End If
    'hInternetConnect = InternetConnect(hInternetSession, CheckUrl, dwPort, _
     txtUsername.Text, txtPassword.Text, INTERNET_SERVICE_HTTP, 0, 0)
    tmp = CheckUrl(TxtURL)
    'Debug.Print "CheckUrl = " & tmp
    hInternetConnect = InternetConnect(hInternetSession, tmp, dwPort, _
                                       Opt_InetUserName, Opt_InetPassword, INTERNET_SERVICE_HTTP, 0, 0)
    If hInternetConnect > 0 Then
        ToDebug "InternetConnect: Ok"

        'Get или пост
        tmp = GetUrlObject(TxtURL)
        
        If POST_flag Then
            sOptionBuffer = "MOVENEXT=MoveNext?CONTRACT=-1&ZIP=94025?STATE=CA"
            lOptionBufferLen = Len(sOptionBuffer)
            ToDebug "POST..."
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "POST", tmp, "HTTP/1.0", sReferer, 0, _
                                               INTERNET_FLAG_RELOAD Or SecFlag, 0)
        Else 'GET (в основном)
        
            sOptionBuffer = vbNullString
            lOptionBufferLen = 0
            'tmp = GetUrlObject(TxtURL)
            'Debug.Print "GetUrlObject = " & tmp
            ToDebug "GET..."
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", tmp, "HTTP/1.0", sReferer, 0, _
                                               INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION Or SecFlag, 0)
        End If

        If CBool(hHttpOpenRequest) Then
            ToDebug "HttpSendRequest: Ok"
            'Debug.Print sOptionBuffer
            Dim sHeader As String

            'sHeader = "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd." & vbCrLf
            'iRetVal = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
            'Debug.Print iRetVal & " " & Len(sHeader)

            sHeader = "Accept-Language: en" & vbCrLf
            iRetVal = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
            'Debug.Print iRetVal & " " & Len(sHeader)

            sHeader = "Connection: Keep-Alive" & vbCrLf
            iRetVal = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
            'Debug.Print iRetVal & " " & Len(sHeader);

            'sHeader = "Content-Type: text/html" & vbCrLf ' "Accept = image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd." & vbCrLf
            'iRetVal = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
            'Debug.Print iRetVal & " " & Len(sHeader)


            iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_CONNECT_TIMEOUT, dwTimeOut, 4)
            'Debug.Print iRetVal & " " & err.LastDllError & " " & "INTERNET_OPTION_CONNECT_TIMEOUT"
            iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, dwTimeOut, 4)
            'Debug.Print iRetVal & " " & "INTERNET_OPTION_RECEIVE_TIMEOUT"
            iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SEND_TIMEOUT, dwTimeOut, 4)
            'Debug.Print iRetVal & " " & "INTERNET_OPTION_SEND_TIMEOUT"

Resend:
            iRetVal = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, sOptionBuffer, lOptionBufferLen)

            If (iRetVal <> 1) And (err.LastDllError = 12045) Then
                'ToDebug "Resend after Invalid CA"
                'Certificate Authority is invalid.
                'Debug.Print "Invalid Cert Auth, resending" & " "
                dwSecFlag = SECURITY_FLAG_IGNORE_UNKNOWN_CA
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SECURITY_FLAGS, dwSecFlag, 4)
                'Debug.Print iRetVal & " " & err.LastDllError & " " & "INTERNET_OPTION_SECURITY_FLAGS"
                GoTo Resend
            End If

            If iRetVal Then
                Dim dwStatus As Long, dwStatusSize As Long
                dwStatusSize = Len(dwStatus)
                HttpQueryInfo hHttpOpenRequest, HTTP_QUERY_FLAG_NUMBER Or HTTP_QUERY_STATUS_CODE, dwStatus, dwStatusSize, 0
                Select Case dwStatus
                Case HTTP_STATUS_PROXY_AUTH_REQ
                    iRetVal = InternetSetOptionStr(hHttpOpenRequest, INTERNET_OPTION_PROXY_USERNAME, _
                                                   "IUSR_WEIHUA1", Len("IUSR_WEIHUA1") + 1)
                    iRetVal = InternetSetOptionStr(hHttpOpenRequest, INTERNET_OPTION_PROXY_PASSWORD, _
                                                   "IUSR_WEIHUA1", Len("IUSR_WEIHUA1") + 1)
                    'ToDebug "Resent after AUTH_REQ"
                    GoTo Resend
                Case HTTP_STATUS_DENIED
                    iRetVal = InternetSetOptionStr(hHttpOpenRequest, INTERNET_OPTION_USERNAME, _
                                                   "IUSR_WEIHUA1", Len("IUSR_WEIHUA1") + 1)
                    iRetVal = InternetSetOptionStr(hHttpOpenRequest, INTERNET_OPTION_PASSWORD, _
                                                   "IUSR_WEIHUA1", Len("IUSR_WEIHUA1") + 1)
                    'ToDebug "Resent after DENIED"
                    GoTo Resend
                End Select

                'Debug.Print "HttpQueryInfo"
                'response headers
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_CONTENT_TYPE
                'Debug.Print "HTTP_QUERY_CONTENT_TYPE = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_CONTENT_LENGTH
                'Debug.Print "HTTP_QUERY_CONTENT_LENGTH = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_LAST_MODIFIED
                'Debug.Print "HTTP_QUERY_LAST_MODIFIED = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_VERSION
                'Debug.Print "HTTP_QUERY_VERSION = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_STATUS_CODE
                'Debug.Print "HTTP_QUERY_STATUS_CODE = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_STATUS_TEXT
                'Debug.Print "HTTP_QUERY_STATUS_TEXT = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_RAW_HEADERS
                'Debug.Print "HTTP_QUERY_RAW_HEADERS = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_RAW_HEADERS_CRLF
                'Debug.Print "HTTP_QUERY_RAW_HEADERS_CRLF = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_FORWARDED
                'Debug.Print "HTTP_QUERY_FORWARDED = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_SERVER
                'Debug.Print "HTTP_QUERY_SERVER = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_REQUEST_METHOD
                'Debug.Print "HTTP_QUERY_REQUEST_METHOD = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_PRAGMA
                'Debug.Print "HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_PRAGMA = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_RAW_HEADERS_CRLF
                'Debug.Print "HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_RAW_HEADERS_CRLF = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_USER_AGENT
                'Debug.Print "HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_USER_AGENT = " & tmp
                '                GetQueryInfo hHttpOpenRequest, tmp, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_REQUEST_METHOD
                'Debug.Print "HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_REQUEST_METHOD = " & tmp
                'Debug.Print "Ready"

                SendRequest = True
            Else

                ' HttpSendRequest failed
                'Debug.Print "HttpSendRequest call failed; Error code: " & err.LastDllError & "."
                'ToDebug "Send Request Error: " & WinInetErr(err.LastDllError)
                ToDebug "Err_SendRequest: " & err.LastDllError
            End If
        Else
            ' HttpOpenRequest failed
            'Debug.Print "HttpOpenRequest call failed; Error code: " & err.LastDllError & "."
            ToDebug "Err_OpenRequest: " & err.LastDllError
        End If
    Else
        ' InternetConnect failed
        'Debug.Print "InternetConnect call failed; Error code: " & err.LastDllError & "."
        ToDebug "Internet Connect Error: " & err.LastDllError
    End If
Else
    ' hInternetSession handle not allocated
    'Debug.Print "InternetOpen call failed: Error code: " & err.LastDllError & "."
    ToDebug "Err_InternetOpen: " & err.LastDllError
End If

Screen.MousePointer = vbNormal

End Function


Private Function GetUrlObject(u As String) As String
'отправим все до первой /
Dim posSlash As Long
posSlash = InStr(u, "/")
If posSlash <> 0 Then
    GetUrlObject = Right$(u, Len(u) - posSlash + 1)
Else
    GetUrlObject = vbNullString
End If
End Function

Private Function CheckUrl(u As String) As String
' http://www.ru -> http: ?
'If Len(u) = 0 Then u = vbNullString '"www.microsoft.com"
Dim posSlash As Long
posSlash = InStr(u, "/")
If posSlash <> 0 Then
    CheckUrl = Left$(u, posSlash - 1)
Else
    CheckUrl = u
End If
End Function

Private Function CheckUrlRev(u As String) As String
' http://www.ru -> www.ru
Dim posSlash As Long
posSlash = InStrRev(u, "//")
If posSlash <> 0 Then
    CheckUrlRev = Right$(u, Len(u) - posSlash - 1)
Else
    CheckUrlRev = u
End If
End Function

'Private Function GetQueryInfo(ByVal hHttpRequest As Long, ByVal lblContentType As Object, ByVal iInfoLevel As Long) As Boolean
'Private Function GetQueryInfo(ByVal hHttpRequest As Long, ByVal sContentType As String, ByVal iInfoLevel As Long) As Boolean
'Dim sBuffer         As String * 1024
'Dim lBufferLength   As Long
'lBufferLength = Len(sBuffer)
'GetQueryInfo = CBool(HttpQueryInfo(hHttpRequest, iInfoLevel, ByVal sBuffer, lBufferLength, 0))
''lblContentType = sBuffer
'sContentType = sBuffer
'End Function

Public Function BlockDelete(s As String, pStart As String, pFin As String)
Dim F As Long, l As Long
Dim p As String
'задача убить по шаблону "&#...;" - спец коды и символы html

'f - позиция "&#" (pStart)
'l - позиция ";" (pFin)
'p - удаляемое

F = InStr(1, s, pStart)
Do While F > 0
    l = 0
    l = InStr(F, s, pFin)
    If l > 0 Then
        p = Mid(s, F, l - F + Len(pFin))
        s = Replace(s, p, "")
        F = InStr(1, s, pStart)
    Else
        Exit Do 'нет конца
    End If    'l
Loop 'f

s = Replace(s, "()", "")
s = Replace(s, "  ", " ")

BlockDelete = Trim$(s)
End Function


Public Function GetReferer(sBaseAddr) As String
'получим из скрипта
'BaseAddress = "http://www.dvdempire.com/Exec/v4_item.asp?item_id="
'BaseAddress = "http://www.dvdempire.com"
'вернем "http://www.dvdempire.com/"
Dim i As Integer, j As Integer

On Error GoTo err

If LCase$(Left$(sBaseAddr, 7)) = "http://" Then
    i = 1
    For j = 1 To 3
        i = InStr(i + 1, sBaseAddr, "/")
    Next j
    If i = 0 Then
        'только //
        GetReferer = sBaseAddr & "/"
    ElseIf i > 7 Then
        GetReferer = Left$(sBaseAddr, i)
    Else
        GetReferer = vbNullString
    End If
Else
    GetReferer = vbNullString
End If
Exit Function

err:
ToDebug "Err_GeRef: " & sBaseAddr
GetReferer = vbNullString
End Function

'Private Function WinInetErr(c As Long) As String
'Select Case c
'Case 12001
'    WinInetErr = "ERROR_INTERNET_OUT_OF_HANDLES" '               No more handles could be generated at this time.
'Case 12002
'    WinInetErr = "ERROR_INTERNET_TIMEOUT" '              The request has timed out.
'Case 12003
'    WinInetErr = "ERROR_INTERNET_EXTENDED_ERROR '  An extended error was returned from the server. This is  typically a string or buffer containing a verbose error message. Call InternetGetLastResponseInfo to retrieve the  error text."
'Case 12004
'    WinInetErr = "ERROR_INTERNET_INTERNAL_ERROR" '               An internal error has occurred.
'Case 12005
'    WinInetErr = "ERROR_INTERNET_INVALID_URL '               The URL is invalid."
'Case 12006
'    WinInetErr = "ERROR_INTERNET_UNRECOGNIZED_SCHEME '               The URL scheme could not be recognized or is not supported."
'Case 12007
'    WinInetErr = "ERROR_INTERNET_NAME_NOT_RESOLVED '               The server name could not be resolved."
'Case 12008
'    WinInetErr = "ERROR_INTERNET_PROTOCOL_NOT_FOUND '               The requested protocol could not be located."
'Case 12009
'    WinInetErr = "ERROR_INTERNET_INVALID_OPTION '     A request to InternetQueryOption or InternetSetOption  specified an invalid option value."
'Case 12010
'    WinInetErr = "ERROR_INTERNET_BAD_OPTION_LENGTH " '               The length of an option supplied to InternetQueryOption or               InternetSetOption is incorrect for the type of option               specified."
'Case 12011
'    WinInetErr = "ERROR_INTERNET_OPTION_NOT_SETTABLE " '               The request option cannot be set, only queried."
'Case 12012
'    WinInetErr = "ERROR_INTERNET_SHUTDOWN " '               The Win32 Internet function support is being shut down or               unloaded."
'Case 12013
'    WinInetErr = "ERROR_INTERNET_INCORRECT_USER_NAME " '               The request to connect and log on to an FTP server could               not be completed because the supplied user name is               incorrect."
'Case 12014
'    WinInetErr = "ERROR_INTERNET_INCORRECT_PASSWORD " '               The request to connect and log on to an FTP server could               not be completed because the supplied password is               incorrect."
'Case 12015
'    WinInetErr = "ERROR_INTERNET_LOGIN_FAILURE " '               The request to connect to and log on to an FTP server               failed."
'Case 12016
'    WinInetErr = "ERROR_INTERNET_INVALID_OPERATION" '               The requested operation is invalid.
'Case 12017
'    WinInetErr = "ERROR_INTERNET_OPERATION_CANCELLED" '               The operation was canceled, usually because the handle on               which the request was operating was closed before the               operation completed.
'Case 12018
'    WinInetErr = "ERROR_INTERNET_INCORRECT_HANDLE_TYPE" '               The type of handle supplied is incorrect for this               operation.
'Case 12019
'    WinInetErr = "ERROR_INTERNET_INCORRECT_HANDLE_STATE" '               The requested operation cannot be carried out because the               handle supplied is not in the correct state.
'Case 12020
'    WinInetErr = "ERROR_INTERNET_NOT_PROXY_REQUEST" '               The request cannot be made via a proxy.
'Case 12021
'    WinInetErr = "ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND"  '             A required registry value could not be located.
'Case 12022
'    WinInetErr = "ERROR_INTERNET_BAD_REGISTRY_PARAMETER" '               A required registry value was located but is an incorrect               type or has an invalid value.
'Case 12023
'    WinInetErr = "ERROR_INTERNET_NO_DIRECT_ACCESS" '               Direct network access cannot be made at this time.
'Case 12024
'    WinInetErr = "ERROR_INTERNET_NO_CONTEXT" '               An asynchronous request could not be made because a zero               context value was supplied.
'Case 12025
'    WinInetErr = "ERROR_INTERNET_NO_CALLBACK" '               An asynchronous request could not be made because a               callback function has not been set.
'Case 12026
'    WinInetErr = "ERROR_INTERNET_REQUEST_PENDING" '               The required operation could not be completed because one               or more requests are pending.
'Case 12027
'    WinInetErr = "ERROR_INTERNET_INCORRECT_FORMAT" '               The format of the request is invalid.
'Case 12028
'    WinInetErr = "ERROR_INTERNET_ITEM_NOT_FOUND" '               The requested item could not be located.
'Case 12029
'    WinInetErr = "ERROR_INTERNET_CANNOT_CONNECT" '               The attempt to connect to the server failed.
'Case 12030
'    WinInetErr = "ERROR_INTERNET_CONNECTION_ABORTED" '               The connection with the server has been terminated.
'Case 12031
'    WinInetErr = "ERROR_INTERNET_CONNECTION_RESET" '               The connection with the server has been reset.
'Case 12032
'    WinInetErr = "ERROR_INTERNET_FORCE_RETRY" '               Calls for the Win32 Internet function to redo the request.
'Case 12033
'    WinInetErr = "ERROR_INTERNET_INVALID_PROXY_REQUEST" '               The request to the proxy was invalid.
'Case 12036
'    WinInetErr = "ERROR_INTERNET_HANDLE_EXISTS" '               The request failed because the handle already exists.
'Case 12037
'    WinInetErr = "ERROR_INTERNET_SEC_CERT_DATE_INVALID" '               SSL certificate date that was received from the server is               bad. The certificate is expired.
'Case 12038
'    WinInetErr = "ERROR_INTERNET_SEC_CERT_CN_INVALID" '               SSL certificate common name (host name field) is incorrect.               For example, if you entered www.server.com and the common               name on the certificate says www.different.com.
'Case 12039
'    WinInetErr = "ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR" '               The application is moving from a non-SSL to an SSL               connection because of a redirect.
'Case 12040
'    WinInetErr = "ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR" '               The application is moving from an SSL to an non-SSL               connection because of a redirect.
'Case 12041
'    WinInetErr = "ERROR_INTERNET_MIXED_SECURITY" '               Indicates that the content is not entirely secure. Some of               the content being viewed may have come from unsecured               servers.
'Case 12042
'    WinInetErr = "ERROR_INTERNET_CHG_POST_IS_NON_SECURE" '               The application is posting and attempting to change               multiple lines of text on a server that is not secure.
'Case 12043
'    WinInetErr = "ERROR_INTERNET_POST_IS_NON_SECURE"  '             The application is posting data to a server that is not               secure.
'Case 12110
'    WinInetErr = "ERROR_FTP_TRANSFER_IN_PROGRESS" '               The requested operation cannot be made on the FTP session               handle because an operation is already in progress.
'Case 12111
'    WinInetErr = "ERROR_FTP_DROPPED" '               The FTP operation was not completed because the session was               aborted.
'Case 12130
'    WinInetErr = "ERROR_GOPHER_PROTOCOL_ERROR" '               An error was detected while parsing data returned from the               gopher server.
'Case 12131
'    WinInetErr = "ERROR_GOPHER_NOT_FILE" '               The request must be made for a file locator.
'Case 12132
'    WinInetErr = "ERROR_GOPHER_DATA_ERROR" '               An error was detected while receiving data from the gopher               server.
'Case 12133
'    WinInetErr = "ERROR_GOPHER_END_OF_DATA" '               The end of the data has been reached.
'Case 12134
'    WinInetErr = "ERROR_GOPHER_INVALID_LOCATOR" '               The supplied locator is not valid.
'Case 12135
'    WinInetErr = "ERROR_GOPHER_INCORRECT_LOCATOR_TYPE" '               The type of the locator is not correct for this operation.
'Case 12136
'    WinInetErr = "ERROR_GOPHER_NOT_GOPHER_PLUS" '               The requested operation can only be made against a Gopher+               server or with a locator that specifies a Gopher+               operation.
'Case 12137
'    WinInetErr = "ERROR_GOPHER_ATTRIBUTE_NOT_FOUND" '               The requested attribute could not be located.
'Case 12138
'    WinInetErr = "ERROR_GOPHER_UNKNOWN_LOCATOR" '               The locator type is unknown.
'Case 12150
'    WinInetErr = "ERROR_HTTP_HEADER_NOT_FOUND" '               The requested header could not be located.
'Case 12151
'    WinInetErr = "ERROR_HTTP_DOWNLEVEL_SERVER" '               The server did not return any headers.
'Case 12152
'    WinInetErr = "ERROR_HTTP_INVALID_SERVER_RESPONSE" '               The server response could not be parsed.
'Case 12153
'    WinInetErr = "ERROR_HTTP_INVALID_HEADER" '               The supplied header is invalid.
'Case 12154
'    WinInetErr = "ERROR_HTTP_INVALID_QUERY_REQUEST" '               The request made to HttpQueryInfo is invalid.
'Case 12155
'    WinInetErr = "ERROR_HTTP_HEADER_ALREADY_EXISTS" '               The header could not be added because it already exists.
'Case 12156
'    WinInetErr = "ERROR_HTTP_REDIRECT_FAILED" '               The redirection failed because either the scheme changed               (for example, HTTP to FTP) or all attempts made to redirect               failed (default is five attempts).
'End Select
'
'End Function


