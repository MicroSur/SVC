Attribute VB_Name = "ModStart"
Option Explicit
      
Private Declare Function CreateMutex Lib "KERNEL32" Alias "CreateMutexA" _
   (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Const ERROR_ALREADY_EXISTS = 183&
Private m_hMutex As Long
Private m_bInDevelopment As Boolean
' Change this line to match your app:
Private Const mcTHISAPPID = "SurVideoCatalog"

Public Const GW_HWNDPREV = 3

Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) _
         As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'
Public kDPI As Single 'коэф пересчета при отличной от 96 масштаба

'Private Declare Function InitCommonControls Lib "COMCTL32.DLL" () As Long
'Private Const ICC_BAR_CLASSES = &H4      'toolbar, statusbar, trackbar, tooltips
'Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()
'Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

'mzt Private Type tagINITCOMMONCONTROLSEX
'mzt    dwSize As Long   ' size of this structure
'mzt    dwICC As Long    ' flags indicating which classes to be initialized.
'mzt End Type

'Private Sub InitComctl32(dwFlags As Long)
'   Dim icc As tagINITCOMMONCONTROLSEX
'   On Error GoTo Err_OldVersion
'   icc.dwSize = Len(icc)
'   icc.dwICC = dwFlags
'   InitCommonControlsEx icc
'   On Error GoTo 0
'   Exit Sub
'Err_OldVersion:
'   InitCommonControls
'End Sub

Public Sub Main()
If ExitSVC Then Exit Sub

   If (WeAreAlone(mcTHISAPPID & "_APPLICATION_MUTEX")) Then
   
   'InitCommonControls
   'InitComctl32 ICC_BAR_CLASSES
   
'коэф масштаба фонтов
kDPI = 15 / Screen.TwipsPerPixelX  '15=1440/96

   frmSplash.Show
   'FrmOptions.Show
   'FrmRule.Show
   'frmActFilt.Show
   
   Else
    ActivatePrevInstance
   End If
End Sub

Private Function WeAreAlone(ByVal sMutex As String) As Boolean
' Don't call Mutex when in VBIDE because it will apply
' for the entire VB IDE session, not just the app's
' session.
If InDevelopment Then
    WeAreAlone = Not (App.PrevInstance)
Else
    ' Ensures we don't run a second instance even
    ' if the first instance is in the start-up phase
    m_hMutex = CreateMutex(ByVal 0&, 1, sMutex)
    If (err.LastDllError = ERROR_ALREADY_EXISTS) Then
        CloseHandle m_hMutex
    Else
        WeAreAlone = True
    End If
End If
End Function
Private Sub ActivatePrevInstance()
Dim OldTitle As String
Dim PrevHndl As Long
'         Dim result As Long

'Save the title of the application.
OldTitle = App.title
'Rename the title of this application so FindWindow
'will not find this application instance.
App.title = "unwanted instance"

'Attempt to get window handle using VB4 class name.
PrevHndl = FindWindow("ThunderRTMain", OldTitle)

'Check for no success.
If PrevHndl = 0 Then
    'Attempt to get window handle using VB5 class name.
    PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
End If

'Check if found
If PrevHndl = 0 Then
    'Attempt to get window handle using VB6 class name
    PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
End If

'Check if found
If PrevHndl = 0 Then
    'No previous instance found.
    Exit Sub
End If

'Get handle to previous window.
PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)

'Restore the program.
'result =
OpenIcon (PrevHndl)

'Activate the application.
'result =
SetForegroundWindow (PrevHndl)

'End the application.
End
End Sub
Public Function InDevelopment() As Boolean
   ' Debug.Assert code not run in an EXE. Therefore
   ' m_bInDevelopment variable is never set.
   Debug.Assert InDevelopmentHack() = True
   InDevelopment = m_bInDevelopment
End Function

Private Function InDevelopmentHack() As Boolean
   m_bInDevelopment = True
   InDevelopmentHack = m_bInDevelopment
End Function

Public Function EndApp()
   ' Call this to remove the Mutex. It will be cleared
   ' anyway by windows, but this ensures it works.
   If (m_hMutex <> 0) Then
      CloseHandle m_hMutex
   End If
   m_hMutex = 0
End Function

