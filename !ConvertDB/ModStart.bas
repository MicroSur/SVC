Attribute VB_Name = "ModStart"
'Private Declare Function InitCommonControls Lib "COMCTL32.DLL" () As Long
Private Const ICC_BAR_CLASSES = &H4      'toolbar, statusbar, trackbar, tooltips
Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()
Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Type tagINITCOMMONCONTROLSEX
    dwSize As Long   ' size of this structure
    dwICC As Long    ' flags indicating which classes to be initialized.
End Type

Public LCID As Long
Public Declare Function GetSystemDefaultLCID Lib "KERNEL32" () As Long


Private Sub InitComctl32(dwFlags As Long)
   Dim icc As tagINITCOMMONCONTROLSEX
   On Error GoTo Err_OldVersion
   icc.dwSize = Len(icc)
   icc.dwICC = dwFlags
   InitCommonControlsEx icc
   On Error GoTo 0
   Exit Sub
Err_OldVersion:
   InitCommonControls
End Sub

Public Sub Main()
InitCommonControls
InitComctl32 ICC_BAR_CLASSES

frmMain.Show

End Sub

