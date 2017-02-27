Attribute VB_Name = "DrivesMod"
Option Explicit
Public Declare Function apiDriveType Lib "KERNEL32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function apiGetDrives Lib "KERNEL32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function drives(Optional chDrive As String) As Boolean
'Public Sub Drives(ByRef intRemovable As Integer, ByRef intNotRemovable As Integer, ByRef intCD As Integer, ByRef intRAM As Integer, ByRef intNetwork As Integer)
Dim intRemovable As Integer
Dim intNotRemovable As Integer
Dim intCD As Integer
Dim intRAM As Integer
Dim intNetwork As Integer
'---------------------------------------------------------------------------
' SUB: Drives
'
' Returns the number of removable, fixed, CD-ROM, RAM, and Network drives
' that are connected to your computer.
'
' THIS FUNCTION USES THE DRIVETYPE FUNCTION, SO IF YOU MODIFY THAT FUNCTION
' YOU MUST ALSO MODIFY THIS FUNCTION.
'
' OUT:  intRemovable        - Integer containing the number of removable drives
'       intNotRemovable     - Integer containing the number of fixed drives
'       intCD               - Integer containing the number of CD drives
'       intRAM              - Integer containing the number of RAM disks
'       int Network         - Integer containing the number of Network drives
'
'---------------------------------------------------------------------------
'
Dim Retrn As Long
Dim Buffer As Long
Dim temp As String
Dim intI As Integer
Dim Read As String
Dim Counter As Integer
Buffer = 10

'почистить
If frmOptFlag Then
temp = FrmOptions.ComboCDHid.Text
FrmOptions.ComboCDHid.Clear
FrmOptions.ComboCDHid.Text = temp
End If

Again:
temp = Space$(Buffer)
Retrn = apiGetDrives(Buffer, temp)
' Call the API function.

If Retrn > Buffer Then ' If the API returned a value that is bigger than Buffer,
    Buffer = Retrn     ' than the Buffer isn't big enough to hold the information.
    GoTo Again         ' In that case adjust the Buffer to the right size (returned by
End If                 ' the API) and try again.


' The API returns something like :
' A:\*B:\*C:\*D:\**  , with  * = NULL character
' 1234123412341234
' \ 1 \ 2 \ 3 \ 4 \
'
' So we start reading three characters, we step 4 further (the three we read + the
' NULL-character), and we read again three characters, step 4, ect.

Counter = 0
For intI = 1 To (Buffer - 4) Step 4

    Counter = Counter + 1
    Read = Mid$(temp, intI, 3)
    
If UCase$(chDrive) = UCase$(Read) Then drives = True
 
    Select Case DriveType(Read)
        Case "Removable"
            intRemovable = intRemovable + 1
        Case "Local"
            intNotRemovable = intNotRemovable + 1
            
            Case "Network"
            intNetwork = intNetwork + 1
        Case "CD-ROM"
        
If frmOptFlag Then FrmOptions.ComboCDHid.AddItem Read
        
        
            intCD = intCD + 1
        Case "RAM-disk"
            intRAM = intRAM + 1
    End Select

Next

End Function

Public Function DriveType(ByVal strRoot As String) As String
' if DriveType("d:\")="CD-ROM" then
'---------------------------------------------------------------------------
' FUNCTION: DriveType
'
' This function returns information about the drive you asked for. It will
' return whether the drive is a Removable drive, a non-removable (fixed)
' drive, a CD-ROM drive, a RAM drive or a Network drive.
'
' IN:  strRoot      - String containing the root of a drive. (e.g. "C:\")
'
' OUT: DriveType    - String containing type of drive.
'
' If the function fails a empty string is returned.
'
' You can also re-program this Function so that it doens't return a string,
' but it returns the value. That can be easier if you want to work with
' the returned information. I let it return a string, so that I can print
' it.
'

Dim lngType As Long
Const DRIVE_CDROM = 5       ' Some API constants required to
Const DRIVE_FIXED = 3       ' get the difference between the
Const DRIVE_RAMDISK = 6     ' drive types.
Const DRIVE_REMOTE = 4
Const DRIVE_REMOVABLE = 2
'Private Const DRIVE_UNKNOWN = 0
'Private Const DRIVE_ABSENT = 1

lngType = apiDriveType(strRoot)
' The API returns a value in lngType. Use the Constants to
' make the strings.

If Left$(strRoot, 2) = "\\" Then lngType = DRIVE_REMOTE 'sur

Select Case lngType
    Case DRIVE_REMOVABLE
        DriveType = "Removable"
    Case DRIVE_FIXED
        DriveType = "HDD"
    Case DRIVE_REMOTE
        DriveType = "Network"
    Case DRIVE_CDROM
        DriveType = "CD-ROM"
    Case DRIVE_RAMDISK
        DriveType = "RAM-disk"
    Case Else
        DriveType = vbNullString   ' If the API returns an error, we return a empty string
End Select

End Function

Public Function GetFirstOptoDrive() As String
Dim Retrn As Long
Dim Buffer As Long
Dim temp As String
Dim intI As Integer
Dim Read As String
Dim Counter As Integer
Buffer = 10


Again:
temp = Space$(Buffer)
Retrn = apiGetDrives(Buffer, temp)
' Call the API function.
If Retrn > Buffer Then    ' If the API returned a value that is bigger than Buffer,
    Buffer = Retrn     ' than the Buffer isn't big enough to hold the information.
    GoTo Again         ' In that case adjust the Buffer to the right size (returned by
End If                 ' the API) and try again.

Counter = 0
For intI = 1 To (Buffer - 4) Step 4

    Counter = Counter + 1
    Read = Mid$(temp, intI, 3)

    Select Case DriveType(Read)
    Case "CD-ROM"

        GetFirstOptoDrive = Read
        Exit For

    End Select

Next
End Function

