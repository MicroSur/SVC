'Enjoy! Sur.

On Error Resume Next
days = WScript.arguments(0)
If Err Then days = 1
On Error Goto 0

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set NewFile = objFSO.CreateTextFile("new.txt", True)

NewFile.WriteLine("Измененные файлы. Кол-во дней: ") & days
NewFile.WriteLine("Сегодня: ") & date()
NewFile.WriteLine()
objStartFolder = "."
Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

For Each objFile in colFiles
If (objFile.DateLastModified > (date() - days)) Then
' Select Case lcase(Right(objFile.Name,3))
'  Case "avi", "vid", "ivx", "mpg", "mp4", "asf", "vob", "wmv", "peg", "mov", "flv", "mpa"
  NewFile.WriteLine(objFile.DateLastModified & " " & objFolder.Path & "\" & objFile.Name)
' End Select
End If
Next

ShowSubfolders objFSO.GetFolder(objStartFolder)

Sub ShowSubFolders(Folder)
For Each Subfolder in Folder.SubFolders
 Set objFolder = objFSO.GetFolder(Subfolder.Path)
 Set colFiles = objFolder.Files
 For Each objFile in colFiles
  If (objFile.DateLastModified > (date() - days)) Then
   'Select Case lcase(Right(objFile.Name,3))
   'Case "avi", "vid", "ivx", "mpg", "mp4", "asf", "vob", "wmv", "peg", "mov", "flv", "mpa"
   NewFile.WriteLine(objFile.DateLastModified & " " & objFolder.Path & "\" & objFile.Name)
   'End Select
  end if
 Next
 ShowSubFolders Subfolder
Next
End Sub

objShell.run ("Notepad.exe new.txt")