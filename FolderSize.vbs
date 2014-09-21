Private Sub LogMessage(Message)
	Dim objStdOut
    Set objStdOut = WScript.StdOut
	If Not IsNull(objStdOut) Then objStdOut.WriteLine Message
    Set objStdOut = Nothing
End Sub

Dim objFSO
dim objFolder
Dim colSubFolders

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Documents and Settings\All Users\Documents\My Music\Rock")
LogMessage "Sizing " & objFolder.Name & "..."
Set colSubfolders = objFolder.Subfolders
For Each objSubfolder in colSubfolders
	LogMessage vbTab & objSubFolder.Name & vbTab & FormatNumber(objSubFolder.Size / 1024 / 1024 / 1024, 2) & " GB"
Next
