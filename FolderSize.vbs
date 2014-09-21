'FolderSize.vbs
'	Visual Basic Script Used to Display Folder Size Information...
'   Copyright © 2006-2010, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'	02/15/10	Ken Clark		Added History and command-line parameters;
'=================================================================================================================================
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X FolderSize.vbs
'=================================================================================================================================
Private Sub LogMessage(Message)
	Dim objStdOut
    Set objStdOut = WScript.StdOut
	If Not IsNull(objStdOut) Then objStdOut.WriteLine Message
    Set objStdOut = Nothing
End Sub
Private Sub DisplayHelp
    LogMessage "Usage:"
    LogMessage "FolderSize.vbs <Folder>"
    LogMessage "  Folder     Specifies the root folder to size"
End Sub

Dim folderName
Dim objFSO
Dim objFolder
Dim colSubFolders

LogMessage "[FolderSize.vbs" & vbTab & Now() & "]"
If WScript.Arguments.Count > 0 Then
	If WScript.Arguments.Count = 1 Then
		folderName = WScript.Arguments(0)
	Else
		DisplayHelp()
		WScript.Quit
    End If
Else
	folderName = CurDir
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFolder = objFSO.GetFolder("C:\Documents and Settings\All Users\Documents\My Music\Rock")
Set objFolder = objFSO.GetFolder(folderName)
LogMessage "Sizing " & objFolder.Name & "..."
Set colSubfolders = objFolder.Subfolders
For Each objSubfolder in colSubfolders
	LogMessage vbTab & objSubFolder.Name & vbTab & FormatNumber(objSubFolder.Size / 1024 / 1024 / 1024, 2) & " GB"
Next
