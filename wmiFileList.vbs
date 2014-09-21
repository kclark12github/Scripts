Dim vArg, aArgs(), iCount
Dim SharedDocuments, BackupFolder, AltBackupFolder
Dim objStdOut
Dim dtmTargetDate, dtmConvertedDate
Dim strFolder, FolderCount, FileCount

Set objStdOut = WScript.StdOut
strComputer = "."

If WScript.Arguments.Count > 0 Then
    If WScript.Arguments.Count <> 2 Then
        If Not IsNull(objStdOut) Then 
            objStdOut.WriteLine "Usage:"
            objStdOut.WriteLine "wmiFileList.vbs [<Folder>, <BackupFileName>]"
        End If
        WScript.Quit
    End If
End If

Set FSO = CreateObject("Scripting.FileSystemObject")
Set dtmTargetDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

ReDim aArgs(wscript.Arguments.Count - 1)
For iCount = 0 To WScript.Arguments.Count - 1
    aArgs(iCount) = WScript.Arguments(iCount)
Next

strFolder = aArgs(0)
strTarget = aArgs(1)

creationDate = GetCreationDate(strTarget)
If IsNull(creationDate) Then
	objStdOut.WriteLine "Proceed with backup..."
Else
	'objStdOut.WriteLine "Creation Date: " & creationDate
	dtmTargetDate.SetVarDate creationDate, LOCAL_TIME
	FolderCount = 0
	FileCount = 0
	If ScanSubfolders(FSO.GetFolder(strFolder), dtmTargetDate, FolderCount, FileCount) Then 
	    objStdOut.WriteLine "Proceed with backup..."
	Else
	    objStdOut.WriteLine "Checked " & FormatNumber(FileCount,0,,,vbTrue) & " Files (in " & FormatNumber(FolderCount,0,,,vbTrue) & " Folders) - Found nothing new to backup..."
	End If
End If

Function GetCreationDate(bksPath)
	Dim strComputer, objWMIService
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")

	GetCreationDate = Null
	'strTarget won't be our real file name, but a template used to date-time stamp the true file name...
	If Not FSO.FileExists(bksPath) Then
	    Set objFile = FSO.CreateTextFile(bksPath)                   'Create a dummy file to ease FileName construction...
	End If
	Set objFile = FSO.GetFile(bksPath)
	ParentFolder = objFile.ParentFolder
	BaseName = FSO.GetBaseName(objFile)
	Extension = FSO.GetExtensionName(objFile)
	If objFile.Size = 0 Then FSO.DeleteFile(objFile.Path)

	'Note that we're not currently handling DST...
	
	'objStdOut.WriteLine "Attempting to find CreationDate for " & bksPath & "..."
    Set colFileList = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_Directory.Name='" & ParentFolder & "'} Where ResultClass = CIM_DataFile")
    For Each objFile In colFileList
		If UCase(Left(objFile.FileName, Len(BaseName))) = UCase(BaseName) And UCase(objFile.Extension) = UCase(Extension) Then
			'objStdOut.WriteLine Now() & vbTab & objFile.FileName & "." & objFile.Extension & " (" & TypeName(objFile) & ")"
            Set varDate = CreateObject("WbemScripting.SWbemDateTime")
            varDate.Value = objFile.CreationDate
            'objStdOut.WriteLine Now() & vbTab & varDate.GetVarDate(True) & " (" & objFile.CreationDate & ") - " & objFile.Name
            GetCreationDate = varDate.GetVarDate(True)
            Exit Function
        End If
    Next
End Function
Function ScanSubFolders(Folder, dtmTargetDate, FolderCount, FileCount)
    FolderCount = FolderCount + 1
    Set colFileList = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_Directory.Name='" & Folder & "'} Where ResultClass = CIM_DataFile")
    For Each objFile In colFileList
        FileCount = FileCount + 1
        If objFIle.LastModified > dtmTargetDate Then 
            'objStdOut.WriteLine Now() & vbTab & TypeName(objFile.LastModified)
            Set varDate = CreateObject("WbemScripting.SWbemDateTime")
            varDate.Value = objFile.LastModified
            objStdOut.WriteLine Now() & vbTab & varDate.GetVarDate(True) & " (" & objFile.LastModified & ") - " & objFile.Name
            ScanSubFolders = True
            Exit Function
        End If
    Next

    For Each Subfolder in Folder.SubFolders
        'objStdOut.WriteLine Subfolder.Path
        If ScanSubFolders(Subfolder, dtmTargetDate, FolderCount, FileCount) Then Exit For
    Next
    ScanSubFolders = False
End Function
