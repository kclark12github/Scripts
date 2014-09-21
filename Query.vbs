Private Sub LogMessage(Message)
	Dim objStdOut
    Set objStdOut = WScript.StdOut
	If Not IsNull(objStdOut) Then objStdOut.WriteLine Message
    Set objStdOut = Nothing
End Sub

Dim strComputer, objWMIService, objFile
Dim colFileList, SQLSource, bksPath

bksPath = "D:\Backups\GZPR141\My Music - Rock.bkf"

Set FSO = CreateObject("Scripting.FileSystemObject")
If Not FSO.FileExists(bksPath) Then
   Set objFile = FSO.CreateTextFile(bksPath)                   'Create a dummy file to ease FileName construction...
End If
Set objFile = FSO.GetFile(bksPath)
ParentFolder = objFile.ParentFolder
BaseName = FSO.GetBaseName(objFile)
Extension = FSO.GetExtensionName(objFile)
If objFile.Size = 0 Then FSO.DeleteFile(objFile.Path)

strComputer = "."
'Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")
SQLSource = "Select * from CIM_DataFile where Path='\\" & Replace(Mid(ParentFolder, 4), "\", "\\") & "\\' And FileName Like '" & BaseName & ".%' And Extension='" & Extension & "'"
Set colFiles = objWMIService.ExecQuery(SQLSource, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objFile In colFiles
	LogMessage vbTab & objFile.Path & "; " & objFile.FileName & "; " & objFile.Extension & "; " & TypeName(objFile)
Next
Set colFiles = Nothing
Set objFile = Nothing
Set objWMIService = Nothing
