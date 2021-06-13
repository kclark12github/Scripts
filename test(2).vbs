Private Sub LogMessage(Message)
	Dim objStdOut
    Set objStdOut = WScript.StdOut
	If Not IsNull(objStdOut) Then objStdOut.WriteLine Message
    Set objStdOut = Nothing
End Sub
Const DeleteReadOnly = TRUE
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Set objFSO = CreateObject("Scripting.FileSystemObject")

strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")
Set objSSA = objFSO.GetFile("H:\Development\Backups\VSS\WSRV08 VSS Database.20090920-020000.ssa")
BaseName = objFSO.GetBaseName(objSSA)
BackupName = Mid(BaseName, 1, Len(BaseName) - Len("yyyyMMdd-HHmmss."))
'Works in Windows XP, but not WSRV08's Windows 2000
'SQLSource = "Select * from CIM_DataFile where Path='\\" & Replace(Mid(objSSA.ParentFolder, 4), "\", "\\") & "\\' And FileName Like '" & BackupName & ".%' And CreationDate <= '" & DateAdd("d",-27,objSSA.DateCreated) & "'"
SQLSource = "Select * from CIM_DataFile where Path='\\" & Replace(Mid(objSSA.ParentFolder, 4), "\", "\\") & "\\' And CreationDate <= '" & DateAdd("d",-27,objSSA.DateCreated) & "'"
Set colFiles = objWMIService.ExecQuery(SQLSource, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objFile in colFiles
    If Left(objFile.FileName, Len(BackupName)+1) = BackupName & "." Then
		LogMessage Now() & vbTab & objFile.Name
	End If
Next    
If Not IsNull(WScript.StdOut) Then WScript.StdOut.Close
WScript.Quit