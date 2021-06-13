'Retrieve.vbs
'	Visual Basic Script Used to Retrieve .NET Components for the FiRRe Application...
'   Copyright © 2006-2013, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'	04/10/13	Ken Clark		Added FiRReData.NET & DBUtility.NET;
'	01/08/13	Ken Clark		Enhanced to support the initial deployment of VB6 DLL and RES files required after a new release
'								is created;
'	05/27/12	Ken Clark		Corrected file naming issues when renaming logs bigger than 10MB;
'   08/23/11    Ken Clark		Created;
'=================================================================================================================================
'Notes:
'Recommended Command-Line:	cscript "Retrieve.vbs" "All"
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X "Retrieve.vbs" "All"
'=================================================================================================================================
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const UnicodeFormat = -1
Const MB = 1048576
Dim WshShell, objFSO, startFolder, Projects, Version, UserName, TargetFolder, Command, i
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
startFolder = WshShell.CurrentDirectory

Private Sub LogMessage(Message)
	Dim objStdOut, objFile, LogFile, BaseName
    Set objStdOut = WScript.StdOut
    On Error Resume Next
    objStdOut.WriteLine Message
    On Error GoTo 0
	
	BaseName = startFolder & "\" & Replace(WScript.ScriptName, ".vbs", "")
	
	LogFile = BaseName & ".log"
    If objFSO.FileExists(LogFile) Then
		Set objFile = objFSO.GetFile(LogFile)
		If objFile.Size > 10*MB Then
            Dim dtModified, NewFileName
            dtModified = objFile.DateLastModified
            NewFileName = BaseName & "." & FormatTimeStamp(dtModified) & ".log"
            objFSO.MoveFile LogFile, NewFileName

            'If we successfully renamed our existing file, now police any older files that need to be deleted...
            'Dim LogDirInfo As New DirectoryInfo(LogFileInfo.DirectoryName)
            'Dim LogFileList() As FileInfo = LogDirInfo.GetFiles(String.Format("{0}.*{1}", Path.GetFileNameWithoutExtension(LogFileInfo.Name), LogFileInfo.Extension))
            'For Each iFileInfo As FileInfo In LogFileList
            '    If DateDiff(DateInterval.DayOfYear, iFileInfo.LastWriteTime, Now) > mSupport.LogRetentionDays Then iFileInfo.Delete()
            'Next
        End If
        Set objFile = Nothing
    End If
    
	Set objFile = objFSO.OpenTextFile(BaseName & ".log", ForAppending, True)
	objFile.WriteLine(Message)
	objFile.Close
	Set objFile = Nothing
    Set objStdOut = Nothing
End Sub
Private Function FormatTimeStamp(TimeStamp)
    iYear = Year(TimeStamp)
    iMonth = Month(TimeStamp)
    iDay = Day(TimeStamp)
    iHour = Hour(TimeStamp)
    iMinute = Minute(TimeStamp)
    iSecond = Second(TimeStamp)
    
    FormatTimeStamp = iYear
    if iMonth < 10 then FormatTimeStamp = FormatTimeStamp & "0"
    FormatTimeStamp = FormatTimeStamp & iMonth
    if iDay < 10 then FormatTimeStamp = FormatTimeStamp & "0"
    FormatTimeStamp = FormatTimeStamp & iDay
    FormatTimeStamp = FormatTimeStamp & "-"
    if iHour < 10 then FormatTimeStamp = FormatTimeStamp & "0"
    FormatTimeStamp = FormatTimeStamp & iHour
    if iMinute < 10 then FormatTimeStamp = FormatTimeStamp & "0"
    FormatTimeStamp = FormatTimeStamp & iMinute
    if iSecond < 10 then FormatTimeStamp = FormatTimeStamp & "0"
    FormatTimeStamp = FormatTimeStamp & iSecond
End Function
Private Function GetEnvironmentVariable(VariableName)
	Const wbemFlagReturnImmediately = &h10
	Const wbemFlagForwardOnly = &h20

    strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Environment Where Name='" & VariableName & "'", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
	For Each objItem In colItems
	    GetEnvironmentVariable = objItem.VariableValue
	    Exit Function
	Next
End Function
Private Function ExecuteWithoutOutput(Command)
	Dim oExec
    'LogMessage("Executing: " & Command)
	Set oExec = WshShell.Exec(Command)
    Do
		WScript.Sleep 10
	Loop Until oExec.Status <> 0
	ExecuteWithoutOutput = oExec.ExitCode
	Set oExec = Nothing
End Function
Private Function Execute(Command)
	Dim oExec, oStdOut, sOutput
    'LogMessage("Executing: " & SS & Command)
	Set oExec = WshShell.Exec(Command)
	Set oStdOut = oExec.StdOut
	sOutput = ""
    Do
		WScript.Sleep 10
		do until oStdOut.AtEndOfStream 
			sOutput = sOutput & oStdOut.ReadAll
		loop 
	Loop Until oExec.Status <> 0 and oStdOut.AtEndOfStream
	Execute = sOutput
	sOutput = ""
	Set oStdOut = Nothing
	Set oExec = Nothing
End Function
Private Sub RetrieveProject(MainProject, Project, RetrieveProjects, Version, TargetFolder)
	Dim RealSource, RealTarget, CommandLine, VersionTag, DoCopy, SourcePath, TargetPath, iCopied, VB6

	If UCase(MainProject) <> UCase(RetrieveProjects) And UCase(RetrieveProjects) <> "ALL" Then Exit Sub
	VB6 = CBool(InStr(Project, ".NET") = 0)
	
	VersionTag = ""
	If Version <> "" Then VersionTag = " Version " & Version
	RealSource = "\\WWS004\" & UserName & "\" & MainProject & VersionTag & "\"
	If Project <> "" Then RealSource = RealSource & Project & "\"
	If Not VB6 Then RealSource = RealSource & "bin\"
	If Not objFSO.FolderExists(Replace(RealSource, "\bin\", "")) Then
		LogMessage("   " & Replace(RealSource, "\bin\", "") & " does not exist!")
		LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
		Exit Sub
	End If

	RealTarget = TargetFolder & MainProject & VersionTag & "\"
	If Project <> "" Then RealTarget = RealTarget & Project & "\"
	If Not VB6 Then RealTarget = RealTarget & "bin\"
	If Not objFSO.FolderExists(RealTarget) Then objFSO.CreateFolder(RealTarget)

	LogMessage("   Retrieving " & Replace(RealSource, "\bin\", "") & " to " & Replace(RealTarget, "\bin\", ""))
	LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
	
	On Error Resume Next
	iCopied = 0
	For Each objFile in	objFSO.GetFolder(RealSource).Files
		DoCopy = False
		If VB6 And MainProject = "FiRRe" And Project = "" Then
			'The FiRRe project may contain a bunch of InterOp DLLs and other files that we do not want to process...
			If LCase(Right(objFile.Name, Len(".res"))) = ".res" And Not objFSO.FileExists(RealTarget & objFile.Name) Then DoCopy = True
		Else
			If LCase(Right(objFile.Name, Len(".dll"))) = ".dll" Then 
				If objFSO.GetFileVersion(objFile.Path) <> objFSO.GetFileVersion(RealTarget & "\" & objFile.Name) Then DoCopy = True
			ElseIf LCase(Right(objFile.Name, Len(".ocx"))) = ".ocx" Then 
				If objFSO.GetFileVersion(objFile.Path) <> objFSO.GetFileVersion(RealTarget & "\" & objFile.Name) Then DoCopy = True
			ElseIf LCase(Right(objFile.Name, Len(".net.exe"))) = ".net.exe" Then 
				If objFSO.GetFileVersion(objFile.Path) <> objFSO.GetFileVersion(RealTarget & "\" & objFile.Name) Then DoCopy = True
			ElseIf LCase(Right(objFile.Name, Len(".tlb"))) = ".tlb" Then 
				If objFSO.GetFile(objFile.Path).DateLastModified <> objFSO.GetFile(RealTarget & "\" & objFile.Name).DateLastModified Then DoCopy = True
			ElseIf LCase(Right(objFile.Name, Len(".snk"))) = ".snk" And Not objFSO.FileExists(RealTarget & objFile.Name) Then 
				DoCopy = True
			ElseIf LCase(Right(objFile.Name, Len(".pdb"))) = ".pdb" Then 
				If objFSO.GetFile(objFile.Path).DateLastModified <> objFSO.GetFile(RealTarget & "\" & objFile.Name).DateLastModified Then DoCopy = True
			ElseIf LCase(Right(objFile.Name, Len(".net.exe.manifest"))) = ".net.exe.manifest" And Not objFSO.FileExists(RealTarget & objFile.Name) Then 
				DoCopy = True
			ElseIf LCase(Right(objFile.Name, Len(".xml"))) = ".xml" And Not objFSO.FileExists(RealTarget & objFile.Name) Then 
				DoCopy = True
			ElseIf LCase(Right(objFile.Name, Len(".res"))) = ".res" And Not objFSO.FileExists(RealTarget & objFile.Name) Then 
				DoCopy = True
			End If
		End If
		If DoCopy Then
			LogMessage("      " & objFile.Name)
			Call objFSO.CopyFile(objFile.Path, RealTarget, True)
			If Err.number <> 0 Then 
				LogMessage("****  " & Err.Description & "; Copying " & objFile.Path & " to " & RealTarget)
			Else
				iCopied = iCopied + 1
			End If
		End If
		Err.Clear
	Next
	If Project = "Utilities\FiRRe.NET" Then
		LogMessage("      FiRReCustomBNYM.NET.*")
		Call objFSO.CopyFile(Replace(RealSource,Project,"Components\FiRReCustomBNYM.NET") & "FiRReCustomBNYM.NET.*", RealTarget, True)
		If Err.number <> 0 Then LogMessage("      " & Err.Description) Else iCopied = iCopied + 1
		LogMessage("      FiRReCustomCSTC.NET.*")
		Call objFSO.CopyFile(Replace(RealSource,Project,"Components\FiRReCustomCSTC.NET") & "FiRReCustomCSTC.NET.*", RealTarget, True)
		If Err.number <> 0 Then LogMessage("      " & Err.Description) Else iCopied = iCopied + 1
	End If
	On Error GoTo 0
	If iCopied = 0 Then LogMessage("      " & MainProject & "\" & Project & " is already current")
	LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
End Sub
'This is an attempt to redirect a script not started using CScript to do so (To date, it is not having the desired effect)...
Private Sub ReDirectToCScript
	On Error Resume Next
	WScript.StdOut.WriteLine ""
	If Err.number <> 0 Then
		Command = "cscript " & WScript.ScriptFullName
		For i = 0 To WScript.Arguments.Count - 1
			Command = Command & " """ & WScript.Arguments(0) & """"
		Next
		Execute Command
		WScript.Quit
	End If
	On Error GoTo 0
End Sub
Private Sub DisplayHelp
    LogMessage("Usage:")
    LogMessage("Retrieve.vbs <""Components"" | ""TestRPC.NET"" | ""FiRRe"" | ""All""> <Version> [<UserName>]")
    LogMessage("  Projects      Determines set of projects to process")
    LogMessage("  Version       Version number in the form <MajorVersion>.<MinorVersion> (i.e. ""4.9"")")
    LogMessage("  UserName      Name of User to use for retrieval (Defaults to Projects)")
    LogMessage("")
    LogMessage("Software will be retrieved from the version-specific folder from the given UserName's V-drive")
    LogMessage("(using the UNC format \\WWS004\<UserName>) to the current user's V-drive.")
End Sub

LogMessage("[" & WScript.ScriptName & vbTab & Now() & "]")
Projects = "All"
Version = ""
UserName = "Projects"
TargetFolder = "V:\"
Select Case WScript.Arguments.Count
	Case 1
		Select Case WScript.Arguments(0)
			Case "?", "/?", "-?"
				DisplayHelp()
				WScript.Quit
		End Select
		Projects = WScript.Arguments(0)
	Case 2
		Projects = WScript.Arguments(0)
		Version = WScript.Arguments(1)
	Case 3
		Projects = WScript.Arguments(0)
		Version = WScript.Arguments(1)
		UserName = WScript.Arguments(2)
	Case Else
		DisplayHelp()
		WScript.Quit
End Select
Select Case Projects
	Case "Components","FiRRe","All"
	Case Else
		DisplayHelp()
		WScript.Quit
End Select
If Not objFSO.FolderExists(TargetFolder) Then
	LogMessage("   Target folder (" & TargetFolder & ") does not exist!")
	WScript.Quit
End If
LogMessage("   Retrieving Projects to " & TargetFolder)
LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")

If Version <> "" Then
	Call RetrieveProject("Components", "SIASUTL", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "CRUFLSIA", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASADO", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASCL", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASCurrency", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASDB", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASEMail", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASRPC", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASRPC\SIASBPE00000", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASRPC\SIASBPE00001", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASRPC\SIASBPE21130", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASRPC\SIASBPE21140", Projects, Version, TargetFolder)
	Call RetrieveProject("Components", "SIASRPC\SIASBPE21150", Projects, Version, TargetFolder)
	'Call RetrieveProject("Components", "SIASRegisterDLLs", Projects, Version, TargetFolder)
End If
Call RetrieveProject("Components", "SIASSupport.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASWinsock.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASFTP.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASCL.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASCurrency.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASEmail.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASDB.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASRPC.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE00000.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE00001.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21090.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21110.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21120.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21130.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21150.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21170.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21180.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21190.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21200.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21210.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21220.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21230.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21240.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21250.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21260.NET", Projects, Version, TargetFolder)
'Call RetrieveProject("Components", "SIASRPC.NET\SIASBPE21270.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASTask.NET", Projects, Version, TargetFolder)
Call RetrieveProject("Components", "SIASWatcher.NET", Projects, Version, TargetFolder)

'Call RetrieveProject("FiRRe", "Components\CRUFLFiRRe", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReBase.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReData.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReCustom.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReCustomBNYM.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReCustomCSTC.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReLookup.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReSetup.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReBilling.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FCTXDemo.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReFile.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReSystem.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReLoader.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReForecast.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReProcessing.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReCollection.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReReports.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReHelp.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Components\FiRReSentinel.NET", Projects, Version, TargetFolder)
If Version <> "" Then
	Call RetrieveProject("FiRRe", "Components\SIASAFP", Projects, Version, TargetFolder)
	Call RetrieveProject("FiRRe", "Components\SIASCustom", Projects, Version, TargetFolder)
	Call RetrieveProject("FiRRe", "Components\SIASCustom\SIASBNYM", Projects, Version, TargetFolder)
	Call RetrieveProject("FiRRe", "Components\SIASCustom\SIASCSTC", Projects, Version, TargetFolder)
	'Call RetrieveProject("FiRRe", "Components\SIASCustom\SIASEXP", Projects, Version, TargetFolder)
	'Call RetrieveProject("FiRRe", "Components\SIASCustom\SIASWB", Projects, Version, TargetFolder)
End If
Call RetrieveProject("FiRRe", "Services\FiRReServer.NET", Projects, Version, TargetFolder)
If Version <> "" Then
	Call RetrieveProject("FiRRe", "Utilities\AFP", Projects, Version, TargetFolder)
	Call RetrieveProject("FiRRe", "Utilities\Benchmark", Projects, Version, TargetFolder)
	Call RetrieveProject("FiRRe", "Utilities\BusinessLine", Projects, Version, TargetFolder)
	Call RetrieveProject("FiRRe", "Utilities\DBUtility", Projects, Version, TargetFolder)
	Call RetrieveProject("FiRRe", "Utilities\PDF", Projects, Version, TargetFolder)
	Call RetrieveProject("FiRRe", "Utilities\Revenue", Projects, Version, TargetFolder)
End If
Call RetrieveProject("FiRRe", "Utilities\LookupTest.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Utilities\FiRReMonitor.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Utilities\TXDemo.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Utilities\TX32001.NET", Projects, Version, TargetFolder)
Call RetrieveProject("FiRRe", "Utilities\DBUtility.NET", Projects, Version, TargetFolder)
If Version <> "" Then
	Call RetrieveProject("FiRRe", "Utilities\FiRRe.NET", Projects, Version, TargetFolder)
Else
	Call RetrieveProject("FiRRe", "", Projects, Version, TargetFolder)
End If
LogMessage(vbCrLf & "Script complete @ " & Now())
Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
