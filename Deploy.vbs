'Deploy.vbs
'	Visual Basic Script Used to Deploy .NET Components for the FiRRe Application...
'   Copyright © 2006-2013, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'	04/10/13	Ken Clark		Added FiRReData.NET & DBUtility.NET;
'	01/08/13	Ken Clark		Enhanced to support the initial deployment of VB6 DLL and RES files required after a new release
'								is created;
'	05/27/12	Ken Clark		Corrected file naming issues when renaming logs bigger than 10MB;
'	05/09/12	Ken Clark		Introduced \\WWS004\Project support;
'	07/21/11	Ken Clark		Added *.pdb to list of files deployed;
'   05/25/11    Ken Clark		Created;
'=================================================================================================================================
'Notes:
'Recommended Command-Line:	cscript Deploy.vbs All
'WWS004\Project Deployment:	cscript Deploy.vbs All "" Projects
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X "Deploy.vbs" "All"
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
Private Sub DeployProject(MainProject, Project, DeployProjects, Version, TargetFolder)
	Dim RealSource, RealTarget, CommandLine, VersionTag, DoCopy, SourcePath, TargetPath, iCopied, VB6, DevDeploy, FileName

	If UCase(MainProject) <> UCase(DeployProjects) And UCase(DeployProjects) <> "ALL" Then Exit Sub
	VB6 = CBool(InStr(Project, ".NET") = 0)
	DevDeploy = False

	VersionTag = ""
	If Version <> "" Then VersionTag = " Version " & Version
	RealSource = "V:\" & MainProject & VersionTag & "\"
	If Project <> "" Then RealSource = RealSource & Project & "\"
	If Not VB6 Then RealSource = RealSource & "bin\"
	If Not objFSO.FolderExists(Replace(RealSource, "\bin\", "")) Then
		LogMessage("   " & Replace(RealSource, "\bin\", "") & " does not exist!")
		LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
		Exit Sub
	End If

	If LCase(Right(TargetFolder, Len("Projects\"))) <> "projects\" Then
		VersionTag = ""
		If Version <> "" Then VersionTag = " v" & Version
	End If
	If TargetFolder = "S:\" Then
		If VB6 Then
			RealTarget = "S:\FiRRe" & VersionTag & "\Common\SunGard Shared\v" & Version & "\"
		Else
			RealTarget = "S:\FiRRe" & VersionTag & "\program files\SunGard\FiRRe" & VersionTag & "\"
		End If
		If Not objFSO.FolderExists(RealTarget) Then objFSO.CreateFolder(RealTarget)
	ElseIf TargetFolder = "\\WWS510\D$\" Then
		If VB6 Then
			RealTarget = "\\WWS510\C$\Program Files\Common Files\SunGard Shared\v" & Version & "\"
		Else
			RealTarget = "\\WWS510\D$\Program Files\SunGard\FiRRe" & VersionTag & "\"
		End If
		If Not objFSO.FolderExists(RealTarget) Then objFSO.CreateFolder(RealTarget)
	Else
		DevDeploy = True
		RealTarget = TargetFolder & MainProject & VersionTag & "\" 
		If Project <> "" Then RealTarget = RealTarget & Project & "\"
		If Not objFSO.FolderExists(RealTarget) Then objFSO.CreateFolder(RealTarget)
		If Not VB6 Then RealTarget = RealTarget & "bin\"
		If Not objFSO.FolderExists(RealTarget) Then objFSO.CreateFolder(RealTarget)
	End If

	LogMessage("   Deploying " & Replace(RealSource, "\bin\", "") & " to " & Replace(RealTarget, "\bin\", ""))
	LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
	On Error Resume Next
	iCopied = 0
	For Each objFile in	objFSO.GetFolder(RealSource).Files
		DoCopy = False
		FileName = objFile.Name
		If DevDeploy And VB6 And MainProject = "FiRRe" And Project = "" Then
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
    LogMessage("Deploy.vbs <""Components"" | ""TestRPC.NET"" | ""FiRRe"" | ""All""> [<Version>] [<UserName>]")
    LogMessage("  Projects      Determines set of projects to process")
    LogMessage("  Version       Version number in the form <MajorVersion>.<MinorVersion> (i.e. ""4.9"")")
    LogMessage("  UserName      Name of User to use for deployment (optional)")
    LogMessage("                Note: UserName of ""Projects"" will deploy to the \\WWS004\Projects share")
    LogMessage("                      UserName of ""WWS010"" will deploy to the server installation")
    LogMessage("")
    LogMessage("Software will be deployed from the version-specific folder from the current user's V-drive")
    LogMessage("to either the standard S-drive deployment location, or if <UserName> is provided, deployment")
    LogMessage("will occur to the provided user's V-drive.")
End Sub

LogMessage("[" & WScript.ScriptName & vbTab & Now() & "]")
Projects = "All"
Version = ""
UserName = "Projects"
TargetFolder = "S:\"
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
If UserName = "Projects" Then
	TargetFolder = "\\WWS004\" & UserName & "\"
ElseIf UserName = "WWS010" Then
	TargetFolder = "\\WWS010\D$\"
Else
	If Not objFSO.DriveExists("H:") Then
		LogMessage("   Drive Letter H must be mapped to \\WWS004\H$ in order to use this feature!")
		WScript.Quit
	End If
	TargetFolder = "H:\Development\" & UserName & "\"
End If
If Not objFSO.FolderExists(TargetFolder) Then
	LogMessage("   Target folder (" & TargetFolder & ") does not exist!")
	WScript.Quit
End If
LogMessage("   Deploying Projects to " & TargetFolder)
LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")

If Version <> "" Then
	Call DeployProject("Components", "SIASUTL", Projects, Version, TargetFolder)
	Call DeployProject("Components", "CRUFLSIA", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASADO", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASCL", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASCurrency", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASDB", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASEMail", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASRPC", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASRPC\SIASBPE00000", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASRPC\SIASBPE00001", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASRPC\SIASBPE21130", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASRPC\SIASBPE21140", Projects, Version, TargetFolder)
	Call DeployProject("Components", "SIASRPC\SIASBPE21150", Projects, Version, TargetFolder)
	'Call DeployProject("Components", "SIASRegisterDLLs", Projects, Version, TargetFolder)
End If
Call DeployProject("Components", "SIASSupport.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASWinsock.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASFTP.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASCL.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASCurrency.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASEmail.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASDB.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE00000.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE00001.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21090.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21110.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21120.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21130.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21150.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21170.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21180.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21190.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21200.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21210.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21220.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21230.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21240.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21250.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21260.NET", Projects, Version, TargetFolder)
'Call DeployProject("Components", "SIASRPC.NET\SIASBPE21270.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASTask.NET", Projects, Version, TargetFolder)
Call DeployProject("Components", "SIASWatcher.NET", Projects, Version, TargetFolder)

'Call DeployProject("FiRRe", "Components\CRUFLFiRRe", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReBase.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReData.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReCustom.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReCustomBNYM.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReCustomCSTC.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReLookup.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReSetup.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReBilling.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FCTXDemo.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReFile.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReSystem.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReLoader.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReForecast.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReProcessing.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReCollection.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReReports.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReHelp.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReSentinel.NET", Projects, Version, TargetFolder)
If Version <> "" Then
	Call DeployProject("FiRRe", "Components\SIASAFP", Projects, Version, TargetFolder)
	Call DeployProject("FiRRe", "Components\SIASCustom", Projects, Version, TargetFolder)
	Call DeployProject("FiRRe", "Components\SIASCustom\SIASBNYM", Projects, Version, TargetFolder)
	Call DeployProject("FiRRe", "Components\SIASCustom\SIASCSTC", Projects, Version, TargetFolder)
	'Call DeployProject("FiRRe", "Components\SIASCustom\SIASEXP", Projects, Version, TargetFolder)
	'Call DeployProject("FiRRe", "Components\SIASCustom\SIASWB", Projects, Version, TargetFolder)
End If
Call DeployProject("FiRRe", "Services\FiRReServer.NET", Projects, Version, TargetFolder)
If Version <> "" Then
	Call DeployProject("FiRRe", "Utilities\AFP", Projects, Version, TargetFolder)
	Call DeployProject("FiRRe", "Utilities\Benchmark", Projects, Version, TargetFolder)
	Call DeployProject("FiRRe", "Utilities\BusinessLine", Projects, Version, TargetFolder)
	Call DeployProject("FiRRe", "Utilities\DBUtility", Projects, Version, TargetFolder)
	Call DeployProject("FiRRe", "Utilities\PDF", Projects, Version, TargetFolder)
	Call DeployProject("FiRRe", "Utilities\Revenue", Projects, Version, TargetFolder)
End If
Call DeployProject("FiRRe", "Utilities\LookupTest.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\FiRReMonitor.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\TXDemo.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\TX32001.NET", Projects, Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\DBUtility.NET", Projects, Version, TargetFolder)
If Version <> "" Then 
	Call DeployProject("FiRRe", "Utilities\FiRRe.NET", Projects, Version, TargetFolder)
Else
	Call DeployProject("FiRRe", "", Projects, Version, TargetFolder)
End If
If UCase(Projects) = "FIRRE" Or UCase(Projects) = "ALL" Then
	VersionTag = ""
	If Version <> "" Then VersionTag = " Version " & Version
	LogMessage("   Copying FiRReCustom Components to FiRRe.NET project folder...")
	If Version <> "" Then
		Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\FiRRe.NET\bin", True)
		Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\FiRRe.NET\bin", True)
		Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\FiRRe.NET\bin", True)
	Else
		Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*", TargetFolder & "FiRRe" & VersionTag & "\bin", True)
		Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", TargetFolder & "FiRRe" & VersionTag & "\bin", True)
		Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", TargetFolder & "FiRRe" & VersionTag & "\bin", True)
	End If	
	LogMessage("   Copying FiRReCustom Components to FiRReServer.NET project folder...")
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Services\FiRReServer.NET\bin", True)
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Services\FiRReServer.NET\bin", True)
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Services\FiRReServer.NET\bin", True)
	LogMessage("   Copying FiRReCustom Components to DBUtility.NET project folder...")
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\DBUtility.NET\bin", True)
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\DBUtility.NET\bin", True)
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\DBUtility.NET\bin", True)
	LogMessage("   Copying FiRReCustom Components to LookupTest.NET project folder...")
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\LookupTest.NET\bin", True)
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\LookupTest.NET\bin", True)
	Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", TargetFolder & "FiRRe" & VersionTag & "\Utilities\LookupTest.NET\bin", True)
End If
LogMessage(vbCrLf & "Script complete @ " & Now())
Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
