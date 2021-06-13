'Deploy All .NET Stuff.vbs
'	Visual Basic Script Used to Deploy All .NET Components for the FiRRe Application...
'   Copyright © 2006-2011, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   05/25/11    Ken Clark		Created;
'=================================================================================================================================
'Notes:
'Recommended Command-Line:	cscript "Deploy All .NET Stuff.vbs"
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X "Deploy All .NET Stuff.vbs"
'=================================================================================================================================
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const UnicodeFormat = -1
Const MB = 1048576
Dim WshShell, objFSO, startFolder, Version, UserName, TargetFolder, Command, i
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
            NewFileName = startFolder & "\" & BaseName & "." & FormatTimeStamp(dtModified) & ".log"
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
Private Sub DeployProject(MainProject, Project, Version, TargetFolder)
	Dim RealSource, RealTarget, CommandLine, VersionTag, DoCopy, SourcePath, TargetPath

	VersionTag = ""
	If Version <> "" Then VersionTag = " Version " & Version
	RealSource = "V:\" & MainProject & VersionTag & "\" & Project & "\bin\"

	VersionTag = ""
	If Version <> "" Then VersionTag = " v" & Version
	If TargetFolder = "S:\" Then
		RealTarget = "S:\FiRRe" & VersionTag & "\program files\SunGard\FiRRe" & VersionTag
	Else
		RealTarget = TargetFolder & MainProject & VersionTag & "\" & Project & "\bin\"
		If Not objFSO.FolderExists(RealTarget) Then objFSO.CreateFolder(RealTarget)
		'If Not objFSO.FolderExists(Left(RealTarget,Len(RealTarget)-1)) Then objFSO.CreateFolder(Left(RealTarget,Len(RealTarget)-1))
	End If

	LogMessage("   Deploying " & Replace(RealSource, "\bin\", "") & " to " & Replace(RealTarget, "\bin\", ""))
	LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
	On Error Resume Next
	For Each objFile in	objFSO.GetFolder(RealSource).Files
		DoCopy = False
		If LCase(Right(objFile.Name, Len(".dll"))) = ".dll" Then 
			If objFSO.GetFileVersion(objFile.Path) <> objFSO.GetFileVersion(RealTarget & "\" & objFile.Name) Then DoCopy = True
		ElseIf LCase(Right(objFile.Name, Len(".snk"))) = ".snk" And Not objFSO.FileExists(RealTarget & objFile.Name) Then 
			If objFSO.GetFileVersion(objFile.Path) <> objFSO.GetFileVersion(RealTarget & "\" & objFile.Name) Then DoCopy = True
		ElseIf LCase(Right(objFile.Name, Len(".net.exe"))) = ".net.exe" Then 
			If objFSO.GetFileVersion(objFile.Path) <> objFSO.GetFileVersion(RealTarget & "\" & objFile.Name) Then DoCopy = True
		ElseIf LCase(Right(objFile.Name, Len(".net.exe.manifest"))) = ".net.exe.manifest" Then 
			If objFSO.GetFileVersion(objFile.Path) <> objFSO.GetFileVersion(RealTarget & "\" & objFile.Name) Then DoCopy = True
		ElseIf LCase(Right(objFile.Name, Len(".xml"))) = ".xml" Then 
			If objFSO.GetFileVersion(objFile.Path) <> objFSO.GetFileVersion(RealTarget & "\" & objFile.Name) Then DoCopy = True
		End If

		If DoCopy Then
			LogMessage("      " & objFile.Name)
			Call objFSO.CopyFile(objFile.Path, RealTarget, True)
		End If
		If Err.number <> 0 Then LogMessage("      " & Err.Description)
		Err.Clear
	Next
	On Error GoTo 0
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
    LogMessage("Deploy All .NET Stuff.vbs <Version>[,<UserName>]")
    LogMessage("  Version       Version number in the form <MajorVersion>.<MinorVersion> (i.e. ""4.9"")")
    LogMessage("  UserName      Name of User to use for deployment (optional)")
    LogMessage("")
    LogMessage("Software will be deployed from the version-specific folder from the current user's V-drive")
    LogMessage("to either the standard S-drive deployment location, or if <UserName> is provided, deployment")
    LogMessage("will occur to the provided user's V-drive.")
End Sub

LogMessage("[" & WScript.ScriptName & vbTab & Now() & "]")
Select Case WScript.Arguments.Count
	Case 1
		Version = WScript.Arguments(0)
		TargetFolder = "S:\"
	Case 2
		Version = WScript.Arguments(0)
		UserName = WScript.Arguments(1)
		If Not objFSO.DriveExists("H:") Then
			LogMessage("   Drive Letter H must be mapped to \\WSRV08\H$ in order to use this feature!")
			WScript.Quit
		End If
		TargetFolder = "H:\Development\" & UserName & "\"
	Case Else
		DisplayHelp()
		WScript.Quit
End Select
If Not objFSO.FolderExists(TargetFolder) Then
	LogMessage("   Target folder (" & TargetFolder & ") does not exist!")
	WScript.Quit
End If
LogMessage("   Deploying Projects to " & TargetFolder)
LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")

Call DeployProject("Components", "SIASSupport.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASWinsock.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASFTP.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASCL.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASCurrency.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASEmail.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASDB.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE00000.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE00001.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21090.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21110.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21120.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21130.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21150.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21170.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21180.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21190.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21200.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21210.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21220.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21230.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21240.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21250.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21260.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASRPC.NET\SIASBPE21270.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASTask.NET", Version, TargetFolder)
Call DeployProject("Components", "SIASWatcher.NET", Version, TargetFolder)

Call DeployProject("FiRRe", "Components\FiRReBase.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReLookup.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FCTXDemo.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReBilling.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReCollection.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReFile.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReForecast.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReHelp.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReLoader.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReProcessing.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReReports.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReSentinel.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReSetup.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Components\FiRReSystem.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Services\FiRReServer.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\LookupTest.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\FiRReMonitor.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\TXDemo.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\TX32001.NET", Version, TargetFolder)
Call DeployProject("FiRRe", "Utilities\FiRRe.NET", Version, TargetFolder)

Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
