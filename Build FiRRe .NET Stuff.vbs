'Build FiRRe .NET Stuff.vbs
'	Visual Basic Script Used to Build .NET Components for the FiRRe Application...
'   Copyright © 2006-2011, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   05/31/11    Ken Clark		Created;
'=================================================================================================================================
'Notes:
'Recommended Command-Line:	cscript "Build FiRRe .NET Stuff.vbs"
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X "Build FiRRe .NET Stuff.vbs"
'=================================================================================================================================
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const UnicodeFormat = -1
Const MB = 1048576
Dim WshShell, objFSO, startFolder
Dim iSucceeded, iFailed, iSkipped
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
Private Function GetProjectName(ProjectFileNameBase)
	Dim iPos
	iPos = InStrRev(ProjectFileNameBase, "\")
	If iPos <> 0 Then
		GetProjectName = Mid(ProjectFileNameBase, iPos+1)
	Else
		GetProjectName = vbNullString
	End If
End Function
Private Sub BuildProject(ProjectFileNameBase, iSucceeded, iFailed, iSkipped)
	Dim vsFileName, CommandLine, ProjectName, objStream, strLine, iPos
	vsFileName = "\Program Files\Microsoft Visual Studio .NET 2003\Common7\IDE\DEVENV.exe"
	If objFSO.FileExists("C:" & vsFileName) Then
		vsFileName = "C:" & vsFileName
	ElseIf objFSO.FileExists("D:" & vsFileName) Then
		vsFileName = "D:" & vsFileName
	Else
		LogMessage("   Error: Unable to determine DEVENV.EXE location")
		Exit Sub
	End If
	
	If Not objFSO.FileExists(ProjectFileNameBase & ".sln") Then
		LogMessage("   Error: " & ProjectFileNameBase & ".sln does not exist!")
		Exit Sub
	ElseIf Not objFSO.FileExists(ProjectFileNameBase & ".vbproj") Then
		LogMessage("   Error: " & ProjectFileNameBase & ".vbproj does not exist!")
		Exit Sub
	End If
	WshShell.CurrentDirectory = objFSO.GetParentFolderName(ProjectFileNameBase & ".sln")
	ProjectName = GetProjectName(ProjectFileNameBase)
	LogMessage("   Building " & ProjectName)
	LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
	If objFSO.FileExists(ProjectName & ".log") Then objFSO.DeleteFile(ProjectName & ".log")
	CommandLine = ProjectName & ".sln /rebuild Debug /project " & ProjectName & ".vbproj /out " & ProjectName & ".log"
	LogMessage("   DEVENV.exe " & CommandLine)
	LogMessage("   " & Execute(vsFileName & " " & CommandLine))
	If objFSO.FileExists(ProjectName & ".log") Then
		Set objStream = objFSO.OpenTextFile(ProjectName & ".log", ForReading, False)
		Do While (Not objStream.AtEndOfStream) 
			strLine = objStream.ReadLine
			If Left(Trim(strLine), Len("Rebuild All: ")) = "Rebuild All: " then
				iPos = InStr(strLine, "Rebuild All: ") + Len("Rebuild All: ")
				'We can cheat here because we only ever build one project at a time (so these values will only ever be 1 or 0)...
				iSucceeded = iSucceeded + CInt(Mid(strLine, iPos, 1))
				iPos = InStr(strLine, " succeeded, ") + Len(" succeeded, ")
				iFailed = iFailed + CInt(Mid(strLine, iPos, 1))
				iPos = InStr(strLine, " failed, ") + Len(" failed, ")
				iSkipped = iSkipped + CInt(Mid(strLine, iPos, 1))
			End If
			LogMessage("   " & strLine)
		Loop
		objStream.Close
	End If
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

LogMessage("[" & WScript.ScriptName & vbTab & Now() & "]")
LogMessage("   Current Directory: " & WshShell.CurrentDirectory)
LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
iSucceeded = 0
iFailed = 0
iSkipped = 0
'Call BuildProject("V:\Components\SIASTask.NET\SIASTask.NET", iSucceeded, iFailed, iSkipped)
'Call BuildProject("V:\Components\SIASWatcher.NET\SIASWatcher.NET", iSucceeded, iFailed, iSkipped)

Call BuildProject("V:\FiRRe\Components\FiRReBase.NET\FiRReBase.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReLookup.NET\FiRReLookup.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FCTXDemo.NET\FCTXDemo.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReFile.NET\FiRReFile.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReSystem.NET\FiRReSystem.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReSetup.NET\FiRReSetup.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReLoader.NET\FiRReLoader.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReForecast.NET\FiRReForecast.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReProcessing.NET\FiRReProcessing.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReBilling.NET\FiRReBilling.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReCollection.NET\FiRReCollection.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReReports.NET\FiRReReports.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReHelp.NET\FiRReHelp.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Components\FiRReSentinel.NET\FiRReSentinel.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Services\FiRReServer.NET\FiRReServer.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Utilities\LookupTest.NET\LookupTest.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Utilities\FiRReMonitor.NET\FiRReMonitor.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Utilities\TXDemo.NET\TXDemo.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Utilities\TX32001.NET\TX32001.NET", iSucceeded, iFailed, iSkipped)
Call BuildProject("V:\FiRRe\Utilities\FiRRe.NET\FiRRe.NET", iSucceeded, iFailed, iSkipped)
LogMessage("   " & iSucceeded & " succeeded, " & iFailed & " failed, " & iSkipped & " skipped")
Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
