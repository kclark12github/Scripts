'ReRegister.vbs
'	Visual Basic Script Used to Re-Register ActiveX and VB.NET Components for FiRRe...
'   Copyright © 2006-2012, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   08/12/14	Ken Clark		Segragated VB6/VS2003 from VS2005/.NET Framework 2.0;
'   04/10/13	Ken Clark		Added FiRReData.NET;
'   04/05/12	Ken Clark		Restored behavior no-argument behavior broken when adding TargetFolder parameter support below;
'   04/04/12	Ken Clark		Added help and parameter support to provide TargetFolder via command-line;
'   05/25/11    Ken Clark		Created;
'=================================================================================================================================
'Notes:
'Recommended Command-Line:	cscript ReRegister.vbs
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X ReRegister.vbs
'=================================================================================================================================
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const UnicodeFormat = -1
Const MB = 1048576
Dim WshShell, objFSO, startFolder, Version, TargetFolder, RealTarget, iPos, Framework
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
Private Sub RegisterActiveX(RealTarget, FileName)
	Dim objFile, CommandLine, ExitCode
	If Not objFSO.FileExists(RealTarget & "\" & FileName) Then Exit Sub
	
	Set objFile = objFSO.GetFile(RealTarget & "\" & FileName)	
	Select Case LCase(objFSO.GetExtensionName(objFile.Path))
		Case "dll", "ocx"
			LogMessage("      " & objFile.Name)
			CommandLine = "REGSVR32.exe /s/c/u """ & objFile.Path & """"
			ExitCode = ExecuteWithoutOutput(CommandLine)
			If ExitCode <> 0 Then LogMessage("         Unregister failed. (" & ExitCode & ")")

			CommandLine = "REGSVR32.exe /s/c """ & objFile.Path & """"
			ExitCode = ExecuteWithoutOutput(CommandLine)
			If ExitCode <> 0 Then LogMessage("         Register failed. (" & ExitCode & ")")
		Case Else
	End Select
End Sub
Private Sub RegisterDotNet(TargetFolder, FileName, Interop)
	Dim objFile, CommandLine, ExitCode

	If Not objFSO.FileExists(TargetFolder & "\" & FileName) Then
		LogMessage("         " & FileName & " not found!")
		Exit Sub
	End If
	Set objFile = objFSO.GetFile(TargetFolder & "\" & FileName)	
	If LCase(Right(objFile.Name, Len(".net.dll"))) = ".net.dll" Then 
		LogMessage("      " & objFile.Name)
		CommandLine = Framework & "\RegAsm.exe /unregister """ & objFile.Path & """ /silent"
		ExitCode = ExecuteWithoutOutput(CommandLine)
		If ExitCode <> 0 Then LogMessage("         Unregister failed. (" & ExitCode & ")")
		
		CommandLine = Framework & "\RegAsm.exe """ & objFile.Path & """ /silent"
		If Interop Then CommandLine = CommandLine & " /tlb"
		ExitCode = ExecuteWithoutOutput(CommandLine)
		If ExitCode <> 0 Then LogMessage("         Register failed. (" & ExitCode & ")")

		'If objFSO.FileExists(objFile.ParentFolder.Path & "\" & objFSO.GetBaseName(objFile.Path) & ".tlb") Then
		'	CommandLine = "S:\Regtlb.exe """ & objFile.ParentFolder.Path & "\" & objFSO.GetBaseName(objFile.Path) & ".tlb"" -q"
		'	ExitCode = ExecuteWithoutOutput(CommandLine)
		'	If ExitCode <> 0 Then LogMessage("         Interop registration failed. (" & ExitCode & ")")
		'End If
	End If
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
    LogMessage("ReRegister.vbs [<TargetFolder>]")
    LogMessage("  TargetFolder  Deployment folder (defaults to location of this script)")
    LogMessage("")
End Sub

LogMessage("[" & WScript.ScriptName & vbTab & Now() & "]")

Version = ""
TargetFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
LogMessage("TargetFolder: " & TargetFolder)
Select Case WScript.Arguments.Count
	Case 0
		'Use default TargetFolder assigned above...
	Case 1
		Select Case WScript.Arguments(0)
			Case "?", "/?", "-?"
				DisplayHelp()
				WScript.Quit
		End Select
		TargetFolder = WScript.Arguments(0)
	Case Else
		DisplayHelp()
		WScript.Quit
End Select
If Not objFSO.FolderExists(TargetFolder) Then
	LogMessage("   Target folder (" & TargetFolder & ") does not exist!")
	WScript.Quit
End If

'If Left(objFSO.GetFolder(TargetFolder).Name, Len("FiRRe")) <> "FiRRe" Then
'	LogMessage("   This Script must be run from a FiRRe deployment folder!")
'	WScript.Quit
'End If
iPos = InStr(objFSO.GetFolder(TargetFolder).Name, " v")
If iPos > 0 Then Version = Mid(objFSO.GetFolder(TargetFolder).Name, iPos+1)
If Version = "" Then
	Version = objFSO.GetFileVersion(TargetFolder & "\FiRRe.NET.exe")
	Version = Mid(Version, 1, InStrRev(Version, ".")-1)			'Remove Build...
	Version = "v" & Mid(Version, 1, InStrRev(Version, ".")-1)	'Remove Revision...
End If

Framework = "C:\Windows\Microsoft.NET\Framework\v2.0.50727"
If Version < "v5" Then
    Framework = "C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322"
	LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
	RealTarget = objFSO.GetParentFolderName(TargetFolder)	'SunGard
	RealTarget = objFSO.GetParentFolderName(RealTarget)		'program files
	RealTarget = objFSO.GetParentFolderName(RealTarget)		'FiRRe/FiRRe vx.y
	RealTarget = RealTarget & "\Common\SunGard Shared\" & Version
	If Not objFSO.FolderExists(RealTarget) Then
		RealTarget = "C:\Program Files\Common Files\SunGard Shared\" & Version	'UAT/Production installation location...
	End If

	LogMessage("   Registering ActiveX Components from " & RealTarget)
	LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
	Call RegisterActiveX(RealTarget, "SIASUTL.dll")
	Call RegisterActiveX(RealTarget, "CRUFLSIA.dll")
	Call RegisterActiveX(RealTarget, "SIASADO.dll")
	Call RegisterActiveX(RealTarget, "SIASCL.dll")
	Call RegisterActiveX(RealTarget, "SIASDB.dll")
	Call RegisterActiveX(RealTarget, "SIASEmail.dll")
	Call RegisterActiveX(RealTarget, "SIASRPC.dll")
	Call RegisterActiveX(RealTarget, "SIASBPE00000.dll")
	Call RegisterActiveX(RealTarget, "SIASBPE00001.dll")
	Call RegisterActiveX(RealTarget, "SIASBPE21130.dll")
	Call RegisterActiveX(RealTarget, "SIASBPE21140.dll")
	Call RegisterActiveX(RealTarget, "SIASBPE21150.dll")
	Call RegisterActiveX(RealTarget, "SIASCurrency.ocx")
	Call RegisterActiveX(RealTarget, "SIASAFP.dll")
	Call RegisterActiveX(RealTarget, "SIASCustom.dll")
	'These are not necessary as they are binary compatible with SIASCustom.dll
	'Call RegisterActiveX(RealTarget, "SIASBNYM.dll")
	'Call RegisterActiveX(RealTarget, "SIASCSTC.dll")
	'Call RegisterActiveX(RealTarget, "SIASEXP.dll")
	'Call RegisterActiveX(RealTarget, "SIASWB.dll")
End If

LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
LogMessage("   Registering .NET Components from " & TargetFolder)
LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
Call RegisterDotNet(TargetFolder, "SIASSupport.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASWinsock.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASFTP.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASCL.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASCurrency.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASDB.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASEmail.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASRPC.NET.dll", False)
Call RegisterDotNet(TargetFolder, "SIASBPE00000.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASBPE00001.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASBPE21130.NET.dll", False)
Call RegisterDotNet(TargetFolder, "SIASBPE21150.NET.dll", False)
Call RegisterDotNet(TargetFolder, "SIASTask.NET.dll", True)
Call RegisterDotNet(TargetFolder, "SIASWatcher.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReBase.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReData.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReCustom.NET.dll", False)
Call RegisterDotNet(TargetFolder, "FiRReCustomBNYM.NET.dll", False)
Call RegisterDotNet(TargetFolder, "FiRReCustomCSTC.NET.dll", False)
Call RegisterDotNet(TargetFolder, "FiRReLookup.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FCTXDemo.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReFile.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReSystem.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReSetup.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReLoader.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReForecast.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReProcessing.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReBilling.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReCollection.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReReports.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReHelp.NET.dll", True)
Call RegisterDotNet(TargetFolder, "FiRReSentinel.NET.dll", False)
LogMessage(vbCrLf & "Script complete @ " & Now())
Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
