'Fix .NET Security.vbs
'	Visual Basic Script Used to Fix .NET Security Settings for Network Share Use of the FiRRe Application...
'   Copyright © 2006-2012, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   09/04/12    Ken Clark		Created;
'=================================================================================================================================
'Notes:
'Recommended Command-Line:	cscript "Fix .NET Security.vbs"
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X "Fix .NET Security.vbs"
'=================================================================================================================================
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const UnicodeFormat = -1
Const MB = 1048576
Dim WshShell, objFSO, objNetwork, startFolder, Projects, Version, UserName, TargetFolder, Command, i
Set objNetwork = CreateObject("WScript.Network")
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
startFolder = WshShell.CurrentDirectory
UserName = objNetwork.UserName

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
Private Sub AddGroup(Command, URL, PermissionSet)
	LogMessage("      " & URL)
	'Note: "1.2" represents LocalIntranet Zone at the Machine-level...
	Execute(Command & " -pp off -m -ag 1.2 -url " & URL & " " & PermissionSet)
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
    LogMessage("""Fix .NET Security.vbs"" [UserName]")
    LogMessage("  UserName      Name of User to use for Development (optional)")
    LogMessage("")
    LogMessage(".NET Framework Security will be disabled and machine-level LocalIntranet Zone URL-based Code Groups ")
    LogMessage("will be defined with FullTrust permissions through the CASPOL utility for both v1.1 and v2.0.")
End Sub

LogMessage("[" & WScript.ScriptName & vbTab & Now() & "]")
Select Case WScript.Arguments.Count
	Case 0
	Case 1
		Select Case WScript.Arguments(0)
			Case "?", "/?", "-?"
				DisplayHelp()
				WScript.Quit
		End Select
	    UserName = WScript.Arguments(0)
	Case Else
		DisplayHelp()
		WScript.Quit
End Select
dim CASPOLv11, CASPOLv20
CASPOLv11 = "C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\CasPol.exe"
LogMessage("   Disabling Microsoft .NET Framework 1.1 Machine-Level Security...")
Execute(CASPOLv11 & " -machine -s off")
LogMessage("   Adding v1.1 Machine-Level LocalIntranet Zone URL-based Code Groups with FullTrust Permissions...")
If objFSO.FolderExists("\\WWS004\" & UCase(UserName) & "\") Then
	Call AddGroup(CASPOLv11, "file://V:/*", "FullTrust")
	Call AddGroup(CASPOLv11, "file://WWS004/" & UCase(UserName) & "/*", "FullTrust")
End If
Call AddGroup(CASPOLv11, "file://WWS004/SUNGARD/*", "FullTrust")
Call AddGroup(CASPOLv11, "file://WWS004/SUNGARDSAN/*", "FullTrust")
LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")

CASPOLv20 = "C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\CasPol.exe"
'LogMessage("   Disabling Microsoft .NET Framework 2.0 Machine-Level Security...")
'Execute(CASPOLv20 & " -machine -s off")
LogMessage("   Adding v2.0 Machine-Level LocalIntranet Zone URL-based Code Groups with FullTrust Permissions...")
If objFSO.FolderExists("\\WWS004\" & UCase(UserName) & "\") Then
	Call AddGroup(CASPOLv20, "file://V:/*", "FullTrust")
	Call AddGroup(CASPOLv20, "file://WWS004/" & UCase(UserName) & "/*", "FullTrust")
End If
Call AddGroup(CASPOLv20, "file://WWS004/SUNGARD/*", "FullTrust")
Call AddGroup(CASPOLv20, "file://WWS004/SUNGARDSAN/*", "FullTrust")
LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")

LogMessage(vbCrLf & "Script complete @ " & Now())
Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
