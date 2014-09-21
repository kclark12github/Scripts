'VSSRename.vbs
'	Visual Basic Script Used to Automate New Product Version Creation (starting with renaming projects)...
'   Copyright © 2006-2010, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   08/13/10    Ken Clark		Created;
'=================================================================================================================================
'Notes:
'Recommended Command-Line:	cscript VSSRename.vbs "WSRV08 VSS Database" "$/FiRRe Version 4.2"
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X VSSRename.vbs
'=================================================================================================================================
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const bWaitOnReturn = True
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const UnicodeFormat = -1
Const MB = 1048576
Dim ProjectList(), iProject, startFolder
Dim WshShell, objFSO, SS, DatabaseName, UserName, UserPassword, RootProject, Version, CurrentVSSProject
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
ReDim ProjectList(0)
iProject = 0

Private Sub LogMessage(Message)
	Dim objStdOut, objFile, LogFile, BaseName
    Set objStdOut = WScript.StdOut
	If Not IsNull(objStdOut) Then objStdOut.WriteLine Message
	
	BaseName = startFolder & "\VSSRename"
	
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
Private Function ExecuteSSwithoutOutput(Command)
	Dim oExec
    'LogMessage("Executing: " & Command)
	Set oExec = WshShell.Exec(SS & Command)
    Do
		WScript.Sleep 10
	Loop Until oExec.Status <> 0
	ExecuteSSwithoutOutput = oExec.ExitCode
	Set oExec = Nothing
End Function
Private Function ExecuteSS(Command)
	Dim oExec, oStdOut, sOutput
    'LogMessage("Executing: " & SS & Command)
	Set oExec = WshShell.Exec(SS & Command)
	Set oStdOut = oExec.StdOut
	sOutput = ""
    Do
		WScript.Sleep 10
		do until oStdOut.AtEndOfStream 
			sOutput = sOutput & oStdOut.ReadAll
		loop 
	Loop Until oExec.Status <> 0 and oStdOut.AtEndOfStream
	ExecuteSS = sOutput
	sOutput = ""
	Set oStdOut = Nothing
	Set oExec = Nothing
End Function
Private Sub OpenSourceSafe(Database, User, Password)
	Dim SCCServerPath, SSDIR
    SCCServerPath = vbNullString
	SSDIR = vbNullString
	DatabaseName = Database
	UserName = User
	UserPassword = Password
	
    strComputer = "."
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    oReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SourceSafe", "SCCServerPath", SCCServerPath
	If SCCServerPath & vbNullString = vbNullString Then
		LogMessage("Unable to locate SCCServerPath from registry")
		WScript.Quit
	End If
    oReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SourceSafe\Databases", DatabaseName, SSDIR
	If SSDIR & vbNullString = vbNullString Then
		LogMessage("Unable to find SourceSafe database named """ & DatabaseName & """ in registry")
		WScript.Quit
	End If

    BinFolder = Left(SCCServerPath, Len(SCCServerPath) - Len("\SSSCC.DLL"))
    SS = Chr(34) & BinFolder & "\SS.exe" & Chr(34) & " "	'"-Y" & UserName & "," & UserPassword & Chr(34)
    WshShell.Environment("PROCESS")("SSDIR") = SSDIR
End Sub
Private Sub GetProjectFiles(searchString)
	LogMessage("      " & searchString & "...")
    dtNow = Now()
    TimeStamp = FormatTimeStamp(dtNow)

	Dim workFile
	workFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\VSSRename.work"
	Set objFile = objFSO.OpenTextFile(workFile, ForWriting, True)
    CommandLine = "DIR " & Chr(34) & searchString & Chr(34) & " -R"
	objFile.WriteLine(ExecuteSS(CommandLine))
	objFile.Close
	
	Dim vssProject, strLine
	Set objFile = objFSO.OpenTextFile(workFile, ForReading, False)
	Do While Not objFile.AtEndOfStream
		strLine = objFile.ReadLine
		If Left(strLine, Len("No items found")) = "No items found" Then
			Do
				strLine = objFile.ReadLine
			Loop Until(Trim(strLine) = "")
		ElseIf Right(strLine, Len("item(s)")) <> "item(s)" Then
			If Left(strLine, 2) = "$/" Then
				Do While(InStr(strLine, "*") = 0)
					vssProject = strLine
					strLine = objFile.ReadLine
					strLine = vssProject & strLine
				Loop
				vssProject = mid(strline, 1, InStr(strLine, "*") - 1)
			ElseIf Trim(strLine) <> "" Then
				iProject = iProject + 1
				ReDim Preserve ProjectList(iProject) 
				ProjectList(iProject) = vssProject & Trim(strLine)
			End If
		End If
	Loop
	objFile.Close
	Set objFile = Nothing
	Set objFile = objFSO.GetFile(workFile)
	objFile.Delete
	Set objFile = Nothing
End Sub
Private Function GetSuffix(Project)
	Dim iPos
	iPos = InStrRev(Project, ".")
	If iPos <> 0 Then
		GetSuffix = LCase(Mid(Project, iPos))
	Else
		GetSuffix = vbNullString
	End If
End Function
Private Function GetVBProject(Project)
	Dim iPos
	iPos = InStrRev(Project, "/")
	If iPos <> 0 Then
		GetVBProject = Mid(Project, iPos+1)
	Else
		GetVBProject = vbNullString
	End If
End Function
Private Function GetVSSProject(Project)
	Dim VBProject, VSSProject
	VBProject = GetVBProject(Project)
	VSSProject = Mid(Project, 1, Len(Project) - Len(VBProject) - 1)
	GetVSSProject = VSSProject
End Function
Private Sub SetCurrentProject(Project)
	Dim CommandLine, VSSProject
	VSSProject = GetVSSProject(Project)
	If VSSProject = CurrentVSSProject Then Exit Sub
	
	CommandLine = "CP " & Chr(34) & VSSProject & Chr(34)	'\\WSRV08\VSS\win32\SS CP "$/FiRRe Version 4.6"
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))
End Sub
Private Sub UpdateSolution(Project, Version)
	Dim CommandLine, VSSProject, VBProject, workingFolder, workFile, sourceFile, targetFile, strLine, renamedVBProject, Solution, Suffix
	
	VSSProject = GetVSSProject(Project)					'$/FiRRe Version 4.6/FiRRe.vbproj
	VBProject = GetVBProject(Project)
	Suffix = GetSuffix(Project)
	If AlreadyRenamed(Project, Version, Suffix) Then	'$/FiRRe Version 4.6/FiRRe v4.6.vbproj
		renamedVBProject = VBProject
	Else												'$/FiRRe Version 4.6/FiRRe.vbproj
		renamedVBProject = Mid(VBProject, 1, Len(VBProject) - Len(Suffix)) & " " & LCase(Version) & Suffix
	End If
	Solution = Replace(renamedVBProject, ".vbproj", ".sln")	'Will always be already renamed...
	'Take the version off. This is what we'll search the solution for...
	VBProject = Mid(renamedVBProject, 1, Len(renamedVBProject) - Len(" " & LCase(Version) & Suffix)) & Suffix	
	
	'Before going any further, determine if we have anything to change...
	CommandLine = "FINDINFILES " & Chr(34) & VBProject & Chr(34) & " " & Chr(34) & VSSProject & "/" & Solution & Chr(34) 
	If ExecuteSSwithoutOutput(CommandLine) = 0 Then Exit Sub
	
	workingFolder = Replace(Replace(VSSProject, "$/", "V:\"), "/", "\")
	If Not objFSO.FolderExists(workingFolder) Then
		LogMessage("         Error: Working Folder [assumed] """ & workingFolder & """ does not exist!")
		Exit Sub
	End If
	WshShell.CurrentDirectory = workingFolder
	'Change the default project to our project (making command-line shorter, and required for CHECKOUT/IN operations)...
	SetCurrentProject Project
	'LogMessage("         CurrentDirectory: " & WshShell.CurrentDirectory)
	CommandLine = "CHECKOUT " & Chr(34) & Solution & Chr(34) 
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))

	LogMessage("         Updating " & Solution & "...")
	Set sourceFile = objFSO.OpenTextFile(Solution, ForReading, False)
	workFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\VSSRename.work"
	Set targetFile = objFSO.OpenTextFile(workFile, ForWriting, True)
	Do While Not sourceFile.AtEndOfStream
		strLine = sourceFile.ReadLine
		If InStr(strLine, VBProject) > 0 Then strLine = Replace(strLine, VBProject, renamedVBProject)
		targetFile.WriteLine(strLine)
	Loop
	sourceFile.Close
	targetFile.Close

	Set sourceFile = objFSO.GetFile(Solution)
	sourceFile.Delete
	Set sourceFile = Nothing
	objFSO.MoveFile workFile, Solution

	CommandLine = "CHECKIN " & Chr(34) & Solution & Chr(34) & " -C""VSSRename automated version update."""
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))
End Sub
Private Function AlreadyRenamed(Project, Version, Suffix)
	Dim VBProject
	VBProject = GetVBProject(Project)
	AlreadyRenamed = False
	'First check to see if this project has already been renamed (and skip if so)...
	If Right(LCase(Mid(Project, 1, Len(Project) - Len(Suffix))), Len(Version)) = LCase(Version) Then AlreadyRenamed = True
End Function
Private Sub DoRename(Project, Version, Suffix)
	Dim VBProject, renamedVBProject

	If AlreadyRenamed(Project, Version, Suffix) Then Exit Sub

	'Change the default project to our project (making command-line shorter, and required for CHECKOUT/IN operations)...
	SetCurrentProject Project
	VBProject = GetVBProject(Project)
	renamedVBProject = Mid(VBProject, 1, Len(VBProject) - Len(Suffix)) & " " & LCase(Version) & Suffix
	
	'\\WSRV08\VSS\win32\SS RENAME "FiRRe.vbp" "FiRRe v4.6.vbp"
	CommandLine = "RENAME " & Chr(34) & VBProject & Chr(34) & " " & Chr(34) & renamedVBProject & Chr(34)
	LogMessage("         SS " & CommandLine & " -S")	'-S)mart mode - renaming the local copy after renaming the VSS master copy.
	LogMessage("         " & ExecuteSS(CommandLine))
End Sub
Private Sub RenameProject(Project, Version)
	Dim CommandLine, VBProject, Suffix
	
	LogMessage("      " & Mid(Project, Len(RootProject)+2))

	VBProject = GetVBProject(Project)
	Suffix = GetSuffix(Project)

	DoRename Project, Version, Suffix
	
	Select Case Suffix
		Case ".vbp"
		Case ".vbproj"
			'We must also rename the associated .NET supporting project files...
			DoRename Replace(Project, ".vbproj", ".sln"), Version, ".sln"						'Solution file
			DoRename Replace(Project, ".vbproj", ".vssscc"), Version, ".vssscc"					'Visual Studio Source Control Project Metadata File
			DoRename Replace(Project, ".vbproj", ".vbproj.vspscc"), Version, ".vbproj.vspscc"	'Visual Studio Source Control Solution Metadata File
			'Lastly, we must CheckOut the solution and update any .vbproj references within to the newly renamed incarnation of the project...
			UpdateSolution Project, Version
	End Select
End Sub
Private Sub DisplayHelp
    LogMessage("Usage:")
    LogMessage("VSSRename.vbs <Database>,<RootProject> [,<UserName>, <Password>]")
    LogMessage("  Database      SourceSafe Database Name (i.e. ""WSRV08 SourceSafe Database"")")
    LogMessage("  RootProject   SourceSafe project to process [recursively] (i.e. ""$/FiRRe Version 4.6"")")
    LogMessage("  User          SourceSafe User Name (optional)")
    LogMessage("  Password      SourceSafe Password (optional)")
End Sub

LogMessage("[VSSRename.vbs" & vbTab & Now() & "]")
LogMessage("   Current Directory: " & WshShell.CurrentDirectory)
Select Case WScript.Arguments.Count
	Case 2
	Case 4
	Case Else
		DisplayHelp()
		WScript.Quit
End Select
RootProject = WScript.Arguments(1)
If InStr(RootProject, " Version ") = 0 Then
	LogMessage("Unable to determine new version number. RootProject expected to be in the form ""<Product> Version x.y""")
	WScript.Quit
End If
Version = "v" & Mid(RootProject, InStr(RootProject, " Version ") + Len(" Version "))
If WScript.Arguments.Count = 2 Then
	OpenSourceSafe WScript.Arguments(0), "", ""
Else
	OpenSourceSafe WScript.Arguments(0), WScript.Arguments(2), WScript.Arguments(3)
End If
startFolder = WshShell.CurrentDirectory

LogMessage("   Scanning " & DatabaseName & "...")
'GetProjectFiles(RootProject & "/*.vbp")
GetProjectFiles(RootProject & "/*.vbproj")
LogMessage("")
LogMessage("   Renaming Projects...")
If Not IsNull(ProjectList) Then
	For i = 1 To UBound(ProjectList)
		'LogMessage("ProjectList(" & i & "): " & ProjectList(i))
		'LogMessage("   Suffix: " & GetSuffix(ProjectList(i)))
		RenameProject ProjectList(i), Version
	Next
End If

'DoRename DatabaseName, Project, User, Password
Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
