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
Dim WshShell, objFSO, SS, DatabaseName, UserName, UserPassword, RootProject, Version, PriorVersion, CurrentVSSProject
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
    CommandLine = "DIR " & Chr(34) & searchString & Chr(34) & " -I- -R"
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
				If Right(strLine, 1) <> ":" Then
					Do
						strLine = objFile.ReadLine
					Loop Until Right(strLine, 1) = ":"
				End If
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
Private Function GetProduct(Project)
	Dim iPos
	iPos = InStr(Project, " Version ")
	If iPos <> 0 Then
		GetProduct = Mid(Project, Len("$/")+1, iPos-(Len("$/")+1))
	Else
		GetProduct = vbNullString
	End If
End Function
Private Function GetProjectFileName(Project)
	Dim iPos
	iPos = InStrRev(Project, "/")
	If iPos <> 0 Then
		GetProjectFileName = Mid(Project, iPos+1)
	Else
		GetProjectFileName = vbNullString
	End If
End Function
Private Function GetVSSProject(Project)
	Dim ProjectFileName, VSSProject
	ProjectFileName = GetProjectFileName(Project)
	VSSProject = Mid(Project, 1, Len(Project) - Len(ProjectFileName) - 1)
	GetVSSProject = VSSProject
End Function
Private Sub SetCurrentProject(Project)
	Dim CommandLine, VSSProject, sOutput
	VSSProject = GetVSSProject(Project)
	If VSSProject = CurrentVSSProject Then Exit Sub
	
	CommandLine = "CP " & Chr(34) & VSSProject & Chr(34) & " -I-"	'\\WSRV08\VSS\win32\SS CP "$/FiRRe Version 4.6"
	'LogMessage("         SS " & CommandLine)
	'LogMessage("         " & ExecuteSS(CommandLine))
	sOutput = ExecuteSS(CommandLine)
	if Trim(sOutput) <> vbNullString Then LogMessage("         " & sOutput)
	sOutput = vbNullString
End Sub
Private Sub UpdateProject(Project, Version)
	Dim CommandLine, Product, VSSProject, ProjectFileName, workingFolder, workFile, sourceFile, targetFile, strLine, Suffix
	Dim searchOutputPath, outputPath, searchHintPath, hintPath
	'Project:											'$/FiRRe Version 4.6/FiRRe v4.6.vbproj
	'Version:											'v4.6
	Product = GetProduct(Project)						'FiRRe
	VSSProject = GetVSSProject(Project)					'$/FiRRe Version 4.6
	ProjectFileName = GetProjectFileName(Project)		'FiRRe v4.6.vbproj
	Suffix = GetSuffix(Project)							'.vbproj
	
	If Not AlreadyRenamed(Project, Version, Suffix) Then ProjectFileName = Mid(ProjectFileName, 1, Len(ProjectFileName) - Len(Suffix)) & " " & LCase(Version) & Suffix	'FiRRe v4.6.vbproj
	
	'What are we looking for...?
	Select Case Product
		Case "Components"
			'We have to look for OutputPath property of the <Config> tag...
			searchOutputPath = "FiRRe\program files\"
			outputPath = "FiRRe " & Version & "\program files\"
			searchHintPath = vbNullString
			hintPath = vbNullString
		Case "FiRRe"
			'We have to look for OutputPath property of the <Config> tag...
			searchOutputPath = "FiRRe\program files\"
			outputPath = "FiRRe " & Version & "\program files\"
			'We also have to look for Component folder references in the HintPath property of the <Reference> tag...
			'Note that <Reference> tags contain relative paths...
			searchHintPath = "..\..\..\Components\"
			hintPath = "..\..\..\Components Version " & Mid(Version, 2) & "\"
		Case Else
			LogMessage("         Error: Unexpected Product (" & Product & "); Project not updated!")
			Exit Sub
	End Select	
	
	'Before going any further, determine if we have anything to change...
	'Note: Last "\" is doubled so as not to confuse it with and escaped-'"'
	CommandLine = "FINDINFILES " & Chr(34) & searchOutputPath & "\" & Chr(34) & " " & Chr(34) & VSSProject & "/" & ProjectFileName & Chr(34) & " -I-" 
	If ExecuteSSwithoutOutput(CommandLine) = 0 Then
		If hintPath = vbNullString Then Exit Sub
		'Last "\" is doubled so as not to confuse it with and escaped-'"'
		CommandLine = "FINDINFILES " & Chr(34) & searchHintPath & "\" & Chr(34) & " " & Chr(34) & VSSProject & "/" & ProjectFileName & Chr(34) & " -I-" 
		If ExecuteSSwithoutOutput(CommandLine) = 0 Then Exit Sub
	End If
	
	workingFolder = Replace(Replace(VSSProject, "$/", "V:\"), "/", "\")
	If Not objFSO.FolderExists(workingFolder) Then
		LogMessage("         Error: Working Folder [assumed] """ & workingFolder & """ does not exist! Get latest version of project and run this utility again.")
		Exit Sub
	End If
	WshShell.CurrentDirectory = workingFolder
	'Change the default project to our project (making command-line shorter, and required for CHECKOUT/IN operations)...
	SetCurrentProject Project
	'LogMessage("         CurrentDirectory: " & WshShell.CurrentDirectory)
	CommandLine = "CHECKOUT " & Chr(34) & ProjectFileName & Chr(34) & " -I-"
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))

	LogMessage("         Updating " & ProjectFileName & "...")
	Set sourceFile = objFSO.OpenTextFile(ProjectFileName, ForReading, False)
	workFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\VSSRename.work"
	Set targetFile = objFSO.OpenTextFile(workFile, ForWriting, True)
	Do While Not sourceFile.AtEndOfStream
		strLine = sourceFile.ReadLine
		
		If InStr(strLine, searchOutputPath) > 0 Then strLine = Replace(strLine, searchOutputPath, outputPath)
		If hintPath <> vbNullString Then
			If InStr(strLine, searchHintPath) > 0 Then strLine = Replace(strLine, searchHintPath, hintPath)
		End If
		targetFile.WriteLine(strLine)
	Loop
	sourceFile.Close
	targetFile.Close

	Set sourceFile = objFSO.GetFile(ProjectFileName)
	sourceFile.Delete
	Set sourceFile = Nothing
	objFSO.MoveFile workFile, ProjectFileName

	CommandLine = "CHECKIN " & Chr(34) & ProjectFileName & Chr(34) & " -I- -C""VSSRename automated version update."""
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))
End Sub
Private Sub UpdateSolution(Project, Version)
	Dim CommandLine, VSSProject, ProjectFileName, workingFolder, workFile, sourceFile, targetFile, strLine, renamedProjectFile, Solution, Suffix
	
	VSSProject = GetVSSProject(Project)					'$/FiRRe Version 4.6/FiRRe.vbproj
	ProjectFileName = GetProjectFileName(Project)
	Suffix = GetSuffix(Project)
	If AlreadyRenamed(Project, Version, Suffix) Then	'$/FiRRe Version 4.6/FiRRe v4.6.vbproj
		renamedProjectFile = ProjectFileName
	Else												'$/FiRRe Version 4.6/FiRRe.vbproj
		renamedProjectFile = Mid(ProjectFileName, 1, Len(ProjectFileName) - Len(Suffix)) & " " & LCase(Version) & Suffix
	End If
	Solution = Replace(renamedProjectFile, ".vbproj", ".sln")	'Will always be already renamed...
	'Take the version off. This is what we'll search the solution for...
	ProjectFileName = Mid(renamedProjectFile, 1, Len(renamedProjectFile) - Len(" " & LCase(Version) & Suffix)) & Suffix	
	
	'Before going any further, determine if we have anything to change...
	CommandLine = "FINDINFILES " & Chr(34) & ProjectFileName & Chr(34) & " " & Chr(34) & VSSProject & "/" & Solution & Chr(34) & " -I-" 
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
	CommandLine = "CHECKOUT " & Chr(34) & Solution & Chr(34) & " -I-"
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))

	LogMessage("         Updating " & Solution & "...")
	Set sourceFile = objFSO.OpenTextFile(Solution, ForReading, False)
	workFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\VSSRename.work"
	Set targetFile = objFSO.OpenTextFile(workFile, ForWriting, True)
	Do While Not sourceFile.AtEndOfStream
		strLine = sourceFile.ReadLine
		If InStr(strLine, ProjectFileName) > 0 Then strLine = Replace(strLine, ProjectFileName, renamedProjectFile)
		targetFile.WriteLine(strLine)
	Loop
	sourceFile.Close
	targetFile.Close

	Set sourceFile = objFSO.GetFile(Solution)
	sourceFile.Delete
	Set sourceFile = Nothing
	objFSO.MoveFile workFile, Solution

	CommandLine = "CHECKIN " & Chr(34) & Solution & Chr(34) & " -I- -C""VSSRename automated version update."""
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))
End Sub
Private Sub UpdateInstallShield(Project, Version, PriorVersion)
	Dim CommandLine, Product, VSSProject, ProjectFileName, workingFolder, workFile, sourceFile, targetFile, strLine, Suffix
	Dim searchSccPath, SccPath, searchString

	'Project:											'$/FiRRe Version 4.6/InstallShield/FiRRe.ism
	'Version:											'v4.6
	Product = GetProduct(Project)						'FiRRe
	VSSProject = GetVSSProject(Project)					'$/FiRRe Version 4.6/InstallShield
	ProjectFileName = GetProjectFileName(Project)		'FiRRe.ism
	Suffix = GetSuffix(Project)							'.ism

	If Not AlreadyRenamed(Project, Version, Suffix) Then ProjectFileName = Mid(ProjectFileName, 1, Len(ProjectFileName) - Len(Suffix)) & " " & LCase(Version) & Suffix	'FiRRe v4.6.ism

	'What are we looking for...?
	searchSccPath = "$/" & Product & "/InstallShield"	'<row><td>SccPath</td><td>"$/FiRRe/InstallShield", ESZAAAAA</td></row>
	SccPath = RootProject & "/InstallShield"

	'Before going any further, determine if we have anything to change...
	CommandLine = "FINDINFILES " & Chr(34) & searchSccPath & Chr(34) & " " & Chr(34) & VSSProject & "/" & ProjectFileName & Chr(34) & " -I-"
	If ExecuteSSwithoutOutput(CommandLine) = 0 Then
		CommandLine = "FINDINFILES " & Chr(34) & "\" & PriorVersion & "<" & Chr(34) & " " & Chr(34) & VSSProject & "/" & ProjectFileName & Chr(34) & " -I-"
		If ExecuteSSwithoutOutput(CommandLine) = 0 Then
			CommandLine = "FINDINFILES " & Chr(34) & "\SunGard\" & Product & "<" & Chr(34) & " " & Chr(34) & VSSProject & "/" & ProjectFileName & Chr(34) & " -I-"
			If ExecuteSSwithoutOutput(CommandLine) = 0 Then
				CommandLine = "FINDINFILES " & Chr(34) & "\Projects\" & Product & "<" & Chr(34) & " " & Chr(34) & VSSProject & "/" & ProjectFileName & Chr(34) & " -I-"
				If ExecuteSSwithoutOutput(CommandLine) = 0 Then Exit Sub
			End If
		End If
	End If
	
	workingFolder = Replace(Replace(VSSProject, "$/", "V:\"), "/", "\")
	If Not objFSO.FolderExists(workingFolder) Then
		LogMessage("         Error: Working Folder [assumed] """ & workingFolder & """ does not exist! Get latest version of project and run this utility again.")
		Exit Sub
	End If
	WshShell.CurrentDirectory = workingFolder
	'Change the default project to our project (making command-line shorter, and required for CHECKOUT/IN operations)...
	SetCurrentProject Project
	'LogMessage("         CurrentDirectory: " & WshShell.CurrentDirectory)
	CommandLine = "CHECKOUT " & Chr(34) & ProjectFileName & Chr(34) & " -I-"
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))

	LogMessage("         Updating " & ProjectFileName & "...")
	Set sourceFile = objFSO.OpenTextFile(ProjectFileName, ForReading, False)
	workFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\VSSRename.work"
	Set targetFile = objFSO.OpenTextFile(workFile, ForWriting, True)
	
	Dim TableName, PropertyName
	TableName = vbNullString
	
	Do While Not sourceFile.AtEndOfStream
		strLine = sourceFile.ReadLine
		If Left(Trim(strLine), Len("<table name=""")) = "<table name=""" Then TableName = Mid(Trim(strLine), Len("<table name=""")+1, Len(Trim(strLine))-Len("<table name=""")-2)
		If Trim(strLine) = "</table>" Then TableName = vbNullString

		'Note: We're not dealing with the Directory or related data in the Component table as these should be updated by 
		'      InstallShield itself the first time the project is built after being renamed...
		
		If TableName = "InstallShield" Then
			'<row><td>SccPath</td><td>"$/FiRRe/InstallShield", ESZAAAAA</td></row>
			If InStr(strLine, "SccPath") > 0 Then strLine = Replace(strLine, searchSccPath, SccPath)
		End If
		If TableName = "ISPathVariable" Then
			searchString = "SunGard Shared\" & PriorVersion & "<"						'<row><td>ComponentSource</td><td><AppServerFolder>\Common\SunGard Shared\v4.3</td><td/><td>2</td></row>
			If Left(Trim(strLine), Len("<row><td>ComponentSource")) = "<row><td>ComponentSource" And InStr(strLine, searchString) > 0 Then strLine = Replace(strLine, searchString, "SunGard Shared\" & Version & "<")
			searchString = "\SunGard\" & Product & "<"									'<row><td>AppServerFolder</td><td>\\WSRV08\SunGard\FiRRe</td><td/><td>2</td></row>
			If Left(Trim(strLine), Len("<row><td>AppServerFolder")) = "<row><td>AppServerFolder" And InStr(strLine, searchString) > 0 Then strLine = Replace(strLine, searchString, "\SunGard\" & Product & " " & Version & "<")
			searchString = "<AppServerFolder>\program files\SunGard\" & Product & "<"	'<row><td>FiRReExePath</td><td><AppServerFolder>\program files\SunGard\FiRRe</td><td/><td>2</td></row>
			If Left(Trim(strLine), Len("<row><td>" & Product & "ExePath")) = "<row><td>" & Product & "ExePath" And InStr(strLine, searchString) > 0 Then strLine = Replace(strLine, searchString, "<AppServerFolder>\program files\SunGard\" & Product & " " & Version & "<")
			searchString = "\Projects\" & Product & "<"									'<row><td>FiRReProject</td><td>\\WSRV08\Projects\FiRRe</td><td/><td>2</td></row>
			If Left(Trim(strLine), Len("<row><td>" & Product & "Project")) = "<row><td>" & Product & "Project" And InStr(strLine, searchString) > 0 Then strLine = Replace(strLine, searchString, "\Projects\" & Product & " Version " & mid(Version, 2) & "<")
		End If
		If TableName = "ISRelease" Then
			searchString = ">" & PriorVersion & "<"										'<row><td>v4.3</td><td>BNY</td><td>C:\InstallShield\FiRRe</td><td>FiRRe</td><td>1</td><td>1033</td><td>2</td><td>2</td><td>Intel</td><td/><td>1033</td><td>3</td><td>0</td><td>0</td><td>0</td><td/><td>0</td><td/><td>\\WSRV08\InstallShield\FiRRe\BNYMv4365</td><td/><td>http://</td><td/><td/><td/><td/><td>73741</td><td/><td/><td/><td/></row>
			If Left(Trim(strLine), Len("<row><td>" & PriorVersion)) = "<row><td>" & PriorVersion And InStr(strLine, searchString) > 0 Then strLine = Replace(strLine, searchString, ">" & Version & "<")
		End If
		If TableName = "ISReleaseExtended" Then
			searchString = ">" & PriorVersion & "<"										'<row><td>v4.3</td><td>BNY</td><td>0</td><td>http://</td><td>0</td><td>install</td><td>install</td><td>[WindowsFolder]Downloaded Installations</td><td>1</td><td>http://www.installengine.com/Msiengine20</td><td>http://www.installengine.com/Msiengine20</td><td>1</td><td>http://www.installengine.com/cert05/isengine</td><td/><td/><td/><td/><td>1</td><td>http://www.installengine.com/cert05/dotnetfx</td><td>1</td><td>1033</td><td/><td/><td/><td/><td>24</td><td>3</td><td>20</td><td/><td/></row>
			If Left(Trim(strLine), Len("<row><td>" & PriorVersion)) = "<row><td>" & PriorVersion And InStr(strLine, searchString) > 0 Then strLine = Replace(strLine, searchString, ">" & Version & "<")
		End If
		If TableName = "Property" Then
			'<row><td>ProductName</td><td>FiRRe Version 4.3.65</td><td/></row>
			If Left(Trim(strLine), Len("<row><td>ProductName</td>")) = "<row><td>ProductName</td>" Then strLine = "<row><td>ProductName</td><td>" & Product & " Version " & mid(Version, 2) & ".0</td><td/></row>"
			'<row><td>ProductVersion</td><td>4.3.65</td><td/></row>
			If Left(Trim(strLine), Len("<row><td>ProductVersion</td>")) = "<row><td>ProductVersion</td>" Then strLine = "<row><td>ProductVersion</td><td>" & mid(Version, 2) & ".0</td><td/></row>"
		End If
		If TableName = "ISString" Then
			searchString = "|" & Product & " " & PriorVersion & "<"						'<row><td>S_FiRRe_ShortLongName</td><td>1033</td><td>FIRREV~1.3|FiRRe v4.3</td><td>0</td><td/><td>-1801705073</td></row>
			If Left(Trim(strLine), Len("<row><td>S_" & Product & "_ShortLongName")) = "<row><td>S_" & Product & "_ShortLongName" And InStr(strLine, searchString) > 0 Then strLine = Replace(strLine, searchString, "|" & Product & " " & Version & "<")
		End If
		targetFile.WriteLine(strLine)
	Loop
	sourceFile.Close
	targetFile.Close

	Set sourceFile = objFSO.GetFile(ProjectFileName)
	sourceFile.Delete
	Set sourceFile = Nothing
	objFSO.MoveFile workFile, ProjectFileName

	CommandLine = "CHECKIN " & Chr(34) & ProjectFileName & Chr(34) & " -I- -C""VSSRename automated version update."""
	LogMessage("         SS " & CommandLine)
	LogMessage("         " & ExecuteSS(CommandLine))
End Sub
Private Function AlreadyRenamed(Project, Version, Suffix)
	AlreadyRenamed = False
	If Right(LCase(Mid(Project, 1, Len(Project) - Len(Suffix))), Len(Version)) = LCase(Version) Then AlreadyRenamed = True
End Function
Private Sub DoRename(Project, Version, Suffix)
	Dim FileName, renamedProjectFile, sOutput

	If AlreadyRenamed(Project, Version, Suffix) Then Exit Sub

	'Change the default project to our project (making command-line shorter, and required for CHECKOUT/IN operations)...
	SetCurrentProject Project
	FileName = GetProjectFileName(Project)
	renamedProjectFile = Mid(FileName, 1, Len(FileName) - Len(Suffix)) & " " & LCase(Version) & Suffix
	
	'\\WSRV08\VSS\win32\SS RENAME "FiRRe.vbp" "FiRRe v4.6.vbp"
	CommandLine = "RENAME " & Chr(34) & FileName & Chr(34) & " " & Chr(34) & renamedProjectFile & Chr(34) & " -I- -S"	'-S)mart mode - renaming the local copy after renaming the VSS master copy.
	LogMessage("         SS " & CommandLine)
	'LogMessage("         " & ExecuteSS(CommandLine))
	sOutput = ExecuteSS(CommandLine)
	if Trim(sOutput) <> vbNullString Then LogMessage("         " & sOutput)
	sOutput = vbNullString
End Sub
Private Sub RenameProject(Project, Version, PriorVersion)
	Dim CommandLine, Suffix
	
	LogMessage("      " & Mid(Project, Len(RootProject)+2))
	Suffix = GetSuffix(Project)
	Select Case Suffix
		Case ".vbp"
			'Rename the VB6 project through SourceSafe...
			DoRename Project, Version, Suffix
		Case ".vbproj"
			'Rename the VB.NET project through SourceSafe...
			DoRename Project, Version, Suffix
			'We must also rename the associated .NET supporting project files...
			DoRename Replace(Project, ".vbproj", ".sln"), Version, ".sln"						'Solution file
			DoRename Replace(Project, ".vbproj", ".vssscc"), Version, ".vssscc"					'Visual Studio Source Control Project Metadata File
			DoRename Replace(Project, ".vbproj", ".vbproj.vspscc"), Version, ".vbproj.vspscc"	'Visual Studio Source Control Solution Metadata File
			'Next, we must update path references in the .vbproj file to reflect the new version
			'	Note: This is rather hard-coded for the FiRRe/Components relationship)...
			UpdateProject Project, Version
			'Lastly, we must CheckOut the solution and update any .vbproj references within to the newly renamed incarnation of the project...
			UpdateSolution Project, Version
		Case ".ism"
			'Rename the InstallShield project through SourceSafe...
			DoRename Project, Version, Suffix
			UpdateInstallShield Project, Version, PriorVersion
	End Select
End Sub
Private Sub DisplayHelp
    LogMessage("Usage:")
    LogMessage("VSSRename.vbs <Database>,<RootProject>,<PriorVersion>[,<UserName>, <Password>]")
    LogMessage("  Database      SourceSafe Database Name (i.e. ""WSRV08 SourceSafe Database"")")
    LogMessage("  RootProject   SourceSafe project to process [recursively] (i.e. ""$/FiRRe Version 4.6"")")
    LogMessage("  PriorVersion  Prior version number in the form <MajorVersion>.<MinorVersion> (i.e. ""4.5"")")
    LogMessage("  User          SourceSafe User Name (optional)")
    LogMessage("  Password      SourceSafe Password (optional)")
End Sub

LogMessage("[VSSRename.vbs" & vbTab & Now() & "]")
Select Case WScript.Arguments.Count
	Case 3
	Case 5
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
PriorVersion = "v" & WScript.Arguments(2)
If WScript.Arguments.Count = 3 Then
	OpenSourceSafe WScript.Arguments(0), "", ""
Else
	OpenSourceSafe WScript.Arguments(0), WScript.Arguments(2), WScript.Arguments(3)
End If
LogMessage("   Current Directory: " & WshShell.CurrentDirectory)
startFolder = WshShell.CurrentDirectory

LogMessage("   Scanning " & DatabaseName & "...")
GetProjectFiles(RootProject & "/*.vbp")
GetProjectFiles(RootProject & "/*.vbproj")
GetProjectFiles(RootProject & "/*.ism")
'GetProjectFiles(RootProject & "/*.vbproj v4.2.vspscc")
LogMessage("")
LogMessage("   Renaming Projects...")
If Not IsNull(ProjectList) Then
	For i = 1 To UBound(ProjectList)
		'LogMessage("ProjectList(" & i & "): " & ProjectList(i))
		'LogMessage("   Suffix: " & GetSuffix(ProjectList(i)))
		RenameProject ProjectList(i), Version, PriorVersion
	Next
End If
Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
