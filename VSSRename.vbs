'VSSRename.vbs
'	Visual Basic Script Used to Automate New Product Version Creation (starting with renaming projects)...
'   Copyright © 2006-2010, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   08/13/10    Ken Clark		Created;
'=================================================================================================================================
'Recommended Command-Line:	cscript VSSRename.vbs "WSRV08 VSS Database" "$/FiRRe Version 4.2"
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X VSSRename.vbs
'=================================================================================================================================
'Notes:
'\\WSRV08\VSS\win32\SS DIR "$/Components Version 4.6/*.vbp" -R
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/CRUFLSIA/CRUFLSIA.vbp" "$/Components Version 4.6/CRUFLSIA/CRUFLSIA v4.6.vbp" -S
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/mapirtf/VBSource/vbmaprtf.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASADO/SIASADO.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASCL/SIASCL.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASCurrency/SIASCurrency.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASDB/SIASDB.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASEMail/SIASEmail.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRegisterDLLs/SIASRegisterDLLs.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRPC/SIASRPC.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRPC/SIASBPE00000/SIASBPE00000.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRPC/SIASBPE00001/SIASBPE00001.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRPC/SIASBPE21130/SIASBPE21130.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRPC/SIASBPE21140/SIASBPE21140.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRPC/SIASBPE21150/SIASBPE21150.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRPC.NET/SIASRPCDemo/SIASRPCDemo.vbp
'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASUTL/SIASUTL.vbp
'=================================================================================================================================
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const bWaitOnReturn = True
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const UnicodeFormat = -1
Const MB = 1048576
Dim ProjectList(), iProject
Dim WshShell, objFSO, SS, DatabaseName, UserName, UserPassword, RootProject, Version
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
ReDim ProjectList(0)
iProject = 0

Private Sub LogMessage(Message)
	Dim objStdOut, objFile, LogFile, BaseName
    Set objStdOut = WScript.StdOut
	If Not IsNull(objStdOut) Then objStdOut.WriteLine Message
	
	BaseName = "VSSRename"
	
	LogFile = BaseName & ".log"
    If objFSO.FileExists(LogFile) Then
		Set objFile = objFSO.GetFile(LogFile)
		If objFile.Size > 10*MB Then
            Dim dtModified, NewFileName
            dtModified = objFile.DateLastModified
            NewFileName = BackupFolder & "\" & BaseName & "." & FormatTimeStamp(dtModified) & ".log"
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
Public Function FormatTimeStamp(TimeStamp)
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
Public Function GetEnvironmentVariable(VariableName)
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
Public Sub CleanUp(FileName)
    Const DeleteReadOnly = TRUE
	Const wbemFlagReturnImmediately = &h10
	Const wbemFlagForwardOnly = &h20

    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")
    Set objSSA = objFSO.GetFile(FileName)
    BaseName = objSSA.ParentFolder & "\" & objFSO.GetBaseName(objSSA) & "."
    BackupName = Mid(BaseName, 1, Len(BaseName) - Len("yyyyMMdd-HHmmss."))
    SQLSource = "Select * from CIM_DataFile where Path='\\" & Replace(Mid(objSSA.ParentFolder, 4), "\", "\\") & "\\'"	' And CreationDate < '" & objSSA.CreationDate & "'""
    Set colFiles = objWMIService.ExecQuery(SQLSource, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
    For Each objFile in colFiles
		If Left(objFile.Name, Len(BackupName)) = LCase(BackupName) And Left(objFile.Name, Len(BaseName)) <> LCase(BaseName) Then 
			'Wscript.Echo objFile.Name
			objFSO.DeleteFile(objFile.Name), DeleteReadOnly
		End If
    Next    
End Sub
Public Function Execute(Command)
	Dim oExec, oStdOut, sOutput
    'LogMessage("Executing: " & Command)
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
Public Sub OpenSourceSafe(Database, User, Password)
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
Public Sub GetProjectFiles(searchString)	'(Database, User, Password, Projects)
	LogMessage("      " & searchString & "...")
    dtNow = Now()
    TimeStamp = FormatTimeStamp(dtNow)

	workFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\VSSRename.work"
	Set objFile = objFSO.OpenTextFile(workFile, ForWriting, True)
    CommandLine = SS & " DIR " & Chr(34) & searchString & Chr(34) & " -R"
	objFile.WriteLine(Execute(CommandLine))
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
Public Function GetSuffix(Project)
	Dim iPos
	iPos = InStrRev(Project, ".")
	If iPos <> 0 Then
		GetSuffix = LCase(Mid(Project, iPos))
	Else
		GetSuffix = vbNullString
	End If
End Function
Public Function GetVBProject(Project)
	Dim iPos
	iPos = InStrRev(Project, "/")
	If iPos <> 0 Then
		GetVBProject = Mid(Project, iPos)
	Else
		GetVBProject = vbNullString
	End If
End Function
Public Sub DoRename(Project, Version)
	Dim CommandLine, VBProject, Suffix
	
	LogMessage("      " & Mid(Project, Len(RootProject)+2))
	VBProject = GetVBProject(Project)
	Suffix = GetSuffix(Project)
	'First check to see if this project has already been renamed (and skip if so)...
	If Right(LCase(Mid(Project, 1, Len(Project) - Len(Suffix))), Len(Version)) = LCase(Version) Then
		'LogMessage("         Project already processed!")
		Exit Sub
	End If
	
	'\\WSRV08\VSS\win32\SS RENAME "$/Components Version 4.6/SIASRegisterDLLs/SIASRegisterDLLs.vbp" "$/Components Version 4.6/SIASRegisterDLLs/SIASRegisterDLLs v4.6.vbp"
	CommandLine = "RENAME " & Chr(34) & Project & Chr(34) & " " & Chr(34) & Mid(Project, 1, Len(Project) - Len(Suffix)) & " " & LCase(Version) & Suffix & Chr(34)
	LogMessage("         SS " & CommandLine)
	'LogMessage("         " & Execute(CommandLine))
	'Select Case Suffix
'		Case ".vbp"
'			If Right(LCase(Project), Len(" v4.6.vbp")) = " v4.6.vbp" Then
'			End If
'		Case ".vbproj"
'			If Right(LCase(Project), Len(" v4.6.vbproj")) = " v4.6.vbproj" Then
'			End If
'	End Select
End Sub
Private Sub DisplayHelp
    LogMessage "Usage:"
    LogMessage "VSSRename.vbs <Database>,<RootProject> [,<UserName>, <Password>]"
    LogMessage "  Database      SourceSafe Database Name (i.e. ""WSRV08 SourceSafe Database"")"
    LogMessage "  RootProject   SourceSafe project to process [recursively] (i.e. ""$/FiRRe Version 4.6"")"
    LogMessage "  User          SourceSafe User Name (optional)"
    LogMessage "  Password      SourceSafe Password (optional)"
End Sub

LogMessage "[VSSRename.vbs" & vbTab & Now() & "]"
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

LogMessage("   Scanning " & DatabaseName & "...")
GetProjectFiles(RootProject & "/*.vbp")
GetProjectFiles(RootProject & "/*.vbproj")
LogMessage("")
LogMessage("   Renaming Projects...")
If Not IsNull(ProjectList) Then
	For i = 1 To UBound(ProjectList)
		'LogMessage("ProjectList(" & i & "): " & ProjectList(i))
		'LogMessage("   Suffix: " & GetSuffix(ProjectList(i)))
		DoRename ProjectList(i), Version
	Next
End If

'DoRename DatabaseName, Project, User, Password
Set objFSO = Nothing
WScript.Quit
