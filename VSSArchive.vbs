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

    Set objFSO = CreateObject("Scripting.FileSystemObject")

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
Public Sub DoArchive(Database, Project, Admin, Password)
    Const HKEY_CURRENT_USER = &H80000001
    Const HKEY_LOCAL_MACHINE = &H80000002


    dtNow = Now()
    TimeStamp = FormatTimeStamp(dtNow)
    SCCServerPath = vbNullString
    VSSini = vbNullString

    strComputer = "."
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    oReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SourceSafe", "SCCServerPath", SCCServerPath
    oReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SourceSafe\Databases", DatabaseName, VSSini

    BinFolder = Left(SCCServerPath, Len(SCCServerPath) - Len("\SSSCC.DLL"))
    SSARCPath = BinFolder & "\SSARC.exe"
    BaseName = DatabaseName 
    If Project <> vbNullString Then BaseName = BaseName & "." & Project
    LogFile = BackupFolder & "\" & BaseName & "." & TimeStamp & ".log"
    ArcFile = BackupFolder & "\" & BaseName & "." & TimeStamp & ".ssa"

    CommandLine = """" & SSARCPath & """ -d- ""-s" & VSSini & """ ""-o" & LogFile & """ -i- -y" & Admin & "," & Password & " """ & ArcFile & """ $/" & Project
    Set objShell = CreateObject("WScript.Shell")
    ExitCode = objShell.Run(CommandLine, 8, True)
    'If we successfully backed-up our database, purge files older than a month (i.e. 28 days)
    If ExitCode = 0 Then CleanUp ArcFile	', 28
End Sub
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript VSSArchive.vbs //X

BackupFolder = GetEnvironmentVariable("BackupFolder")
AltBackupFolder = "E" & Mid(BackupFolder, 2)
DatabaseName = "Ken's Home VSS Database"
Project = ""
'Project = "VSSarchive.NET"	'for Testing...
Admin = "Admin"
Password = "tomcat"
DoArchive DatabaseName, Project, Admin, Password
