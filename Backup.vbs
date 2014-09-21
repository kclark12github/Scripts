'Ntbackup
'Perform backup operations at a command prompt or from a batch file using the ntbackup command followed by various parameters.
'
'Syntax
'   ntbackup backup [systemstate] "@bks file name" /J {"job name"} [/P {"pool name"}] [/G {"guid name"}] [/T { "tape name"}] [/N {"media name"}] [/F {"file name"}] [/D {"set description"}] [/DS {"server name"}] [/IS {"server name"}] [/A] [/V:{yes|no}] [/R:{yes|no}] [/L:{f|s|n}] [/M {backup type}] [/RS:{yes|no}] [/HC:{on|off}] [/SNAP:{on|off}]
'
'Parameters
'   systemstate             Specifies that you want to back up the System State data. When you select this option, the backup type will be forced to 
'                           normal or copy. 
'   @bks file name          Specifies the name of the backup selection file (.bks file) to be used for this backup operation. The at (@) character must 
'                           precede the name of the backup selection file. A backup selection file contains information on the files and folders you 
'                           have selected for backup. You have to create the file using the graphical user interface (GUI) version of Backup. 
'   /J {"job name"}         Specifies the job name to be used in the log file. The job name usually describes the files and folders you are backing up 
'                           in the current backup job as well as the date and time you backed up the files. 
'   /P {"pool name"}        Specifies the media pool from which you want to use media. This is usually a subpool of the Backup media pool, such as 4mm 
'                           DDS. If you select this you cannot use the /A, /G, /F, or /T command-line options. 
'   /G {"guid name"}        Overwrites or appends to this tape. Do not use this switch in conjunction with /P. 
'   /T {"tape name"}        Overwrites or appends to this tape. Do not use this switch in conjunction with /P. 
'   /N {"media name"}       Specifies the new tape name. You must not use /A with this switch. 
'   /F {"file name"}        Logical disk path and file name. You must not use the following switches with this switch: /P /G /T. 
'   /D {"set description"}  Specifies a label for each backup set. 
'   /DS {"server name"}     Backs up the directory service file for the specified Microsoft Exchange Server. 
'   /IS {"server name"}     Backs up the Information Store file for the specified Microsoft Exchange Server. 
'   /A                      Performs an append operation. Either /G or /T must be used in conjunction with this switch. Do not use this switch in 
'                           conjunction with /P. 
'   /V:{yes|no}             Verifies the data after the backup is complete. 
'   /R:{yes|no}             Restricts access to this tape to the owner or members of the Administrators group. 
'   /L:{f|s|n}              Specifies the type of log file: f=full, s=summary, n=none (no log file is created). 
'   /M {backup type}        Specifies the backup type. It must be one of the following: normal, copy, differential, incremental, or daily. 
'   /RS:{yes|no}            Backs up the migrated data files located in Remote Storage. The /RS command-line option is not required to back up the local 
'                           Removable Storage database (that contains the Remote Storage placeholder files). When you backup the %systemroot% folder, 
'                           Backup automatically backs up the Removable Storage database as well. 
'   /HC:{on|off}            Uses hardware compression, if available, on the tape drive. 
'   /SNAP:{on|off}          Specifies whether or not the backup is a volume shadow copy. 
'   /M {backup type}        Specifies the backup type. It must be one of the following: normal, copy, differential, incremental, or daily. 
'   /?                      Displays help at the command prompt. 
'Remarks
'   - You cannot restore files from the command line using the ntbackup command. 
'   - The following command-line options default to what you have already set using the graphical user interface (GUI) version of Backup unless they are 
'     changed by a command-line option: /V /R /L /M /RS /HC. For example, if hardware compression is turned on in the Options dialog box in Backup, it 
'     will be used if /HC is not specified on the command line. However, if you specify /HC:off at the command line, it overrides the Option dialog box 
'     setting and compression is not used. 
'   - If you have Windows Media Services running on your computer, and you want to back up the files associated with these services, see "Running Backup 
'     with Windows Media Services" in the Windows Media Services online documentation. You must follow the procedures outlined in the Windows Media 
'     Services online documentation before you can back up or restore files associated with Windows Media Services. 
'   - You can only back up the System State data on a local computer. You cannot back up the System State data on a remote computer 
'   - If you are using Removable Storage to manage media, or you are using the Remote Storage to store data, then you should regularly back up the files 
'     that are in the following folders: 
'       Systemroot\System32\Ntmsdata
'       Systemroot\System32\Remotestorage
'     This ensures that all Removable Storage and Remote Storage data can be restored.
Public Function IsDST()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	For Each objItem In colItems
	  'WScript.Echo "Current Time Zone (Hours Offset From GMT): " & (objItem.CurrentTimeZone / 60)
	  'WScript.Echo "Daylight Saving In Effect: " & objItem.DaylightInEffect
	  IsDST = objItem.DaylightInEffect
	  Exit Function
	Next
End Function
Public Function BaseCTime()
	'CTime := # Seconds since 01/01/1970 GMT (adjusted for EST/EDT)...
	If IsDST() Then
		BaseCTime = #12/31/1969 8:00:00 PM#
	Else
		BaseCTime = #12/31/1969 7:00:00 PM#
	End If
End Function
Public Function FormatCTime(CTime)
	TimeStamp = DateAdd("s", CTime, BaseCTime())
    iYear = Year(TimeStamp)
    iMonth = Month(TimeStamp)
    iDay = Day(TimeStamp)
    iHour = Hour(TimeStamp)
    iMinute = Minute(TimeStamp)
    iSecond = Second(TimeStamp)
    If iHour > 12 Then AMPM = "PM":iHour = iHour - 12 Else AMPM = "AM"
    If iHour = 0 then iHour = "12"
    
    if iMonth < 10 then FormatCTime = FormatCTime & "0"
    FormatCTime = FormatCTime & iMonth & "/"
    if iDay < 10 then FormatCTime = FormatCTime & "0"
    FormatCTime = FormatCTime & iDay & "/"
    FormatCTime = FormatCTime & iYear
    
    FormatCTime = FormatCTime & " "
    if iHour < 10 then FormatCTime = FormatCTime & "0"
    FormatCTime = FormatCTime & iHour & ":"
    if iMinute < 10 then FormatCTime = FormatCTime & "0"
    FormatCTime = FormatCTime & iMinute & ":"
    if iSecond < 10 then FormatCTime = FormatCTime & "0"
    FormatCTime = FormatCTime & iSecond
    FormatCTime = FormatCTime & " " & AMPM
End Function
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
Public Function GetLogFile(StartTimeStamp, MyJobName)
    Const LOCAL_APPLICATION_DATA = &H1c&
    Const HKEY_CURRENT_USER = &H80000001
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const REG_SZ = 1
    Const REG_EXPAND_SZ = 2
    Const REG_BINARY = 3
    Const REG_DWORD = 4
    Const REG_MULTI_SZ = 7
    
    'Identify the Log file just created for relocation to store with the Backup file itself...    
	GetLogFile = vbNullString	'Return value if no log greater than our StartTimeStamp is found...
    strComputer = "."
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    oReg.GetDWORDValue HKEY_CURRENT_USER, "Software\Microsoft\Ntbackup\Log Files", "Log File Count", LogFileCount
    
    oReg.EnumKey HKEY_CURRENT_USER, "Software\Microsoft\Ntbackup\Log Files", arrSubKeys
    For Each subkey In arrSubKeys
        oReg.GetStringValue HKEY_CURRENT_USER, "Software\Microsoft\Ntbackup\Log Files\" & subkey, "Job Name", JobName
        oReg.GetDWORDValue HKEY_CURRENT_USER, "Software\Microsoft\Ntbackup\Log Files\" & subkey, "Date/Time Used", DateTimeUsed
        If DateTimeUsed > StartTimeStamp Then 'Or UCase(JobName) = UCase(MyJobName) Then
            Set objAppShell = CreateObject("Shell.Application")
            Set objFolder = objAppShell.Namespace(LOCAL_APPLICATION_DATA)
            Set objFolderItem = objFolder.Self
            GetLogFile = objFolderItem.Path & "\Microsoft\Windows NT\NTBackup\data\backup"
            If UCase(subkey) <> "LOG#10" Then 
                GetLogFile = GetLogFile & "0" & Mid(subkey, 5)
            Else
                GetLogFile = GetLogFile & "10"
            End If
            GetLogFile = GetLogFile & ".log"
            Exit For
        End If
    Next
End Function
Public Sub CleanUp(FileName)
    Const DeleteReadOnly = TRUE
	Const wbemFlagReturnImmediately = &h10
	Const wbemFlagForwardOnly = &h20

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objStdOut = WScript.StdOut
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")
    Set objBKF = objFSO.GetFile(FileName)
    BaseName = objBKF.ParentFolder & "\" & objFSO.GetBaseName(objBKF) & "."
    BackupName = Mid(BaseName, 1, Len(BaseName) - Len("yyyyMMdd-HHmmss."))
    SQLSource = "Select * from CIM_DataFile where Path='\\" & Replace(Mid(objBKF.ParentFolder, 4), "\", "\\") & "\\'"
    Set colFiles = objWMIService.ExecQuery(SQLSource, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
    For Each objFile in colFiles
		If Left(objFile.Name, Len(BackupName)) = LCase(BackupName) And Left(objFile.Name, Len(BaseName)) <> LCase(BaseName) Then 
			If Not IsNull(objStdOut) Then objStdOut.WriteLine Now() & vbTab & "Deleting " & objFile.Name
			objFSO.DeleteFile(objFile.Name), DeleteReadOnly
		End If
    Next    
End Sub
Public Sub DoBackup(bks, FileName, JobName, Description, iJob, totalJobs)
	Const ForReading = 1
	COnst UnicodeFormat = -1
    Const OverwriteExisting = TRUE

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")
    Set objStdOut = WScript.StdOut

    dtNow = Now()
    StartTimeStamp = DateDiff("s", BaseCTime(), dtNow)      
    If Not objFSO.FileExists(FileName) Then
        Set objFile = objFSO.CreateTextFile(FileName)                   'Create a dummy file to ease FileName construction...
    End If
    Set objFile = objFSO.GetFile(FileName)
    FileName = objFile.ParentFolder & "\" & objFSO.GetBaseName(objFile) & "."  & FormatTimeStamp(dtNow) & "." & objFSO.GetExtensionName(objFile)
    If objFile.Size = 0 Then objFSO.DeleteFile(objFile.Path)
    
    CommandLine = "NTBACKUP backup """ & bks & """ /v:yes /r:no /rs:no /m normal /j """ & JobName & """ /l:f /f """ & FileName & """ /d """ & Description & """"
    If Not IsNull(objStdOut) Then WScript.StdOut.WriteLine Now() & vbTab & CommandLine
    ExitCode = objShell.Run("cmd /c " & CommandLine, 8, True)

    SourceLog = GetLogFile(StartTimeStamp, JobName)
    If SourceLog <> vbNullString Then
		Set objFile = objFSO.GetFile(FileName)
		TargetLog = objFile.ParentFolder & "\" & objFSO.GetBaseName(objFile) & ".log"
		If Not IsNull(objStdOut) Then WScript.StdOut.WriteLine Now() & vbTab & "Copy """ & SourceLog & """ """ & TargetLog & """"
		objFSO.CopyFile SourceLog, TargetLog, OverwriteExisting

		Success = True
		Set objFile = objFSO.OpenTextFile(TargetLog, ForReading, False, UnicodeFormat)
		Do While Not objFile.AtEndOfStream
		    strLine = objFile.ReadLine
		    If strLine = "The operation did not successfully complete." Then Success = False : Exit Do
		Loop
		objFile.Close				
		If Success And ExitCode = 0 Then CleanUp(FileName)
    End If
	If Not IsNull(objStdOut) Then WScript.StdOut.WriteLine Now() & vbTab & iJob & " of " & totalJobs & " Complete"
End Sub
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript Backup.vbs //X
Dim vArg, aArgs(), iCount
Dim SharedDocuments, BackupFolder, AltBackupFolder
Dim objStdOut
Dim iJob, totalJobs

SharedDocuments = GetEnvironmentVariable("SharedDocuments")
BackupFolder = GetEnvironmentVariable("BackupFolder")
AltBackupFolder = GetEnvironmentVariable("AltBackupFolder")

Set objStdOut = WScript.StdOut
If Not IsNull(objStdOut) Then objStdOut.WriteLine "[Backup.vbs" & vbTab & Now() & "]"
If WScript.Arguments.Count > 0 Then
    If WScript.Arguments.Count <> 4 Then
        If Not IsNull(objStdOut) Then 
            objStdOut.WriteLine "Usage:"
            objStdOut.WriteLine "Backup.vbs [<bks>, <FileName>, <JobName>, <Description>]"
            objStdOut.WriteLine " (either none or all arguments accepted)"
            objStdOut.WriteLine "  bks             Specifies the name of the backup selection file (.bks file) to be used for this backup operation. The @ character"
            objStdOut.WriteLine "                  must precede the name of the backup selection file. A backup selection file contains information on the files and"
            objStdOut.WriteLine "                  folders you have selected for backup. You have to create the file using the graphical user interface (GUI)"
            objStdOut.WriteLine "                  version of Backup. If no bks file is to be used, this argument may be ""systemstate"" or the pathname of the folder"
            objStdOut.WriteLine "                  being backed-up."
            objStdOut.WriteLine "  FileName        Logical disk path and file name."
            objStdOut.WriteLine "  JobName         Specifies the job name to be used in the log file. The job name usually describes the files and folders you are"
            objStdOut.WriteLine "                  backing up in the current backup job as well as the date and time you backed up the files."
            objStdOut.WriteLine "  Description     Specifies a label for backup set. "
        End If
        WScript.Quit
    End If
    ReDim aArgs(wscript.Arguments.Count - 1)
    For iCount = 0 To WScript.Arguments.Count - 1
        aArgs(iCount) = WScript.Arguments(iCount)
    Next
    'DoBackup(bks, FileName, JobName, Description)
    DoBackup aArgs(0), aArgs(1), aArgs(2), aArgs(3), 1, 1
Else
    'Every Day...
    Select Case WeekDayName(WeekDay(Date))
        Case "Sunday"
            totalJobs = 2
        Case "Monday"
            totalJobs = 18
        Case "Tuesday"
            totalJobs = 8
        Case "Wednesday"
            totalJobs = 4
        Case "Thursday"
            totalJobs = 2
        Case "Friday"
            totalJobs = 2
        Case "Saturday"
            totalJobs = 2
    End Select

    iJob = 1
    DoBackup "@" & SharedDocuments & "\SystemState.bks",                BackupFolder & "\SystemState.bkf",                          "SystemState",                       "SystemState Backup",              iJob, totalJobs:    iJob = iJob + 1
    DoBackup SharedDocuments & "\Finance",                              BackupFolder & "\Shared Documents - Finance.bkf",           "Shared Documents - Finance",        "Shared Documents - Finance",      iJob, totalJobs:    iJob = iJob + 1
    Select Case WeekDayName(WeekDay(Date))
        Case "Sunday"
        Case "Monday"
            'DoBackup SharedDocuments & "\My Music",                                   AltBackupFolder & "\Shared Documents - My Music.bkf",      "My Music",                  "My Music",                  iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Jimmy Buffett.bks",   AltBackupFolder & "\My Music - Rock - Jimmy Buffett.bkf",  "My Music - Jimmy Buffett",  "My Music - Jimmy Buffett",  iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Eric Clapton.bks",    AltBackupFolder & "\My Music - Rock - Eric Clapton.bkf",   "My Music - Eric Clapton",   "My Music - Eric Clapton",   iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - ELO.bks",             AltBackupFolder & "\My Music - Rock - ELO.bkf",            "My Music - ELO",            "My Music - ELO",            iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Fleetwood Mac.bks",   AltBackupFolder & "\My Music - Rock - Fleetwood Mac.bkf",  "My Music - Fleetwood Mac",  "My Music - Fleetwood Mac",  iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Genesis.bks",         AltBackupFolder & "\My Music - Rock - Genesis.bkf",        "My Music - Genesis",        "My Music - Genesis",        iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Elton John.bks",      AltBackupFolder & "\My Music - Rock - Elton John.bkf",     "My Music - Elton John",     "My Music - Elton John",     iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Kinks.bks",           AltBackupFolder & "\My Music - Rock - Kinks.bkf",          "My Music - Kinks",          "My Music - Kinks",          iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Alan Parsons.bks",    AltBackupFolder & "\My Music - Rock - Alan Parsons.bkf",   "My Music - Alan Parsons",   "My Music - Parsons, Alan",  iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Pink Floyd.bks",      AltBackupFolder & "\My Music - Rock - Pink Floyd.bkf",     "My Music - Pink Floyd",     "My Music - Pink Floyd",     iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Queen.bks",           AltBackupFolder & "\My Music - Rock - Queen.bkf",          "My Music - Queen",          "My Music - Queen",          iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Rush.bks",            AltBackupFolder & "\My Music - Rock - Rush.bkf",           "My Music - Rush",           "My Music - Rush",           iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Styx.bks",            AltBackupFolder & "\My Music - Rock - Styx.bkf",           "My Music - Styx",           "My Music - Styx",           iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Joe Walsh.bks",       AltBackupFolder & "\My Music - Rock - Joe Walsh.bkf",      "My Music - Joe Walsh",      "My Music - Joe Walsh",      iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock - Yes.bks",             AltBackupFolder & "\My Music - Rock - Yes.bkf",            "My Music - Yes",            "My Music - Yes",            iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music - Rock.bks",                   AltBackupFolder & "\My Music - Rock.bkf",                  "My Music - Rock",           "My Music - Rock",           iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\My Music.bks",                          AltBackupFolder & "\My Music.bkf",                         "My Music",                  "My Music",                  iJob, totalJobs:    iJob = iJob + 1

            'DoBackup SharedDocuments & "\Game Images",                  AltBackupFolder & "\Shared Documents - Game Images.bkf",    "Game Images",    "Game Images",                                        iJob, totalJobs:    iJob = iJob + 1
            'DoBackup SharedDocuments & "\Software Images",              AltBackupFolder & "\Shared Documents - Software Images.bkf","Software Images","Software Images",                                    iJob, totalJobs:    iJob = iJob + 1
        Case "Tuesday"
            DoBackup "@" & SharedDocuments & "\My Profile.bks",         BackupFolder & "\My Profile.bkf",                           "My Profile",                       "My Profile",                       iJob, totalJobs:    iJob = iJob + 1
            DoBackup "C:\Documents and Settings\kclark\My Documents",   BackupFolder & "\GZPR141 My Documents.bkf",                 "GZPR141 My Documents",             "GZPR141 My Documents",             iJob, totalJobs:    iJob = iJob + 1
            DoBackup "@" & SharedDocuments & "\SharedDocuments.bks",    BackupFolder & "\Shared Documents.bkf",                     "Shared Documents",                 "Shared Documents",                 iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Controls",           BackupFolder & "\Downloads - Controls.bkf",					"Downloads - Controls",				"Downloads - Controls",             iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Games",              BackupFolder & "\Downloads - Games.bkf",					"Downloads - Games",				"Downloads - Games",                iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Hardware",           BackupFolder & "\Downloads - Hardware.bkf",					"Downloads - Hardware",				"Downloads - Hardware",             iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Personal Finance",   BackupFolder & "\Downloads - Personal Finance.bkf",         "Downloads - Personal Finance",		"Downloads - Personal Finance",     iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\SunGard",            BackupFolder & "\Downloads - SunGard.bkf",					"Downloads - SunGard",				"Downloads - SunGard",              iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Tools - Development",BackupFolder & "\Downloads - Tools - Development.bkf",      "Downloads - Tools(Development)",	"Downloads - Tools(Development)",   iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Tools - Music",      BackupFolder & "\Downloads - Tools - Music.bkf",			"Downloads - Tools(Music)",			"Downloads - Tools(Music)",         iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Tools - PC",         BackupFolder & "\Downloads - Tools - PC.bkf",				"Downloads - Tools(PC)",			"Downloads - Tools(PC)",            iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Tools - Publishing", BackupFolder & "\Downloads - Tools - Publishing.bkf",       "Downloads - Tools(Publishing)",	"Downloads - Tools(Publishing)",    iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Tools - Web",        BackupFolder & "\Downloads - Tools - Web.bkf",				"Downloads - Tools(Web)",			"Downloads - Tools(Web)",           iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\Downloads\Web Downloads",      BackupFolder & "\Downloads - Web Downloads.bkf",			"Downloads - Web Downloads",		"Downloads - Web Downloads",        iJob, totalJobs:    iJob = iJob + 1
        Case "Wednesday"
            DoBackup "C:\WebShare\wwwroot",                             BackupFolder & "\WebShare - wwwroot.bkf",                   "WebShare - wwwroot",                "WebShare - wwwroot",              iJob, totalJobs:    iJob = iJob + 1
            DoBackup "C:\WebShare\wwwArchive",                          BackupFolder & "\WebShare - wwwArchive.bkf",                "WebShare - wwwArchive",             "WebShare - wwwArchive",           iJob, totalJobs:    iJob = iJob + 1
            DoBackup "C:\Projects",                                     BackupFolder & "\Projects.bkf",                             "Projects",                          "Projects",                        iJob, totalJobs:    iJob = iJob + 1
            DoBackup SharedDocuments & "\My Pictures",                  BackupFolder & "\Shared Documents - My Pictures.bkf",       "Shared Documents - My Pictures",    "My Pictures",                     iJob, totalJobs:    iJob = iJob + 1
        Case "Thursday"
            'VSSArchive Runs Thursdays...
        Case "Friday"
            'FileListDBs Run Fridays...
        Case "Saturday"
    End Select
    'Remote Machine...
    ''DoBackup "\\EUKB6\My Documents",                            BackupFolder & "\EUKB6 My Documents.bkf",                   "EUKB6 My Documents",                   "Full NT Backup of EUKB6 My Documents",           iJob, totalJobs:    iJob = iJob + 1
End If

If Not IsNull(objStdOut) Then objStdOut.Close
WScript.Quit