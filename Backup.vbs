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
        If DateTimeUsed > StartTimeStamp Or UCase(JobName) = UCase(MyJobName) Then
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
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")
    Set objBKF = objFSO.GetFile(FileName)
    BaseName = objBKF.ParentFolder & "\" & objFSO.GetBaseName(objBKF) & "."
    BackupName = Mid(BaseName, 1, Len(BaseName) - Len("yyyyMMdd-HHmmss."))
    SQLSource = "Select * from CIM_DataFile where Path='\\" & Replace(Mid(objBKF.ParentFolder, 4), "\", "\\") & "\\'"
    Set colFiles = objWMIService.ExecQuery(SQLSource, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
    For Each objFile in colFiles
		If Left(objFile.Name, Len(BackupName)) = LCase(BackupName) And Left(objFile.Name, Len(BaseName)) <> LCase(BaseName) Then 
			'Wscript.Echo objFile.Name
			objFSO.DeleteFile(objFile.Name), DeleteReadOnly
		End If
    Next    
End Sub
Public Sub DoBackup(bks, FileName, JobName, Description)
	Const ForReading = 1
	COnst UnicodeFormat = -1
    Const OverwriteExisting = TRUE

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")
    dtNow = Now()
    StartTimeStamp = DateDiff("s", #12/31/1969 7:00:00 PM#, dtNow)      'CTime := # Seconds since 01/01/1970 GMT (adjusted for EST)...
    If Not objFSO.FileExists(FileName) Then
        Set objFile = objFSO.CreateTextFile(FileName)                   'Create a dummy file to ease FileName construction...
    End If
    Set objFile = objFSO.GetFile(FileName)
    FileName = objFile.ParentFolder & "\" & objFSO.GetBaseName(objFile) & "."  & FormatTimeStamp(dtNow) & "." & objFSO.GetExtensionName(objFile)
    If objFile.Size = 0 Then objFSO.DeleteFile(objFile.Path)
    
    CommandLine = "NTBACKUP backup """ & bks & """ /v:yes /r:no /rs:no /m normal /j """ & JobName & """ /l:f /f """ & FileName & """ /d """ & Description & """"
    'MsgBox("CommandLine: " & CommandLine)
    ExitCode = objShell.Run("cmd /c " & CommandLine, 8, True)

    SourceLog = GetLogFile(StartTimeStamp, JobName)
    If SourceLog <> vbNullString Then
		Set objFile = objFSO.GetFile(FileName)
		TargetLog = objFile.ParentFolder & "\" & objFSO.GetBaseName(objFile) & ".log"
		'MsgBox "Copy """ & SourceLog & """ """ & TargetLog & """"
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
End Sub
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript Backup.vbs //X

SharedDocuments = GetEnvironmentVariable("SharedDocuments")
BackupFolder = GetEnvironmentVariable("BackupFolder")
AltBackupFolder = "E" & Mid(BackupFolder, 2)

'Every Day...
DoBackup "@" & SharedDocuments & "\SystemState.bks",                BackupFolder & "\SystemState.bkf",                          "SystemState",                       "SystemState Backup"                                   '  5 minutes
DoBackup SharedDocuments & "\Finance",                              BackupFolder & "\Shared Documents - Finance.bkf",           "Shared Documents - Finance",        "Full NT Backup of Shared Documents - Finance"         '  5 minutes
Select Case WeekDayName(WeekDay(Date))
    Case "Sunday"
    Case "Monday"
        DoBackup SharedDocuments & "\My Music",                     AltBackupFolder & "\Shared Documents - My Music.bkf",       "Shared Documents - My Music",       "Full NT Backup of Shared Documents - My Music"        '210 minutes
        DoBackup SharedDocuments & "\Game Images",                  AltBackupFolder & "\Shared Documents - Game Images.bkf",    "Shared Documents - Game Images",    "Full NT Backup of Shared Documents - Game Images"     '360 minutes+
        DoBackup SharedDocuments & "\Software Images",              AltBackupFolder & "\Shared Documents - Software Images.bkf","Shared Documents - Software Images","Full NT Backup of Shared Documents - Software Images" ' 15 minutes
    Case "Tuesday"
        DoBackup "C:\Documents and Settings\kclark\My Documents",   BackupFolder & "\GZPR141 My Documents.bkf",                 "GZPR141 My Documents",              "Full NT Backup of GZPR141 My Documents"               ' 30 minutes
        DoBackup "C:\Projects",                                     BackupFolder & "\Projects.bkf",                             "Projects",                          "Full NT Backup of Projects"                           ' 40 minutes
        DoBackup "@" & SharedDocuments & "\SharedDocuments.bks",    BackupFolder & "\Shared Documents.bkf",                     "Shared Documents",                  "Full NT Backup of Shared Documents"                   '210 minutes
        DoBackup SharedDocuments & "\Downloads",                    BackupFolder & "\Shared Documents - Downloads.bkf",         "Shared Documents - Downloads",      "Full NT Backup of Shared Documents - Downloads"       '160 minutes
        DoBackup SharedDocuments & "\My Pictures",                  BackupFolder & "\Shared Documents - My Pictures.bkf",       "Shared Documents - My Pictures",    "Full NT Backup of Shared Documents - My Pictures"     '  5 minutes
        DoBackup "@" & SharedDocuments & "\My Profile.bks",         BackupFolder & "\My Profile.bkf",                           "My Profile",                        "Full NT Backup of My Profile"                         ' 35 minutes
    Case "Wednesday"
        DoBackup "C:\WebShare\wwwroot",                             BackupFolder & "\WebShare - wwwroot.bkf",                   "WebShare - wwwroot",                "Full NT Backup of WebShare - wwwroot"                 ' 15 minutes
        DoBackup "C:\WebShare\wwwArchive",                          BackupFolder & "\WebShare - wwwArchive.bkf",                "WebShare - wwwArchive",             "Full NT Backup of WebShare - wwwArchive"              '300 minutes
    Case "Thursday"
        'VSSArchive Runs Thursdays...
    Case "Friday"
        'FileListDBs Run Fridays...
    Case "Saturday"
End Select


'Remote Machine...
''DoBackup "\\EUKB6\My Documents",                            BackupFolder & "\EUKB6 My Documents.bkf",                   "EUKB6 My Documents",                   "Full NT Backup of EUKB6 My Documents"                         ' minutes
