Private Sub LogMessage(Message)
	Const ForAppending = 8
	Const UnicodeFormat = -1
	Const MB = 1048576
	Dim objStdOut, objFSO, objFile, LogFile
    Set objStdOut = WScript.StdOut
	If Not IsNull(objStdOut) Then objStdOut.WriteLine Message
	
	LogFile = BackupFolder & "\WSArchive.log"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(LogFile) Then
		Set objFile = objFSO.GetFile(LogFile)
		If objFile.Size > 10*MB Then
            Dim dtModified, NewFileName
            dtModified = objFile.DateLastModified
            NewFileName = BackupFolder & "\WSArchive." & FormatTimeStamp(dtModified) & ".log"
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
    
	Set objFile = objFSO.OpenTextFile(BackupFolder & "\WSArchive.log", ForAppending, True)
	objFile.WriteLine(Message)
	objFile.Close
	
	Set objFile = Nothing
	Set objFSO = Nothing
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
Public Sub CleanUp(BackupFolder)
    Const DeleteReadOnly = TRUE
    Const wbemFlagReturnImmediately = &h10
    Const wbemFlagForwardOnly = &h20

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")
    SQLSource = "Select * from CIM_DataFile where Path='" & Replace(Mid(BackupFolder, 3), "\", "\\") & "\\' And Extension='log' And CreationDate <= '" & DateAdd("d",-27,Now()) & "'"
    Set colFiles = objWMIService.ExecQuery(SQLSource, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
    For Each objFile in colFiles
        LogMessage Now() & vbTab & "Deleting: " & objFile.Name
        objFSO.DeleteFile(objFile.Name), DeleteReadOnly
    Next    
End Sub
Public Sub DoArchive
    LogMessage "[WSArchive Version 1.0]"
    dtNow = Now()
    TimeStamp = FormatTimeStamp(dtNow)

    CommandLine = "XCOPY D:\Workspaces\WayneDevelopment\*.* V:\Workspaces\WayneDevelopment /S /D /C /X /K /R /L /Y"
    LogMessage Now() & vbTab & CommandLine

    Set objShell = CreateObject("WScript.Shell")
    ExitCode = objShell.Run(CommandLine, 8, True)
    If ExitCode = 0 Then 
      LogMessage Now() & vbTab & "Archive Complete."
    Else
      LogMessage Now() & vbTab & "Archive Failed."
    End If
End Sub
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript WSArchive.vbs //X

BackupFolder = "D:\Workspaces\WayneDevelopment"
DoArchive
If Not IsNull(WScript.StdOut) Then WScript.StdOut.Close
WScript.Quit