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
Public Function FormatCTime(CTime)
	If IsDST() Then
		TimeStamp = DateAdd("s", CTime, #12/31/1969 8:00:00 PM#)
	Else
		TimeStamp = DateAdd("s", CTime, #12/31/1969 7:00:00 PM#)
	End If
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
strOutput = ""
oReg.EnumKey HKEY_CURRENT_USER, "Software\Microsoft\Ntbackup\Log Files", arrSubKeys
For Each subkey In arrSubKeys
    oReg.GetStringValue HKEY_CURRENT_USER, "Software\Microsoft\Ntbackup\Log Files\" & subkey, "Job Name", JobName
    oReg.GetDWORDValue HKEY_CURRENT_USER, "Software\Microsoft\Ntbackup\Log Files\" & subkey, "Date/Time Used", DateTimeUsed
    
	strOutput = strOutput & subkey & ": " & FormatCTime(DateTimeUsed) & " - " & JobName & vbCrLf
Next
MsgBox(strOutput)
