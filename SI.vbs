'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X SI.vbs
'or... (Note that arguments are enclosed in double-quotes due to embedded spaces, and arguments are separated by spaces - not commas)
'	cscript//X SI.vbs "arg1" "arg2" "arg3" "arg4"
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const ForReading = 1

strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

Set objRange = objExcel.Worksheets	'Range("A1","G5")
objRange.Font.Size = 14
'Format cells
objExcel.Range("A1:S1").Select
objExcel.Selection.Font.bold = True
objExcel.Selection.Interior.ColorIndex = 4
objExcel.Selection.Interior.Pattern = 1
objExcel.Selection.Font.ColorIndex = 1

objExcel.Cells(1, 1).Value = "Computer Name":		objExcel.Columns(1).ColumnWidth = 15 
objExcel.Cells(1, 2).Value = "Description":			objExcel.Columns(2).ColumnWidth = 15 
objExcel.Cells(1, 3).Value = "UserName":			objExcel.Columns(3).ColumnWidth = 15 
objExcel.Cells(1, 4).Value = "Publisher":			objExcel.Columns(4).ColumnWidth = 15 
objExcel.Cells(1, 5).Value = "Application Name":	objExcel.Columns(5).ColumnWidth = 40 
objExcel.Cells(1, 6).Value = "Version":				objExcel.Columns(6).ColumnWidth = 15 
objExcel.Cells(1, 7).Value = "Install Date":		objExcel.Columns(7).ColumnWidth = 15 
objExcel.Cells(1, 8).Value = "Estimated Size":		objExcel.Columns(8).ColumnWidth = 15 
objExcel.Cells(1, 9).Value = "Install Location":	objExcel.Columns(9).ColumnWidth = 40 
objExcel.Cells(1, 10).Value = "Uninstall String":	objExcel.Columns(10).ColumnWidth = 40 

Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
	strComputer = objItem.Path_.Server: Exit For
Next

x = 2
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"DefaultUserName",strUserName

Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "System\CurrentControlSet\Services\lanmanserver\parameters"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "srvcomment", strDescription

Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey, arrSubkeys
For Each strSubkey In arrSubkeys
	doThisEntry = True
	intStatus = objReg.GetStringValue(HKLM, strKey & strSubkey, "ReleaseType", strValue)
	if intStatus = 0 And strValue = "Security Update" Then doThisEntry = False
	intStatus = objReg.GetDWORDValue(HKLM, strKey & strSubkey, "SystemComponent", intValue)
	if intStatus = 0 And intValue = 1 Then doThisEntry = False

	if doThisEntry Then
		objExcel.Cells(x, 1).Value = strComputer
		objExcel.Cells(x, 2).Value = strDescription
		objExcel.Cells(x, 3).Value = strUserName
		
		intStatus = objReg.GetStringValue(HKLM, strKey & strSubkey, "Publisher", strValue)
		If intStatus = 0 Then objExcel.Cells(x, 4).Value = strValue
		
		intStatus = objReg.GetStringValue(HKLM, strKey & strSubkey, "DisplayName", strApplication)
		If intStatus <> 0 Then intStatus = objReg.GetStringValue(HKLM, strKey & strSubkey, "QuietDisplayName", strApplication)
		If strApplication <> "" Then 
			objExcel.Cells(x, 5).Value = strApplication
		Else
			objExcel.Cells(x, 5).Value = strSubKey
		End If
		
		objReg.GetStringValue HKLM, strKey & strSubkey, "InstallDate", strValue
		If strValue <> "" Then objExcel.Cells(x, 6).Value = strValue
		
		objReg.GetStringValue HKLM, strKey & strSubkey, "DisplayVersion", strValue
		If strValue <> "" Then 
			objExcel.Cells(x, 7).Value = strValue
		Else
			objReg.GetDWORDValue HKLM, strKey & strSubkey, "VersionMajor", intValue1
			objReg.GetDWORDValue HKLM, strKey & strSubkey, "VersionMinor", intValue2
			If intValue1 <> 0 Then objExcel.Cells(x, 7).Value = "'" & intValue1 & "." & intValue2
		End If
		
		intStatus = objReg.GetDWORDValue(HKLM, strKey & strSubkey, "EstimatedSize", intValue)
		If intStatus = 0 Then objExcel.Cells(x, 8).Value = Round(intValue/1024, 3) & " MB"
		
		intStatus = objReg.GetStringValue(HKLM, strKey & strSubkey, "InstallLocation", strValue)
		If intStatus = 0 Then 
			objExcel.Cells(x, 9).Value = strValue
		Else
			intStatus = objReg.GetStringValue(HKLM, strKey & strSubkey, "EXE", strValue)
			If intStatus = 0 Then objExcel.Cells(x, 9).Value = strValue
		End If
		
		intStatus = objReg.GetStringValue(HKLM, strKey & strSubkey, "UninstallString", strValue)
		If intStatus = 0 Then objExcel.Cells(x, 10).Value = strValue

		x = x + 1
	End If
Next

objFile.Close
'Save file
objExcel.ActiveWorkbook.SaveAs Replace(WScript.ScriptFullName, ".vbs", ".xls")
objExcel.ActiveWorkbook.Close
objExcel.Quit

