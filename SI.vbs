'SI.vbs
'	Visual Basic Script Used to Replicate Control Panel's Programs and Features Display (Output in Excel)...
'   Copyright © 2015, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   06/20/16    Ken Clark     Additional Scheduled Task refinements;
'   09/26/15    Ken Clark     Enabled/tested running as a Scheduled Task;
'   09/21/14	  Ken Clark		  Created;
'=================================================================================================================================
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'	cscript//X SI.vbs
'or... (Note that arguments are enclosed in double-quotes due to embedded spaces, and arguments are separated by spaces - not commas)
'	cscript//X SI.vbs "arg1" "arg2" "arg3" "arg4"
'=================================================================================================================================
Private Function Fill(num)
    If(Len(num)=1) Then Fill = "0" & num Else Fill = num
End Function
Public Function FormatTimeStamp(DateTime)
    yyyy = Year(DateTime)
    MM = Fill(Month(DateTime))    
    dd = Fill(Day(DateTime))
    hh = Fill(Hour(DateTime))
    m = Fill(Minute(DateTime))
    ss = Fill(Second(DateTime))
    FormatTimeStamp = yyyy & MM & dd & "." & hh & m & ss
End Function
Private Sub LogMessage(LogFile,Message)
    Const ForAppending = 8
    Const UnicodeFormat = -1
    Const MB = 1048576
    Dim objStdOut : Set objStdOut = WScript.StdOut : If Not IsNull(objStdOut) Then objStdOut.WriteLine Message
    Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile
    
    If objFSO.FileExists(LogFile) Then
        Set objFile = objFSO.GetFile(LogFile)
        If objFile.Size > 10*MB Then
            Dim dtModified, NewFileName
            dtModified = objFile.DateLastModified
            NewFileName = Replace(LogFile, ".log", "." & FormatTimeStamp(dtModified) & ".log")
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
    
    Set objFile = objFSO.OpenTextFile(LogFile, ForAppending, True)
    objFile.WriteLine(Message)
    objFile.Close
    
    Set objFile = Nothing
    Set objFSO = Nothing
    Set objStdOut = Nothing
End Sub

On Error Resume Next
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Dim vArg, aArgs(), iCount, BackupServer
BackupServer = "ALPHA"

Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
    strComputer = objItem.Path_.Server: Exit For
Next

If WScript.Arguments.Count > 0 Then
	If WScript.Arguments.Count = 1 Then
		BackupServer = UCase(WScript.Arguments(0))
  End If
End If

Dim LogFile : LogFile = "\\" & BackupServer & "\Backups\" & strComputer & "\SI." & strComputer & "." & FormatTimeStamp(Now()) & ".log"
LogMessage LogFile, "[SI.vbs" & vbTab & Now() & "]"

Set Regions=CreateObject("Scripting.Dictionary")
Regions.Add HKEY_LOCAL_MACHINE, HKEY_LOCAL_MACHINE
Regions.Add HKEY_CURRENT_USER, HKEY_CURRENT_USER

Set Keys=CreateObject("Scripting.Dictionary")
Keys.Add "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\",1
Keys.Add "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\",2

LogMessage LogFile, "Creating Excel Object..."
Dim objExcel : Set objExcel = CreateObject("Excel.Application")
If objExcel Is Nothing Then
    LogMessage LogFile, "Unable to invoke Excel, continuing with tab-delimited data..."

    LogMessage LogFile, "Examining Computer System..."
    x = 2
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"DefaultUserName",strUserName

    Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
    strKeyPath = "System\CurrentControlSet\Services\lanmanserver\parameters"
    objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "srvcomment", strDescription

    LogMessage LogFile, "Scanning " & strComputer & " Registry for Installed Software..."
    LogMessage LogFile, "Application Name" & vbTab & _
                        "Version" & vbTab & _
                        "Publisher" & vbTab & _
                        "Install Date" & vbTab & _
                        "Estimated Size" & vbTab & _
                        "Install Location" & vbTab & _
                        "Uninstall String" & vbTab & _
                        "Source"
    Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
    For Each Region in Regions
        For Each Key in Keys
            objReg.EnumKey Region, Key, arrSubkeys
            For Each subKey In arrSubkeys
                DisplayName = ""
                DisplayVersion = ""
                EstimatedSize = ""
                InstallDate = ""
                InstallLocation = ""
                Publisher = ""
                Source = ""
                doThisEntry = True
                ReleaseType = ""

                intStatus = objReg.GetStringValue(Region, Key & subKey, "ReleaseType", strValue)
                if intStatus = 0 And strValue = "Security Update" Then doThisEntry = False
                intStatus = objReg.GetDWORDValue(Region, Key & subKey, "SystemComponent", intValue)
                if intStatus = 0 And intValue = 1 Then doThisEntry = False

                intStatus = objReg.GetStringValue(Region, Key & subKey, "DisplayName", DisplayName)
                If intStatus <> 0 Then intStatus = objReg.GetStringValue(Region, Key & subKey, "QuietDisplayName", DisplayName)
                If intStatus <> 0 Then DisplayName = ""
                If Len(DisplayName) = 0 Then doThisEntry = False

                if CBool(doThisEntry) Then
                    'LogMessage LogFile, "DisplayName: """ & DisplayName & """ (" & Len(DisplayName) & "); doThisEntry: " & doThisEntry

                    Set iProduct = Nothing
                    Set installer = CreateObject("WindowsInstaller.Installer")
                    For Each product In installer.Products
                        If DisplayName = installer.ProductInfo(product, "ProductName") Then Set iProduct = product : Exit For
                    Next
                    if Region = HKEY_LOCAL_MACHINE then Source = "HKEY_LOCAL_MACHINE\" & Key else Source = "HKEY_CURRENT_USER\" & Key
                    objReg.GetStringValue Region, Key & subKey, "Publisher", Publisher
                    If Publisher = "" And Not iProduct Is Nothing Then Publisher = installer.ProductInfo(iProduct, "Publisher") 
        
                    objReg.GetStringValue Region, Key & subKey, "InstallDate", InstallDate
                    If InstallDate = "" And Not iProduct Is Nothing Then InstallDate = installer.ProductInfo(iProduct, "InstallDate")
                    If InstallDate <> "" Then InstallDate = Mid(InstallDate,5,2) & "/" & Right(InstallDate,2) & "/" & Left(InstallDate,4)
        
                    objReg.GetStringValue Region, Key & subKey, "DisplayVersion", DisplayVersion
                    objReg.GetDWORDValue Region, Key & subKey, "VersionMajor", intValue1
                    objReg.GetDWORDValue Region, Key & subKey, "VersionMinor", intValue2
                    If DisplayVersion = "" And Not iProduct Is Nothing Then DisplayVersion = installer.ProductInfo(iProduct, "DisplayVersion") 
                    If DisplayVersion = "" And intValue1 <> 0 Then DisplayVersion = intValue1 & "." & intValue2
        
                    intStatus = objReg.GetDWORDValue(Region, Key & subKey, "EstimatedSize", intValue)
                    If intStatus = 0 Then EstimatedSize = Round(intValue/1024, 3) & " MB"
        
                    intStatus = objReg.GetStringValue(Region, Key & subKey, "InstallLocation", InstallLocation)
                    intStatus = objReg.GetStringValue(Region, Key & subKey, "UninstallString", UninstallString)

                    LogMessage LogFile, DisplayName & vbTab & _
                                        DisplayVersion & vbTab & _
                                        Publisher & vbTab & _
                                        InstallDate & vbTab & _
                                        EstimatedSize & vbTab & _
                                        InstallLocation & vbTab & _
                                        UninstallString & vbTab & _
                                        Source
                    x = x + 1
                End If
            Next
        Next
    Next
    LogMessage LogFile, "Process Complete - " & x & " entries"
Else
    objExcel.Visible = False
    objExcel.Workbooks.Add

    Set objRange = objExcel.Worksheets	'Range("A1","G5")
    objRange.Font.Size = 14
    'Format cells
    objExcel.Range("A1:H1").Select
    objExcel.Selection.Font.bold = True
    objExcel.Selection.Interior.ColorIndex = 4
    objExcel.Selection.Interior.Pattern = 1
    objExcel.Selection.Font.ColorIndex = 1
    objExcel.ActiveWindow.SplitColumn = 1
    objExcel.ActiveWindow.SplitRow = 1
    objExcel.ActiveWindow.FreezePanes = True

    'iComp=1:    objExcel.Cells(1, iComp).Value = "Computer Name":		objExcel.Columns(iComp).ColumnWidth = 15 
    'iDesc=2:    objExcel.Cells(1, iDesc).Value = "Description":		objExcel.Columns(iDesc).ColumnWidth = 15 
    'iUser=3:    objExcel.Cells(1, iUser).Value = "UserName":			objExcel.Columns(iUser).ColumnWidth = 15 
    iName=1:    objExcel.Cells(1, iName).Value = "Application Name":	objExcel.Columns(iName).ColumnWidth = 40 
    iVer=2:     objExcel.Cells(1, iVer).Value = "Version":				objExcel.Columns(iVer).ColumnWidth = 15
    iPub=3:     objExcel.Cells(1, iPub).Value = "Publisher":			objExcel.Columns(iPub).ColumnWidth = 15 
    iDate=4:    objExcel.Cells(1, iDate).Value = "Install Date":		objExcel.Columns(iDate).ColumnWidth = 15 
    iSize=5:    objExcel.Cells(1, iSize).Value = "Estimated Size":		objExcel.Columns(iSize).ColumnWidth = 15 
    iLoc=6:     objExcel.Cells(1, iLoc).Value = "Install Location":	    objExcel.Columns(iLoc).ColumnWidth = 40 
    iUnins=7:   objExcel.Cells(1, iUnins).Value = "Uninstall String":	objExcel.Columns(iUnins).ColumnWidth = 40 
    iSource=8:  objExcel.Cells(1, iSource).Value = "Source":	        objExcel.Columns(iSource).ColumnWidth = 40 

    LogMessage LogFile, "Examining Computer System..."
    objExcel.ActiveSheet.Name = strComputer
    'xlsFile = Replace(WScript.ScriptFullName, ".vbs", "." & strComputer & "." & FormatTimeStamp(Now()) & ".xlsx")
    xlsFile = "\\" & BackupServer & "\Backups\" & strComputer & "\SI." & strComputer & "." & FormatTimeStamp(Now()) & ".xlsx"
    Set fso = CreateObject("Scripting.FileSystemObject")
    if fso.FileExists(xlsFile) then fso.DeleteFile(xlsFile)

    x = 2
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"DefaultUserName",strUserName

    Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
    strKeyPath = "System\CurrentControlSet\Services\lanmanserver\parameters"
    objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "srvcomment", strDescription

    LogMessage LogFile, "Scanning " & strComputer & " Registry for Installed Software..."
    Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
    For Each Region in Regions
        For Each Key in Keys
            objReg.EnumKey Region, Key, arrSubkeys
            For Each subKey In arrSubkeys
                DisplayName = ""
                DisplayVersion = ""
                EstimatedSize = ""
                InstallDate = ""
                InstallLocation = ""
                Publisher = ""
                Source = ""
                doThisEntry = True
                intStatus = objReg.GetStringValue(Region, Key & subKey, "ReleaseType", strValue)
                if intStatus = 0 And strValue = "Security Update" Then doThisEntry = False
                intStatus = objReg.GetDWORDValue(Region, Key & subKey, "SystemComponent", intValue)
                if intStatus = 0 And intValue = 1 Then doThisEntry = False

                intStatus = objReg.GetStringValue(Region, Key & subKey, "DisplayName", DisplayName)
                If intStatus <> 0 Then intStatus = objReg.GetStringValue(Region, Key & subKey, "QuietDisplayName", DisplayName)
                'If Len(DisplayName) = 0 <> "" Then objExcel.Cells(x, iName).Value = DisplayName Else objExcel.Cells(x, iName).Value = subKey
                If Len(DisplayName) <> 0 Then objExcel.Cells(x, iName).Value = DisplayName Else doThisEntry = False

                if doThisEntry Then
                    Set iProduct = Nothing
                    Set installer = CreateObject("WindowsInstaller.Installer")
                    For Each product In installer.Products
                        If DisplayName = installer.ProductInfo(product, "ProductName") Then Set iProduct = product : Exit For
                    Next
                    if Region = HKEY_LOCAL_MACHINE then Source = "HKEY_LOCAL_MACHINE\" & Key else Source = "HKEY_CURRENT_USER\" & Key
                    objExcel.Cells(x, iSource).Value = Source
                    'objExcel.Cells(x, iComp).Value = strComputer
                    'objExcel.Cells(x, iDesc).Value = strDescription
                    'objExcel.Cells(x, iUser).Value = strUserName
        
                    objReg.GetStringValue Region, Key & subKey, "Publisher", Publisher
                    If Publisher = "" And iProduct Is Not Nothing Then Publisher = installer.ProductInfo(iProduct, "Publisher") 
                    If Publisher <> "" Then objExcel.Cells(x, iPub).Value = Publisher
        
                    objReg.GetStringValue Region, Key & subKey, "InstallDate", InstallDate
                    If InstallDate = "" And iProduct Is Not Nothing Then InstallDate = installer.ProductInfo(iProduct, "InstallDate")
                    If InstallDate <> "" Then 
                        InstallDate = Mid(InstallDate,5,2) & "/" & Right(InstallDate,2) & "/" & Left(InstallDate,4)
                        objExcel.Cells(x, iDate).Value = "'" & InstallDate
                    End If
        
                    objReg.GetStringValue Region, Key & subKey, "DisplayVersion", DisplayVersion
                    objReg.GetDWORDValue Region, Key & subKey, "VersionMajor", intValue1
                    objReg.GetDWORDValue Region, Key & subKey, "VersionMinor", intValue2
                    If DisplayVersion = "" And iProduct Is Not Nothing Then DisplayVersion = installer.ProductInfo(iProduct, "DisplayVersion") 
                    If DisplayVersion = "" And intValue1 <> 0 Then DisplayVersion = intValue1 & "." & intValue2
                    If DisplayVersion <> "" Then objExcel.Cells(x, iVer).Value = "'" & DisplayVersion
        
                    intStatus = objReg.GetDWORDValue(Region, Key & subKey, "EstimatedSize", intValue)
                    If intStatus = 0 Then objExcel.Cells(x, iSize).Value = Round(intValue/1024, 3) & " MB"
        
                    intStatus = objReg.GetStringValue(Region, Key & subKey, "InstallLocation", InstallLocation)
                    If intStatus = 0 Then 
                        objExcel.Cells(x, iLoc).Value = InstallLocation
                    Else
                        intStatus = objReg.GetStringValue(Region, Key & subKey, "EXE", InstallLocation)
                        If intStatus = 0 Then objExcel.Cells(x, iLoc).Value = InstallLocation
                    End If
        
                    intStatus = objReg.GetStringValue(Region, Key & subKey, "UninstallString", UninstallString)
                    If intStatus = 0 Then objExcel.Cells(x, iUnins).Value = UninstallString

                    x = x + 1
                End If
            Next
        Next
    Next
    LogMessage LogFile, "Sorting Excel Data..."
    objExcel.ActiveSheet.Cells.Select
    objExcel.ActiveSheet.Cells.EntireColumn.AutoFit
    With objExcel.ActiveSheet.Sort 
        .Clear
        .SortFields.Add objExcel.Range("A2:A" & CInt(x-1))
        .SetRange objExcel.Range("A2:H" & CInt(x-1))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With 
    objExcel.Range("A1").Select
    LogMessage LogFile, "Saving Excel Workbook..."
    objExcel.ActiveWorkbook.SaveAs xlsFile
    objExcel.ActiveWorkbook.Close
    objExcel.Quit
    LogMessage LogFile, "Process Complete - " & xlsFile
End If

If Not IsNull(WScript.StdOut) Then WScript.StdOut.Close
WScript.Quit