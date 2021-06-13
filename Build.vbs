'Build.vbs
'    Visual Basic Script Used to Build .NET Components for the FiRRe Application...
'   Copyright © 2006-2014, SunGard
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:      Description:
'   10/14/14    Ken Clark       Updated to support VS2013;
'   03/21/14    Ken Clark       Updated to better support VS2005;
'   04/10/13    Ken Clark       Added FiRReData.NET & DBUtility.NET;
'   01/09/13    Ken Clark       Handled space-in-the-project-name issues;
'   07/10/12    Ken Clark       Upgraded to handle either Visual Studio 2003 or 2005 dynamically;
'   05/27/12    Ken Clark       Corrected file naming issues when renaming logs bigger than 10MB;
'   05/25/11    Ken Clark       Created;
'=================================================================================================================================
'Notes:
'Recommended Command-Line:   cscript "Build.vbs" "All"
'Script can be debugged by opening a CMD window and executing the following command (note that the two slashes are not a typo)...
'   cscript//X "Build.vbs" "All"
'=================================================================================================================================
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const UnicodeFormat = -1
Const MB = 1048576
Dim WshShell, objFSO, startFolder, vsFileName, vs2003FileName, vs2005FileName, vs2013FileName, Projects, BuildOption, Version, VersionTag, ProjectTag
Dim iSucceeded, iFailed, iSkipped, FailedList
Dim defaultVS2003Option : defaultVS2003Option = "Debug"
Dim defaultVS2005Option : defaultVS2005Option = "Debug|x86"
Dim defaultVS2013Option : defaultVS2013Option = "Debug|x86"
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
startFolder = WshShell.CurrentDirectory

Private Sub LogMessage(Message)
    Dim objStdOut, objFile, LogFile, BaseName
    Set objStdOut = WScript.StdOut
    On Error Resume Next
    objStdOut.WriteLine Message
    On Error GoTo 0
   
    BaseName = startFolder & "\" & Replace(WScript.ScriptName, ".vbs", "") 
    If Version <> "" Then BaseName = BaseName & " " & Version
    
    LogFile = BaseName & ".log"
    If objFSO.FileExists(LogFile) Then
        Set objFile = objFSO.GetFile(LogFile)
        If objFile.Size > 10*MB Then
            Dim dtModified, NewFileName
            dtModified = objFile.DateLastModified
            NewFileName = BaseName & "." & FormatTimeStamp(dtModified) & ".log"
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
Private Function ExecuteWithoutOutput(Command)
    Dim oExec
    'LogMessage("Executing: " & Command)
    Set oExec = WshShell.Exec(Command)
    Do
        WScript.Sleep 10
    Loop Until oExec.Status <> 0
    ExecuteWithoutOutput = oExec.ExitCode
    Set oExec = Nothing
End Function
Private Function Execute(Command)
    Dim oExec, oStdOut, sOutput
    'LogMessage("Executing: " & SS & Command)
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
Private Function GetProjectName(ProjectFileNameBase)
    Dim iPos
    GetProjectName = ProjectFileNameBase
    If Right(LCase(ProjectFileNameBase), 4) = ".sln" Then GetProjectName = Left(ProjectFileNameBase, Len(ProjectFileNameBase)-4) : Exit Function
    iPos = InStrRev(ProjectFileNameBase, "\")
    If iPos <> 0 Then GetProjectName = Mid(ProjectFileNameBase, iPos+1)
End Function
Private Sub RegisterDotNet(FileName, Interop)
    Dim objFile, CommandLine, ExitCode

    LogMessage("   Force Registering...")
    If Not objFSO.FileExists(FileName) Then
        LogMessage("      " & FileName & " not found!")
        Exit Sub
    End If
    Set objFile = objFSO.GetFile(FileName)    
    If LCase(Right(objFile.Name, Len(".net.dll"))) = ".net.dll" Then 
        CommandLine = "C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\RegAsm.exe /unregister """ & objFile.Path & """ /silent"
        ExitCode = ExecuteWithoutOutput(CommandLine)
        If ExitCode <> 0 Then LogMessage("      Unregister failed. (" & ExitCode & ")")
        
        CommandLine = "C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\RegAsm.exe """ & objFile.Path & """ /silent"
        ExitCode = ExecuteWithoutOutput(CommandLine)
        If ExitCode <> 0 Then LogMessage("      Register failed. (" & ExitCode & ")")

        If Interop Then
            LogMessage("   Registering for COM interop...")
            CommandLine = "C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\RegAsm.exe  """ & objFile.Path & """ /tlb /silent"
            ExitCode = ExecuteWithoutOutput(CommandLine)
            If ExitCode <> 0 Then LogMessage("      Type Library registration failed. (" & ExitCode & ")")
        End If
    End If
End Sub
Private Sub BuildProject(MainProject, Project, BuildProjects, BuildOption, Version, Interop, iSucceeded, iFailed, iSkipped)
    Dim i, CommandLine, ProjectName, objStream, strLine, iPos, RealSource, VersionTag, ProjectFileNameBase, vsVersion

    If UCase(MainProject) <> UCase(BuildProjects) And UCase(BuildProjects) <> "ALL" Then Exit Sub
    i = 0
    VersionTag = ""
    If Version <> "" Then VersionTag = " Version " & Version
    If Right(LCase(Project), 4) = ".sln" Then
        RealSource = "V:\" & MainProject & VersionTag & "\"
    Else
        RealSource = "V:\" & MainProject & VersionTag & "\" & Project & "\"
    End If

    ProjectFileNameBase = RealSource & GetProjectName(Project)
    If Version <> "" And objFSO.FileExists(RealSource & GetProjectName(Project) & " v" & Version) Then
        ProjectFileNameBase = ProjectFileNameBase & " v" & Version
    End If

    If Not objFSO.FileExists(ProjectFileNameBase & ".sln") Then
        LogMessage("   ERROR: " & ProjectFileNameBase & ".sln does not exist!")
        Exit Sub
    ElseIf Not objFSO.FileExists(ProjectFileNameBase & ".vbproj") Then
        LogMessage("   ERROR: " & ProjectFileNameBase & ".vbproj does not exist!")
        Exit Sub
    End If

    WshShell.CurrentDirectory = objFSO.GetParentFolderName(ProjectFileNameBase & ".sln")
    ProjectName = GetProjectName(ProjectFileNameBase)
    LogMessage("   Building " & ProjectName)
    LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
    Set objStream = objFSO.OpenTextFile(ProjectFileNameBase & ".sln", ForReading, False)
    strLine = objStream.ReadLine
    objStream.Close
    If strLine = "Microsoft Visual Studio Solution File, Format Version 8.00" Then
        vsVersion = "2003"
        If vs2003FileName = "" Then 
            LogMessage("   ERROR: Unable to determine Visual Studio .NET 2003 DEVENV.EXE location") 
            LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
            Exit Sub
        End If
        vsFileName = vs2003FileName
        If BuildOption = "" Then BuildOption = defaultVS2003Option
    ElseIf strLine = "Microsoft Visual Studio Solution File, Format Version 9.00" Then
        vsVersion = "2005"
        If vs2005FileName = "" Then 
            LogMessage("   ERROR: Unable to determine Visual Studio 2005 DEVENV.EXE location") 
            LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
            Exit Sub
        End If
        vsFileName = vs2005FileName
        If BuildOption = "" Then BuildOption = defaultVS2005Option
    ElseIf strLine = "Microsoft Visual Studio Solution File, Format Version 12.00" Then
        vsVersion = "2013"
        If vs2013FileName = "" Then 
            LogMessage("   ERROR: Unable to determine Visual Studio 2013 DEVENV.EXE location") 
            LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
            Exit Sub
        End If
        vsFileName = vs2013FileName
        If BuildOption = "" Then BuildOption = defaultVS2013Option
    Else
        LogMessage("   ERROR: Cannot determine Visual Studio version from " & ProjectFileNameBase & ".sln!")
        LogMessage("   " & strLine)
        LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
        Exit Sub
    End If

    If objFSO.FileExists(ProjectName & ".log") Then objFSO.DeleteFile(ProjectName & ".log")
    CommandLine = """" & ProjectName & ".sln"" /rebuild " & BuildOption & " /project """ & ProjectName & ".vbproj"" /out """ & ProjectName & ".log"""
    LogMessage("   DEVENV.exe " & CommandLine)
    LogMessage("   " & Execute(vsFileName & " " & CommandLine))
    If objFSO.FileExists(ProjectName & ".log") Then
        Set objStream = objFSO.OpenTextFile(ProjectName & ".log", ForReading, False)
        Do While (Not objStream.AtEndOfStream) 
            strLine = objStream.ReadLine
            If (vsVersion = "2003" And Left(Trim(strLine), Len("Rebuild All: ")) = "Rebuild All: ") Or _
               (vsVersion = "2005" And Left(Trim(strLine), Len("========== Rebuild All: ")) = "========== Rebuild All: ") Or _
               (vsVersion = "2013" And Left(Trim(strLine), Len("========== Rebuild All: ")) = "========== Rebuild All: ") Then
                If vsVersion = "2003" Then 
                    iPos = InStr(strLine, "Rebuild All: ") + Len("Rebuild All: ")
                ElseIf vsVersion = "2005" Then
                    iPos = InStr(strLine, "========== Rebuild All: ") + Len("========== Rebuild All: ")
                ElseIf vsVersion = "2013" Then
                    iPos = InStr(strLine, "========== Rebuild All: ") + Len("========== Rebuild All: ")
                End If
                'We can cheat here because we only ever build one project at a time (so these values will only ever be 1 or 0)...
                i = CInt(Mid(strLine, iPos, 1))
                iSucceeded = iSucceeded + i
                iPos = InStr(strLine, " succeeded, ") + Len(" succeeded, ")
                iFailed = iFailed + CInt(Mid(strLine, iPos, 1)) : If CInt(Mid(strLine, iPos, 1)) > 0 Then FailedList = FailedList & "   " & ProjectName & vbCrLf
                iPos = InStr(strLine, " failed, ") + Len(" failed, ")
                iSkipped = iSkipped + CInt(Mid(strLine, iPos, 1))
            End If
            If InStr(LCase(strLine), "vbc.exe ") = 0 Then LogMessage("   " & strLine)
        Loop
        'If i <> 0 And objFSO.FileExists(RealSource & "bin\" & ProjectName & ".dll") Then Call RegisterDotNet(RealSource & "bin\" & ProjectName & ".dll", Interop)
        objStream.Close
    End If
    LogMessage("   Running Total: " & iSucceeded & " succeeded, " & iFailed & " failed, " & iSkipped & " skipped")
    If FailedList <> "" Then LogMessage(vbCrLf & "   Failed Projects:" & vbCrLf & FailedList)
    LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")
End Sub
'This is an attempt to redirect a script not started using CScript to do so (To date, it is not having the desired effect)...
Private Sub ReDirectToCScript
    On Error Resume Next
    WScript.StdOut.WriteLine ""
    If Err.number <> 0 Then
        Command = "cscript " & WScript.ScriptFullName
        For i = 0 To WScript.Arguments.Count - 1
            Command = Command & " """ & WScript.Arguments(0) & """"
        Next
        Execute Command
        WScript.Quit
    End If
    On Error GoTo 0
End Sub
Private Sub DisplayHelp
    LogMessage("Usage:")
    LogMessage("Build.vbs <""Components"" | ""TestRPC.NET"" | ""FiRRe"" | ""All"" | ""COPY""> <Version> <BuildOption>")
    LogMessage("  Projects      Determines set of projects to process")
    LogMessage("  Version       Version number in the form <MajorVersion>.<MinorVersion> (i.e. ""5.1"")")
    LogMessage("  BuildOptions  Valid DEVENV command line entries (defaults to ""Debug|x86"") (optional)")
    LogMessage("")
    LogMessage("Software will be built to the developer's own V-Drive folder structure.")
End Sub

LogMessage("[" & WScript.ScriptName & vbTab & Now() & "]")
Projects = "FiRRe"
Version = ""
BuildOptions = ""
Select Case WScript.Arguments.Count
    Case 1
        Select Case WScript.Arguments(0)
            Case "?", "/?", "-?"
                DisplayHelp()
                WScript.Quit
        End Select
        Projects = WScript.Arguments(0)
    Case 2
        Projects = WScript.Arguments(0)
        Version = WScript.Arguments(1)
    Case 3
        Projects = WScript.Arguments(0)
        Version = WScript.Arguments(1)
        BuildOptions = WScript.Arguments(2)
    Case Else
        DisplayHelp()
        WScript.Quit
End Select
If Projects = "" then
    DisplayHelp()
    WScript.Quit
End If
LogMessage("   Current Directory: " & WshShell.CurrentDirectory)

vs2003FileName = "\Program Files\Microsoft Visual Studio .NET 2003\Common7\IDE\DEVENV.exe"
If objFSO.FileExists("C:" & vs2003FileName) Then
    vs2003FileName = "C:" & vs2003FileName
ElseIf objFSO.FileExists("D:" & vs2003FileName) Then
    vs2003FileName = "D:" & vs2003FileName
Else
    vs2003FileName = ""
End If
vs2005FileName = "\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\devenv.exe"
If objFSO.FileExists("C:" & vs2005FileName) Then
    vs2005FileName = "C:" & vs2005FileName
ElseIf objFSO.FileExists("D:" & vs2005FileName) Then
    vs2005FileName = "D:" & vs2005FileName
Else
    vs2005FileName = "\Program Files\Microsoft Visual Studio 8\Common7\IDE\devenv.exe"
    If objFSO.FileExists("C:" & vs2005FileName) Then
        vs2005FileName = "C:" & vs2005FileName
    ElseIf objFSO.FileExists("D:" & vs2005FileName) Then
        vs2005FileName = "D:" & vs2005FileName
    Else
        vs2005FileName = ""
    End If
End If
vs2013FileName = "\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\devenv.exe"
If objFSO.FileExists("C:" & vs2013FileName) Then
    vs2013FileName = "C:" & vs2013FileName
ElseIf objFSO.FileExists("D:" & vs2013FileName) Then
    vs2013FileName = "D:" & vs2013FileName
Else
    vs2013FileName = "\Program Files\Microsoft Visual Studio 12.0\Common7\IDE\devenv.exe"
    If objFSO.FileExists("C:" & vs2013FileName) Then
        vs2013FileName = "C:" & vs2013FileName
    ElseIf objFSO.FileExists("D:" & vs2013FileName) Then
        vs2013FileName = "D:" & vs2013FileName
    Else
        vs2013FileName = ""
    End If
End If

If vs2003FileName = "" And vs2005FileName = "" And vs2013FileName = "" Then
    LogMessage("   ERROR: Unable to determine DEVENV.EXE location")
    Set objFSO = Nothing
    WshShell.CurrentDirectory = startFolder
    WScript.Quit
Else
    If vs2003FileName = "" Then LogMessage("   WARNING: Visual Studio .NET 2003 does not appear to be installed.")
    If vs2005FileName = "" Then LogMessage("   WARNING: Visual Studio 2005 does not appear to be installed.")
    If vs2013FileName = "" Then LogMessage("   WARNING: Visual Studio 2013 does not appear to be installed.")
End If
LogMessage("   --------------------------------------------------------------------------------------------------------------------------------")

iSucceeded = 0
iFailed = 0 : FailedList = ""
iSkipped = 0
Call BuildProject("Components", "SIASSupport.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASWinsock.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASFTP.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASCL.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASCurrency.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASEmail.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASDB.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASRPC.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASRPC.NET\SIASBPE00000.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASRPC.NET\SIASBPE00001.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21090.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21110.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21120.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASRPC.NET\SIASBPE21130.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASRPC.NET\SIASBPE21150.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21170.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21180.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21190.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21200.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21210.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21220.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21230.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21240.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21250.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21260.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("Components", "SIASRPC.NET\SIASBPE21270.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASTask.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("Components", "SIASWatcher.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
'Call BuildProject("TestRPC.NET", "TestRPC.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)    'Needs all the omitted SIASBPE##### components above...
Call BuildProject("FiRRe", "Components\FiRReBase.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReData.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\CRUFLFiRRe.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReCustom.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReCustomBNYM.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReCustomCSTC.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReLookup.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReSetup.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReBilling.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FCTXDemo.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReFile.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReSystem.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReLoader.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReForecast.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReProcessing.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReCollection.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReReports.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReHelp.NET", Projects, BuildOptions, Version, True, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Components\FiRReSentinel.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Services\FiRReServer.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
'Call BuildProject("FiRRe", "Services\vbFiRReFTPmonitor.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Utilities\CRExplorer.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Utilities\LookupTest.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Utilities\FiRReMonitor.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Utilities\SplitFile.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Utilities\TXDemo.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Utilities\TX32001.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)
Call BuildProject("FiRRe", "Utilities\DBUtility.NET", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped)

Dim FiRRePath : FiRRePath = ""
VersionTag = "" : ProjectTag = ""
If Version <> "" Then VersionTag = " Version " & Version : ProjectTag = " v" & Version
If objFSO.FileExists("V:\FiRRe" & VersionTag & "\Utilities\FiRRe.NET\FiRRe.NET" & ProjectTag & ".sln") Then 
    FiRRePath = "Utilities\FiRRe.NET"
    Call BuildProject("FiRRe", FiRRePath, Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped) : FiRRePath = FiRRePath & "\bin"
Else
    Call BuildProject("FiRRe", "FiRRe.NET.sln", Projects, BuildOptions, Version, False, iSucceeded, iFailed, iSkipped) : FiRRePath = "bin"
End If

If UCase(Projects) = "FIRRE" Or UCase(Projects) = "ALL"  Or UCase(Projects) = "COPY" Then
    LogMessage("   Copying FiRReCustom Components to FiRRe.NET project folder...")
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*",         "V:\FiRRe" & VersionTag & "\" & FiRRePath, True)
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", "V:\FiRRe" & VersionTag & "\" & FiRRePath, True)
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", "V:\FiRRe" & VersionTag & "\" & FiRRePath, True)
    LogMessage("   Copying FiRReCustom Components to FiRReServer.NET project folder...")
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*",         "V:\FiRRe" & VersionTag & "\Services\FiRReServer.NET\bin", True)
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", "V:\FiRRe" & VersionTag & "\Services\FiRReServer.NET\bin", True)
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", "V:\FiRRe" & VersionTag & "\Services\FiRReServer.NET\bin", True)
    LogMessage("   Copying FiRReCustom Components to DBUtility.NET project folder...")
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*",         "V:\FiRRe" & VersionTag & "\Utilities\DBUtility.NET\bin", True)
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", "V:\FiRRe" & VersionTag & "\Utilities\DBUtility.NET\bin", True)
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", "V:\FiRRe" & VersionTag & "\Utilities\DBUtility.NET\bin", True)
    LogMessage("   Copying FiRReCustom Components to LookupTest.NET project folder...")
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustom.NET\bin\FiRReCustom.NET.*",         "V:\FiRRe" & VersionTag & "\Utilities\LookupTest.NET\bin", True)
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomBNYM.NET\bin\FiRReCustomBNYM.NET.*", "V:\FiRRe" & VersionTag & "\Utilities\LookupTest.NET\bin", True)
    Call objFSO.CopyFile("V:\FiRRe" & VersionTag & "\Components\FiRReCustomCSTC.NET\bin\FiRReCustomCSTC.NET.*", "V:\FiRRe" & VersionTag & "\Utilities\LookupTest.NET\bin", True)
End If
'Cannot update registry without Administrator access rights...
'If objFSO.FileExists(startFolder & "\ProjectMRUList.reg") Then 
'    LogMessage("Resetting Visual Studio MRU List...")
'    Execute "RegEdit.exe /s " & startFolder & "\ProjectMRUList.reg"
'End If
LogMessage(vbCrLf & "Script complete @ " & Now())
Set objFSO = Nothing
WshShell.CurrentDirectory = startFolder
WScript.Quit
