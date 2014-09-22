Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002Const REG_SZ = 1Const REG_EXPAND_SZ = 2Const REG_BINARY = 3Const REG_DWORD = 4Const REG_MULTI_SZ = 7Dim WshShell:	Set WshShell = CreateObject("WScript.Shell")
Dim objFSO:		Set objFSO = CreateObject("Scripting.FileSystemObject")

WScript.StdOut.WriteLine WshShell.RegRead("HKCR\SIASUTL.clsTrace\Clsid\")

Dim stringvalue
Dim strComputer:	strComputer = "."
Dim oReg:			Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
'Dim strKeyPath:		strKeyPath = "SIASUTL.*"'strKeyPath = "SYSTEM\CurrentControlSet\Control\Lsa"'oReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes'For i=0 To UBound(arrValueNames)'    stringvalue = arrValueNames(i) '    Select Case arrValueTypes(i)'        Case REG_SZ'            stringvalue = stringValue & "(String)"'        Case REG_EXPAND_SZ'            stringvalue = stringValue & "(Expanded String)"'        Case REG_BINARY'            stringvalue = stringValue & "(Binary)"'        Case REG_DWORD'            stringvalue = stringValue & "(DWORD)"'        Case REG_MULTI_SZ'            stringvalue = stringValue & "(Multi String)"'    End Select '    Wscript.Echo stringValue'NextOn Error Resume NextstrComputer = "."Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")Set colItems = objWMIService.ExecQuery("Select * From Win32_ProgIDSpecification Where ProgID Like 'SIASUTL.%'")For Each objItem in colItems    Wscript.Echo "Caption: " & objItem.Caption    Wscript.Echo "Check ID: " & objItem.CheckID    Wscript.Echo "Name: " & objItem.Name    Wscript.Echo "Parent: " & objItem.Parent    Wscript.Echo "ProgID: " & objItem.ProgIDNext