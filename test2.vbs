strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM CIM_DataFile where Path='H:\\Development\\Backups\\TFS\\' And CreationDate<='" & DateAdd("d",-10,Now()) & "'") 
For Each objItem in colItems 
    Wscript.Echo "-----------------------------------"
    Wscript.Echo "CIM_DataFile instance"
    Wscript.Echo "-----------------------------------"
    Wscript.Echo "AccessMask: " & objItem.AccessMask
    Wscript.Echo "Archive: " & objItem.Archive
    Wscript.Echo "LastAccessed: " & objItem.LastAccessed
    Wscript.Echo "LastModified: " & objItem.LastModified
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "Readable: " & objItem.Readable
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "System: " & objItem.System
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "Writeable: " & objItem.Writeable
Next