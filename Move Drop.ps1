#Move Drop.ps1
#   Script Move Local Azure-Updated Drop Folder to NAS Folder Using RoboCopy...
#   Copyright © 1998-2021, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   06/13/21    Created;
#=================================================================================================================================
#Notes to Self:
#Apparently the locally hosted Azure Agent cannot access NAS drives (probably to do with SYSTEM account access), so as a 
#workaround I've configured it to write to a local Drop folder, and this script periodically moves it out to the NAS.
#=================================================================================================================================
# 

#Enable -Verbose option
[CmdletBinding()]
param([switch]$Test)

$Title = "Move Drop"
$Source = "D:\Drop"
$Target = "\\Alpha\Public\Drop"
$LogFile = "\\Alpha\Public\Drop\Move Drop.log"
&"$PSScriptRoot\RoboCopyMove.ps1" -Title $Title -Source $Source -Target $Target -LogFile $LogFile -SuppressFailure
exit $LASTEXITCODE