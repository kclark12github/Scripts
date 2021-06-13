#B54J71N Database Backup.ps1
#   Script Move SQL Server Database Backups from Local Folder to NAS Backup Folder Using RoboCopy...
#   Copyright © 1998-2021, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   06/13/21    Created;
#=================================================================================================================================
#Notes to Self:
#=================================================================================================================================
# 

#Enable -Verbose option
[CmdletBinding()]
param([switch]$Test)

$Title = "B54J71N Database Backup"
$Source = "D:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\Backups"
$Target = "\\Alpha\Backups\Databases\B54J71N"
$LogFile = "\\Alpha\backups\Databases\B54J71N\B54J71N Database Backup.log"
&"$PSScriptRoot\RoboCopyMirror.ps1" -Title $Title -Source $Source -Target $Target -LogFile $LogFile
exit $LASTEXITCODE