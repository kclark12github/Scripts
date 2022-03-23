#B54J71N Database Backup.ps1
#   Script Move SQL Server Database Backups from Local Folder to NAS Backup Folder Using RoboCopy...
#   Copyright © 1998-2022, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   03/23/22    Replaced hard-coded \\Alpha\Backups with $Env:BackupRoot for greater flexibility;
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
$Target = "$($env:BackupRoot)\Databases\B54J71N\Backups"
$LogFile = "$($env:BackupRoot)\Databases\B54J71N\B54J71N Database Backup.log"
&"$PSScriptRoot\RoboCopyMirror.ps1" -Title $Title -Source $Source -Target $Target -LogFile $LogFile
$DStatus = $LASTEXITCODE


$Title = "B54J71N Database Backup Logs"
$Source = "C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\Log\*.txt"
$Target = "$($env:BackupRoot)\Databases\B54J71N"
$LogFile = "$($env:BackupRoot)\Databases\B54J71N\B54J71N Database Backup Logs.log"
&"$PSScriptRoot\RoboCopyMirror.ps1" -Title $Title -Source $Source -Target $Target -Files "*.txt" -LogFile $LogFile
$LStatus = $LASTEXITCODE

exit ($DStatus -bor $LStatus)