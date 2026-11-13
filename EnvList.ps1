#EnvList.ps1
#   Script to export environment variable definitions  (for backup/restore purposes)...
#   Copyright © 1998-2026, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   11/13/26    Added History;
#=================================================================================================================================
param([string]$Root="$($env:COMPUTERNAME)",[string]$BackupFolder,[string]$LogPath)
Import-Module MyTools
function Write-Vars {
    Param($OutFile, $LogPath)

    $Message = "Listing Environment Variables on $Root to $OutFile"
    Write-Message -Message $Message -Path $LogPath

    $fileName = Split-Path -Path $OutFile -Leaf
    $tempOutFile = Join-Path -Path $env:TEMP -ChildPath $fileName

    $ExportData = [PSCustomObject]@{
        User    = [System.Environment]::GetEnvironmentVariables("User")
        Machine = [System.Environment]::GetEnvironmentVariables("Machine")
    }

    $ExportData | ConvertTo-Json -Depth 3 | Out-File -FilePath $tempOutFile -Encoding utf8
    Move-Item -Path $tempOutFile -Destination $OutFile -Force

    $Message + "<br /><br />"
}
.{
    $AppName = "EnvList"
    $StartTime = Get-Date
    $Message = "[$AppName © $("{0:yyyy}" -f $StartTime), Ken Clark                       $("{0:MM/dd/yy} {0:hh:mm:ss tt}" -f $StartTime)]"

    if ($BackupFolder.Equals("")) {$BackupFolder = "$($Env:OneDrive)\Backups\$Root\"}  #$($Env:BackupRoot)\
    if ($LogPath.Equals("")) {$LogPath = "$($BackupFolder)$Root.EnvList.log"}
    Write-Message -Message $Message -Path $LogPath
    $EmailBody = $Message.Replace("©", "&copy;") + "<br />"
    $Message = "Root: $Root; BackupFolder: $BackupFolder; LogPath: $LogPath;"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"

    if ($Root.Contains("TEST")) {
        $Message = "TEST Script"
        Write-Message -Message $Message -Path $LogPath
        $EmailBody += $Message + "<br /><br />"
    }

    $EmailBody += Write-Vars -OutFile "$($BackupFolder)$($Root).Environment Variables.json" -LogPath $LogPath
    
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "<br />$Message"
    Write-Separator -Path $LogPath

    Send-EmailResults  -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
}