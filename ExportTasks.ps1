#ExportTasks.ps1
#   Script to export custom Task Scheduler Tasks (for backup/restore purposes)...
#   Copyright © 1998-2026, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   11/13/26    Added History;
#=================================================================================================================================
param([string]$Root="$($env:COMPUTERNAME)",[string]$BackupFolder,[string]$LogPath)
Import-Module MyTools
function Export-Tasks {
    Param($OutFile, $LogPath)

    $Message = "Exporting Ken's Scheduled Tasks from $Root to $OutFile"
    Write-Message -Message $Message -Path $LogPath

    Get-ScheduledTask -TaskPath "\Ken's Tasks\" | ForEach-Object{Export-ScheduledTask -TaskPath $_.TaskPath -TaskName $_.TaskName } > "$OutFile"

    $Message + "<br /><br />"
}
.{
    $AppName = "ExportTasks"
    $StartTime = Get-Date
    $Message = "[$AppName © $("{0:yyyy}" -f $StartTime), Ken Clark                       $("{0:MM/dd/yy} {0:hh:mm:ss tt}" -f $StartTime)]"

    if ($BackupFolder.Equals("")) {$BackupFolder = "$($Env:OneDrive)\Backups\$Root\"}  #$($Env:BackupRoot)\
    if ($LogPath.Equals("")) {$LogPath = "$($BackupFolder)$Root.ExportTask.log"}
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

    $EmailBody += Export-Tasks -OutFile "$($BackupFolder)Ken's Tasks.xml" -LogPath $LogPath
    
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "<br />$Message"
    Write-Separator -Path $LogPath
    Send-EmailResults  -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
}