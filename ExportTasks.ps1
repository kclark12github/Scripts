param([string]$Root="$($env:COMPUTERNAME)",[string]$BackupFolder,[string]$LogPath)
function Format-Elapsed {
    Param($Start, $End)
    $Elapsed = ""
    $ts = New-TimeSpan -start $Start -end $End
    if ($ts.Days -gt 0) {$Elapsed += "$($ts.Days) Days, "}
    if ($ts.Hours -gt 0) {$Elapsed += "$($ts.Hours) Hours, "}
    if ($ts.Minutes -gt 0) {$Elapsed += "$($ts.Minutes) Minutes, "}
    $Elapsed += "{0}.{1:000} Seconds" -f $ts.Seconds, $ts.Milliseconds
    return $Elapsed
}
function Send-Mail {
    Param($AppName, $EmailBody, $LogPath)
    $PW = ConvertTo-SecureString $env:SMTP_PW -AsPlainText -Force
    $Creds = New-Object System.Management.Automation.PSCredential ($env:SMTP_USER, $PW)
    $Server = $env:SMTP_ADDRESS+":"+$env:SMTP_PORT
    Send-MailMessage -From "$env:USERNAME <$env:My_EMAIL>" -To "$env:USERNAME <$env:My_EMAIL>" -Subject "$AppName Succeeded!" -Body $EmailBody -BodyAsHtml -Attachments $LogPath -SmtpServer $env:SMTP_ADDRESS -Port $env:SMTP_PORT -Credential $Creds
}
function Write-Log {
    Param($Message, $Path = ".")
    function TS {return "[{0:MM/dd/yy} {0:HH:mm:ss tt}]" -f (Get-Date)}
    Write-Message -Message "$(TS) $Message" -Path $Path
}
function Write-Message {
    Param($Message, $Path = ".")
    "$Message" | Tee-Object -FilePath $Path -Append | Write-Output
}
function Write-Separator {
    Param($Path = ".")
    #                      1         2         3         4         5         6         7         8
    #             12345678901234567890123456789012345678901234567890123456789012345678901234567890
    $Separator = "--------------------------------------------------------------------------------"
    Write-Message -Message $Separator -Path $Path
}
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

    $EmailBody += Export-Tasks -OutFile "$($BackupFolder)$($Root)\Ken's Tasks.xml" -LogPath $LogPath
    
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "<br />$Message"
    Write-Separator -Path $LogPath


    #write-output "$EmailBody"


    &"$PSScriptRoot\eMailResults.ps1" -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
    #Send-Mail -AppName $AppName -EmailBody $EmailBody -LogPath $LogPath
}