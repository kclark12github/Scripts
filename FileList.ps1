#FileList.ps1
#   Script to inventory files (for backup/restore purposes)...
#   Copyright © 1998-2026, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   11/13/26    Added History;
#=================================================================================================================================
param([string]$Root="$($env:COMPUTERNAME)",[string]$BackupFolder,[string]$LogPath)
Import-Module MyTools
function Write-Files {
    Param($Path = ".", $OutFile, $LogPath)
    $errMessage = ""
    $Message = "Listing Files from $Path to $OutFile"
    Write-Message -Message $Message -Path $LogPath

    Get-ChildItem -Path "$Path" -Force -Recurse -ErrorVariable errMessage | Select Fullname,Length,CreationTime,LastWriteTime | Export-Csv -Path "$OutFile" -NoTypeInformation
    if (-not $errMessage.Equals("")) {Write-Message -Message $errMessage -Path $LogPath}
    "<br />" #Return break tags appended to message
}
.{
    $AppName = "FileList"
    $StartTime = Get-Date
    $Message = "[$AppName © $("{0:yyyy}" -f $StartTime), Ken Clark                       $("{0:MM/dd/yy} {0:hh:mm:ss tt}" -f $StartTime)]"

    if ($BackupFolder.Equals("")) {$BackupFolder = "$($Env:OneDrive)\Backups\$Root\"}  #$($Env:BackupRoot)\
    if ($LogPath.Equals("")) {$LogPath = "$($BackupFolder)$Root.FileList.log"}
    Write-Message -Message $Message -Path $LogPath
    $EmailBody = $Message.Replace("©", "&copy;") + "<br />"
    $Message = "Root: $Root; BackupFolder: $BackupFolder; LogPath: $LogPath;"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br /><br />"

    if ($Root.Contains("TEST")) {
        $EmailBody += Write-Files -Path "\\Alpha\Public\Pluralsight\Entity Framework Core 2 - Getting Started\03 - demos\M3 Demos\SamuraiApp\packages\Microsoft.Extensions.DependencyInjection.Abstractions.2.0.0\lib" -OutFile "$($BackupFolder)TEST\Test.csv" -LogPath $LogPath
    }

    if ($Root.Contains("NAS")) {
        $EmailBody += Write-Files -Path "\\Alpha\Public" -OutFile "$($BackupFolder)ALPHA\Alpha Public Files.csv" -LogPath $LogPath
        #$EmailBody += Write-Files -Path "\\Alpha\Backups" -OutFile "$($BackupFolder)ALPHA\Alpha Backups Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Bravo\Public" -OutFile "$($BackupFolder)BRAVO\Bravo Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Charlie\Public" -OutFile "$($BackupFolder)CHARLIE\Charlie Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Delta\Public" -OutFile "$($BackupFolder)DELTA\Delta Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Echo\Public" -OutFile "$($BackupFolder)ECHO\Echo Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Foxtrot\Public" -OutFile "$($BackupFolder)FOXTROT\Foxtrot Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Golf\Public" -OutFile "$($BackupFolder)GOLF\Golf Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Golf\Backups" -OutFile "$($BackupFolder)GOLF\Golf Backups Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Hotel\Public" -OutFile "$($BackupFolder)HOTEL\Hotel Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Hotel\Backups" -OutFile "$($BackupFolder)HOTEL\Hotel Backups Files.csv" -LogPath $LogPath
    }

    if ($Root.Contains("7X4L243")) {
        $EmailBody += Write-Files -Path "\\7X4L243\C" -OutFile "$($BackupFolder)7X4L243 C Files.csv" -EmailBody $EmailBody -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\7X4L243\D" -OutFile "$($BackupFolder)7X4L243 D Files.csv" -EmailBody $EmailBody -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\7X4L243\E" -OutFile "$($BackupFolder)7X4L243 E Files.csv" -EmailBody $EmailBody -LogPath $LogPath
    }

    if ($Root.Contains("AAIUCI4")) {
        $EmailBody += Write-Files -Path "\\AAIUCI4\C" -OutFile "$($BackupFolder)AAIUCI4 C Files.csv" -EmailBody $EmailBody -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\AAIUCI4\D" -OutFile "$($BackupFolder)AAIUCI4 D Files.csv" -EmailBody $EmailBody -LogPath $LogPath
    }
    if ($Root.Contains("GGGSCP1")) {
        $EmailBody += Write-Files -Path "\\GGGSCP1\C" -OutFile "$($BackupFolder)GGGSCP1 C Files.csv" -EmailBody $EmailBody -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\GGGSCP1\D" -OutFile "$($BackupFolder)GGGSCP1 D Files.csv" -EmailBody $EmailBody -LogPath $LogPath
    }
    
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br /><br />"
    Write-Separator -Path $LogPath

    Send-EmailResults  -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
}