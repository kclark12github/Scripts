#MyTools.psm1
#   Utility Function Module...
#   Copyright © 1998-2026, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   11/13/26    Created;
#=================================================================================================================================
#Notes to Self:
#  This module needs to be deployed in PowerShell's standard module directory: $HOME\Documents\PowerShell\Modules\MyTools
#=================================================================================================================================
function Send-EmailResults {
    param([string]$Subject, [string]$Body, [string]$LogFile, [switch]$AsHTML=$False)

    $User = $env:SMTP_USER
    $PWord = ConvertTo-SecureString $env:SMTP_PW -AsPlainText -Force; $PWord.MakeReadOnly()
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord

    #$LogContent = Get-Content -Path $LogFile -Raw
    #$Body = $Body + "<br />" + $LogContent.Replace("`n", "<br />").Replace("©", "&copy;")

    $email = @{
        From = "$env:USERNAME <$env:My_EMAIL>"
        To = "$env:USERNAME <$env:My_EMAIL>"
        Subject = $Subject
        Server = $env:SMTP_ADDRESS
        Port = [int]$env:SMTP_PORT
        Attachments = $LogFile
        Credential = $Credential
        SecureSocketOptions = 'Auto' # Automatically negotiates SSL/TLS cleanly
    }
    if ($AsHTML) {
        $email['HTML'] = $Body
    } else {
        $email['Text'] = $Body
    }
    $MailResult = Send-EmailMessage @email
    if (-not $MailResult.Status) {
        Write-Warning "Email delivery failed: $($MailResult.Error)"
    }
}
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
function ConvertTo-SizeString {
    Param($size)
    $SizeMB = $null
    if ($null -ne $size -and $size -gt 0) {
        $CalculatedMB = [math]::Round(($size / 1024), 2)
        $SizeMB = "$CalculatedMB MB"
    }
    $SizeMB
}
function ConvertTo-Date {
    Param($DirtyDate)
    $ParsedDate = [DateTime]::MinValue
    $CleanDate  = $null
    if (-not [string]::IsNullOrWhiteSpace($DirtyDate)) {
        if ([DateTime]::TryParseExact($DirtyDate.Trim(), "yyyyMMdd", $null, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
            $CleanDate = $ParsedDate.Date
        }
    }
    $CleanDate
}
function Write-Log {
    Param($Message, $Path = ".")
    function TS {return "[{0:MM/dd/yy} {0:HH:mm:ss tt}]" -f (Get-Date)}
    Write-Message -Message "$(TS) $Message" -Path $Path
}
function Write-Message {
    Param($Message, $Path = ".")
    #"$Message" | Tee-Object -FilePath $Path -Append | Write-Output
    Add-Content -Path $Path -Value $Message -Encoding utf8
    Write-Output $Message
}
function Write-Separator {
    Param($Path = ".")
    #                      1         2         3         4         5         6         7         8
    #             12345678901234567890123456789012345678901234567890123456789012345678901234567890
    $Separator = "--------------------------------------------------------------------------------"
    Write-Message -Message $Separator -Path $Path
}
