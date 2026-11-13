#eMailResults.ps1
#   Script to E-Mail Script Output Results ...
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
