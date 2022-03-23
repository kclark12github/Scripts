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
$PWord = ConvertTo-SecureString $env:SMTP_PW -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord

$LogContent = Get-Content -Path $LogFile -Raw
$Body = $Body + $LogContent

$email = @{
    From = "$env:USERNAME <$env:My_EMAIL>"
    To = "$env:USERNAME <$env:My_EMAIL>"
    Subject = $Subject
    SMTPServer = $env:SMTP_ADDRESS
    Port = $env:SMTP_PORT
    Body = $Body
    BodyAsHtml = $AsHTML
    Attachments = $LogFile
    Credential = $Credential
}
Send-MailMessage @email -UseSsl

