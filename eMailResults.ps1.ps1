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
param([string]$Subject, [string]$Body, [string]$LogFile)

$User = "kfc12"
$PWord = ConvertTo-SecureString -String "cvn80BigE!" -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord

$LogContent = Get-Content -Path $LogFile -Raw
$Body = $Body + $LogContent

$email = @{
    From = "kfc12@comcast.net"
    To = "kfc12@comcast.net"
    Subject = $Subject
    SMTPServer = "smtp.comcast.net"
    Port = 587
    Body = $Body
    Attachments = $LogFile
    Credential = $Credential
}
send-mailmessage @email -UseSsl
