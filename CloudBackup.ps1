#CloudBackup.ps1
#	PowerShell Script Used to Synchronize Two WD myCloud Mirror Folders (Suitable for Scheduling)...
#   Copyright © 2006-2015, Ken Clark
#*********************************************************************************************************************************
#
#   Modification History:
#   Date:       Developer:		Description:
#   09/30/15    Ken Clark		Created;
#=================================================================================================================================
[CmdletBinding()]
param ([parameter(Mandatory=$true)] [string]$sourceCloud, 
    [parameter(Mandatory=$true)] [string]$targetCloud, 
    [parameter(Mandatory=$true)] [string]$Folder,
    [switch]$test)
$body = "";$subject = "";$fileCount = 0
$Error.Clear()
function escapedPath([string]$path){
    if ($path -contains "``") 
    {
        $path.Replace("``","````")
    }
    else
    {
        $path.Replace("[","``[").Replace("]","``]")
    }
}
function fileExists([string]$path){
    ([System.IO.FileInfo]$path).Exists
}
function folderExists([string]$path){
    ([System.IO.DirectoryInfo]$path).Exists
}
function formatNow{
    Get-Date -Format 'MM/dd/yyyy HH:mm:ss'
}
function logMessage([string]$message){
    Write-Host $message; "$message`n"
}
function sendMail{
     #Creating SMTP server object
     $smtpClient = new-object Net.Mail.SmtpClient("smtp.comcast.net",587)   #465
     $smtpClient.Credentials = new-object System.Net.NetworkCredential("kfc12", "cvn65BigE")
     $smtpClient.EnableSsl = $true
     #Email structure 
     $msg = new-object Net.Mail.MailMessage("kfc12@comcast.net","kfc12@comcast.net")
     $msg.sender = "kfc12@comcast.net"
     $msg.subject = $subject
     $msg.body = $body

     #Write-Host "Sending Email"
     $smtpClient.Send($msg)
}
try
{
    $body += logMessage("[CloudBackup.ps1`t$(formatNow)]"); $sw = [Diagnostics.StopWatch]::StartNew()
    $subject = "CloudBackup \\$sourceCloud\$Folder"
    if($sourceCloud -eq "") {throw "sourceCloud must be provided!"}; if(!(test-path -path "\\$sourceCloud\$Folder" -PathType Container)) {throw "\\$sourceCloud\$Folder does not exist or is inaccessible!"}
    if($targetCloud -eq "") {throw "targetCloud must be provided!"}; if(!(test-path -path "\\$targetCloud\$Folder" -PathType Container)) {throw "\\$targetCloud\$Folder does not exist or is inaccessible!"}
    if($Folder -eq "") {throw "Folder must be provided!"} elseif ($Folder.ToUpper() -eq "PUBLIC") {throw "Public Folder is too big to backup!"}
    $sourceFolder = "\\$sourceCloud\$Folder"
    $body += logMessage("Scanning $sourceFolder...")
    if ($test) {$body += logMessage("**** TEST MODE ****")}
    #Get the directories in $sourceFolder...
    $directories = Get-ChildItem $sourceFolder -Recurse | Where-Object { $_.PSIsContainer } | Sort-Object -Property Name
    foreach($directory in $directories)
    {
        #Write-Host "Directory: $($directory.FullName)"
        $targetFolder = $directory.FullName.Replace("\\$sourceCloud\$Folder","\\$targetCloud\$Folder")
        if (-not (folderExists($targetFolder)))
        {
            $body += logMessage("`t$(formatNow)`tCreating $targetFolder")
            if ($test) 
            {
                New-Item -Path $(escapedPath($targetFolder)) -ItemType directory -Force -WhatIf
            }
            else
            {
                New-Item -Path $(escapedPath($targetFolder)) -ItemType directory -Force
            }
        }

        $files = Get-ChildItem -LiteralPath $directory.FullName | Where-Object {! ($_.PSIsContainer)}  | Sort-Object -Property Name
        foreach($file in $files)
        {
            $targetPath = "$targetFolder\$($file.Name)"
            $doCopy = $false
            if (-not (fileExists($targetPath)))
            {
                $doCopy = $true
            }
            elseif ((Get-ItemProperty -LiteralPath $targetPath).LastWriteTime -lt $file.LastWriteTime)
            {
                $doCopy = $true
            }
            if ($doCopy)
            {
                $fileCount += 1
                $body += logMessage("`t$(formatNow)`t$($file.FullName)")
                if ($test) 
                {
                    Copy-Item -LiteralPath $file.FullName -Destination $targetFolder -Force -WhatIf
                }
                else
                {
                    Copy-Item -LiteralPath $file.FullName -Destination $targetFolder -Force
                }
            }
        }
    }
    $subject += " Succeeded"
}
catch [System.Exception]
{
    $body += logMessage($Error[0].Exception.Message+"`n"+$Error[0].Exception+"`n"+$Error[0].ScriptStackTrace)
    $subject += " Failed"
}
finally
{
    $body += logMessage("`t$fileCount Total Files copied")
    if ($test) {$body += logMessage("**** TEST MODE ****")}
    $sw.Stop(); $elapsed = $sw.Elapsed.ToString()
    $body += logMessage("CloudBackup Complete @ $(formatNow) ($elapsed elapsed)")
    sendMail
}

##--------------------------------------------------------------------------------  
#$date = Get-Date -Format d.MMMM.yyyy 
#New-PSDrive -Name "Backup" -PSProvider Filesystem -Root "\\T_Server\Tally" 
#Remove-PSDrive "Backup"   
#$destination = "backup:\$date" 
##--------------------------------------------------------------------------------  
#$source = "C:\fso"
#$destination = "C:\backup\FSO_Backup.zip"
#If(Test-path $destination) {Remove-item $destination}
#Add-Type -assembly "system.io.compression.filesystem"
#[io.compression.zipfile]::CreateFromDirectory($Source, $destination)
#Send-MailMessage -From "ScriptingGuys@Outlook.com" -To "ScriptingGuys@Outlook.com" `
# -Attachments $destination -Subject "$(Split-Path $destination -Leaf)" -Body "File attached" `
# -SmtpServer "smtp-mail.outlook.com" -UseSsl -Credential "ScriptingGUys@Outlook.com"
# Remove-Item $destination
##--------------------------------------------------------------------------------  
#$newFolders = dir \\ALPHA\Backups | ? {$_.PSIsContainer} | ? {$_.LastWriteTime -gt (Get-Date).AddDays(-7)} 
#$newFolders | % { copy $_.FullName  c:\temp\archive -Recurse -Force}
##--------------------------------------------------------------------------------  
#$sevenDaysAgo = (Get-Date).AddDays(-7);
#$newFolders = dir \\ALPHA\Backups | ? {$_.PSIsContainer}
## Get the directories in X:\EmpInfo.
#$directories = Get-ChildItem \\ALPHA\Backups | Where-Object { $_.PSIsContainer };
## Loop through the directories.
#foreach($directory in $directories)
#{
#    # Check in the directory for a file within the last seven days.
#    if (Get-ChildItem .\UnitTests -Recurse | Where-Object { $_.LastWriteTime -ge $sevenDaysAgo })
#    {
#        $directory
#    }
#}
