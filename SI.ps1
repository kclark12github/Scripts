$array = @()

#Define the variable to hold the location of Currently Installed Programs
$UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"

#Create an instance of the Registry Object and open the HKLM base key
$reg32=[microsoft.win32.registrykey]::OpenBaseKey('LocalMachine','Default') #Registry32

#Drill down into the Uninstall key using the OpenSubKey Method
$regkey=$reg32.OpenSubKey($UninstallKey) 

#Retrieve an array of string that contain all the subkey names
$subkeys=$regkey.GetSubKeyNames() 

#Open each Subkey and use GetValue Method to return the required values for each
foreach($key in $subkeys){
    $thisKey=$UninstallKey+"\\"+$key 
    $thisSubKey=$reg32.OpenSubKey($thisKey) 
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "Source" -Value 'Standard'
    $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
    $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
    $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
    $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
    $array += $obj
} 

#Now do it all again for Registry64...
$UninstallKey="SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
$reg64=[microsoft.win32.registrykey]::OpenBaseKey('LocalMachine','Default') #Registry64
$regkey=$reg64.OpenSubKey($UninstallKey) 
$subkeys=$regkey.GetSubKeyNames() 
foreach($key in $subkeys){
    $thisKey=$UninstallKey+"\\"+$key 
    $thisSubKey=$reg64.OpenSubKey($thisKey) 
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "Source" -Value 'Wow6432Node'
    $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
    $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
    $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
    $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
    $array += $obj
} 

$array | Where-Object { $_.DisplayName } | sort Displayname | select Source, DisplayName, DisplayVersion, Publisher | ft -auto

#Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall | 
#ForEach-Object {Get-ItemProperty $_.PsPath} | 
#where {$_.Displayname -notlike "*update*" -and $_.DisplayName -notlike ""} | 
#sort Displayname | select DisplayName, Publisher 

#Get-ChildItem HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | 
#ForEach-Object {Get-ItemProperty $_.PsPath} | 
#where {$_.Displayname -notlike "*update*" -and $_.DisplayName -notlike ""} | 
#sort Displayname | select DisplayName, Publisher 
