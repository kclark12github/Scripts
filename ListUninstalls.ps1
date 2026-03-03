    $Objs = @()
    $RegKey = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

    $InstalledAppsInfos = Get-ItemProperty -Path $RegKey

    Foreach($InstalledAppsInfo in $InstalledAppsInfos)
    {
        $Obj = [PSCustomObject]@{Computer=$ComputerName;
                                 DisplayName = $InstalledAppsInfo.DisplayName;
                                 DisplayVersion = $InstalledAppsInfo.DisplayVersion;
                                 Publisher = $InstalledAppsInfo.Publisher}
        $Objs += $Obj
    }
    $Objs | Where-Object { $_.DisplayName } 
