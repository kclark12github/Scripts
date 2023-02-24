    $Objs2 = @()
    Get-AppxPackage | 
        ForEach-Object{

            $Obj2 = [PSCustomObject]@{
                Computer=$ComputerName;
                DisplayName = $DisplayName;
                #DisplayVersion = $_.Package.Identity.Version;
                #Publisher = $_.Package.Properties.PublisherDisplayName
                #Application = $_.Package.Applications.Application;
                Identity = $_.Package.Identity;
                Properties = $_.Package.Properties;
                }
            $Objs2 += $Obj2
        }
        $Objs2 | Where-Object { $_.DisplayName } 
