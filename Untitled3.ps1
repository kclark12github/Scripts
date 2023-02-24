$Objs2 = @()
$Count = 0
Get-AppxPackage | 
    ForEach-Object {
        $Manifest = Get-AppxPackageManifest $_;
        $Package = $Manifest.Package;
        $DisplayName = $_.Name;
        $Version = $_.Version;
        $Publisher = "";
        $InstallDate = "";
        $EstimatedSize = "";
        $InstallLocation = "";
        $UninstallString = "";
        $Source = "Get-AppxPackage";

        $IsFramework = $_.IsFramework;
        $SignatureKind = $_.SignatureKind;

        if ($null -ne $Package.Identity -and $null -eq $Version) {
            $Version = $Package.Identity.Version;
        }

        if ($null -ne $Package.Properties) {
            $DisplayName = $Package.Properties.DisplayName;
            $Publisher = $Package.Properties.PublisherDisplayName;
        }
        
        $DisplayName

        if ($InstallDate -eq "") {
            $InstallDate = "";
        }
 
#            if ($null -ne $_.Package.Applications) {
#                $DisplayName = "Application:"+$_.Package.Applications.Application.VisualElements.DisplayName;
#            } elseif ($null -ne $_.Package.Identity) {
#                $DisplayName = "Identity:"+$_.Package.Identity.Name;
#                $Version = "Identity:"+$_.Package.Identity.Version;
#                $Publisher = "Identity:"+$_.Package.Identity.Publisher
#            } else {
#                $DisplayName = "Unknown";
#            }

        $Obj2 = [PSCustomObject]@{
            DisplayName = $DisplayName;
            Version = $Version;
            Publisher = $Publisher;
            InstallDate = $InstallDate;
            EstimatedSize = $EstimatedSize;
            InstallLocation = $InstallLocation;
            UninstallString = $UninstallString;
            Source = $Source;
            IsFramework = $IsFramework;
            SignatureKind  = $SignatureKind;
            }
        $Objs2 += $Obj2
        $Count += 1
    }
    $Objs2 | Where-Object { $_.DisplayName }  | Sort-Object -Property DisplayName
    $Count
