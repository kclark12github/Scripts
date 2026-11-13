$DevFolder      = "D:\Users\kfc12\Source\repos\kclark12github\Scripts"        # Your Git / VS Code folder
$LocalProdPath  = "D:\Program Files (x86)\KClark Software"                      # Local folder where Task Scheduler runs scripts
$RemoteLaptop   = "AAIUCI4"                # Hostname or IP of your laptop
$RemoteProdPath = "D:\Program Files (x86)\KClark Software"                      # Destination folder on your laptop

$ModuleName      = "MyTools"
$DevModuleFolder = Join-Path $DevFolder "Modules\$ModuleName"
$LocalBaseModDir = ($env:PSModulePath -split ';' | Where-Object { $_ -notlike "$env:ProgramFiles*" -and $_ -notlike "$env:SystemRoot*" })[0]
$LocalModulePath = Join-Path $LocalBaseModDir $ModuleName

# List only the production scripts you want deployed (keeps test scripts out)
$ScriptsToDeploy = @(
    "eMailResults.ps1",
    "EnvList.ps1",
    "ExportTasks.ps1",
    "FileList.ps1",
    "HarvestGOGGoodies.ps1",
    "SoftwareInventory.ps1"
)

# Deploy to Local Installation Folder
Write-Host "Deploying to local installation folder ($LocalProdPath)..." -ForegroundColor Cyan
if (-not (Test-Path $LocalProdPath)) {
    New-Item -ItemType Directory -Path $LocalProdPath -Force | Out-Null
}

foreach ($Script in $ScriptsToDeploy) {
    $SourceFile = Join-Path $DevFolder $Script
    if (Test-Path $SourceFile) {
        Copy-Item -Path $SourceFile -Destination $LocalProdPath -Force
        Write-Host "  [Local] Copied $Script" -ForegroundColor Green
    } else {
        Write-Warning "  [Local] Missing source file: $SourceFile"
    }
}

if (Test-Path $DevModuleFolder) {
    Copy-Item -Path "$DevModuleFolder\*" -Destination $LocalModulePath -Recurse -Force
    Write-Host "  [Local] Deployed module: $ModuleName" -ForegroundColor Green
} else {
    Write-Warning "  [Local] Missing module folder: $DevModuleFolder"
}

# Deploy to Remote Laptop (if online)
Write-Host "`nChecking if $RemoteLaptop is reachable..." -ForegroundColor Cyan
if (Test-Connection -ComputerName $RemoteLaptop -Count 1 -Quiet) {
    try {
        $Session = New-PSSession -ComputerName $RemoteLaptop -ErrorAction Stop

        # Ensure destination folder exists on the laptop
        Invoke-Command -Session $Session -ScriptBlock {
            param($Path)
            if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
        } -ArgumentList $RemoteProdPath

        # Resolve the remote user's Documents\PowerShell\Modules\MyTools path and create folders
        $RemoteModulePath = Invoke-Command -Session $Session -ScriptBlock {
            param($ProdPath, $ModName)
            $RemoteBaseModDir = ($env:PSModulePath -split ';' | Where-Object { $_ -notlike "$env:ProgramFiles*" -and $_ -notlike "$env:SystemRoot*" })[0]
            $ModPath          = Join-Path $RemoteBaseModDir $ModName
            if (-not (Test-Path $ProdPath)) { New-Item -ItemType Directory -Path $ProdPath -Force | Out-Null }
            if (-not (Test-Path $ModPath))  { New-Item -ItemType Directory -Path $ModPath -Force | Out-Null }
            return $ModPath
        } -ArgumentList $RemoteProdPath, $ModuleName
        
        # Copy each production script over the PowerShell session
        foreach ($Script in $ScriptsToDeploy) {
            $SourceFile = Join-Path $DevFolder $Script
            if (Test-Path $SourceFile) {
                Copy-Item -Path $SourceFile -Destination $RemoteProdPath -ToSession $Session -Force
                Write-Host "  [$RemoteLaptop] Copied $Script" -ForegroundColor Green
            }
        }

        # Copy the MyTools module files over the session
        if (Test-Path $DevModuleFolder) {
            Copy-Item -Path "$DevModuleFolder\*" -Destination $RemoteModulePath -ToSession $Session -Recurse -Force
            Write-Host "  [$RemoteLaptop] Deployed module: $ModuleName" -ForegroundColor Green
        }

        Remove-PSSession $Session
    }
    catch {
        Write-Warning "Could not connect via PowerShell Remoting to $RemoteLaptop : $($_.Exception.Message)"
    }
}
else {
    Write-Host "  [$RemoteLaptop] Offline or unreachable. Skipped remote deployment." -ForegroundColor Yellow
}
