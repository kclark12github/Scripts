$DevFolder      = "D:\Users\kfc12\Source\repos\kclark12github\Scripts"        # Your Git / VS Code folder
$LocalProdPath  = "D:\Program Files (x86)\KClark Software"                      # Local folder where Task Scheduler runs scripts
$RemoteLaptop   = "AAIUCI4"                # Hostname or IP of your laptop
$RemoteProdPath = "D:\Program Files (x86)\KClark Software"                      # Destination folder on your laptop

# 2. List only the production scripts you want deployed (keeps test scripts out)
$ScriptsToDeploy = @(
    "eMailResults.ps1",
    "EnvList.ps1",
    "ExportTasks.ps1",
    "FileList.ps1",
    "HarvestGOGGoodies.ps1",
    "SoftwareInventory.ps1"
)

# 3. Deploy to Local Installation Folder
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

# 4. Deploy to Remote Laptop (if online)
Write-Host "`nChecking if $RemoteLaptop is reachable..." -ForegroundColor Cyan
if (Test-Connection -ComputerName $RemoteLaptop -Count 1 -Quiet) {
    try {
        $Session = New-PSSession -ComputerName $RemoteLaptop -ErrorAction Stop

        # Ensure destination folder exists on the laptop
        Invoke-Command -Session $Session -ScriptBlock {
            param($Path)
            if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
        } -ArgumentList $RemoteProdPath

        # Copy each production script over the PowerShell session
        foreach ($Script in $ScriptsToDeploy) {
            $SourceFile = Join-Path $DevFolder $Script
            if (Test-Path $SourceFile) {
                Copy-Item -Path $SourceFile -Destination $RemoteProdPath -ToSession $Session -Force
                Write-Host "  [$RemoteLaptop] Copied $Script" -ForegroundColor Green
            }
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
