 #DeployScripts.ps1
#   Deployment Script...
#   Copyright � 1998-2026, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   11/13/26    Created;
#=================================================================================================================================
#Notes to self:
#On failure to connect to AAIUCI4, performed the following on the development computer to add AAIUCI4 to the TrustedHosts list
#(preserving any existing entries):
#   Start-Service WinRM
#   $CurrentHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
#   if ([string]::IsNullOrWhiteSpace($CurrentHosts)) {
#       Set-Item WSMan:\localhost\Client\TrustedHosts -Value "AAIUCI4" -Force
#   } elseif ($CurrentHosts -notlike "*AAIUCI4*") {
#       Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$CurrentHosts,AAIUCI4" -Force
#   }
#
#=================================================================================================================================
Import-Module MyTools

$AppName = "DeployScripts"
$StartTime = Get-Date
$Message = "[$AppName © $("{0:yyyy}" -f $StartTime), Ken Clark                       $("{0:MM/dd/yy} {0:hh:mm:ss tt}" -f $StartTime)]"
Write-Output $Message

$RemoteCred = Import-Clixml -Path "$env:USERPROFILE\msa_deploy.xml" #Get-Credential -UserName "MicrosoftAccount\kfc12@comcast.net" -Message "Enter your Microsoft Account password"

$DevFolder       = "D:\Users\kfc12\Source\repos\kclark12github\Scripts"        # Your Git / VS Code folder
$ProdScriptsPath  = "D:\Program Files (x86)\KClark Software"                      # Local folder where Task Scheduler runs scripts
$CustomModuleName = "MyTools"                         # Your custom helper module name
$DevModuleFolder = Join-Path $DevFolder "Modules\$ModuleName"
$RequiredModules  = @("SqlServer", "Mailozaurr")      # Public modules to ensure exist

# List only the production scripts you want deployed (keeps test scripts out)
$ScriptsToDeploy = @(
    "EnvList.ps1",
    "ExportTasks.ps1",
    "FileList.ps1",
    "HarvestGOGGoodies.ps1",
    "SoftwareInventory.ps1"
)
$RemoteComputers = @("AAIUCI4"
)               # Hostname or IP of your laptop

Write-Host "Updating Local Machine ($env:COMPUTERNAME)" -ForegroundColor Cyan
# Ensure required public modules exist locally
foreach ($Mod in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [Local] Installing missing module: $Mod..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
}
$LocalBaseModDir = ($env:PSModulePath -split ';' | Where-Object { $_ -notlike "$env:ProgramFiles*" -and $_ -notlike "$env:SystemRoot*" })[0]
$LocalCustomModPath = Join-Path $LocalBaseModDir $CustomModuleName

if (-not (Test-Path $ProdScriptsPath))    { New-Item -ItemType Directory -Path $ProdScriptsPath -Force | Out-Null }
if (-not (Test-Path $LocalCustomModPath)) { New-Item -ItemType Directory -Path $LocalCustomModPath -Force | Out-Null }

# Deploy to Local Installation Folder
Write-Host "Deploying to local installation folder ($ProdScriptPath)..." -ForegroundColor Cyan
foreach ($Script in $ScriptsToDeploy) {
    $SourceFile = Join-Path $DevFolder $Script
    if (Test-Path $SourceFile) {
        Copy-Item -Path $SourceFile -Destination $ProdScriptPath -Force
        Write-Host "  [Local] Copied $Script" -ForegroundColor Green
    } else {
        Write-Warning "  [Local] Missing source file: $SourceFile"
    }
}
Write-Host "Deploying to local PowerShell module folder ($LocalCustomModPath)..." -ForegroundColor Cyan
if (Test-Path $DevModuleFolder) {
    Copy-Item -Path "$DevModuleFolder\*" -Destination $LocalCustomModPath -Recurse -Force
    Write-Host "  [Local] Deployed module: $CustomModuleName" -ForegroundColor Green
} else {
    Write-Warning "  [Local] Missing module folder: $DevModuleFolder"
}

# Deploy to Remote Remote Computers (if online)
foreach ($Computer in $RemoteComputers) {
    Write-Host "`nChecking Remote Computer: $Computer" -ForegroundColor Cyan

    if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
        Write-Host "  [$Computer] Offline or unreachable. Skipping." -ForegroundColor Yellow
        continue
    }

    try {
        $Session = New-PSSession -ComputerName $Computer -Credential $RemoteCred -ErrorAction Stop

        # Step A: Bootstrap PowerShell 7, Public Modules, Env Vars, and Resolve Target Paths
        $RemoteModDestination = Invoke-Command -Session $Session -ScriptBlock {
            param($ProdPath, $CustomMod, $PubModules, $SetEnv, $User, $Pass)

            #Ensure PowerShell 7 is installed
            $PwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            $PwshExe = if ($PwshCmd) { $PwshCmd.Source } else { $null }
            #If still not found, check common fallback directories or install via winget
            if (-not $PwshExe) {
                $FallbackPaths = @(
                    "$env:ProgramFiles\PowerShell\7\pwsh.exe",
                    "$env:LOCALAPPDATA\Microsoft\WindowsApps\pwsh.exe"
                )
                $PwshExe = $FallbackPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
            }
            if (-not $PwshExe) {
                Write-Output "Installing PowerShell 7 on $env:COMPUTERNAME..."
                winget install --id Microsoft.PowerShell --source winget --silent --accept-package-agreements --accept-source-agreements | Out-Null
                $PwshExe = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source
            }
            if (-not (Test-Path $PwshExe)) {
                Write-Output "Installing PowerShell 7 on $env:COMPUTERNAME..."
                winget install --id Microsoft.PowerShell --source winget --silent --accept-package-agreements --accept-source-agreements | Out-Null
            }

            #Inside PowerShell 7: install missing public modules & resolve PS 7 user PSModulePath
            $ResolvedModPath = & $PwshExe -NoProfile -Command {
                param($Mods, $MyModName)
                foreach ($m in $Mods) {
                    if (-not (Get-Module -ListAvailable -Name $m)) {
                        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
                        Install-Module -Name $m -Scope AllUsers -Force -AllowClobber
                    }
                }
                $BaseDir = ($env:PSModulePath -split ';' | Where-Object { 
                    $_ -notlike "$env:ProgramFiles*" -and $_ -notlike "$env:SystemRoot*" -and -not [string]::IsNullOrWhiteSpace($_)
                })[0]
                return (Join-Path $BaseDir $MyModName)
            } -Args $PubModules, $CustomMod

            #Create destination folders on the remote machine
            if (-not (Test-Path $ProdPath))        { New-Item -ItemType Directory -Path $ProdPath -Force | Out-Null }
            if (-not (Test-Path $ResolvedModPath)) { New-Item -ItemType Directory -Path $ResolvedModPath -Force | Out-Null }

            return $ResolvedModPath
        } -ArgumentList $ProdScriptsPath, $CustomModuleName, $RequiredModules, $PushEnvVars

        # Step B: Copy production scripts over the session
        foreach ($Script in $ScriptsToDeploy) {
            $SourceFile = Join-Path $DevFolder $Script
            if (Test-Path $SourceFile) {
                Copy-Item -Path $SourceFile -Destination $ProdScriptsPath -ToSession $Session -Force
                Write-Host "  [$Computer] Deployed script: $Script" -ForegroundColor Green
            }
        }

        # Step C: Copy custom MyTools module files over the session
        if (Test-Path $DevModuleFolder) {
            Copy-Item -Path "$DevModuleFolder\*" -Destination $RemoteModDestination -ToSession $Session -Recurse -Force
            Write-Host "  [$Computer] Deployed custom module: $CustomModuleName -> $RemoteModDestination" -ForegroundColor Green
        }

        Remove-PSSession $Session
    }
    catch {
        Write-Warning "  [$Computer] Deployment failed: $($_.Exception.Message)"
    }
}
Write-Host "`nDeployment complete!" -ForegroundColor Cyan
