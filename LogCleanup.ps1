# Define the root folder where you want to start searching
$RootFolder = "B:\Databases"

# Define the path where the master text log file will be created (Saves to Desktop)
$LogFilePath = "B:\Databases\LogCleanup.log"

# Initialize or overwrite the log file with a clean header
$Header = @"
===================================================================
SQL SERVER LOG CLEANUP MASTER REPORT
Executed on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Root Directory: $RootFolder
===================================================================
"@
$Header | Out-File -FilePath $LogFilePath -Encoding utf8

# Get all subfolders (including the root folder itself)
$Folders = Get-ChildItem -Path $RootFolder -Directory -Recurse

foreach ($Folder in $Folders) {
    # 1. Find the earliest (oldest) .bak file in this specific folder
    $EarliestBak = Get-ChildItem -Path $Folder.FullName -Filter "*.bak" -File | 
                   Sort-Object $LastWriteTime | 
                   Select-Object -First 1

    # If a .bak file exists in this folder, proceed
    if ($EarliestBak) {
        $BakCutoffTime = $EarliestBak.LastWriteTime
        
        # Log folder details to console and file text buffer
        $FolderSummary = "`r`nFolder: $($Folder.FullName)`r`n  -> Oldest .bak file: $($EarliestBak.Name) (Modified: $($BakCutoffTime))"
        Write-Host "Folder: $($Folder.FullName)" -ForegroundColor Cyan
        Write-Host "  -> Oldest .bak file: \$($EarliestBak.Name) ($BakCutoffTime)" -ForegroundColor Gray
        $FolderSummary | Out-File -FilePath $LogFilePath -Append -Encoding utf8

        # 2. Find all .log files in this specific folder older than the cutoff time
        $LogsToDelete = Get-ChildItem -Path $Folder.FullName -Filter "*.log" -File |
                        Where-Object { $_.LastWriteTime -lt $BakCutoffTime }

        if ($LogsToDelete.Count -eq 0) {
            "  -> No matching .log files found to delete." | Out-File -FilePath $LogFilePath -Append -Encoding utf8
        }

        # 3. Delete and log the matching .log files
        foreach ($Log in $($LogsToDelete)) {$DeleteMessage = "  [DELETING] $($Log.Name) (Modified: $($Log.LastWriteTime))"
            Write-Host $DeleteMessage -ForegroundColor Yellow
            $DeleteMessage | Out-File -FilePath $LogFilePath -Append -Encoding utf8
            
            # UNCOMMENT the line below to actually delete files once you verify the output text
            Remove-Item -Path $Log.FullName -Force
        }
    }
}

# Final execution summary stamp
$Footer = "`r`n===================================================================`r`nCleanup execution completed safely.`r`n==================================================================="
$Footer | Out-File -FilePath $LogFilePath -Append -Encoding utf8

Write-Host "`r`nMaster cleanup log successfully saved to: $LogFilePath" -ForegroundColor Green
