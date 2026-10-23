# 1. Define the parent directory you want to scan (use "." for the current directory)
$TargetFolder = "."

# 2. Get all subdirectories recursively
#$Subfolders = Get-ChildItem -Path $TargetFolder -Recurse -Directory
$Subfolders = (Get-ChildItem -Path $TargetFolder -Recurse -Directory) + (Get-Item -Path $TargetFolder)


# 3. Loop through each subfolder, count images, and calculate total size
$Result = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($Folder in $Subfolders) {
    # Gather all .jpg and .jpeg files in the current subfolder
    $Images = Get-ChildItem -Path $Folder.FullName -File -ErrorAction SilentlyContinue | 
              Where-Object { $_.Extension -in '.jpg', '.jpeg' }

    if ($Images) {
        $TotalSizeMB = ($Images | Measure-Object -Property Length -Sum).Sum / 1MB
        $JpgCount    = $Images.Count
    } else {
        $TotalSizeMB = 0
        $JpgCount    = 0
    }

    # Add the current folder's data to our list
    $Result.Add([PSCustomObject]@{
        "Folder Name"     = $Folder.Name
        "JPG Count"       = $JpgCount
        "Total Size (MB)" = [Math]::Round($TotalSizeMB, 2)
        "Full Path"       = $Folder.FullName
    })
}

# 4. Calculate Grand Totals across all folders
$GrandTotalCount = ($Result | Measure-Object -Property "JPG Count" -Sum).Sum
$GrandTotalSize  = ($Result | Measure-Object -Property "Total Size (MB)" -Sum).Sum

# 5. Append the Grand Total row to the end of the collection
$Result.Add([PSCustomObject]@{
    "Folder Name"     = "--- GRAND TOTAL ---"
    "JPG Count"       = $GrandTotalCount
    "Total Size (MB)" = [Math]::Round($GrandTotalSize, 2)
    "Full Path"       = ""
})

# 6. Display the results in a clean table
#$Result | Format-Table -AutoSize
$Result | Export-Csv -Path ".\JpgCountReport.csv" -NoTypeInformation
