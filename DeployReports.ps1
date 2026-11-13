# Define source and destination folders
$SourceFolder = "D:\Users\kfc12\Source\repos\kclark12github\TC2-Reports\TC2-Reports"
$DestinationFolder = "D:\Users\kfc12\Source\repos\kclark12github\TreasureChest2\Reports"

# Ensure the destination folder exists
if (-not (Test-Path -Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
}

# Get all .rdl files from the source folder
$RdlFiles = Get-ChildItem -Path $SourceFolder -Filter "*.rdl"

# Loop through each file, copy, and rename it to .rdlc in the destination
foreach ($File in $RdlFiles) {
    # Construct the new file name with .rdlc extension
    $NewName = $File.BaseName + ".rdlc"
    $DestinationPath = Join-Path -Path $DestinationFolder -ChildPath $NewName
    
    # Copy the file to the new destination
    Copy-Item -Path $File.FullName -Destination $DestinationPath -Force
    
    Write-Host "Deployed: $($File.Name) -> $NewName"
}
