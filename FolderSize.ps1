$colItems = Get-ChildItem $startFolder | Where-Object {$_.PSIsContainer -eq $true} | Sort-Object
foreach ($i in $colItems)
{
    $subFolderItems = Get-ChildItem -LiteralPath $i.FullName -recurse -force | Where-Object {$_.PSIsContainer -eq $false} | Measure-Object -property Length -sum | Select-Object Sum
    $i.Name + "|" + "{0:N2}" -f ($subFolderItems.sum / 1MB) + " MB"
}