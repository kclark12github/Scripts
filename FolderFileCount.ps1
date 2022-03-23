$colItems = Get-ChildItem $startFolder | Where-Object {$_.PSIsContainer -eq $true} | Sort-Object
foreach ($i in $colItems) {

  $subFolderItems = (Get-ChildItem $i.FullName -recurse -force | Where-Object {-not ($_.PSIsContainer)}).Count
  $i.Name +"|" + $subFolderItems
}