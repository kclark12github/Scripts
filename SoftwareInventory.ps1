#SoftwareInventory.ps1
#   Build Utility Script...
#   Copyright © 1998-2026, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   11/12/26    Moved UserName and Password into environment variables & added Inventoried date;
#               Reorganized so registry information is only gathered once;
#               Upgraded to PowerShell 7.7.0;
#   11/11/26    Enhanced to store installation data in SQL database;
#               Added History;
#=================================================================================================================================
#Notes to Self:
#  This script is scheduled to run in the Task Scheduler.
#=================================================================================================================================
# 
param([string]$Root="$($env:COMPUTERNAME)",[string]$BackupFolder,[string]$LogPath)
# Note:
#  * Accepts input only via the pipeline, either line by line, 
#    or as a single, multi-line string.
#  * The input is assumed to have a header line whose column names
#    mark the start of each field
#    * Column names are assumed to be *single words* (must not contain spaces).
#  * The header line is assumed to be followed by a separator line
#    (its format doesn't matter).
Import-Module MyTools
function ConvertFrom-FixedColumnTable {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline)] [string] $InputObject
  )
  
  begin {
    Set-StrictMode -Version 1
    $lineNdx = 0
  }
  
  process {
    $lines = 
      if ($InputObject.Contains("`n")) { $InputObject.TrimEnd("`r", "`n") -split '\r?\n' }
      else { $InputObject }
    foreach ($line in $lines) {
      ++$lineNdx
      if ($lineNdx -eq 1) { 
        # header line
        $headerLine = $line 
      }
      elseif ($lineNdx -eq 2) { 
        # separator line
        # Get the indices where the fields start.
        $fieldStartIndices = [regex]::Matches($headerLine, '\b\S').Index
        # Calculate the field lengths.
        $fieldLengths = foreach ($i in 1..($fieldStartIndices.Count-1)) { 
          $fieldStartIndices[$i] - $fieldStartIndices[$i - 1] - 1
        }
        # Get the column names
        $colNames = foreach ($i in 0..($fieldStartIndices.Count-1)) {
          if ($i -eq $fieldStartIndices.Count-1) {
            $headerLine.Substring($fieldStartIndices[$i]).Trim()
          } else {
            $headerLine.Substring($fieldStartIndices[$i], $fieldLengths[$i]).Trim()
          }
        } 
      }
      else {
        # data line
        $oht = [ordered] @{} # ordered helper hashtable for object constructions.
        $i = 0
        foreach ($colName in $colNames) {
          $oht[$colName] = 
            if ($fieldStartIndices[$i] -lt $line.Length) {
              if ($fieldLengths[$i] -and $fieldStartIndices[$i] + $fieldLengths[$i] -le $line.Length) {
                $line.Substring($fieldStartIndices[$i], $fieldLengths[$i]).Trim()
              }
              else {
                $line.Substring($fieldStartIndices[$i]).Trim()
              }
            }
          ++$i
        }
        # Convert the helper hashable to an object and output it.
        [pscustomobject] $oht
      }
    }
  }
  
}
function DoExcel {
    #==============================================================================================================================
    $Message = "`nCreating new Excel file...";   Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br />"
    $excel = New-Object -ComObject excel.application 
    $excel.visible = $false
    $workbook = $excel.Workbooks.Add()

    #==============================================================================================================================
    $EmailBody += Write-To-Excel -excel $excel -workbook $workbook -Infos $RegKey32Infos -Tab "Uninstall"   -BackColor 14 -FontColor 2 -LogPath $LogPath  #Teal/White
    $EmailBody += Write-To-Excel -excel $excel -workbook $workbook -Infos $RegKey64Infos -Tab "Wow6432Node" -BackColor 10 -FontColor 6 -LogPath $LogPath  #Dark Green/Yellow
    $EmailBody += Write-To-Excel -excel $excel -workbook $workbook -Infos $WinGetInfos   -Tab "winget"      -BackColor 11 -FontColor 2 -LogPath $LogPath  #Dark Blue/White
    $EmailBody += "<br />"

    $fileName = Split-Path -Path $xlsFile -Leaf
    $tempXlsFile = Join-Path -Path $env:TEMP -ChildPath $fileName

    $excel.DisplayAlerts = $false
    $workbook.SaveAs($tempXlsFile,51); $workbook.Saved = $true; $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    $targetFolder = Split-Path -Path $xlsFile -Parent
    if (-not (Test-Path -Path $targetFolder)) {New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null}
    Move-Item -Path $tempXlsFile -Destination $xlsFile -Force

    #Need to clean-up BackupFolder or these files will eventually get out of hand (filling the disk)...
    $Threshold = (Get-Date).AddMonths(-1)
    $Pattern = "$($BackupFolder)SI.$Root.*.xlsx"
    $CleanList = Get-ChildItem $Pattern -File | Where-Object { ($_.LastWriteTime –ge $Threshold)}
    $PurgeList = Get-ChildItem $Pattern -File | Where-Object { ($_.LastWriteTime –lt $Threshold)}
    if ($CleanList.Count -gt 1 -and $PurgeList.Count -gt 4) {
        $Message = "`nPurging $($PurgeList.Count) old files...";        Write-Message -Message $Message -Path $LogPath;        $EmailBody += "$Message<br />"
        $PurgeList | ForEach-Object {
            Write-Message -Message $_.Name -Path $LogPath; $EmailBody += "$($_.Name)<br />"
        }
        Get-ChildItem $Pattern -File | Where-Object { ($_.LastWriteTime –lt $Threshold)} | Remove-Item -ErrorAction SilentlyContinue -Force
        $EmailBody += "<br />"
    }
}
function DoSQL {
    $ServerName = "7X4L243"; $DatabaseName = "TreasureChest2"; $TableName = "Apps"
    $SQLUser = $env:TC2User; $SQLPassword = ConvertTo-SecureString $env:TC2Password -AsPlainText -Force; $SQLPassword.MakeReadOnly()
    $SQLcredentials = New-Object System.Management.Automation.PSCredential ($SQLUser, $SQLPassword)
    $DeleteQuery = "DELETE FROM [$TableName] WHERE [System] = '$Root';"
    Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Credential $SQLcredentials -Query $DeleteQuery -TrustServerCertificate

    #==============================================================================================================================
    #$Message = "`nSaving installation data in $ServerName.$DatabaseName.$TableName...";   Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br />"
    $wteMessage = "`nSaving data to $ServerName.$DatabaseName.$TableName..."
    Write-Message -Message $wteMessage -Path $LogPath
    $EmailBody += Write-To-SQL -Infos $RegKey32Infos -LogPath $LogPath
    $EmailBody += Write-To-SQL -Infos $RegKey64Infos -LogPath $LogPath
    $EmailBody += Write-To-SQL -Infos $WinGetInfos   -LogPath $LogPath
    $EmailBody += "<br />"
 }
function GetInstalledAppsInfo {
    Param($RegKey, $Tag)
    #$wteMessage = "Gathering installation information from registry ($Tag)..."; Write-Message -Message $wteMessage -Path $LogPath
    $Objs = @()
    $InstalledAppsInfos = Get-ItemProperty -Path $RegKey
    Foreach($InstalledAppsInfo in $InstalledAppsInfos)
    {
        $Obj = [PSCustomObject]@{System=$Root;
                                    ApplicationName = $InstalledAppsInfo.DisplayName;
                                    Version = $InstalledAppsInfo.DisplayVersion;
                                    Publisher = $InstalledAppsInfo.Publisher;
                                    InstallDate = ConvertTo-Date($InstalledAppsInfo.InstallDate);
                                    EstimatedSize = ConvertTo-SizeString($InstalledAppsInfo.EstimatedSize);
                                    InstallLocation = $InstalledAppsInfo.InstallLocation;
                                    UninstallString = $InstalledAppsInfo.UninstallString;
                                    Source = $InstalledAppsInfo.PSPath;
                                    RegKey = $Tag;
                                    Inventoried = (Get-Date).Date;
                                    }
        $Objs += $Obj
    }
    $Objs
}
function GetWinGetInfo {
    #$wteMessage = "Gathering installation information from registry (winget)..."; Write-Message -Message $wteMessage -Path $LogPath
    $Objs = @()
    #$OutputEncoding = [System.Text.Encoding]::UTF8
    try {[Console]::OutputEncoding = [System.Text.Encoding]::UTF8} catch {}   # Ignore invalid handle errors in ISE or background Task Scheduler sessions
    $wingetObjs = (winget list) -match '^(\p{L}|-)' |               # filter out progress-display lines
                                ConvertFrom-FixedColumnTable |  # parse output into objects
                                Sort-Object Name |              # sort by the Name property (column)
        ForEach-Object {
            $Name = $_.Name
            $Version = $_.Version
            $Source = $_.Source
            $Publisher = $null
            $InstallDate = $null
            $EstimatedSize = $null
            $InstallLocation = $null
            $ID = $_.ID
            if ($Source -eq "winget") {
                #Parse winget show --id $_.ID to get Pulisher and other desired data
                $show = (winget show --id $ID) -split "`n"
                foreach($item in $show) {
                    if ($item.StartsWith("Found ")) {
                        $Name = $item.Substring("Found ".Length).Replace("[$ID]","")
                    } elseif ($item.StartsWith("Publisher: ")) {
                        $Publisher = $item.Substring("Publisher: ".Length)
                    }
                }
            }
            else {
                $ErrorActionPreference = "Stop"
                Try {
                    #First look in the 32-bit Uninstall Registry Key...
                    $RegKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ID"
                    $reg = Get-ItemProperty -Path $RegKey
                    $Source = "Uninstall"
                    $Name = $reg.DisplayName
                    $Publisher = $reg.Publisher
                    $InstallDate = ConvertTo-Date($reg.InstallDate)
                    $EstimatedSize = ConvertTo-SizeString($reg.EstimatedSize)
                    $InstallLocation = $reg.InstallLocation
                } Catch [System.Management.Automation.ItemNotFoundException]{
                    Try {
                        #Next look in the 64-bit Uninstall Registry Key...
                        $RegKey = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ID"
                        $reg = Get-ItemProperty -Path $RegKey
                        $Source = "Wow5432Node"
                        $Name = $reg.DisplayName
                        $Publisher = $reg.Publisher
                        $InstallDate = ConvertTo-Date($reg.InstallDate)
                        $EstimatedSize = ConvertTo-SizeString($reg.EstimatedSize)
                        $InstallLocation = $reg.InstallLocation
                    } Catch [System.Management.Automation.ItemNotFoundException]{
                        #Where Next?
                        $Source = $null
                    }
                } Finally { 
                    $ErrorActionPreference = "Continue"
                    $null = $reg
                    $null = $RegKey
                }
            }
            $Obj = [PSCustomObject]@{System=$Root;
                                    ApplicationName = $Name;
                                    Version = $Version;
                                    Publisher = $Publisher;
                                    InstallDate = ConvertTo-Date($InstalledAppsInfo.InstallDate);
                                    EstimatedSize = ConvertTo-SizeString($InstalledAppsInfo.EstimatedSize);
                                    InstallLocation = $InstallLocation;
                                    UninstallString = $null;
                                    Source = $Source;
                                    RegKey = "winget";
                                    Inventoried = (Get-Date).Date;
                                    }
            $Objs += $Obj
        }
    $Objs
}
function Write-To-Excel {
    Param($excel, $workbook, $Infos, $Tab, $BackColor, $FontColor=2, $LogPath)
    $errMessage = ""
    $Range = "A1:H1"
    "<br />" #Write-Message is causing messages to be returned as the return value for this function, so inject line breaks where necessary

    #Save $objs to Excel $Tab tab...
    $wteMessage = "Writing data to Excel($Tab)..."; Write-Message -Message $wteMessage -Path $LogPath
    "<br />" #Write-Message is causing messages to be returned as the return value for this function, so inject line breaks where necessary

    if ($workbook.Worksheets.Item(1).Name -eq "Sheet1") {
        $ws = $workbook.Worksheets.Item(1) #Use the new empty sheet
    } else {
        $ws = $workbook.Worksheets.add()
    }
    $ws.Name = $Tab
    $ws.Tab.ColorIndex = $BackColor
    #Format column headers
    $excel.Range($Range).Font.bold = $true
    $excel.Range($Range).Interior.ColorIndex = $BackColor
    $excel.Range($Range).Interior.Pattern = 1
    $excel.Range($Range).Font.ColorIndex = $FontColor
    $excel.ActiveWindow.SplitColumn = 1
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    $iName=1;        $ws.Cells(1, $iName).Value = "Application Name";	$ws.Columns($iName).ColumnWidth = 40 
    $iVer=2;         $ws.Cells(1, $iVer).Value = "Version";				$ws.Columns($iVer).ColumnWidth = 15
    $iPub=3;         $ws.Cells(1, $iPub).Value = "Publisher";			$ws.Columns($iPub).ColumnWidth = 15 
    $iDate=4;        $ws.Cells(1, $iDate).Value = "Install Date";		$ws.Columns($iDate).ColumnWidth = 15 
    $iSize=5;        $ws.Cells(1, $iSize).Value = "Estimated Size";		$ws.Columns($iSize).ColumnWidth = 15 
    $iLoc=6;         $ws.Cells(1, $iLoc).Value = "Install Location";	$ws.Columns($iLoc).ColumnWidth = 40 
    $iUnins=7;       $ws.Cells(1, $iUnins).Value = "Uninstall String";	$ws.Columns($iUnins).ColumnWidth = 40 
    $iSource=8;      $ws.Cells(1, $iSource).Value = "Source";	        $ws.Columns($iSource).ColumnWidth = 40 
    $iInventoried=9; $ws.Cells(1, $iInventoried).Value = "Inventoried";	$ws.Columns($iInventoried).ColumnWidth = 15 

    $Row = 2

    Foreach($Info in $Infos) {
        $ws.Cells.Item($Row, $iName) = $Info.ApplicationName
        $ws.Cells.Item($Row, $iVer) = $Info.Version
        $ws.Cells.Item($Row, $iPub) = $Info.Publisher
        $ws.Cells.Item($Row, $iDate) = $Info.InstallDate
        $ws.Cells.Item($Row, $iSize) = $Info.EstimatedSize
        $ws.Cells.Item($Row, $iLoc) = $Info.InstallLocation
        $ws.Cells.Item($Row, $iUnins) = $Info.UninstallString
        $ws.Cells.Item($Row, $iSource) = $Info.Source
        $ws.Cells.Item($Row, $iInventoried) = $Info.Inventoried
        $Row += 1
    }
    $ws.UsedRange.EntireColumn.AutoFilter() | Out-Null
    $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

    if (-not $errMessage.Equals("")) {
        Write-Message -Message $errMessage -Path $LogPath
        "<br />" #Write-Message is causing messages to be returned as the return value for this function, so inject line breaks where necessary
    }

    #==============================================================================================================================
    #Get-AppxPackage does not provide any significant improvement over these other methods, but preserving for posterity...
    #   - Seems to work off HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages
    #   - winget lists these guys, but doesn't provide as much data as available for Source=winget items
    #
    #$Objs = @()
    #Get-AppxPackage | Get-AppxPackageManifest | 
    #    ForEach-Object{
    #    
    #        $Obj = [PSCustomObject]@{
    #            Computer=$ComputerName;
    #            DisplayName = $_.Package.Applications.Application.VisualElements.DisplayName;
    #            DisplayVersion = $_.Package.Applications.Application.VisualElements.DisplayVersion;
    #            Publisher = $_.Package.Applications.Application.VisualElements.Publisher
    #            }
    #        $Objs += $Obj
    #    }
    #    $Objs | Where-Object { $_.DisplayName } 

    #==============================================================================================================================
    #wmic does not provide any significant improvement over these other methods, but preserving for posterity...
    #wmic product get Name,Vendor,Version,Description,InstallDate,InstallLocation
    #==============================================================================================================================
}
function Write-To-SQL {
    Param($Infos, $LogPath)
    $Infos | Write-SqlTableData -ServerInstance $ServerName -DatabaseName $DatabaseName -Credential $SQLCredentials -SchemaName "dbo" -TableName $TableName -TrustServerCertificate
}
.{
    $AppName = "SoftwareInventory"
    $StartTime = Get-Date
    $Message = "[$AppName © $("{0:yyyy}" -f $StartTime), Ken Clark                       $("{0:MM/dd/yy} {0:hh:mm:ss tt}" -f $StartTime)]"

    if ($BackupFolder.Equals("")) {$BackupFolder = "$($Env:OneDrive)\Backups\$Root\"}  #$($Env:BackupRoot)\
    if ($LogPath.Equals("")) {$LogPath = "$BackupFolder$Root.$AppName.log"}

    Write-Message -Message $Message -Path $LogPath;     $EmailBody = $Message.Replace("©", "&copy;") + "<br />"
    $Message = "Root: $Root";                    Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br />"
    $Message = "BackupFolder: $BackupFolder";    Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br />"
    $Message = "LogPath: $LogPath";              Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br />"
    $now = Get-Date -Format "yyyyMMdd.HHmmss"
    $xlsFile = "$($BackupFolder)SI.$Root.$now.xlsx";    
    $Message = "xlsFile: $xlsFile";              Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br /><br />"
    
    $RegKey32Infos = GetInstalledAppsInfo -RegKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -Tag "Uninstall"
    $RegKey64Infos = GetInstalledAppsInfo -RegKey "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -Tag "Wow6432Node"
    $WinGetInfos = GetWinGetInfo

    DoExcel
    DoSQL
 
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"; 
    Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br /><br />"
    Write-Separator -Path $LogPath
        
    Send-EmailResults  -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
}
