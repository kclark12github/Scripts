param([string]$Root="$($env:COMPUTERNAME)",[string]$BackupFolder,[string]$LogPath)
# Note:
#  * Accepts input only via the pipeline, either line by line, 
#    or as a single, multi-line string.
#  * The input is assumed to have a header line whose column names
#    mark the start of each field
#    * Column names are assumed to be *single words* (must not contain spaces).
#  * The header line is assumed to be followed by a separator line
#    (its format doesn't matter).
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
function Format-Elapsed {
    Param($Start, $End)
    $Elapsed = ""
    $ts = New-TimeSpan -start $Start -end $End
    if ($ts.Days -gt 0) {$Elapsed += "$($ts.Days) Days, "}
    if ($ts.Hours -gt 0) {$Elapsed += "$($ts.Hours) Hours, "}
    if ($ts.Minutes -gt 0) {$Elapsed += "$($ts.Minutes) Minutes, "}
    $Elapsed += "{0}.{1:000} Seconds" -f $ts.Seconds, $ts.Milliseconds
    return $Elapsed
}
function Send-Mail {
    Param($AppName, $EmailBody, $LogPath)
    $PW = ConvertTo-SecureString $env:SMTP_PW -AsPlainText -Force
    $Creds = New-Object System.Management.Automation.PSCredential ($env:SMTP_USER, $PW)
    $Server = $env:SMTP_ADDRESS+":"+$env:SMTP_PORT
    Send-MailMessage -From "$env:USERNAME <$env:My_EMAIL>" -To "$env:USERNAME <$env:My_EMAIL>" -Subject "$AppName Succeeded!" -Body $EmailBody -BodyAsHtml -Attachments $LogPath -SmtpServer $env:SMTP_ADDRESS -Port $env:SMTP_PORT -Credential $Creds
}
function Write-Log {
    Param($Message, $Path = ".")
    function TS {return "[{0:MM/dd/yy} {0:HH:mm:ss tt}]" -f (Get-Date)}
    Write-Message -Message "$(TS) $Message" -Path $Path
}
function Write-Message {
    Param($Message, $Path = ".")
    "$Message" | Tee-Object -FilePath $Path -Append | Write-Output
}
function Write-Separator {
    Param($Path = ".")
    #                      1         2         3         4         5         6         7         8
    #             12345678901234567890123456789012345678901234567890123456789012345678901234567890
    $Separator = "--------------------------------------------------------------------------------"
    Write-Message -Message $Separator -Path $Path
}
function Write-Files {
    Param($Path = ".", $OutFile, $LogPath)
    $errMessage = ""
    $Message = "Listing Files from $Path to $OutFile"
    Write-Message -Message $Message -Path $LogPath

    Get-ChildItem -Path "$Path" -Force -Recurse -ErrorVariable errMessage | Select Fullname,Length,CreationTime,LastWriteTime | Export-Csv -Path "$OutFile" -NoTypeInformation
    if (-not $errMessage.Equals("")) {Write-Message -Message $errMessage -Path $LogPath}
    "<br />" #Return break tags appended to message
}
.{
    $AppName = "SoftwareInventory"
    $StartTime = Get-Date
    $Message = "[$AppName © $("{0:yyyy}" -f $StartTime), Ken Clark                       $("{0:MM/dd/yy} {0:hh:mm:ss tt}" -f $StartTime)]"

    if ($BackupFolder.Equals("")) {$BackupFolder = "$($Env:OneDrive)\Backups\"}  #$($Env:BackupRoot)\
    if ($LogPath.Equals("")) {$LogPath = "$($BackupFolder)$Root.$AppName.log"}

    Write-Message -Message $Message -Path $LogPath
    $EmailBody = $Message.Replace("©", "&copy;") + "<br />"
    $Message = "Root: $Root; "
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"
    $Message = "BackupFolder: $BackupFolder; "
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"
    $Message = "LogPath: $LogPath; "
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"
    $now = Get-Date -Format "yyyyMMdd.HHmmss"
    $xlsFile = "$($BackupFolder)$Root\SI.$Root.$now.xlsx"
    $Message = "xlsFile: $xlsFile;"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br /><br />"
    
    #==============================================================================================================================
    $Message = "Creating new Excel file..."
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"
    $excel = New-Object -ComObject excel.application 
    $excel.visible = $false
    $workbook = $excel.Workbooks.Add()
    #$workbook.Worksheets.Item(3).Delete()
    #$Range = $excel.Worksheets	#Range("A1","G5")
    #$Range.Font.Size = 14

    #==============================================================================================================================
    $RegKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $Message = "Gathering installation information from registry ($RegKey)..."
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"

    $Objs = @()
    $InstalledAppsInfos = Get-ItemProperty -Path $RegKey
    Foreach($InstalledAppsInfo in $InstalledAppsInfos)
    {
        $Obj = [PSCustomObject]@{Computer=$env:ComputerName;
                                 DisplayName = $InstalledAppsInfo.DisplayName;
                                 Version = $InstalledAppsInfo.DisplayVersion;
                                 Publisher = $InstalledAppsInfo.Publisher;
                                 InstallDate = $InstalledAppsInfo.InstallDate;
                                 EstimatedSize = $InstalledAppsInfo.EstimatedSize;
                                 InstallLocation = $InstalledAppsInfo.InstallLocation;
                                 UninstallString = $InstalledAppsInfo.UninstallString;
                                 Source = $InstalledAppsInfo.PSPath;
                                 }
        $Objs += $Obj
    }
    #Save $objs to Excel Uninstall tab...
    $Message = "Writing data to Excel(Uninstall)..."
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"
    $ws = $workbook.Worksheets.Item(1) 
    $ws.Name = 'Uninstall'
    #Format cells
    $excel.Range("A1:H1").Font.bold = $true
    $excel.Range("A1:H1").Interior.ColorIndex = 14   #Teal
    $excel.Range("A1:H1").Interior.Pattern = 1
    $excel.Range("A1:H1").Font.ColorIndex = 2        #White
    $excel.ActiveWindow.SplitColumn = 1
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
    $iName=1;    $ws.Cells(1, $iName).Value = "Application Name";	$ws.Columns($iName).ColumnWidth = 40 
    $iVer=2;     $ws.Cells(1, $iVer).Value = "Version";				$ws.Columns($iVer).ColumnWidth = 15
    $iPub=3;     $ws.Cells(1, $iPub).Value = "Publisher";			$ws.Columns($iPub).ColumnWidth = 15 
    $iDate=4;    $ws.Cells(1, $iDate).Value = "Install Date";		$ws.Columns($iDate).ColumnWidth = 15 
    $iSize=5;    $ws.Cells(1, $iSize).Value = "Estimated Size";		$ws.Columns($iSize).ColumnWidth = 15 
    $iLoc=6;     $ws.Cells(1, $iLoc).Value = "Install Location";	$ws.Columns($iLoc).ColumnWidth = 40 
    $iUnins=7;   $ws.Cells(1, $iUnins).Value = "Uninstall String";	$ws.Columns($iUnins).ColumnWidth = 40 
    $iSource=8;  $ws.Cells(1, $iSource).Value = "Source";	        $ws.Columns($iSource).ColumnWidth = 40 
    $Row = 2
    $Objs | Where-Object { $_.DisplayName } | Sort-Object -Property DisplayName |
        ForEach-Object {
            $ws.Cells.Item($Row, $iName) = $_.DisplayName
            $ws.Cells.Item($Row, $iVer) = $_.Version
            $ws.Cells.Item($Row, $iPub) = $_.Publisher
            $ws.Cells.Item($Row, $iDate) = $_.InstallDate
            $ws.Cells.Item($Row, $iSize) = $_.EstimatedSize
            $ws.Cells.Item($Row, $iLoc) = $_.InstallLocation
            $ws.Cells.Item($Row, $iUnins) = $_.UninstallString
            $ws.Cells.Item($Row, $iSource) = $_.Source
            $Row += 1
        }
    $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

    #==============================================================================================================================
    $RegKey = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $Message = "Gathering installation information from registry ($RegKey)..."
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"

    $Objs = @()
    $InstalledAppsInfos = Get-ItemProperty -Path $RegKey
    Foreach($InstalledAppsInfo in $InstalledAppsInfos)
    {
        $Obj = [PSCustomObject]@{Computer=$env:ComputerName;
                                 DisplayName = $InstalledAppsInfo.DisplayName;
                                 Version = $InstalledAppsInfo.DisplayVersion;
                                 Publisher = $InstalledAppsInfo.Publisher;
                                 InstallDate = $InstalledAppsInfo.InstallDate;
                                 EstimatedSize = $InstalledAppsInfo.EstimatedSize;
                                 InstallLocation = $InstalledAppsInfo.InstallLocation;
                                 UninstallString = $InstalledAppsInfo.UninstallString;
                                 Source = $InstalledAppsInfo.PSPath;
                                 }
        $Objs += $Obj
    }
    #Save $objs to Excel Wow6432Node tab...
    $Message = "Writing data to Excel(Wow6432Node)..."
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"
    $ws = $workbook.Worksheets.add()
    $ws.Name = 'Wow6432Node'
    $ws.Activate()
    #Format cells
    $excel.Range("A1:H1").Font.bold = $true
    $excel.Range("A1:H1").Interior.ColorIndex = 10    #Dark Green
    $excel.Range("A1:H1").Interior.Pattern = 1
    $excel.Range("A1:H1").Font.ColorIndex = 6         #Yellow
    $excel.ActiveWindow.SplitColumn = 1
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
    $iName=1;    $ws.Cells(1, $iName).Value = "Application Name";	$ws.Columns($iName).ColumnWidth = 40 
    $iVer=2;     $ws.Cells(1, $iVer).Value = "Version";				$ws.Columns($iVer).ColumnWidth = 15
    $iPub=3;     $ws.Cells(1, $iPub).Value = "Publisher";			$ws.Columns($iPub).ColumnWidth = 15 
    $iDate=4;    $ws.Cells(1, $iDate).Value = "Install Date";		$ws.Columns($iDate).ColumnWidth = 15 
    $iSize=5;    $ws.Cells(1, $iSize).Value = "Estimated Size";		$ws.Columns($iSize).ColumnWidth = 15 
    $iLoc=6;     $ws.Cells(1, $iLoc).Value = "Install Location";	$ws.Columns($iLoc).ColumnWidth = 40 
    $iUnins=7;   $ws.Cells(1, $iUnins).Value = "Uninstall String";	$ws.Columns($iUnins).ColumnWidth = 40 
    $iSource=8;  $ws.Cells(1, $iSource).Value = "Source";	        $ws.Columns($iSource).ColumnWidth = 40 
    $Row = 2
    $Objs | Where-Object { $_.DisplayName } | Sort-Object -Property DisplayName |
        ForEach-Object {
            $ws.Cells.Item($Row, $iName) = $_.DisplayName
            $ws.Cells.Item($Row, $iVer) = $_.Version
            $ws.Cells.Item($Row, $iPub) = $_.Publisher
            $ws.Cells.Item($Row, $iDate) = $_.InstallDate
            $ws.Cells.Item($Row, $iSize) = $_.EstimatedSize
            $ws.Cells.Item($Row, $iLoc) = $_.InstallLocation
            $ws.Cells.Item($Row, $iUnins) = $_.UninstallString
            $ws.Cells.Item($Row, $iSource) = $_.Source
            $Row += 1
        }
    $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

    #==============================================================================================================================
    $Message = "Gathering installation information using winget list..."
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br />"

    #Save winget output to Excel winget tab...
    $Message = "Writing data to Excel(winget)..."
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br /><br />"
    $ws = $workbook.Worksheets.add()
    $ws.Name = 'winget'
    $ws.Activate()
    #Format cells
    $excel.Range("A1:E1").Font.bold = $true
    $excel.Range("A1:E1").Interior.ColorIndex = 11    #Dark Blue
    $excel.Range("A1:E1").Interior.Pattern = 1
    $excel.Range("A1:E1").Font.ColorIndex = 2         #White
    $excel.ActiveWindow.SplitColumn = 1
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
    $iName=1;    $ws.Cells(1, $iName).Value = "Application Name";	$ws.Columns($iName).ColumnWidth = 40 
    $iVer=2;     $ws.Cells(1, $iVer).Value = "Version";				$ws.Columns($iVer).ColumnWidth = 15
    $iAvail=3;   $ws.Cells(1, $iAvail).Value = "Available";			$ws.Columns($iAvail).ColumnWidth = 15
    $iID=4;      $ws.Cells(1, $iID).Value = "ID";	                $ws.Columns($iID).ColumnWidth = 40 
    $iSource=5;  $ws.Cells(1, $iSource).Value = "Source";	        $ws.Columns($iSource).ColumnWidth = 40 
    $Row = 2
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new() 
    $Objs = (winget list) -match '^(\p{L}|-)' |               # filter out progress-display lines
                              ConvertFrom-FixedColumnTable |  # parse output into objects
                              Sort-Object Name |              # sort by the Name property (column)
        ForEach-Object {
            $ws.Cells.Item($Row, $iName) = $_.Name
            $ws.Cells.Item($Row, $iVer) = $_.Version
            $ws.Cells.Item($Row, $iAvail) = $_.Available
            $ws.Cells.Item($Row, $iID) = $_.ID
            $ws.Cells.Item($Row, $iSource) = $_.Source
            $Row += 1
        }
    $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

    #==============================================================================================================================
    #Get-AppxPackage does not provide any significant improvement over these other methods, but preserving for posterity...
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

    $workbook.SaveAs($xlsFile) 
    $excel.Quit()
    
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br /><br />"
    Write-Separator -Path $LogPath

    #write-output "$EmailBody"

    &"$PSScriptRoot\eMailResults.ps1" -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
    #Send-Mail -AppName $AppName -EmailBody $EmailBody -LogPath $LogPath
}
