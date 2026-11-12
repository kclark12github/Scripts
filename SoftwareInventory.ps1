#SoftwareInventory.ps1
#   Build Utility Script...
#   Copyright © 1998-2026, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   11/12/26    Moved UserName and Password into environment variables & added Inventoried date;
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
    $EmailBody += Write-To-Excel -excel $excel -workbook $workbook -RegKey $RegKey32 -Tab "Uninstall"   -BackColor 14 -FontColor 2 -LogPath $LogPath  #Teal/White
    $EmailBody += Write-To-Excel -excel $excel -workbook $workbook -RegKey $RegKey64 -Tab "Wow6432Node" -BackColor 10 -FontColor 6 -LogPath $LogPath  #Dark Green/Yellow
    $EmailBody += Write-To-Excel -excel $excel -workbook $workbook                   -Tab "winget"      -BackColor 11 -FontColor 2 -LogPath $LogPath  #Dark Blue/White
    $EmailBody += "<br />"

    $workbook.SaveAs($xlsFile) 
    $excel.Quit()
  
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
    #==============================================================================================================================
    $Message = "`nSaving installation data in $ServerName.$DatabaseName.$TableName...";   Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br />"
    $EmailBody += Write-To-SQL -Tag $RegKey32 -LogPath $LogPath
    $EmailBody += Write-To-SQL -Tag $RegKey64 -LogPath $LogPath
    $EmailBody += Write-To-SQL -Tag "winget"  -LogPath $LogPath
    $EmailBody += "<br />"
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
function Write-To-Excel {
    Param($excel, $workbook, $RegKey, $Tab, $BackColor, $FontColor=2, $LogPath)
    $errMessage = ""
    if ($Tab -eq "winget") {
        $wteMessage = "Gathering installation information using winget list..."
        $Range = "A1:I1"
    } else {
        $wteMessage = "Gathering installation information from registry ($RegKey)..."
        $Range = "A1:H1"
    }
    Write-Message -Message $wteMessage -Path $LogPath
    "<br />" #Write-Message is causing messages to be returned as the return value for this function, so inject line breaks where necessary

    if ($Tab -ne "winget") {
        $Objs = @()
        $InstalledAppsInfos = Get-ItemProperty -Path $RegKey64
        Foreach($InstalledAppsInfo in $InstalledAppsInfos)
        {
            $Obj = [PSCustomObject]@{Computer=$Root;
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
    }

    #Save $objs to Excel $Tab tab...
    $wteMessage = "Writing data to Excel($Tab)..."
    Write-Message -Message $wteMessage -Path $LogPath
    "<br />" #Write-Message is causing messages to be returned as the return value for this function, so inject line breaks where necessary

    if ($workbook.Worksheets.Item(1).Name -eq "Sheet1") {
        $ws = $workbook.Worksheets.Item(1) #Use the new empty sheet
    } else {
        $ws = $workbook.Worksheets.add()
    }
    $ws.Name = $Tab
    $ws.Tab.ColorIndex = $BackColor
    #Format cells
    $excel.Range($Range).Font.bold = $true
    $excel.Range($Range).Interior.ColorIndex = $BackColor
    $excel.Range($Range).Interior.Pattern = 1
    $excel.Range($Range).Font.ColorIndex = $FontColor
    $excel.ActiveWindow.SplitColumn = 1
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
    if ($Tab -eq "winget") {
        $iName=1;    $ws.Cells(1, $iName).Value = "Application Name";	$ws.Columns($iName).ColumnWidth = 40 
        $iVer=2;     $ws.Cells(1, $iVer).Value = "Version";				$ws.Columns($iVer).ColumnWidth = 15
        $iAvail=3;   $ws.Cells(1, $iAvail).Value = "Available";			$ws.Columns($iAvail).ColumnWidth = 15
        $iSource=4;  $ws.Cells(1, $iSource).Value = "Source";	        $ws.Columns($iSource).ColumnWidth = 40 
        $iPub=5;     $ws.Cells(1, $iPub).Value = "Publisher";			$ws.Columns($iPub).ColumnWidth = 15 
        $iDate=6;    $ws.Cells(1, $iDate).Value = "Install Date";		$ws.Columns($iDate).ColumnWidth = 15 
        $iSize=7;    $ws.Cells(1, $iSize).Value = "Estimated Size";		$ws.Columns($iSize).ColumnWidth = 15 
        $iLoc=8;     $ws.Cells(1, $iLoc).Value = "Install Location";	$ws.Columns($iLoc).ColumnWidth = 40 
        $iID=9;      $ws.Cells(1, $iID).Value = "ID";	                $ws.Columns($iID).ColumnWidth = 40 
    } else {
        $iName=1;    $ws.Cells(1, $iName).Value = "Application Name";	$ws.Columns($iName).ColumnWidth = 40 
        $iVer=2;     $ws.Cells(1, $iVer).Value = "Version";				$ws.Columns($iVer).ColumnWidth = 15
        $iPub=3;     $ws.Cells(1, $iPub).Value = "Publisher";			$ws.Columns($iPub).ColumnWidth = 15 
        $iDate=4;    $ws.Cells(1, $iDate).Value = "Install Date";		$ws.Columns($iDate).ColumnWidth = 15 
        $iSize=5;    $ws.Cells(1, $iSize).Value = "Estimated Size";		$ws.Columns($iSize).ColumnWidth = 15 
        $iLoc=6;     $ws.Cells(1, $iLoc).Value = "Install Location";	$ws.Columns($iLoc).ColumnWidth = 40 
        $iUnins=7;   $ws.Cells(1, $iUnins).Value = "Uninstall String";	$ws.Columns($iUnins).ColumnWidth = 40 
        $iSource=8;  $ws.Cells(1, $iSource).Value = "Source";	        $ws.Columns($iSource).ColumnWidth = 40 
    }

    $Row = 2

    if ($Tab -eq "winget") {
        #[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new() 
        try {[Console]::OutputEncoding = [System.Text.Encoding]::UTF8} catch {}   # Ignore invalid handle errors in ISE or background Task Scheduler sessions
        $Objs = (winget list) -match '^(\p{L}|-)' |               # filter out progress-display lines
                                  ConvertFrom-FixedColumnTable |  # parse output into objects
                                  Sort-Object Name |              # sort by the Name property (column)
            ForEach-Object {
                $Name = $_.Name
                $Version = $_.Version
                $Available = $_.Available
                $Source = $_.Source
                $Publisher = "Unavailable"
                $InstallDate = "Unavailable"
                $EstimatedSize = "Unavailable"
                $InstallLocation = "Unavailable"
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
                        $RegKey = $RegKey32.Replace("*", $ID)
                        $reg = Get-ItemProperty -Path $RegKey
                        $Source = "Uninstall"
                        $Name = $reg.DisplayName
                        $Publisher = $reg.Publisher
                        $InstallDate = $reg.InstallDate
                        $EstimatedSize = $reg.EstimatedSize
                        $InstallLocation = $reg.InstallLocation
                    } Catch [System.Management.Automation.ItemNotFoundException]{
                        Try {
                            #Next look in the 64-bit Uninstall Registry Key...
                            $RegKey = $RegKey64.Replace("*", $ID)
                            $reg = Get-ItemProperty -Path $RegKey
                            $Source = "Wow5432Node"
                            $Name = $reg.DisplayName
                            $Publisher = $reg.Publisher
                            $InstallDate = $reg.InstallDate
                            $EstimatedSize = $reg.EstimatedSize
                            $InstallLocation = $reg.InstallLocation
                        } Catch [System.Management.Automation.ItemNotFoundException]{
                            #Where Next?
                            $Source = "Unavailable"
                        }
                    } Finally { 
                        $ErrorActionPreference = "Continue"
                        $null = $reg
                        $null = $RegKey
                    }
                }

                $ws.Cells.Item($Row, $iName) = $Name
                $ws.Cells.Item($Row, $iVer) = $Version
                $ws.Cells.Item($Row, $iAvail) = $Available
                $ws.Cells.Item($Row, $iSource) = $Source
                $ws.Cells.Item($Row, $iPub) = $Publisher
                $ws.Cells.Item($Row, $iDate) = $InstallDate
                $ws.Cells.Item($Row, $iSize) = $EstimatedSize
                $ws.Cells.Item($Row, $iLoc) = $InstallLocation
                $ws.Cells.Item($Row, $iID) = $ID
                $Row += 1
            }
    } else {
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
    Param($Tag, $LogPath)

    switch ($Tag) {
        "$RegKey32" {$Tag = "Uninstall";   break}
        "$RegKey64" {$Tag = "Wow6432Node"; break}
        default {break}
        }
 
    $username = $env:TC2User; $password = ConvertTo-SecureString $env:TC2Password -AsPlainText -Force; $password.MakeReadOnly()
    $cred = New-Object System.Management.Automation.PSCredential ($username, $password)
    $DeleteQuery = "DELETE FROM [$TableName] WHERE [System] = '$Root' And RegKey='$Tag';"
    Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Credential $cred -Query $DeleteQuery -TrustServerCertificate

    $Inventoried = (Get-Date).Date

    $Objs = @()
    if ($Tag -ne "winget") {
        $InstalledAppsInfos = Get-ItemProperty -Path $RegKey64 | Where-Object { $_.DisplayName -ne $null }
        Foreach($InstalledAppsInfo in $InstalledAppsInfos)
        {
            $ParsedDate = [DateTime]::MinValue
            $CleanDate  = $null
            if (-not [string]::IsNullOrWhiteSpace($InstalledAppsInfo.InstallDate)) {
                if ([DateTime]::TryParseExact($InstalledAppsInfo.InstallDate.Trim(), "yyyyMMdd", $null, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
                    $CleanDate = $ParsedDate.Date
                }
            }
            $SizeMB = $null
            if ($null -ne $InstalledAppsInfo.EstimatedSize -and $InstalledAppsInfo.EstimatedSize -gt 0) {
                $CalculatedMB = [math]::Round(($InstalledAppsInfo.EstimatedSize / 1024), 2)
                $SizeMB = "$CalculatedMB MB"
            }

            $Obj = [PSCustomObject]@{System=$Root;
                                     ApplicationName = $InstalledAppsInfo.DisplayName;
                                     Version = $InstalledAppsInfo.DisplayVersion;
                                     Publisher = $InstalledAppsInfo.Publisher;
                                     InstallDate = $CleanDate
                                     EstimatedSize = $SizeMB
                                     InstallLocation = $InstalledAppsInfo.InstallLocation;
                                     UninstallString = $InstalledAppsInfo.UninstallString;
                                     Source = $InstalledAppsInfo.PSPath;
                                     RegKey = $Tag;
                                     Inventoried = $Inventoried;
                                     }
            $Objs += $Obj
        }
    } else {
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
                        $RegKey = $RegKey32.Replace("*", $ID)
                        $reg = Get-ItemProperty -Path $RegKey
                        $Source = "Uninstall"
                        $Name = $reg.DisplayName
                        $Publisher = $reg.Publisher
                        $InstallDate = $reg.InstallDate
                        $EstimatedSize = $reg.EstimatedSize
                        $InstallLocation = $reg.InstallLocation
                    } Catch [System.Management.Automation.ItemNotFoundException]{
                        Try {
                            #Next look in the 64-bit Uninstall Registry Key...
                            $RegKey = $RegKey64.Replace("*", $ID)
                            $reg = Get-ItemProperty -Path $RegKey
                            $Source = "Wow5432Node"
                            $Name = $reg.DisplayName
                            $Publisher = $reg.Publisher
                            $InstallDate = $reg.InstallDate
                            $EstimatedSize = $reg.EstimatedSize
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
                $ParsedDate = [DateTime]::MinValue
                $CleanDate  = $null
                if (-not [string]::IsNullOrWhiteSpace($InstallDate)) {
                    if ([DateTime]::TryParseExact($InstallDate.Trim(), "yyyyMMdd", $null, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
                        $CleanDate = $ParsedDate.Date
                    }
                }
                $SizeMB = $null
                if ($null -ne $EstimatedSize -and $EstimatedSize -gt 0) {
                    $CalculatedMB = [math]::Round(($EstimatedSize / 1024), 2)
                    $SizeMB = "$CalculatedMB MB"
                }
                $Obj = [PSCustomObject]@{System=$Root;
                                     ApplicationName = $Name;
                                     Version = $Version;
                                     Publisher = $Publisher;
                                     InstallDate = $CleanDate
                                     EstimatedSize = $SizeMB
                                     InstallLocation = $InstallLocation;
                                     UninstallString = $null;
                                     Source = $Source;
                                     RegKey = $Tag;
                                     Inventoried = $Inventoried
                                     }
                $Objs += $Obj
            }
    }

    # Insert the fresh data

    $Objs | Write-SqlTableData -ServerInstance $ServerName -DatabaseName $DatabaseName -Credential $cred -SchemaName "dbo" -TableName $TableName -TrustServerCertificate
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
    
    $ServerName = "7X4L243"; $DatabaseName = "TreasureChest2"; $TableName = "Apps"
    $RegKey32 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $RegKey64 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    
    #DoExcel
    DoSQL
 
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"; 
    Write-Message -Message $Message -Path $LogPath;    $EmailBody += "$Message<br /><br />"
    Write-Separator -Path $LogPath
        
    &"$PSScriptRoot\eMailResults.ps1" -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
    #Send-Mail -AppName $AppName -EmailBody $EmailBody -LogPath $LogPath
}
