#RoboCopyMove.ps1
#   RoboCopy Script to Move Provided Folders...
#   Copyright © 1998-2021, Ken Clark
#*********************************************************************************************************************************
#   Modification History:
#   Date:       Description:
#   06/13/21    Created;
#=================================================================================================================================
#Notes to Self:
#This "Move" version should probably be rewritten into a generic RoboCopy shell script supporting both Mirror and Move, but I 
#don't have the ambition (or time) to putz with capturing the arguments or passing options as a string and dealing with it all.
#=================================================================================================================================
# 

#Enable -Verbose option
[CmdletBinding()]
param([string]$Title, [string]$Source, [string]$Target, [string]$LogFile, [switch]$SuppressFailure)

$Success = $True

$Body = "RoboCopy ""$Source"" ""$Target"" /s /move /np /log:""$LogFile""`n`n"
RoboCopy $Source $Target /mir /np /log:$LogFile
switch($LASTEXITCODE) {
    0 { 
        $Body = $Body + "No files were copied. No failure was encountered. No files were mismatched. The files already exist in the destination directory; therefore, the copy operation was skipped."
	}
    1 {
        $Body = $Body + "All files were copied successfully."
    }
    2 {
        $Body = $Body + "There are some additional files in the destination directory that are not present in the source directory. No files were copied."
    }
    3 {
        $Body = $Body + "Some files were copied. Additional files were present. No failure was encountered."
    }
    4 {
        $Body = $Body + "Return code (4) is not defined!"
        $Success = $False
    }
    5 {
        $Body = $Body + "Some files were copied. Some files were mismatched. No failure was encountered."
    }
    6 {
        $Body = $Body + "Additional files and mismatched files exist. No files were copied and no failures were encountered. This means that the files already exist in the destination directory."
    }
    7 {
        $Body = $Body + "Files were copied, a file mismatch was present, and additional files were present."
    }
    8 {
        $Body = $Body + "Several files did not copy."
        $Success = $False
    }
    default { 
		$Body = $Body + "There was at least one failure during the copy operation."
        $Success = $False
	}
}

$Body = $Body + "`n`nLog: $LogFile`n"
if ($Success) {
    $Subject = "$Title Succeeded"
} else {
    $Subject = "$Title Failed"
}

if ($Success -or -not $SuppressFailure) {
    &"$PSScriptRoot\eMailResults.ps1" -Subject $Subject -Body $Body -LogFile $LogFile 
} else {
    Write-Host "Suppressed Failure E-mail"
}
if ($Success -or $SuppressFailure) {
    exit 0
} else {
    exit 1
}
