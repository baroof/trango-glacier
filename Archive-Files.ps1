<#
.SYNOPSIS
    Selectively copy modified files to archive folder

.DESCRIPTION
    Scans -from folder, matching last-modified-date to files already in the -to folder, skipping any matches.

    Newly modified files are copied over, optionally w/ filenames changed to include last modified date.

    Optionally restricted to specified file extensions.

    Optionally sweeps the archive folder, deleting anything older than X months.

.PARAMETER from
    required: where the files come from

.PARAMETER to
    required: where should the copies go? MUST ALREADY EXIST

.PARAMETER extensions
    optional: comma-delimited list of file extensions to grab (jpg,txt,gif)

.PARAMETER sweep
    optional: delete archived items older than [#] months

.PARAMETER modified
    optional: add the last-modified-date to the filename.

.OUTPUTS
  - success/fail messages from Test-Path and Copy-Item
  - logs to C:\Windows\Logs\Archive-Files.ps1.[date].log

.NOTES
  Version:        1.0
  Author:         Will Mooreston
  Creation Date:  2018-01-26
    
.EXAMPLE
    Archive-Files.ps1 -from \\server1\sourceFolder -to \\server1\sourceFolder\Archive

    This copies all files under sourceFolder to sourceFolder\Archive

.EXAMPLE
    Archive-Files.ps1 -from \\server1\sourceFolder -to \\server1\sourceFolder\Archive -sweep 1 -extensions jpg,jpeg
    
    This copies all .jpg and .jpeg files under sourceFolder to sourceFolder\Archive, 
    then DELETES anything in Archive last modified more than 1 month ago)

.EXAMPLE
    Archive-Files.ps1 -from \\server1\sourceFolder -to \\server1\sourceFolder\Archive -modifed
    
    This copies all files under sourceFolder to sourceFolder\Archive, and adds last-modified-date to each filename (yyyyMMdd_HHmmss)
    filename.jpg -> filename_lastMod20170602_100023.jpg
#>

param(
    [ValidateScript({Test-Path $_})]
    [string]
    $from
,
    [ValidateScript({Test-Path $_})]
    [string]
    $to
,
    [string[]]
    $extensions = "*"
,
    [int32]
    $sweep = 0
,
    [switch]
    $modified = $false
)

function Log($message) {
# print message to screen as well as log
    $message
    if ($logfile) {
        $timestamp = Get-Date -Format o 
        $message = "$timestamp | $message"
        $message | Out-File -Append $logfile
    }
}

function CopyFile($ff, $tf) {
    Log "Trying to copy: $ff"

    Try{
        $result = Copy-Item -LiteralPath $ff -Destination $tf -PassThru -ErrorAction Stop  
    }
    Catch{
        Log "`nERROR on COPY: "+$_.exception.message
    }

    if (Test-Path $tf) {
        Log "Succcesful copy to $tf"
    }
}

# exit unless required parameters are set
if (-not $from -or -not $to) {Log("ERROR: missing required parameter: -from ($from) | -to ($to)"); exit}

#set up logging
$logfile =  "C:\Windows\Logs\"+$MyInvocation.MyCommand+"."+(Get-Date).ToString("yyyyMMdd")+".log"
Log "... starting run ..."

foreach ($ext in $extensions) {
    $ext = "*."+$ext
    $targetFiles += Get-ChildItem $from"\*" -file -Include $ext
}


foreach ($file in $targetFiles) {
    $fileFullName = $file.FullName
    $fileLastMod = Get-Date(Get-Item $fileFullName).LastWriteTime
    $fileLastModString = $fileLastMod.ToString("yyyyMMdd_HHmmss")
    if ($modified -eq $True) {
        $toFileFullName = $to+"\"+$file.BaseName+"_lastMod"+$fileLastModString+$file.Extension
    } else {
        $toFileFullName = $to+"\"+$file.name
    }

    if (Test-Path $toFileFullName) {
    #check if we already copied this file over, based on last modified date
        $toFileLastMod = Get-Date(Get-Item $toFileFullName).LastWriteTime -ErrorAction Stop
        if ($fileLastMod -eq $toFileLastMod) {
            Log "Skipping: '$fileFullName' b/c '$toFileFullName' exists | last-mod:$fileLastModString"
        } else {
            CopyFile $fileFullName $toFileFullName
        }
    } else {
        CopyFile $fileFullName $toFileFullName
    }
}

<#
# archive anything new, checking modified time in case they manually edit a file or something
# add the mod time to the new file name
foreach ($file in $targetFiles) {
    $fileFullName = $file.FullName
    $lastMod = (Get-Date(Get-Item $fileFullName).LastWriteTime).ToString("yyyyMMdd_HHmmss")
    $newFileName = $to+"\"+$file.BaseName+"_lastMod"+$lastMod+$file.Extension

    if (Test-Path $newFileName) {
        Log "Skipping: $fileFullName because $newFileName already exists!"
    } else {
        Log "Trying to copy: $fileFullName"

        Try{
            $result = Copy-Item -LiteralPath $fileFullName -Destination $newFileName -PassThru -ErrorAction Stop  
        }
        Catch{
            Log "`nERROR on COPY: "+$_.exception.message
        }

        if (Test-Path $newFileName) {
            Log "Succcesful copy to $newFileName"
        }
    }
}
#>
if ($sweep -gt 0) {
# delete anything more than [$sweep] months old, again checking mod time
# default of 0 months skips the sweep
    $archivedFiles = Get-ChildItem $to -file
    foreach ($file in $archivedFiles) {
        $fileFullName = $file.FullName
        $lastMod = Get-Date(Get-Item $fileFullName).LastWriteTime
    
        if ($lastMod -le (Get-Date).AddMonths(-$sweep)) {
            Log "Trying to delete old file: $fileFullName, modified $lastMod"

            Try{
                $result = Remove-Item -LiteralPath $fileFullName -ErrorAction Stop 
            }
            Catch{
                Log "`nERROR on DELETION: "+$_.exception.message
            }

            if (-not (Test-Path $fileFullName)) {
                Log "successully deleted: $fileFullName"
            }
        }
    }
}
    
Log("... ending run ...")