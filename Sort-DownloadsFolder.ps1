<#
.SYNOPSIS
    Organizes the Downloads folder by file type and flags oversized files.

.DESCRIPTION
    Scans files in the root of the configured DownloadsPath and:
      - Moves each file into an extension-mapped subfolder (Installers, Archives, etc.)
      - Prepends LargeFlagPrefix to any file at or above LargeFlagSizeMB, causing it to
        sort to the top of the directory listing when sorted by name.

    Files with unmapped extensions are left in place.
    Files that are locked (actively being downloaded) are skipped and retried next run.
    Designed to be run on a schedule via Task Scheduler (e.g. every 30 minutes).

.NOTES
    Requires : PowerShell 7.4+
    Scheduler: pwsh.exe -NonInteractive -File "C:\Scripts\Sort-DownloadsFolder.ps1"
    Logs     : <LogDirectory>\Sort-DownloadsFolder.log  (cleared on first run each day)
#>

#Requires -Version 7.4

[CmdletBinding()]
param(
    [Parameter()]
    [string]$DownloadsPath = (Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'),

    [Parameter()]
    [string]$ExtensionMapJson = "$PSScriptRoot\ExtensionMap.json",

    [Parameter()]
    [int]$LargeFlagSizeMB = 500,

    [Parameter()]
    [string]$LargeFlagPrefix = '!LRG_',

    [Parameter()]
    [string]$LogDirectory = $PSScriptRoot
)

#region Configuration

$ExtensionMap = (
    Get-Content $ExtensionMapJson -Raw |
    ConvertFrom-Json -AsHashtable
)

#endregion Configuration

#region Logging

$script:logFile = $null

function Initialize-Log
{
    if (-not (Test-Path -LiteralPath $LogDirectory))
    { New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null }

    $script:logFile = Join-Path $LogDirectory 'Sort-DownloadsFolder.log'

    if ((Test-Path -LiteralPath $script:logFile) -and
        (Get-Item -LiteralPath $script:logFile).LastWriteTime.Date -lt (Get-Date).Date)
    { Clear-Content -LiteralPath $script:logFile }
}

function Write-Log
{
    param(
        [Parameter(Mandatory)]
        [string] $Message,

        [ValidateSet('INFO', 'VERBOSE', 'WARN', 'ERROR')]
        [string] $Level = 'INFO'
    )

    $entry = '[{0}] [{1}] {2}' -f (Get-Date -F 's'), $Level, $Message

    $addParams = @{
        LiteralPath = $script:logFile
        Value       = $entry
        Encoding    = 'UTF8'
    }
    Add-Content @addParams

    switch ($Level)
    {
        'INFO'    { Write-Information $Message }
        'VERBOSE' { Write-Verbose     $Message }
        'WARN'    { Write-Warning     $Message }
        'ERROR'   { Write-Error       $Message }
    }
}

#endregion Logging

#region Helpers

function Resolve-DestinationPath
{
    param(
        [Parameter(Mandatory)]
        [string]$Directory,

        [Parameter(Mandatory)]
        [string]$FileName
    )

    $base      = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext       = [System.IO.Path]::GetExtension($FileName)
    $n         = 0
    $candidate = Join-Path $Directory $FileName

    while (Test-Path -LiteralPath $candidate)
    {
        $n++
        $candidate = Join-Path $Directory ('{0} ({1}){2}' -f @($base, $n, $ext))
    }

    return $candidate
}

function Test-FileLocked
{
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    try
    {
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $stream.Close()
        $stream.Dispose()

        return $false
    }
    catch
    { return $true }
}

#endregion Helpers

#region Main

Initialize-Log

Write-Log (@(
    "Sort-DownloadsFolder started.`n"
    "Downloads path: $($DownloadsPath)"
    "Size threshold: $($LargeFlagSizeMB)MB"
    "Large prefix  : $($LargeFlagPrefix)`n"
) -join "`n")

if (-not (Test-Path -LiteralPath $DownloadsPath))
{
    Write-Log "Downloads path not found: '$($DownloadsPath)'" 'ERROR'
    throw "Downloads path not found: '$($DownloadsPath)'"
}

$thresholdBytes = $LargeFlagSizeMB * 1MB
$files          = Get-ChildItem -LiteralPath $DownloadsPath -File
$stats          = @{ Moved = 0; Flagged = 0; Skipped = 0; Errors = 0 }

Write-Log "Processing $($files.Count) files in current user's Downloads folder..."

foreach ($file in $files)
{
    try
    {
        if (Test-FileLocked -Path $file.FullName)
        {
            Write-Log "File is locked: '$($file.Name)'" 'VERBOSE'
            $stats.Skipped++
            continue
        }

        $fileExtension = $file.Extension.ToLower()
        $targetFolder  = $null

        foreach ($folder in $ExtensionMap.Keys)
        {
            if ($ExtensionMap[$folder] -contains $fileExtension)
            {
                $targetFolder = $folder
                Write-Log "Matched '$($file.Name)' type: $folder" 'VERBOSE'
                break
            }
        }

        if (-not $targetFolder)
        {
            Write-Log "No type match: '$($file.Name)'" 'VERBOSE'
            continue
        }

        $currentPath = $file.FullName
        $currentName = $file.Name

        if ($file.Length -ge $thresholdBytes -and
            -not $currentName.StartsWith($LargeFlagPrefix))      # Flag !LRG_ if over LargeFlagSizeMB
        {
            $flaggedName      = $LargeFlagPrefix + $currentName
            $resolvedFlagPath = Resolve-DestinationPath -Directory $file.DirectoryName -FileName $flaggedName
            $resolvedFlagName = [System.IO.Path]::GetFileName($resolvedFlagPath)

            Rename-Item -LiteralPath $currentPath -NewName $resolvedFlagName -EA 'Stop'
            Write-Log "Flagged large file '$currentName' ($([math]::Round($file.Length / 1MB, 1))MB)" 'VERBOSE'
            $stats.Flagged++

            $currentName = $resolvedFlagName
            $currentPath = Join-Path $file.DirectoryName $currentName
        }

        $targetDir = Join-Path $DownloadsPath $targetFolder

        if (-not (Test-Path -LiteralPath $targetDir))
        {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            Write-Log "Created missing file type subfolder: '$targetFolder'" 'VERBOSE'
        }

        $destPath = Resolve-DestinationPath -Directory $targetDir -FileName $currentName

        Move-Item -LiteralPath $currentPath -Destination $destPath -EA 'Stop'
        Write-Log "Moved '$currentName' to: '$targetFolder'"
        $stats.Moved++
    }
    catch
    {
        Write-Log "Error processing '$($file.Name)':`n$_" 'ERROR'
        $stats.Errors++
    }
}

Write-Log (@(
    "Done.`n"
    "Summary:"
    "Moved  : $($stats.Moved)"
    "Flagged: $($stats.Flagged)"
    "Skipped: $($stats.Skipped)"
    "Errors : $($stats.Errors)`n"
) -join "`n")

Write-Log "Sort-DownloadsFolder complete."

#endregion Main
