<#
.SYNOPSIS
    Organizes the Downloads folder by file type and flags oversized files.

.DESCRIPTION
    Scans files in the root of the configured DownloadsPath and:
      - Moves each file into an extension-mapped subfolder (Installers, Archives, etc.)
      - Prepends LargeFlagPrefix to any file at or above LargeFlagSizeMB, causing it to
        sort to the top of the directory listing when sorted by name.

    Files that are locked (actively being downloaded) are skipped and retried next run.
    Designed to be run on a schedule via Task Scheduler (e.g. every 30 minutes).

.NOTES
    Requires : PowerShell 7.4+
    Scheduler: pwsh.exe -NonInteractive -File "C:\Scripts\Sort-DownloadsFolder.ps1"
    Logs     : <LogDirectory>\Sort-DownloadsFolder_yyyyMMdd.log  (daily rollover)
#>

#Requires -Version 7.4

#region Configuration

$config = @{
    DownloadsPath   = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
    LargeFlagSizeMB = 500
    LargeFlagPrefix = '!LRG_'
    CatchAllFolder  = '_Unsorted'
    LogDirectory    = $PSScriptRoot
    ExtensionMap    = [ordered]@{
        'Installers' = @('.exe', '.msi', '.msix', '.appx', '.pkg', '.dmg', '.run', '.deb', '.rpm')
        'Archives'   = @('.zip', '.7z', '.rar', '.tar', '.gz', '.bz2', '.xz', '.cab', '.iso', '.img')
        'Documents'  = @('.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.csv', '.rtf', '.odt', '.ods', '.odp', '.md')
        'Images'     = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp', '.tiff', '.tif', '.ico', '.heic', '.raw', '.cr2', '.nef')
        'Audio'      = @('.mp3', '.flac', '.wav', '.aac', '.ogg', '.m4a', '.wma', '.opus', '.aiff')
        'Video'      = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.webm', '.m4v', '.mpg', '.mpeg')
        'Code'       = @('.ps1', '.psm1', '.psd1', '.py', '.js', '.ts', '.json', '.xml', '.yaml', '.yml', '.sh', '.bat', '.cmd', '.cs', '.cpp', '.h', '.java', '.rb', '.go', '.rs')
    }
}

#endregion Configuration

#region Logging

$script:logFile = $null

function Initialize-Log
{
    if (-not (Test-Path -LiteralPath $config.LogDirectory))
    { New-Item -ItemType Directory -Path $config.LogDirectory -Force | Out-Null }

    $script:logFile = Join-Path $config.LogDirectory 'Sort-DownloadsFolder_RollingLog.log'
}

function Write-Log
{
    param(
        [Parameter(Mandatory)]
        [string] $Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
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
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Error   $Message }
        default { Write-Verbose $Message }
    }
}

#endregion Logging

#region Helpers

function Get-TargetFolder
{
    <#
    .SYNOPSIS
        Returns the mapped subfolder name for a file extension, or CatchAllFolder if unmapped.
    #>
    param(
        [Parameter(Mandatory)]
        [string] $Extension
    )

    $ext = $Extension.ToLower()

    foreach ($folder in $config.ExtensionMap.Keys)
    {
        if ($config.ExtensionMap[$folder] -contains $ext)
        {
            return $folder
        }
    }

    return $config.CatchAllFolder
}

function Resolve-DestinationPath
{
    <#
    .SYNOPSIS
        Returns a non-colliding destination path, appending _N to the base name if the
        target already exists.
    #>
    param(
        [Parameter(Mandatory)]
        [string] $Directory,

        [Parameter(Mandatory)]
        [string] $FileName
    )

    $candidate = Join-Path $Directory $FileName
    if (-not (Test-Path -LiteralPath $candidate))
    {
        return $candidate
    }

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext  = [System.IO.Path]::GetExtension($FileName)
    $n    = 1

    do
    {
        $candidate = Join-Path $Directory ('{0}_{1}{2}' -f $base, $n, $ext)
        $n++
    }
    while (Test-Path -LiteralPath $candidate)

    return $candidate
}

function Test-FileLocked
{
    <#
    .SYNOPSIS
        Returns $true if the file cannot be opened exclusively — e.g. it is still being
        written to by a browser download.
    #>
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
    'Sort-DownloadsFolder started.'
    "Downloads path  : $($config.DownloadsPath)"
    "Size threshold  : $($config.LargeFlagSizeMB)MB"
    "Large prefix    : $($config.LargeFlagPrefix)"
    "Catch-all folder: $($config.CatchAllFolder)"
) -join "`n")

if (-not (Test-Path -LiteralPath $config.DownloadsPath))
{
    Write-Log "Downloads path not found: $($config.DownloadsPath)" -Level ERROR
    throw "Downloads path not found: $($config.DownloadsPath)"
}

$thresholdBytes = $config.LargeFlagSizeMB * 1MB
$files          = Get-ChildItem -LiteralPath $config.DownloadsPath -File
$stats          = @{ Moved = 0; Flagged = 0; Skipped = 0; Errors = 0 }

Write-Log "Files in Downloads root: $($files.Count)"

foreach ($file in $files)
{
    try
    {
        if (Test-FileLocked -Path $file.FullName)
        {
            Write-Log "Skipped (locked): $($file.Name)"
            $stats.Skipped++
            continue
        }

        $currentPath = $file.FullName
        $currentName = $file.Name

        if ($file.Length -ge $thresholdBytes -and
            -not $currentName.StartsWith($config.LargeFlagPrefix))      # Flag !LRG_ if over LargeFlagSizeMB
        {
            $flaggedName      = $config.LargeFlagPrefix + $currentName
            $resolvedFlagPath = Resolve-DestinationPath -Directory $file.DirectoryName -FileName $flaggedName
            $resolvedFlagName = [System.IO.Path]::GetFileName($resolvedFlagPath)

            Rename-Item -LiteralPath $currentPath -NewName $resolvedFlagName -EA 'Stop'
            Write-Log "Flagged '$currentName' ($([math]::Round($file.Length / 1MB, 1))MB)"
            $stats.Flagged++

            $currentName = $resolvedFlagName
            $currentPath = Join-Path $file.DirectoryName $currentName
        }

        $targetFolder = Get-TargetFolder -Extension $file.Extension     # Sort into type folder if applicable
        $targetDir    = Join-Path $config.DownloadsPath $targetFolder

        if (-not (Test-Path -LiteralPath $targetDir))
        {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            Write-Log "Created missing file type subfolder: $targetFolder"
        }

        $destPath = Resolve-DestinationPath -Directory $targetDir -FileName $currentName

        Move-Item -LiteralPath $currentPath -Destination $destPath -EA 'Stop'
        Write-Log "Moved '$currentName' to: '$targetFolder'"
        $stats.Moved++
    }
    catch
    {
        Write-Log "Error processing '$($file.Name)':`n$_" -Level ERROR
        $stats.Errors++
    }
}

Write-Log (@(
    'Sort-DownloadsFolder complete.'
    "Moved  : $($stats.Moved)"
    "Flagged: $($stats.Flagged)"
    "Skipped: $($stats.Skipped)"
    "Errors : $($stats.Errors)"
) -join "`n")

#endregion Main
