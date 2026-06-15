<#PSScriptInfo
.VERSION 1.0.0
.AUTHOR Ryan Eaton
#>

<#
.SYNOPSIS
    Sorts Downloads folder files into type-based subfolders and flags large files.

.DESCRIPTION
    On each run, processes all files and subfolders in the root of DownloadsPath:

    Files:
      - Skipped if locked (in-progress download) or created within the last MinAgeMinutes.
      - Renamed with LargeFlagPrefix if at or above LargeFlagSizeMB (floats to top when sorted by name).
      - Moved into a type-matched subfolder (prefixed with '!!' by default to sort above unmanaged folders).
      - Moved to StaleFolderName if unmatched and older than StaleDaysThreshold days,
        unless large (large unmatched files are flagged and left in the root).

    Subfolders:
      - Skipped if the name starts with CategoryFolderPrefix or matches StaleFolderName.
      - Renamed with LargeFlagPrefix if at or above LargeFlagSizeMB.
      - Moved to StaleFolderName if older than StaleDaysThreshold days and not large.

    A reminder file is created in the root if the skipped file count reaches CleanReminderThreshold.

.PARAMETER DownloadsPath
    Folder to process. Defaults to the current user's Downloads folder.

.PARAMETER LargeFlagPrefix
    Prefix prepended to large files and folders. Default: '!LRG_'.

.PARAMETER CategoryFolderPrefix
    Prefix for script-managed category subfolders. Default: '!!'.
    Folders with this prefix sort above user-floated '!' folders and are skipped by the folder loop.

.PARAMETER StaleFolderName
    Name appended to CategoryFolderPrefix for the stale items folder. Default: 'zStale'.
    The 'z' causes it to sort below other category folders.

.PARAMETER LogDirectory
    Directory for the rolling log file. Defaults to the script directory.

.PARAMETER NoLogFile
    Disables all file logging for the run.

.PARAMETER Quiet
    Suppresses all console output. Log file is still written unless -NoLogFile is also set.

.PARAMETER ExtensionMapJson
    Path to a JSON file fully replacing the built-in extension map.
    Format: { "Category": [".ext1", ".ext2"], ... }

.PARAMETER ExtensionMapOverrideJson
    Path to a JSON file that adds to or overrides categories in the built-in extension map.
    Extensions moved to a new category are automatically removed from their original one.
    To customize without an external file, edit $scriptExtensionMapOverride in-script.

.PARAMETER LargeFlagSizeMB
    Size threshold in MB for large-file and large-folder flagging. Default: 100. Set to 0 to disable.

.PARAMETER MinAgeMinutes
    Files created within this many minutes are left in place regardless of other rules. Default: 30. Set to 0 to disable.

.PARAMETER StaleDaysThreshold
    Files and folders last modified more than this many days ago are considered stale. Default: 7. Set to 0 to disable.

.PARAMETER CleanReminderThreshold
    Creates a visible reminder file in the root when skipped file count reaches this value. Default: 15. Set to 0 to disable.

.PARAMETER Skip
    Skips processing of 'Files', 'Folders', or both. Accepts multiple values.
    Example: -Skip Files,Folders

.PARAMETER RestoreDownloadsRoot
    Moves all contents of script-managed category folders back into the Downloads root and
    removes the now-empty category folders. Large-flagged prefixes are not removed from filenames.

.PARAMETER NoSizeLabels
    Skips the size label pass at the end of each run. Category folder names will not be updated
    with their current size.

.NOTES
    Requires : PowerShell 7.4+
    Log      : <LogDirectory>\Sort-DownloadsFolder.log  (cleared on first run each day)

    Register as a per-user scheduled task running every 30 minutes (no elevation required):

        $script = 'C:\Scripts\Sort-DownloadsFolder.ps1'

        $action = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument "-NonInteractive -File `"$script`""

        $triggerParams = @{
            Once               = $true
            At                 = Get-Date
            RepetitionInterval = New-TimeSpan -Minutes 30
        }
        $trigger = New-ScheduledTaskTrigger @triggerParams

        $settingsParams = @{
            ExecutionTimeLimit          = New-TimeSpan -Minutes 5
            StartWhenAvailable          = $true
            RunOnlyIfNetworkAvailable   = $false
        }
        $settings = New-ScheduledTaskSettingsSet @settingsParams

        $taskParams = @{
            TaskName  = 'Sort-DownloadsFolder'
            Action    = $action
            Trigger   = $trigger
            Settings  = $settings
            RunLevel  = 'Limited'
            Force     = $true
        }
        Register-ScheduledTask @taskParams
#>

#Requires -Version 7.4

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateNotNullOrWhiteSpace()]
    [string]$DownloadsPath      = (Join-Path ([System.Environment]::GetFolderPath('UserProfile')) 'Downloads'),

    [ValidateNotNullOrWhiteSpace()]
    [string]$LargeFlagPrefix    = '!LRG_',

    [ValidatePattern('^[^\\/:*?"<>|]+$')]
    [ValidateNotNullOrWhiteSpace()]
    [string]$CategoryFolderPrefix = '!!',

    [ValidatePattern('^[^\\/:*?"<>|!][^\\/:*?"<>|]*$')]
    [ValidateNotNullOrWhiteSpace()]
    [string]$StaleFolderName    = 'zStale',

    [ValidateNotNullOrWhiteSpace()]
    [string]$LogDirectory       = $PSScriptRoot,

    [ValidateNotNullOrWhiteSpace()]
    [string]$ExtensionMapJson,          # Use to fully replace the default map
    
    [ValidateNotNullOrWhiteSpace()]
    [string]$ExtensionMapOverrideJson,  # Use to add to or modify the default map

    [ValidateSet('Files','Folders')]
    [string[]]$Skip,

    [ValidateRange(0,102400)]
    [int]$LargeFlagSizeMB        = 100,

    [ValidateRange(0,518400)]
    [int]$MinAgeMinutes          = 30,

    [ValidateRange(0,10000)]
    [int]$StaleDaysThreshold     = 7,

    [ValidateRange(0,1000000)]
    [int]$CleanReminderThreshold = 15,

    [switch]$RestoreDownloadsRoot,
    [switch]$NoSizeLabels,
    [switch]$NoLogFile,
    [switch]$Quiet
)

[version]$script:Version = '1.0.0'

#region In-Script Map Override

$scriptExtMapOverrideEnabled = $false   # Turn this override map on or off

if ($scriptExtMapOverrideEnabled -and
    (-not $ExtensionMapOverrideJson))
{
    $scriptExtensionMapOverride = @{
        CSV = @(
            ".csv"                      # Puts .csv in !!CSV instead of !!Data
        )
    }
}

#endregion In-Script Map Override

#region Extension Map

if (-not $ExtensionMapJson)
{
    $defaultExtensionMap = @{
        Android = @(
            '.aab'
            '.apk'
            '.xapk'
        )
        Archives = @(
            '.7z'
            '.bz2'
            '.cab'
            '.gz'
            '.lz4'
            '.rar'
            '.tar'
            '.tgz'
            '.xz'
            '.zip'
            '.zst'
        )
        Audio = @(
            '.aac'
            '.aiff'
            '.alac'
            '.flac'
            '.m4a'
            '.mp3'
            '.ogg'
            '.opus'
            '.wav'
            '.wma'
        )
        Certs = @(
            '.cer'
            '.crt'
            '.csr'
            '.der'
            '.key'
            '.p12'
            '.p7b'
            '.p8'
            '.pem'
            '.pfx'
        )
        Code = @(
            '.c'
            '.cpp'
            '.cs'
            '.go'
            '.h'
            '.hpp'
            '.java'
            '.js'
            '.kt'
            '.lua'
            '.rb'
            '.rs'
            '.swift'
            '.ts'
            '.vb'
        )
        ConfigFiles = @(
            '.cfg'
            '.conf'
            '.config'
            '.env'
            '.ini'
            '.properties'
            '.toml'
            '.yaml'
            '.yml'
        )
        Data = @(
            '.accdb'
            '.csv'
            '.db'
            '.json'
            '.mdb'
            '.ods'
            '.parquet'
            '.sqlite'
            '.tsv'
            '.xls'
            '.xlsx'
            '.xml'
        )
        DiskImages = @(
            '.img'
            '.iso'
            '.ova'
            '.ovf'
            '.vdi'
            '.vhd'
            '.vhdx'
            '.vmdk'
        )
        Documents = @(
            '.doc'
            '.docx'
            '.epub'
            '.html'
            '.md'
            '.odp'
            '.odt'
            '.pdf'
            '.ppt'
            '.pptx'
            '.rtf'
            '.txt'
            '.xps'
        )
        Fonts = @(
            '.eot'
            '.otf'
            '.ttf'
            '.woff'
            '.woff2'
        )
        Installers = @(
            '.appx'
            '.appxbundle'
            '.exe'
            '.msi'
            '.msix'
            '.msixbundle'
            '.msu'
        )
        LinuxPkgs = @(
            '.appimage'
            '.deb'
            '.flatpak'
            '.rpm'
            '.run'
            '.snap'
        )
        macOSPkgs = @(
            '.dmg'
            '.pkg'
        )
        Photos = @(
            '.arw'
            '.bmp'
            '.cr2'
            '.dng'
            '.gif'
            '.heic'
            '.ico'
            '.jpeg'
            '.jpg'
            '.nef'
            '.orf'
            '.png'
            '.psd'
            '.raw'
            '.rw2'
            '.svg'
            '.tif'
            '.tiff'
            '.webp'
        )
        Scripts = @(
            '.ahk'
            '.bat'
            '.cmd'
            '.pl'
            '.ps1'
            '.psd1'
            '.psm1'
            '.py'
            '.reg'
            '.sh'
            '.sql'
            '.vbs'
            '.wsf'
        )
        Video = @(
            '.3gp'
            '.avi'
            '.flv'
            '.m4v'
            '.mkv'
            '.mov'
            '.mp4'
            '.mpeg'
            '.mpg'
            '.ogv'
            '.webm'
            '.wmv'
        )
    }

    $ExtensionMap = $defaultExtensionMap
}
else
{
    $testJsonParams = @{
        LiteralPath = $ExtensionMapJson
        Options     = @('AllowTrailingCommas','IgnoreComments')
        ErrorAction = 'Stop'
    }
    if (-not (Test-Json @testJsonParams))
    { throw "Invalid JSON in extension map file: '$ExtensionMapJson'" }

    $ExtensionMap = Get-Content -LiteralPath $ExtensionMapJson -Raw | ConvertFrom-Json -AsHashtable
}

#endregion Extension Map

#region Functions

function Write-Log
{
    param(
        [Parameter(Mandatory)]
        [string] $Message,

        [ValidateSet('INFO','VERBOSE','WARN','ERROR')]
        [string] $Level = 'INFO'
    )

    if (-not $script:Quiet)
    {
        switch ($Level)
        {
            'INFO'    { Write-Information $Message }
            'VERBOSE' { Write-Verbose     $Message }
            'WARN'    { Write-Warning     $Message }
            'ERROR'   { Write-Error       $Message }
        }
    }

    if ($script:NoLogFile)
    { return }

    if ($Level -eq 'VERBOSE' -and
        ($script:VerbosePreference -ne 'Continue'))
    { return }

    $entry = '[{0}] [{1}] {2}' -f (Get-Date -F 's'),$Level,$Message

    $addParams = @{
        LiteralPath = $logFile
        Value       = $entry
        Encoding    = 'UTF8'
    }
    Add-Content @addParams
}

function Resolve-DestinationPath
{
    param(
        [Parameter(Mandatory)]
        [string]$Directory,

        [Parameter(Mandatory)]
        [string]$FileName
    )

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext  = [System.IO.Path]::GetExtension($FileName)
    $num  = 0
    $try  = Join-Path $Directory $FileName

    if (-not (Test-Path -LiteralPath $try))
    { return $try }

    $destPath = $null

    while (-not $destPath)
    {
        $num++
        $try = Join-Path $Directory ('{0} ({1}){2}' -f @($base,$num,$ext))

        if (-not (Test-Path -LiteralPath $try))
        { $destPath = $try }
    }

    return $destPath
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

function Restore-DownloadsRoot
{
    $validPrefixes = @(
        $ExtensionMap.Keys.ForEach({ "${CategoryFolderPrefix}$_" })
        "${CategoryFolderPrefix}${StaleFolderName}"
    )

    $categoryDirs = (
        Get-ChildItem -LiteralPath $DownloadsPath -Directory |
        Where-Object {
            $dir = $_
            $validPrefixes.Where({ $dir.Name.StartsWith($_) })
        }
    )

    if (-not $categoryDirs)
    {
        Write-Log 'Restore-DownloadsRoot: No Sort-DownloadsFolder category folders found to restore.' 'VERBOSE'
        return
    }

    foreach ($dir in $categoryDirs)
    {
        $items = Get-ChildItem -LiteralPath $dir.FullName

        foreach ($item in $items)
        {
            $destPath = Resolve-DestinationPath -Directory $DownloadsPath -FileName $item.Name

            Move-Item -LiteralPath $item.FullName -Destination $destPath -EA 'Stop'

            Write-Log "Restored '$($item.Name)' from '$($dir.Name)' to Downloads root" 'VERBOSE'
        }

        try
        {
            if (-not (Get-ChildItem -LiteralPath $dir.FullName))
            {
                Remove-Item -LiteralPath $dir.FullName -EA 'Stop'
                
                Write-Log "Removed empty category folder '$($dir.Name)'"
            }
            else
            { 
                Write-Log (@(
                    "Restore could not be fully completed."
                    "Category folder '$($dir.Name)' still has contents, left in place."
                ) -join "`n") 'WARN'
            }
        }
        catch
        { Write-Log "Failed to remove empty category folder '$($dir.Name)'." 'WARN' }
    }
}

#endregion Functions

#region Setup

if (-not $NoLogFile)
{
    if (-not (Test-Path -LiteralPath $LogDirectory))
    {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
        $newLogDir = $true
    }

    $logFile = Join-Path $LogDirectory 'Sort-DownloadsFolder.log'

    if ((Test-Path -LiteralPath $logFile) -and
        (Get-Item -LiteralPath $logFile).LastWriteTime.Date -lt (Get-Date).Date)
    {
        Clear-Content -LiteralPath $logFile
        $clearedLog = $true
    }

    Write-Log (@(
        if ($newLogDir)  { "Log directory '$LogDirectory' was created." }
        if ($clearedLog) { "It's a new day; logfile was wiped." }
        "Sort-DownloadsFolder started.`n"
        "Script version   : $($script:Version)"
        "Downloads folder : $($DownloadsPath)"
        "Large file prefix: $($LargeFlagPrefix)"
        "Sort dir. prefix : $($CategoryFolderPrefix)"
        "Stale folder name: $($StaleFolderName)"
        "Log directory    : $($LogDirectory)"
        "Custom ext. map  : $($ExtensionMapJson ? "Yes" : "No")"
        "Overrides applied: $(($ExtensionMapOverrideJson -or $scriptExtensionMapOverride) ? "Yes" : "No")"
        "Large file size  : >$($LargeFlagSizeMB)MB"
        "Min. file age    : $($MinAgeMinutes) minutes"
        "Files stale after: $($StaleDaysThreshold) days"
        "Cleanup reminder : $($CleanReminderThreshold) loose files`n"
    ) -join "`n")
}

if (-not (Test-Path -LiteralPath $DownloadsPath))
{
    Write-Log "Downloads path not found: '$($DownloadsPath)'" 'ERROR'
    throw "Downloads path not found: '$($DownloadsPath)'"
}

if ($RestoreDownloadsRoot)
{
    Write-Log (@(
        "-RestoreDownloadsRoot specified, emptying all '$CategoryFolderPrefix' folders back into Downloads root."
        "Please note, large-flagged files/folders will not have their name prefix removed."
    ) -join "`n")

    Restore-DownloadsRoot

    Write-Log "Done."
    Write-Log "`nSort-DownloadsFolder complete."

    return
}

if (($ExtensionMapOverrideJson -and
    (Test-Path -LiteralPath $ExtensionMapOverrideJson)) -or
    $scriptExtensionMapOverride)
{
    if ($ExtensionMapOverrideJson)
    {
        $testJsonParams = @{
            LiteralPath = $ExtensionMapOverrideJson
            Options     = @('AllowTrailingCommas','IgnoreComments')
            ErrorAction = 'Stop'
        }
        if (-not (Test-Json @testJsonParams))
        { throw "Invalid JSON in extension map override file: '$ExtensionMapOverrideJson'" }
    }

    $overrides = (
        $ExtensionMapOverrideJson ?
        (Get-Content $ExtensionMapOverrideJson -Raw | ConvertFrom-Json -AsHashtable) :
        $scriptExtensionMapOverride
    )

    $lookup = @{}

    foreach ($map in @($ExtensionMap,$overrides))
    {
        $map.GetEnumerator() |
        ForEach-Object {
            $c = $_.Key
            $_.Value | ForEach-Object { $lookup[$_] = $c }
        }
    }

    $ExtensionMap = @{}

    $lookup.GetEnumerator() |
    ForEach-Object { $ExtensionMap[$_.Value] += @($_.Key) }
}

#endregion Setup

#region Main

$categoryDirCache = @{}

Get-ChildItem -LiteralPath $DownloadsPath -Directory |
Where-Object { $_.Name.StartsWith($CategoryFolderPrefix) } |
ForEach-Object {
    $cleanName = $_.Name -replace '\s+\[[\d.]+ [KMGT]?B\]$'
    $categoryDirCache[$cleanName] = $_.FullName
}

if ($LargeFlagSizeMB -gt 0)
{ $flagBytes = $LargeFlagSizeMB * 1MB }

$rootFiles = Get-ChildItem -LiteralPath $DownloadsPath -File

$fileStats = @{
    Moved   = 0
    Flagged = 0
    Skipped = 0
    Errors  = 0
}

if (-not ($Skip -contains 'Files') -and
    ($rootFiles.Count -gt 0))
{
    Write-Log "Processing $($rootFiles.Count) file(s) in specified Downloads folder..."

    foreach ($file in $rootFiles)
    {
        try
        {
            if (Test-FileLocked -Path $file.FullName)
            {
                Write-Log "File is locked: '$($file.Name)'" 'VERBOSE'
                $fileStats.Skipped++
                continue
            }

            if ($MinAgeMinutes -gt 0 -and
                ($file.CreationTime -ge (Get-Date).AddMinutes(-$MinAgeMinutes)))
            {
                Write-Log "Ignoring recent download: '$($file.Name)'" 'VERBOSE'
                $fileStats.Skipped++
                continue
            }

            $currentPath = $file.FullName
            $currentName = $file.Name

            if ($LargeFlagSizeMB -gt 0 -and
                ($file.Length -ge $flagBytes) -and
                (-not $currentName.StartsWith($LargeFlagPrefix)))
            {
                $flaggedName      = $LargeFlagPrefix + $currentName     # Flag !LRG_ if over LargeFlagSizeMB
                $resolvedFlagPath = Resolve-DestinationPath -Directory $file.DirectoryName -FileName $flaggedName
                $resolvedFlagName = [System.IO.Path]::GetFileName($resolvedFlagPath)

                Rename-Item -LiteralPath $currentPath -NewName $resolvedFlagName -EA 'Stop'

                Write-Log "Flagged large file '$currentName' ($([math]::Round($file.Length / 1MB, 1))MB)" 'VERBOSE'

                $fileStats.Flagged++

                $currentName = $resolvedFlagName
                $currentPath = Join-Path $file.DirectoryName $currentName
            }

            $fileExtension = $file.Extension.ToLower()
            $targetFolder  = $null

            foreach ($folder in $ExtensionMap.Keys)
            {
                if ($ExtensionMap[$folder] -contains $fileExtension)
                {
                    $targetFolder = "${CategoryFolderPrefix}${folder}"
                    Write-Log "Matched '$($file.Name)' type: $folder" 'VERBOSE'
                    break                       # First match of extension wins
                }
            }

            if (-not $targetFolder)
            {
                if ($LargeFlagSizeMB -gt 0 -and
                    ($file.Length -ge $flagBytes))
                {
                    $fileStats.Skipped++
                    continue                            # Leave stale but large files in root
                }
                elseif ($StaleDaysThreshold -gt 0 -and
                        ($file.LastWriteTime -lt (Get-Date).AddDays(-$StaleDaysThreshold)))
                {
                    $targetFolder = "${CategoryFolderPrefix}${StaleFolderName}"    # Default name has a z after $CategoryFolderPrefix for sorting
                    
                    Write-Log "Stale unmatched file: '$($file.Name)' (modified $($file.LastWriteTime.ToString('yyyy-MM-dd')))" 'VERBOSE'
                }
                else
                {
                    Write-Log "No type match: '$($file.Name)'" 'VERBOSE'
                    $fileStats.Skipped++
                    continue
                }
            }

            if (-not $categoryDirCache.ContainsKey($targetFolder))
            {
                $newDir = Join-Path $DownloadsPath $targetFolder

                New-Item -ItemType Directory -Path $newDir -Force | Out-Null

                Write-Log "Created file category subfolder: '$targetFolder'" 'VERBOSE'

                $categoryDirCache[$targetFolder] = $newDir
            }

            $targetDir = $categoryDirCache[$targetFolder]

            $destPath = Resolve-DestinationPath -Directory $targetDir -FileName $currentName

            Move-Item -LiteralPath $currentPath -Destination $destPath -EA 'Stop'

            Write-Log "Moved '$currentName' to: '$targetFolder'"

            $fileStats.Moved++
        }
        catch
        {
            Write-Log "Error processing '$($file.Name)':`n$_" 'ERROR'
            $fileStats.Errors++
        }
    }

    Write-Log "Done."
}
elseif ($Skip -contains 'Files')
{ Write-Log "-Skip specified with 'Files', file processing skipped." }
elseif ($rootFiles.Count -eq 0)
{ Write-Log "No files to process in the root of '$DownloadsPath'." }

$rootDirs = Get-ChildItem -LiteralPath $DownloadsPath -Directory

$dirCount = $rootDirs.Count - ($rootDirs.Where({ $_.Name.StartsWith($CategoryFolderPrefix) })).Count

if ($rootFiles.Count -eq 0 -and
    $dirCount -eq 0)
{
    Write-Log (@(
        "No files or folders found to process in the root of '$DownloadsPath'."
        "Sort-DownloadsFolder complete."
    ) -join "`n")

    return
}

if (-not ($Skip -contains 'Folders') -and
    ($dirCount -gt 0))
{
    Write-Log "Processing $dirCount subfolder(s) in specified Downloads folder..."

    foreach ($dir in $rootDirs)
    {
        try
        {
            if ($dir.Name.StartsWith($LargeFlagPrefix))
            {
                Write-Log "Skipped large-flagged folder '$($dir.Name)'" 'VERBOSE'
                $fileStats.Skipped++
                continue
            }

            if ($dir.Name.StartsWith($CategoryFolderPrefix))
            {
                Write-Log "Skipped file category folder '$($dir.Name)'" 'VERBOSE'
                continue
            }

            if ($StaleDaysThreshold -eq 0 -or
                ($dir.LastWriteTime -ge (Get-Date).AddDays(-$StaleDaysThreshold)))
            {
                Write-Log "Folder not stale: '$($dir.Name)'" 'VERBOSE'
                continue
            }

            $dirSizeBytes = (
                Get-ChildItem -LiteralPath $dir.FullName -Recurse -File |
                Measure-Object -Property Length -Sum
            ).Sum ?? 0

            Write-Log "Folder size (bytes): $dirSizeBytes" 'VERBOSE'

            if ($LargeFlagSizeMB -gt 0 -and
                ($dirSizeBytes -ge $flagBytes))
            {
                $flaggedName      = $LargeFlagPrefix + $dir.Name
                $resolvedFlagPath = Resolve-DestinationPath -Directory $DownloadsPath -FileName $flaggedName
                $resolvedFlagName = [System.IO.Path]::GetFileName($resolvedFlagPath)

                Rename-Item -LiteralPath $dir.FullName -NewName $resolvedFlagName -EA 'Stop'

                Write-Log "Flagged large folder '$($dir.Name)' ($([math]::Round($dirSizeBytes / 1MB, 1))MB)" 'VERBOSE'

                $fileStats.Flagged++
                $fileStats.Skipped++

                continue
            }

            $staleFolderKey = "${CategoryFolderPrefix}${StaleFolderName}"

            if (-not $categoryDirCache.ContainsKey($staleFolderKey))
            {
                $newDir = Join-Path $DownloadsPath $staleFolderKey

                New-Item -ItemType Directory -Path $newDir -Force | Out-Null

                Write-Log "Created stale unmatched item subfolder: '$staleFolderKey'" 'VERBOSE'

                $categoryDirCache[$staleFolderKey] = $newDir
            }

            $staleDir = $categoryDirCache[$staleFolderKey]

            $destPath = Resolve-DestinationPath -Directory $staleDir -FileName $dir.Name

            Move-Item -LiteralPath $dir.FullName -Destination $destPath -EA 'Stop'

            Write-Log "Moved stale folder '$($dir.Name)' to: '$StaleFolderName'"

            $fileStats.Moved++
        }
        catch
        {
            Write-Log "Error processing folder '$($dir.Name)':`n$_" 'ERROR'
            $fileStats.Errors++
        }
    }

    Write-Log "Done."
}
elseif ($Skip -contains 'Folders')
{ Write-Log "-Skip specified with 'Folders', folder processing skipped." }
elseif ($dirCount -eq 0)
{ Write-Log "No folders to process in the root of '$DownloadsPath'." }

if ($CleanReminderThreshold -gt 0 -and
    ($fileStats.Skipped -ge $CleanReminderThreshold))
{
    $cleanupReminder = "!!_OVER $CleanReminderThreshold LOOSE ITEMS - YOU SHOULD CLEAN UP_!!.txt"

    if (-not (Test-Path -LiteralPath "$DownloadsPath\$cleanupReminder"))
    { New-Item -Path $DownloadsPath -Name $cleanupReminder -ItemType 'File' -EA 'SilentlyContinue' }
}

$validPrefixes = @(
    $ExtensionMap.Keys.ForEach({ "${CategoryFolderPrefix}$_" })
    "${CategoryFolderPrefix}${StaleFolderName}"
)

$categoryFolders = (
    Get-ChildItem -LiteralPath $DownloadsPath -Directory |
    Where-Object {
        $folder = $_
        $validPrefixes.Where({ $folder.Name.StartsWith($_) })
    }
)

if ($categoryFolders -and
    (-not $NoSizeLabels))
{
    Write-Log "Calculating and applying size tags to category folders..."

    $folderSizes = (
        $categoryFolders |
            ForEach-Object -Parallel {
                $bytes = (
                    Get-ChildItem -LiteralPath $_.FullName -Recurse -File |
                    Measure-Object -Property Length -Sum
                ).Sum ?? 0

                [PSCustomObject]@{
                    Path  = $_.FullName
                    Name  = $_.Name
                    Bytes = $bytes
                }
        } -ThrottleLimit 8
    )

    foreach ($entry in $folderSizes)
    {
        $baseName = $entry.Name -replace '\s+\[[\d.]+ [KMGT]?B\]$'

        $sizeStr = switch ($entry.Bytes)
        {
            { $_ -ge 1GB } { '{0:N1} GB' -f ($entry.Bytes / 1GB); break }
            { $_ -ge 1MB } { '{0:N1} MB' -f ($entry.Bytes / 1MB); break }
            { $_ -ge 1KB } { '{0:N1} KB' -f ($entry.Bytes / 1KB); break }
            default        { '{0} B'     -f  $entry.Bytes }
        }

        $newName = '{0} [{1}]' -f $baseName, $sizeStr

        if ($newName -ne $entry.Name)
        {
            Rename-Item -LiteralPath $entry.Path -NewName $newName -EA 'Stop'

            Write-Log "Applied size label to '$baseName': $sizeStr" 'VERBOSE'

            $categoryDirCache[$baseName] = Join-Path $DownloadsPath $newName
        }
    }

    Write-Log "Done."
}

Write-Log (@(
    "`nSummary:"
    "Moved  : $($fileStats.Moved)"
    "Flagged: $($fileStats.Flagged)"
    "Skipped: $($fileStats.Skipped)"
    "Errors : $($fileStats.Errors)`n"
) -join "`n")

Write-Log "Sort-DownloadsFolder complete."

#endregion Main
