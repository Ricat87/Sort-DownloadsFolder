# Sort-DownloadsFolder

Keeps your (my) Downloads folder from turning into a disaster.

Sorts files into extension-based category folders, flags large files/folders, consolidates stale uncategorized files,
and stamps each category folder with its current size on each run.
The plan is to set a scheduled task and forget it.

*"Your warranty is void", I take no responsibility for what you do with this, yada yada yada.*

## Features

- Sorts files into named category folders (`!!Installers`, `!!Scripts`, `!!Documents`, etc.)
- Category folders use `!!` prefix so they sort (by name) above any user-floated `!` folders
- Flags files and folders at or above 100MB with `!LRG_` so they sort to the top
- Consolidates stale (>7d) uncategorized files and stale subfolders into `!!zStale`
- Stamps each category folder with its current size on each run (e.g. `!!Installers [1.2 GB]`)
- Skips files that are locked (in use) or downloaded within the last 30 minutes
- Places a reminder to clean up in the root if too many items (15) pile up
- **Everything above is fully configurable via parameters/directly in script**
- File extension maps are configurable in script or via external JSON files
- Supports `-WhatIf`, because it should
- Built-in extension map with 16 categories and 155 file types

## Quick Start

```powershell
# Dry run. See what would happen without touching anything
.\Sort-DownloadsFolder.ps1 -WhatIf -InformationAction Continue  # Optional: -Verbose

# Manual run of the script with default settings
.\Sort-DownloadsFolder.ps1 -InformationAction Continue

# Target a specific folder
.\Sort-DownloadsFolder.ps1 -DownloadsPath 'D:\Downloads'

# Sort files only, skip folders
.\Sort-DownloadsFolder.ps1 -Skip Folders

# Lower the large-file threshold and shorten the stale window
.\Sort-DownloadsFolder.ps1 -LargeFlagSizeMB 50 -StaleDaysThreshold 3

# Move everything back out of category folders into the Downloads root
.\Sort-DownloadsFolder.ps1 -RestoreDownloadsRoot  # Only works on categories currently in the map
```

---

## Installation

**Requires PowerShell 7.4+. This will not run on Windows PowerShell 5.1.**
   ```powershell
   winget install Microsoft.PowerShell
   ```

1. Put the script somewhere stable, **NOT your Downloads folder**.
   e.g. C:\Scripts\Sort-DownloadsFolder or C:\Users\yourusername\Sort-DownloadsFolder

2. Unblock the file from Defender SmartScreen via Properties or:
   ```powershell
   Unblock-File -Path 'C:\Scripts\Sort-DownloadsFolder\Sort-DownloadsFolder.ps1'
   ```

3. Check your execution policy:
   ```powershell
   Get-ExecutionPolicy -Scope CurrentUser
   ```
   If it's `Restricted` or `AllSigned`, fix it:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```
   No elevation needed. Don't touch the `LocalMachine` scope unless you have a reason to.
   As per Microsoft, ExecutionPolicy is not a security feature, but still.

---

## Scheduled Task Setup

Run this once from a PowerShell 7 session. No elevation required.
It's a good idea to check Task Scheduler later to see that it's there and working.

```powershell
$script = 'C:\Scripts\Sort-DownloadsFolder\Sort-DownloadsFolder.ps1'

$action = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument "-NonInteractive -File `"$script`""

$triggerParams = @{
    Once               = $true
    At                 = Get-Date
    RepetitionInterval = New-TimeSpan -Minutes 30
}
$trigger = New-ScheduledTaskTrigger @triggerParams

$settingsParams = @{
    ExecutionTimeLimit        = New-TimeSpan -Minutes 5
    StartWhenAvailable        = $true
    RunOnlyIfNetworkAvailable = $false
}
$settings = New-ScheduledTaskSettingsSet @settingsParams

$taskParams = @{
    TaskName = 'Sort-DownloadsFolder'
    Action   = $action
    Trigger  = $trigger
    Settings = $settings
    RunLevel = 'Limited'
    Force    = $true
}
Register-ScheduledTask @taskParams
```

Make sure `pwsh.exe` is on your PATH before running this. The task calls `pwsh.exe` (PS7)
directly, not `powershell.exe` (PS5.1). You can verify with `Get-Command pwsh`.

---

## How It Works

Every run processes loose files and subfolders in your Downloads root:

**Files**
- Left alone if locked (in use) or too recently created (within `MinAgeMinutes`, default 30).
- Renamed with `!LRG_` (default) if at or above `LargeFlagSizeMB` (default 100 MB) so they float to
  the top when sorted by name.
- Moved into a type-matched subfolder like `!!Installers`, `!!Scripts`, `!!Documents`, etc.
  The `!!` prefix (default) keeps them sorted above user-floated `!` folders and everything else.
- Unmatched files older than `StaleDaysThreshold` (default 7) are moved into `!!zStale` (default).
  Large stale files stay in the root.
- Disable with `-Skip Files`, if you want to?

**Folders**
- Known (in current map) category folders are skipped.
- Flagged with `!LRG_` (default) if large, moved to `!!zStale` (default) if stale and not large.
- Disable with `-Skip Folders`.

**Category Folder Sizes**
- During each run, all category folders have their current total size appended, e.g.
  `!!Installers [1.2 GB]`.
- Disable with `-NoSizeLabels`.

**Cleanup Reminder**
- Once the skipped file/folder count hits `CleanReminderThreshold` (default 15), a `.txt` reminder file
  appears in the root.
- Disable with `-CleanReminderThreshold`.

**Logging**
- The script creates a single log file in the script directory and clears it daily.
- Logging to console is done via `Write-Information` and will not appear by default.
- Use `-InformationPreference Continue` or set `$InformationPreference Continue` to see it.
- `VERBOSE` lines only appear in the log if you run the script with `-Verbose`.

**Restore**
- `-RestoreDownloadsRoot` moves the contents of all (in current map) category folders back into
  the Downloads root, then deletes each folder if it was successfully emptied.
  Large file prefixes are not removed from file/folder names.

---

## Configuration

Everything is a parameter, with a default set to what made sense for myself.
Editing the defaults directly is easier than using parameters in your scheduled task.
Setting integer parameters to 0 disables that feature of the script.
Extremely high values are blocked by parameter validation, remove it if you need to.

| Parameter | Default | Description |
|---|---|---|
| `DownloadsPath` | Current user's Downloads | Folder to process |
| `LargeFlagPrefix` | `!LRG_` | Prefix for large files and folders |
| `CategoryFolderPrefix` | `!!` | Prefix for script-managed category subfolders |
| `StaleFolderName` | `zStale` | Stale folder name, gets same prefix as category folders |
| `LogDirectory` | Script directory | Where the log file goes |
| `LargeFlagSizeMB` | `100` | Large file flag threshold in MB. `0` disables |
| `MinAgeMinutes` | `30` | Don't touch downloaded files newer than this. `0` disables |
| `StaleDaysThreshold` | `7` | Days until an unmatched file/folder is considered stale. `0` disables |
| `CleanReminderThreshold` | `15` | This many skipped items will create a cleanup reminder file. `0` disables |
| `ExtensionMapJson` | *(none)* | JSON file to fully replace the built-in extension map |
| `ExtensionMapOverrideJson` | *(none)* | JSON file to add to or override specific categories/extensions |
| `Skip` | *(none)* | Skip `Files`, `Folders`, or both: `-Skip Files,Folders` (*Why though*) |
| `RestoreDownloadsRoot` | *(off)* | Undo everything this script has done |
| `NoSizeLabels` | *(off)* | Skip stamping category folders with their current size |
| `NoLogFile` | *(off)* | Disables logging to file |
| `Quiet` | *(off)* | Suppresses all console output from the script itself (except terminating errors) |

### Extension Map

Built-in categories:

`Android` `Archives` `Audio` `Certs` `Code` `ConfigFiles` `Data` `DiskImages`
`Documents` `Fonts` `Installers` `LinuxPkgs` `macOSPkgs` `Photos` `Scripts` `Video`

**In-script override** (no extra files needed): set `$scriptExtMapOverrideEnabled = $true`
in `#region In-Script Map Override` in the script and edit `$scriptExtensionMapOverride`.
You can add and modify categories in either the default map or an `-ExtensionMapJson` one.
Extensions you reassign here are overriden from the main map.

```powershell
$scriptExtensionMapOverride = @{
    CSV = @(
        '.csv'
        '.tsv'
    )
}
```

**External override JSON** (same idea, just in a file):

```json
{ "CSV": [".csv", ".tsv"] }
```

```powershell
.\Sort-DownloadsFolder.ps1 -ExtensionMapOverrideJson 'C:\Scripts\MyOverrides.json'
```

---

## Other Notes

- **Always use -WhatIf the first time.**
- **Remote use over SMB:** works fine with `-DownloadsPath '\\PC\C$\Users\name\Downloads'` if
  you have local admin rights on the target.
- **Name collisions:** follows Windows convention and appends `(1)`, `(2)`, etc. rather than
  overwriting.
- **Other uses:** Technically, you could point this at anything. **Do this at your own risk.**
