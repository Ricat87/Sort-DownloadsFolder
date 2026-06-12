# Sort-DownloadsFolder

Keeps your Downloads folder from turning into a disaster. Sorts files into type-based
subfolders, flags large items so they float to the top, and moves stale clutter to a `_Stale`
folder automatically. Set it up once as a scheduled task and forget about it.

## Features

- Sorts files into named subfolders (`!Installers`, `!Scripts`, `!Documents`, etc.)
- Flags files and folders at or above a size threshold with `!LRG_` so they sort to the top
- Moves unrecognized stale files and old subfolders to `_Stale`
- Skips files that are locked (mid-download) or too recently created
- Drops a reminder file in the root when too many items pile up
- Fully configurable via parameters -- sensible defaults out of the box
- Built-in extension map with 16 categories, overridable without editing the script

## Quick Start

```powershell
# Run once manually to see what it does
.\Sort-DownloadsFolder.ps1 -Verbose

# Target a specific folder
.\Sort-DownloadsFolder.ps1 -DownloadsPath 'D:\Downloads' -Verbose

# Sort files only, skip subfolders
.\Sort-DownloadsFolder.ps1 -Skip Folders

# Lower the large-file threshold and shorten the stale window
.\Sort-DownloadsFolder.ps1 -LargeFlagSizeMB 250 -StaleDaysThreshold 3
```

---

## Installation

1. Requires **PowerShell 7.4+**. This will not run on Windows PowerShell 5.1.
   ```
   winget install Microsoft.PowerShell
   ```

2. Put the script somewhere stable. Avoid putting it in Downloads itself since the script
   directory is used for logging and the script file would get processed on every run.

3. Unblock it (Windows marks downloaded files and PowerShell may refuse to run them):
   ```powershell
   Unblock-File -Path 'C:\Scripts\Sort-DownloadsFolder\Sort-DownloadsFolder.ps1'
   ```

4. Check your execution policy:
   ```powershell
   Get-ExecutionPolicy -Scope CurrentUser
   ```
   If it's `Restricted` or `AllSigned`, fix it:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```
   No elevation needed. Don't touch `LocalMachine` scope unless you have a reason to.

---

## Scheduled Task Setup

Run this once from a PowerShell 7 session. No elevation required.

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

Make sure `pwsh.exe` is on your PATH before running this. The task calls `pwsh.exe` directly,
not `powershell.exe`. You can verify with `Get-Command pwsh`.

---

## How It Works

Every run processes loose files and subfolders in your Downloads root:

**Files**
- Left alone if locked (mid-download) or too recently created (within `MinAgeMinutes`, default 30).
- Renamed with `!LRG_` if at or above `LargeFlagSizeMB` (default 100 MB) so they float to
  the top when sorted by name.
- Moved into a type-matched subfolder like `!Installers`, `!Scripts`, `!Documents`, etc.
  The `!` prefix keeps them sorted above everything else.
- Unrecognized extensions older than `StaleDaysThreshold` days (default 7) go to `_Stale`.
  Large unmatched files get flagged and stay put.

**Subfolders**
- Skipped if the name starts with `!` or matches `StaleFolderName`.
- Flagged with `!LRG_` if large, moved to `_Stale` if stale and not large.

**Cleanup reminder**
Once skipped file count hits `CleanReminderThreshold` (default 15), a `.txt` reminder file
appears in the root.

**Logging**
Single log file in the script directory, cleared daily. VERBOSE lines only appear in the log
if you run with `-Verbose`.

---

## Configuration

Everything is a parameter with a sensible default. Most people won't need to change anything.

| Parameter | Default | Description |
|---|---|---|
| `DownloadsPath` | Current user's Downloads | Folder to process |
| `LargeFlagPrefix` | `!LRG_` | Prefix for large files and folders |
| `StaleFolderName` | `_Stale` | Destination for stale unmatched items |
| `LogDirectory` | Script directory | Where the log file goes |
| `LargeFlagSizeMB` | `100` | Flag threshold in MB. `0` disables |
| `MinAgeMinutes` | `30` | Leave files alone if newer than this. `0` disables |
| `StaleDaysThreshold` | `7` | Days until an unmatched item is considered stale. `0` disables |
| `CleanReminderThreshold` | `15` | Skipped file count before reminder appears. `0` disables |
| `ExtensionMapJson` | *(none)* | JSON file to fully replace the built-in extension map |
| `ExtensionMapOverrideJson` | *(none)* | JSON file to add to or override specific categories |
| `Skip` | *(none)* | Skip `Files`, `Folders`, or both: `-Skip Files,Folders` |
| `NoLog` | *(off)* | Disables log file for the run |
| `Quiet` | *(off)* | Suppresses console output |

### Extension Map

Built-in categories:

`Android` `Archives` `Audio` `Certs` `Code` `ConfigFiles` `Data` `DiskImages`
`Documents` `Fonts` `Installers` `LinuxPkgs` `macOSPkgs` `Photos` `Scripts` `Video`

**In-script override** (no extra files needed): set `$scriptExtMapOverrideEnabled = $true`
near the top of the script and edit `$scriptExtensionMapOverride`. Extensions you reassign are
automatically removed from their original category:

```powershell
$scriptExtensionMapOverride = @{
    Data = @('.csv', '.tsv')
}
```

**External override JSON** (same idea, just in a file):

```json
{ "Data": [".csv", ".tsv"] }
```

```powershell
.\Sort-DownloadsFolder.ps1 -ExtensionMapOverrideJson 'C:\Scripts\MyOverrides.json'
```

**Full replacement:** pass a complete map via `-ExtensionMapJson` to throw out the defaults
entirely. Can be combined with `-ExtensionMapOverrideJson`; the replacement loads first, then
the override is merged on top.

---

## A Few Gotchas

- **First run on a messy folder:** use `-Skip Folders` to sort files first, check the results,
  then run again to handle folders.
- **Remote use over SMB:** works fine with `-DownloadsPath '\\PC\C$\Users\name\Downloads'` if
  you have local admin on the target. Run it interactively rather than via scheduled task, and
  always pass `-DownloadsPath` explicitly since the default resolves to your local profile.
- **Name collisions:** follows Windows convention and appends `(1)`, `(2)`, etc. rather than
  overwriting.