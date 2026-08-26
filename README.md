# Purview Activity Explorer — Full Export Automation

Automates a **complete** export of Microsoft Purview Activity Explorer data — bypassing the
10,000-row limit of the Activity Explorer portal UI — by calling the underlying
`Export-ActivityExplorerData` PowerShell cmdlet directly and paging through every result.

Produces one dated folder per exported date, each containing a full-fidelity JSON backup and a
column-normalized, Excel-ready CSV, plus a combined CSV across the whole run.

## Why this exists

The Activity Explorer web portal caps any single export at 10,000 rows. Tenants generating more
than that per period can't get a complete export through the UI, and doing this by hand every
month doesn't scale. This script calls the cmdlet behind that button directly — which has no
row cap, only pagination — and loops until every record for the requested period is retrieved.

## Contents

| File | Purpose |
|---|---|
| `Get & Export-ActivityExplorerFull.ps1` | The export script. |
| `Activity Explorer Export - User Guide.docx` | Full user guide: architecture, every parameter, every filter value, troubleshooting, official sources. |

## Requirements

- PowerShell 5.1 or PowerShell 7+, Windows
- `ExchangeOnlineManagement` module: `Install-Module ExchangeOnlineManagement`
- An account with Activity Explorer / Purview permissions (e.g. Information Protection Analysts
  or Compliance Administrator role group)

## Running this in PowerShell (step-by-step)

If you've never run this script before, here's the full walkthrough from download to output:

1. **Download the script.** Save `Export-ActivityExplorerFull.ps1` anywhere on your computer —
   for example your Desktop, or a dedicated folder like `C:\Scripts\ActivityExplorer\`.
2. **Open PowerShell.** Click Start, type `PowerShell`, and open it (Windows PowerShell or
   PowerShell 7 both work).
3. **Move into the folder where you saved the script**, using `cd` (change directory) followed
   by the full path. For example, if you saved it to `C:\Scripts\ActivityExplorer\`:
   ```powershell
   cd C:\Scripts\ActivityExplorer
   ```
   Tip: you don't have to type the path by hand — in File Explorer, open that folder, click in
   the address bar, copy the path, then paste it after `cd ` in PowerShell.
4. **Confirm the script is there** (optional, but useful the first time):
   ```powershell
   dir
   ```
   You should see `Export-ActivityExplorerFull.ps1` listed.
5. **Run the export command.** This is the same command from Quick Start above — run it from
   inside the folder you just moved into:
   ```powershell
   .\Export-ActivityExplorerFull.ps1  
   ```  
   Along with the desired filter  of activity, workload, date 
   The leading `.\` tells PowerShell "run the script sitting right here in this folder."
6. **Sign in when prompted.** A login window will appear — sign in with an account that has
   Activity Explorer/Purview permissions, completing MFA if required.
7. **Let it run.** You'll see console output per date as it pages through results (`Page 1: 5000
   records...`, etc.). This can take a while for busy tenants — that's expected.
8. **Check the output.** Once it finishes, a new folder (`ActivityExplorerExport` by default)
   appears inside the same folder you ran the command from, containing one subfolder per date
   plus the combined CSV.

> If you get a message about scripts being disabled ("running scripts is disabled on this
> system"), run this once first, then try step 5 again:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```
> This only relaxes the restriction for the current PowerShell window, not system-wide.

## Quick start

```powershell
# Simplest run — interactive sign-in, last 29 days, one folder per date
.\Export-ActivityExplorerFull.ps1 -UserPrincipalName admin@yourtenant.onmicrosoft.com

# Custom output path, shorter window, already connected in this session
.\Export-ActivityExplorerFull.ps1 -PastDays 7 -OutputRootPath "D:\PurviewExports" -SkipConnect
```

## Parameters

| Parameter | Type | Default | Purpose |
|---|---|---|---|
| `-UserPrincipalName` | string | *(none)* | UPN used to sign in via `Connect-IPPSSession`. |
| `-SkipConnect` | switch | off | Skip connecting — use when already connected in this session. |
| `-OutputRootPath` | string | `.\ActivityExplorerExport` | Root folder; one dated subfolder per exported date. |
| `-PageSize` | int (1–5000) | 5000 | Records requested per page. Lower only if you see PageCookie-expiry errors. |
| `-PastDays` | int | 29 | Trailing window ending yesterday. Ignored if `-SpecificDates` or `-StartDate`/`-EndDate` is supplied. |
| `-StartDate` / `-EndDate` | date | *(none)* | Explicit contiguous date range (inclusive). Use both together. |
| `-SpecificDates` | date[] | *(none)* | Explicit, non-contiguous list of individual dates. Highest priority of the three date modes. |
| `-Activities` | string[] | *(none)* | Filter to specific activity types (OR'd together). Full list below. |
| `-Workloads` | string[] | *(none)* | Filter to specific workloads (OR'd together). Full list below. |

## Date selection — three modes, checked in this order

```powershell
# 1. Specific, non-contiguous dates (highest priority)
.\Export-ActivityExplorerFull.ps1 -SpecificDates "2026-08-05","2026-08-12","2026-08-19"

# 2. Explicit contiguous range
.\Export-ActivityExplorerFull.ps1 -StartDate "2026-08-01" -EndDate "2026-08-07"

# 3. Trailing window (default if neither above is given)
.\Export-ActivityExplorerFull.ps1 -PastDays 14
```

Activity Explorer only reliably retains **30 days** of data, and the service checks that boundary
against a precise UTC timestamp rather than a calendar date — so the default is 29 days plus an
internal safety buffer, not the full 30. Any date landing outside the safe window, or that is
today/future, is skipped automatically with a console warning.

## Activity filter (`-Activities`)

```powershell
.\Export-ActivityExplorerFull.ps1 -Activities "FileArchived","LabelApplied"
```

Valid values (case-sensitive):

```
AIAppInteraction, ArchiveCreated, AutoLabelingSimulation, ChangeProtection,
ClassificationAdded, ClassificationDeleted, ClassificationUpdated, CopilotInteraction,
DLPInfo, DLPRuleEnforce, DLPRuleMatch, DLPRuleUndo, DlpClassification, DownloadFile,
DownloadText, FileAccessedByUnallowedApp, FileArchived, FileCopiedToClipboard,
FileCopiedToNetworkShare, FileCopiedToRemoteDesktopSession, FileCopiedToRemovableMedia,
FileCreated, FileCreatedOnNetworkShare, FileCreatedOnRemovableMedia, FileDeleted,
FileDiscovered, FileModified, FilePrinted, FileRead, FileRenamed,
FileTransferredByBluetooth, FileUploadedToCloud, LabelApplied, LabelChanged,
LabelRecommended, LabelRecommendedAndDismissed, LabelRemoved, NewProtection,
PastedToBrowser, RemoveProtection, ScreenCapture, UploadFile, UploadText,
WebpageCopiedToClipboard, WebpagePrinted, WebpageSavedToLocal
```

## Workload filter (`-Workloads`)

```powershell
.\Export-ActivityExplorerFull.ps1 -Workloads "Exchange","SharePoint"
```

Valid values:

```
Copilot, Endpoint, Exchange, OnPremisesFileShareScanner, OnPremisesSharePointScanner,
OneDrive, PowerBI, PurviewDataMap, SharePoint
```

## Combining filters with dates

Filters and date-selection are independent and can always be combined:

```powershell
# Filter + date range
.\Export-ActivityExplorerFull.ps1 -Activities "FileArchived","LabelApplied" -StartDate "2026-08-01" -EndDate "2026-08-07"

# Both filters + a date range
.\Export-ActivityExplorerFull.ps1 -Activities "FileArchived","LabelApplied" -Workloads "SharePoint" -StartDate "2026-08-01" -EndDate "2026-08-07"
```

`-Activities` and `-Workloads` combine with **AND**; multiple values within one filter combine
with **OR** — e.g. `(Activity = FileArchived OR LabelApplied) AND (Workload = SharePoint)`.

## Output structure

```
ActivityExplorerExport\
  2026-08-05\
    2026-08-05-ActivityExplorer.csv        <- normalized, all columns, Excel-ready
    2026-08-05-ActivityExplorer-Raw.json   <- full-fidelity backup
  2026-08-12\
    2026-08-12-ActivityExplorer.csv
    2026-08-12-ActivityExplorer-Raw.json
  Combined-2-Dates-20260805-to-20260812.csv <- all dates in one file
```

Every CSV is normalized against the same fixed ~76-column schema, so columns stay identical and
in order across every day and the combined file, regardless of which activity types appear on a
given date. Missing fields are left blank, never omitted.



| Symptom | Cause | Fix |
|---|---|---|
| `Date range must be within the past 30 days and cannot include future dates` | Day boundaries built from local time instead of UTC, or grazing the 30-day/now edge. | Already handled by UTC-based windowing + safety buffer. If it recurs, reduce `-PastDays` or narrow `-StartDate`/`-EndDate`. |
| `Invalid Filter Name` | `-Filter2` sent with no `-Filter1` (e.g. only `-Workloads` supplied). | Already handled — filters are packed sequentially starting at `-Filter1` with no gaps. |
| PageCookie / paging errors on high-volume days | 120-second PageCookie validity exceeded. | Lower `-PageSize`, or check network latency to the compliance endpoint. |


## Official Microsoft source

- [Export-ActivityExplorerData (ExchangePowerShell) — Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/export-activityexplorerdata?view=exchange-ps)
