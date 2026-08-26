<#
.SYNOPSIS
    Fully exports Microsoft Purview Activity Explorer data for the past N days (default 30),
    one dated folder per day, bypassing the 10,000-row portal UI download limit by calling
    Export-ActivityExplorerData directly and paging through ALL results per day.

.DESCRIPTION
    Design notes (why it's built this way):

    1. THE 10,000-ROW LIMIT DOES NOT APPLY HERE.
       That cap belongs to the "Export" button in the Activity Explorer web UI. This script
       never uses the UI — it calls the Export-ActivityExplorerData cmdlet directly, which is
       paginated via -PageSize (max 5000) and -PageCookie/Watermark. We loop per day until
       $res.LastPage -eq $true, so every record for that day is retrieved, regardless of volume.

    2. ONE DAY AT A TIME.
       Microsoft's own documentation recommends smaller StartTime/EndTime windows over one big
       range to avoid timeouts. This also happens to match your requirement of "each different
       folder for each date," so daily windowing solves both problems at once.

    3. PAGECOOKIE EXPIRES IN 120 SECONDS.
       We do the minimum work between page calls (just accumulate raw JSON) and defer all
       parsing/flattening/column-normalization until after a day's pages are fully retrieved.

    4. NOT ALL ACTIVITIES SHARE THE SAME COLUMNS.
       Per the Export-ActivityExplorerData reference, different activity types populate
       different subsets of ~76 possible columns. A naive per-page CSV merge produces ragged,
       inconsistent columns. This script normalizes every record against the FULL documented
       column list (see $AllColumns below) so every day's CSV — and the combined monthly CSV —
       has identical, complete columns. Missing fields are left blank, never omitted.

    5. 30-DAY ROLLING WINDOW.
       Activity Explorer itself only retains 30 days of data. Schedule this script to run at
       least monthly (weekly is safer) so you never lose data to the rolling window — see the
       scheduling note at the bottom of this file.

.PARAMETER Activities
    Optional. One or more Activity Explorer activity names to filter to (OR'd together),
    e.g. -Activities "FileArchived","LabelApplied". Valid values include: AIAppInteraction,
    ArchiveCreated, AutoLabelingSimulation, ChangeProtection, ClassificationAdded,
    ClassificationDeleted, ClassificationUpdated, CopilotInteraction, DLPInfo, DLPRuleEnforce,
    DLPRuleMatch, DLPRuleUndo, DlpClassification, DownloadFile, DownloadText,
    FileAccessedByUnallowedApp, FileArchived, FileCopiedToClipboard, FileCopiedToNetworkShare,
    FileCopiedToRemoteDesktopSession, FileCopiedToRemovableMedia, FileCreated,
    FileCreatedOnNetworkShare, FileCreatedOnRemovableMedia, FileDeleted, FileDiscovered,
    FileModified, FilePrinted, FileRead, FileRenamed, FileTransferredByBluetooth,
    FileUploadedToCloud, LabelApplied, LabelChanged, LabelRecommended,
    LabelRecommendedAndDismissed, LabelRemoved, NewProtection, PastedToBrowser,
    RemoveProtection, ScreenCapture, UploadFile, UploadText, WebpageCopiedToClipboard,
    WebpagePrinted, WebpageSavedToLocal. Omit to export all activities.

.PARAMETER Workloads
    Optional. One or more workloads to filter to (OR'd together), e.g. -Workloads
    "Exchange","SharePoint". Valid values: Copilot, Endpoint, Exchange,
    OnPremisesFileShareScanner, OnPremisesSharePointScanner, OneDrive, PowerBI,
    PurviewDataMap, SharePoint. Combined with -Activities using AND (if both are given).
    Omit to export all workloads.

.PARAMETER StartDate
    Optional. Explicit start of a custom date range (date only, e.g. "2026-08-01"), inclusive.
    Use together with -EndDate to export a specific contiguous range instead of the trailing
    -PastDays window. Both must fall within the service's ~30-day retention.

.PARAMETER EndDate
    Optional. Explicit end of a custom date range (date only), inclusive. Use with -StartDate.

.PARAMETER SpecificDates
    Optional. An explicit, non-contiguous list of individual dates to export, e.g.
    -SpecificDates "2026-08-05","2026-08-12","2026-08-19". Takes priority over -StartDate/
    -EndDate and -PastDays if supplied.

.PARAMETER PastDays
    How many days back to export, ending yesterday (UTC). Default 29 — Activity Explorer's
    documented retention is "up to 30 days," but the service enforces that boundary against a
    precise UTC timestamp, not a calendar date, so requesting the full 30 tends to graze the
    edge and get rejected. 29 plus the script's own safety buffer stays reliably inside it.
    Each day becomes its own subfolder. Ignored if -SpecificDates or -StartDate/-EndDate is used.

.PARAMETER OutputRootPath
    Root folder under which dated subfolders (yyyy-MM-dd) are created. Default: .\ActivityExplorerExport

.PARAMETER PageSize
    Records per page requested from the cmdlet. Max 5000 (Microsoft's own ceiling). Default 5000.
    Lower this (e.g. 1000-2000) only if you see PageCookie-expiry errors on a very high-volume tenant.

.PARAMETER UserPrincipalName
    UPN used to connect to Security & Compliance PowerShell (Connect-IPPSSession). If you're
    already connected in this session, omit this and the script will skip reconnecting.

.PARAMETER SkipConnect
    Use this switch if you already have an active Connect-IPPSSession in the current session.

.EXAMPLE
    .\Export-ActivityExplorerFull.ps1 -UserPrincipalName admin@ansisolutions.onmicrosoft.com

.EXAMPLE
    .\Export-ActivityExplorerFull.ps1 -PastDays 7 -OutputRootPath "D:\PurviewExports" -SkipConnect

.EXAMPLE
    # Only FileArchived / LabelApplied activities, last 29 days
    .\Export-ActivityExplorerFull.ps1 -Activities "FileArchived","LabelApplied"

.EXAMPLE
    # Only SharePoint workload, specific contiguous date range
    .\Export-ActivityExplorerFull.ps1 -Workloads "SharePoint" -StartDate "2026-08-01" -EndDate "2026-08-07"

.EXAMPLE
    # Just three specific, non-contiguous dates
    .\Export-ActivityExplorerFull.ps1 -SpecificDates "2026-08-05","2026-08-12","2026-08-19"

.NOTES
    Requires: ExchangeOnlineManagement module (for Connect-IPPSSession / Export-ActivityExplorerData).
    Requires: An account with Activity Explorer / Purview permissions (e.g. Information Protection
    Analysts or Compliance Administrator role group).

    << PLACEHOLDER FOR HUMAN COMPLETION >>
    - Confirm $UserPrincipalName / connection method matches your tenant's auth model (interactive
      MFA vs certificate-based app-only auth). This script uses interactive Connect-IPPSSession by
      default — replace that block if you're wiring this into unattended Azure Automation.

      Use this for specefic date and filter 
      .\Get-Activity Explorer Data Last 30 Days With Filter Workload and Start End Date.ps1'-Activities "FileRenamed" -SpecificDates "2026-08-08","2026-08-13","2026-08-21"
#>

[CmdletBinding()]
param(
    [int]$PastDays = 29,
    [string]$OutputRootPath = ".\ActivityExplorerExport",
    [ValidateRange(1,5000)]
    [int]$PageSize = 5000,
    [string]$UserPrincipalName,
    [switch]$SkipConnect,
    [string[]]$Activities,
    [string[]]$Workloads,
    [datetime]$StartDate,
    [datetime]$EndDate,
    [datetime[]]$SpecificDates
)

# ---------------------------------------------------------------------------
# Full documented column list from the Export-ActivityExplorerData reference.
# Every exported row is normalized to exactly this column set (order preserved)
# so CSVs are consistent regardless of which activity types appear on a given day.
# << PLACEHOLDER: if Microsoft adds new columns in future, add them here. >>
# ---------------------------------------------------------------------------
$AllColumns = @(
    "Activity","Application","ArtifactType","AssociatedAdminUnits","AuthorizedGroupId",
    "AuthorizedGroupName","ClientIP","DataState","DestinationLocationType","DeviceName",
    "DlpPolicyMatchId","EndpointOperation","EnforcementMode","EntityProperties","EvaluationTime",
    "FalsePositive","FileExtension","FilePath","FileSize","FileType","FullUrl","GroupId",
    "GroupName","GroupType","Happened","Hidden","HowApplied","HowAppliedDetail","IRMContentId",
    "IsCorporateNetwork","IsProtected","IsProtectedBefore","JitTriggered","Justification",
    "LabelEventType","Manufacturer","MatchedWithV1DetailedScheme","MDATPDeviceId","Model",
    "OldRetentionLabel","OldSensitivityLabel","OriginatingDomain","ParentArchiveHash","Platform",
    "PolicyId","PolicyMode","PolicyName","PreviousFileName","PreviousFilePath",
    "PreviousProtectionOwner","ProcessName","ProductVersion","ProtectionEventType",
    "ProtectionOwner","ProtectionType","Reason","Receivers","RecordIdentity","RetentionLabel",
    "RMSEncrypted","RuleActions","RuleId","RuleName","Sender","SensitiveInfoTypeBucketsData",
    "SensitiveInfoTypeData","SensitivityLabel","SensitivityLabelPolicyId","SerialNumber","Sha1",
    "Sha256","SourceLocationType","StorageName","Subject","TargetDomain","TargetFilePath",
    "TargetPrinterName","TemplateId","User","UserSku","UserType","VpnNetworkAddress",
    "VpnServerAddress","Workload"
)

function Connect-IfNeeded {
    if ($SkipConnect) { return }
    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        Write-Host "Already connected to Security & Compliance PowerShell." -ForegroundColor Green
    }
    catch {
        Write-Host "Connecting to Security & Compliance PowerShell..." -ForegroundColor Cyan
        if ($UserPrincipalName) {
            Connect-IPPSSession -UserPrincipalName $UserPrincipalName
        }
        else {
            # << PLACEHOLDER: supply -UserPrincipalName, or hardcode your admin UPN here >>
            Connect-IPPSSession
        }
    }
}

function Get-FullDayActivityData {
    param(
        [datetime]$DayStartUtc,
        [datetime]$DayEndUtc,
        [int]$PageSize,
        [string[]]$Activities,
        [string[]]$Workloads
    )

    $allRecords = New-Object System.Collections.Generic.List[object]
    $pageCookie = $null
    $pageNumber = 0
    $lastPage = $false

    while (-not $lastPage) {
        $pageNumber++
        $params = @{
            StartTime    = $DayStartUtc
            EndTime      = $DayEndUtc
            PageSize     = $PageSize
            OutputFormat = "Json"
        }
        if ($pageCookie) { $params["PageCookie"] = $pageCookie }

        # Activity/Workload filters must be assigned sequentially into Filter1, Filter2, ...
        # -Filter2 (or higher) is REJECTED by the service as "Invalid Filter Name" if the
        # slots before it are empty — e.g. sending only -Filter2 with no -Filter1 fails,
        # even though conceptually you only wanted one filter. So we build whichever
        # filters were actually supplied and pack them starting at Filter1 with no gaps.
        $filterList = @()
        if ($Activities -and $Activities.Count -gt 0) { $filterList += ,(@("Activity") + $Activities) }
        if ($Workloads -and $Workloads.Count -gt 0)    { $filterList += ,(@("Workload") + $Workloads) }
        for ($f = 0; $f -lt $filterList.Count; $f++) {
            $params["Filter$($f + 1)"] = $filterList[$f]
        }

        try {
            $res = Export-ActivityExplorerData @params
        }
        catch {
            Write-Warning "  Page $pageNumber failed: $($_.Exception.Message). Retrying once..."
            Start-Sleep -Seconds 3
            $res = Export-ActivityExplorerData @params
        }

        if ($res.ResultData) {
            # ResultData is a JSON string of the page's records — parse immediately,
            # do it fast, then move on before the 120-second PageCookie window closes.
            $parsed = $res.ResultData | ConvertFrom-Json
            foreach ($record in $parsed) { $allRecords.Add($record) }
            Write-Host "  Page $pageNumber : $($parsed.Count) records (running total: $($allRecords.Count))"
        }

        $lastPage   = $res.LastPage
        $pageCookie = $res.Watermark
    }

    return $allRecords
}

function Export-NormalizedCsv {
    param(
        [System.Collections.Generic.List[object]]$Records,
        [string]$Path
    )

    if ($Records.Count -eq 0) {
        Write-Host "  No records — writing empty file with full header row for consistency." -ForegroundColor Yellow
        ($AllColumns -join ",") | Out-File -FilePath $Path -Encoding UTF8
        return
    }

    # Force every record through the full canonical column list, in order.
    # This is what guarantees a clean, uniform, all-columns CSV even though
    # different activity types populate different subsets of fields.
    $normalized = foreach ($record in $Records) {
        $row = [ordered]@{}
        foreach ($col in $AllColumns) {
            $value = $record.$col
            if ($null -ne $value -and $value -isnot [string]) {
                # Flatten nested objects/arrays (e.g. EntityProperties, SensitiveInfoTypeData) to JSON text
                $value = ($value | ConvertTo-Json -Compress -Depth 6)
            }
            $row[$col] = $value
        }
        [PSCustomObject]$row
    }

    $normalized | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

# ===========================================================================
# MAIN
# ===========================================================================

Connect-IfNeeded

if (-not (Test-Path $OutputRootPath)) {
    New-Item -ItemType Directory -Path $OutputRootPath -Force | Out-Null
}

$summary = @()
$combinedRecords = New-Object System.Collections.Generic.List[object]

# ---------------------------------------------------------------------------
# IMPORTANT: all boundaries below are computed in UTC, with the DateTime Kind
# explicitly stamped as Utc. Building windows from local (e.g. IST) time and
# letting PowerShell pass them through as "Unspecified" is what causes the
# "Date range must be within the past 30 days and cannot include future dates"
# error on every call — the service treats an Unspecified-kind DateTime as if
# it were already UTC, so a local-time window silently shifts by your UTC
# offset and can graze both edges of the 30-day window at once.
#
# We also leave a small safety buffer on both edges rather than requesting
# exactly "now" or exactly "30 days ago," since the service's own boundary
# check appears to compare precise timestamps, not calendar dates.
# ---------------------------------------------------------------------------
$NowUtc          = (Get-Date).ToUniversalTime()
$SafeWindowEnd   = $NowUtc.AddMinutes(-15)          # never request up to the literal current instant
$SafeWindowStart = $NowUtc.Date.AddDays(-29).AddHours(1)  # stay a bit inside the 30-day wall, not exactly on it

# ---------------------------------------------------------------------------
# Build the list of dates to export, in priority order:
#   1. -SpecificDates  (explicit, possibly non-contiguous list)
#   2. -StartDate/-EndDate (explicit contiguous range)
#   3. -PastDays        (trailing window — the default)
# ---------------------------------------------------------------------------
$datesToExport = New-Object System.Collections.Generic.List[datetime]

if ($SpecificDates -and $SpecificDates.Count -gt 0) {
    foreach ($d in $SpecificDates) { $datesToExport.Add($d.Date) }
}
elseif ($StartDate -and $EndDate) {
    if ($EndDate -lt $StartDate) {
        throw "-EndDate ($EndDate) cannot be earlier than -StartDate ($StartDate)."
    }
    $cursor = $StartDate.Date
    while ($cursor -le $EndDate.Date) {
        $datesToExport.Add($cursor)
        $cursor = $cursor.AddDays(1)
    }
}
else {
    if ($PastDays -gt 29) {
        Write-Warning "PastDays=$PastDays requested, but capping the oldest day to 29 days back plus a 1-hour buffer to stay safely inside the service's 30-day window."
    }
    for ($i = 1; $i -le $PastDays; $i++) {
        $datesToExport.Add($NowUtc.Date.AddDays(-$i))
    }
}

if ($Activities -and $Activities.Count -gt 0) {
    Write-Host "Activity filter: $($Activities -join ', ')" -ForegroundColor DarkCyan
}
if ($Workloads -and $Workloads.Count -gt 0) {
    Write-Host "Workload filter: $($Workloads -join ', ')" -ForegroundColor DarkCyan
}

# Export each requested date. Activity Explorer only reliably reports "up to 30 days,"
# so today (partial/still-changing) is never included even if requested explicitly.
foreach ($date in $datesToExport) {
    $dayLabel = $date.ToString("yyyy-MM-dd")
    $dayStart = [datetime]::SpecifyKind($date, [DateTimeKind]::Utc)
    $dayEnd   = [datetime]::SpecifyKind($date.AddHours(23).AddMinutes(59).AddSeconds(59), [DateTimeKind]::Utc)

    # Clamp both edges into the safe window computed above.
    if ($dayEnd -gt $SafeWindowEnd)     { $dayEnd = $SafeWindowEnd }
    if ($dayStart -lt $SafeWindowStart) {
        Write-Warning "Skipping $dayLabel — falls outside the safe 30-day window (too close to the service's retention boundary)."
        continue
    }
    if ($dayStart -gt $SafeWindowEnd) {
        Write-Warning "Skipping $dayLabel — this date is today or in the future; Activity Explorer can't export it."
        continue
    }

    Write-Host "`n=== Exporting $dayLabel (UTC $dayStart to $dayEnd) ===" -ForegroundColor Cyan

    $dayFolder = Join-Path $OutputRootPath $dayLabel
    if (-not (Test-Path $dayFolder)) {
        New-Item -ItemType Directory -Path $dayFolder -Force | Out-Null
    }

    $records = Get-FullDayActivityData -DayStartUtc $dayStart -DayEndUtc $dayEnd -PageSize $PageSize -Activities $Activities -Workloads $Workloads

    # Raw JSON kept for full-fidelity/audit purposes (nothing normalized or dropped)
    $jsonPath = Join-Path $dayFolder "$dayLabel-ActivityExplorer-Raw.json"
    $records | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8

    # Normalized, all-columns, Excel-readable CSV
    $csvPath = Join-Path $dayFolder "$dayLabel-ActivityExplorer.csv"
    Export-NormalizedCsv -Records $records -Path $csvPath

    foreach ($r in $records) { $combinedRecords.Add($r) }

    $summary += [PSCustomObject]@{
        Date        = $dayLabel
        RecordCount = $records.Count
        CsvPath     = $csvPath
    }
}

# One combined file across the whole export for convenience (monthly report, pivoting, etc.)
$sortedDates = $datesToExport | Sort-Object
$combinedLabel = "$($datesToExport.Count)-Dates-$($sortedDates[0].ToString('yyyyMMdd'))-to-$($sortedDates[-1].ToString('yyyyMMdd'))"
$combinedPath = Join-Path $OutputRootPath "Combined-$combinedLabel.csv"
Export-NormalizedCsv -Records $combinedRecords -Path $combinedPath

Write-Host "`n=== Export summary ===" -ForegroundColor Green
$summary | Format-Table -AutoSize
Write-Host "Total records across $($datesToExport.Count) date(s): $($combinedRecords.Count)"
Write-Host "Combined file: $combinedPath"

<#
SCHEDULING NOTE (30-day rolling window):
Activity Explorer only retains 30 days of data, so this needs to run on a cadence
that never leaves a gap larger than 30 days. Options:
  - Windows Task Scheduler running this .ps1 monthly (or better, weekly) on a
    management VM/jump box where the ExchangeOnlineManagement module is installed.
  - Azure Automation Runbook (see Cruz2812/purview-content-export-automation's
    automation-runbook-updates.ps1 pattern) using a certificate-based app-only
    connection instead of interactive Connect-IPPSSession, bound to a monthly
    or weekly schedule.
<< PLACEHOLDER: decide cadence + hosting (Task Scheduler vs Azure Automation) and
   fill in the corresponding connection block above before productionizing. >>
#>
