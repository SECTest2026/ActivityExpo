<#
.SYNOPSIS
    Microsoft Purview Activity Explorer V2 exporter.

.DESCRIPTION
    Exports Activity Explorer data using Export-ActivityExplorerData.

    OUTPUTS PER DAY
      1. Raw JSON
         - Full-fidelity records returned by Microsoft.

      2. Backend CSV
         - Microsoft documented Export-ActivityExplorerData properties.

      3. UI80 CSV
         - Normalized representation of the 80 Activity Explorer
           "Customize columns" fields identified from the supplied screenshots.

    ALSO CREATES
      - MappingReport.csv
      - Combined Backend CSV
      - Combined UI80 CSV

    IMPORTANT
      The Activity Explorer UI has fields that do not have a direct,
      documented property in Export-ActivityExplorerData.

      Those fields are:
        - populated directly when a documented property exists
        - derived when they can safely be calculated
        - parsed from nested JSON where possible
        - left blank when Microsoft does not expose the field through
          Export-ActivityExplorerData

      The script never fabricates values.

.REQUIREMENTS
      ExchangeOnlineManagement
      Connect-IPPSSession
      Export-ActivityExplorerData
      Purview permissions sufficient for Activity Explorer

.NOTES
      Activity Explorer reports up to 30 days of data.
      PageSize maximum is 5000.
      PageCookie/Watermark is used to retrieve subsequent pages.
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

    [datetime[]]$SpecificDates,

    [int]$PageRetryCount = 3
)

# ============================================================================
# 1. MICROSOFT BACKEND EXPORT SCHEMA
# ============================================================================
# These are the documented Export-ActivityExplorerData properties used by
# the current exporter.

$BackendColumns = @(
    "Activity",
    "Application",
    "ArtifactType",
    "AssociatedAdminUnits",
    "AuthorizedGroupId",
    "AuthorizedGroupName",
    "ClientIP",
    "DataState",
    "DestinationLocationType",
    "DeviceName",
    "DlpPolicyMatchId",
    "EndpointOperation",
    "EnforcementMode",
    "EntityProperties",
    "EvaluationTime",
    "FalsePositive",
    "FileExtension",
    "FilePath",
    "FileSize",
    "FileType",
    "FullUrl",
    "GroupId",
    "GroupName",
    "GroupType",
    "Happened",
    "Hidden",
    "HowApplied",
    "HowAppliedDetail",
    "IRMContentId",
    "IsCorporateNetwork",
    "IsProtected",
    "IsProtectedBefore",
    "JitTriggered",
    "Justification",
    "LabelEventType",
    "Manufacturer",
    "MatchedWithV1DetailedScheme",
    "MDATPDeviceId",
    "Model",
    "OldRetentionLabel",
    "OldSensitivityLabel",
    "OriginatingDomain",
    "ParentArchiveHash",
    "Platform",
    "PolicyId",
    "PolicyMode",
    "PolicyName",
    "PreviousFileName",
    "PreviousFilePath",
    "PreviousProtectionOwner",
    "ProcessName",
    "ProductVersion",
    "ProtectionEventType",
    "ProtectionOwner",
    "ProtectionType",
    "Reason",
    "Receivers",
    "RecordIdentity",
    "RetentionLabel",
    "RMSEncrypted",
    "RuleActions",
    "RuleId",
    "RuleName",
    "Sender",
    "SensitiveInfoTypeBucketsData",
    "SensitiveInfoTypeData",
    "SensitivityLabel",
    "SensitivityLabelPolicyId",
    "SerialNumber",
    "Sha1",
    "Sha256",
    "SourceLocationType",
    "StorageName",
    "Subject",
    "TargetDomain",
    "TargetFilePath",
    "TargetPrinterName",
    "TemplateId",
    "User",
    "UserSku",
    "UserType",
    "VpnNetworkAddress",
    "VpnServerAddress",
    "Workload"
)

# ============================================================================
# 2. ACTIVITY EXPLORER UI - 80 COLUMNS
# ============================================================================
# Exact names based on the supplied Activity Explorer screenshots.

$UI80Columns = @(
    "Activity",
    "File",
    "Location",
    "Enforcement plane",
    "User",
    "Happened",
    "Sensitivity label",
    "Old sensitivity label",
    "Sensitivity label policy",
    "Retention label",
    "Old retention label",
    "Sensitive info type",
    "Trainable classifier",
    "Sensitive info type count",
    "Sensitive info type - metadata",
    "Policy",
    "Rule",
    "Policy mode",
    "Rule actions",
    "Email subject",
    "Email sender",
    "Email recipient",
    "File extension",
    "Client IP",
    "File size",
    "Source location type",
    "Destination location type",
    "Originating domain",
    "How applied",
    "How applied detail",
    "Label event type",
    "Content type",
    "Sha1",
    "Sha256",
    "Platform",
    "Application",
    "Device name",
    "RMS encrypted",
    "Enforcement mode",
    "Previous file name",
    "Previous file path",
    "Target printer name",
    "Target domain",
    "Target file path",
    "Removable media device manufacturer",
    "Removable media device model",
    "Removable media device serial number",
    "Justification",
    "Power BI Item Type",
    "Group ID",
    "Group name",
    "Record identity",
    "JIT triggered",
    "False Positive",
    "Reason",
    "Printer Group ID",
    "Printer Group Name",
    "Printer Alias",
    "Removable USB device ID",
    "Removable USB device group name",
    "Removable USB device display name",
    "Network Share Group ID",
    "Network share group name",
    "Site group ID",
    "Site group name",
    "Site URL",
    "VPN server address",
    "VPN network address",
    "Is corporate network",
    "Activity type",
    "Site template",
    "Data state",
    "Protection event type",
    "Is protected",
    "Is protected before",
    "Protection owner",
    "Protection owner before",
    "Product version",
    "Protection type",
    "Template Id"
)

# Verify UI list really contains 80 entries.
if ($UI80Columns.Count -ne 80) {
    throw "UI80Columns contains $($UI80Columns.Count) columns instead of 80."
}

# ============================================================================
# 3. HELPER FUNCTIONS
# ============================================================================

function Get-PropertyValue {
    param(
        [object]$Object,
        [string[]]$Names
    )

    if ($null -eq $Object) {
        return $null
    }

    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties |
            Where-Object { $_.Name -ieq $name } |
            Select-Object -First 1

        if ($property) {
            return $property.Value
        }
    }

    return $null
}

function Convert-ValueToCsvSafeString {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    if ($Value -is [string]) {
        return $Value
    }

    try {
        return ($Value | ConvertTo-Json -Compress -Depth 10)
    }
    catch {
        return [string]$Value
    }
}

function Get-FileNameFromPath {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    try {
        return [System.IO.Path]::GetFileName($Path)
    }
    catch {
        return $Path
    }
}

function Get-FirstJsonValue {
    param(
        [object]$Value,
        [string[]]$PropertyNames
    )

    if ($null -eq $Value) {
        return $null
    }

    # Already an object
    if ($Value -isnot [string]) {
        return Get-PropertyValue -Object $Value -Names $PropertyNames
    }

    # Try JSON parsing
    try {
        $parsed = $Value | ConvertFrom-Json -ErrorAction Stop
        return Get-PropertyValue -Object $parsed -Names $PropertyNames
    }
    catch {
        return $null
    }
}

function Get-SensitiveInfoDetails {
    param(
        [object]$Record
    )

    $raw = Get-PropertyValue -Object $Record -Names @(
        "SensitiveInfoTypeData"
    )

    if ($null -eq $raw) {
        return [PSCustomObject]@{
            Type       = ""
            Count      = ""
            Metadata   = ""
            Classifier = ""
        }
    }

    $jsonText = Convert-ValueToCsvSafeString $raw

    $typeValue = ""
    $countValue = ""
    $metadataValue = ""
    $classifierValue = ""

    try {
        $parsed = $jsonText | ConvertFrom-Json -ErrorAction Stop

        if ($parsed -is [System.Array]) {
            $items = $parsed
        }
        else {
            $items = @($parsed)
        }

        $typeValues = New-Object System.Collections.Generic.List[string]
        $countValues = New-Object System.Collections.Generic.List[string]
        $classifierValues = New-Object System.Collections.Generic.List[string]

        foreach ($item in $items) {

            $type = Get-PropertyValue $item @(
                "Name",
                "name",
                "SensitiveInfoTypeName",
                "SensitiveInfoType",
                "Type",
                "type"
            )

            if ($null -ne $type) {
                $typeValues.Add([string]$type)
            }

            $count = Get-PropertyValue $item @(
                "Count",
                "count",
                "SensitiveInfoTypeCount",
                "MatchCount"
            )

            if ($null -ne $count) {
                $countValues.Add([string]$count)
            }

            $classifier = Get-PropertyValue $item @(
                "Classifier",
                "ClassifierName",
                "TrainableClassifier",
                "TrainableClassifierName"
            )

            if ($null -ne $classifier) {
                $classifierValues.Add([string]$classifier)
            }
        }

        $typeValue = ($typeValues | Select-Object -Unique) -join "; "
        $countValue = ($countValues | Select-Object -Unique) -join "; "
        $classifierValue = ($classifierValues | Select-Object -Unique) -join "; "

        $metadataValue = $jsonText
    }
    catch {
        # Preserve raw data if Microsoft changes the structure.
        $metadataValue = $jsonText
    }

    return [PSCustomObject]@{
        Type       = $typeValue
        Count      = $countValue
        Metadata   = $metadataValue
        Classifier = $classifierValue
    }
}

# ============================================================================
# 4. CONNECT TO PURVIEW
# ============================================================================

function Connect-IfNeeded {

    if ($SkipConnect) {
        Write-Host "Skipping connection as requested." -ForegroundColor Yellow
        return
    }

    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null

        Write-Host `
            "Already connected to Security & Compliance PowerShell." `
            -ForegroundColor Green
    }
    catch {

        Write-Host `
            "Connecting to Security & Compliance PowerShell..." `
            -ForegroundColor Cyan

        if ($UserPrincipalName) {
            Connect-IPPSSession `
                -UserPrincipalName $UserPrincipalName
        }
        else {
            Connect-IPPSSession
        }
    }
}

# ============================================================================
# 5. FILTER BUILDER
# ============================================================================

function Add-ActivityFilters {

    param(
        [hashtable]$Params
    )

    $filterList = @()

    if ($Activities -and $Activities.Count -gt 0) {
        $filterList += ,(@("Activity") + $Activities)
    }

    if ($Workloads -and $Workloads.Count -gt 0) {
        $filterList += ,(@("Workload") + $Workloads)
    }

    for ($i = 0; $i -lt $filterList.Count; $i++) {

        $filterNumber = $i + 1

        if ($filterNumber -gt 5) {
            throw "Maximum supported Export-ActivityExplorerData filters are Filter1 through Filter5."
        }

        $Params["Filter$filterNumber"] = $filterList[$i]
    }
}

# ============================================================================
# 6. GET ALL PAGES FOR ONE DAY
# ============================================================================

function Get-FullDayActivityData {

    param(
        [datetime]$DayStartUtc,
        [datetime]$DayEndUtc,
        [int]$PageSize
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

        if ($pageCookie) {
            $params["PageCookie"] = $pageCookie
        }

        Add-ActivityFilters -Params $params

        $res = $null
        $success = $false

        for ($attempt = 1; $attempt -le $PageRetryCount; $attempt++) {

            try {

                Write-Host `
                    "  Requesting page $pageNumber (attempt $attempt)..." `
                    -ForegroundColor DarkGray

                $res = Export-ActivityExplorerData @params

                $success = $true
                break
            }
            catch {

                Write-Warning `
                    "  Page $pageNumber attempt $attempt failed: $($_.Exception.Message)"

                if ($attempt -lt $PageRetryCount) {
                    Start-Sleep -Seconds ([Math]::Min(5 * $attempt, 15))
                }
            }
        }

        if (-not $success) {
            throw "Unable to retrieve Activity Explorer page $pageNumber after $PageRetryCount attempts."
        }

        if ($res.ResultData) {

            $parsed = $res.ResultData | ConvertFrom-Json

            if ($parsed -isnot [System.Array]) {
                $parsed = @($parsed)
            }

            foreach ($record in $parsed) {
                $allRecords.Add($record)
            }

            Write-Host `
                "  Page $pageNumber : $($parsed.Count) records | Total: $($allRecords.Count)" `
                -ForegroundColor Gray
        }
        else {

            Write-Host `
                "  Page $pageNumber : 0 records" `
                -ForegroundColor Gray
        }

        $lastPage = [bool]$res.LastPage

        if (-not $lastPage) {

            if ([string]::IsNullOrWhiteSpace([string]$res.Watermark)) {
                throw "Microsoft returned LastPage=False but no Watermark/PageCookie."
            }

            $pageCookie = $res.Watermark
        }
    }

    return $allRecords
}

# ============================================================================
# 7. BACKEND 84-COLUMN CSV
# ============================================================================

function Export-BackendCsv {

    param(
        [System.Collections.Generic.List[object]]$Records,
        [string]$Path
    )

    if ($Records.Count -eq 0) {

        $headerObject = [ordered]@{}

        foreach ($column in $BackendColumns) {
            $headerObject[$column] = ""
        }

        [PSCustomObject]$headerObject |
            Export-Csv `
                -Path $Path `
                -NoTypeInformation `
                -Encoding UTF8

        return
    }

    $rows = foreach ($record in $Records) {

        $row = [ordered]@{}

        foreach ($column in $BackendColumns) {

            $value = Get-PropertyValue `
                -Object $record `
                -Names @($column)

            $row[$column] = Convert-ValueToCsvSafeString $value
        }

        [PSCustomObject]$row
    }

    $rows |
        Export-Csv `
            -Path $Path `
            -NoTypeInformation `
            -Encoding UTF8
}

# ============================================================================
# 8. UI 80-COLUMN MAPPING
# ============================================================================

function Convert-ToActivityExplorerUI80 {

    param(
        [object]$Record
    )

    $sit = Get-SensitiveInfoDetails -Record $Record

    # ------------------------------------------------------------------------
    # Direct / derived values
    # ------------------------------------------------------------------------

    $filePath = Get-PropertyValue $Record @(
        "FilePath"
    )

    $fileName = Get-FileNameFromPath $filePath

    $policyName = Get-PropertyValue $Record @(
        "PolicyName"
    )

    $ruleName = Get-PropertyValue $Record @(
        "RuleName"
    )

    $subject = Get-PropertyValue $Record @(
        "Subject"
    )

    $sender = Get-PropertyValue $Record @(
        "Sender"
    )

    $receivers = Get-PropertyValue $Record @(
        "Receivers"
    )

    $sourceLocation = Get-PropertyValue $Record @(
        "SourceLocationType"
    )

    $destinationLocation = Get-PropertyValue $Record @(
        "DestinationLocationType"
    )

    $location = ""

    # Best-effort location representation.
    # Microsoft does not document a standalone "Location" export property.
    if ($destinationLocation) {
        $location = $destinationLocation
    }
    elseif ($sourceLocation) {
        $location = $sourceLocation
    }

    # Content type:
    # Prefer FileType, then ArtifactType.
    $contentType = Get-PropertyValue $Record @(
        "FileType",
        "ArtifactType"
    )

    # Activity type:
    # EndpointOperation is useful for endpoint events, otherwise Activity.
    $activityType = Get-PropertyValue $Record @(
        "EndpointOperation",
        "Activity"
    )

    # Policy mapping
    $policy = $policyName

    # Rule mapping
    $rule = $ruleName

    # Sensitivity label policy
    $sensitivityLabelPolicy = Get-PropertyValue $Record @(
        "SensitivityLabelPolicyId"
    )

    # Protection owner before
    $protectionOwnerBefore = Get-PropertyValue $Record @(
        "PreviousProtectionOwner"
    )

    # ------------------------------------------------------------------------
    # Build exactly 80 columns
    # ------------------------------------------------------------------------

    $row = [ordered]@{}

    foreach ($column in $UI80Columns) {

        $value = ""

        switch ($column) {

            "Activity" {
                $value = Get-PropertyValue $Record @("Activity")
            }

            "File" {
                $value = $fileName
            }

            "Location" {
                $value = $location
            }

            "Enforcement plane" {
                # No documented direct Export-ActivityExplorerData property.
                $value = ""
            }

            "User" {
                $value = Get-PropertyValue $Record @("User")
            }

            "Happened" {
                $value = Get-PropertyValue $Record @("Happened")
            }

            "Sensitivity label" {
                $value = Get-PropertyValue $Record @("SensitivityLabel")
            }

            "Old sensitivity label" {
                $value = Get-PropertyValue $Record @("OldSensitivityLabel")
            }

            "Sensitivity label policy" {
                $value = $sensitivityLabelPolicy
            }

            "Retention label" {
                $value = Get-PropertyValue $Record @("RetentionLabel")
            }

            "Old retention label" {
                $value = Get-PropertyValue $Record @("OldRetentionLabel")
            }

            "Sensitive info type" {
                $value = $sit.Type
            }

            "Trainable classifier" {
                $value = $sit.Classifier
            }

            "Sensitive info type count" {
                $value = $sit.Count
            }

            "Sensitive info type - metadata" {
                $value = $sit.Metadata
            }

            "Policy" {
                $value = $policy
            }

            "Rule" {
                $value = $rule
            }

            "Policy mode" {
                $value = Get-PropertyValue $Record @("PolicyMode")
            }

            "Rule actions" {
                $value = Get-PropertyValue $Record @("RuleActions")
            }

            "Email subject" {
                $value = $subject
            }

            "Email sender" {
                $value = $sender
            }

            "Email recipient" {
                $value = $receivers
            }

            "File extension" {
                $value = Get-PropertyValue $Record @("FileExtension")
            }

            "Client IP" {
                $value = Get-PropertyValue $Record @("ClientIP")
            }

            "File size" {
                $value = Get-PropertyValue $Record @("FileSize")
            }

            "Source location type" {
                $value = $sourceLocation
            }

            "Destination location type" {
                $value = $destinationLocation
            }

            "Originating domain" {
                $value = Get-PropertyValue $Record @("OriginatingDomain")
            }

            "How applied" {
                $value = Get-PropertyValue $Record @("HowApplied")
            }

            "How applied detail" {
                $value = Get-PropertyValue $Record @("HowAppliedDetail")
            }

            "Label event type" {
                $value = Get-PropertyValue $Record @("LabelEventType")
            }

            "Content type" {
                $value = $contentType
            }

            "Sha1" {
                $value = Get-PropertyValue $Record @("Sha1")
            }

            "Sha256" {
                $value = Get-PropertyValue $Record @("Sha256")
            }

            "Platform" {
                $value = Get-PropertyValue $Record @("Platform")
            }

            "Application" {
                $value = Get-PropertyValue $Record @("Application")
            }

            "Device name" {
                $value = Get-PropertyValue $Record @("DeviceName")
            }

            "RMS encrypted" {
                $value = Get-PropertyValue $Record @("RMSEncrypted")
            }

            "Enforcement mode" {
                $value = Get-PropertyValue $Record @("EnforcementMode")
            }

            "Previous file name" {
                $value = Get-PropertyValue $Record @("PreviousFileName")
            }

            "Previous file path" {
                $value = Get-PropertyValue $Record @("PreviousFilePath")
            }

            "Target printer name" {
                $value = Get-PropertyValue $Record @("TargetPrinterName")
            }

            "Target domain" {
                $value = Get-PropertyValue $Record @("TargetDomain")
            }

            "Target file path" {
                $value = Get-PropertyValue $Record @("TargetFilePath")
            }

            "Removable media device manufacturer" {
                $value = Get-PropertyValue $Record @("Manufacturer")
            }

            "Removable media device model" {
                $value = Get-PropertyValue $Record @("Model")
            }

            "Removable media device serial number" {
                $value = Get-PropertyValue $Record @("SerialNumber")
            }

            "Justification" {
                $value = Get-PropertyValue $Record @("Justification")
            }

            "Power BI Item Type" {
                $value = Get-PropertyValue $Record @(
                    "ArtifactType"
                )
            }

            "Group ID" {
                $value = Get-PropertyValue $Record @("GroupId")
            }

            "Group name" {
                $value = Get-PropertyValue $Record @("GroupName")
            }

            "Record identity" {
                $value = Get-PropertyValue $Record @("RecordIdentity")
            }

            "JIT triggered" {
                $value = Get-PropertyValue $Record @("JitTriggered")
            }

            "False Positive" {
                $value = Get-PropertyValue $Record @("FalsePositive")
            }

            "Reason" {
                $value = Get-PropertyValue $Record @("Reason")
            }

            "Printer Group ID" {
                $value = ""
            }

            "Printer Group Name" {
                $value = ""
            }

            "Printer Alias" {
                $value = ""
            }

            "Removable USB device ID" {
                $value = ""
            }

            "Removable USB device group name" {
                $value = ""
            }

            "Removable USB device display name" {
                $value = ""
            }

            "Network Share Group ID" {
                $value = ""
            }

            "Network share group name" {
                $value = ""
            }

            "Site group ID" {
                $value = ""
            }

            "Site group name" {
                $value = ""
            }

            "Site URL" {
                $value = ""
            }

            "VPN server address" {
                $value = Get-PropertyValue $Record @("VpnServerAddress")
            }

            "VPN network address" {
                $value = Get-PropertyValue $Record @("VpnNetworkAddress")
            }

            "Is corporate network" {
                $value = Get-PropertyValue $Record @("IsCorporateNetwork")
            }

            "Activity type" {
                $value = $activityType
            }

            "Site template" {
                $value = ""
            }

            "Data state" {
                $value = Get-PropertyValue $Record @("DataState")
            }

            "Protection event type" {
                $value = Get-PropertyValue $Record @("ProtectionEventType")
            }

            "Is protected" {
                $value = Get-PropertyValue $Record @("IsProtected")
            }

            "Is protected before" {
                $value = Get-PropertyValue $Record @("IsProtectedBefore")
            }

            "Protection owner" {
                $value = Get-PropertyValue $Record @("ProtectionOwner")
            }

            "Protection owner before" {
                $value = $protectionOwnerBefore
            }

            "Product version" {
                $value = Get-PropertyValue $Record @("ProductVersion")
            }

            "Protection type" {
                $value = Get-PropertyValue $Record @("ProtectionType")
            }

            "Template Id" {
                $value = Get-PropertyValue $Record @("TemplateId")
            }

            default {
                $value = ""
            }
        }

        $row[$column] = Convert-ValueToCsvSafeString $value
    }

    return [PSCustomObject]$row
}

# ============================================================================
# 9. EXPORT UI80 CSV
# ============================================================================

function Export-UI80Csv {

    param(
        [System.Collections.Generic.List[object]]$Records,
        [string]$Path
    )

    if ($Records.Count -eq 0) {

        $header = [ordered]@{}

        foreach ($column in $UI80Columns) {
            $header[$column] = ""
        }

        [PSCustomObject]$header |
            Export-Csv `
                -Path $Path `
                -NoTypeInformation `
                -Encoding UTF8

        return
    }

    $rows = foreach ($record in $Records) {
        Convert-ToActivityExplorerUI80 -Record $record
    }

    $rows |
        Export-Csv `
            -Path $Path `
            -NoTypeInformation `
            -Encoding UTF8
}

# ============================================================================
# 10. MAPPING REPORT
# ============================================================================

function New-MappingReport {

    param(
        [string]$Path
    )

    $direct = @(
        "Activity",
        "User",
        "Happened",
        "Sensitivity label",
        "Old sensitivity label",
        "Retention label",
        "Old retention label",
        "Policy mode",
        "Rule actions",
        "File extension",
        "Client IP",
        "File size",
        "Source location type",
        "Destination location type",
        "Originating domain",
        "How applied",
        "How applied detail",
        "Label event type",
        "Sha1",
        "Sha256",
        "Platform",
        "Application",
        "Device name",
        "RMS encrypted",
        "Enforcement mode",
        "Previous file name",
        "Previous file path",
        "Target printer name",
        "Target domain",
        "Target file path",
        "Removable media device manufacturer",
        "Removable media device model",
        "Removable media device serial number",
        "Justification",
        "Group ID",
        "Group name",
        "Record identity",
        "JIT triggered",
        "False Positive",
        "Reason",
        "VPN server address",
        "VPN network address",
        "Is corporate network",
        "Data state",
        "Protection event type",
        "Is protected",
        "Is protected before",
        "Protection owner",
        "Product version",
        "Protection type",
        "Template Id"
    )

    $derived = @(
        "File",
        "Location",
        "Policy",
        "Rule",
        "Email subject",
        "Email sender",
        "Email recipient",
        "Content type",
        "Power BI Item Type",
        "Activity type",
        "Protection owner before"
    )

    $nested = @(
        "Sensitive info type",
        "Trainable classifier",
        "Sensitive info type count",
        "Sensitive info type - metadata"
    )

    $unavailable = @(
        "Enforcement plane",
        "Printer Group ID",
        "Printer Group Name",
        "Printer Alias",
        "Removable USB device ID",
        "Removable USB device group name",
        "Removable USB device display name",
        "Network Share Group ID",
        "Network share group name",
        "Site group ID",
        "Site group name",
        "Site URL",
        "Site template"
    )

    $report = foreach ($column in $UI80Columns) {

        $status = "Unavailable"
        $source = ""
        $notes = ""

        if ($direct -contains $column) {
            $status = "Direct"
            $source = "Export-ActivityExplorerData"
        }
        elseif ($derived -contains $column) {
            $status = "Derived"
            $source = "Export-ActivityExplorerData"
        }
        elseif ($nested -contains $column) {
            $status = "Nested"
            $source = "SensitiveInfoTypeData"
        }
        elseif ($unavailable -contains $column) {
            $status = "Unavailable"
            $source = "Not exposed as documented Export-ActivityExplorerData property"
            $notes = "Left blank; raw JSON retained."
        }

        [PSCustomObject]@{
            UIColumn    = $column
            Status      = $status
            Source      = $source
            Notes       = $notes
        }
    }

    $report |
        Export-Csv `
            -Path $Path `
            -NoTypeInformation `
            -Encoding UTF8
}

# ============================================================================
# 11. MAIN
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Microsoft Purview Activity Explorer Exporter V2" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Backend columns : $($BackendColumns.Count)"
Write-Host "UI columns      : $($UI80Columns.Count)"
Write-Host ""

Connect-IfNeeded

# ---------------------------------------------------------------------------
# Create output directory
# ---------------------------------------------------------------------------

if (-not (Test-Path $OutputRootPath)) {

    New-Item `
        -ItemType Directory `
        -Path $OutputRootPath `
        -Force |
        Out-Null
}

# ---------------------------------------------------------------------------
# Create mapping report
# ---------------------------------------------------------------------------

$mappingPath = Join-Path `
    $OutputRootPath `
    "MappingReport.csv"

New-MappingReport -Path $mappingPath

Write-Host "Mapping report: $mappingPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# UTC safe window
# ---------------------------------------------------------------------------

$NowUtc = (Get-Date).ToUniversalTime()

# Stay away from the exact current instant.
$SafeWindowEnd = $NowUtc.AddMinutes(-15)

# Stay inside the rolling Activity Explorer retention boundary.
$SafeWindowStart = $NowUtc.Date.AddDays(-29).AddHours(1)

Write-Host ""
Write-Host "Current UTC       : $NowUtc"
Write-Host "Safe window start : $SafeWindowStart"
Write-Host "Safe window end   : $SafeWindowEnd"
Write-Host ""

# ============================================================================
# 12. BUILD DATE LIST
# ============================================================================

$datesToExport = New-Object System.Collections.Generic.List[datetime]

if ($SpecificDates -and $SpecificDates.Count -gt 0) {

    foreach ($date in $SpecificDates) {
        $datesToExport.Add($date.Date)
    }
}
elseif ($StartDate -and $EndDate) {

    if ($EndDate -lt $StartDate) {
        throw "-EndDate cannot be earlier than -StartDate."
    }

    $cursor = $StartDate.Date

    while ($cursor -le $EndDate.Date) {

        $datesToExport.Add($cursor)

        $cursor = $cursor.AddDays(1)
    }
}
else {

    if ($PastDays -gt 29) {

        Write-Warning `
            "PastDays > 29 requested. Activity Explorer is being capped to the safe 29-day range."
    }

    $effectivePastDays = [Math]::Min($PastDays, 29)

    for ($i = 1; $i -le $effectivePastDays; $i++) {

        $datesToExport.Add(
            $NowUtc.Date.AddDays(-$i)
        )
    }
}

if ($datesToExport.Count -eq 0) {
    throw "No dates selected for export."
}

# Remove duplicates and sort.
$datesToExport = $datesToExport |
    Sort-Object |
    Get-Unique

# ============================================================================
# 13. FILTER INFORMATION
# ============================================================================

if ($Activities -and $Activities.Count -gt 0) {

    Write-Host `
        "Activity filter: $($Activities -join ', ')" `
        -ForegroundColor DarkCyan
}

if ($Workloads -and $Workloads.Count -gt 0) {

    Write-Host `
        "Workload filter: $($Workloads -join ', ')" `
        -ForegroundColor DarkCyan
}

Write-Host ""
Write-Host "Dates to export: $($datesToExport.Count)"
Write-Host ""

# ============================================================================
# 14. EXPORT
# ============================================================================

$summary = New-Object System.Collections.Generic.List[object]

$combinedRecords =
    New-Object System.Collections.Generic.List[object]

foreach ($date in $datesToExport) {

    $dayLabel = $date.ToString("yyyy-MM-dd")

    $dayStart = [datetime]::SpecifyKind(
        $date.Date,
        [DateTimeKind]::Utc
    )

    $dayEnd = [datetime]::SpecifyKind(
        $date.Date.AddHours(23).AddMinutes(59).AddSeconds(59),
        [DateTimeKind]::Utc
    )

    # Clamp to safe window.

    if ($dayEnd -gt $SafeWindowEnd) {
        $dayEnd = $SafeWindowEnd
    }

    if ($dayStart -lt $SafeWindowStart) {

        Write-Warning `
            "Skipping $dayLabel — outside safe 30-day retention boundary."

        continue
    }

    if ($dayStart -gt $SafeWindowEnd) {

        Write-Warning `
            "Skipping $dayLabel — future/current partial day."

        continue
    }

    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor Cyan

    Write-Host `
        "EXPORTING $dayLabel" `
        -ForegroundColor Cyan

    Write-Host `
        "UTC: $dayStart -> $dayEnd" `
        -ForegroundColor DarkGray

    Write-Host "============================================================" `
        -ForegroundColor Cyan

    # ------------------------------------------------------------------------
    # Folder
    # ------------------------------------------------------------------------

    $dayFolder = Join-Path `
        $OutputRootPath `
        $dayLabel

    if (-not (Test-Path $dayFolder)) {

        New-Item `
            -ItemType Directory `
            -Path $dayFolder `
            -Force |
            Out-Null
    }

    # ------------------------------------------------------------------------
    # Retrieve all pages
    # ------------------------------------------------------------------------

    $records = Get-FullDayActivityData `
        -DayStartUtc $dayStart `
        -DayEndUtc $dayEnd `
        -PageSize $PageSize

    Write-Host ""
    Write-Host `
        "Total records for $dayLabel : $($records.Count)" `
        -ForegroundColor Green

    # ------------------------------------------------------------------------
    # Raw JSON
    # ------------------------------------------------------------------------

    $jsonPath = Join-Path `
        $dayFolder `
        "$dayLabel-ActivityExplorer-Raw.json"

    $records |
        ConvertTo-Json `
            -Depth 20 |
        Out-File `
            -FilePath $jsonPath `
            -Encoding UTF8

    # ------------------------------------------------------------------------
    # Backend CSV
    # ------------------------------------------------------------------------

    $backendPath = Join-Path `
        $dayFolder `
        "$dayLabel-ActivityExplorer-Backend.csv"

    Export-BackendCsv `
        -Records $records `
        -Path $backendPath

    # ------------------------------------------------------------------------
    # UI 80 CSV
    # ------------------------------------------------------------------------

    $uiPath = Join-Path `
        $dayFolder `
        "$dayLabel-ActivityExplorer-UI80.csv"

    Export-UI80Csv `
        -Records $records `
        -Path $uiPath

    # ------------------------------------------------------------------------
    # Add to combined dataset
    # ------------------------------------------------------------------------

    foreach ($record in $records) {

        $combinedRecords.Add($record)
    }

    $summary.Add(
        [PSCustomObject]@{
            Date        = $dayLabel
            RecordCount = $records.Count
            RawJson     = $jsonPath
            BackendCsv  = $backendPath
            UI80Csv     = $uiPath
        }
    )
}

# ============================================================================
# 15. COMBINED EXPORTS
# ============================================================================

if ($datesToExport.Count -gt 0) {

    $sortedDates = $datesToExport | Sort-Object

    $combinedLabel = "$($datesToExport.Count)-Dates-$($sortedDates[0].ToString('yyyyMMdd'))-to-$($sortedDates[-1].ToString('yyyyMMdd'))"

    # ------------------------------------------------------------------------
    # Combined Backend
    # ------------------------------------------------------------------------

    $combinedBackendPath = Join-Path `
        $OutputRootPath `
        "Combined-$combinedLabel-Backend.csv"

    Export-BackendCsv `
        -Records $combinedRecords `
        -Path $combinedBackendPath

    # ------------------------------------------------------------------------
    # Combined UI80
    # ------------------------------------------------------------------------

    $combinedUIPath = Join-Path `
        $OutputRootPath `
        "Combined-$combinedLabel-UI80.csv"

    Export-UI80Csv `
        -Records $combinedRecords `
        -Path $combinedUIPath

    # ------------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------------

    $summaryPath = Join-Path `
        $OutputRootPath `
        "Export-Summary.csv"

    $summary |
        Export-Csv `
            -Path $summaryPath `
            -NoTypeInformation `
            -Encoding UTF8

    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor Green

    Write-Host "EXPORT COMPLETE" `
        -ForegroundColor Green

    Write-Host "============================================================" `
        -ForegroundColor Green

    Write-Host ""
    Write-Host "Total records: $($combinedRecords.Count)" `
        -ForegroundColor Green

    Write-Host ""
    Write-Host "Combined Backend CSV:"
    Write-Host $combinedBackendPath

    Write-Host ""
    Write-Host "Combined UI80 CSV:"
    Write-Host $combinedUIPath

    Write-Host ""
    Write-Host "Mapping Report:"
    Write-Host $mappingPath

    Write-Host ""
    Write-Host "Summary:"
    $summary | Format-Table -AutoSize
}
