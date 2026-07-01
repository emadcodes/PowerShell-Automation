# ============================================================================
# Tracked License Report
# ============================================================================

$ErrorActionPreference = "Stop"
$scriptStartTime = Get-Date

$CONFIG = @{
    # When IncludeAllUsers is $false, only these EmployeeType values are included.
    # Use one value or many, for example: @("Employee") or @("Employee", "Student")
    EmployeeTypes      = @("Contractor","Employee")
    # Edit this map to add or remove tracked license groups without changing report logic.
    TrackedLicenseGroups = [ordered]@{
        F3 = @("SPE_F1", "M365_F1")
        E3 = @("SPE_E3", "ENTERPRISEPACK")
        E5 = @("SPE_E5", "ENTERPRISEPREMIUM")
    }
    # When $true, the script ignores EmployeeTypes and reports on all users.
    IncludeAllUsers    = $true
    # Fast mode: when $false, no group display-name lookup calls are made.
    ResolveGroupNames  = $false
    OutputFolder       = Join-Path $PSScriptRoot "reports"
}

function Print-Info { Write-Host "[INFO] $args" -ForegroundColor Cyan }
function Print-Ok   { Write-Host "[OK] $args" -ForegroundColor Green }
function Print-Err  { Write-Host "[ERROR] $args" -ForegroundColor Red }

function Format-Duration {
    param([double]$Seconds)
    $ts = [timespan]::FromSeconds([math]::Max(0, $Seconds))
    return "{0}h {1}m {2}s" -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds
}

function Connect-Services {
    $graphScopes = @(
        "User.Read.All"
        "Directory.Read.All"
    )

    Print-Info "Connecting to Microsoft Graph interactively..."
    Connect-MgGraph -Scopes $graphScopes -NoWelcome | Out-Null
    Print-Ok "Connected to Microsoft Graph"
}

function Get-TrackedLicenseReport {
    param(
        [string[]]$EmployeeTypes = @("Employee"),
        [hashtable]$TrackedLicenseGroups,
        [bool]$IncludeAll = $false,
        [bool]$ResolveGroupNames = $false
    )

    Print-Info "Querying users..."
    $userProperties = @(
        "id",
        "userPrincipalName",
        "displayName",
        "mail",
        "employeeType",
        "accountEnabled",
        "assignedLicenses",
        "licenseAssignmentStates",
        "department",
        "jobTitle",
        "usageLocation"
    )

    $allUsers = Get-MgUser -All -Property ($userProperties -join ",") -ConsistencyLevel eventual

    if ($IncludeAll) {
        $users = $allUsers
        Print-Info "Including all users"
    } else {
        $employeeTypesToInclude = @($EmployeeTypes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
        if ($employeeTypesToInclude.Count -eq 0) {
            throw "IncludeAllUsers is set to `$false, but EmployeeTypes is empty. Specify one or more EmployeeType values."
        }

        $users = $allUsers | Where-Object { $_.EmployeeType -in $employeeTypesToInclude }
        Print-Info "Filtered to $($users.Count) users with EmployeeType in: $($employeeTypesToInclude -join ', ')"
    }

    $trackedLicenseNames = @($TrackedLicenseGroups.Keys)
    if ($trackedLicenseNames.Count -eq 0) {
        throw "TrackedLicenseGroups is empty. Specify one or more tracked license groups in the config."
    }

    $allSkus = Get-MgSubscribedSku -All

    $trackedSkuMap = @{}
    foreach ($licenseName in $trackedLicenseNames) {
        $matchingSkus = $allSkus | Where-Object { $_.SkuPartNumber -in $TrackedLicenseGroups[$licenseName] }
        foreach ($sku in $matchingSkus) {
            $trackedSkuMap[[string]$sku.SkuId] = @{
                Name       = $licenseName
                PartNumber = $sku.SkuPartNumber
            }
        }
    }

    $groupNameCache = @{}
    $results = [System.Collections.Generic.List[object]]::new()
    $userCount = @($users).Count
    $processedCount = 0
    $stageStart = Get-Date

    foreach ($user in $users) {
        $processedCount++
        $percentComplete = if ($userCount -gt 0) { [math]::Min(100, [int](($processedCount / $userCount) * 100)) } else { 100 }
        
        $elapsed = ((Get-Date) - $stageStart).TotalSeconds
        $avgTimePerUser = if ($processedCount -gt 0) { $elapsed / $processedCount } else { 0 }
        $remainingUsers = $userCount - $processedCount
        $estimatedRemainingSeconds = $avgTimePerUser * $remainingUsers
        
        Write-Progress -Activity "Building report" -Status "User $processedCount of $userCount | Elapsed: $(Format-Duration $elapsed) | ETA: $(Format-Duration $estimatedRemainingSeconds)" -PercentComplete $percentComplete
        
        $labelPartNumbers = @{}
        $labelAssignmentPaths = @{}
        foreach ($licenseName in $trackedLicenseNames) {
            $labelPartNumbers[$licenseName] = @()
            $labelAssignmentPaths[$licenseName] = @()
        }

        $userSkuIds = @($user.AssignedLicenses | ForEach-Object { [string]$_.SkuId })
        foreach ($userSkuId in $userSkuIds) {
            if ($trackedSkuMap.ContainsKey($userSkuId)) {
                $skuInfo = $trackedSkuMap[$userSkuId]
                $labelPartNumbers[$skuInfo.Name] += $skuInfo.PartNumber
            }
        }

        foreach ($state in @($user.LicenseAssignmentStates)) {
            $stateSkuId = [string]$state.SkuId
            if (-not $trackedSkuMap.ContainsKey($stateSkuId)) {
                continue
            }

            $licenseName = $trackedSkuMap[$stateSkuId].Name
            if ($state.AssignedByGroup) {
                $groupId = [string]$state.AssignedByGroup
                if (-not $ResolveGroupNames) {
                    $labelAssignmentPaths[$licenseName] += "Group: $groupId"
                    continue
                }

                if (-not $groupNameCache.ContainsKey($groupId)) {
                    $group = Get-MgGroup -GroupId $groupId -Property "displayName" -ErrorAction SilentlyContinue
                    $groupNameCache[$groupId] = if ($group -and $group.DisplayName) { $group.DisplayName } else { "Unknown group" }
                }
                $labelAssignmentPaths[$licenseName] += "Group: $($groupNameCache[$groupId]) ($groupId)"
            } else {
                $labelAssignmentPaths[$licenseName] += "Direct"
            }
        }

        $trackedLicenseSummary = @()
        $trackedLicenseAssignmentSummary = @()
        foreach ($licenseName in $trackedLicenseNames) {
            $partNumbers = @($labelPartNumbers[$licenseName] | Select-Object -Unique)
            if ($partNumbers.Count -gt 0) {
                $trackedLicenseSummary += "$licenseName ($($partNumbers -join ', '))"
            }

            $assignmentPaths = @($labelAssignmentPaths[$licenseName] | Select-Object -Unique)
            if ($assignmentPaths.Count -gt 0) {
                $trackedLicenseAssignmentSummary += "${licenseName}: $($assignmentPaths -join ', ')"
            }
        }

        $results.Add([PSCustomObject]@{
            ObjectId           = $user.Id
            UserPrincipalName  = $user.UserPrincipalName
            DisplayName        = $user.DisplayName
            Mail               = $user.Mail
            EmployeeType       = $user.EmployeeType
            Department         = $user.Department
            JobTitle           = $user.JobTitle
            UsageLocation      = $user.UsageLocation
            AccountEnabled     = $user.AccountEnabled
            HasTrackedLicense  = ($trackedLicenseSummary.Count -gt 0)
            TrackedLicenseSummary = if ($trackedLicenseSummary.Count -gt 0) { $trackedLicenseSummary -join '; ' } else { "None" }
            TrackedLicenseAssignmentPaths = if ($trackedLicenseAssignmentSummary.Count -gt 0) { $trackedLicenseAssignmentSummary -join '; ' } else { "None" }
        }) | Out-Null
    }

    Write-Progress -Activity "Building report" -Completed

    Print-Ok "Built report for $($results.Count) users"
    return @($results.ToArray())
}

function Save-ReportOutput {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Report,
        [Parameter(Mandatory = $true)]
        [string]$OutputFolder
    )

    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileName = "tracked_license_report_$timestamp.csv"
    $filePath = Join-Path $OutputFolder $fileName

    $Report | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Print-Ok "CSV saved locally: $filePath"
}

Connect-Services
$report = Get-TrackedLicenseReport -EmployeeTypes $CONFIG.EmployeeTypes -TrackedLicenseGroups $CONFIG.TrackedLicenseGroups -IncludeAll $CONFIG.IncludeAllUsers -ResolveGroupNames $CONFIG.ResolveGroupNames
Save-ReportOutput -Report $report -OutputFolder $CONFIG.OutputFolder

# Calculate and display total execution time
$scriptEndTime = Get-Date
$totalElapsed = $scriptEndTime - $scriptStartTime
Print-Ok "Script execution completed in: $(Format-Duration $totalElapsed.TotalSeconds)"
