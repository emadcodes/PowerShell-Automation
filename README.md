# Entra ID License Tracking Scripts

This repository contains PowerShell scripts to track Microsoft 365 license assignments in Entra ID and export results to CSV for reporting.

## Primary Scenario

Use this when you need to answer questions like:
- Which users have one or more licenses from a tracked list (for example F3/E3/E5/Power Apps)?
- Were those licenses assigned directly or through groups?
- Which employee populations should be included (all users or selected `EmployeeType` values)?

This is useful for:
- license governance
- cleanup and right-sizing initiatives
- monthly audit snapshots
- preparing data for Excel or Power BI

## Main Script

- `AssignCoreLicense_Simple.ps1`

This script generates a tracked-license report based on the SKU groups defined in `TrackedLicenseGroups`.

## How It Works

1. Connects to Microsoft Graph with read-only directory/user scopes.
2. Pulls users and optional employee-type filtered population.
3. Resolves your configured tracked SKU part numbers to tenant SKU IDs.
4. Checks each user for matching tracked licenses.
5. Builds summary fields including assignment path (`Direct` or `Group: <id/name>`).
6. Exports CSV to the `reports` folder.

## Prerequisites

- PowerShell 7+ (recommended)
- Microsoft Graph PowerShell SDK installed

Install modules if needed:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Required delegated Graph permissions during sign-in:
- `User.Read.All`
- `Directory.Read.All`

## Quick Start

Run from this folder:

```powershell
pwsh .\AssignCoreLicense_Simple.ps1
```

Output file pattern:

- `reports/tracked_license_report_yyyyMMdd_HHmmss.csv`

## Configuration

Edit the `$CONFIG` block in `AssignCoreLicense_Simple.ps1`.

### 1) User Scope Controls

- `IncludeAllUsers = $true`
  - includes all users
- `IncludeAllUsers = $false`
  - filters users by `EmployeeTypes`

Example:

```powershell
EmployeeTypes   = @("Contractor", "Employee")
IncludeAllUsers = $false
```

### 2) Tracked License Groups

`TrackedLicenseGroups` is the key extension point.

```powershell
TrackedLicenseGroups = [ordered]@{
    F3        = @("SPE_F1", "M365_F1")
    E3        = @("SPE_E3", "ENTERPRISEPACK")
    E5        = @("SPE_E5", "ENTERPRISEPREMIUM")
    PowerApps = @("POWERAPPS_VIRAL", "POWERAPPS_PER_USER")
}
```

To add or remove what is tracked, only edit this map.

### 3) Group Display Names vs Fast Mode

- `ResolveGroupNames = $false` (faster)
  - assignment path contains group IDs
- `ResolveGroupNames = $true` (slower)
  - assignment path includes group display names

## CSV Output Columns

Core identity fields:
- `ObjectId`
- `UserPrincipalName`
- `DisplayName`
- `Mail`
- `EmployeeType`
- `Department`
- `JobTitle`
- `UsageLocation`
- `AccountEnabled`

Tracked-license fields:
- `HasTrackedLicense`
- `TrackedLicenseSummary`
- `TrackedLicenseAssignmentPaths`

### Delimiters

- Between license groups: `; `
- Within one group details: `, `

Example values:

```text
TrackedLicenseSummary = E3 (SPE_E3); PowerApps (POWERAPPS_PER_USER)
TrackedLicenseAssignmentPaths = E3: Direct; PowerApps: Group: 00000000-0000-0000-0000-000000000000
```

## Behavior Notes

- `HasTrackedLicense` is boolean only:
  - `True` if user has at least one tracked group
  - `False` if user has none
- If you track 5 groups and user has 2, summary fields contain only those 2.
- CSV schema remains stable while summary content changes based on your tracked map.

## Troubleshooting

- If prompted for sign-in each run, verify your Graph session and tenant context.
- If no matches are found, verify SKU part numbers in `TrackedLicenseGroups` are valid for your tenant.
- If assignment paths are only group IDs, set `ResolveGroupNames = $true`.

## Repository Hygiene

Recommended `.gitignore` entries:

```gitignore
reports/
*.csv
```

Keep scripts in source control, but avoid committing generated report outputs.
