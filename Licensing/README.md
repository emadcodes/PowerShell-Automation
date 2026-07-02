# M365 License Audit Report (PowerShell)

This repository contains a PowerShell script that generates a CSV report of Microsoft 365 users and whether they have any configured tracked licenses, including how those licenses are assigned (direct vs. group-based).

Primary script:
- `M365LicenseAuditReport.ps1`

## Why this is helpful

This report is useful for people who need quick, reliable license visibility without manually checking users one by one in Microsoft 365 portals.

Relevant audiences:
- IT administrators: validate licensing coverage and assignment method.
- Identity and access teams: detect group-based licensing patterns and direct assignments.
- Procurement and finance partners: support license planning and audits.
- Security and compliance teams: verify account and assignment consistency.
- Managers and HR operations partners: understand employee/contractor license distribution when employee type filtering is used.

Benefits:
- Reduces manual effort and portal-click work.
- Produces timestamped CSV evidence for audits and reviews.
- Helps spot assignment inconsistencies across users.
- Supports recurring operational reporting.

## What the script does

`M365LicenseAuditReport.ps1`:
1. Connects to Microsoft Graph interactively.
2. Retrieves users and selected user properties.
3. Optionally filters users by `EmployeeType`.
4. Maps tracked labels (your choice) to one or more SKU part numbers.
5. Determines tracked license presence and assignment path per user.
6. Exports a timestamped CSV file to the `reports` folder.

## Prerequisites

- PowerShell 7+ (recommended) or Windows PowerShell 5.1.
- Microsoft Graph PowerShell SDK modules:
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Users`
  - `Microsoft.Graph.Identity.DirectoryManagement`
  - `Microsoft.Graph.Groups` (needed when resolving group names)
- Microsoft Entra permissions consented for your account:
  - `User.Read.All`
  - `Directory.Read.All`

Install Graph modules (if not already installed):

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

## Configuration

Edit the `$CONFIG` block in `M365LicenseAuditReport.ps1`.

- `EmployeeTypes`: list of employee types to include when `IncludeAllUsers = $false`.
- `TrackedLicenseGroups`: map friendly labels to SKU part numbers.
- `IncludeAllUsers`: when `$true`, ignores `EmployeeTypes` and reports on all users.
- `ResolveGroupNames`: when `$true`, resolves group display names for group-based assignments (slower).
- `OutputFolder`: destination folder for exported CSV files.

Current tracked map in this repository (example only, fully customizable):

- `F3` -> `SPE_F1`, `M365_F1`
- `E3` -> `SPE_E3`, `ENTERPRISEPACK`
- `E5` -> `SPE_E5`, `ENTERPRISEPREMIUM`

## How to run

From the repository root:

```powershell
pwsh .\M365LicenseAuditReport.ps1
```

Or with Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\M365LicenseAuditReport.ps1
```

The script will prompt for interactive sign-in to Microsoft Graph.

## Output

The script creates a CSV in:
- `reports\m365_license_audit_report_yyyyMMdd_HHmmss.csv`

Key output columns include:
- `ObjectId`
- `UserPrincipalName`
- `DisplayName`
- `EmployeeType`
- `AccountEnabled`
- `HasTrackedLicense`
- `TrackedLicenseSummary`
- `TrackedLicenseAssignmentPaths`

## Example use cases

- Monthly contractor license review.
- Track any E*/F* family combinations (for example, E3, F3, E5) to understand how each license is assigned.
- Identify users with direct assignment where group-based assignment is preferred.
- Validate post-migration license state (e.g., F3 to E3 transitions).
- Compare assignment behavior across tracked licenses for operational cleanup, governance, and audit readiness.
- Produce evidence for compliance or internal control checks.

## Important note

This script is not limited to E3, F3, or E5.

You can track any license set by editing `TrackedLicenseGroups` in `M365LicenseAuditReport.ps1`. The labels are just friendly names, and each label can map to one or many SKU part numbers.

## Performance notes

- `ResolveGroupNames = $false` runs faster because it avoids per-group display name lookup.
- Very large tenants may take time because all users are enumerated.
- Progress and ETA are shown while building the report.

## Troubleshooting

- `Connect-MgGraph` permission errors:
  - Ensure your account can consent to required Graph scopes.
- Empty output or unexpected user set:
  - Check `IncludeAllUsers` and `EmployeeTypes` values.
- Missing license matches:
  - Confirm SKU part numbers in `TrackedLicenseGroups` match your tenant SKUs.

## Security and operational guidance

- Treat exported reports as potentially sensitive identity data.
- Store CSV outputs in approved locations.
- Limit access to report files based on least privilege.