# Azure Monthly Cost Script

`Get-AzureMonthlyCosts.ps1` collects Azure cost totals per month (`YYYY-MM`) for a selected subscription.

## What It Does

- Reuses existing Az login context by default; signs in only when no context is available.
- Lists only active subscriptions and keeps only those where your signed-in account is `Owner`.
- Lets you select one subscription from that filtered list.
- Supports optional `-SubscriptionId` to skip interactive subscription selection.
- Supports optional `-ForceLogin` to force a new interactive login and switch context.
- Queries Cost Management data and aggregates totals per month.
- Supports single year or year range input (`YYYY` or `YYYY:YYYY`).
- Skips future years/months automatically.

## Prerequisites

- PowerShell 7+ (`pwsh`) recommended.
- Network access to Azure.
- Azure permissions:
  - Subscription visibility (`Get-AzSubscription`).
  - Ability to read role assignments for Owner filtering.
  - Cost data access for selected subscription (for Cost Management query).
- Modules (script auto-installs if missing):
  - `Az.Accounts`
  - `Az.Resources`

## Usage

Run with default range (current UTC year):

```powershell
pwsh -File .\Get-AzureMonthlyCosts.ps1
```

Run for one year:

```powershell
pwsh -File .\Get-AzureMonthlyCosts.ps1 -YearRange 2024
```

Run for a range:

```powershell
pwsh -File .\Get-AzureMonthlyCosts.ps1 -YearRange 2023:2026
```

Run for a specific subscription ID (no selection prompt):

```powershell
pwsh -File .\Get-AzureMonthlyCosts.ps1 -YearRange 2024 -SubscriptionId 00000000-0000-0000-0000-000000000000
```

Include currency column in output:

```powershell
pwsh -File .\Get-AzureMonthlyCosts.ps1 -YearRange 2024 -ShowCurrency
```

Copy tab-delimited output to clipboard for Excel paste:

```powershell
pwsh -File .\Get-AzureMonthlyCosts.ps1 -YearRange 2024 -CopyForExcel
```

Force interactive login (ignore existing context):

```powershell
pwsh -File .\Get-AzureMonthlyCosts.ps1 -YearRange 2024 -ForceLogin
```

## `-SubscriptionId` Behavior

- Optional parameter.
- When provided, the script:
  - skips interactive subscription selection
  - validates that the subscription is active
  - validates that your current principal has `Owner` role on that subscription
- If validation fails, the script exits with a clear error.

## `-ShowCurrency` Behavior

- Optional switch.
- Default: off (currency column hidden).
- When enabled, output includes `Currency`.

## `-CopyForExcel` Behavior

- Optional switch.
- Default: off.
- When enabled, script copies tab-delimited output (with headers) to clipboard.
- Paste directly into Excel to get separate columns.

## `-ForceLogin` Behavior

- Optional switch.
- Default: off (script tries to reuse existing Az context first).
- When enabled, script always prompts interactive login (`Connect-AzAccount`).
- Useful when you want to change account/tenant/subscription context explicitly.

## `-YearRange` Format

- Allowed:
  - `YYYY` (example: `2024`)
  - `YYYY:YYYY` (example: `2023:2026`)
- Validation:
  - years must be between `2000` and `2100`
  - start year must be `<=` end year

## Important Implementation Details

- Data source:
  - `Microsoft.CostManagement/query` (`api-version=2023-03-01`)
  - `type = ActualCost`
  - `granularity = Daily`
  - aggregation on `PreTaxCost`
- Aggregation:
  - Script aggregates API rows to monthly keys `YYYY-MM`.
- Request pattern:
  - One API request per year in effective range.
  - `3` second delay between yearly requests.
  - Retry on `429` with fixed `10` second backoff.
  - Retry on transient `5xx` errors with exponential backoff.
- Future handling:
  - Future years are skipped.
  - Future months are excluded.
  - If requested range is fully future, script exits cleanly with a message.
- Subscription filtering:
  - Excludes states: `Disabled`, `Deleted`, `Expired`.
  - Keeps only subscriptions where your current principal has `Owner` at subscription scope.

## Azure Rate Limiting (Cost Management Query API)

- Why this section matters:
  - When you query multiple years, the script sends multiple API calls (one per year).
  - Even with retries and delays, Azure can still throttle requests and return `429 Too many requests`.
  - This rate-limit guidance helps you understand why those errors happen and how to tune query ranges/reruns.
- Cost Management Query API uses QPU-based throttling (tenant-level).
- Microsoft recommends calling Cost Management APIs no more than once per day; data is typically refreshed every 4 hours.
- Documented QPU quotas:
  - `12` QPU per `10` seconds
  - `60` QPU per `1` minute
  - `600` QPU per `1` hour
- Current documented QPU model:
  - about `1` QPU per `1` month of queried data.
- Useful response headers:
  - `x-ms-ratelimit-microsoft.costmanagement-qpu-retry-after`
  - `x-ms-ratelimit-microsoft.costmanagement-qpu-consumed`
  - `x-ms-ratelimit-microsoft.costmanagement-qpu-remaining`
- In this script:
  - one request per year in range
  - `3` second delay between yearly requests
  - retry with fixed `10` second backoff on `429`
  - retry with exponential backoff on `5xx`
- Reference:
  - https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/manage-automation

## Output

- Table columns:
  - `Month` (`YYYY-MM`)
  - `Cost`
  - `Currency` (only when `-ShowCurrency` is used)
- If no non-zero costs are found in the effective period, script prints a clear message and exits.

## Notes

- Multi-currency months are shown as comma-separated currencies in `Currency`.
- Cost Management API has rate limits; retries and fixed inter-request delay are built in.
- Azure login UX may still show many subscriptions; the script applies its own filtering after sign-in.

## Disclaimer and License

- This script is provided `as is`, without any warranty or support obligation.
- The authors/contributors are not liable for any claims, damages, or other liability arising from use of this code.
- Redistribution and modification are allowed under the MIT License.
- Full license text: `LICENSE.md`.
