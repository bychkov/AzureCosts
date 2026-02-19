param(
    # Supported formats: YYYY or YYYY:YYYY. Default is current UTC year.
    [string]$YearRange = ([DateTime]::UtcNow.Year.ToString()),
    # Month sort order for output: asc (historical) or desc (newest first).
    [ValidateSet('asc', 'desc')]
    [string]$SortOrder = 'asc',
    # Optional direct subscription selector. When provided, interactive selection is skipped.
    [string]$SubscriptionId,
    # Optional flag to include currency column in output. Off by default.
    [switch]$ShowCurrency,
    # Optional flag to copy tab-delimited results to clipboard for Excel paste.
    [switch]$CopyForExcel,
    # Force interactive login even if an existing Az context is available.
    [switch]$ForceLogin
)

$ErrorActionPreference = 'Stop'
$script:OwnerRoleGraphReauthAttempted = $false

# Ensures a required Az module exists, installs if missing, then imports it.
function Require-Module {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Module '$Name' not found. Installing for current user..."
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        catch {
            throw "Failed to install module '$Name'. Run: Install-Module $Name -Scope CurrentUser -Force -AllowClobber"
        }
    }

    Import-Module $Name -ErrorAction Stop | Out-Null
}

# Parses REST response JSON safely; returns $null for empty/invalid content.
function Parse-ResponseJson {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Content
    )

    if ([string]::IsNullOrWhiteSpace($Content)) {
        return $null
    }

    try {
        return $Content | ConvertFrom-Json -Depth 100
    }
    catch {
        return $null
    }
}

# Extracts a readable API error message from standard Azure error payloads.
function Get-ApiErrorMessage {
    param(
        [Parameter(Mandatory = $false)]
        $Body
    )

    if ($null -eq $Body) {
        return 'Empty response body.'
    }

    if ($null -ne $Body.error) {
        $code = [string]$Body.error.code
        $message = [string]$Body.error.message

        if (-not [string]::IsNullOrWhiteSpace($code) -and -not [string]::IsNullOrWhiteSpace($message)) {
            return "${code}: $message"
        }

        if (-not [string]::IsNullOrWhiteSpace($message)) {
            return $message
        }

        if (-not [string]::IsNullOrWhiteSpace($code)) {
            return $code
        }
    }

    return 'Unexpected API response.'
}

# Finds the first matching column index by name in API response metadata.
function Get-ColumnIndex {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Columns,
        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    for ($i = 0; $i -lt $Columns.Count; $i++) {
        if ($Names -contains $Columns[$i].name) {
            return $i
        }
    }

    return -1
}

# Converts Azure month/date values (yyyyMM, yyyyMMdd, or parseable date text) to DateTime.
function Parse-MonthValue {
    param(
        [Parameter(Mandatory = $true)]
        $RawValue
    )

    $text = [string]$RawValue

    if ($text -match '^\d{6}$') {
        return [DateTime]::ParseExact($text + '01', 'yyyyMMdd', $null)
    }

    if ($text -match '^\d{8}$') {
        return [DateTime]::ParseExact($text, 'yyyyMMdd', $null)
    }

    return [DateTime]::Parse($text)
}

# Parses and validates -YearRange input (YYYY or YYYY:YYYY).
function Parse-YearRange {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $text = $Value.Trim()
    $startYear = 0
    $endYear = 0

    if ($text -match '^(?<year>\d{4})$') {
        $startYear = [int]$Matches['year']
        $endYear = $startYear
    }
    elseif ($text -match '^(?<start>\d{4})\s*:\s*(?<end>\d{4})$') {
        $startYear = [int]$Matches['start']
        $endYear = [int]$Matches['end']
    }
    else {
        throw "Invalid -YearRange value '$Value'. Use 'YYYY' or 'YYYY:YYYY'."
    }

    if ($startYear -lt 2000 -or $startYear -gt 2100 -or $endYear -lt 2000 -or $endYear -gt 2100) {
        throw "Year range must be between 2000 and 2100. Received '$Value'."
    }

    if ($startYear -gt $endYear) {
        throw "Invalid -YearRange '$Value'. Start year must be less than or equal to end year."
    }

    return [PSCustomObject]@{
        StartYear = $startYear
        EndYear = $endYear
    }
}

# Returns true for Graph/MFA auth errors that can be fixed by re-login with AuthScope.
function Test-IsGraphAuthScopeRequiredError {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $false
    }

    return (
        $Message -match '(?i)MicrosoftGraphEndpointResourceId' -or
        $Message -match '(?i)Authentication failed against resource' -or
        $Message -match '(?i)Azure credentials have not been set up or have expired'
    )
}

# Checks whether the signed-in principal has Owner role at subscription scope.
function Test-OwnerAccessForSubscription {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId,
        [Parameter(Mandatory = $true)]
        $Account
    )

    $scope = "/subscriptions/$SubscriptionId"
    $accountType = [string]$Account.Type
    $accountId = [string]$Account.Id

    try {
        if ($accountType -eq 'User') {
            $assignments = @(Get-AzRoleAssignment -SignInName $accountId -Scope $scope -RoleDefinitionName 'Owner' -ErrorAction Stop)
            return $assignments.Count -gt 0
        }

        if ($accountType -eq 'ServicePrincipal') {
            $assignments = @(Get-AzRoleAssignment -ServicePrincipalName $accountId -Scope $scope -RoleDefinitionName 'Owner' -ErrorAction Stop)
            return $assignments.Count -gt 0
        }

        if ($accountId -match '^[0-9a-fA-F-]{36}$') {
            $assignments = @(Get-AzRoleAssignment -ObjectId $accountId -Scope $scope -RoleDefinitionName 'Owner' -ErrorAction Stop)
            return $assignments.Count -gt 0
        }

        $assignments = @(Get-AzRoleAssignment -SignInName $accountId -Scope $scope -RoleDefinitionName 'Owner' -ErrorAction Stop)
        return $assignments.Count -gt 0
    }
    catch {
        $message = [string]$_.Exception.Message
        $shouldRetryWithGraphAuth = (
            -not $script:OwnerRoleGraphReauthAttempted -and
            (Test-IsGraphAuthScopeRequiredError -Message $message)
        )

        if ($shouldRetryWithGraphAuth) {
            $script:OwnerRoleGraphReauthAttempted = $true
            Write-Warning 'Owner role checks require refreshed Microsoft Graph authentication. Opening login...'
            Connect-AzAccount -SkipContextPopulation -AuthScope MicrosoftGraphEndpointResourceId -ErrorAction Stop | Out-Null

            $refreshedContext = Get-AzContext -ErrorAction SilentlyContinue
            if ($null -ne $refreshedContext -and $null -ne $refreshedContext.Account) {
                return Test-OwnerAccessForSubscription -SubscriptionId $SubscriptionId -Account $refreshedContext.Account
            }

            throw 'Graph authentication refresh succeeded but no active context is available.'
        }

        throw
    }
}

# Queries Cost Management for a time window, then aggregates daily rows to monthly totals.
function Get-BillingMonthTotals {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId,
        [Parameter(Mandatory = $true)]
        [datetime]$FromUtc,
        [Parameter(Mandatory = $true)]
        [datetime]$ToUtc
    )

    $apiVersion = '2023-03-01'

    $payload = @{
        type = 'ActualCost'
        timeframe = 'Custom'
        timePeriod = @{
            from = $FromUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
            to = $ToUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
        }
        dataset = @{
            granularity = 'Daily'
            aggregation = @{
                totalCost = @{
                    name = 'PreTaxCost'
                    function = 'Sum'
                }
            }
        }
    } | ConvertTo-Json -Depth 10 -Compress

    $path = "/subscriptions/$SubscriptionId/providers/Microsoft.CostManagement/query?api-version=$apiVersion"

    # Retry transient throttling/server errors with exponential backoff.
    $maxAttempts = 6
    $response = $null

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $response = Invoke-AzRestMethod -Method POST -Path $path -Payload $payload -ErrorAction Stop
            break
        }
        catch {
            $message = [string]$_.Exception.Message
            $isThrottled = $message -match '(?i)\b429\b' -or $message -match '(?i)too many requests'
            $isServerError = $message -match '(?i)\b50[0-9]\b'

            if ($isThrottled -and $attempt -lt $maxAttempts) {
                $delaySeconds = 10
                Write-Warning "HTTP 429 received (attempt $attempt/$maxAttempts). Retrying in $delaySeconds seconds..."
                Start-Sleep -Seconds $delaySeconds
                continue
            }

            if ($isServerError -and $attempt -lt $maxAttempts) {
                $delaySeconds = [Math]::Min(60, [int][Math]::Pow(2, $attempt))
                Write-Warning "Transient server error (attempt $attempt/$maxAttempts). Retrying in $delaySeconds seconds..."
                Start-Sleep -Seconds $delaySeconds
                continue
            }

            throw
        }
    }

    if ($null -eq $response) {
        throw 'Failed to get response from Cost Management API after retries.'
    }

    $statusCode = [int]$response.StatusCode
    $body = Parse-ResponseJson -Content $response.Content

    if ($statusCode -lt 200 -or $statusCode -ge 300) {
        $apiMessage = Get-ApiErrorMessage -Body $body
        throw "HTTP $statusCode. $apiMessage"
    }

    if ($null -eq $body -or $null -eq $body.properties) {
        throw 'API response did not contain cost properties.'
    }

    $columns = @($body.properties.columns)
    $rows = @($body.properties.rows)

    if (-not $columns) {
        $apiMessage = Get-ApiErrorMessage -Body $body
        throw "No columns in response. $apiMessage"
    }

    $costIndex = Get-ColumnIndex -Columns $columns -Names @('PreTaxCost', 'Cost', 'totalCost')
    $monthIndex = Get-ColumnIndex -Columns $columns -Names @('UsageDate', 'BillingMonth', 'UsageMonth')
    $currencyIndex = Get-ColumnIndex -Columns $columns -Names @('Currency')

    if ($costIndex -lt 0 -or $monthIndex -lt 0) {
        throw 'Missing required columns (BillingMonth and cost) in API response.'
    }

    $result = @{}

    # Normalize all returned rows into month buckets (YYYY-MM).
    foreach ($row in $rows) {
        $monthDate = Parse-MonthValue -RawValue $row[$monthIndex]
        $monthKey = $monthDate.ToString('yyyy-MM')

        if (-not $result.ContainsKey($monthKey)) {
            $result[$monthKey] = @{
                Cost = [decimal]0
                Currencies = @{}
            }
        }

        $result[$monthKey].Cost += [decimal]$row[$costIndex]

        if ($currencyIndex -ge 0) {
            $currency = [string]$row[$currencyIndex]
            if (-not [string]::IsNullOrWhiteSpace($currency)) {
                $result[$monthKey].Currencies[$currency] = $true
            }
        }
    }

    $normalized = @{}
    foreach ($monthKey in $result.Keys) {
        $currency = ''
        if ($result[$monthKey].Currencies.Count -gt 0) {
            $currency = ($result[$monthKey].Currencies.Keys | Sort-Object) -join ','
        }

        $normalized[$monthKey] = [PSCustomObject]@{
            Cost = $result[$monthKey].Cost
            Currency = $currency
        }
    }

    return $normalized
}

# Module bootstrap
Require-Module -Name Az.Accounts
Require-Module -Name Az.Resources

# Azure sign-in / context reuse
Write-Host 'Preparing Azure context...'
try {
    Update-AzConfig -LoginExperienceV2 Off -Scope Process -ErrorAction Stop | Out-Null
}
catch {
}

if ($ForceLogin) {
    Write-Host 'ForceLogin enabled. Signing in to Azure...'
    Connect-AzAccount -SkipContextPopulation -ErrorAction Stop | Out-Null
}
else {
    try {
        Enable-AzContextAutosave -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
    }

    $existingContext = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $existingContext) {
        $availableContexts = @(Get-AzContext -ListAvailable -ErrorAction SilentlyContinue)
        if ($availableContexts.Count -gt 0) {
            try {
                Set-AzContext -Context $availableContexts[0] -ErrorAction Stop | Out-Null
                $existingContext = Get-AzContext -ErrorAction SilentlyContinue
            }
            catch {
            }
        }
    }

    if ($existingContext) {
        Write-Host "Using existing Azure context for account '$($existingContext.Account.Id)'."
    }
    else {
        Write-Host 'No saved Azure context found. Signing in to Azure...'
        Connect-AzAccount -SkipContextPopulation -ErrorAction Stop | Out-Null
    }
}

$initializedContext = Get-AzContext -ErrorAction SilentlyContinue
if ($null -eq $initializedContext -or $null -eq $initializedContext.Account) {
    throw 'No active Azure context is available after context initialization.'
}

# Resolve subscription scope:
# - If -SubscriptionId is provided, use it directly (no interactive selection).
# - Otherwise list active subscriptions where current principal is Owner and prompt if needed.
$inactiveStates = @('Disabled', 'Deleted', 'Expired')
$currentAccount = (Get-AzContext -ErrorAction Stop).Account

if (-not [string]::IsNullOrWhiteSpace($SubscriptionId)) {
    $requestedSubscriptionId = $SubscriptionId.Trim()

    try {
        $selectedSubscription = Get-AzSubscription -SubscriptionId $requestedSubscriptionId -ErrorAction Stop
    }
    catch {
        throw "Subscription '$requestedSubscriptionId' was not found or is not accessible for this account."
    }

    if ($inactiveStates -contains [string]$selectedSubscription.State) {
        throw "Subscription '$($selectedSubscription.Name)' ($($selectedSubscription.Id)) is not active. Current state: $($selectedSubscription.State)."
    }

    try {
        $isOwner = Test-OwnerAccessForSubscription -SubscriptionId $selectedSubscription.Id -Account $currentAccount
    }
    catch {
        throw "Could not verify Owner role for subscription '$($selectedSubscription.Name)' ($($selectedSubscription.Id)): $($_.Exception.Message)"
    }

    if (-not $isOwner) {
        throw "Account '$($currentAccount.Id)' does not have Owner role on subscription '$($selectedSubscription.Name)' ($($selectedSubscription.Id))."
    }
}
else {
    $activeSubscriptions = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $inactiveStates -notcontains $_.State })
    if (-not $activeSubscriptions) {
        throw "No active subscriptions were found for this account. Excluded states: $($inactiveStates -join ', ')."
    }

    $ownerSubscriptions = @()
    for ($i = 0; $i -lt $activeSubscriptions.Count; $i++) {
        $item = $activeSubscriptions[$i]
        $percent = [int]((($i + 1) * 100) / $activeSubscriptions.Count)
        Write-Progress -Activity 'Filtering subscriptions by Owner role' -Status "$($item.Name) ($($i + 1)/$($activeSubscriptions.Count))" -CurrentOperation "Checking Owner access for $($currentAccount.Id)" -PercentComplete $percent

        try {
            if (Test-OwnerAccessForSubscription -SubscriptionId $item.Id -Account $currentAccount) {
                $ownerSubscriptions += $item
            }
        }
        catch {
            Write-Warning "Could not verify Owner role for subscription '$($item.Name)' ($($item.Id)): $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity 'Filtering subscriptions by Owner role' -Completed

    # Interactive subscription selection (if multiple candidates remain).
    $activeSubscriptions = @($ownerSubscriptions)
    if (-not $activeSubscriptions) {
        throw "No active subscriptions found where account '$($currentAccount.Id)' has Owner role."
    }

    if ($activeSubscriptions.Count -eq 1) {
        $selectedSubscription = $activeSubscriptions[0]
    }
    else {
        Write-Host ''
        Write-Host "Active subscriptions where '$($currentAccount.Id)' is Owner (excluded states: $($inactiveStates -join ', ')):"
        for ($i = 0; $i -lt $activeSubscriptions.Count; $i++) {
            $item = $activeSubscriptions[$i]
            Write-Host ("[{0}] {1} ({2}, State:{3})" -f ($i + 1), $item.Name, $item.Id, $item.State)
        }

        $selectedIndex = $null
        while ($null -eq $selectedIndex) {
            $raw = Read-Host ("Select active subscription (1-{0})" -f $activeSubscriptions.Count)
            $parsed = 0
            if ([int]::TryParse($raw, [ref]$parsed) -and $parsed -ge 1 -and $parsed -le $activeSubscriptions.Count) {
                $selectedIndex = $parsed - 1
            }
            else {
                Write-Warning 'Invalid selection. Enter a number from the list.'
            }
        }

        $selectedSubscription = $activeSubscriptions[$selectedIndex]
    }
}

# Set context explicitly to the chosen subscription for all subsequent calls.
if (-not [string]::IsNullOrWhiteSpace([string]$selectedSubscription.TenantId)) {
    Set-AzContext -SubscriptionId $selectedSubscription.Id -TenantId $selectedSubscription.TenantId -ErrorAction Stop | Out-Null
}
else {
    Set-AzContext -SubscriptionId $selectedSubscription.Id -ErrorAction Stop | Out-Null
}

Write-Host "Scope: selected subscription only: $($selectedSubscription.Name) ($($selectedSubscription.Id))."

# Parse requested year range and trim future years/months from effective scope.
$parsedYearRange = Parse-YearRange -Value $YearRange
$startYear = $parsedYearRange.StartYear
$endYear = $parsedYearRange.EndYear
$utcNow = [DateTime]::UtcNow
$currentYear = $utcNow.Year
$currentMonthStartUtc = [DateTime]::new($utcNow.Year, $utcNow.Month, 1, 0, 0, 0, [DateTimeKind]::Utc)
$effectiveEndYear = [Math]::Min($endYear, $currentYear)

if ($startYear -gt $currentYear) {
    Write-Host ''
    Write-Host "Requested range '$YearRange' is fully in the future. Current UTC month is $($currentMonthStartUtc.ToString('yyyy-MM')). Nothing to query."
    Write-Host 'Done.'
    return
}

# Build requested month keys (YYYY-MM), excluding any future months.
$monthStarts = @(
    for ($year = $startYear; $year -le $effectiveEndYear; $year++) {
        for ($month = 1; $month -le 12; $month++) {
            $monthStart = [DateTime]::new($year, $month, 1, 0, 0, 0, [DateTimeKind]::Utc)
            if ($monthStart -le $currentMonthStartUtc) {
                $monthStart
            }
        }
    }
)

if (-not $monthStarts) {
    Write-Host ''
    Write-Host "No non-future months found in requested range '$YearRange'. Current UTC month is $($currentMonthStartUtc.ToString('yyyy-MM'))."
    Write-Host 'Done.'
    return
}

$rangeStart = $monthStarts[0].ToString('yyyy-MM')
$rangeEnd = $monthStarts[$monthStarts.Count - 1].ToString('yyyy-MM')
$periodLabel = if ($startYear -eq $effectiveEndYear) { "calendar year $startYear" } else { "years ${startYear}:$effectiveEndYear" }

Write-Host 'Data source: Cost Management Query (single yearly request; aggregated to month in script).'
if ($effectiveEndYear -lt $endYear) {
    Write-Host "Requested period: years ${startYear}:$endYear"
    Write-Host "Effective period (future years skipped): $periodLabel ($rangeStart to $rangeEnd)."
}
else {
    Write-Host "Period: $periodLabel ($rangeStart to $rangeEnd)."
}

# Query one request per year in the effective range and merge results.
try {
    $monthTotals = @{}
    $totalYears = $effectiveEndYear - $startYear + 1

    for ($year = $startYear; $year -le $effectiveEndYear; $year++) {
        $yearNumber = ($year - $startYear + 1)
        $yearStart = [DateTime]::new($year, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
        if ($year -eq $currentYear) {
            $yearEnd = $currentMonthStartUtc.AddMonths(1).AddSeconds(-1)
        }
        else {
            $yearEnd = $yearStart.AddYears(1).AddSeconds(-1)
        }
        $yearRangeStart = $yearStart.ToString('yyyy-MM')
        $yearRangeEnd = $yearEnd.ToString('yyyy-MM')
        $progressPercent = [int]((($yearNumber - 1) * 100) / $totalYears)

        Write-Progress -Activity 'Querying billing months' -Status "Querying year $year ($yearNumber/$totalYears)" -CurrentOperation "$yearRangeStart..$yearRangeEnd" -PercentComplete $progressPercent
        $yearTotals = Get-BillingMonthTotals -SubscriptionId $selectedSubscription.Id -FromUtc $yearStart -ToUtc $yearEnd

        # Merge year result into final month map.
        foreach ($monthKey in $yearTotals.Keys) {
            $monthTotals[$monthKey] = $yearTotals[$monthKey]
        }

        if ($yearNumber -lt $totalYears) {
            Start-Sleep -Seconds 3
        }
    }

    Write-Progress -Activity 'Querying billing months' -Status 'Aggregating monthly totals' -CurrentOperation "$rangeStart..$rangeEnd" -PercentComplete 90
}
catch {
    throw "Failed to query monthly totals: $($_.Exception.Message)"
}
finally {
    Write-Progress -Activity 'Querying billing months' -Completed
}

# Render final monthly table in requested range order.
$rows = @()
foreach ($monthStart in $monthStarts) {
    $monthKey = $monthStart.ToString('yyyy-MM')

    if ($monthTotals.ContainsKey($monthKey)) {
        $monthData = $monthTotals[$monthKey]
        $cost = $monthData.Cost
        $currency = $monthData.Currency
    }
    else {
        $cost = [decimal]0
        $currency = ''
    }

    $rows += [PSCustomObject]@{
        Month = $monthKey
        Cost = $cost
        Currency = $currency
    }
}

# Apply requested month sort order for output/export.
if ($SortOrder -eq 'desc') {
    $rows = @($rows | Sort-Object Month -Descending)
}
else {
    $rows = @($rows | Sort-Object Month)
}

# Summary and guard for empty activity.
$totalCost = [decimal](($rows.Cost | Measure-Object -Sum).Sum)
if ($totalCost -eq 0) {
    Write-Host ''
    Write-Host "No activity found for the selected subscription in $periodLabel (all monthly totals are zero)."
    Write-Host 'Done.'
    return
}

# Optional warning if totals include multiple currencies in a month.
if ($ShowCurrency) {
    $multiCurrencyMonths = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Currency) -and $_.Currency.Contains(',') })
    if ($multiCurrencyMonths.Count -gt 0) {
        Write-Warning 'One or more months contain multiple currencies. Cost is shown as a simple sum across currencies.'
    }
}

Write-Host ''
Write-Host "Subscription: $($selectedSubscription.Name)"
Write-Host "Id:           $($selectedSubscription.Id)"

$tableColumns = @('Month', 'Cost')
if ($ShowCurrency) {
    $tableColumns += 'Currency'
}
$rows | Format-Table -Property $tableColumns -AutoSize

if ($CopyForExcel) {
    try {
        $tsvLines = $rows | Select-Object $tableColumns | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation
        ($tsvLines -join [Environment]::NewLine) | Set-Clipboard
        Write-Host ''
        Write-Host 'Tab-delimited data copied to clipboard. Paste directly into Excel.'
    }
    catch {
        Write-Warning "Could not copy results to clipboard: $($_.Exception.Message)"
    }
}

Write-Host ''
Write-Host 'Done.'
