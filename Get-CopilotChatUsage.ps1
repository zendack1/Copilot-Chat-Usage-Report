<#
.SYNOPSIS
    Reports all users who used Copilot Chat (free / unlicensed) over a configurable
    window with engagement tiers, active-day counts, and department breakdowns -
    actionable data for license-expansion decisions.

.DESCRIPTION
    Queries the Microsoft 365 Unified Audit Log via the Microsoft Graph beta
    'auditLogQuery' API (no 1,000-row Admin Center cap, no 5,000-row
    Search-UnifiedAuditLog cap) for 'CopilotInteraction' events.

    Cross-references each active user against assigned licenses to distinguish
    free Copilot Chat usage from paid Microsoft 365 Copilot usage, and enriches
    each row with directory properties (department, job title, office) so admins
    can target conversations by org segment.

    Output:
      - CSV with one row per user (UPN, interactions, active days, engagement
        tier, first/last use, app hosts, department, job title, office,
        license flag).
      - Self-contained interactive HTML report with KPI tiles, top-20 bar chart,
        app-host distribution chart, engagement histogram, and a sortable /
        searchable table.

.PREREQUISITES
    1. PowerShell 7+ (5.1 works for the script body but ValidateSet behavior
       differs slightly).
    2. Modules:
         Install-Module Microsoft.Graph      -Scope CurrentUser
         Install-Module Microsoft.Graph.Beta -Scope CurrentUser
    3. Delegated runs: signed-in account holds Global Admin / Global Reader /
       Security Admin / Security Reader / Audit Log Reader, with delegated
       scopes 'AuditLogsQuery.Read.All', 'User.Read.All',
       'Organization.Read.All' (consented on first run).
    4. Unattended runs: app registration with the same three scopes as
       Application permissions, admin-consented; pass -ClientId, -TenantId,
       -CertificateThumbprint.

.PARAMETER Days
    Lookback window in days. Must be 30, 60, or 90. Default 30.

.PARAMETER OutputPath
    Path for the CSV output. Default: .\CopilotChatUsage_<utc-timestamp>_utc.csv

.PARAMETER HtmlPath
    Path for the HTML report. Defaults to OutputPath with .html extension.

.PARAMETER IncludeLicensed
    Include users who hold a Microsoft 365 Copilot license. By default they
    are excluded so the result matches the "free Copilot Chat" definition.

.PARAMETER NoHtml
    Skip HTML report generation; emit CSV only.

.PARAMETER ClientId
    App-only auth: client (application) id of the registered app. Triggers
    the AppOnly parameter set; -TenantId and -CertificateThumbprint are also
    required.

.PARAMETER TenantId
    App-only auth: tenant id (GUID or domain).

.PARAMETER CertificateThumbprint
    App-only auth: thumbprint of the certificate stored in CurrentUser\My
    or LocalMachine\My.

.PARAMETER CopilotSkuPattern
    Wildcard pattern matched against SkuPartNumber to identify paid Copilot
    SKUs. Default '*Copilot*' catches Microsoft_365_Copilot, Copilot_Pro,
    Microsoft_Copilot_for_Sales, Microsoft_Copilot_for_Service, etc.

.PARAMETER QueryTimeoutMinutes
    How long to wait for the Graph audit query to finish. Default 60.

.PARAMETER TestConnection
    Verify sign-in, scopes, and SKU enumeration without submitting the audit
    query. Useful before scheduling unattended runs.

.PARAMETER SkipDirectoryEnrichment
    Skip per-user Get-MgUser lookups. Faster on very large tenants, but the
    report won't include department, job title, or office location.

.EXAMPLE
    .\Get-CopilotChatUsage.ps1
    Last 30 days, free Copilot Chat users only, CSV + HTML.

.EXAMPLE
    .\Get-CopilotChatUsage.ps1 -Days 90 -IncludeLicensed
    Every Copilot Chat user (licensed and unlicensed) in the last 90 days.

.EXAMPLE
    .\Get-CopilotChatUsage.ps1 -ClientId 'aaaaaaaa-...' -TenantId 'contoso.onmicrosoft.com' -CertificateThumbprint 'AB12...' -Days 30
    Unattended run for a scheduled task.

.EXAMPLE
    .\Get-CopilotChatUsage.ps1 -TestConnection
    Verifies auth, scopes, and SKU access; submits no audit query.

.NOTES
    Audit records can take up to ~30 minutes to surface after the actual
    interaction, so a run "right now" may slightly under-count today.
#>

[CmdletBinding(DefaultParameterSetName = 'Delegated')]
param(
    [ValidateSet(30, 60, 90)]
    [int]$Days = 30,

    [string]$OutputPath = ".\CopilotChatUsage_$((Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss'))_utc.csv",

    [string]$HtmlPath,

    [switch]$IncludeLicensed,

    [switch]$NoHtml,

    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$TenantId,

    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$CertificateThumbprint,

    [string]$CopilotSkuPattern = '*Copilot*',

    [int]$QueryTimeoutMinutes = 60,

    [switch]$TestConnection,

    [switch]$SkipDirectoryEnrichment
)

if (-not $HtmlPath) {
    $HtmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')
}

$ErrorActionPreference = 'Stop'

# ---------- Helpers ----------

function Connect-Graph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    $requiredScopes = @(
        'AuditLogsQuery.Read.All',
        'User.Read.All',
        'Organization.Read.All'
    )

    if ($PSCmdlet.ParameterSetName -eq 'AppOnly') {
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId `
            -CertificateThumbprint $CertificateThumbprint -NoWelcome
        Write-Host "  Connected app-only as $ClientId to tenant $TenantId" -ForegroundColor Green
    } else {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        $ctx = Get-MgContext
        Write-Host "  Connected as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Green

        $missing = $requiredScopes | Where-Object { $_ -notin $ctx.Scopes }
        if ($missing) {
            throw "Missing required scopes: $($missing -join ', '). Re-run after consent."
        }
    }
}

function Invoke-GraphWithRetry {
    param(
        [string]$Method = 'GET',
        [Parameter(Mandatory)] [string]$Uri,
        [object]$Body,
        [string]$ContentType = 'application/json',
        [int]$MaxAttempts = 5
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($PSBoundParameters.ContainsKey('Body') -and $Body) {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType $ContentType
            } else {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri
            }
        } catch {
            $code = 0
            $retryAfter = 0
            try {
                $resp = $_.Exception.Response
                if ($resp) {
                    $code = [int]$resp.StatusCode
                    if ($resp.Headers -and $resp.Headers.Contains('Retry-After')) {
                        $ra = $resp.Headers.GetValues('Retry-After') | Select-Object -First 1
                        [void][int]::TryParse($ra, [ref]$retryAfter)
                    }
                }
            } catch { }

            $retryable = $code -in 408, 429, 500, 502, 503, 504
            if (-not $retryable -or $attempt -ge $MaxAttempts) { throw }
            $delay = if ($retryAfter -gt 0) { $retryAfter } else { [Math]::Min(60, [Math]::Pow(2, $attempt)) }
            Write-Verbose "Graph $code on attempt $attempt; sleeping ${delay}s and retrying."
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-CopilotLicensedUserIds {
    param([string]$SkuPattern = '*Copilot*')

    Write-Host "Identifying users with a paid Copilot license (SKU pattern: $SkuPattern)..." -ForegroundColor Cyan

    $skus = Get-MgSubscribedSku -All
    $matchedSkus = $skus | Where-Object { $_.SkuPartNumber -like $SkuPattern }

    if (-not $matchedSkus) {
        Write-Host "  No paid Copilot SKUs found matching '$SkuPattern' in this tenant." -ForegroundColor Yellow
        return @{}
    }

    Write-Host ("  Matched {0} SKU(s):" -f @($matchedSkus).Count) -ForegroundColor Cyan
    foreach ($s in $matchedSkus) {
        Write-Host ("    - {0}" -f $s.SkuPartNumber) -ForegroundColor Gray
    }

    $licensed = @{}
    foreach ($sku in $matchedSkus) {
        $skuId = $sku.SkuId
        $filter = "assignedLicenses/any(x:x/skuId eq $skuId)"
        $users = Get-MgUser -Filter $filter -All -Property 'Id,UserPrincipalName' `
            -ConsistencyLevel eventual -CountVariable c
        foreach ($u in $users) {
            if ($u.UserPrincipalName) {
                $licensed[$u.UserPrincipalName.ToLowerInvariant()] = $u.Id
            }
        }
    }
    Write-Host "  $($licensed.Count) user(s) currently hold a paid Copilot license." -ForegroundColor Green
    return $licensed
}

function Get-UserDirectoryProperties {
    param([string[]]$Upns)

    $result = @{}
    if (-not $Upns -or $Upns.Count -eq 0) { return $result }

    Write-Host ("Enriching {0} user(s) with directory properties..." -f $Upns.Count) -ForegroundColor Cyan

    $batchSize = 15
    $total = $Upns.Count
    $processed = 0
    for ($i = 0; $i -lt $total; $i += $batchSize) {
        $end = [Math]::Min($i + $batchSize - 1, $total - 1)
        $batch = $Upns[$i..$end]
        $quoted = ($batch | ForEach-Object { "'$($_ -replace "'", "''")'" }) -join ','
        $filter = "userPrincipalName in ($quoted)"
        try {
            $users = Get-MgUser -Filter $filter -ConsistencyLevel eventual -CountVariable c `
                -Property 'UserPrincipalName,DisplayName,Department,JobTitle,OfficeLocation' -All
            foreach ($u in $users) {
                if ($u.UserPrincipalName) {
                    $result[$u.UserPrincipalName.ToLowerInvariant()] = [PSCustomObject]@{
                        DisplayName    = $u.DisplayName
                        Department     = $u.Department
                        JobTitle       = $u.JobTitle
                        OfficeLocation = $u.OfficeLocation
                    }
                }
            }
        } catch {
            Write-Warning "Directory lookup failed for batch starting at index $i. Continuing without enrichment for that batch."
        }
        $processed += $batch.Count
        Write-Verbose "  Enriched $processed / $total"
    }
    Write-Host "  Resolved directory data for $($result.Count) of $total user(s)." -ForegroundColor Green
    return $result
}

function Get-EngagementTier {
    param([int]$Interactions, [int]$ActiveDays)
    if ($Interactions -ge 100 -or $ActiveDays -ge 15) { return 'Power' }
    if ($Interactions -ge 20  -or $ActiveDays -ge 5)  { return 'Regular' }
    if ($Interactions -ge 2)                          { return 'Casual' }
    return 'OneTime'
}

function Get-Median {
    param([int[]]$Values)
    if (-not $Values -or $Values.Count -eq 0) { return 0 }
    $sorted = @($Values | Sort-Object)
    $n = $sorted.Count
    if ($n % 2 -eq 0) {
        return [Math]::Round(($sorted[$n/2 - 1] + $sorted[$n/2]) / 2.0, 1)
    }
    return $sorted[[Math]::Floor($n/2)]
}

function Submit-AuditQuery {
    param(
        [datetime]$Start,
        [datetime]$End
    )

    Write-Host "Submitting Unified Audit Log query for $($Start.ToString('u')) to $($End.ToString('u'))..." -ForegroundColor Cyan

    $body = @{
        '@odata.type'        = '#microsoft.graph.security.auditLogQuery'
        displayName          = "Copilot Chat usage $($Start.ToString('yyyyMMdd'))-$($End.ToString('yyyyMMdd'))"
        filterStartDateTime  = $Start.ToString('yyyy-MM-ddTHH:mm:ssZ')
        filterEndDateTime    = $End.ToString('yyyy-MM-ddTHH:mm:ssZ')
        operationFilters     = @('CopilotInteraction')
    } | ConvertTo-Json -Depth 5

    $resp = Invoke-GraphWithRetry -Method POST `
        -Uri 'https://graph.microsoft.com/beta/security/auditLog/queries' `
        -Body $body -ContentType 'application/json'

    Write-Host "  Query id: $($resp.id) - status: $($resp.status)" -ForegroundColor Gray
    return $resp.id
}

function Wait-AuditQuery {
    param(
        [string]$QueryId,
        [int]$TimeoutMinutes = 60
    )

    Write-Host "Waiting for query to complete (up to $TimeoutMinutes min)..." -ForegroundColor Cyan
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ((Get-Date) -lt $deadline) {
        $r = Invoke-GraphWithRetry -Method GET `
            -Uri "https://graph.microsoft.com/beta/security/auditLog/queries/$QueryId"
        $status = $r.status
        Write-Verbose ("  status = {0}" -f $status)
        if ($status -eq 'succeeded') { return }
        if ($status -in 'failed', 'cancelled') {
            throw "Audit query ended with status '$status'."
        }
        Start-Sleep -Seconds 30
    }
    throw "Audit query did not complete within $TimeoutMinutes minutes."
}

function Get-AuditQueryRecords {
    param([string]$QueryId)

    Write-Host "Downloading audit records (paging through all results)..." -ForegroundColor Cyan
    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$QueryId/records?`$top=999"
    $all = New-Object System.Collections.Generic.List[object]
    $page = 0
    while ($uri) {
        $page++
        $r = Invoke-GraphWithRetry -Method GET -Uri $uri
        if ($r.value) { $all.AddRange([object[]]$r.value) }
        Write-Verbose ("  page {0}: {1} records (running total: {2})" -f $page, ($r.value).Count, $all.Count)
        $uri = $r.'@odata.nextLink'
    }
    Write-Host "  Total audit records retrieved: $($all.Count)" -ForegroundColor Green
    return $all
}

function Write-HtmlReport {
    param(
        [Parameter(Mandatory)] $Rows,
        [Parameter(Mandatory)] [string]$OutPath,
        [Parameter(Mandatory)] [datetime]$StartUtc,
        [Parameter(Mandatory)] [datetime]$EndUtc,
        [switch]$IncludeLicensed
    )

    $rowsArray = @($Rows)
    $totalUsers  = $rowsArray.Count
    $totalEvents = ($rowsArray | Measure-Object InteractionCount -Sum).Sum
    if (-not $totalEvents) { $totalEvents = 0 }
    $avgPerUser  = if ($totalUsers) { [Math]::Round($totalEvents / $totalUsers, 1) } else { 0 }
    $medianPerUser = Get-Median -Values @($rowsArray | ForEach-Object { [int]$_.InteractionCount })

    # Engagement histogram buckets.
    $tierCounts = @{
        Power   = 0
        Regular = 0
        Casual  = 0
        OneTime = 0
    }
    foreach ($r in $rowsArray) { $tierCounts[$r.EngagementTier]++ }
    $tiers = @(
        @{ key = 'Power';   label = 'Power (100+ events or 15+ active days)'; count = $tierCounts.Power },
        @{ key = 'Regular'; label = 'Regular (20-99 events or 5-14 days)';    count = $tierCounts.Regular },
        @{ key = 'Casual';  label = 'Casual (2-19 events)';                   count = $tierCounts.Casual },
        @{ key = 'OneTime'; label = 'One-time (1 event)';                     count = $tierCounts.OneTime }
    )

    # App host distribution (count distinct users per host).
    $hostCounts = @{}
    foreach ($r in $rowsArray) {
        if ($r.AppHostsUsed) {
            foreach ($h in $r.AppHostsUsed.Split(';')) {
                $h = $h.Trim()
                if ($h) {
                    if (-not $hostCounts.ContainsKey($h)) { $hostCounts[$h] = 0 }
                    $hostCounts[$h]++
                }
            }
        }
    }
    $hosts = @(
        $hostCounts.GetEnumerator() |
            Sort-Object Value -Descending |
            ForEach-Object { @{ label = $_.Key; count = $_.Value } }
    )

    $generatedAt = (Get-Date).ToUniversalTime().ToString('u')
    $scopeText   = if ($IncludeLicensed) {
        'All Copilot Chat users (licensed and unlicensed)'
    } else {
        'Free / unlicensed Copilot Chat users only'
    }
    $titleRange  = "$($StartUtc.ToString('yyyy-MM-dd')) to $($EndUtc.ToString('yyyy-MM-dd'))"

    $payload = @{
        rows            = $rowsArray
        startUtc        = $StartUtc.ToString('u')
        endUtc          = $EndUtc.ToString('u')
        generatedAt     = $generatedAt
        scopeText       = $scopeText
        totalUsers      = $totalUsers
        totalEvents     = $totalEvents
        avgPerUser      = $avgPerUser
        medianPerUser   = $medianPerUser
        tiers           = $tiers
        hosts           = $hosts
        titleRange      = $titleRange
        includeLicensed = [bool]$IncludeLicensed
    } | ConvertTo-Json -Depth 6 -Compress

    $template = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Copilot Chat Usage Report - __TITLE_RANGE__</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  :root {
    --bg:#f4f5f7;
    --card:#ffffff;
    --ink:#0b1320;
    --ink-2:#283041;
    --muted:#5a6473;
    --muted-2:#8a93a3;
    --border:#e5e8ee;
    --border-2:#eef0f4;
    --accent:#0f6cbd;
    --accent-soft:#e8f0f9;
    --shadow-sm:0 1px 2px rgba(15,23,42,.04), 0 1px 1px rgba(15,23,42,.02);
    --shadow-md:0 1px 3px rgba(15,23,42,.05), 0 4px 16px rgba(15,23,42,.04);
    --power:#0f4a82;
    --regular:#2076c4;
    --casual:#7eb1de;
    --onetime:#c4d5e6;
    --radius:10px;
    --radius-sm:6px;
  }
  * { box-sizing:border-box; }
  html, body { margin:0; padding:0; }
  body {
    font:14px/1.55 "Segoe UI Variable","Segoe UI",-apple-system,BlinkMacSystemFont,"Helvetica Neue",Arial,sans-serif;
    color:var(--ink);
    background:var(--bg);
    -webkit-font-smoothing:antialiased;
    text-rendering:optimizeLegibility;
  }
  .accent-bar { height:4px; background:linear-gradient(90deg,#0f4a82 0%,#0f6cbd 50%,#2899f5 100%); }
  header.report-head {
    background:var(--card);
    border-bottom:1px solid var(--border);
    padding:32px 40px 28px;
  }
  header.report-head .inner { max-width:1280px; margin:0 auto; }
  header.report-head .eyebrow {
    font-size:11px; font-weight:600; letter-spacing:.14em; text-transform:uppercase;
    color:var(--accent); margin-bottom:10px;
  }
  header.report-head h1 {
    margin:0 0 6px; font-size:28px; font-weight:600; letter-spacing:-.015em;
    color:var(--ink);
  }
  header.report-head .sub { font-size:13.5px; color:var(--muted); }
  header.report-head .meta-strip {
    margin-top:22px; display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
    gap:0; border-top:1px solid var(--border-2);
  }
  header.report-head .meta-strip .meta-item {
    padding:14px 0 0; padding-right:24px;
  }
  header.report-head .meta-strip .meta-item + .meta-item { padding-left:24px; border-left:1px solid var(--border-2); }
  header.report-head .meta-label {
    font-size:11px; font-weight:600; letter-spacing:.08em; text-transform:uppercase;
    color:var(--muted-2); margin-bottom:4px;
  }
  header.report-head .meta-value {
    font-size:13px; color:var(--ink-2); font-variant-numeric:tabular-nums;
  }
  main { max-width:1280px; margin:0 auto; padding:28px 40px 64px; }
  .section-label {
    font-size:11px; font-weight:600; letter-spacing:.12em; text-transform:uppercase;
    color:var(--muted-2); margin:24px 0 12px;
  }
  .section-label:first-child { margin-top:0; }
  .kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:16px; }
  .kpi {
    background:var(--card); border:1px solid var(--border);
    border-radius:var(--radius); padding:18px 22px 20px;
    box-shadow:var(--shadow-sm); position:relative; overflow:hidden;
  }
  .kpi::before {
    content:""; position:absolute; left:0; top:0; bottom:0; width:3px;
    background:var(--accent); opacity:.85;
  }
  .kpi .label {
    font-size:11px; font-weight:600; color:var(--muted);
    text-transform:uppercase; letter-spacing:.08em; margin-bottom:8px;
  }
  .kpi .value {
    font-size:30px; font-weight:600; letter-spacing:-.02em;
    color:var(--ink); font-variant-numeric:tabular-nums; line-height:1.1;
  }
  .kpi .sub { font-size:11.5px; color:var(--muted); margin-top:4px; }
  .panel {
    background:var(--card); border:1px solid var(--border);
    border-radius:var(--radius); padding:22px 26px;
    box-shadow:var(--shadow-sm); margin-bottom:18px;
  }
  .panel-head {
    display:flex; align-items:baseline; justify-content:space-between;
    margin:0 0 16px; gap:16px; flex-wrap:wrap;
  }
  .panel-head h2 {
    margin:0; font-size:15px; font-weight:600; color:var(--ink); letter-spacing:-.005em;
  }
  .panel-head .panel-sub { font-size:12px; color:var(--muted); }
  .grid-2 { display:grid; grid-template-columns:1fr 1fr; gap:18px; margin-bottom:18px; }
  @media (max-width:920px) { .grid-2 { grid-template-columns:1fr; } }
  .controls {
    display:flex; gap:10px; align-items:center; margin-bottom:14px; flex-wrap:wrap;
  }
  .controls input[type=search] {
    flex:1; min-width:220px; padding:9px 14px;
    border:1px solid var(--border); border-radius:var(--radius-sm);
    font:inherit; color:var(--ink); background:#fdfdfe;
    transition:border-color .15s, box-shadow .15s;
  }
  .controls input[type=search]:focus {
    outline:none; border-color:var(--accent);
    box-shadow:0 0 0 3px rgba(15,108,189,.12);
  }
  .controls .count {
    color:var(--muted); font-size:12.5px; font-variant-numeric:tabular-nums;
  }
  .btn {
    padding:9px 14px; border:1px solid var(--border); background:#fff;
    border-radius:var(--radius-sm); cursor:pointer; font:inherit; font-size:13px;
    color:var(--ink-2); font-weight:500; transition:background .15s, border-color .15s;
  }
  .btn:hover { background:var(--accent-soft); border-color:#cfdcec; color:var(--accent); }
  .btn:focus-visible {
    outline:none; border-color:var(--accent);
    box-shadow:0 0 0 3px rgba(15,108,189,.18);
  }
  .table-wrap {
    max-height:600px; overflow:auto;
    border:1px solid var(--border); border-radius:var(--radius-sm);
    background:var(--card);
  }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  th, td {
    text-align:left; padding:11px 14px;
    border-bottom:1px solid var(--border-2); vertical-align:middle;
  }
  th {
    font-weight:600; color:var(--muted); font-size:11px;
    text-transform:uppercase; letter-spacing:.06em;
    cursor:pointer; user-select:none; position:sticky; top:0;
    background:#fafbfc; z-index:1;
    border-bottom:1px solid var(--border);
  }
  th:hover { color:var(--ink); }
  th.sorted { color:var(--accent); }
  th.sorted::after { content:" \25BE"; }
  th.sorted.asc::after { content:" \25B4"; }
  tbody tr { transition:background .12s; }
  tbody tr:nth-child(even) { background:#fbfcfd; }
  tbody tr:hover { background:var(--accent-soft); }
  td.num { text-align:right; font-variant-numeric:tabular-nums; }
  td .upn { font-family:"Segoe UI",sans-serif; color:var(--ink-2); }
  .pill {
    display:inline-block; padding:3px 10px; border-radius:999px;
    font-size:11px; font-weight:600; letter-spacing:.02em;
    line-height:1.5;
  }
  .pill.tier-Power   { background:#dde9f7; color:#0f4a82; }
  .pill.tier-Regular { background:#e8f1fa; color:#1a5fa3; }
  .pill.tier-Casual  { background:#eef3f9; color:#475a73; }
  .pill.tier-OneTime { background:#f3f1ec; color:#7c6116; }
  .bar-chart { display:grid; gap:10px; }
  .bar-row {
    display:grid; grid-template-columns:240px 1fr 90px;
    gap:14px; align-items:center; font-size:12.5px;
  }
  .bar-row .name {
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    color:var(--ink-2);
  }
  .bar-row .num {
    text-align:right; font-variant-numeric:tabular-nums;
    color:var(--ink-2); font-weight:500;
  }
  .bar-track { background:var(--border-2); border-radius:999px; height:8px; overflow:hidden; }
  .bar {
    height:100%; background:var(--accent);
    border-radius:999px; transition:width .3s ease-out;
  }
  .bar.tier-Power   { background:var(--power); }
  .bar.tier-Regular { background:var(--regular); }
  .bar.tier-Casual  { background:var(--casual); }
  .bar.tier-OneTime { background:var(--onetime); }
  footer.report-foot {
    color:var(--muted); font-size:12px; margin-top:32px; line-height:1.7;
    border-top:1px solid var(--border-2); padding-top:18px;
  }
  footer .note {
    display:flex; gap:10px; align-items:flex-start;
    background:#fbf6e9; border:1px solid #ead9a9; color:#6b5418;
    padding:11px 14px; border-radius:var(--radius-sm); margin-bottom:12px;
    font-size:12.5px;
  }
  footer .note::before {
    content:""; flex:0 0 14px; height:14px; margin-top:2px;
    background:#b8901c; border-radius:50%;
    -webkit-mask:radial-gradient(circle, transparent 4px, #000 4px) center/contain;
            mask:radial-gradient(circle, transparent 4px, #000 4px) center/contain;
  }
  .empty { padding:40px; text-align:center; color:var(--muted); font-size:13px; }
  @media (max-width:680px) {
    header.report-head { padding:24px 20px 22px; }
    main { padding:20px; }
    .bar-row { grid-template-columns:140px 1fr 70px; }
    header.report-head h1 { font-size:23px; }
    header.report-head .meta-strip { grid-template-columns:1fr 1fr; }
    header.report-head .meta-strip .meta-item + .meta-item { padding-left:16px; }
  }
  @media print {
    body { background:#fff; }
    .accent-bar { background:#0f6cbd !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    header.report-head { padding:16px 0 14px; }
    .panel, .kpi { break-inside:avoid; box-shadow:none; }
    .controls .btn, #csv-btn { display:none; }
    main { max-width:none; padding:12px 0; }
    table { font-size:11px; }
    th { position:static; }
    .table-wrap { max-height:none !important; overflow:visible !important; border:none !important; }
    tbody tr:hover { background:transparent; }
    .kpi::before { background:#0f6cbd !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  }
</style>
</head>
<body>
<div class="accent-bar"></div>
<header class="report-head">
  <div class="inner">
    <div class="eyebrow">Microsoft 365 &middot; Copilot adoption</div>
    <h1>Copilot Chat Usage Report</h1>
    <div class="sub" id="sub"></div>
    <div class="meta-strip" id="meta-strip"></div>
  </div>
</header>
<main>
  <div class="section-label">Overview</div>
  <section class="kpis">
    <div class="kpi"><div class="label">Active users</div><div class="value" id="kpi-users">0</div><div class="sub">in window</div></div>
    <div class="kpi"><div class="label">Total interactions</div><div class="value" id="kpi-events">0</div><div class="sub">across all users</div></div>
    <div class="kpi"><div class="label">Average / user</div><div class="value" id="kpi-avg">0</div><div class="sub">interactions</div></div>
    <div class="kpi"><div class="label">Median / user</div><div class="value" id="kpi-median">0</div><div class="sub">interactions</div></div>
  </section>

  <div class="section-label">Distribution</div>
  <section class="grid-2">
    <div class="panel">
      <div class="panel-head">
        <h2>Engagement tiers</h2>
        <span class="panel-sub">users by activity level</span>
      </div>
      <div class="bar-chart" id="tier-chart"></div>
    </div>
    <div class="panel">
      <div class="panel-head">
        <h2>App host distribution</h2>
        <span class="panel-sub">users per surface</span>
      </div>
      <div class="bar-chart" id="host-chart"></div>
      <div id="host-empty" class="empty" style="display:none">No app host data.</div>
    </div>
  </section>

  <div class="section-label">Top users</div>
  <section class="panel">
    <div class="panel-head">
      <h2>Top 20 by interaction count</h2>
      <span class="panel-sub">highest-volume Copilot Chat users in window</span>
    </div>
    <div class="bar-chart" id="chart"></div>
    <div id="chart-empty" class="empty" style="display:none">No data.</div>
  </section>

  <div class="section-label">Detail</div>
  <section class="panel">
    <div class="panel-head">
      <h2>All users</h2>
      <span class="panel-sub">sortable, searchable, exportable</span>
    </div>
    <div class="controls">
      <input type="search" id="filter" placeholder="Filter by UPN, app host, department, tier&hellip;">
      <span class="count" id="count"></span>
      <button id="csv-btn" class="btn" type="button">Download filtered CSV</button>
    </div>
    <div class="table-wrap">
      <table id="tbl">
        <thead>
          <tr>
            <th data-key="UserPrincipalName">User Principal Name</th>
            <th data-key="EngagementTier">Tier</th>
            <th data-key="InteractionCount" class="num">Events</th>
            <th data-key="ActiveDays" class="num">Active days</th>
            <th data-key="Department">Department</th>
            <th data-key="JobTitle">Job title</th>
            <th data-key="FirstUsedUtc">First used (UTC)</th>
            <th data-key="LastUsedUtc">Last used (UTC)</th>
            <th data-key="AppHostsUsed">App hosts</th>
          </tr>
        </thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>
  </section>

  <footer class="report-foot" id="foot">
    <div class="note">
      Audit records can take up to ~30 minutes to surface. Activity in the last
      30 minutes of the window may be under-counted.
    </div>
    <div id="foot-text"></div>
  </footer>
</main>

<script id="report-data" type="application/json">__DATA_JSON__</script>
<script>
(function(){
  const data = JSON.parse(document.getElementById('report-data').textContent);
  const rows = (data.rows || []).slice();

  const fmt = n => (n==null) ? '' : Number(n).toLocaleString();
  document.getElementById('sub').textContent = data.scopeText;

  // Meta strip in the header.
  const metaItems = [
    { label:'Reporting period', value:`${data.startUtc} → ${data.endUtc}` },
    { label:'Generated', value:data.generatedAt },
    { label:'Scope', value:data.includeLicensed ? 'Licensed + unlicensed' : 'Unlicensed only' },
    { label:'Users in report', value:fmt(data.totalUsers) }
  ];
  function escapeHtml(s){return String(s==null?'':s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
  document.getElementById('meta-strip').innerHTML = metaItems.map(m =>
    `<div class="meta-item"><div class="meta-label">${escapeHtml(m.label)}</div><div class="meta-value">${escapeHtml(m.value)}</div></div>`
  ).join('');

  document.getElementById('kpi-users').textContent  = fmt(data.totalUsers);
  document.getElementById('kpi-events').textContent = fmt(data.totalEvents);
  document.getElementById('kpi-avg').textContent    = fmt(data.avgPerUser);
  document.getElementById('kpi-median').textContent = fmt(data.medianPerUser);

  function renderBars(containerId, items, options){
    options = options || {};
    const max = items.length ? items.reduce((m,i)=>Math.max(m, i.count||0), 0) : 0;
    const el = document.getElementById(containerId);
    if (!items.length) { return false; }
    el.innerHTML = items.map(i => {
      const pct = max ? ((i.count||0) / max * 100) : 0;
      let cls = 'bar';
      if (options.tierKey) { cls += ' tier-' + options.tierKey(i); }
      return `<div class="bar-row">` +
        `<div class="name" title="${escapeHtml(i.label)}">${escapeHtml(i.label)}</div>` +
        `<div class="bar-track"><div class="${cls}" style="width:${pct}%"></div></div>` +
        `<div class="num">${fmt(i.count)}</div>` +
        `</div>`;
    }).join('');
    return true;
  }

  // Engagement tier chart.
  renderBars('tier-chart', data.tiers || [], { tierKey: i => i.key });

  // App host distribution.
  if (!renderBars('host-chart', data.hosts || [])) {
    document.getElementById('host-empty').style.display = 'block';
  }

  // Top-20 users.
  const top = rows.slice().sort((a,b)=>b.InteractionCount-a.InteractionCount).slice(0,20)
    .map(r => ({ label: r.UserPrincipalName, count: r.InteractionCount }));
  if (!renderBars('chart', top)) {
    document.getElementById('chart-empty').style.display = 'block';
  }

  // Table.
  const tbody = document.getElementById('tbody');
  const filter = document.getElementById('filter');
  const countEl = document.getElementById('count');
  let sortKey = 'InteractionCount';
  let sortAsc = false;
  let currentView = rows.slice();

  function tierBadge(t){
    const safe = String(t||'').replace(/[^A-Za-z]/g,'');
    return `<span class="pill tier-${safe}">${escapeHtml(t||'')}</span>`;
  }

  function applyFilter() {
    const q = filter.value.trim().toLowerCase();
    let view = rows;
    if (q) {
      view = rows.filter(r =>
        (r.UserPrincipalName||'').toLowerCase().includes(q) ||
        (r.AppHostsUsed||'').toLowerCase().includes(q) ||
        (r.Department||'').toLowerCase().includes(q) ||
        (r.JobTitle||'').toLowerCase().includes(q) ||
        (r.EngagementTier||'').toLowerCase().includes(q)
      );
    }
    return view.slice().sort((a,b)=>{
      let av=a[sortKey], bv=b[sortKey];
      if (typeof av==='string') av=av.toLowerCase();
      if (typeof bv==='string') bv=bv.toLowerCase();
      if (av==null) av = (typeof bv==='number') ? -Infinity : '';
      if (bv==null) bv = (typeof av==='number') ? -Infinity : '';
      if (av<bv) return sortAsc?-1:1;
      if (av>bv) return sortAsc?1:-1;
      return 0;
    });
  }

  function render() {
    currentView = applyFilter();
    tbody.innerHTML = currentView.map(r => `
      <tr>
        <td><span class="upn">${escapeHtml(r.UserPrincipalName||'')}</span></td>
        <td>${tierBadge(r.EngagementTier)}</td>
        <td class="num">${fmt(r.InteractionCount)}</td>
        <td class="num">${fmt(r.ActiveDays)}</td>
        <td>${escapeHtml(r.Department||'')}</td>
        <td>${escapeHtml(r.JobTitle||'')}</td>
        <td>${escapeHtml(r.FirstUsedUtc||'')}</td>
        <td>${escapeHtml(r.LastUsedUtc||'')}</td>
        <td>${escapeHtml(r.AppHostsUsed||'')}</td>
      </tr>`).join('');
    countEl.textContent = `${fmt(currentView.length)} of ${fmt(rows.length)} shown`;
    document.querySelectorAll('th').forEach(th => {
      th.classList.toggle('sorted', th.dataset.key===sortKey);
      th.classList.toggle('asc', th.dataset.key===sortKey && sortAsc);
    });
  }

  document.querySelectorAll('th[data-key]').forEach(th => {
    th.addEventListener('click', () => {
      const k = th.dataset.key;
      const stringCols = new Set(['UserPrincipalName','AppHostsUsed','Department','JobTitle','EngagementTier','FirstUsedUtc','LastUsedUtc']);
      if (k===sortKey) sortAsc = !sortAsc;
      else { sortKey = k; sortAsc = stringCols.has(k); }
      render();
    });
  });
  filter.addEventListener('input', render);

  document.getElementById('csv-btn').addEventListener('click', () => {
    const cols = ['UserPrincipalName','UserId','EngagementTier','InteractionCount','ActiveDays','Department','JobTitle','OfficeLocation','FirstUsedUtc','LastUsedUtc','AppHostsUsed'];
    const csv = [cols.join(',')].concat(currentView.map(r =>
      cols.map(c => {
        const v = r[c]==null ? '' : String(r[c]);
        return /[,"\n]/.test(v) ? `"${v.replace(/"/g,'""')}"` : v;
      }).join(',')
    )).join('\n');
    const blob = new Blob([csv],{type:'text/csv;charset=utf-8'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'CopilotChatUsage_filtered.csv';
    a.click();
  });

  document.getElementById('foot-text').textContent =
    `Generated ${data.generatedAt} · ${fmt(rows.length)} users · ${fmt(data.totalEvents)} total interactions`;

  render();
})();
</script>
</body>
</html>
'@

    # Inject the JSON payload safely.
    $safe = $payload -replace '</', '<\/'
    $html = $template.Replace('__DATA_JSON__', $safe).Replace('__TITLE_RANGE__', $titleRange)

    Set-Content -Path $OutPath -Value $html -Encoding UTF8
}

# ---------- Main ----------

Connect-Graph

if ($TestConnection) {
    Write-Host ""
    Write-Host "Test-connection mode:" -ForegroundColor Cyan
    $null = Get-CopilotLicensedUserIds -SkuPattern $CopilotSkuPattern
    Write-Host ""
    Write-Host "Connection test passed. All required permissions are in place." -ForegroundColor Green
    return
}

$endUtc   = (Get-Date).ToUniversalTime()
$startUtc = $endUtc.AddDays(-1 * $Days)

$licensedMap = if ($IncludeLicensed) { @{} } else { Get-CopilotLicensedUserIds -SkuPattern $CopilotSkuPattern }

$queryId = Submit-AuditQuery -Start $startUtc -End $endUtc
Wait-AuditQuery -QueryId $queryId -TimeoutMinutes $QueryTimeoutMinutes
$records = Get-AuditQueryRecords -QueryId $queryId

if ($records.Count -eq 0) {
    Write-Host "No Copilot interactions returned for the window." -ForegroundColor Yellow
    return
}

Write-Host "Aggregating per-user interaction counts..." -ForegroundColor Cyan

$rows = foreach ($rec in $records) {
    $data = $rec.auditData
    if ($data -is [string]) {
        try { $data = $data | ConvertFrom-Json } catch { $data = $null }
    }

    $upn      = $data.UserId
    $userKey  = $data.UserKey
    $appHost  = $data.CopilotEventData.AppHost
    if (-not $appHost) { $appHost = $data.AppHost }
    $created  = $rec.createdDateTime
    if (-not $created) { $created = $data.CreationTime }

    $timestamp = $null
    if ($created) {
        try { $timestamp = [datetime]$created } catch { $timestamp = $null }
    }
    if (-not $timestamp) { continue }

    [PSCustomObject]@{
        Upn       = $upn
        UserId    = $userKey
        AppHost   = $appHost
        TimeStamp = $timestamp
    }
}

# Group by UPN to compute per-user metrics.
$grouped = $rows |
    Where-Object { $_.Upn } |
    Group-Object -Property Upn |
    ForEach-Object {
        $g = $_
        $first  = ($g.Group | Sort-Object TimeStamp | Select-Object -First 1).TimeStamp
        $last   = ($g.Group | Sort-Object TimeStamp -Descending | Select-Object -First 1).TimeStamp
        $hosts  = ($g.Group.AppHost | Where-Object { $_ } | Sort-Object -Unique) -join ';'
        $uid    = ($g.Group | Where-Object UserId | Select-Object -First 1).UserId
        $upnKey = $g.Name.ToLowerInvariant()
        $hasLic = $licensedMap.ContainsKey($upnKey)
        $activeDays = @($g.Group.TimeStamp | ForEach-Object { $_.ToUniversalTime().Date } |
                        Sort-Object -Unique).Count
        $tier = Get-EngagementTier -Interactions $g.Count -ActiveDays $activeDays

        [PSCustomObject]@{
            UserPrincipalName  = $g.Name
            UserId             = $uid
            InteractionCount   = $g.Count
            ActiveDays         = $activeDays
            EngagementTier     = $tier
            FirstUsedUtc       = $first.ToString('u')
            LastUsedUtc        = $last.ToString('u')
            AppHostsUsed       = $hosts
            HasCopilotLicense  = $hasLic
            DisplayName        = $null
            Department         = $null
            JobTitle           = $null
            OfficeLocation     = $null
        }
    }

if (-not $IncludeLicensed) {
    $grouped = $grouped | Where-Object { -not $_.HasCopilotLicense }
}

$grouped = @($grouped | Sort-Object InteractionCount -Descending)

# Directory enrichment.
if (-not $SkipDirectoryEnrichment -and $grouped.Count -gt 0) {
    $dirMap = Get-UserDirectoryProperties -Upns ($grouped.UserPrincipalName)
    foreach ($g in $grouped) {
        $key = $g.UserPrincipalName.ToLowerInvariant()
        if ($dirMap.ContainsKey($key)) {
            $info = $dirMap[$key]
            $g.DisplayName    = $info.DisplayName
            $g.Department     = $info.Department
            $g.JobTitle       = $info.JobTitle
            $g.OfficeLocation = $info.OfficeLocation
        }
    }
}

$grouped | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

if (-not $NoHtml) {
    Write-Host "Building HTML report..." -ForegroundColor Cyan
    Write-HtmlReport -Rows $grouped `
                     -OutPath $HtmlPath `
                     -StartUtc $startUtc `
                     -EndUtc $endUtc `
                     -IncludeLicensed:$IncludeLicensed
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host ("  Users in report : {0}" -f $grouped.Count) -ForegroundColor Green
Write-Host ("  Total events    : {0}" -f ($grouped | Measure-Object InteractionCount -Sum).Sum) -ForegroundColor Green
Write-Host ("  CSV file        : {0}" -f (Resolve-Path $OutputPath)) -ForegroundColor Green
if (-not $NoHtml) {
    Write-Host ("  HTML report     : {0}" -f (Resolve-Path $HtmlPath)) -ForegroundColor Green
}
