<#
.SYNOPSIS
    Reports all users who used Copilot Chat (free / unlicensed) in the past 30 days,
    along with their UPN and total interaction count.

.DESCRIPTION
    Queries the Microsoft 365 Unified Audit Log via the Microsoft Graph
    auditLogQuery API (no 1,000-row Admin Center cap and no 5,000-row
    Search-UnifiedAuditLog cap) for "CopilotInteraction" events in the last
    30 days. Cross-references each active user against assigned licenses to
    distinguish "free" Copilot Chat usage (no Microsoft 365 Copilot license)
    from paid M365 Copilot usage.

    Output: CSV with one row per user — UPN, DisplayName, InteractionCount,
    FirstUsed, LastUsed, AppHostsUsed, HasCopilotLicense.

.PREREQUISITES
    1. PowerShell 7+
    2. Microsoft.Graph module:
         Install-Module Microsoft.Graph -Scope CurrentUser
       Microsoft.Graph.Beta module (auditLogQuery is currently beta):
         Install-Module Microsoft.Graph.Beta -Scope CurrentUser
    3. The signed-in account must hold one of:
         - Global Administrator
         - Global Reader
         - Security Administrator / Reader
         - Audit Log Reader
       AND have consented to these Graph scopes:
         - AuditLogsQuery.Read.All
         - User.Read.All
         - Organization.Read.All

.PARAMETER Days
    Lookback window in days. Must be 30, 60, or 90. Default 30.

.PARAMETER OutputPath
    Path for the CSV output. Default: .\CopilotChatUsage_<timestamp>.csv

.PARAMETER IncludeLicensed
    If set, includes users who have a Microsoft 365 Copilot license.
    By default they are excluded so the report matches the "free Copilot Chat"
    definition in the Admin Center.

.EXAMPLE
    .\Get-CopilotChatUsage.ps1
    Runs interactively for the last 30 days, excludes licensed Copilot users.

.EXAMPLE
    .\Get-CopilotChatUsage.ps1 -Days 60
    Runs for the last 60 days.

.EXAMPLE
    .\Get-CopilotChatUsage.ps1 -Days 90 -IncludeLicensed -OutputPath .\AllCopilotUsers.csv
    Reports every Copilot Chat user (licensed and unlicensed) in the last 90 days.

.NOTES
    For unattended / scheduled runs, replace the Connect-MgGraph call with
    an app-only connection:
        Connect-MgGraph -ClientId <appId> -TenantId <tenantId> -CertificateThumbprint <thumb>
    The app registration needs the application permissions:
        AuditLogsQuery.Read.All, User.Read.All, Organization.Read.All
#>

[CmdletBinding()]
param(
    [ValidateSet(30, 60, 90)]
    [int]$Days = 30,
    [string]$OutputPath = ".\CopilotChatUsage_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [string]$HtmlPath,
    [switch]$IncludeLicensed,
    [switch]$NoHtml
)

if (-not $HtmlPath) {
    $HtmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')
}

$ErrorActionPreference = 'Stop'

# Microsoft 365 Copilot license SKU part numbers.
# Add tenant-specific SKUs here if needed (Get-MgSubscribedSku to list).
$CopilotSkuPartNumbers = @(
    'Microsoft_365_Copilot',
    'M365_Copilot',
    'Microsoft_365_Copilot_for_Business'
)

function Connect-Graph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    $requiredScopes = @(
        'AuditLogsQuery.Read.All',
        'User.Read.All',
        'Organization.Read.All'
    )
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    $ctx = Get-MgContext
    Write-Host "Connected as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Green
}

function Get-CopilotLicensedUserIds {
    Write-Host "Identifying users with a Microsoft 365 Copilot license..." -ForegroundColor Cyan

    $skus = Get-MgSubscribedSku -All
    $copilotSkuIds = $skus |
        Where-Object {
            $part = $_.SkuPartNumber
            $CopilotSkuPartNumbers | Where-Object { $part -like "*$_*" }
        } |
        Select-Object -ExpandProperty SkuId

    if (-not $copilotSkuIds -or $copilotSkuIds.Count -eq 0) {
        Write-Host "  No Microsoft 365 Copilot SKUs found in this tenant." -ForegroundColor Yellow
        return @{}
    }

    Write-Host "  Found $($copilotSkuIds.Count) Copilot SKU(s). Enumerating assigned users..." -ForegroundColor Cyan

    # Key the map by UPN (lowercased) — audit-record UserKey is not reliably the
    # Entra object id, but UserId in Copilot audit data is the UPN.
    $licensed = @{}
    foreach ($skuId in $copilotSkuIds) {
        $filter = "assignedLicenses/any(x:x/skuId eq $skuId)"
        $users = Get-MgUser -Filter $filter -All -Property 'Id,UserPrincipalName' -ConsistencyLevel eventual -CountVariable c
        foreach ($u in $users) {
            if ($u.UserPrincipalName) {
                $licensed[$u.UserPrincipalName.ToLowerInvariant()] = $u.Id
            }
        }
    }
    Write-Host "  $($licensed.Count) user(s) currently hold an M365 Copilot license." -ForegroundColor Green
    return $licensed
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

    $resp = Invoke-MgGraphRequest -Method POST `
        -Uri 'https://graph.microsoft.com/beta/security/auditLog/queries' `
        -Body $body -ContentType 'application/json'

    Write-Host "  Query id: $($resp.id) — status: $($resp.status)" -ForegroundColor Gray
    return $resp.id
}

function Wait-AuditQuery {
    param([string]$QueryId)

    Write-Host "Waiting for query to complete (this can take several minutes for 30-day windows)..." -ForegroundColor Cyan
    $deadline = (Get-Date).AddMinutes(60)
    while ((Get-Date) -lt $deadline) {
        $r = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/beta/security/auditLog/queries/$QueryId"
        $status = $r.status
        Write-Host ("  status = {0}" -f $status) -ForegroundColor Gray
        if ($status -eq 'succeeded') { return }
        if ($status -in 'failed','cancelled') {
            throw "Audit query ended with status '$status'."
        }
        Start-Sleep -Seconds 30
    }
    throw "Audit query did not complete within 60 minutes."
}

function Get-AuditQueryRecords {
    param([string]$QueryId)

    Write-Host "Downloading audit records (paging through all results)..." -ForegroundColor Cyan
    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$QueryId/records?`$top=999"
    $all = New-Object System.Collections.Generic.List[object]
    $page = 0
    while ($uri) {
        $page++
        $r = Invoke-MgGraphRequest -Method GET -Uri $uri
        if ($r.value) { $all.AddRange([object[]]$r.value) }
        Write-Host ("  page {0}: {1} records (running total: {2})" -f $page, ($r.value).Count, $all.Count) -ForegroundColor Gray
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
    $licCount    = ($rowsArray | Where-Object HasCopilotLicense).Count
    $unlicCount  = $totalUsers - $licCount

    $generatedAt = (Get-Date).ToString('u')
    $tenant      = (Get-MgContext).TenantId
    $scopeText   = if ($IncludeLicensed) { 'All Copilot Chat users (licensed and unlicensed)' } else { 'Free / unlicensed Copilot Chat users only' }

    # Embed data as JSON for client-side sort/filter.
    $payload = @{
        rows           = $rowsArray
        startUtc       = $StartUtc.ToString('u')
        endUtc         = $EndUtc.ToString('u')
        generatedAt    = $generatedAt
        tenantId       = $tenant
        scopeText      = $scopeText
        totalUsers     = $totalUsers
        totalEvents    = $totalEvents
        licensedCount  = $licCount
        unlicensedCount = $unlicCount
    } | ConvertTo-Json -Depth 6 -Compress

    $template = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Copilot Chat Usage Report</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  :root {
    --bg:#f5f6fa; --card:#fff; --ink:#1f2328; --muted:#5b6573;
    --accent:#0f6cbd; --accent-2:#2899f5; --border:#e3e6ec;
    --good:#107c10; --warn:#a47100;
  }
  * { box-sizing: border-box; }
  body { margin:0; font:14px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
         color:var(--ink); background:var(--bg); }
  header { background:linear-gradient(135deg,#0f6cbd 0%,#2899f5 100%); color:#fff;
           padding:24px 32px; }
  header h1 { margin:0 0 4px; font-size:22px; font-weight:600; letter-spacing:-.01em; }
  header .sub { font-size:13px; opacity:.9; }
  main { max-width:1280px; margin:0 auto; padding:24px 32px 64px; }
  .kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
          gap:16px; margin-bottom:24px; }
  .kpi { background:var(--card); border:1px solid var(--border); border-radius:8px;
         padding:16px 20px; }
  .kpi .label { font-size:12px; color:var(--muted); text-transform:uppercase;
                letter-spacing:.04em; margin-bottom:6px; }
  .kpi .value { font-size:28px; font-weight:600; letter-spacing:-.02em; }
  .panel { background:var(--card); border:1px solid var(--border); border-radius:8px;
           padding:20px 24px; margin-bottom:20px; }
  .panel h2 { margin:0 0 12px; font-size:15px; font-weight:600; }
  .controls { display:flex; gap:12px; align-items:center; margin-bottom:12px;
              flex-wrap:wrap; }
  .controls input[type=search] { flex:1; min-width:200px; padding:8px 12px;
              border:1px solid var(--border); border-radius:6px; font:inherit; }
  .controls .count { color:var(--muted); font-size:13px; }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  th, td { text-align:left; padding:8px 10px; border-bottom:1px solid var(--border);
           vertical-align:top; }
  th { font-weight:600; color:var(--muted); font-size:12px; text-transform:uppercase;
       letter-spacing:.04em; cursor:pointer; user-select:none; position:sticky; top:0;
       background:var(--card); }
  th:hover { color:var(--ink); }
  th.sorted::after { content:" \25BE"; color:var(--accent); }
  th.sorted.asc::after { content:" \25B4"; }
  tbody tr:hover { background:#fafbfc; }
  td.num { text-align:right; font-variant-numeric: tabular-nums; }
  .pill { display:inline-block; padding:2px 8px; border-radius:999px; font-size:11px;
          background:#eef4fb; color:var(--accent); }
  .pill.lic { background:#fef6e7; color:var(--warn); }
  .bar-chart { display:grid; gap:6px; }
  .bar-row { display:grid; grid-template-columns:240px 1fr 60px; gap:8px;
             align-items:center; font-size:12px; }
  .bar-row .name { white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .bar-row .num  { text-align:right; font-variant-numeric: tabular-nums; color:var(--muted); }
  .bar { height:14px; background:linear-gradient(90deg,var(--accent),var(--accent-2));
         border-radius:4px; }
  footer { color:var(--muted); font-size:12px; margin-top:24px; }
  .muted { color:var(--muted); }
  .empty { padding:40px; text-align:center; color:var(--muted); }
  @media (max-width:640px) {
    .bar-row { grid-template-columns:140px 1fr 50px; }
    main { padding:16px; }
  }
</style>
</head>
<body>
<header>
  <h1>Copilot Chat Usage Report</h1>
  <div class="sub" id="sub"></div>
</header>
<main>
  <section class="kpis">
    <div class="kpi"><div class="label">Users</div><div class="value" id="kpi-users">0</div></div>
    <div class="kpi"><div class="label">Total interactions</div><div class="value" id="kpi-events">0</div></div>
    <div class="kpi"><div class="label">Avg / user</div><div class="value" id="kpi-avg">0</div></div>
  </section>

  <section class="panel">
    <h2>Top 20 users by interaction count</h2>
    <div class="bar-chart" id="chart"></div>
    <div id="chart-empty" class="empty" style="display:none">No data.</div>
  </section>

  <section class="panel">
    <h2>All users</h2>
    <div class="controls">
      <input type="search" id="filter" placeholder="Filter by UPN or app host…">
      <span class="count" id="count"></span>
      <button id="csv-btn" type="button" style="padding:8px 12px;border:1px solid var(--border);background:#fff;border-radius:6px;cursor:pointer;">Download CSV</button>
    </div>
    <div style="max-height:560px;overflow:auto;border:1px solid var(--border);border-radius:6px;">
      <table id="tbl">
        <thead>
          <tr>
            <th data-key="UserPrincipalName">User Principal Name</th>
            <th data-key="InteractionCount" class="num">Interactions</th>
            <th data-key="FirstUsedUtc">First used (UTC)</th>
            <th data-key="LastUsedUtc">Last used (UTC)</th>
            <th data-key="AppHostsUsed">App hosts</th>
          </tr>
        </thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>
  </section>

  <footer id="foot"></footer>
</main>

<script id="report-data" type="application/json">__DATA_JSON__</script>
<script>
(function(){
  const data = JSON.parse(document.getElementById('report-data').textContent);
  const rows = (data.rows || []).slice();

  const fmt = n => (n==null) ? '' : Number(n).toLocaleString();
  const subtitle = `${data.scopeText} • ${data.startUtc} → ${data.endUtc} • Tenant ${data.tenantId}`;
  document.getElementById('sub').textContent = subtitle;

  document.getElementById('kpi-users').textContent  = fmt(data.totalUsers);
  document.getElementById('kpi-events').textContent = fmt(data.totalEvents);
  const avg = data.totalUsers ? Math.round(data.totalEvents / data.totalUsers) : 0;
  document.getElementById('kpi-avg').textContent    = fmt(avg);

  // Bar chart - top 20.
  const top = rows.slice().sort((a,b)=>b.InteractionCount-a.InteractionCount).slice(0,20);
  const max = top.length ? top[0].InteractionCount : 0;
  const chart = document.getElementById('chart');
  if (!top.length) {
    document.getElementById('chart-empty').style.display='block';
  } else {
    top.forEach(r => {
      const pct = max ? (r.InteractionCount / max * 100) : 0;
      const row = document.createElement('div');
      row.className = 'bar-row';
      row.innerHTML =
        `<div class="name" title="${r.UserPrincipalName}">${r.UserPrincipalName}</div>` +
        `<div><div class="bar" style="width:${pct}%"></div></div>` +
        `<div class="num">${fmt(r.InteractionCount)}</div>`;
      chart.appendChild(row);
    });
  }

  // Table.
  const tbody = document.getElementById('tbody');
  const filter = document.getElementById('filter');
  const countEl = document.getElementById('count');
  let sortKey = 'InteractionCount';
  let sortAsc = false;

  function render() {
    const q = filter.value.trim().toLowerCase();
    let view = rows;
    if (q) {
      view = rows.filter(r =>
        (r.UserPrincipalName||'').toLowerCase().includes(q) ||
        (r.AppHostsUsed||'').toLowerCase().includes(q)
      );
    }
    view = view.slice().sort((a,b)=>{
      let av=a[sortKey], bv=b[sortKey];
      if (typeof av==='string') av=av.toLowerCase();
      if (typeof bv==='string') bv=bv.toLowerCase();
      if (av<bv) return sortAsc?-1:1;
      if (av>bv) return sortAsc?1:-1;
      return 0;
    });
    tbody.innerHTML = view.map(r => `
      <tr>
        <td>${escapeHtml(r.UserPrincipalName||'')}</td>
        <td class="num">${fmt(r.InteractionCount)}</td>
        <td>${escapeHtml(r.FirstUsedUtc||'')}</td>
        <td>${escapeHtml(r.LastUsedUtc||'')}</td>
        <td>${escapeHtml(r.AppHostsUsed||'')}</td>
      </tr>`).join('');
    countEl.textContent = `${fmt(view.length)} of ${fmt(rows.length)} shown`;
    document.querySelectorAll('th').forEach(th => {
      th.classList.toggle('sorted', th.dataset.key===sortKey);
      th.classList.toggle('asc', th.dataset.key===sortKey && sortAsc);
    });
  }
  function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}

  document.querySelectorAll('th[data-key]').forEach(th => {
    th.addEventListener('click', () => {
      const k = th.dataset.key;
      if (k===sortKey) sortAsc = !sortAsc;
      else { sortKey = k; sortAsc = (k==='UserPrincipalName' || k==='AppHostsUsed'); }
      render();
    });
  });
  filter.addEventListener('input', render);

  document.getElementById('csv-btn').addEventListener('click', () => {
    const cols = ['UserPrincipalName','UserId','InteractionCount','FirstUsedUtc','LastUsedUtc','AppHostsUsed'];
    const csv = [cols.join(',')].concat(rows.map(r =>
      cols.map(c => {
        const v = r[c]==null ? '' : String(r[c]);
        return /[,"\n]/.test(v) ? `"${v.replace(/"/g,'""')}"` : v;
      }).join(',')
    )).join('\n');
    const blob = new Blob([csv],{type:'text/csv;charset=utf-8'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'CopilotChatUsage.csv';
    a.click();
  });

  document.getElementById('foot').textContent = `Generated ${data.generatedAt} • ${fmt(rows.length)} users • ${fmt(data.totalEvents)} interactions`;

  render();
})();
</script>
</body>
</html>
'@

    # Inject the JSON payload safely (no broken </script> sequences in data).
    $safe = $payload -replace '</', '<\/'
    $html = $template.Replace('__DATA_JSON__', $safe)

    Set-Content -Path $OutPath -Value $html -Encoding UTF8
}

# ---------- Main ----------

Connect-Graph

$endUtc   = (Get-Date).ToUniversalTime()
$startUtc = $endUtc.AddDays(-1 * $Days)

$licensedMap = if ($IncludeLicensed) { @{} } else { Get-CopilotLicensedUserIds }

$queryId = Submit-AuditQuery -Start $startUtc -End $endUtc
Wait-AuditQuery -QueryId $queryId
$records = Get-AuditQueryRecords -QueryId $queryId

if ($records.Count -eq 0) {
    Write-Host "No Copilot interactions returned for the window." -ForegroundColor Yellow
    return
}

Write-Host "Aggregating per-user interaction counts..." -ForegroundColor Cyan

# auditData payload structure varies; pull the most useful fields defensively.
$rows = foreach ($rec in $records) {
    $data = $rec.auditData
    if ($data -is [string]) {
        try { $data = $data | ConvertFrom-Json } catch { $data = $null }
    }

    $upn      = $data.UserId
    $userKey  = $data.UserKey         # AAD object id when present
    $appHost  = $data.CopilotEventData.AppHost
    if (-not $appHost) { $appHost = $data.AppHost }
    $created  = $rec.createdDateTime
    if (-not $created) { $created = $data.CreationTime }

    [PSCustomObject]@{
        Upn       = $upn
        UserId    = $userKey
        AppHost   = $appHost
        TimeStamp = [datetime]$created
    }
}

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

        [PSCustomObject]@{
            UserPrincipalName  = $g.Name
            UserId             = $uid
            InteractionCount   = $g.Count
            FirstUsedUtc       = $first.ToString('u')
            LastUsedUtc        = $last.ToString('u')
            AppHostsUsed       = $hosts
            HasCopilotLicense  = $hasLic
        }
    }

if (-not $IncludeLicensed) {
    $grouped = $grouped | Where-Object { -not $_.HasCopilotLicense }
}

$grouped = $grouped | Sort-Object InteractionCount -Descending

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
