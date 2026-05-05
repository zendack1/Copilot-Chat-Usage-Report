# Copilot Chat Usage Report

A PowerShell script that produces a complete report of every user in your Microsoft 365 tenant who has used the **free / unlicensed Copilot Chat** experience over the last 30, 60, or 90 days — bypassing the 1,000-user cap of the M365 Admin Center.

Each run produces:

- a **CSV** with one row per user (UPN, interaction count, active-day count, engagement tier, first/last use, surfaces used, department, job title, office)
- a **self-contained interactive HTML report** with KPI tiles (users / total / avg / median), an engagement-tier histogram, an app-host distribution chart, a Top-20 bar chart, and a sortable / searchable / printable table — opens by double-clicking, no server, no internet, no dependencies

*See [`examples/mock_report.html`](examples/mock_report.html) for a sample.*

---

## Why this exists

The Microsoft 365 Admin Center's Copilot usage report caps the user list at **1,000 rows**. For tenants larger than that, you can't enumerate every unlicensed Copilot Chat user from the UI. This script uses the Microsoft Graph **`auditLogQuery`** API — which has no such cap and pages through the full result set — to give you everyone.

It then cross-references each active user against assigned licenses and (by default) drops anyone holding a Microsoft 365 Copilot SKU, leaving only the free-Copilot-Chat population. Each user is enriched with department, job title, and office location so admins can target conversations by org segment, and classified into an engagement tier (Power / Regular / Casual / One-time) based on event count and active-day count — making the report actionable for license-expansion decisions, not just a flat user list.

## Prerequisites

1. **PowerShell 7+**
2. Modules:
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   Install-Module Microsoft.Graph.Beta -Scope CurrentUser
   ```
3. **Admin role** — sign in as one of:
   - Global Administrator, Global Reader
   - Security Administrator / Security Reader
   - Audit Log Reader
4. Graph scopes (delegated for interactive runs, application for unattended):
   - `AuditLogsQuery.Read.All`
   - `User.Read.All`
   - `Organization.Read.All`
5. **Unified audit log** must be enabled in the tenant (default for new tenants since 2023).

## Quick start

```powershell
git clone https://github.com/<your-handle>/copilot-chat-usage-report.git
cd copilot-chat-usage-report
.\Get-CopilotChatUsage.ps1
```

Default behavior: last 30 days, free Copilot Chat users only, both CSV and HTML written to the current directory with a UTC-stamped filename.

## Usage

| Command | Result |
|---|---|
| `.\Get-CopilotChatUsage.ps1` | Last 30 days, unlicensed users only |
| `.\Get-CopilotChatUsage.ps1 -Days 60` | Last 60 days |
| `.\Get-CopilotChatUsage.ps1 -Days 90` | Last 90 days |
| `.\Get-CopilotChatUsage.ps1 -IncludeLicensed` | Include paid Copilot users in the output |
| `.\Get-CopilotChatUsage.ps1 -NoHtml` | CSV only, skip the HTML report |
| `.\Get-CopilotChatUsage.ps1 -OutputPath .\report.csv -HtmlPath .\report.html` | Custom file locations |
| `.\Get-CopilotChatUsage.ps1 -SkipDirectoryEnrichment` | Skip per-user `Get-MgUser` lookups — faster on very large tenants but no department / job-title columns |
| `.\Get-CopilotChatUsage.ps1 -CopilotSkuPattern '*Copilot*'` | Custom SKU match pattern (default `*Copilot*` covers all current Copilot SKUs) |
| `.\Get-CopilotChatUsage.ps1 -QueryTimeoutMinutes 90` | Increase audit-query wait timeout (default 60 min) |
| `.\Get-CopilotChatUsage.ps1 -TestConnection` | Verify auth + scopes + SKU enumeration without submitting an audit query |

`-Days` is validated to **30**, **60**, or **90**. Other values are rejected at parameter binding.

The query runs asynchronously on the service side; expect 2–10 minutes for a 30-day window before records start streaming back. The script polls every 30 seconds and prints status, retrying transient Graph errors (408 / 429 / 5xx) with exponential backoff and `Retry-After` honoring.

## Output columns

| Column | Meaning |
|---|---|
| `UserPrincipalName` | The user's UPN |
| `UserId` | Entra (AAD) object id when the audit record carried it |
| `InteractionCount` | Number of Copilot Chat events in the window |
| `ActiveDays` | Distinct UTC dates the user interacted with Copilot Chat |
| `EngagementTier` | `Power` (100+ events or 15+ active days), `Regular` (20+ events or 5+ days), `Casual` (2+ events), `OneTime` (1 event) |
| `FirstUsedUtc` / `LastUsedUtc` | First and last interaction timestamps |
| `AppHostsUsed` | Surfaces used (e.g. `Bing`, `M365App`, `Teams`). See [Microsoft's CopilotInteraction schema](https://learn.microsoft.com/en-us/office/office-365-management-api/copilot-schema) for the full value set. |
| `HasCopilotLicense` | Present in the CSV for verification; hidden in the HTML when running default scope |
| `DisplayName` / `Department` / `JobTitle` / `OfficeLocation` | Directory enrichment (omitted when `-SkipDirectoryEnrichment` is set) |

## Unattended / scheduled runs

Pass app-only credentials to skip interactive sign-in:

```powershell
.\Get-CopilotChatUsage.ps1 `
    -ClientId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' `
    -TenantId 'contoso.onmicrosoft.com' `
    -CertificateThumbprint 'AB12CD34EF56...' `
    -Days 30
```

Grant the same three Graph scopes as **Application** permissions on the app registration and admin-consent them. The certificate must be installed in `CurrentUser\My` (or `LocalMachine\My`) on the host running the script.

Use `-TestConnection` first to verify the registration before scheduling:

```powershell
.\Get-CopilotChatUsage.ps1 -ClientId ... -TenantId ... -CertificateThumbprint ... -TestConnection
```

## HTML report features

- **KPI tiles**: total users, total interactions, average per user, median per user
- **Engagement-tier histogram**: Power / Regular / Casual / One-time bucket counts
- **App-host distribution**: which surfaces are driving usage (Bing, M365App, Teams, Edge, Word, etc.)
- **Top-20 bar chart**: highest-volume users
- **Sortable / searchable table** with all enriched columns
- **Filtered CSV download**: export the current filtered view as a separate CSV without re-running the script
- **Print stylesheet**: prints cleanly to PDF / paper for distribution
- **Safe to share**: the report carries no tenant id, customer name, or directory metadata beyond the filtered user list — paste a screenshot in a deck without redaction
- **Hostile-input safe**: embedded JSON is escaped (`</` → `<\/`) so a UPN containing `</script>` cannot break out of the data block

## Tests

The repository includes a Python test harness (`tests/test_script.py`) that performs structural validation, simulates the HTML rendering pipeline with a synthetic 1,000-user dataset, and asserts:

- PowerShell brace / paren / here-string balance
- Required helper functions defined and called
- All new parameters wired through to `Main`
- AppOnly parameter set declares mandatory `ClientId` / `TenantId` / `CertificateThumbprint`
- `-Days` parameter validates to 30/60/90
- Engagement-tier classification rules match the script's logic
- License-map keying uses lower-cased UPN
- SKU detection uses a wildcard pattern (not a hardcoded list)
- All licensed users are filtered out under the default scope (zero leaks)
- HTML template contains all four KPIs, both charts, the print stylesheet, the latency disclaimer, and the filtered-CSV download
- Embedded JSON parses; `</script>` injection in user names is escaped safely
- Top-20 chart slice has at most 20 rows and the highest count first

Run the tests:

```bash
python tests/test_script.py
```

Note: the harness is Python because PowerShell isn't required to validate the output structure. Live Graph calls are **not** exercised — those need a real admin sign-in.

## Tuning notes

- The default SKU pattern (`*Copilot*`) covers `Microsoft_365_Copilot`, `Copilot_Pro`, `Microsoft_Copilot_for_Sales`, `Microsoft_Copilot_for_Service`, and any future SKU whose name contains "Copilot". If your tenant uses a paid Copilot SKU with a different naming convention, override with `-CopilotSkuPattern '<pattern>'`. Run `Get-MgSubscribedSku | Select SkuPartNumber, SkuId` to list SKUs in your tenant.
- Audit records can take up to ~30 minutes to surface after the actual interaction, so a run "right now" may slightly under-count today's activity. The HTML report calls this out in the footer.
- For very large tenants (10K+ active users in the window), `-SkipDirectoryEnrichment` can shave several minutes off the run at the cost of empty department / job-title / office columns.

## License

[MIT](LICENSE)

## Contributing

Issues and PRs welcome. Please run `python tests/test_script.py` before submitting changes — it should report all checks passing.
