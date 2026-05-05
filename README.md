# Copilot Chat Usage Report

A PowerShell script that produces a complete report of every user in your Microsoft 365 tenant who has used the **free / unlicensed Copilot Chat** experience over the last 30, 60, or 90 days — bypassing the 1,000-user cap of the M365 Admin Center.

Each run produces:

- a **CSV** with one row per user (UPN, interaction count, first/last use, surfaces used)
- a **self-contained interactive HTML report** with KPI tiles, a Top-20 bar chart, and a sortable / searchable table — opens by double-clicking, no server, no internet, no dependencies

![Copilot Chat usage report screenshot](mock_report.html)
*See [`mock_report.html`](mock_report.html) for a sample.*

---

## Why this exists

The Microsoft 365 Admin Center's Copilot usage report caps the user list at **1,000 rows**. For tenants larger than that, you can't enumerate every unlicensed Copilot Chat user from the UI. This script uses the Microsoft Graph **`auditLogQuery`** API — which has no such cap and pages through the full result set — to give you everyone.

It then cross-references each active user against assigned licenses and (by default) drops anyone holding a Microsoft 365 Copilot SKU, leaving only the free-Copilot-Chat population.

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
4. Graph delegated scopes (consented on first run):
   - `AuditLogsQuery.Read.All`
   - `User.Read.All`
   - `Organization.Read.All`
5. **Unified audit log** must be enabled in the tenant (default for new tenants since 2023).

## Quick start

```powershell
git clone https://github.com/zendack1/copilot-chat-usage-report.git
cd copilot-chat-usage-report
.\Get-CopilotChatUsage.ps1
```

Default behavior: last 30 days, free Copilot Chat users only, both CSV and HTML written to the current directory.

## Usage

| Command | Result |
|---|---|
| `.\Get-CopilotChatUsage.ps1` | Last 30 days, unlicensed users only |
| `.\Get-CopilotChatUsage.ps1 -Days 60` | Last 60 days |
| `.\Get-CopilotChatUsage.ps1 -Days 90` | Last 90 days |
| `.\Get-CopilotChatUsage.ps1 -IncludeLicensed` | Include paid Copilot users in the output |
| `.\Get-CopilotChatUsage.ps1 -NoHtml` | CSV only, skip the HTML report |
| `.\Get-CopilotChatUsage.ps1 -OutputPath .\report.csv -HtmlPath .\report.html` | Custom file locations |

`-Days` is validated to **30**, **60**, or **90**. Other values are rejected at parameter binding.

The query runs asynchronously on the service side; expect 2–10 minutes for a 30-day window before records start streaming back. The script polls every 30 seconds and prints status.

## Output columns

| Column | Meaning |
|---|---|
| `UserPrincipalName` | The user's UPN |
| `UserId` | Entra (AAD) object id when the audit record carried it |
| `InteractionCount` | Number of Copilot Chat events in the window |
| `FirstUsedUtc` / `LastUsedUtc` | First and last interaction timestamps |
| `AppHostsUsed` | Surfaces used (e.g. `Bing`, `M365App`, `Teams`). See [Microsoft's CopilotInteraction schema](https://learn.microsoft.com/en-us/office/office-365-management-api/copilot-schema) for the full value set. |
| `HasCopilotLicense` | Present in the CSV for verification; hidden in the HTML when running default scope |

## Unattended / scheduled runs

Replace the interactive `Connect-MgGraph` call with an app-only connection:

```powershell
Connect-MgGraph -ClientId <appId> -TenantId <tenantId> -CertificateThumbprint <thumb>
```

Grant the same three permissions as **Application** permissions on the app registration and admin-consent them.

## Tests

The repository includes a Python test harness (`tests/test_script.py`) that performs structural validation, simulates the HTML rendering pipeline with a synthetic 1,000-user dataset, and asserts:

- PowerShell brace / paren / here-string balance
- Required helper functions defined and called
- `-Days` parameter validates to 30/60/90
- License-map keying uses lower-cased UPN
- All licensed users are filtered out under the default scope (zero leaks)
- Embedded JSON parses; `</script>` injection in user names is escaped safely
- Top-20 chart slice has at most 20 rows and the highest count first

Run the tests:

```bash
python tests/test_script.py
```

Note: the harness is Python because PowerShell isn't required to validate the output structure. Live Graph calls are **not** exercised — those need a real admin sign-in.

## Tuning notes

- If your tenant uses a Copilot SKU whose `SkuPartNumber` doesn't contain `Copilot`, add it to `$CopilotSkuPartNumbers` at the top of the script. Run `Get-MgSubscribedSku | Select SkuPartNumber, SkuId` to list SKUs in your tenant.
- Audit records can take up to ~30 minutes to surface after the actual interaction, so a run "right now" may slightly under-count today's activity.

## License

[MIT](LICENSE)

## Contributing

Issues and PRs welcome. Please run `python tests/test_script.py` before submitting changes — it should report all checks passing.
