# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2026-05-05

### Added
- **App-only / unattended auth path**: new `-ClientId`, `-TenantId`, and
  `-CertificateThumbprint` parameters (parameter set `AppOnly`, all mandatory
  together) for scheduled-task runs without an interactive sign-in.
- **`-TestConnection` switch**: verifies sign-in, granted scopes, and SKU
  enumeration without submitting an audit query — useful for validating an
  app registration before scheduling.
- **`-CopilotSkuPattern` parameter** (default `*Copilot*`): replaces the
  prior hardcoded SKU list. Future Copilot SKUs are picked up automatically
  as long as their `SkuPartNumber` contains "Copilot".
- **`-QueryTimeoutMinutes` parameter** (default 60): configurable wait for
  the Graph audit query to finish, so long-running queries on big tenants
  don't time out at a fixed boundary.
- **`-RedactTenant` switch**: replaces the tenant id in the HTML subtitle
  with `[redacted]` — safe to share screenshots externally.
- **`-SkipDirectoryEnrichment` switch**: skip per-user `Get-MgUser` lookups.
  Faster on very large tenants at the cost of empty department / job-title /
  office columns.
- **Engagement tier classification** (`Power` / `Regular` / `Casual` /
  `OneTime`) — combined event-count + active-day rule, surfacing the high-
  value population for license-expansion decisions instead of a flat list.
- **Active-day metric**: distinct UTC dates per user, in addition to the raw
  event count.
- **Directory enrichment** via batched `Get-MgUser` (15 UPNs per request):
  adds `DisplayName`, `Department`, `JobTitle`, and `OfficeLocation` columns
  to the CSV.
- **Median interactions / user** alongside the existing average — robust
  central-tendency signal when a few power users skew the mean.
- **App-host distribution chart** in the HTML: which surfaces (Bing, Teams,
  M365App, Edge, etc.) are driving Copilot Chat usage.
- **Engagement-tier histogram** in the HTML.
- **Filtered CSV download** in the HTML report — the in-browser table filter
  feeds a client-side CSV export, so admins can slice by department / tier /
  app host without re-running the script.
- **Print stylesheet** in the HTML report — prints cleanly to PDF for
  distribution; non-printing controls are hidden.
- **Latency disclaimer** banner in the HTML footer ("audit records can take
  up to ~30 minutes to surface").
- **Title date range** in `<title>` and HTML subtitle.
- **Graph retry wrapper** (`Invoke-GraphWithRetry`): exponential backoff on
  408 / 429 / 5xx with `Retry-After` header honoring. All audit-query and
  paging requests now go through it.

### Changed
- Default `OutputPath` filename now stamps the **UTC** time (not local), and
  the suffix `_utc.csv` makes the timezone explicit.
- License-map keying clarified to lower-cased UPN throughout (audit payloads
  may differ in case from directory entries).
- Page-status output for audit-record paging moved to `Write-Verbose` so the
  default console output stays clean.

### Fixed
- Guard against null / unparseable `createdDateTime` values in audit records:
  rows with no usable timestamp are now skipped instead of throwing during
  the `[datetime]` cast.

### Notes
- Existing `-Days`, `-OutputPath`, `-HtmlPath`, `-IncludeLicensed`, and
  `-NoHtml` parameters are unchanged. Existing CSVs from v1.0.0 will load in
  the v1.1.0 HTML viewer, but the new columns will be empty.

## [1.0.0] - 2026-05-05

### Added
- Initial release of `Get-CopilotChatUsage.ps1`.
- Pulls every Copilot Chat user from the Microsoft Graph beta `auditLogQuery`
  API — bypasses the M365 Admin Center 1,000-user cap and the
  `Search-UnifiedAuditLog` 5,000-row cap.
- Per-user aggregation: UPN, interaction count, first/last use timestamps,
  app hosts used (Bing, M365App, Teams, etc.).
- License cross-reference: drops users holding a Microsoft 365 Copilot SKU
  by default so the output matches the Admin Center's "free Copilot Chat"
  definition. License map keyed by lower-cased UPN to match the audit
  payload's `UserId` field reliably.
- `-Days` parameter validated to **30**, **60**, or **90**.
- `-IncludeLicensed` switch to include paid Copilot users in the output.
- `-NoHtml` switch to emit CSV only.
- `-OutputPath` and `-HtmlPath` for custom file locations.
- Self-contained interactive HTML report — KPI tiles, top-20 bar chart,
  sortable / searchable table, client-side CSV export. No server, no
  internet, no external dependencies.
- Embedded JSON payload is escaped (`</` → `<\/`) so hostile UPNs
  containing `</script>` cannot break out of the data block.
- Python test harness (`tests/test_script.py`) — structural lint, helper
  function presence, parameter wiring, HTML template extraction, and
  end-to-end render simulation against a synthetic 1,000-user dataset
  with a `</script>` injection probe.
- Sample report (`examples/mock_report.html`) generated from the test
  harness for preview without running the live script.

### Notes
- The audit query runs asynchronously on the service side; expect 2–10
  minutes for a 30-day window before records start streaming back.
- Records can take up to ~30 minutes to surface in the audit log, so a
  run "right now" may slightly under-count today's activity.
