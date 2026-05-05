# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
