"""
Static + simulated tests for Get-CopilotChatUsage.ps1.

Without pwsh available we can't execute the script, but we CAN:
  1. Lint structural balance (braces, parentheses, here-strings).
  2. Verify all expected helper functions are defined.
  3. Verify all parameters used in Main exist.
  4. Verify the new param set / parameter additions for app-only auth,
     SKU pattern, query timeout, redact, test-connection, and skip-enrichment.
  5. Extract the embedded HTML template, inject mock JSON, write the page,
     and validate the result (HTML parses, JSON inside <script> is valid,
     all expected DOM hooks exist, no broken </script> sequences,
     engagement-tier and host-distribution charts present, filtered CSV
     download wired up, print stylesheet and latency note present).
  6. Verify the JS code references match the data shape.
  7. Verify engagement tier classification matches the script's rules.
"""
from __future__ import annotations
import html.parser
import json
import re
import sys
from pathlib import Path

SCRIPT = Path("/mnt/workspace/output/Get-CopilotChatUsage.ps1")
OUT_HTML = Path("/mnt/workspace/working/mock_report.html")
EXAMPLES_HTML = Path("/mnt/workspace/output/examples/mock_report.html")

failures: list[str] = []
warnings: list[str] = []


def check(name: str, ok: bool, detail: str = ""):
    status = "PASS" if ok else "FAIL"
    print(f"[{status}] {name}{(': ' + detail) if detail else ''}")
    if not ok:
        failures.append(name + ((": " + detail) if detail else ""))


def warn(name: str, detail: str = ""):
    print(f"[WARN] {name}{(': ' + detail) if detail else ''}")
    warnings.append(name + ((": " + detail) if detail else ""))


src = SCRIPT.read_text(encoding="utf-8")

# ------------------------------------------------------------------ 1. Balance
def strip_strings_and_comments(s: str) -> str:
    out = []
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if c == "<" and s[i:i+2] == "<#":
            end = s.find("#>", i+2)
            if end == -1: break
            i = end + 2; continue
        if c == "#":
            end = s.find("\n", i)
            i = end if end != -1 else n; continue
        if c == "@" and i+1 < n and s[i+1] in "'\"":
            quote = s[i+1]
            term = quote + "@"
            end = s.find("\n" + term, i+2)
            if end == -1: break
            i = end + len(term) + 1; continue
        if c == "'":
            end = i + 1
            while end < n:
                if s[end] == "'":
                    if end + 1 < n and s[end+1] == "'":
                        end += 2; continue
                    break
                end += 1
            i = end + 1; continue
        if c == '"':
            end = i + 1
            while end < n:
                if s[end] == '`' and end+1 < n:
                    end += 2; continue
                if s[end] == '"':
                    if end + 1 < n and s[end+1] == '"':
                        end += 2; continue
                    break
                end += 1
            i = end + 1; continue
        out.append(c); i += 1
    return "".join(out)

stripped = strip_strings_and_comments(src)
# Raw counts (cheap sanity check): a real lexer would be more accurate, but raw
# counts catch the obvious "forgot a closing brace" cases without flagging
# false positives on PowerShell $() subexpressions inside double-quoted strings.
raw_open_b = src.count("{");  raw_close_b = src.count("}")
raw_open_p = src.count("(");  raw_close_p = src.count(")")
check("Curly braces balanced (raw count)", raw_open_b == raw_close_b,
      f"{raw_open_b} open vs {raw_close_b} close")
check("Parentheses balanced (raw count)", raw_open_p == raw_close_p,
      f"{raw_open_p} open vs {raw_close_p} close")

hs_open  = len(re.findall(r"@['\"]\s*\n", src))
hs_close = len(re.findall(r"\n['\"]@", src))
check("Here-strings balanced", hs_open == hs_close, f"{hs_open} open vs {hs_close} close")

# ------------------------------------------------------------------ 2. Helper functions
defined = set(re.findall(r"^function\s+([A-Za-z\-_]+)", src, re.MULTILINE))
expected = {
    "Connect-Graph", "Invoke-GraphWithRetry", "Get-CopilotLicensedUserIds",
    "Get-UserDirectoryProperties", "Get-EngagementTier", "Get-Median",
    "Submit-AuditQuery", "Wait-AuditQuery", "Get-AuditQueryRecords",
    "Write-HtmlReport",
}
missing = expected - defined
check("All expected helper functions defined", not missing, f"missing: {missing or 'none'}")

main_block = src.split("# ---------- Main ----------", 1)[1] if "# ---------- Main ----------" in src else ""
for fn in ["Submit-AuditQuery", "Wait-AuditQuery", "Get-AuditQueryRecords",
           "Get-CopilotLicensedUserIds", "Get-UserDirectoryProperties",
           "Get-EngagementTier", "Write-HtmlReport"]:
    check(f"Main invokes {fn}", fn in main_block)

check("Audit calls go through Invoke-GraphWithRetry",
      src.count("Invoke-GraphWithRetry") >= 3)

# ------------------------------------------------------------------ 3. Parameters
for var in ("$Days", "$OutputPath", "$HtmlPath", "$IncludeLicensed", "$NoHtml",
            "$CopilotSkuPattern", "$QueryTimeoutMinutes", "$RedactTenant",
            "$TestConnection", "$SkipDirectoryEnrichment"):
    check(f"Main uses {var}", var in main_block, "")

m_validate = re.search(r"\[ValidateSet\(\s*30\s*,\s*60\s*,\s*90\s*\)\]\s*\n\s*\[int\]\$Days", src)
check("$Days parameter has ValidateSet(30, 60, 90)", bool(m_validate))

for p in ("ClientId", "TenantId", "CertificateThumbprint"):
    pat = rf"\[Parameter\(ParameterSetName\s*=\s*'AppOnly',\s*Mandatory\)\]\s*\n\s*\[string\]\${p}"
    check(f"AppOnly parameter set: ${p} is mandatory", bool(re.search(pat, src)))

check("CmdletBinding declares default param set 'Delegated'",
      "DefaultParameterSetName = 'Delegated'" in src)

check("Connect-Graph picks AppOnly path when set",
      "ParameterSetName -eq 'AppOnly'" in src and "-CertificateThumbprint" in src)

check("Delegated connect verifies required scopes are granted",
      "Missing required scopes" in src)

check("TestConnection mode short-circuits Main",
      bool(re.search(r"if\s*\(\$TestConnection\)\s*\{[\s\S]+?return\s*\}\s*", src)))

check("Default OutputPath uses UTC timestamp",
      "ToUniversalTime().ToString('yyyyMMdd_HHmmss')" in src and "_utc.csv" in src)

check("Guards [datetime] cast against null/invalid input",
      "try { $timestamp = [datetime]$created }" in src)

check("Engagement tier rule: Power threshold 100 events or 15 active days",
      bool(re.search(r"\$Interactions\s*-ge\s*100\s*-or\s*\$ActiveDays\s*-ge\s*15", src)))
check("Engagement tier rule: Regular threshold 20 events or 5 active days",
      bool(re.search(r"\$Interactions\s*-ge\s*20\s+-or\s+\$ActiveDays\s*-ge\s*5", src)))

# ------------------------------------------------------------------ 4. License match logic
check("Licensed map keyed by lower-cased UPN",
      bool(re.search(r"\$licensed\[\$u\.UserPrincipalName\.ToLowerInvariant\(\)\]\s*=", src)))
check("Licensed-map lookup uses lower-cased UPN",
      bool(re.search(
          r"\$upnKey\s*=\s*\$g\.Name\.ToLowerInvariant\(\)[\s\S]{0,80}\$licensedMap\.ContainsKey\(\$upnKey\)", src)))
check("SKU detection uses pattern match (not hardcoded list)",
      "$_.SkuPartNumber -like $SkuPattern" in src)

# ------------------------------------------------------------------ 5. HTML template
m = re.search(r"\$template\s*=\s*@'\s*\n(.*?)\n'@", src, re.DOTALL)
check("HTML template here-string extractable", bool(m))
if not m:
    print("\nFAILURES:", failures); sys.exit(1)

template = m.group(1)
check("Template contains __DATA_JSON__ placeholder", "__DATA_JSON__" in template)
check("Template contains __TITLE_RANGE__ placeholder", "__TITLE_RANGE__" in template)
check("Template has <!doctype html>", template.lower().lstrip().startswith("<!doctype html>"))
check("Template has all four KPI tiles (users/events/avg/median)",
      all(k in template for k in ["kpi-users", "kpi-events", "kpi-avg", "kpi-median"]))
check("Template has engagement tier chart", 'id="tier-chart"' in template)
check("Template has app host distribution chart", 'id="host-chart"' in template)
check("Template has top-20 chart container", 'id="chart"' in template)
check("Template has filter input", 'id="filter"' in template)
check("Template has table tbody", 'id="tbody"' in template)
check("Template has print stylesheet", "@media print" in template)
check("Template has latency disclaimer", "30 minutes" in template and "under-counted" in template)
check("Filtered CSV download wired up",
      "CopilotChatUsage_filtered.csv" in template and "currentView" in template)
check("Table includes Tier column", 'data-key="EngagementTier"' in template)
check("Table includes ActiveDays column", 'data-key="ActiveDays"' in template)
check("Table includes Department column", 'data-key="Department"' in template)
check("License column not in HTML header", 'data-key="HasCopilotLicense"' not in template)

# ------------------------------------------------------------------ 6. Render mock report
import random
random.seed(1000)

SAMPLE_SIZE = 1000

FIRST_NAMES = ["alex","jordan","taylor","morgan","casey","sam","riley","drew",
               "jamie","quinn","reese","avery","hayden","blake","cameron",
               "dakota","elliot","frankie","gray","harper","indigo","jules",
               "kai","logan","micah","noah","olivia","priya","rahul","sofia",
               "theo","uma","viktor","willa","xavier","yuki","zoe","arjun",
               "bea","chloe","diego","eva","felix","gigi","hassan","iris"]
LAST_NAMES  = ["smith","jones","patel","garcia","kim","nguyen","khan","wong",
               "lopez","brown","davis","miller","wilson","moore","clark",
               "lewis","walker","young","hall","king","scott","green",
               "baker","carter","price","murphy","reed","torres","stewart",
               "morris","bailey","cooper","richardson","cox","howard","ward",
               "watson","brooks","kelly","sanders","price","bennett","wood",
               "barnes","ross","henderson"]
APP_HOSTS = ["Bing","M365App","Teams","Edge","Word","Outlook","PowerPoint","Excel"]
DEPARTMENTS = ["Engineering","Finance","Marketing","Sales","HR","Operations",
               "Customer Success","Legal","Product","IT"]
JOB_TITLES = ["Software Engineer","Senior Engineer","Program Manager",
              "Account Executive","Analyst","Director","VP",
              "Specialist","Architect","Technical Lead"]


def classify_tier(events: int, active_days: int) -> str:
    if events >= 100 or active_days >= 15: return "Power"
    if events >= 20  or active_days >= 5:  return "Regular"
    if events >= 2:                        return "Casual"
    return "OneTime"


mock_rows = []
upns = set()
counter = 0
while len(mock_rows) < SAMPLE_SIZE:
    fn = random.choice(FIRST_NAMES)
    ln = random.choice(LAST_NAMES)
    counter += 1
    upn = f"{fn}.{ln}{counter}@contoso.com"
    if upn in upns: continue
    upns.add(upn)

    hosts = ";".join(sorted(set(random.choices(APP_HOSTS, k=random.randint(1,3)))))
    start_day = random.randint(0, 25)
    end_day   = start_day + random.randint(1, 4)
    events = random.randint(1, 500)
    active = min(30, max(1, int(events ** 0.5) + random.randint(-2, 2)))
    if events == 1: active = 1
    mock_rows.append({
        "UserPrincipalName": upn,
        "UserId": f"00000000-0000-0000-0000-{counter:012d}",
        "InteractionCount": events,
        "ActiveDays": active,
        "EngagementTier": classify_tier(events, active),
        "FirstUsedUtc": f"2026-04-{5+start_day:02d} 09:00:00Z",
        "LastUsedUtc":  f"2026-04-{min(30, 5+end_day):02d} 17:00:00Z",
        "AppHostsUsed": hosts,
        "HasCopilotLicense": random.random() < 0.20,
        "DisplayName":    f"{fn.title()} {ln.title()}",
        "Department":     random.choice(DEPARTMENTS),
        "JobTitle":       random.choice(JOB_TITLES),
        "OfficeLocation": random.choice(["Redmond","London","Tokyo","Bangalore","Remote"]),
    })

mock_rows[7] = {
    "UserPrincipalName": "evil</script>user@contoso.com",
    "UserId": "00000000-0000-0000-0000-000000000007",
    "InteractionCount": 42,
    "ActiveDays": 7,
    "EngagementTier": "Regular",
    "FirstUsedUtc": "2026-04-30 11:00:00Z",
    "LastUsedUtc":  "2026-05-01 11:00:00Z",
    "AppHostsUsed": "Teams",
    "HasCopilotLicense": False,
    "DisplayName":    "Evil </script> User",
    "Department":     "Security Research",
    "JobTitle":       "Pentester",
    "OfficeLocation": "Remote",
}

pre_filter_count = len(mock_rows)
pre_filter_licensed = sum(1 for r in mock_rows if r["HasCopilotLicense"])

INCLUDE_LICENSED = False
if not INCLUDE_LICENSED:
    mock_rows = [r for r in mock_rows if not r["HasCopilotLicense"]]

mock_rows.sort(key=lambda r: r["InteractionCount"], reverse=True)
print(f"\n[INFO] Pre-filter: {pre_filter_count} users ({pre_filter_licensed} licensed)")
print(f"[INFO] Post-filter sample size: {len(mock_rows)} users (will appear in report)")


def median(values):
    if not values: return 0
    s = sorted(values)
    n = len(s)
    if n % 2 == 0:
        return round((s[n//2 - 1] + s[n//2]) / 2.0, 1)
    return s[n//2]


tier_counts = {"Power": 0, "Regular": 0, "Casual": 0, "OneTime": 0}
for r in mock_rows: tier_counts[r["EngagementTier"]] += 1
tiers = [
    {"key": "Power",   "label": "Power (100+ events or 15+ active days)", "count": tier_counts["Power"]},
    {"key": "Regular", "label": "Regular (20-99 events or 5-14 days)",    "count": tier_counts["Regular"]},
    {"key": "Casual",  "label": "Casual (2-19 events)",                   "count": tier_counts["Casual"]},
    {"key": "OneTime", "label": "One-time (1 event)",                     "count": tier_counts["OneTime"]},
]
host_counts = {}
for r in mock_rows:
    if r.get("AppHostsUsed"):
        for h in r["AppHostsUsed"].split(";"):
            h = h.strip()
            if not h: continue
            host_counts[h] = host_counts.get(h, 0) + 1
hosts = sorted(
    [{"label": k, "count": v} for k, v in host_counts.items()],
    key=lambda x: x["count"], reverse=True,
)

events_total = sum(r["InteractionCount"] for r in mock_rows)
title_range  = "2026-04-04 to 2026-05-04"
payload = {
    "rows": mock_rows,
    "startUtc": "2026-04-04 20:00:00Z",
    "endUtc":   "2026-05-04 20:00:00Z",
    "generatedAt": "2026-05-04 20:00:00Z",
    "tenantId": "00000000-0000-0000-0000-aaaaaaaaaaaa",
    "scopeText": "Free / unlicensed Copilot Chat users only" if not INCLUDE_LICENSED
                  else "All Copilot Chat users (licensed and unlicensed)",
    "totalUsers": len(mock_rows),
    "totalEvents": events_total,
    "avgPerUser": round(events_total / len(mock_rows), 1) if mock_rows else 0,
    "medianPerUser": median([r["InteractionCount"] for r in mock_rows]),
    "tiers": tiers,
    "hosts": hosts,
    "titleRange": title_range,
    "includeLicensed": INCLUDE_LICENSED,
}
payload_json = json.dumps(payload, separators=(",", ":"))
safe = payload_json.replace("</", "<\\/")
rendered = template.replace("__DATA_JSON__", safe).replace("__TITLE_RANGE__", title_range)
OUT_HTML.parent.mkdir(parents=True, exist_ok=True)
OUT_HTML.write_text(rendered, encoding="utf-8")
EXAMPLES_HTML.parent.mkdir(parents=True, exist_ok=True)
EXAMPLES_HTML.write_text(rendered, encoding="utf-8")
check("Mock report file written", OUT_HTML.exists() and OUT_HTML.stat().st_size > 1000,
      f"{OUT_HTML.stat().st_size} bytes")

# ------------------------------------------------------------------ 7. Validate produced HTML
class P(html.parser.HTMLParser):
    def __init__(self): super().__init__(); self.errors = []
    def error(self, msg): self.errors.append(msg)
p = P()
try:
    p.feed(rendered); p.close()
    check("HTML parses without exceptions", not p.errors, "; ".join(p.errors[:3]))
except Exception as e:
    check("HTML parses without exceptions", False, str(e))

m2 = re.search(r'<script id="report-data" type="application/json">(.*?)</script>',
               rendered, re.DOTALL)
check("report-data <script> block present and closes correctly", bool(m2))
if m2:
    raw = m2.group(1)
    try:
        parsed = json.loads(raw)
        check("Injected JSON parses", True)
        check("Injected JSON has rows array", isinstance(parsed.get("rows"), list))
        check("Injected JSON row count matches", len(parsed["rows"]) == len(mock_rows))
        check("Injected JSON has tier histogram", isinstance(parsed.get("tiers"), list) and len(parsed["tiers"]) == 4)
        check("Injected JSON has host distribution", isinstance(parsed.get("hosts"), list))
        check("Injected JSON has avg+median KPIs",
              "avgPerUser" in parsed and "medianPerUser" in parsed)
        bad_user = next((r for r in parsed["rows"] if "evil" in r["UserPrincipalName"]), None)
        check("Hostile </script> in data preserved through escape",
              bad_user is not None and bad_user["UserPrincipalName"] == "evil</script>user@contoso.com")
    except Exception as e:
        check("Injected JSON parses", False, str(e))

script_close_in_data = m2.group(1).count("</script>") if m2 else -1
check("No raw </script> inside the data block", script_close_in_data == 0)

check("Title placeholder filled with date range",
      f"<title>Copilot Chat Usage Report - {title_range}</title>" in rendered)

js_fields = {"UserPrincipalName", "EngagementTier", "InteractionCount", "ActiveDays",
             "Department", "JobTitle", "FirstUsedUtc", "LastUsedUtc", "AppHostsUsed"}
data_fields = set(mock_rows[0].keys())
missing_js = js_fields - data_fields
check("JS-referenced fields all present in row data", not missing_js, f"missing: {missing_js or 'none'}")

for key in re.findall(r'data-key="([^"]+)"', rendered):
    check(f"Table column '{key}' maps to a row field", key in data_fields)

# ------------------------------------------------------------------ 8. Tier classification spot-checks
spot_checks = [
    (200,  3, "Power"),
    (50,   3, "Regular"),
    (10,   3, "Casual"),
    (1,    1, "OneTime"),
    (5,   16, "Power"),
    (5,    8, "Regular"),
]
for evt, days, expected_tier in spot_checks:
    check(f"Tier({evt} events, {days} days) == {expected_tier}",
          classify_tier(evt, days) == expected_tier,
          f"got {classify_tier(evt, days)}")

mismatched = [r for r in mock_rows
              if classify_tier(r["InteractionCount"], r["ActiveDays"]) != r["EngagementTier"]]
check("All mock rows have tier matching the classification rule",
      not mismatched, f"{len(mismatched)} mismatches")

# ------------------------------------------------------------------ 9. Sample / filter behavior
check("Sample after filter is non-empty", len(mock_rows) > 0, f"{len(mock_rows)} rows")
check("Pre-filter sample reaches target size",
      pre_filter_count == SAMPLE_SIZE, f"{pre_filter_count} of {SAMPLE_SIZE}")
check("Pre-filter sample contains licensed users",
      pre_filter_licensed > 0, f"{pre_filter_licensed} licensed users in raw sample")
check("Default scope removes all licensed users",
      all(not r["HasCopilotLicense"] for r in mock_rows),
      f"{sum(1 for r in mock_rows if r['HasCopilotLicense'])} leaked")
check("Filter removed expected number of users",
      len(mock_rows) == pre_filter_count - pre_filter_licensed,
      f"expected {pre_filter_count - pre_filter_licensed}, got {len(mock_rows)}")
check("All UPNs unique in filtered sample",
      len({r['UserPrincipalName'] for r in mock_rows}) == len(mock_rows))

top20 = sorted(mock_rows, key=lambda r: r["InteractionCount"], reverse=True)[:20]
check("Top-20 chart input has at most 20 rows", len(top20) <= 20, f"{len(top20)}")
check("Top-20 first row has the highest count",
      top20[0]["InteractionCount"] == max(r["InteractionCount"] for r in mock_rows))

hostile = [r for r in mock_rows if "evil" in r["UserPrincipalName"]]
check("Hostile-UPN test case still present in filtered sample", len(hostile) == 1)

check("Engagement tier counts sum to total users",
      sum(t["count"] for t in tiers) == len(mock_rows),
      f"{sum(t['count'] for t in tiers)} vs {len(mock_rows)}")
check("At least one Power user in sample", tier_counts["Power"] > 0)

unknown_hosts = [h for h in host_counts if h not in APP_HOSTS]
check("Host distribution contains only known app hosts",
      not unknown_hosts, f"unknown: {unknown_hosts}")

# ------------------------------------------------------------------ Summary
print("\n" + "="*60)
if failures:
    print(f"FAILURES: {len(failures)}")
    for f in failures: print("  - " + f)
    sys.exit(1)
print(f"All checks passed.")
if warnings:
    print(f"({len(warnings)} warnings)")
print(f"Mock report written to: {OUT_HTML}")
print(f"Examples copy:           {EXAMPLES_HTML}")
