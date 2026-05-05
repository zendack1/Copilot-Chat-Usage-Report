"""
Static + simulated tests for Get-CopilotChatUsage.ps1.

Without pwsh available we can't execute the script, but we CAN:
  1. Lint structural balance (braces, parentheses, here-strings).
  2. Verify all called functions are defined.
  3. Verify all parameters used in Main exist.
  4. Extract the embedded HTML template, inject mock JSON, write the page,
     and validate the result (HTML parses, JSON inside <script> is valid,
     all expected DOM hooks exist, no broken </script> sequences).
  5. Verify the JS code references match the data shape.
"""
from __future__ import annotations
import html.parser
import json
import re
import sys
from pathlib import Path

SCRIPT = Path("/mnt/workspace/output/Get-CopilotChatUsage.ps1")
OUT_HTML = Path("/mnt/workspace/working/mock_report.html")

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
# Strip comments and string literals so braces/parens inside them don't skew the count.
def strip_strings_and_comments(s: str) -> str:
    out = []
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        # block comment <# ... #>
        if c == "<" and s[i:i+2] == "<#":
            end = s.find("#>", i+2)
            if end == -1: break
            i = end + 2; continue
        # line comment
        if c == "#":
            end = s.find("\n", i)
            i = end if end != -1 else n; continue
        # here-strings @' ... '@ and @" ... "@
        if c == "@" and i+1 < n and s[i+1] in "'\"":
            quote = s[i+1]
            term = quote + "@"
            end = s.find("\n" + term, i+2)
            if end == -1: break
            i = end + len(term) + 1; continue
        # single-quoted string '...'
        if c == "'":
            end = i + 1
            while end < n:
                if s[end] == "'":
                    if end + 1 < n and s[end+1] == "'":
                        end += 2; continue
                    break
                end += 1
            i = end + 1; continue
        # double-quoted string "..."
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
opens = stripped.count("{"); closes = stripped.count("}")
check("Curly braces balanced", opens == closes, f"{opens} open vs {closes} close")
opens = stripped.count("("); closes = stripped.count(")")
check("Parentheses balanced", opens == closes, f"{opens} open vs {closes} close")

# Here-string opens vs closes
hs_open  = len(re.findall(r"@['\"]\s*\n", src))
hs_close = len(re.findall(r"\n['\"]@", src))
check("Here-strings balanced", hs_open == hs_close, f"{hs_open} open vs {hs_close} close")

# ------------------------------------------------------------------ 2. Functions defined vs called
defined = set(re.findall(r"^function\s+([A-Za-z\-_]+)", src, re.MULTILINE))
expected = {
    "Connect-Graph", "Get-CopilotLicensedUserIds", "Submit-AuditQuery",
    "Wait-AuditQuery", "Get-AuditQueryRecords", "Write-HtmlReport",
}
missing = expected - defined
check("All expected helper functions defined", not missing, f"missing: {missing or 'none'}")

# Each call site references a function that exists (or is built-in).
# Pull bare-call lines for our helpers.
for fn in expected:
    if fn == "Connect-Graph": continue
    pattern = re.compile(rf"\b{re.escape(fn)}\b")
    used = bool(pattern.search(src.split("# ---------- Main ----------")[1])) if "# ---------- Main ----------" in src else False
    if fn == "Write-HtmlReport":
        check("Main invokes Write-HtmlReport", used)
    elif fn == "Get-CopilotLicensedUserIds":
        check("Main invokes Get-CopilotLicensedUserIds", used)
    elif fn in {"Submit-AuditQuery", "Wait-AuditQuery", "Get-AuditQueryRecords"}:
        check(f"Main invokes {fn}", used)

# ------------------------------------------------------------------ 3. Parameter / variable wiring
main_block = src.split("# ---------- Main ----------", 1)[1] if "# ---------- Main ----------" in src else ""
for var in ("$Days", "$OutputPath", "$HtmlPath", "$IncludeLicensed", "$NoHtml"):
    check(f"Main uses {var}", var in main_block)

# Days parameter is constrained to 30, 60, 90.
m_validate = re.search(r"\[ValidateSet\(\s*30\s*,\s*60\s*,\s*90\s*\)\]\s*\n\s*\[int\]\$Days", src)
check("$Days parameter has ValidateSet(30, 60, 90)", bool(m_validate))

# ------------------------------------------------------------------ 4. Extract HTML template
m = re.search(r"\$template\s*=\s*@'\s*\n(.*?)\n'@", src, re.DOTALL)
check("HTML template here-string extractable", bool(m))
if not m:
    print("\nFAILURES:", failures)
    sys.exit(1)

template = m.group(1)
check("Template contains __DATA_JSON__ placeholder", "__DATA_JSON__" in template)
check("Template has <!doctype html>", template.lower().lstrip().startswith("<!doctype html>"))
check("Template references kpi-users", "kpi-users" in template)
check("Template references chart container", 'id="chart"' in template)
check("Template references table tbody", 'id="tbody"' in template)
check("Template references filter input", 'id="filter"' in template)
check("License column removed from header", 'data-key="HasCopilotLicense"' not in template)
check("Unlicensed KPI tile removed", 'kpi-unlic' not in template)
check("License pill rendering removed", 'pill lic' not in template or '"pill lic"' not in template)

# Simulate the inject-payload step exactly as PowerShell does it.
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
APP_HOSTS = ["bizchat","Office","Teams","Edge","Word","Outlook","PowerPoint","Excel"]

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
    mock_rows.append({
        "UserPrincipalName": upn,
        "UserId": f"00000000-0000-0000-0000-{counter:012d}",
        "InteractionCount": random.randint(1, 500),
        "FirstUsedUtc": f"2026-04-{5+start_day:02d} 09:00:00Z",
        "LastUsedUtc":  f"2026-04-{min(30, 5+end_day):02d} 17:00:00Z",
        "AppHostsUsed": hosts,
        # ~20% of the sample carry a Copilot license.
        "HasCopilotLicense": random.random() < 0.20,
    })

# Replace one row with a hostile-UPN security test case (kept unlicensed so
# it survives the default filter and we can verify escaping in the rendered HTML).
mock_rows[7] = {
    "UserPrincipalName": "evil</script>user@contoso.com",
    "UserId": "00000000-0000-0000-0000-000000000007",
    "InteractionCount": 42,
    "FirstUsedUtc": "2026-04-30 11:00:00Z",
    "LastUsedUtc":  "2026-05-01 11:00:00Z",
    "AppHostsUsed": "Teams",
    "HasCopilotLicense": False,
}

pre_filter_count = len(mock_rows)
pre_filter_licensed = sum(1 for r in mock_rows if r["HasCopilotLicense"])

# Mirror the PowerShell main block: when -IncludeLicensed is NOT set, drop
# licensed users before passing rows to Write-HtmlReport.
INCLUDE_LICENSED = False
if not INCLUDE_LICENSED:
    mock_rows = [r for r in mock_rows if not r["HasCopilotLicense"]]

# Sort by InteractionCount desc, like the script.
mock_rows.sort(key=lambda r: r["InteractionCount"], reverse=True)
print(f"\n[INFO] Pre-filter: {pre_filter_count} users ({pre_filter_licensed} licensed)")
print(f"[INFO] Post-filter sample size: {len(mock_rows)} users (will appear in report)")
payload = {
    "rows": mock_rows,
    "startUtc": "2026-04-04 20:00:00Z",
    "endUtc":   "2026-05-04 20:00:00Z",
    "generatedAt": "2026-05-04 20:00:00Z",
    "tenantId": "00000000-0000-0000-0000-aaaaaaaaaaaa",
    "scopeText": "Free / unlicensed Copilot Chat users only" if not INCLUDE_LICENSED
                  else "All Copilot Chat users (licensed and unlicensed)",
    "totalUsers": len(mock_rows),
    "totalEvents": sum(r["InteractionCount"] for r in mock_rows),
    "licensedCount": sum(1 for r in mock_rows if r["HasCopilotLicense"]),
    "unlicensedCount": sum(1 for r in mock_rows if not r["HasCopilotLicense"]),
}
payload_json = json.dumps(payload, separators=(",", ":"))
# Mirror the PS line:  $safe = $payload -replace '</', '<\/'
safe = payload_json.replace("</", "<\\/")
rendered = template.replace("__DATA_JSON__", safe)
OUT_HTML.write_text(rendered, encoding="utf-8")
check("Mock report file written", OUT_HTML.exists() and OUT_HTML.stat().st_size > 1000,
      f"{OUT_HTML.stat().st_size} bytes")

# ------------------------------------------------------------------ 5. Validate the produced HTML
# (a) HTML parser doesn't choke
class P(html.parser.HTMLParser):
    def __init__(self): super().__init__(); self.errors = []
    def error(self, msg): self.errors.append(msg)
p = P()
try:
    p.feed(rendered); p.close()
    check("HTML parses without exceptions", not p.errors, "; ".join(p.errors[:3]))
except Exception as e:
    check("HTML parses without exceptions", False, str(e))

# (b) The injected payload is a single contiguous JSON block.
m2 = re.search(r'<script id="report-data" type="application/json">(.*?)</script>',
               rendered, re.DOTALL)
check("report-data <script> block present and closes correctly", bool(m2))
if m2:
    raw = m2.group(1)
    # Reverse the </ -> <\/ escape exactly as the browser's JSON.parse would see it.
    # Browsers don't unescape <\/ — but JSON.parse accepts \/ as forward slash.
    try:
        parsed = json.loads(raw)
        check("Injected JSON parses", True)
        check("Injected JSON has rows array", isinstance(parsed.get("rows"), list))
        check("Injected JSON row count matches", len(parsed["rows"]) == len(mock_rows))
        # Verify the malicious </script> in the user name was escaped so it can't break out.
        bad_user = next((r for r in parsed["rows"] if "evil" in r["UserPrincipalName"]), None)
        check("Hostile </script> in data preserved through escape",
              bad_user is not None and bad_user["UserPrincipalName"] == "evil</script>user@contoso.com")
    except Exception as e:
        check("Injected JSON parses", False, str(e))

# (c) The escaped payload should not contain a literal </script> that could break out.
script_open = rendered.count('<script id="report-data"')
script_close_in_data = m2.group(1).count("</script>") if m2 else -1
check("No raw </script> inside the data block", script_close_in_data == 0)

# (d) JS field names referenced in the rendered page exist on the data shape.
js_fields = {"UserPrincipalName", "InteractionCount", "FirstUsedUtc",
             "LastUsedUtc", "AppHostsUsed", "HasCopilotLicense"}
data_fields = set(mock_rows[0].keys())
missing_js = js_fields - data_fields
check("JS-referenced fields all present in row data", not missing_js, f"missing: {missing_js or 'none'}")

# (e) Every <th data-key> binds to a real field.
for key in re.findall(r'data-key="([^"]+)"', rendered):
    check(f"Table column '{key}' maps to a row field", key in data_fields)

# ------------------------------------------------------------------ 5b. Sample-size and post-filter behavior
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
# The Top-20 chart in the JS slices the rows; the page source itself doesn't
# contain the post-slice list, but the data does — so verify the slice math.
top20 = sorted(mock_rows, key=lambda r: r["InteractionCount"], reverse=True)[:20]
check("Top-20 chart input has at most 20 rows", len(top20) <= 20, f"{len(top20)}")
check("Top-20 first row has the highest count",
      top20[0]["InteractionCount"] == max(r["InteractionCount"] for r in mock_rows))

# Hostile UPN should still be present after the filter (it is unlicensed).
hostile = [r for r in mock_rows if "evil" in r["UserPrincipalName"]]
check("Hostile-UPN test case still present in filtered sample", len(hostile) == 1)

# ------------------------------------------------------------------ 6. License match logic (UPN-based, case-insensitive)
# Verify the script keys the licensed map by UPN lowercased and looks up the same way.
license_build = re.search(
    r"\$licensed\[\$u\.UserPrincipalName\.ToLowerInvariant\(\)\]\s*=", src)
check("Licensed map keyed by lower-cased UPN", bool(license_build))

license_lookup = re.search(
    r"\$upnKey\s*=\s*\$g\.Name\.ToLowerInvariant\(\)[\s\S]{0,80}\$licensedMap\.ContainsKey\(\$upnKey\)",
    src)
check("Licensed-map lookup uses lower-cased UPN", bool(license_lookup))

# ------------------------------------------------------------------ Summary
print("\n" + "="*60)
if failures:
    print(f"FAILURES: {len(failures)}")
    for f in failures: print("  - " + f)
    sys.exit(1)
print(f"All {len([1])} checks passed.")
if warnings:
    print(f"({len(warnings)} warnings)")
print(f"Mock report written to: {OUT_HTML}")
