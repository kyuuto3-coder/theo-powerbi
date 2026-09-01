# Power BI Skill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A zero-install Claude Code project skill that turns `rawdata/*.csv|xlsx` + a coworker's description into an openable Power BI Project (`.pbip`), using PowerShell 5.1 scripts so Claude only reads a compact profile and writes one small spec.

**Architecture:** Two PowerShell entry points (`profile-data.ps1`, `build-pbip.ps1`) share small dot-sourced libraries under `scripts/lib/`. The profiler reads xlsx (zip+XML) and csv (encoding-detected) into a compact text summary. The builder validates `report-spec.json` and deterministically emits a TMDL semantic model + PBIR report from an in-code visual catalog and a handful of static templates copied from the WWI reference project. A dependency-free test harness runs on Windows PowerShell 5.1 (the real target) via SSH to the Parallels VM.

**Tech Stack:** Windows PowerShell 5.1 (also PowerShell 7 compatible), .NET Framework (`System.IO.Compression`, `Microsoft.VisualBasic.FileIO.TextFieldParser`), PBIR report format v2.0.0 (visualContainer 2.12.0 / page 2.1.0 / report 3.3.0), TMDL (compatibilityLevel 1606), Power BI Desktop for acceptance.

**Spec:** `docs/superpowers/specs/2026-09-01-powerbi-skill-design.md`

**Repo root (all paths below are relative to it):** `/Users/theo/Desktop/Claude/powerbi/skills` — remote `https://github.com/kyuuto3-coder/theo-powerbi.git`, branch `main`.

**Reference project (read-only, outside repo):** `/Users/theo/Desktop/Claude/powerbi/worldwideimporters/` (PBIR + TMDL as written by Power BI Desktop).

---

## Conventions used in every task

- **Running PowerShell for tests:** macOS has no `pwsh`. Run all PowerShell on the Parallels VM (Windows PowerShell 5.1) through the helper created in Task 1: `bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh <script-path-relative-to-repo> [args]`. It maps the repo to `\\Mac\Home\Desktop\Claude\powerbi\skills` and runs `powershell -NoProfile -ExecutionPolicy Bypass -File …`.
- **PowerShell 5.1 rules:** no ternary `?:`, no `??`, no `ConvertFrom-Json -AsHashtable`, no `Join-String`. Use `[ordered]@{}` for any object serialized with `ConvertTo-Json` so key order is stable. Build arrays with `New-Object System.Collections.Generic.List[object]` and `.ToArray()` — never rely on pipeline output for arrays (single items unroll). Every file is written with `Write-Utf8File` (UTF-8, no BOM, CRLF).
- **Commits:** after each task, `git add` the listed files and commit with the message given. Commit trailer on every commit:

```
Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Q7hWZrWvGjmj7GtPVa6hBf
```

## File structure (final)

```
.gitignore
README.md                                  Korean coworker guide
CLAUDE.md                                  points Claude at the skill, Korean
rawdata/.gitkeep                           coworker input
output/.gitkeep                            generated projects
docs/superpowers/{specs,plans}/…
.claude/skills/powerbi/
  SKILL.md                                 workflow + token rules
  reference/spec-schema.md                 report-spec.json contract + example
  reference/visual-catalog.md              visual → roles, when to use
  reference/dax-patterns.md                measure recipes
  examples/app-data-spec.json              worked flat-table example
  templates/report.json                    copied from WWI reference
  templates/theme/Fluent2-CY26SU08.json    copied from WWI reference
  scripts/profile-data.ps1                 entry: rawdata → profile text
  scripts/build-pbip.ps1                   entry: spec → PBIP
  scripts/lib/io.ps1                       file writing, ids, paths
  scripts/lib/xlsx.ps1                     xlsx reader (zip + XML)
  scripts/lib/csvread.ps1                  csv encoding/delimiter detection + reader
  scripts/lib/infer.ps1                    header/type/key/dictionary inference + formatting
  scripts/lib/spec.ps1                     spec loading, field-ref resolution, validation
  scripts/lib/tmdl.ps1                     semantic model emitter (M queries, TMDL)
  scripts/lib/pbir.ps1                     report emitter (visual catalog, grid, filters)
  scripts/tests/run-tests.ps1              harness (no Pester)
  scripts/tests/*.test.ps1                 test files
  scripts/tests/fixtures/                  apps.xlsx, sales_cp949.csv, regions.csv, specs
  scripts/tests/fixtures/make-fixtures.py  dev-only generator (python3 + openpyxl on the Mac)
```

---

### Task 1: Repo scaffold, dev helper, static templates

**Files:**
- Create: `.gitignore`, `rawdata/.gitkeep`, `output/.gitkeep`, `CLAUDE.md`
- Create: `.claude/skills/powerbi/templates/report.json`, `.claude/skills/powerbi/templates/theme/Fluent2-CY26SU08.json` (copies)
- Create (outside repo): `/Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh`

- [ ] **Step 1: Write `.gitignore` and placeholders**

```gitignore
# coworker-owned folders: never committed
rawdata/*
!rawdata/.gitkeep
output/*
!output/.gitkeep

# local Claude Code / MCP settings
.mcp.json
.claude/settings.local.json

# Power BI Desktop caches
**/.pbi/localSettings.json
**/.pbi/cache.abf

# OS
.DS_Store
Thumbs.db
```

```bash
cd /Users/theo/Desktop/Claude/powerbi/skills
touch rawdata/.gitkeep && mkdir -p output && touch output/.gitkeep
```

- [ ] **Step 2: Write `CLAUDE.md`**

```markdown
# 이 폴더의 규칙

- 데이터 분석, 대시보드, 리포트, Power BI, 차트, 시각화 요청이 오면 **반드시 `powerbi` 스킬(.claude/skills/powerbi/SKILL.md)** 을 먼저 읽고 그 절차를 따른다.
- 사용자에게는 **한국어**로 말한다. 스킬 내부 문서는 영어이다.
- `rawdata/`, `output/`, `templates/` 안의 파일은 직접 열어보지 않는다 (토큰 절약). 스크립트가 대신 읽는다.
```

- [ ] **Step 3: Copy static templates from the WWI reference**

```bash
cd /Users/theo/Desktop/Claude/powerbi/skills
mkdir -p .claude/skills/powerbi/templates/theme
cp ../worldwideimporters/WWI.Report/definition/report.json .claude/skills/powerbi/templates/report.json
cp ../worldwideimporters/WWI.Report/StaticResources/SharedResources/BaseThemes/Fluent2-CY26SU08.json .claude/skills/powerbi/templates/theme/Fluent2-CY26SU08.json
head -c 3 .claude/skills/powerbi/templates/report.json | xxd   # expect 7b0d 0a (no BOM)
```

- [ ] **Step 4: Create the VM runner helper (outside the repo)**

`/Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh`:

```bash
#!/usr/bin/env bash
# Run a PowerShell script from the skills repo on the Parallels Windows VM (Windows PowerShell 5.1).
# usage: vm-ps.sh <path relative to repo root> [args...]
set -euo pipefail
REL="$1"; shift
WIN_REPO='\\Mac\Home\Desktop\Claude\powerbi\skills'
WIN_PATH="$WIN_REPO\\$(echo "$REL" | sed 's#/#\\#g')"
ARGS="$*"
ssh -q -i ~/.ssh/parallels_win11 -o BatchMode=yes -o ConnectTimeout=10 theo@10.211.55.3 \
  "powershell -NoProfile -ExecutionPolicy Bypass -Command \"[Console]::OutputEncoding=[Text.Encoding]::UTF8; & '$WIN_PATH' $ARGS; exit \$LASTEXITCODE\""
```

```bash
chmod +x /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh
```

- [ ] **Step 5: Smoke-test the helper**

Create a throwaway `scripts/tests/hello.ps1` containing `"hello from $($PSVersionTable.PSVersion) at $PSScriptRoot"`, run:

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/hello.ps1
```

Expected: `hello from 5.1.… at \\Mac\Home\Desktop\Claude\powerbi\skills\.claude\skills\powerbi\scripts\tests`. Then delete `hello.ps1`.

- [ ] **Step 6: Commit**

```bash
git add .gitignore rawdata/.gitkeep output/.gitkeep CLAUDE.md .claude/skills/powerbi/templates
git commit -m "chore: repo scaffold, CLAUDE.md, static PBIR templates"
```

---

### Task 2: `lib/io.ps1` + test harness

**Files:**
- Create: `.claude/skills/powerbi/scripts/lib/io.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/run-tests.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/io.test.ps1`

- [ ] **Step 1: Write the harness**

`scripts/tests/run-tests.ps1`:

```powershell
# Minimal test harness (no Pester dependency). Runs every *.test.ps1 in this folder.
param([string]$Filter = '*')
$ErrorActionPreference = 'Stop'
$script:TestsRun = 0; $script:TestsFailed = 0; $script:Failures = New-Object System.Collections.Generic.List[string]

function Test-Case {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][scriptblock]$Body)
    $script:TestsRun++
    try { & $Body; Write-Host ("  PASS  " + $Name) }
    catch { $script:TestsFailed++; $script:Failures.Add("$Name :: $($_.Exception.Message)"); Write-Host ("  FAIL  " + $Name + "`n        " + $_.Exception.Message) -ForegroundColor Red }
}
function Assert-True  { param($Condition, [string]$Message = 'expected true')  if (-not $Condition) { throw $Message } }
function Assert-Equal { param($Expected, $Actual, [string]$Message = '')
    if ("$Expected" -ne "$Actual") { throw ("expected [" + $Expected + "] got [" + $Actual + "] " + $Message) } }
function Assert-Match { param([string]$Pattern, [string]$Text, [string]$Message = '')
    if ($Text -notmatch $Pattern) { throw ("expected match /" + $Pattern + "/ in: " + $Text.Substring(0, [Math]::Min(300, $Text.Length)) + " " + $Message) } }
function Assert-Throws { param([scriptblock]$Body, [string]$Pattern = '.')
    $threw = $false; try { & $Body } catch { $threw = $true; if ($_.Exception.Message -notmatch $Pattern) { throw ("threw but message [" + $_.Exception.Message + "] !~ /" + $Pattern + "/") } }
    if (-not $threw) { throw 'expected an exception' } }

$fixtures = Join-Path $PSScriptRoot 'fixtures'
$libDir   = Join-Path (Split-Path -Parent $PSScriptRoot) 'lib'
$tmpRoot  = Join-Path ([IO.Path]::GetTempPath()) ('pbi-skill-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null

foreach ($file in Get-ChildItem -Path $PSScriptRoot -Filter "$Filter.test.ps1" | Sort-Object Name) {
    Write-Host ("== " + $file.Name)
    . $file.FullName
}
Write-Host ""
Write-Host ("{0} tests, {1} failed" -f $script:TestsRun, $script:TestsFailed)
foreach ($f in $script:Failures) { Write-Host ("  - " + $f) -ForegroundColor Red }
Remove-Item -Recurse -Force $tmpRoot -ErrorAction SilentlyContinue
if ($script:TestsFailed -gt 0) { exit 1 } else { exit 0 }
```

- [ ] **Step 2: Write the failing test**

`scripts/tests/io.test.ps1`:

```powershell
. (Join-Path $libDir 'io.ps1')

Test-Case 'Write-Utf8File writes UTF-8 without BOM and CRLF' {
    $p = Join-Path $tmpRoot 'a\b\c.txt'
    Write-Utf8File -Path $p -Content "line1`nline2 한글"
    $bytes = [IO.File]::ReadAllBytes($p)
    Assert-True ($bytes[0] -ne 0xEF) 'must not start with BOM'
    $text = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
    Assert-Equal "line1`r`nline2 한글" $text
}
Test-Case 'New-HexId returns 20 lowercase hex chars, unique' {
    $a = New-HexId; $b = New-HexId
    Assert-Match '^[0-9a-f]{20}$' $a
    Assert-True ($a -ne $b) 'ids must differ'
}
Test-Case 'ConvertTo-JsonFile round-trips ordered objects' {
    $p = Join-Path $tmpRoot 'obj.json'
    ConvertTo-JsonFile -Object ([ordered]@{ b = 1; a = @('x') }) -Path $p
    $raw = [IO.File]::ReadAllText($p)
    Assert-True ($raw.IndexOf('"b"') -lt $raw.IndexOf('"a"')) 'key order preserved'
    $back = $raw | ConvertFrom-Json
    Assert-Equal 'x' $back.a[0]
}
Test-Case 'Get-WorkspaceRoot resolves four levels above scripts/' {
    $root = Get-WorkspaceRoot -ScriptRoot (Split-Path -Parent $libDir)
    Assert-True (Test-Path (Join-Path $root 'CLAUDE.md')) "root should contain CLAUDE.md, got $root"
}
```

- [ ] **Step 3: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter io
```
Expected: error that `lib\io.ps1` cannot be found (exit 1).

- [ ] **Step 4: Write `lib/io.ps1`**

```powershell
# io.ps1 — file writing, identifiers, paths. Dot-source from scripts.
$ErrorActionPreference = 'Stop'

function Write-Utf8File {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][AllowEmptyString()][string]$Content)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $text = $Content -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($false)))
}

function New-HexId { -join ((1..20) | ForEach-Object { '{0:x}' -f (Get-Random -Maximum 16) }) }

function New-LineageTag { [guid]::NewGuid().ToString() }

function ConvertTo-JsonFile {
    param([Parameter(Mandatory)]$Object, [Parameter(Mandatory)][string]$Path)
    Write-Utf8File -Path $Path -Content ($Object | ConvertTo-Json -Depth 64)
}

function Get-WorkspaceRoot {
    # <root>/.claude/skills/powerbi/scripts → <root>
    param([Parameter(Mandatory)][string]$ScriptRoot)
    [System.IO.Path]::GetFullPath((Join-Path $ScriptRoot '../../../..'))
}

function Get-Prop {
    # Safe property read for PSCustomObject/hashtable; returns $Default when missing or null.
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } else { return $Default } }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

function ConvertTo-Array {
    # Always returns object[] (empty for $null), never unrolls.
    param($Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Array]) { return $Value }
    return ,@($Value)
}
```

- [ ] **Step 5: Run to verify it passes**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter io
```
Expected: `4 tests, 0 failed`, exit 0.

- [ ] **Step 6: Commit**

```bash
git add .claude/skills/powerbi/scripts
git commit -m "feat(scripts): io helpers and dependency-free test harness"
```

---

### Task 3: Test fixtures

**Files:**
- Create: `.claude/skills/powerbi/scripts/tests/fixtures/make-fixtures.py`
- Generated + committed: `fixtures/apps.xlsx`, `fixtures/sales_cp949.csv`, `fixtures/regions.csv`

- [ ] **Step 1: Write the generator (dev-only, runs on the Mac with `/usr/bin/python3` + openpyxl)**

```python
#!/usr/bin/env python3
"""Generates test fixtures. Run once on the dev machine; outputs are committed."""
import csv, datetime, os, random
from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
random.seed(7)

# --- apps.xlsx: mimics rawdata/App Data.xlsx (unnamed index col A, header row 1, dictionary sheet at B2)
wb = Workbook()
ws = wb.active; ws.title = "Data"
ws.append([None, "id", "track_name", "size_bytes", "price", "rating_count_tot", "user_rating", "prime_genre", "release_date"])
genres = ["Games", "Education", "Weather", "Productivity", "Music"]
for i in range(40):
    ws.append([i, 281000000 + i, f"App {i}", random.randint(10_000_000, 900_000_000),
               random.choice([0, 0.99, 2.99, 4.99]), random.randint(0, 200000),
               random.choice([3, 3.5, 4, 4.5, 5]), random.choice(genres),
               datetime.date(2016, 1, 1) + datetime.timedelta(days=random.randint(0, 700))])
for row in ws.iter_rows(min_row=2, min_col=9, max_col=9):
    row[0].number_format = "yyyy-mm-dd"
wd = wb.create_sheet("Data Defintion")
wd["B2"] = "Name"; wd["C2"] = "Description"
for r, (n, d) in enumerate([("id", "App ID"), ("track_name", "Track Name"), ("size_bytes", "App Size in Bytes"),
                            ("price", "Price"), ("rating_count_tot", "Number of Rating"), ("user_rating", "User Rating"),
                            ("prime_genre", "Prime Genre"), ("release_date", "Release Date")], start=3):
    wd.cell(row=r, column=2, value=n); wd.cell(row=r, column=3, value=d)
wb.save(os.path.join(HERE, "apps.xlsx"))

# --- sales_cp949.csv + regions.csv: a tiny star schema with Korean headers in CP949
with open(os.path.join(HERE, "regions.csv"), "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f); w.writerow(["지역코드", "지역명", "권역"])
    for code, name, zone in [(1, "서울", "수도권"), (2, "부산", "영남"), (3, "대구", "영남"), (4, "광주", "호남")]:
        w.writerow([code, name, zone])
with open(os.path.join(HERE, "sales_cp949.csv"), "w", encoding="cp949", newline="") as f:
    w = csv.writer(f); w.writerow(["주문일", "지역코드", "제품", "수량", "매출액"])
    for i in range(30):
        d = datetime.date(2025, 1, 1) + datetime.timedelta(days=i * 7)
        w.writerow([d.isoformat(), random.randint(1, 4), random.choice(["A", "B", "C"]), random.randint(1, 20), round(random.uniform(1000, 90000), 2)])
print("fixtures written to", HERE)
```

- [ ] **Step 2: Generate and verify**

```bash
cd /Users/theo/Desktop/Claude/powerbi/skills/.claude/skills/powerbi/scripts/tests/fixtures
python3 make-fixtures.py && ls -la && file sales_cp949.csv && python3 -c "open('sales_cp949.csv','rb').read().decode('utf-8')" 2>&1 | tail -1
```
Expected: three files; `sales_cp949.csv: ISO-8859 text` or similar; the strict UTF-8 decode line ends with `UnicodeDecodeError` (proves it is not UTF-8).

- [ ] **Step 3: Commit**

```bash
git add .claude/skills/powerbi/scripts/tests/fixtures
git commit -m "test: fixtures (xlsx with dictionary sheet, CP949 csv star schema)"
```

---

### Task 4: `lib/xlsx.ps1` — workbook reader

**Files:**
- Create: `.claude/skills/powerbi/scripts/lib/xlsx.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/xlsx.test.ps1`

- [ ] **Step 1: Write the failing test**

```powershell
. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'xlsx.ps1')

Test-Case 'ConvertFrom-ColumnLetters maps A, Z, AA, AB' {
    Assert-Equal 1 (ConvertFrom-ColumnLetters 'A'); Assert-Equal 26 (ConvertFrom-ColumnLetters 'Z')
    Assert-Equal 27 (ConvertFrom-ColumnLetters 'AA'); Assert-Equal 28 (ConvertFrom-ColumnLetters 'AB')
}
Test-Case 'Read-XlsxWorkbook lists sheets in workbook order with dimensions' {
    $wb = Read-XlsxWorkbook -Path (Join-Path $fixtures 'apps.xlsx')
    Assert-Equal 2 $wb.Sheets.Count
    Assert-Equal 'Data' $wb.Sheets[0].Name
    Assert-Equal 'Data Defintion' $wb.Sheets[1].Name
    Assert-Equal 1 $wb.Sheets[0].FirstRow; Assert-Equal 1 $wb.Sheets[0].FirstCol
    Assert-Equal 41 $wb.Sheets[0].LastRow; Assert-Equal 9 $wb.Sheets[0].LastCol
    Assert-Equal 2 $wb.Sheets[1].FirstRow; Assert-Equal 2 $wb.Sheets[1].FirstCol
}
Test-Case 'cells carry typed values: shared strings, numbers, dates' {
    $wb = Read-XlsxWorkbook -Path (Join-Path $fixtures 'apps.xlsx')
    $data = $wb.Sheets[0]
    $hdr = $data.Rows[0]
    Assert-Equal 1 $hdr.R
    Assert-True ($null -eq $hdr.Cells[1]) 'A1 is empty'
    Assert-Equal 'id' $hdr.Cells[2]
    $r2 = $data.Rows[1]
    Assert-Equal 0 $r2.Cells[1]
    Assert-True ($r2.Cells[2] -is [double]) 'id is numeric'
    Assert-True ($r2.Cells[9] -is [datetime]) "release_date should be datetime, got $($r2.Cells[9].GetType().Name)"
    Assert-True ($r2.Cells[9].Year -ge 2016) 'date decoded from serial'
}
Test-Case 'MaxRows limits rows read but LastRow still reports the sheet size' {
    $wb = Read-XlsxWorkbook -Path (Join-Path $fixtures 'apps.xlsx') -MaxRows 5
    Assert-Equal 5 $wb.Sheets[0].Rows.Count
    Assert-Equal 41 $wb.Sheets[0].LastRow
}
```

- [ ] **Step 2: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter xlsx
```
Expected: fails loading `xlsx.ps1`.

- [ ] **Step 3: Write `lib/xlsx.ps1`**

```powershell
# xlsx.ps1 — minimal .xlsx reader (zip + XML). No Excel, no modules.
# Read-XlsxWorkbook returns @{ Sheets = @( @{ Name; FirstRow; FirstCol; LastRow; LastCol; Rows = @( @{ R=<absRow>; Cells=@{ <colIndex> = <value> } } ) } ) }
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function ConvertFrom-ColumnLetters {
    param([Parameter(Mandatory)][string]$Letters)
    $n = 0
    foreach ($ch in $Letters.ToCharArray()) { $n = $n * 26 + ([int][char]::ToUpper($ch) - 64) }
    return $n
}

function Get-XmlText {
    param($Node)
    if ($null -eq $Node) { return '' }
    if ($Node -is [string]) { return $Node }
    $t = $Node.'#text'
    if ($null -ne $t) { return [string]$t }
    return [string]$Node.InnerText
}

function Read-ZipEntryText {
    param($Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { $entry = $Zip.GetEntry($Name.TrimStart('/')) }
    if ($null -eq $entry) { return $null }
    $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

function Test-DateFormatCode {
    param([string]$Code)
    if (-not $Code) { return $false }
    $clean = $Code -replace '"[^"]*"', '' -replace '\[[^\]]*\]', '' -replace '\\.', ''
    return ($clean -match '(?i)[ymdh]')
}

function Read-XlsxStyles {
    param($Zip)
    $isDate = New-Object System.Collections.Generic.List[bool]
    $xml = Read-ZipEntryText -Zip $Zip -Name 'xl/styles.xml'
    if (-not $xml) { return $isDate }
    $doc = [xml]$xml
    $custom = @{}
    if ($doc.styleSheet.numFmts -and $doc.styleSheet.numFmts.numFmt) {
        foreach ($nf in $doc.styleSheet.numFmts.numFmt) { $custom[[int]$nf.numFmtId] = [string]$nf.formatCode }
    }
    $builtin = @(14,15,16,17,18,19,20,21,22,27,28,29,30,31,32,33,34,35,36,45,46,47,50,51,52,53,54,55,56,57,58)
    if ($doc.styleSheet.cellXfs -and $doc.styleSheet.cellXfs.xf) {
        foreach ($xf in $doc.styleSheet.cellXfs.xf) {
            $id = 0; if ($xf.numFmtId) { $id = [int]$xf.numFmtId }
            $d = $false
            if ($builtin -contains $id) { $d = $true }
            elseif ($custom.ContainsKey($id)) { $d = Test-DateFormatCode $custom[$id] }
            $isDate.Add($d)
        }
    }
    return $isDate
}

function Read-XlsxSharedStrings {
    param($Zip)
    $list = New-Object System.Collections.Generic.List[string]
    $xml = Read-ZipEntryText -Zip $Zip -Name 'xl/sharedStrings.xml'
    if (-not $xml) { return $list }
    $doc = [xml]$xml
    foreach ($si in $doc.sst.si) {
        $text = ''
        if ($null -ne $si.t) { $text = Get-XmlText $si.t }
        elseif ($null -ne $si.r) { foreach ($run in $si.r) { $text += (Get-XmlText $run.t) } }
        $list.Add($text)
    }
    return $list
}

function ConvertFrom-CellRef {
    # "AB12" → @{ Col=28; Row=12 }
    param([string]$Ref)
    if ($Ref -match '^([A-Z]+)(\d+)$') { return @{ Col = (ConvertFrom-ColumnLetters $Matches[1]); Row = [int]$Matches[2] } }
    throw "bad cell ref '$Ref'"
}

function Read-XlsxSheet {
    param($Zip, [string]$EntryName, [string]$SheetName, $Shared, $IsDateStyle, [int]$MaxRows)
    $doc = [xml](Read-ZipEntryText -Zip $Zip -Name $EntryName)
    $firstRow = 1; $firstCol = 1; $lastRow = 0; $lastCol = 0
    $dim = $null
    if ($doc.worksheet.dimension) { $dim = [string]$doc.worksheet.dimension.ref }
    if ($dim) {
        $parts = $dim.Split(':')
        $a = ConvertFrom-CellRef $parts[0]; $firstRow = $a.Row; $firstCol = $a.Col
        if ($parts.Count -gt 1) { $b = ConvertFrom-CellRef $parts[1]; $lastRow = $b.Row; $lastCol = $b.Col } else { $lastRow = $a.Row; $lastCol = $a.Col }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    $sheetData = $doc.worksheet.sheetData
    if ($sheetData -and $sheetData.row) {
        foreach ($row in $sheetData.row) {
            if ($rows.Count -ge $MaxRows) { break }
            $r = [int]$row.r
            $cells = @{}
            if ($row.c) {
                foreach ($c in $row.c) {
                    $ref = ConvertFrom-CellRef ([string]$c.r)
                    $t = [string]$c.t
                    $v = $null
                    if ($null -ne $c.v) { $v = Get-XmlText $c.v }
                    $val = $null
                    switch ($t) {
                        's'         { if ($v -ne '') { $val = $Shared[[int]$v] } }
                        'str'       { $val = $v }
                        'inlineStr' { if ($c.is) { $val = Get-XmlText $c.is.t } }
                        'b'         { $val = ($v -eq '1') }
                        'e'         { $val = $null }
                        default {
                            if ($v -ne $null -and $v -ne '') {
                                $d = [double]::Parse($v, [Globalization.CultureInfo]::InvariantCulture)
                                $s = -1; if ($c.s) { $s = [int]$c.s }
                                if ($s -ge 0 -and $s -lt $IsDateStyle.Count -and $IsDateStyle[$s]) { $val = [DateTime]::FromOADate($d) } else { $val = $d }
                            }
                        }
                    }
                    if ($null -ne $val -and -not ($val -is [string] -and $val -eq '')) { $cells[$ref.Col] = $val }
                    if ($ref.Col -gt $lastCol) { $lastCol = $ref.Col }
                }
            }
            $rows.Add(@{ R = $r; Cells = $cells })
            if ($r -gt $lastRow) { $lastRow = $r }
        }
    }
    return @{ Name = $SheetName; FirstRow = $firstRow; FirstCol = $firstCol; LastRow = $lastRow; LastCol = $lastCol; Rows = $rows.ToArray() }
}

function Read-XlsxWorkbook {
    param([Parameter(Mandatory)][string]$Path, [int]$MaxRows = 2001)
    $full = (Get-Item -LiteralPath $Path).FullName
    $zip = [System.IO.Compression.ZipFile]::OpenRead($full)
    try {
        $shared = Read-XlsxSharedStrings -Zip $zip
        $isDate = Read-XlsxStyles -Zip $zip
        $wbDoc = [xml](Read-ZipEntryText -Zip $zip -Name 'xl/workbook.xml')
        $relDoc = [xml](Read-ZipEntryText -Zip $zip -Name 'xl/_rels/workbook.xml.rels')
        $targets = @{}
        foreach ($rel in $relDoc.Relationships.Relationship) {
            $target = [string]$rel.Target
            if ($target.StartsWith('/')) { $target = $target.TrimStart('/') } else { $target = 'xl/' + $target }
            $targets[[string]$rel.Id] = $target
        }
        $sheets = New-Object System.Collections.Generic.List[object]
        $rns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        foreach ($s in $wbDoc.workbook.sheets.sheet) {
            $rid = $s.GetAttribute('id', $rns)
            $sheets.Add((Read-XlsxSheet -Zip $zip -EntryName $targets[$rid] -SheetName ([string]$s.name) -Shared $shared -IsDateStyle $isDate -MaxRows $MaxRows))
        }
        return @{ Sheets = $sheets.ToArray() }
    } finally { $zip.Dispose() }
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter xlsx
```
Expected: `4 tests, 0 failed`. If the date assertion fails, print `$IsDateStyle` and the cell's `s` attribute — openpyxl writes custom `yyyy-mm-dd` as a numFmt ≥ 164, which `Test-DateFormatCode` must flag.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/powerbi/scripts/lib/xlsx.ps1 .claude/skills/powerbi/scripts/tests/xlsx.test.ps1
git commit -m "feat(scripts): xlsx reader (zip+XML, shared strings, date styles)"
```

---

### Task 5: `lib/csvread.ps1` — encoding/delimiter detection + reader

**Files:**
- Create: `.claude/skills/powerbi/scripts/lib/csvread.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/csvread.test.ps1`

- [ ] **Step 1: Write the failing test**

```powershell
. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'csvread.ps1')

Test-Case 'Get-CsvEncoding detects CP949 and UTF-8' {
    Assert-Equal 949   (Get-CsvEncoding -Path (Join-Path $fixtures 'sales_cp949.csv'))
    Assert-Equal 65001 (Get-CsvEncoding -Path (Join-Path $fixtures 'regions.csv'))
}
Test-Case 'Get-CsvEncoding detects UTF-8 BOM' {
    $p = Join-Path $tmpRoot 'bom.csv'
    [IO.File]::WriteAllText($p, "a,b`n1,2`n", (New-Object System.Text.UTF8Encoding($true)))
    Assert-Equal 65001 (Get-CsvEncoding -Path $p)
}
Test-Case 'Get-CsvDelimiter picks the most frequent candidate on the header line' {
    Assert-Equal ',' (Get-CsvDelimiter -Path (Join-Path $fixtures 'regions.csv') -CodePage 65001)
    $p = Join-Path $tmpRoot 'semi.csv'; [IO.File]::WriteAllText($p, "a;b;c`n1;2;3`n")
    Assert-Equal ';' (Get-CsvDelimiter -Path $p -CodePage 65001)
}
Test-Case 'Read-CsvRows decodes Korean headers from CP949 and counts all rows' {
    $r = Read-CsvRows -Path (Join-Path $fixtures 'sales_cp949.csv') -CodePage 949 -Delimiter ',' -MaxRows 10
    Assert-Equal '주문일' $r.Rows[0][0]
    Assert-Equal '매출액' $r.Rows[0][4]
    Assert-Equal 10 $r.Rows.Count
    Assert-Equal 31 $r.TotalLines
}
Test-Case 'Read-CsvRows honours quoted fields with embedded delimiters' {
    $p = Join-Path $tmpRoot 'q.csv'; [IO.File]::WriteAllText($p, "name,note`n""Doe, John"",""says """"hi""""""`n")
    $r = Read-CsvRows -Path $p -CodePage 65001 -Delimiter ',' -MaxRows 10
    Assert-Equal 'Doe, John' $r.Rows[1][0]
    Assert-Equal 'says "hi"' $r.Rows[1][1]
}
```

- [ ] **Step 2: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter csvread
```
Expected: fails loading `csvread.ps1`.

- [ ] **Step 3: Write `lib/csvread.ps1`**

```powershell
# csvread.ps1 — encoding + delimiter detection and a bounded CSV reader.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic

function Get-CsvEncoding {
    # Returns 65001 (UTF-8) or 949 (CP949 / EUC-KR family). Heuristic: BOM → UTF-8; strict UTF-8 decode of the first 64 KB; else 949.
    param([Parameter(Mandatory)][string]$Path)
    $fs = [IO.File]::OpenRead($Path)
    try {
        $buf = New-Object byte[] 65536
        $n = $fs.Read($buf, 0, $buf.Length)
    } finally { $fs.Dispose() }
    if ($n -ge 3 -and $buf[0] -eq 0xEF -and $buf[1] -eq 0xBB -and $buf[2] -eq 0xBF) { return 65001 }
    $strict = New-Object System.Text.UTF8Encoding($false, $true)
    foreach ($trim in 0, 1, 2, 3) {
        $len = $n - $trim
        if ($len -le 0) { break }
        try { [void]$strict.GetString($buf, 0, $len); return 65001 } catch { }
    }
    return 949
}

function Get-CsvDelimiter {
    param([Parameter(Mandatory)][string]$Path, [int]$CodePage = 65001)
    $sr = New-Object IO.StreamReader($Path, [Text.Encoding]::GetEncoding($CodePage))
    try { $line = $sr.ReadLine() } finally { $sr.Dispose() }
    if ($null -eq $line) { return ',' }
    $best = ','; $bestCount = -1
    foreach ($d in ',', ';', "`t", '|') {
        $count = ($line.ToCharArray() | Where-Object { $_ -eq $d }).Count
        if ($count -gt $bestCount) { $best = $d; $bestCount = $count }
    }
    return $best
}

function Read-CsvRows {
    # Returns @{ Rows = @( string[] ... ) (first MaxRows physical records incl. header); TotalLines = number of non-empty lines in the file }
    param([Parameter(Mandatory)][string]$Path, [int]$CodePage = 65001, [string]$Delimiter = ',', [int]$MaxRows = 2001)
    $enc = [Text.Encoding]::GetEncoding($CodePage)
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, $enc)
    $rows = New-Object System.Collections.Generic.List[object]
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters($Delimiter)
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        while (-not $parser.EndOfData -and $rows.Count -lt $MaxRows) {
            $fields = $parser.ReadFields()
            $rows.Add([string[]]$fields)
        }
    } finally { $parser.Close() }
    $total = 0
    $sr = New-Object IO.StreamReader($Path, $enc)
    try { while ($null -ne ($l = $sr.ReadLine())) { if ($l.Trim().Length -gt 0) { $total++ } } } finally { $sr.Dispose() }
    return @{ Rows = $rows.ToArray(); TotalLines = $total }
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter csvread
```
Expected: `5 tests, 0 failed`.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/powerbi/scripts/lib/csvread.ps1 .claude/skills/powerbi/scripts/tests/csvread.test.ps1
git commit -m "feat(scripts): csv encoding/delimiter detection and bounded reader"
```

---

### Task 6: `lib/infer.ps1` — header/type/key/dictionary inference

**Files:**
- Create: `.claude/skills/powerbi/scripts/lib/infer.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/infer.test.ps1`

- [ ] **Step 1: Write the failing test**

```powershell
. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'infer.ps1')

Test-Case 'Get-ValueKind classifies typed and string values' {
    Assert-Equal 'empty'  (Get-ValueKind $null)
    Assert-Equal 'empty'  (Get-ValueKind '  ')
    Assert-Equal 'number' (Get-ValueKind 3.5)
    Assert-Equal 'number' (Get-ValueKind '1,234.50')
    Assert-Equal 'date'   (Get-ValueKind ([datetime]'2025-01-02'))
    Assert-Equal 'date'   (Get-ValueKind '2025-01-02')
    Assert-Equal 'date'   (Get-ValueKind '2025/1/2 13:45')
    Assert-Equal 'bool'   (Get-ValueKind 'TRUE')
    Assert-Equal 'text'   (Get-ValueKind 'Games')
}
Test-Case 'Find-HeaderRow finds first mostly-text row with ≥2 cells' {
    $grid = @(
        ,@($null, $null, $null),
        ,@($null, 'Name', 'Description'),
        ,@($null, 'id', 'App ID')
    )
    Assert-Equal 1 (Find-HeaderRow -Grid $grid)
    $grid2 = @( ,@('id', 'name', 'price'), ,@(1, 'a', 2.5) )
    Assert-Equal 0 (Find-HeaderRow -Grid $grid2)
}
Test-Case 'Get-ColumnProfile infers int64 / decimal / double / date / text with stats' {
    $ints  = Get-ColumnProfile -Values @(1, 2, 3, 3)
    Assert-Equal 'int64' $ints.Type; Assert-Equal 3 $ints.Distinct; Assert-Equal 1 $ints.Min; Assert-Equal 3 $ints.Max
    $dec   = Get-ColumnProfile -Values @('0.99', '2.5', $null)
    Assert-Equal 'decimal' $dec.Type; Assert-Equal 33 $dec.NullPct
    $dbl   = Get-ColumnProfile -Values @(0.123456, 1.5)
    Assert-Equal 'double' $dbl.Type
    $dates = Get-ColumnProfile -Values @('2025-01-01', '2025-02-01')
    Assert-Equal 'date' $dates.Type; Assert-Equal '2025-01-01' $dates.Min
    $text  = Get-ColumnProfile -Values @('a', 'b', 'a')
    Assert-Equal 'text' $text.Type; Assert-Equal 2 $text.Distinct; Assert-Equal 'a' $text.Samples[0]
}
Test-Case 'Get-ColumnProfile flags unique non-null columns as key candidates' {
    $k = Get-ColumnProfile -Values @(10, 11, 12)
    Assert-True $k.IsUnique 'unique ints'
    $nk = Get-ColumnProfile -Values @(10, 10, 12)
    Assert-True (-not $nk.IsUnique) 'duplicates'
}
Test-Case 'Test-DictionarySheet recognises Name/Description sheets' {
    Assert-True (Test-DictionarySheet -Headers @('Name', 'Description') -DataRowCount 8 -OtherHeaders @())
    Assert-True (Test-DictionarySheet -Headers @('항목', '설명') -DataRowCount 8 -OtherHeaders @())
    Assert-True (Test-DictionarySheet -Headers @('col', 'meaning') -DataRowCount 5 -OtherHeaders @('id','track_name','price') -FirstColumnValues @('id','track_name','price','x'))
    Assert-True (-not (Test-DictionarySheet -Headers @('id', 'name', 'price') -DataRowCount 500 -OtherHeaders @()))
}
Test-Case 'Find-KeyMatches pairs same-named columns where one side is unique' {
    $t1 = @{ Name = 'sales';   Columns = @( @{ Name = '지역코드'; IsUnique = $false }, @{ Name = '수량'; IsUnique = $false } ) }
    $t2 = @{ Name = 'regions'; Columns = @( @{ Name = '지역코드'; IsUnique = $true } ) }
    $m = @(Find-KeyMatches -Tables @($t1, $t2))
    Assert-Equal 1 $m.Count
    Assert-Equal 'sales.지역코드 -> regions.지역코드' $m[0]
}
```

- [ ] **Step 2: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter infer
```
Expected: fails loading `infer.ps1`.

- [ ] **Step 3: Write `lib/infer.ps1`**

```powershell
# infer.ps1 — value classification, header detection, column profiling, dictionary/key heuristics.
$ErrorActionPreference = 'Stop'
$script:Inv = [Globalization.CultureInfo]::InvariantCulture
$script:DateFormats = @('yyyy-M-d', 'yyyy/M/d', 'yyyy.M.d', 'yyyy-M-d H:mm', 'yyyy-M-d H:mm:ss', 'yyyy/M/d H:mm', 'yyyy/M/d H:mm:ss', 'yyyy-M-dTH:mm:ss', 'M/d/yyyy', 'M/d/yyyy H:mm', 'yyyyMMdd')

function Test-NumberString { param([string]$S) return ($S -match '^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$') }
function Test-BoolString   { param([string]$S) return ($S -match '^(?i)(true|false|yes|no)$') }
function ConvertTo-DateValue {
    param([string]$S)
    $out = [datetime]::MinValue
    if ([datetime]::TryParseExact($S.Trim(), [string[]]$script:DateFormats, $script:Inv, [Globalization.DateTimeStyles]::None, [ref]$out)) { return $out }
    if ($S -match '^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일$') { return (New-Object datetime ([int]$Matches[1]), ([int]$Matches[2]), ([int]$Matches[3])) }
    return $null
}
function Test-DateString { param([string]$S) return ($null -ne (ConvertTo-DateValue $S)) }
function ConvertTo-NumberValue { param([string]$S) return [double]::Parse($S.Replace(',', ''), $script:Inv) }

function Get-ValueKind {
    param($V)
    if ($null -eq $V) { return 'empty' }
    if ($V -is [datetime]) { return 'date' }
    if ($V -is [bool]) { return 'bool' }
    if ($V -is [double] -or $V -is [int] -or $V -is [long] -or $V -is [decimal] -or $V -is [single]) { return 'number' }
    $s = ([string]$V).Trim()
    if ($s -eq '') { return 'empty' }
    if (Test-NumberString $s) { return 'number' }
    if (Test-DateString $s) { return 'date' }
    if (Test-BoolString $s) { return 'bool' }
    return 'text'
}

function Find-HeaderRow {
    # $Grid: array of object[] rows. Returns 0-based index of the header row, or -1.
    param([Parameter(Mandatory)]$Grid)
    $limit = [Math]::Min(20, $Grid.Count)
    for ($i = 0; $i -lt $limit; $i++) {
        $row = $Grid[$i]
        $nonEmpty = 0; $text = 0
        for ($j = 0; $j -lt $row.Count; $j++) {
            $k = Get-ValueKind $row[$j]
            if ($k -ne 'empty') { $nonEmpty++ }
            if ($k -eq 'text') { $text++ }
        }
        if ($nonEmpty -ge 2 -and ($text / $nonEmpty) -ge 0.6) { return $i }
    }
    return -1
}

function Get-ColumnProfile {
    # Profiles one column's sampled values. Returns @{ Type; NullPct; Distinct; Samples; Min; Max; IsUnique; NonEmpty }
    param([Parameter(Mandatory)][AllowEmptyCollection()]$Values)
    $n = $Values.Count
    $counts = @{ empty = 0; number = 0; date = 0; bool = 0; text = 0 }
    $distinct = New-Object 'System.Collections.Generic.HashSet[string]'
    $samples = New-Object System.Collections.Generic.List[string]
    $allInt = $true; $maxDecimals = 0; $hasTime = $false
    $minN = $null; $maxN = $null; $minD = $null; $maxD = $null
    for ($i = 0; $i -lt $n; $i++) {
        $v = $Values[$i]
        $k = Get-ValueKind $v
        $counts[$k]++
        if ($k -eq 'empty') { continue }
        $str = $null
        switch ($k) {
            'number' {
                $d = if ($v -is [string]) { ConvertTo-NumberValue $v } else { [double]$v }
                if ([Math]::Abs($d - [Math]::Round($d)) -gt 1e-9) { $allInt = $false
                    $dec = ("{0:R}" -f $d); if ($dec -match '\.(\d+)$') { if ($Matches[1].Length -gt $maxDecimals) { $maxDecimals = $Matches[1].Length } } }
                if ($null -eq $minN -or $d -lt $minN) { $minN = $d }
                if ($null -eq $maxN -or $d -gt $maxN) { $maxN = $d }
                $str = $d.ToString('R', $script:Inv)
            }
            'date' {
                $dt = if ($v -is [datetime]) { $v } else { ConvertTo-DateValue $v }
                if ($dt.TimeOfDay.TotalSeconds -ne 0) { $hasTime = $true }
                if ($null -eq $minD -or $dt -lt $minD) { $minD = $dt }
                if ($null -eq $maxD -or $dt -gt $maxD) { $maxD = $dt }
                $str = $dt.ToString('yyyy-MM-dd')
            }
            default { $str = ([string]$v).Trim() }
        }
        if ($distinct.Add($str) -and $samples.Count -lt 2) { $samples.Add($str) }
    }
    $nonEmpty = $n - $counts.empty
    $type = 'text'
    if ($nonEmpty -gt 0) {
        if ($counts.number / $nonEmpty -ge 0.98) { if ($allInt) { $type = 'int64' } elseif ($maxDecimals -le 4) { $type = 'decimal' } else { $type = 'double' } }
        elseif ($counts.date / $nonEmpty -ge 0.95) { $type = if ($hasTime) { 'datetime' } else { 'date' } }
        elseif ($counts.bool -eq $nonEmpty) { $type = 'boolean' }
    }
    $min = $null; $max = $null
    if ($type -in 'int64', 'decimal', 'double') { $min = $minN; $max = $maxN }
    elseif ($type -in 'date', 'datetime') { if ($minD) { $min = $minD.ToString('yyyy-MM-dd'); $max = $maxD.ToString('yyyy-MM-dd') } }
    $nullPct = if ($n -gt 0) { [int][Math]::Round(100.0 * $counts.empty / $n) } else { 0 }
    return @{ Type = $type; NullPct = $nullPct; Distinct = $distinct.Count; Samples = $samples.ToArray(); Min = $min; Max = $max
              IsUnique = ($nonEmpty -gt 0 -and $counts.empty -eq 0 -and $distinct.Count -eq $nonEmpty); NonEmpty = $nonEmpty }
}

function Test-DictionarySheet {
    param([string[]]$Headers, [int]$DataRowCount, [string[]]$OtherHeaders, [string[]]$FirstColumnValues = @())
    $nonEmpty = @($Headers | Where-Object { $_ -and $_.Trim() -ne '' })
    if ($nonEmpty.Count -gt 3 -or $DataRowCount -gt 60) { return $false }
    $joined = ($nonEmpty -join ' ')
    $nameish = $joined -match '(?i)(name|field|column|컬럼|항목|변수|필드)'
    $descish = $joined -match '(?i)(desc|meaning|definition|설명|의미|정의)'
    if ($nameish -and $descish) { return $true }
    if ($OtherHeaders.Count -gt 0 -and $FirstColumnValues.Count -gt 0) {
        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($h in $OtherHeaders) { if ($h) { [void]$set.Add($h.Trim()) } }
        $hits = 0; foreach ($v in $FirstColumnValues) { if ($v -and $set.Contains($v.Trim())) { $hits++ } }
        if ($hits / $FirstColumnValues.Count -ge 0.7) { return $true }
    }
    return $false
}

function Get-NormalizedName { param([string]$Name) return (($Name -replace '[^\p{L}\p{N}]', '').ToLowerInvariant()) }

function Find-KeyMatches {
    # $Tables: @( @{ Name; Columns = @( @{ Name; IsUnique } ) } ). Returns strings "many.col -> one.col".
    param([Parameter(Mandatory)]$Tables)
    $out = New-Object System.Collections.Generic.List[string]
    for ($a = 0; $a -lt $Tables.Count; $a++) {
        for ($b = 0; $b -lt $Tables.Count; $b++) {
            if ($a -eq $b) { continue }
            foreach ($ca in $Tables[$a].Columns) {
                if ($ca.IsUnique) { continue }
                foreach ($cb in $Tables[$b].Columns) {
                    if (-not $cb.IsUnique) { continue }
                    if ((Get-NormalizedName $ca.Name) -eq (Get-NormalizedName $cb.Name)) {
                        $out.Add(("{0}.{1} -> {2}.{3}" -f $Tables[$a].Name, $ca.Name, $Tables[$b].Name, $cb.Name))
                    }
                }
            }
        }
    }
    return $out.ToArray()
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter infer
```
Expected: `6 tests, 0 failed`.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/powerbi/scripts/lib/infer.ps1 .claude/skills/powerbi/scripts/tests/infer.test.ps1
git commit -m "feat(scripts): type/header/key/dictionary inference"
```

---

### Task 7: `profile-data.ps1` — entry point

**Files:**
- Create: `.claude/skills/powerbi/scripts/profile-data.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/profile.test.ps1`

- [ ] **Step 1: Write the failing test**

```powershell
$profileScript = Join-Path (Split-Path -Parent $libDir) 'profile-data.ps1'

Test-Case 'profile of fixtures lists both files, dictionary sheet, encoding and key match' {
    $out = (& $profileScript -DataFolder $fixtures | Out-String)
    Assert-Match 'PROFILE .*\(3 files\)' $out
    Assert-Match 'FILE "apps.xlsx"' $out
    Assert-Match 'SHEET "Data"\s+rows=40\s+header_row=1\s+first_col=B\s+-> table candidate' $out
    Assert-Match '\(unnamed\)\s+int64' $out
    Assert-Match 'id\s+int64\s+0%\s+40\s+281000000, 281000001\s+min=281000000 max=281000039 KEY' $out
    Assert-Match 'release_date\s+date\s+0%' $out
    Assert-Match 'SHEET "Data Defintion"\s+rows=8\s+-> dictionary \(not loaded\)' $out
    Assert-Match 'id=App ID; track_name=Track Name' $out
    Assert-Match 'FILE "sales_cp949.csv"\s+encoding=949\s+delimiter=","\s+rows=30\s+header_row=1' $out
    Assert-Match '주문일\s+date' $out
    Assert-Match '매출액\s+decimal' $out
    Assert-Match 'KEY MATCHES.*\n.*sales_cp949.지역코드 -> regions.지역코드' $out
}
Test-Case 'profile of an empty folder exits 2 with a clear message' {
    $empty = Join-Path $tmpRoot 'empty'; New-Item -ItemType Directory -Path $empty -Force | Out-Null
    $out = (& $profileScript -DataFolder $empty 2>&1 | Out-String)
    Assert-Match 'no \.csv/\.xlsx files' $out
    Assert-Equal 2 $LASTEXITCODE
}
```

- [ ] **Step 2: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter profile
```
Expected: fails (script not found).

- [ ] **Step 3: Write `scripts/profile-data.ps1`**

```powershell
<#
.SYNOPSIS  Profiles every .csv/.xlsx in rawdata/ into a compact text summary for Claude.
.EXAMPLE   powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/profile-data.ps1
#>
param([string]$DataFolder = '', [int]$SampleRows = 2000, [int]$MaxColumns = 25)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib/io.ps1')
. (Join-Path $PSScriptRoot 'lib/xlsx.ps1')
. (Join-Path $PSScriptRoot 'lib/csvread.ps1')
. (Join-Path $PSScriptRoot 'lib/infer.ps1')

if (-not $DataFolder) { $DataFolder = Join-Path (Get-WorkspaceRoot -ScriptRoot $PSScriptRoot) 'rawdata' }
$files = @(Get-ChildItem -LiteralPath $DataFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.csv', '.xlsx', '.xlsm' } | Sort-Object Name)
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add(("PROFILE {0}  ({1} files)" -f $DataFolder, $files.Count))
if ($files.Count -eq 0) { $lines.Add('  no .csv/.xlsx files found - ask the user to put data files in rawdata/'); $lines | ForEach-Object { $_ }; exit 2 }

function ConvertTo-ColumnLetters { param([int]$Index) $s = ''; $n = $Index; while ($n -gt 0) { $r = ($n - 1) % 26; $s = [char](65 + $r) + $s; $n = [int](($n - 1) / 26) }; return $s }

function Format-Cell { param($V) if ($null -eq $V) { return '' } ; $s = [string]$V; if ($s.Length -gt 24) { $s = $s.Substring(0, 22) + '..' }; return $s }

function Get-GridProfile {
    # $Grid: object[] rows (absolute column indices, 0-based). Returns table profile hashtable or $null when no header.
    param($Grid, [int]$TotalDataRows, [string]$Label, [string]$ColPrefix)
    $h = Find-HeaderRow -Grid $Grid
    if ($h -lt 0) { return $null }
    $header = $Grid[$h]
    $width = 0; foreach ($row in $Grid) { if ($row.Count -gt $width) { $width = $row.Count } }
    $firstCol = -1
    for ($j = 0; $j -lt $header.Count; $j++) { if ((Get-ValueKind $header[$j]) -ne 'empty') { $firstCol = $j; break } }
    $cols = New-Object System.Collections.Generic.List[object]
    for ($j = 0; $j -lt $width; $j++) {
        $vals = New-Object System.Collections.Generic.List[object]
        for ($i = $h + 1; $i -lt $Grid.Count; $i++) { $row = $Grid[$i]; if ($j -lt $row.Count) { $vals.Add($row[$j]) } else { $vals.Add($null) } }
        $p = Get-ColumnProfile -Values $vals.ToArray()
        if ($p.NonEmpty -eq 0 -and (Get-ValueKind $header[$j]) -eq 'empty') { continue }   # fully empty column
        $name = if ($j -lt $header.Count -and (Get-ValueKind $header[$j]) -ne 'empty') { ([string]$header[$j]).Trim() } else { '(unnamed)' }
        $cols.Add(@{ Index = $j; Name = $name; Profile = $p })
    }
    return @{ Label = $Label; HeaderIndex = $h; FirstCol = $firstCol; Columns = $cols.ToArray(); DataRows = $TotalDataRows; SampleRows = ($Grid.Count - $h - 1) }
}

function Add-TableLines {
    param($T, [string]$ColPrefixMode)   # 'letters' for xlsx, 'index' for csv
    $lines.Add(("    {0,-4} {1,-24} {2,-9} {3,-6} {4,-8} {5,-28} {6}" -f 'col', 'name', 'type', 'nulls', 'distinct', 'sample', 'note'))
    $shown = 0
    foreach ($c in $T.Columns) {
        if ($shown -ge $MaxColumns) { $lines.Add(("    +{0} more columns" -f ($T.Columns.Count - $shown))); break }
        $shown++
        $p = $c.Profile
        $colLabel = if ($ColPrefixMode -eq 'letters') { ConvertTo-ColumnLetters ($c.Index + 1) } else { [string]($c.Index + 1) }
        $notes = New-Object System.Collections.Generic.List[string]
        if ($null -ne $p.Min -and $null -ne $p.Max) { $notes.Add(("min={0} max={1}" -f $p.Min, $p.Max)) }
        if ($c.Name -eq '(unnamed)') { $notes.Add('index? drop') }
        elseif ($p.IsUnique -and $T.SampleRows -ge $T.DataRows) { $notes.Add('KEY') }
        elseif ($p.IsUnique) { $notes.Add('KEY?') }
        $sample = (($p.Samples | ForEach-Object { if ($p.Type -eq 'text') { '"' + (Format-Cell $_) + '"' } else { Format-Cell $_ } }) -join ', ')
        $lines.Add(("    {0,-4} {1,-24} {2,-9} {3,-6} {4,-8} {5,-28} {6}" -f $colLabel, (Format-Cell $c.Name), $p.Type, ("{0}%" -f $p.NullPct), $p.Distinct, $sample, ($notes -join ' ')))
    }
}

$tables = New-Object System.Collections.Generic.List[object]   # for key matching: @{ Name; Columns=@(@{Name;IsUnique}) }
foreach ($f in $files) {
    if ($f.Extension -eq '.csv') {
        $cp = Get-CsvEncoding -Path $f.FullName
        $delim = Get-CsvDelimiter -Path $f.FullName -CodePage $cp
        $r = Read-CsvRows -Path $f.FullName -CodePage $cp -Delimiter $delim -MaxRows ($SampleRows + 1)
        $grid = @(); foreach ($row in $r.Rows) { $grid += ,([object[]]$row) }
        $hdrGuess = Find-HeaderRow -Grid $grid
        $dataRows = if ($hdrGuess -ge 0) { $r.TotalLines - $hdrGuess - 1 } else { $r.TotalLines }
        $t = Get-GridProfile -Grid $grid -TotalDataRows $dataRows -Label $f.Name -ColPrefixMode 'index'
        $delimLabel = if ($delim -eq "`t") { '\t' } else { $delim }
        if ($null -eq $t) { $lines.Add(("FILE ""{0}""  encoding={1}  -> no header row found (skipped)" -f $f.Name, $cp)); continue }
        $lines.Add(("FILE ""{0}""  encoding={1}  delimiter=""{2}""  rows={3}  header_row={4}  -> table candidate" -f $f.Name, $cp, $delimLabel, $t.DataRows, ($t.HeaderIndex + 1)))
        Add-TableLines -T $t -ColPrefixMode 'index'
        $tables.Add(@{ Name = [IO.Path]::GetFileNameWithoutExtension($f.Name); Columns = @($t.Columns | ForEach-Object { @{ Name = $_.Name; IsUnique = $_.Profile.IsUnique } }) })
    }
    else {
        $wb = Read-XlsxWorkbook -Path $f.FullName -MaxRows ($SampleRows + 40)
        $lines.Add(("FILE ""{0}""" -f $f.Name))
        # first pass: profiles per sheet
        $profiles = @{}
        foreach ($s in $wb.Sheets) {
            $grid = @()
            foreach ($row in $s.Rows) { $arr = New-Object object[] ($s.LastCol); foreach ($k in $row.Cells.Keys) { $arr[$k - 1] = $row.Cells[$k] }; $grid += ,$arr }
            if ($grid.Count -eq 0) { $profiles[$s.Name] = $null; continue }
            $h = Find-HeaderRow -Grid $grid
            $absHeaderRow = if ($h -ge 0) { $s.Rows[$h].R } else { 0 }
            $dataRows = if ($h -ge 0) { $s.LastRow - $absHeaderRow } else { 0 }
            $t = Get-GridProfile -Grid $grid -TotalDataRows $dataRows -Label $s.Name -ColPrefixMode 'letters'
            if ($null -ne $t) { $t.HeaderRowRelative = $absHeaderRow - $s.FirstRow + 1; $t.Grid = $grid }
            $profiles[$s.Name] = $t
        }
        $allHeaders = @(); foreach ($k in $profiles.Keys) { if ($profiles[$k]) { $allHeaders += @($profiles[$k].Columns | ForEach-Object { $_.Name }) } }
        foreach ($s in $wb.Sheets) {
            $t = $profiles[$s.Name]
            if ($null -eq $t) { $lines.Add(("  SHEET ""{0}""  -> empty (skipped)" -f $s.Name)); continue }
            $headers = @($t.Columns | ForEach-Object { $_.Name })
            $others = @($allHeaders | Where-Object { $headers -notcontains $_ })
            $firstColVals = @(); if ($t.Columns.Count -gt 0) { $ci = $t.Columns[0].Index; foreach ($row in $t.Grid[($t.HeaderIndex + 1)..($t.Grid.Count - 1)]) { if ($ci -lt $row.Count -and $null -ne $row[$ci]) { $firstColVals += [string]$row[$ci] } } }
            if (Test-DictionarySheet -Headers $headers -DataRowCount $t.DataRows -OtherHeaders $others -FirstColumnValues $firstColVals) {
                $lines.Add(("  SHEET ""{0}""  rows={1}  -> dictionary (not loaded)" -f $s.Name, $t.DataRows))
                $pairs = New-Object System.Collections.Generic.List[string]
                $c0 = $t.Columns[0].Index; $c1 = if ($t.Columns.Count -gt 1) { $t.Columns[1].Index } else { -1 }
                foreach ($row in $t.Grid[($t.HeaderIndex + 1)..($t.Grid.Count - 1)]) {
                    $k = if ($c0 -lt $row.Count) { [string]$row[$c0] } else { '' }
                    $v = if ($c1 -ge 0 -and $c1 -lt $row.Count) { [string]$row[$c1] } else { '' }
                    if ($k.Trim()) { $pairs.Add(("{0}={1}" -f $k.Trim(), $v.Trim())) }
                }
                $lines.Add('    ' + ($pairs -join '; '))
                continue
            }
            $lines.Add(("  SHEET ""{0}""  rows={1}  header_row={2}  first_col={3}  -> table candidate" -f $s.Name, $t.DataRows, $t.HeaderRowRelative, (ConvertTo-ColumnLetters ($t.FirstCol + 1))))
            Add-TableLines -T $t -ColPrefixMode 'letters'
            $tables.Add(@{ Name = $s.Name; Columns = @($t.Columns | ForEach-Object { @{ Name = $_.Name; IsUnique = $_.Profile.IsUnique } }) })
        }
    }
}
if ($tables.Count -ge 2) {
    $matches = @(Find-KeyMatches -Tables $tables.ToArray())
    $lines.Add('KEY MATCHES (star candidates; same column name, one side unique):')
    if ($matches.Count -eq 0) { $lines.Add('    none - treat tables as independent (flat) unless the user says otherwise') }
    foreach ($m in $matches) { $lines.Add('    ' + $m) }
}
$lines.Add('NOTE header_row is relative to the sheet''s used range (what Power BI sees); for csv it is the physical line number.')
$lines | ForEach-Object { $_ }
exit 0
```

- [ ] **Step 4: Run to verify it passes; eyeball the output**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter profile
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/profile-data.ps1 -DataFolder '\\Mac\Home\Desktop\Claude\powerbi\skills\.claude\skills\powerbi\scripts\tests\fixtures'
```
Expected: `2 tests, 0 failed`; the second command prints the profile (≈ 35 lines) with Korean headers rendered correctly. Adjust regexes in the test only if the *format* was intentionally changed, never to paper over wrong values.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/powerbi/scripts/profile-data.ps1 .claude/skills/powerbi/scripts/tests/profile.test.ps1
git commit -m "feat(scripts): profile-data entry point"
```

---

### Task 8: `lib/spec.ps1` — spec model, field-ref resolution, validation, visual catalog

**Files:**
- Create: `.claude/skills/powerbi/scripts/lib/spec.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/fixtures/spec-flat.json`, `spec-star.json`, `spec-broken.json`
- Create: `.claude/skills/powerbi/scripts/tests/spec.test.ps1`

- [ ] **Step 1: Write the fixture specs**

`fixtures/spec-flat.json` (uses `apps.xlsx`):

```json
{
  "name": "AppsFlat",
  "locale": "ko-KR",
  "autoDateTable": true,
  "tables": [
    {
      "name": "Apps",
      "file": "apps.xlsx",
      "sheet": "Data",
      "headerRow": 1,
      "columns": [
        { "name": "id", "type": "int64", "key": true, "hidden": true, "description": "App ID" },
        { "name": "track_name", "rename": "앱 이름", "type": "text", "description": "Track Name" },
        { "name": "size_bytes", "type": "int64", "format": "#,0" },
        { "name": "price", "type": "decimal", "format": "$#,0.00" },
        { "name": "rating_count_tot", "type": "int64" },
        { "name": "user_rating", "type": "double", "summarizeBy": "average", "format": "0.0" },
        { "name": "prime_genre", "rename": "장르", "type": "text" },
        { "name": "release_date", "rename": "출시일", "type": "date" }
      ]
    }
  ],
  "measures": [
    { "table": "Apps", "name": "앱 수", "dax": "COUNTROWS(Apps)", "format": "#,0" },
    { "table": "Apps", "name": "평균 평점", "dax": "AVERAGE(Apps[user_rating])", "format": "0.00" },
    { "table": "Apps", "name": "평균 가격", "dax": "AVERAGE(Apps[price])", "format": "$#,0.00" }
  ],
  "pages": [
    {
      "name": "개요",
      "visuals": [
        { "type": "textbox", "text": "앱스토어 대시보드", "grid": [0, 0, 12, 1], "fontSize": 24 },
        { "type": "card", "title": "앱 수", "fields": { "Values": ["Apps.앱 수"] }, "grid": [0, 1, 3, 2] },
        { "type": "card", "title": "평균 평점", "fields": { "Values": ["Apps.평균 평점"] }, "grid": [3, 1, 3, 2] },
        { "type": "clusteredBarChart", "title": "장르별 앱 수", "fields": { "Category": ["Apps.장르"], "Y": ["Apps.앱 수"] },
          "sort": { "field": "Apps.앱 수", "direction": "desc" }, "topN": { "n": 3, "by": "Apps.앱 수" }, "grid": [6, 1, 6, 4] },
        { "type": "lineChart", "title": "월별 출시 앱 수", "fields": { "Category": ["Calendar.YearMonth"], "Y": ["Apps.앱 수"] }, "grid": [0, 3, 6, 3] },
        { "type": "scatterChart", "title": "가격 vs 평점", "fields": { "Category": ["Apps.앱 이름"], "X": ["avg:Apps.price"], "Y": ["Apps.평균 평점"], "Size": ["sum:Apps.rating_count_tot"] }, "grid": [6, 5, 6, 3] },
        { "type": "slicer", "fields": { "Values": ["Apps.장르"] }, "grid": [0, 6, 3, 2] },
        { "type": "tableEx", "title": "상위 앱", "fields": { "Values": ["Apps.앱 이름", "Apps.장르", "Apps.평균 평점"] }, "grid": [3, 6, 3, 2] }
      ]
    },
    {
      "name": "상세",
      "filters": [ { "field": "Apps.장르", "in": ["Games", "Education"] } ],
      "visuals": [
        { "type": "pivotTable", "fields": { "Rows": ["Apps.장르"], "Columns": ["Calendar.Year"], "Values": ["Apps.앱 수", "Apps.평균 가격"] }, "grid": [0, 0, 12, 4] },
        { "type": "donutChart", "fields": { "Category": ["Apps.장르"], "Y": ["Apps.앱 수"] }, "grid": [0, 4, 6, 4] },
        { "type": "gauge", "fields": { "Y": ["Apps.평균 평점"], "MaxValue": ["max:Apps.user_rating"] }, "grid": [6, 4, 6, 4] }
      ]
    }
  ]
}
```

`fixtures/spec-star.json` (uses `sales_cp949.csv` + `regions.csv`):

```json
{
  "name": "SalesStar",
  "locale": "ko-KR",
  "autoDateTable": true,
  "tables": [
    { "name": "Sales", "file": "sales_cp949.csv", "encoding": 949, "delimiter": ",",
      "columns": [
        { "name": "주문일", "type": "date" }, { "name": "지역코드", "type": "int64", "hidden": true },
        { "name": "제품", "type": "text" }, { "name": "수량", "type": "int64" }, { "name": "매출액", "type": "decimal", "format": "#,0" } ] },
    { "name": "Regions", "file": "regions.csv", "encoding": 65001,
      "columns": [ { "name": "지역코드", "type": "int64", "key": true, "hidden": true }, { "name": "지역명", "type": "text" }, { "name": "권역", "type": "text" } ] }
  ],
  "relationships": [ { "from": "Sales.지역코드", "to": "Regions.지역코드" } ],
  "measures": [
    { "table": "Sales", "name": "총 매출", "dax": "SUM(Sales[매출액])", "format": "#,0" },
    { "table": "Sales", "name": "총 수량", "dax": "SUM(Sales[수량])", "format": "#,0" }
  ],
  "pages": [ { "name": "매출", "visuals": [
    { "type": "card", "fields": { "Values": ["Sales.총 매출"] }, "grid": [0, 0, 3, 2] },
    { "type": "clusteredColumnChart", "fields": { "Category": ["Regions.권역"], "Y": ["Sales.총 매출"], "Series": ["Sales.제품"] }, "grid": [3, 0, 9, 4] },
    { "type": "areaChart", "fields": { "Category": ["Calendar.Month"], "Y": ["Sales.총 수량"] }, "grid": [0, 4, 12, 4] } ] } ]
}
```

`fixtures/spec-broken.json` — three deliberate errors (unknown column, bad role, bad type):

```json
{
  "name": "Broken",
  "tables": [ { "name": "Apps", "file": "apps.xlsx", "sheet": "Data",
    "columns": [ { "name": "id", "type": "int64" }, { "name": "price", "type": "money" } ] } ],
  "measures": [ { "table": "Apps", "name": "앱 수", "dax": "COUNTROWS(Apps)" } ],
  "pages": [ { "name": "p", "visuals": [
    { "type": "card", "fields": { "Values": ["Apps.nope"] }, "grid": [0, 0, 3, 2] },
    { "type": "lineChart", "fields": { "Category": ["Apps.id"], "Y": ["Apps.앱 수"], "Bogus": ["Apps.id"] }, "grid": [0, 2, 6, 3] } ] } ]
}
```

- [ ] **Step 2: Write the failing test**

```powershell
. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'spec.ps1')

Test-Case 'Get-SpecModel builds tables with final column names, measures, and a Calendar table' {
    $spec = Read-Spec -Path (Join-Path $fixtures 'spec-flat.json')
    $m = Get-SpecModel -Spec $spec
    Assert-Equal 'AppsFlat' $m.Name
    Assert-True ($m.Tables.Contains('Apps')) 'Apps table'
    Assert-Equal '앱 이름' $m.Tables['Apps'].Columns[1].Name
    Assert-Equal 'track_name' $m.Tables['Apps'].Columns[1].Source
    Assert-Equal 'none' $m.Tables['Apps'].Columns[0].SummarizeBy
    Assert-Equal 'sum' $m.Tables['Apps'].Columns[2].SummarizeBy
    Assert-Equal 3 $m.Tables['Apps'].Measures.Count
    Assert-True ($m.Tables.Contains('Calendar')) 'Calendar generated'
    Assert-Equal 1 $m.Relationships.Count
    Assert-Equal 'Apps' $m.Relationships[0].FromTable; Assert-Equal '출시일' $m.Relationships[0].FromColumn
    Assert-Equal 'Calendar' $m.Relationships[0].ToTable; Assert-Equal 'Date' $m.Relationships[0].ToColumn
}
Test-Case 'Resolve-FieldRef distinguishes measure, column, aggregation' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $a = Resolve-FieldRef -Model $m -Ref 'Apps.앱 수';         Assert-Equal 'Measure' $a.Kind
    $b = Resolve-FieldRef -Model $m -Ref 'Apps.장르';          Assert-Equal 'Column' $b.Kind; Assert-Equal 'Apps' $b.Table
    $c = Resolve-FieldRef -Model $m -Ref 'sum:Apps.rating_count_tot'; Assert-Equal 'Aggregation' $c.Kind; Assert-Equal 0 $c.Function
    $d = Resolve-FieldRef -Model $m -Ref 'countDistinct:Apps.장르'; Assert-Equal 2 $d.Function
    $e = Resolve-FieldRef -Model $m -Ref 'Calendar.YearMonth'; Assert-Equal 'Column' $e.Kind
    Assert-Throws { Resolve-FieldRef -Model $m -Ref 'Apps.nope' } 'unknown field'
    Assert-Throws { Resolve-FieldRef -Model $m -Ref 'sum:Apps.앱 수' } 'aggregation prefix on a measure'
    Assert-Throws { Resolve-FieldRef -Model $m -Ref 'noDot' } 'Table.Name'
}
Test-Case 'Get-SpecValidation passes the good fixtures' {
    foreach ($f in 'spec-flat.json', 'spec-star.json') {
        $v = Get-SpecValidation -Spec (Read-Spec -Path (Join-Path $fixtures $f)) -RawDataDir $fixtures
        Assert-Equal 0 $v.Errors.Count ("errors in " + $f + ": " + ($v.Errors -join ' | '))
    }
}
Test-Case 'Get-SpecValidation reports the three deliberate errors with spec paths' {
    $v = Get-SpecValidation -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-broken.json')) -RawDataDir $fixtures
    Assert-Equal 3 $v.Errors.Count ($v.Errors -join ' | ')
    Assert-Match "tables\[0\]\.columns\[1\]\.type.*money" ($v.Errors -join "`n")
    Assert-Match "pages\[0\]\.visuals\[0\]\.fields\.Values\[0\].*unknown field" ($v.Errors -join "`n")
    Assert-Match "pages\[0\]\.visuals\[1\]\.fields\.Bogus.*not a role of lineChart" ($v.Errors -join "`n")
}
Test-Case 'Get-SpecValidation catches missing file, bad grid, duplicate names' {
    $spec = Read-Spec -Path (Join-Path $fixtures 'spec-flat.json')
    $spec.tables[0].file = 'missing.xlsx'
    $spec.pages[0].visuals[1].grid = @(10, 0, 4, 2)
    $spec.measures[1].name = '앱 수'
    $v = Get-SpecValidation -Spec $spec -RawDataDir $fixtures
    $all = $v.Errors -join "`n"
    Assert-Match 'tables\[0\]\.file.*missing\.xlsx' $all
    Assert-Match 'pages\[0\]\.visuals\[1\]\.grid.*exceeds' $all
    Assert-Match 'measures\[1\]\.name.*duplicate' $all
}
Test-Case 'Get-GridPosition converts the 12x8 grid into pixels' {
    $p = Get-GridPosition -Grid @(0, 0, 3, 2)
    Assert-Equal 40 $p.x; Assert-Equal 40 $p.y; Assert-Equal 425 $p.width; Assert-Equal 235 $p.height
    $q = Get-GridPosition -Grid @(6, 1, 6, 4)
    Assert-Equal 970 $q.x; Assert-Equal 168 $q.y; Assert-Equal 910 $q.width; Assert-Equal 490 $q.height
}
```

- [ ] **Step 3: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter spec
```
Expected: fails loading `spec.ps1`.

- [ ] **Step 4: Write `lib/spec.ps1`**

```powershell
# spec.ps1 — report-spec.json loading, model building, field-ref resolution, validation, visual catalog, grid.
$ErrorActionPreference = 'Stop'

$script:AllowedTypes = @('int64', 'double', 'decimal', 'text', 'date', 'datetime', 'boolean')
$script:AllowedSummarize = @('none', 'sum', 'average', 'count', 'min', 'max', 'distinctCount')
$script:AggFunctions = [ordered]@{ sum = 0; avg = 1; countDistinct = 2; min = 3; max = 4; count = 5 }

# Spec role names → PBIR queryState keys. Required/Optional/Max are in spec-role terms.
$script:VisualCatalog = [ordered]@{
    card                 = @{ PbirType = 'cardVisual';           Required = @('Values');            Optional = @();                                  Max = @{ Values = 1 }; RoleMap = @{ Values = 'Data' } }
    multiRowCard         = @{ PbirType = 'multiRowCard';         Required = @('Values');            Optional = @();                                  Max = @{} ; RoleMap = @{} }
    gauge                = @{ PbirType = 'gauge';                Required = @('Y');                 Optional = @('MinValue', 'MaxValue', 'TargetValue'); Max = @{ Y = 1; MinValue = 1; MaxValue = 1; TargetValue = 1 }; RoleMap = @{} }
    lineChart            = @{ PbirType = 'lineChart';            Required = @('Category', 'Y');     Optional = @('Series');                          Max = @{ Category = 1; Series = 1 }; RoleMap = @{} }
    areaChart            = @{ PbirType = 'areaChart';            Required = @('Category', 'Y');     Optional = @('Series');                          Max = @{ Category = 1; Series = 1 }; RoleMap = @{} }
    clusteredColumnChart = @{ PbirType = 'clusteredColumnChart'; Required = @('Category', 'Y');     Optional = @('Series');                          Max = @{ Category = 1; Series = 1 }; RoleMap = @{} }
    stackedColumnChart   = @{ PbirType = 'columnChart';          Required = @('Category', 'Y');     Optional = @('Series');                          Max = @{ Category = 1; Series = 1 }; RoleMap = @{} }
    clusteredBarChart    = @{ PbirType = 'clusteredBarChart';    Required = @('Category', 'Y');     Optional = @('Series');                          Max = @{ Category = 1; Series = 1 }; RoleMap = @{} }
    stackedBarChart      = @{ PbirType = 'barChart';             Required = @('Category', 'Y');     Optional = @('Series');                          Max = @{ Category = 1; Series = 1 }; RoleMap = @{} }
    pieChart             = @{ PbirType = 'pieChart';             Required = @('Category', 'Y');     Optional = @();                                  Max = @{ Category = 1; Y = 1 }; RoleMap = @{} }
    donutChart           = @{ PbirType = 'donutChart';           Required = @('Category', 'Y');     Optional = @();                                  Max = @{ Category = 1; Y = 1 }; RoleMap = @{} }
    scatterChart         = @{ PbirType = 'scatterChart';         Required = @('Category', 'X', 'Y'); Optional = @('Size');                           Max = @{ Category = 1; X = 1; Y = 1; Size = 1 }; RoleMap = @{} }
    tableEx              = @{ PbirType = 'tableEx';              Required = @('Values');            Optional = @();                                  Max = @{}; RoleMap = @{} }
    pivotTable           = @{ PbirType = 'pivotTable';           Required = @('Rows', 'Values');    Optional = @('Columns');                         Max = @{}; RoleMap = @{} }
    slicer               = @{ PbirType = 'slicer';               Required = @('Values');            Optional = @();                                  Max = @{ Values = 1 }; RoleMap = @{} }
    textbox              = @{ PbirType = 'textbox';              Required = @();                    Optional = @();                                  Max = @{}; RoleMap = @{}; IsText = $true }
}

function Read-Spec {
    param([Parameter(Mandatory)][string]$Path)
    $raw = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    return ($raw | ConvertFrom-Json)
}

function Get-DefaultSummarize { param([string]$Type, [bool]$Key) if ($Key) { return 'none' }; if ($Type -in 'int64', 'double', 'decimal') { return 'sum' }; return 'none' }

function Get-SpecModel {
    # Builds the normalized model. Tolerant: invalid entries are skipped here and reported by Get-SpecValidation.
    param([Parameter(Mandatory)]$Spec)
    $tables = [ordered]@{}
    foreach ($t in (ConvertTo-Array (Get-Prop $Spec 'tables'))) {
        $name = [string](Get-Prop $t 'name' '')
        if (-not $name) { continue }
        $file = [string](Get-Prop $t 'file' '')
        $cols = New-Object System.Collections.Generic.List[object]
        foreach ($c in (ConvertTo-Array (Get-Prop $t 'columns'))) {
            $src = [string](Get-Prop $c 'name' ''); if (-not $src) { continue }
            $final = [string](Get-Prop $c 'rename' $src); if (-not $final) { $final = $src }
            $type = [string](Get-Prop $c 'type' 'text')
            $key = [bool](Get-Prop $c 'key' $false)
            $cols.Add(@{ Source = $src; Name = $final; Type = $type; Key = $key; Hidden = [bool](Get-Prop $c 'hidden' $false)
                         Format = (Get-Prop $c 'format' $null); SummarizeBy = [string](Get-Prop $c 'summarizeBy' (Get-DefaultSummarize $type $key))
                         Description = (Get-Prop $c 'description' $null); SortBy = (Get-Prop $c 'sortBy' $null) })
        }
        $tables[$name] = @{ Name = $name; File = $file; Sheet = (Get-Prop $t 'sheet' $null); HeaderRow = [int](Get-Prop $t 'headerRow' 1)
                            Encoding = [int](Get-Prop $t 'encoding' 65001); Delimiter = [string](Get-Prop $t 'delimiter' ',')
                            IsXlsx = ($file -match '(?i)\.xls[xm]$'); Columns = $cols.ToArray(); Measures = @(); IsCalculated = $false }
    }
    foreach ($m in (ConvertTo-Array (Get-Prop $Spec 'measures'))) {
        $tn = [string](Get-Prop $m 'table' '')
        if (-not $tables.Contains($tn)) { continue }
        $tables[$tn].Measures = @($tables[$tn].Measures) + @(@{ Name = [string](Get-Prop $m 'name' ''); Dax = [string](Get-Prop $m 'dax' ''); Format = (Get-Prop $m 'format' $null); Description = (Get-Prop $m 'description' $null) })
    }
    $rels = New-Object System.Collections.Generic.List[object]
    foreach ($r in (ConvertTo-Array (Get-Prop $Spec 'relationships'))) {
        $from = [string](Get-Prop $r 'from' ''); $to = [string](Get-Prop $r 'to' '')
        $fi = $from.IndexOf('.'); $ti = $to.IndexOf('.')
        if ($fi -lt 1 -or $ti -lt 1) { continue }
        $rels.Add(@{ FromTable = $from.Substring(0, $fi); FromColumn = $from.Substring($fi + 1); ToTable = $to.Substring(0, $ti); ToColumn = $to.Substring($ti + 1)
                     Active = [bool](Get-Prop $r 'active' $true); CrossFilter = [string](Get-Prop $r 'crossFilter' 'single') })
    }
    # Calendar table
    $calendar = $null
    if ([bool](Get-Prop $Spec 'autoDateTable' $false)) {
        $dateDimColumns = @()
        foreach ($r in $rels) { if ($tables.Contains($r.ToTable)) { foreach ($c in $tables[$r.ToTable].Columns) { if ($c.Name -eq $r.ToColumn -and $c.Type -in 'date', 'datetime') { $dateDimColumns += ($r.FromTable + '.' + $r.FromColumn) } } } }
        $refs = New-Object System.Collections.Generic.List[object]
        foreach ($tn in $tables.Keys) {
            $first = $true
            foreach ($c in $tables[$tn].Columns) {
                if ($c.Type -in 'date', 'datetime' -and ($dateDimColumns -notcontains ($tn + '.' + $c.Name))) {
                    $refs.Add(@{ Table = $tn; Column = $c.Name; Active = $first }); $first = $false
                }
            }
        }
        if ($refs.Count -gt 0) {
            $calendar = @{ DateRefs = $refs.ToArray() }
            $calCols = @(
                @{ Source = 'Date';      Name = 'Date';      Type = 'date';  Key = $true;  Hidden = $false; Format = 'yyyy-MM-dd'; SummarizeBy = 'none'; Description = $null; SortBy = $null },
                @{ Source = 'Year';      Name = 'Year';      Type = 'int64'; Key = $false; Hidden = $false; Format = '0';          SummarizeBy = 'none'; Description = $null; SortBy = $null },
                @{ Source = 'Quarter';   Name = 'Quarter';   Type = 'text';  Key = $false; Hidden = $false; Format = $null;        SummarizeBy = 'none'; Description = $null; SortBy = $null },
                @{ Source = 'MonthNo';   Name = 'MonthNo';   Type = 'int64'; Key = $false; Hidden = $true;  Format = '0';          SummarizeBy = 'none'; Description = $null; SortBy = $null },
                @{ Source = 'Month';     Name = 'Month';     Type = 'text';  Key = $false; Hidden = $false; Format = $null;        SummarizeBy = 'none'; Description = $null; SortBy = 'MonthNo' },
                @{ Source = 'YearMonth'; Name = 'YearMonth'; Type = 'text';  Key = $false; Hidden = $false; Format = $null;        SummarizeBy = 'none'; Description = $null; SortBy = $null }
            )
            $tables['Calendar'] = @{ Name = 'Calendar'; File = $null; Sheet = $null; HeaderRow = 1; Encoding = 65001; Delimiter = ','; IsXlsx = $false; Columns = $calCols; Measures = @(); IsCalculated = $true }
            foreach ($ref in $refs) { $rels.Add(@{ FromTable = $ref.Table; FromColumn = $ref.Column; ToTable = 'Calendar'; ToColumn = 'Date'; Active = $ref.Active; CrossFilter = 'single' }) }
        }
    }
    return @{ Name = [string](Get-Prop $Spec 'name' ''); Locale = [string](Get-Prop $Spec 'locale' 'en-US'); Tables = $tables; Relationships = $rels.ToArray(); Calendar = $calendar; Pages = (ConvertTo-Array (Get-Prop $Spec 'pages')) }
}

function Resolve-FieldRef {
    # 'Table.Name' | 'agg:Table.Column' → @{ Kind='Column'|'Measure'|'Aggregation'; Table; Name; Function; FunctionName; Type }
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Ref)
    $agg = $null; $body = $Ref
    if ($Ref -match '^([A-Za-z]+):(.+)$') { $agg = $Matches[1]; $body = $Matches[2] }
    if ($null -ne $agg -and -not $script:AggFunctions.Contains($agg)) { throw "unknown aggregation prefix '$agg:' in '$Ref' (allowed: $($script:AggFunctions.Keys -join ', '))" }
    $i = $body.IndexOf('.')
    if ($i -lt 1 -or $i -eq $body.Length - 1) { throw "field '$Ref' must be Table.Name" }
    $tn = $body.Substring(0, $i); $fn = $body.Substring($i + 1)
    if (-not $Model.Tables.Contains($tn)) { throw "unknown table '$tn' in '$Ref'" }
    $t = $Model.Tables[$tn]
    foreach ($m in $t.Measures) { if ($m.Name -eq $fn) {
        if ($null -ne $agg) { throw "aggregation prefix on a measure is not allowed: '$Ref'" }
        return @{ Kind = 'Measure'; Table = $tn; Name = $fn; Function = $null; FunctionName = $null; Type = $null } } }
    foreach ($c in $t.Columns) { if ($c.Name -eq $fn) {
        if ($null -ne $agg) { return @{ Kind = 'Aggregation'; Table = $tn; Name = $fn; Function = $script:AggFunctions[$agg]; FunctionName = $agg; Type = $c.Type } }
        return @{ Kind = 'Column'; Table = $tn; Name = $fn; Function = $null; FunctionName = $null; Type = $c.Type } } }
    throw "unknown field '$fn' in table '$tn' ('$Ref'); known: $((@($t.Columns | ForEach-Object { $_.Name }) + @($t.Measures | ForEach-Object { $_.Name })) -join ', ')"
}

function Get-GridPosition {
    # [col,row,w,h] on 12x8 → pixels (1920x1080, 40 margin, 20 gutter; cell 135 x 107.5)
    param([Parameter(Mandatory)]$Grid)
    $col = [int]$Grid[0]; $row = [int]$Grid[1]; $w = [int]$Grid[2]; $h = [int]$Grid[3]
    return [ordered]@{ x = 40 + $col * 155; y = [int][Math]::Round(40 + $row * 127.5); width = $w * 135 + ($w - 1) * 20; height = [int][Math]::Round($h * 107.5 + ($h - 1) * 20) }
}

function Get-SpecValidation {
    param([Parameter(Mandatory)]$Spec, [Parameter(Mandatory)][string]$RawDataDir)
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    $name = [string](Get-Prop $Spec 'name' '')
    if ($name -notmatch '^[A-Za-z0-9_-]+$') { $errors.Add("name: must match ^[A-Za-z0-9_-]+$ (got '$name')") }
    $tablesRaw = ConvertTo-Array (Get-Prop $Spec 'tables')
    if ($tablesRaw.Count -eq 0) { $errors.Add('tables: at least one table is required') }
    $seenTables = @{}
    for ($ti = 0; $ti -lt $tablesRaw.Count; $ti++) {
        $t = $tablesRaw[$ti]; $p = "tables[$ti]"
        $tn = [string](Get-Prop $t 'name' '')
        if (-not $tn) { $errors.Add("$p.name: required"); continue }
        if ($tn -eq 'Calendar') { $errors.Add("$p.name: 'Calendar' is reserved for the auto date table") }
        if ($tn.Contains('.')) { $errors.Add("$p.name: table names may not contain '.'") }
        if ($seenTables.ContainsKey($tn)) { $errors.Add("$p.name: duplicate table name '$tn'") }; $seenTables[$tn] = $true
        $file = [string](Get-Prop $t 'file' '')
        if (-not $file) { $errors.Add("$p.file: required") }
        elseif (-not (Test-Path -LiteralPath (Join-Path $RawDataDir $file))) { $errors.Add("$p.file: '$file' not found in $RawDataDir") }
        elseif ($file -match '(?i)\.xls[xm]$' -and -not (Get-Prop $t 'sheet' $null)) { $errors.Add("$p.sheet: required for Excel files") }
        $cols = ConvertTo-Array (Get-Prop $t 'columns')
        if ($cols.Count -eq 0) { $errors.Add("$p.columns: at least one column is required") }
        $seenCols = @{}
        for ($ci = 0; $ci -lt $cols.Count; $ci++) {
            $c = $cols[$ci]; $cp = "$p.columns[$ci]"
            $src = [string](Get-Prop $c 'name' ''); if (-not $src) { $errors.Add("$cp.name: required"); continue }
            $final = [string](Get-Prop $c 'rename' $src)
            if ($seenCols.ContainsKey($final)) { $errors.Add("$cp: duplicate column name '$final'") }; $seenCols[$final] = $true
            $type = [string](Get-Prop $c 'type' 'text')
            if ($script:AllowedTypes -notcontains $type) { $errors.Add("$cp.type: '$type' is not one of $($script:AllowedTypes -join '|')") }
            $sb = Get-Prop $c 'summarizeBy' $null
            if ($null -ne $sb -and $script:AllowedSummarize -notcontains [string]$sb) { $errors.Add("$cp.summarizeBy: '$sb' is not one of $($script:AllowedSummarize -join '|')") }
        }
    }
    $model = Get-SpecModel -Spec $Spec
    $measuresRaw = ConvertTo-Array (Get-Prop $Spec 'measures')
    $seenMeasures = @{}
    for ($mi = 0; $mi -lt $measuresRaw.Count; $mi++) {
        $m = $measuresRaw[$mi]; $p = "measures[$mi]"
        $tn = [string](Get-Prop $m 'table' ''); $mn = [string](Get-Prop $m 'name' '')
        if (-not $model.Tables.Contains($tn)) { $errors.Add("$p.table: unknown table '$tn'"); continue }
        if (-not $mn) { $errors.Add("$p.name: required"); continue }
        if (-not [string](Get-Prop $m 'dax' '')) { $errors.Add("$p.dax: required") }
        $k = "$tn.$mn"
        if ($seenMeasures.ContainsKey($k)) { $errors.Add("$p.name: duplicate measure '$mn' in table '$tn'") }; $seenMeasures[$k] = $true
        foreach ($c in $model.Tables[$tn].Columns) { if ($c.Name -eq $mn) { $errors.Add("$p.name: '$mn' collides with a column of '$tn'") } }
    }
    $relsRaw = ConvertTo-Array (Get-Prop $Spec 'relationships')
    for ($ri = 0; $ri -lt $relsRaw.Count; $ri++) {
        $p = "relationships[$ri]"
        foreach ($side in 'from', 'to') {
            $ref = [string](Get-Prop $relsRaw[$ri] $side '')
            try { $r = Resolve-FieldRef -Model $model -Ref $ref; if ($r.Kind -ne 'Column') { $errors.Add("$p.$side: must reference a column") } }
            catch { $errors.Add("$p.$side: " + $_.Exception.Message) }
        }
    }
    $pagesRaw = ConvertTo-Array (Get-Prop $Spec 'pages')
    if ($pagesRaw.Count -eq 0) { $errors.Add('pages: at least one page is required') }
    for ($pi = 0; $pi -lt $pagesRaw.Count; $pi++) {
        $page = $pagesRaw[$pi]; $pp = "pages[$pi]"
        if (-not [string](Get-Prop $page 'name' '')) { $errors.Add("$pp.name: required") }
        $pf = ConvertTo-Array (Get-Prop $page 'filters')
        for ($fi = 0; $fi -lt $pf.Count; $fi++) {
            try { $r = Resolve-FieldRef -Model $model -Ref ([string](Get-Prop $pf[$fi] 'field' '')); if ($r.Kind -ne 'Column') { $errors.Add("$pp.filters[$fi].field: must be a column") } } catch { $errors.Add("$pp.filters[$fi].field: " + $_.Exception.Message) }
            if ((ConvertTo-Array (Get-Prop $pf[$fi] 'in')).Count -eq 0) { $errors.Add("$pp.filters[$fi].in: non-empty list required") }
        }
        $visuals = ConvertTo-Array (Get-Prop $page 'visuals')
        $occupied = New-Object System.Collections.Generic.List[object]
        for ($vi = 0; $vi -lt $visuals.Count; $vi++) {
            $v = $visuals[$vi]; $vp = "$pp.visuals[$vi]"
            $type = [string](Get-Prop $v 'type' '')
            if (-not $script:VisualCatalog.Contains($type)) { $errors.Add("$vp.type: unknown visual type '$type' (allowed: $($script:VisualCatalog.Keys -join ', '))"); continue }
            $cat = $script:VisualCatalog[$type]
            $grid = ConvertTo-Array (Get-Prop $v 'grid')
            if ($grid.Count -ne 4) { $errors.Add("$vp.grid: must be [col,row,w,h]") }
            else {
                $g = @($grid | ForEach-Object { [int]$_ })
                if ($g[0] -lt 0 -or $g[1] -lt 0 -or $g[2] -lt 1 -or $g[3] -lt 1 -or ($g[0] + $g[2]) -gt 12 -or ($g[1] + $g[3]) -gt 8) { $errors.Add("$vp.grid: [$($g -join ',')] exceeds the 12x8 grid") }
                else {
                    foreach ($o in $occupied) { if ($g[0] -lt $o[0] + $o[2] -and $o[0] -lt $g[0] + $g[2] -and $g[1] -lt $o[1] + $o[3] -and $o[1] -lt $g[1] + $g[3]) { $warnings.Add("$vp.grid overlaps an earlier visual on this page"); break } }
                    $occupied.Add($g)
                }
            }
            if ($cat.IsText) { if (-not [string](Get-Prop $v 'text' '')) { $errors.Add("$vp.text: required for textbox") }; continue }
            $fields = Get-Prop $v 'fields' $null
            $roleNames = @(); if ($null -ne $fields) { $roleNames = @($fields.PSObject.Properties | ForEach-Object { $_.Name }) }
            foreach ($req in $cat.Required) { if ($roleNames -notcontains $req -or (ConvertTo-Array (Get-Prop $fields $req)).Count -eq 0) { $errors.Add("$vp.fields.$req: required for $type") } }
            foreach ($role in $roleNames) {
                if (($cat.Required + $cat.Optional) -notcontains $role) { $errors.Add("$vp.fields.$role: not a role of $type (roles: $(($cat.Required + $cat.Optional) -join ', '))"); continue }
                $refs = ConvertTo-Array (Get-Prop $fields $role)
                if ($cat.Max.ContainsKey($role) -and $refs.Count -gt $cat.Max[$role]) { $errors.Add("$vp.fields.$role: at most $($cat.Max[$role]) field(s)") }
                for ($fi = 0; $fi -lt $refs.Count; $fi++) { try { [void](Resolve-FieldRef -Model $model -Ref ([string]$refs[$fi])) } catch { $errors.Add("$vp.fields.$role[$fi]: " + $_.Exception.Message) } }
            }
            $sort = Get-Prop $v 'sort' $null
            if ($null -ne $sort) { try { [void](Resolve-FieldRef -Model $model -Ref ([string](Get-Prop $sort 'field' ''))) } catch { $errors.Add("$vp.sort.field: " + $_.Exception.Message) }
                $dir = [string](Get-Prop $sort 'direction' 'desc'); if ($dir -notin 'asc', 'desc') { $errors.Add("$vp.sort.direction: 'asc' or 'desc'") } }
            $topN = Get-Prop $v 'topN' $null
            if ($null -ne $topN) {
                if ([int](Get-Prop $topN 'n' 0) -lt 1) { $errors.Add("$vp.topN.n: must be >= 1") }
                try { $r = Resolve-FieldRef -Model $model -Ref ([string](Get-Prop $topN 'by' '')); if ($r.Kind -eq 'Column') { $errors.Add("$vp.topN.by: must be a measure or an aggregated column (e.g. sum:Table.Col)") } } catch { $errors.Add("$vp.topN.by: " + $_.Exception.Message) }
                $catRefs = ConvertTo-Array (Get-Prop $fields 'Category'); $rowRefs = ConvertTo-Array (Get-Prop $fields 'Rows')
                if ($catRefs.Count -eq 0 -and $rowRefs.Count -eq 0) { $errors.Add("$vp.topN: needs a Category (or Rows) field to rank") }
            }
            $vf = ConvertTo-Array (Get-Prop $v 'filters')
            for ($fi = 0; $fi -lt $vf.Count; $fi++) {
                try { $r = Resolve-FieldRef -Model $model -Ref ([string](Get-Prop $vf[$fi] 'field' '')); if ($r.Kind -ne 'Column') { $errors.Add("$vp.filters[$fi].field: must be a column") } } catch { $errors.Add("$vp.filters[$fi].field: " + $_.Exception.Message) }
                if ((ConvertTo-Array (Get-Prop $vf[$fi] 'in')).Count -eq 0) { $errors.Add("$vp.filters[$fi].in: non-empty list required") }
            }
        }
    }
    return @{ Errors = $errors.ToArray(); Warnings = $warnings.ToArray(); Model = $model }
}
```

- [ ] **Step 5: Run to verify it passes**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter spec
```
Expected: `6 tests, 0 failed`.

- [ ] **Step 6: Commit**

```bash
git add .claude/skills/powerbi/scripts/lib/spec.ps1 .claude/skills/powerbi/scripts/tests/spec.test.ps1 .claude/skills/powerbi/scripts/tests/fixtures/spec-*.json
git commit -m "feat(scripts): spec model, field-ref resolution, validation, visual catalog"
```

---

### Task 9: `lib/tmdl.ps1` — semantic model emitter

**Files:**
- Create: `.claude/skills/powerbi/scripts/lib/tmdl.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/tmdl.test.ps1`

- [ ] **Step 1: Write the failing test**

```powershell
. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'spec.ps1')
. (Join-Path $libDir 'tmdl.ps1')

Test-Case 'Format-TmdlName quotes anything that is not a plain identifier' {
    Assert-Equal 'Apps' (Format-TmdlName 'Apps')
    Assert-Equal "'앱 이름'" (Format-TmdlName '앱 이름')
    Assert-Equal "'It''s'" (Format-TmdlName "It's")
}
Test-Case 'New-MQuery for xlsx: sheet navigation, select, rename, types' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $q = (New-MQuery -Table $m.Tables['Apps'] -AbsPath 'C:\data\App Data.xlsx') -join "`n"
    Assert-Match 'Source = Excel\.Workbook\(File\.Contents\("C:\\data\\App Data\.xlsx"\), null, true\),' $q
    Assert-Match 'Sheet = Source\{\[Item="Data",Kind="Sheet"\]\}\[Data\],' $q
    Assert-Match 'Promoted = Table\.PromoteHeaders\(Sheet, \[PromoteAllScalars=true\]\),' $q
    Assert-Match 'Selected = Table\.SelectColumns\(Promoted, \{"id", "track_name", "size_bytes"' $q
    Assert-Match 'Renamed = Table\.RenameColumns\(Selected, \{\{"track_name", "앱 이름"\}, \{"prime_genre", "장르"\}, \{"release_date", "출시일"\}\}\),' $q
    Assert-Match 'Typed = Table\.TransformColumnTypes\(Renamed, \{\{"id", Int64\.Type\}, \{"앱 이름", type text\}, \{"size_bytes", Int64\.Type\}, \{"price", Currency\.Type\}, \{"rating_count_tot", Int64\.Type\}, \{"user_rating", type number\}, \{"장르", type text\}, \{"출시일", type date\}\}\)' $q
    Assert-True (-not $q.Contains('Skipped')) 'headerRow 1 → no skip'
}
Test-Case 'New-MQuery for csv: encoding, delimiter, header skip, no rename step' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-star.json'))
    $t = $m.Tables['Sales']; $t.HeaderRow = 3
    $q = (New-MQuery -Table $t -AbsPath '\\srv\share\sales.csv') -join "`n"
    Assert-Match 'Csv\.Document\(File\.Contents\("\\\\srv\\share\\sales\.csv"\),\[Delimiter=",", Encoding=949, QuoteStyle=QuoteStyle\.Csv\]\),' $q
    Assert-Match 'Skipped = Table\.Skip\(Source, 2\),' $q
    Assert-Match 'Promoted = Table\.PromoteHeaders\(Skipped' $q
    Assert-Match 'Typed = Table\.TransformColumnTypes\(Selected, ' $q
    Assert-True (-not $q.Contains('RenameColumns')) 'no renames'
}
Test-Case 'New-TableTmdl emits columns, measures, partition' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $txt = New-TableTmdl -Table $m.Tables['Apps'] -Model $m -AbsPath 'C:\data\apps.xlsx'
    Assert-Match "(?m)^table Apps\r?\n\tlineageTag: [0-9a-f-]{36}" $txt
    Assert-Match "(?m)^\t/// App ID\r?\n\tcolumn id\r?\n\t\tdataType: int64\r?\n\t\tisHidden\r?\n\t\tisKey\r?\n\t\tlineageTag: .*\r?\n\t\tsummarizeBy: none\r?\n\t\tsourceColumn: id" $txt
    Assert-Match "(?m)^\tcolumn '출시일'\r?\n\t\tdataType: dateTime\r?\n\t\tformatString: yyyy-MM-dd\r?\n(.*\r?\n)+?\t\tannotation UnderlyingDateTimeDataType = Date" $txt
    Assert-Match "(?m)^\tcolumn user_rating\r?\n\t\tdataType: double\r?\n\t\tformatString: 0\.0\r?\n\t\tlineageTag: .*\r?\n\t\tsummarizeBy: average" $txt
    Assert-Match "(?m)^\tmeasure '앱 수' = COUNTROWS\(Apps\)\r?\n\t\tformatString: #,0\r?\n\t\tlineageTag: " $txt
    Assert-Match "formatString: \\\`$#,0\.00" $txt
    Assert-Match "(?m)^\tpartition Apps = m\r?\n\t\tmode: import\r?\n\t\tsource =\r?\n\t\t\t\tlet\r?\n\t\t\t\t    Source = Excel" $txt
}
Test-Case 'multi-line DAX measures are indented under the measure line' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $m.Tables['Apps'].Measures = @(@{ Name = 'X'; Dax = "VAR a = 1`nRETURN a"; Format = $null; Description = 'two lines' })
    $txt = New-TableTmdl -Table $m.Tables['Apps'] -Model $m -AbsPath 'C:\x.xlsx'
    Assert-Match "(?m)^\t/// two lines\r?\n\tmeasure X =\r?\n\t\t\tVAR a = 1\r?\n\t\t\tRETURN a\r?\n\t\tlineageTag: " $txt
}
Test-Case 'Calendar table is a calculated table marked as date table' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-star.json'))
    $txt = New-TableTmdl -Table $m.Tables['Calendar'] -Model $m -AbsPath $null
    Assert-Match "(?m)^table Calendar\r?\n\tlineageTag: .*\r?\n\tdataCategory: Time" $txt
    Assert-Match "(?m)^\tcolumn Date\r?\n\t\tdataType: dateTime\r?\n\t\tisKey\r?\n\t\tformatString: yyyy-MM-dd\r?\n\t\tlineageTag: .*\r?\n\t\tsummarizeBy: none\r?\n\t\tisNameInferred\r?\n\t\tisDataTypeInferred\r?\n\t\tsourceColumn: \[Date\]" $txt
    Assert-Match "(?m)^\tcolumn Month\r?\n(.*\r?\n)+?\t\tsortByColumn: MonthNo" $txt
    Assert-Match "(?m)^\tpartition Calendar = calculated\r?\n\t\tmode: import\r?\n\t\tsource = ADDCOLUMNS \( CALENDAR \( DATE \( YEAR \( MIN \( 'Sales'\[주문일\] \) \), 1, 1 \), DATE \( YEAR \( MAX \( 'Sales'\[주문일\] \) \), 12, 31 \) \), ""Year"", YEAR \( \[Date\] \), ""Quarter"", ""Q"" & FORMAT \( \[Date\], ""q"" \), ""MonthNo"", MONTH \( \[Date\] \), ""Month"", FORMAT \( \[Date\], ""MMM"" \), ""YearMonth"", FORMAT \( \[Date\], ""yyyy-MM"" \) \)" $txt
}
Test-Case 'New-ModelTmdl and New-RelationshipsTmdl' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-star.json'))
    $mt = New-ModelTmdl -Model $m
    Assert-Match "(?m)^model Model\r?\n\tculture: ko-KR\r?\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\r?\n\tsourceQueryCulture: ko-KR\r?\n\tdataAccessOptions\r?\n\t\tlegacyRedirects\r?\n\t\treturnErrorValuesAsNull" $mt
    Assert-Match 'annotation __PBI_TimeIntelligenceEnabled = 0' $mt
    Assert-Match 'annotation PBI_QueryOrder = \["Sales","Regions"\]' $mt
    Assert-Match "(?m)^ref table Sales\r?\n^ref table Regions\r?\n^ref table Calendar" $mt
    $rt = New-RelationshipsTmdl -Model $m
    Assert-Match "(?m)^relationship Sales_지역코드_Regions\r?\n\tfromColumn: Sales\.'지역코드'\r?\n\ttoColumn: Regions\.'지역코드'\r?\n" $rt
    Assert-Match "(?m)^relationship Sales_주문일_Calendar\r?\n\tfromColumn: Sales\.'주문일'\r?\n\ttoColumn: Calendar\.Date" $rt
    $m.Relationships[0].Active = $false; $m.Relationships[0].CrossFilter = 'both'
    $rt2 = New-RelationshipsTmdl -Model $m
    Assert-Match "(?m)^relationship Sales_지역코드_Regions\r?\n\tisActive: false\r?\n\tcrossFilteringBehavior: bothDirections\r?\n\tfromColumn" $rt2
}
Test-Case 'Write-SemanticModel writes the full folder' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $dir = Join-Path $tmpRoot 'sm'
    $r = Write-SemanticModel -Model $m -Dir $dir -RawDataDir $fixtures
    foreach ($f in 'definition.pbism', 'definition/database.tmdl', 'definition/model.tmdl', 'definition/relationships.tmdl', 'definition/tables/Apps.tmdl', 'definition/tables/Calendar.tmdl') { Assert-True (Test-Path (Join-Path $dir $f)) "missing $f" }
    $apps = [IO.File]::ReadAllText((Join-Path $dir 'definition/tables/Apps.tmdl'))
    Assert-Match 'File\.Contents\("(\\\\|[A-Za-z]:\\)[^"]*apps\.xlsx"\)' $apps
    Assert-Equal 0 $r.Warnings.Count ($r.Warnings -join ' | ')
}
```

- [ ] **Step 2: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter tmdl
```
Expected: fails loading `tmdl.ps1`.

- [ ] **Step 3: Write `lib/tmdl.ps1`**

```powershell
# tmdl.ps1 — semantic model (TMDL) emitter. Requires io.ps1 and spec.ps1 to be dot-sourced first.
$ErrorActionPreference = 'Stop'

function Format-TmdlName { param([string]$Name) if ($Name -match '^[A-Za-z_][A-Za-z0-9_]*$') { return $Name }; return "'" + $Name.Replace("'", "''") + "'" }
function ConvertTo-MString { param([string]$S) return '"' + $S.Replace('"', '""') + '"' }
function Format-DaxColumn { param([string]$Table, [string]$Column) return "'" + $Table.Replace("'", "''") + "'[" + $Column + "]" }
function Format-TmdlFormatString { param([string]$F) return $F.Replace('$', '\$') }
function Get-MType {
    param([string]$Type)
    switch ($Type) { 'int64' { 'Int64.Type' } 'double' { 'type number' } 'decimal' { 'Currency.Type' } 'date' { 'type date' } 'datetime' { 'type datetime' } 'boolean' { 'type logical' } default { 'type text' } }
}
function Get-TmdlDataType {
    param([string]$Type)
    switch ($Type) { 'int64' { 'int64' } 'double' { 'double' } 'decimal' { 'decimal' } 'date' { 'dateTime' } 'datetime' { 'dateTime' } 'boolean' { 'boolean' } default { 'string' } }
}

function New-MQuery {
    # Returns the M query as an array of lines (unindented).
    param([Parameter(Mandatory)]$Table, [Parameter(Mandatory)][string]$AbsPath)
    $l = New-Object System.Collections.Generic.List[string]
    $l.Add('let')
    if ($Table.IsXlsx) {
        $l.Add('    Source = Excel.Workbook(File.Contents(' + (ConvertTo-MString $AbsPath) + '), null, true),')
        $l.Add('    Sheet = Source{[Item=' + (ConvertTo-MString ([string]$Table.Sheet)) + ',Kind="Sheet"]}[Data],')
        $prev = 'Sheet'
    } else {
        $delim = if ($Table.Delimiter -eq "`t") { '#(tab)' } else { ([string]$Table.Delimiter).Replace('"', '""') }
        $l.Add('    Source = Csv.Document(File.Contents(' + (ConvertTo-MString $AbsPath) + '),[Delimiter="' + $delim + '", Encoding=' + $Table.Encoding + ', QuoteStyle=QuoteStyle.Csv]),')
        $prev = 'Source'
    }
    if ($Table.HeaderRow -gt 1) { $l.Add('    Skipped = Table.Skip(' + $prev + ', ' + ($Table.HeaderRow - 1) + '),'); $prev = 'Skipped' }
    $l.Add('    Promoted = Table.PromoteHeaders(' + $prev + ', [PromoteAllScalars=true]),')
    $sel = @($Table.Columns | ForEach-Object { ConvertTo-MString $_.Source }) -join ', '
    $l.Add('    Selected = Table.SelectColumns(Promoted, {' + $sel + '}),')
    $prev = 'Selected'
    $renames = @($Table.Columns | Where-Object { $_.Source -ne $_.Name } | ForEach-Object { '{' + (ConvertTo-MString $_.Source) + ', ' + (ConvertTo-MString $_.Name) + '}' })
    if ($renames.Count -gt 0) { $l.Add('    Renamed = Table.RenameColumns(Selected, {' + ($renames -join ', ') + '}),'); $prev = 'Renamed' }
    $types = @($Table.Columns | ForEach-Object { '{' + (ConvertTo-MString $_.Name) + ', ' + (Get-MType $_.Type) + '}' }) -join ', '
    $l.Add('    Typed = Table.TransformColumnTypes(' + $prev + ', {' + $types + '})')
    $l.Add('in')
    $l.Add('    Typed')
    return $l.ToArray()
}

function New-ColumnTmdl {
    param([Parameter(Mandatory)]$C, [bool]$Calculated = $false)
    $l = New-Object System.Collections.Generic.List[string]
    if ($C.Description) { $l.Add("`t/// " + (([string]$C.Description) -replace '\r?\n', ' ')) }
    $l.Add("`tcolumn " + (Format-TmdlName $C.Name))
    $l.Add("`t`tdataType: " + (Get-TmdlDataType $C.Type))
    if ($C.Hidden) { $l.Add("`t`tisHidden") }
    if ($C.Key) { $l.Add("`t`tisKey") }
    $fmt = $C.Format
    if (-not $fmt -and $C.Type -eq 'date') { $fmt = 'yyyy-MM-dd' } elseif (-not $fmt -and $C.Type -eq 'datetime') { $fmt = 'yyyy-MM-dd HH:mm:ss' }
    if ($fmt) { $l.Add("`t`tformatString: " + (Format-TmdlFormatString ([string]$fmt))) }
    $l.Add("`t`tlineageTag: " + (New-LineageTag))
    $l.Add("`t`tsummarizeBy: " + $C.SummarizeBy)
    if ($Calculated) { $l.Add("`t`tisNameInferred"); $l.Add("`t`tisDataTypeInferred"); $l.Add("`t`tsourceColumn: [" + $C.Name + "]") }
    else { $l.Add("`t`tsourceColumn: " + $C.Name) }
    if ($C.SortBy) { $l.Add("`t`tsortByColumn: " + (Format-TmdlName ([string]$C.SortBy))) }
    if ($C.Type -eq 'date') { $l.Add(''); $l.Add("`t`tannotation UnderlyingDateTimeDataType = Date") }
    $l.Add('')
    return $l.ToArray()
}

function New-MeasureTmdl {
    param([Parameter(Mandatory)]$M)
    $l = New-Object System.Collections.Generic.List[string]
    if ($M.Description) { $l.Add("`t/// " + (([string]$M.Description) -replace '\r?\n', ' ')) }
    $dax = ([string]$M.Dax).Trim()
    if ($dax -match '\r?\n') {
        $l.Add("`tmeasure " + (Format-TmdlName $M.Name) + " =")
        foreach ($line in ($dax -split '\r?\n')) { $l.Add("`t`t`t" + $line) }
    } else { $l.Add("`tmeasure " + (Format-TmdlName $M.Name) + " = " + $dax) }
    if ($M.Format) { $l.Add("`t`tformatString: " + (Format-TmdlFormatString ([string]$M.Format))) }
    $l.Add("`t`tlineageTag: " + (New-LineageTag))
    $l.Add('')
    return $l.ToArray()
}

function New-CalendarDax {
    param([Parameter(Mandatory)]$Model)
    $mins = @($Model.Calendar.DateRefs | ForEach-Object { 'MIN ( ' + (Format-DaxColumn $_.Table $_.Column) + ' )' })
    $maxs = @($Model.Calendar.DateRefs | ForEach-Object { 'MAX ( ' + (Format-DaxColumn $_.Table $_.Column) + ' )' })
    $minExpr = $mins[$mins.Count - 1]; for ($i = $mins.Count - 2; $i -ge 0; $i--) { $minExpr = 'MIN ( ' + $mins[$i] + ', ' + $minExpr + ' )' }
    $maxExpr = $maxs[$maxs.Count - 1]; for ($i = $maxs.Count - 2; $i -ge 0; $i--) { $maxExpr = 'MAX ( ' + $maxs[$i] + ', ' + $maxExpr + ' )' }
    return 'ADDCOLUMNS ( CALENDAR ( DATE ( YEAR ( ' + $minExpr + ' ), 1, 1 ), DATE ( YEAR ( ' + $maxExpr + ' ), 12, 31 ) ), "Year", YEAR ( [Date] ), "Quarter", "Q" & FORMAT ( [Date], "q" ), "MonthNo", MONTH ( [Date] ), "Month", FORMAT ( [Date], "MMM" ), "YearMonth", FORMAT ( [Date], "yyyy-MM" ) )'
}

function New-TableTmdl {
    param([Parameter(Mandatory)]$Table, [Parameter(Mandatory)]$Model, [AllowNull()][string]$AbsPath)
    $l = New-Object System.Collections.Generic.List[string]
    $l.Add('table ' + (Format-TmdlName $Table.Name))
    $l.Add("`tlineageTag: " + (New-LineageTag))
    if ($Table.IsCalculated) { $l.Add("`tdataCategory: Time") }
    $l.Add('')
    foreach ($m in $Table.Measures) { $l.AddRange([string[]](New-MeasureTmdl -M $m)) }
    foreach ($c in $Table.Columns) { $l.AddRange([string[]](New-ColumnTmdl -C $c -Calculated $Table.IsCalculated)) }
    if ($Table.IsCalculated) {
        $l.Add("`tpartition " + (Format-TmdlName $Table.Name) + " = calculated")
        $l.Add("`t`tmode: import")
        $l.Add("`t`tsource = " + (New-CalendarDax -Model $Model))
    } else {
        $l.Add("`tpartition " + (Format-TmdlName $Table.Name) + " = m")
        $l.Add("`t`tmode: import")
        $l.Add("`t`tsource =")
        foreach ($line in (New-MQuery -Table $Table -AbsPath $AbsPath)) { $l.Add("`t`t`t`t" + $line) }
    }
    $l.Add('')
    return ($l -join "`r`n")
}

function New-ModelTmdl {
    param([Parameter(Mandatory)]$Model)
    $l = New-Object System.Collections.Generic.List[string]
    $l.Add('model Model')
    $l.Add("`tculture: " + $Model.Locale)
    $l.Add("`tdefaultPowerBIDataSourceVersion: powerBI_V3")
    $l.Add("`tsourceQueryCulture: " + $Model.Locale)
    $l.Add("`tdataAccessOptions")
    $l.Add("`t`tlegacyRedirects")
    $l.Add("`t`treturnErrorValuesAsNull")
    $l.Add('')
    $ti = if ($null -ne $Model.Calendar) { 0 } else { 1 }
    $l.Add('annotation __PBI_TimeIntelligenceEnabled = ' + $ti)
    $l.Add('')
    $order = @($Model.Tables.Keys | Where-Object { -not $Model.Tables[$_].IsCalculated } | ForEach-Object { '"' + $_.Replace('"', '\"') + '"' })
    $l.Add('annotation PBI_QueryOrder = [' + ($order -join ',') + ']')
    $l.Add('')
    foreach ($n in $Model.Tables.Keys) { $l.Add('ref table ' + (Format-TmdlName $n)) }
    $l.Add('')
    return ($l -join "`r`n")
}

function New-RelationshipsTmdl {
    param([Parameter(Mandatory)]$Model)
    $l = New-Object System.Collections.Generic.List[string]
    foreach ($r in $Model.Relationships) {
        $name = (($r.FromTable + '_' + $r.FromColumn + '_' + $r.ToTable) -replace '[^\p{L}\p{N}_]', '_')
        $l.Add('relationship ' + $name)
        if (-not $r.Active) { $l.Add("`tisActive: false") }
        if ($r.CrossFilter -eq 'both') { $l.Add("`tcrossFilteringBehavior: bothDirections") }
        $l.Add("`tfromColumn: " + (Format-TmdlName $r.FromTable) + '.' + (Format-TmdlName $r.FromColumn))
        $l.Add("`ttoColumn: " + (Format-TmdlName $r.ToTable) + '.' + (Format-TmdlName $r.ToColumn))
        $l.Add('')
    }
    return ($l -join "`r`n")
}

function Write-SemanticModel {
    # Writes <Dir>/definition.pbism and <Dir>/definition/**. Returns @{ Warnings = @() }
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Dir, [Parameter(Mandatory)][string]$RawDataDir)
    $warnings = New-Object System.Collections.Generic.List[string]
    $def = Join-Path $Dir 'definition'
    ConvertTo-JsonFile -Path (Join-Path $Dir 'definition.pbism') -Object ([ordered]@{ '$schema' = 'https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json'; version = '4.2'; settings = [ordered]@{} })
    Write-Utf8File -Path (Join-Path $def 'database.tmdl') -Content "database`r`n`tcompatibilityLevel: 1606`r`n"
    Write-Utf8File -Path (Join-Path $def 'model.tmdl') -Content (New-ModelTmdl -Model $Model)
    if ($Model.Relationships.Count -gt 0) { Write-Utf8File -Path (Join-Path $def 'relationships.tmdl') -Content (New-RelationshipsTmdl -Model $Model) }
    foreach ($n in $Model.Tables.Keys) {
        $t = $Model.Tables[$n]
        $abs = $null
        if (-not $t.IsCalculated) {
            $abs = (Get-Item -LiteralPath (Join-Path $RawDataDir $t.File)).FullName
            if ($abs -notmatch '^([A-Za-z]:\\|\\\\)') { $warnings.Add("data path '$abs' is not a Windows path - rebuild on Windows before opening in Power BI Desktop") }
        }
        Write-Utf8File -Path (Join-Path (Join-Path $def 'tables') ($n + '.tmdl')) -Content (New-TableTmdl -Table $t -Model $Model -AbsPath $abs)
    }
    return @{ Warnings = $warnings.ToArray() }
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter tmdl
```
Expected: `8 tests, 0 failed`.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/powerbi/scripts/lib/tmdl.ps1 .claude/skills/powerbi/scripts/tests/tmdl.test.ps1
git commit -m "feat(scripts): TMDL semantic model emitter with M queries and Calendar table"
```

---

### Task 10: `lib/pbir.ps1` — report emitter

**Files:**
- Create: `.claude/skills/powerbi/scripts/lib/pbir.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/pbir.test.ps1`

- [ ] **Step 1: Write the failing test**

```powershell
. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'spec.ps1')
. (Join-Path $libDir 'pbir.ps1')
$flat = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))

Test-Case 'ConvertTo-EntityExpr renders Column / Measure / Aggregation' {
    $c = ConvertTo-EntityExpr -Ref (Resolve-FieldRef -Model $flat -Ref 'Apps.장르')
    Assert-Equal 'Apps' $c.Column.Expression.SourceRef.Entity; Assert-Equal '장르' $c.Column.Property
    $m = ConvertTo-EntityExpr -Ref (Resolve-FieldRef -Model $flat -Ref 'Apps.앱 수')
    Assert-Equal '앱 수' $m.Measure.Property
    $a = ConvertTo-EntityExpr -Ref (Resolve-FieldRef -Model $flat -Ref 'avg:Apps.price')
    Assert-Equal 1 $a.Aggregation.Function; Assert-Equal 'price' $a.Aggregation.Expression.Column.Property
}
Test-Case 'New-Projection sets queryRef and nativeQueryRef' {
    $p = New-Projection -Ref (Resolve-FieldRef -Model $flat -Ref 'sum:Apps.rating_count_tot')
    Assert-Equal 'Sum(Apps.rating_count_tot)' $p.queryRef; Assert-Equal 'Sum of rating_count_tot' $p.nativeQueryRef
    $q = New-Projection -Ref (Resolve-FieldRef -Model $flat -Ref 'Apps.앱 수')
    Assert-Equal 'Apps.앱 수' $q.queryRef; Assert-Equal '앱 수' $q.nativeQueryRef
}
Test-Case 'ConvertTo-FilterLiteral formats strings, ints, decimals, bools' {
    Assert-Equal "'Games'" (ConvertTo-FilterLiteral 'Games')
    Assert-Equal "'It''s'" (ConvertTo-FilterLiteral "It's")
    Assert-Equal '10L' (ConvertTo-FilterLiteral 10)
    Assert-Equal '2.5D' (ConvertTo-FilterLiteral 2.5)
    Assert-Equal '2.5D' (ConvertTo-FilterLiteral ([decimal]2.5))
    Assert-Equal 'true' (ConvertTo-FilterLiteral $true)
}
Test-Case 'New-VisualJson: card with title' {
    $v = $flat.Pages[0].visuals[1]
    $j = New-VisualJson -Model $flat -Visual $v -Index 1
    Assert-Match '^[0-9a-f]{20}$' $j.name
    Assert-Equal 'cardVisual' $j.visual.visualType
    Assert-Equal '앱 수' $j.visual.query.queryState.Data.projections[0].field.Measure.Property
    Assert-Equal "'앱 수'" $j.visual.visualContainerObjects.title[0].properties.text.expr.Literal.Value
    Assert-Equal 40 $j.position.x; Assert-Equal 168 $j.position.y; Assert-Equal 425 $j.position.width; Assert-Equal 235 $j.position.height; Assert-Equal 1000 $j.position.tabOrder
    Assert-True $j.visual.drillFilterOtherVisuals 'drill flag'
}
Test-Case 'New-VisualJson: bar chart with sort and TopN filter' {
    $j = New-VisualJson -Model $flat -Visual $flat.Pages[0].visuals[3] -Index 3
    Assert-Equal 'clusteredBarChart' $j.visual.visualType
    Assert-Equal '장르' $j.visual.query.queryState.Category.projections[0].field.Column.Property
    Assert-Equal 'Descending' $j.visual.query.sortDefinition.sort[0].direction
    Assert-Equal '앱 수' $j.visual.query.sortDefinition.sort[0].field.Measure.Property
    Assert-True $j.visual.query.sortDefinition.isDefaultSort 'default sort'
    $f = $j.filterConfig.filters[0]
    Assert-Equal 'TopN' $f.type; Assert-Equal '장르' $f.field.Column.Property
    Assert-Equal 2 $f.filter.Version; Assert-Equal 'Apps' $f.filter.From[0].Entity; Assert-Equal 't' $f.filter.From[0].Name
    $top = $f.filter.Where[0].Condition.TopN
    Assert-Equal 3 $top.Count; Assert-Equal 't' $top.Expression.Column.Expression.SourceRef.Source
    Assert-Equal 2 $top.OrderBy[0].Direction; Assert-Equal '앱 수' $top.OrderBy[0].Expression.Measure.Property
}
Test-Case 'New-VisualJson: scatter aggregations and textbox' {
    $s = New-VisualJson -Model $flat -Visual $flat.Pages[0].visuals[5] -Index 5
    Assert-Equal 1 $s.visual.query.queryState.X.projections[0].field.Aggregation.Function
    Assert-Equal 0 $s.visual.query.queryState.Size.projections[0].field.Aggregation.Function
    $t = New-VisualJson -Model $flat -Visual $flat.Pages[0].visuals[0] -Index 0
    Assert-Equal 'textbox' $t.visual.visualType
    $run = $t.visual.objects.general[0].properties.paragraphs[0].textRuns[0]
    Assert-Equal '앱스토어 대시보드' $run.value; Assert-Equal '24pt' $run.textStyle.fontSize
    Assert-True ($null -eq $t.visual.query) 'textbox has no query'
}
Test-Case 'New-InFilter builds a Categorical In filter with nested value arrays' {
    $f = New-InFilter -Model $flat -Field 'Apps.장르' -Values @('Games', 'Education')
    Assert-Equal 'Categorical' $f.type
    $in = $f.filter.Where[0].Condition.In
    Assert-Equal '장르' $in.Expressions[0].Column.Property
    Assert-Equal 2 $in.Values.Count
    Assert-Equal "'Games'" $in.Values[0][0].Literal.Value
    $json = $f | ConvertTo-Json -Depth 32
    Assert-Match '"Values":\s*\[\s*\[\s*\{' $json
}
Test-Case 'Write-Report writes pages, visuals, static resources and metadata' {
    $dir = Join-Path $tmpRoot 'rep'
    $r = Write-Report -Model $flat -Dir $dir -TemplatesDir (Join-Path (Split-Path -Parent (Split-Path -Parent $libDir)) 'templates') -ModelDirName 'AppsFlat.SemanticModel'
    Assert-Equal 2 $r.Pages; Assert-Equal 11 $r.Visuals
    $pbir = [IO.File]::ReadAllText((Join-Path $dir 'definition.pbir')) | ConvertFrom-Json
    Assert-Equal '../AppsFlat.SemanticModel' $pbir.datasetReference.byPath.path
    $pages = [IO.File]::ReadAllText((Join-Path $dir 'definition/pages/pages.json')) | ConvertFrom-Json
    Assert-Equal 2 $pages.pageOrder.Count; Assert-Equal $pages.pageOrder[0] $pages.activePageName
    $p2 = [IO.File]::ReadAllText((Join-Path $dir ('definition/pages/' + $pages.pageOrder[1] + '/page.json'))) | ConvertFrom-Json
    Assert-Equal '상세' $p2.displayName; Assert-Equal 1920 $p2.width; Assert-Equal 'FitToPage' $p2.displayOption
    Assert-Equal 'Categorical' $p2.filterConfig.filters[0].type
    Assert-Equal 8 (Get-ChildItem (Join-Path $dir ('definition/pages/' + $pages.pageOrder[0] + '/visuals')) -Directory).Count
    Assert-True (Test-Path (Join-Path $dir 'StaticResources/SharedResources/BaseThemes/Fluent2-CY26SU08.json')) 'theme'
    Assert-True (Test-Path (Join-Path $dir 'definition/report.json')) 'report.json'
    Assert-Equal '2.0.0' (([IO.File]::ReadAllText((Join-Path $dir 'definition/version.json')) | ConvertFrom-Json).version)
}
```

- [ ] **Step 2: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter pbir
```
Expected: fails loading `pbir.ps1`.

- [ ] **Step 3: Write `lib/pbir.ps1`**

```powershell
# pbir.ps1 — PBIR report emitter. Requires io.ps1 and spec.ps1 dot-sourced first.
$ErrorActionPreference = 'Stop'
$script:SchemaBase = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition'
$script:Schemas = @{
    visual  = "$script:SchemaBase/visualContainer/2.12.0/schema.json"
    page    = "$script:SchemaBase/page/2.1.0/schema.json"
    pages   = "$script:SchemaBase/pagesMetadata/1.1.0/schema.json"
    version = "$script:SchemaBase/versionMetadata/1.0.0/schema.json"
    pbir    = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json'
}
$script:AggLabels = @{ sum = 'Sum'; avg = 'Avg'; countDistinct = 'CountDistinct'; min = 'Min'; max = 'Max'; count = 'Count' }

function ConvertTo-EntityExpr {
    param([Parameter(Mandatory)]$Ref)
    $src = [ordered]@{ Expression = [ordered]@{ SourceRef = [ordered]@{ Entity = $Ref.Table } }; Property = $Ref.Name }
    switch ($Ref.Kind) {
        'Measure'     { return [ordered]@{ Measure = $src } }
        'Column'      { return [ordered]@{ Column = $src } }
        'Aggregation' { return [ordered]@{ Aggregation = [ordered]@{ Expression = [ordered]@{ Column = $src }; Function = $Ref.Function } } }
    }
    throw "unsupported field kind $($Ref.Kind)"
}

function ConvertTo-SourceExpr {
    # Same as ConvertTo-EntityExpr but with SourceRef.Source = alias (used inside filter Where clauses).
    param([Parameter(Mandatory)]$Ref, [Parameter(Mandatory)][string]$Alias)
    $src = [ordered]@{ Expression = [ordered]@{ SourceRef = [ordered]@{ Source = $Alias } }; Property = $Ref.Name }
    switch ($Ref.Kind) {
        'Measure'     { return [ordered]@{ Measure = $src } }
        'Column'      { return [ordered]@{ Column = $src } }
        'Aggregation' { return [ordered]@{ Aggregation = [ordered]@{ Expression = [ordered]@{ Column = $src }; Function = $Ref.Function } } }
    }
    throw "unsupported field kind $($Ref.Kind)"
}

function New-Projection {
    param([Parameter(Mandatory)]$Ref)
    if ($Ref.Kind -eq 'Aggregation') {
        $label = $script:AggLabels[$Ref.FunctionName]
        return [ordered]@{ field = (ConvertTo-EntityExpr -Ref $Ref); queryRef = "$label($($Ref.Table).$($Ref.Name))"; nativeQueryRef = "$label of $($Ref.Name)" }
    }
    return [ordered]@{ field = (ConvertTo-EntityExpr -Ref $Ref); queryRef = "$($Ref.Table).$($Ref.Name)"; nativeQueryRef = $Ref.Name }
}

function ConvertTo-FilterLiteral {
    param($V)
    if ($V -is [bool]) { return $V.ToString().ToLowerInvariant() }
    if ($V -is [int] -or $V -is [long] -or $V -is [int16] -or $V -is [byte]) { return ([long]$V).ToString() + 'L' }
    if ($V -is [double] -or $V -is [decimal] -or $V -is [single]) {
        $d = [double]$V
        if ($d -eq [Math]::Floor($d) -and [Math]::Abs($d) -lt 9e15) { return ([long]$d).ToString() + 'L' }
        return $d.ToString('R', [Globalization.CultureInfo]::InvariantCulture) + 'D'
    }
    return "'" + ([string]$V).Replace("'", "''") + "'"
}

function New-InFilter {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Field, [Parameter(Mandatory)]$Values)
    $r = Resolve-FieldRef -Model $Model -Ref $Field
    $vals = New-Object System.Collections.Generic.List[object]
    foreach ($v in (ConvertTo-Array $Values)) { $vals.Add([object[]]@([ordered]@{ Literal = [ordered]@{ Value = (ConvertTo-FilterLiteral $v) } })) }
    return [ordered]@{
        name = (New-HexId); field = (ConvertTo-EntityExpr -Ref $r); type = 'Categorical'
        filter = [ordered]@{
            Version = 2
            From = @([ordered]@{ Name = 't'; Entity = $r.Table; Type = 0 })
            Where = @([ordered]@{ Condition = [ordered]@{ In = [ordered]@{ Expressions = @((ConvertTo-SourceExpr -Ref $r -Alias 't')); Values = $vals.ToArray() } } })
        }
    }
}

function New-TopNFilter {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$CategoryField, [Parameter(Mandatory)][string]$ByField, [Parameter(Mandatory)][int]$N)
    $c = Resolve-FieldRef -Model $Model -Ref $CategoryField
    $b = Resolve-FieldRef -Model $Model -Ref $ByField
    $from = New-Object System.Collections.Generic.List[object]
    $from.Add([ordered]@{ Name = 't'; Entity = $c.Table; Type = 0 })
    $bAlias = 't'
    if ($b.Table -ne $c.Table) { $from.Add([ordered]@{ Name = 'm'; Entity = $b.Table; Type = 0 }); $bAlias = 'm' }
    return [ordered]@{
        name = (New-HexId); field = (ConvertTo-EntityExpr -Ref $c); type = 'TopN'
        filter = [ordered]@{
            Version = 2
            From = $from.ToArray()
            Where = @([ordered]@{ Condition = [ordered]@{ TopN = [ordered]@{
                Expression = (ConvertTo-SourceExpr -Ref $c -Alias 't'); Count = $N
                OrderBy = @([ordered]@{ Expression = (ConvertTo-SourceExpr -Ref $b -Alias $bAlias); Direction = 2 }) } } })
        }
    }
}

function New-TitleObjects {
    param([string]$Title)
    return [ordered]@{ title = @([ordered]@{ properties = [ordered]@{
        show = [ordered]@{ expr = [ordered]@{ Literal = [ordered]@{ Value = 'true' } } }
        text = [ordered]@{ expr = [ordered]@{ Literal = [ordered]@{ Value = "'" + $Title.Replace("'", "''") + "'" } } } } }) }
}

function New-VisualJson {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)]$Visual, [Parameter(Mandatory)][int]$Index)
    $type = [string](Get-Prop $Visual 'type')
    $cat = $script:VisualCatalog[$type]
    $pos = Get-GridPosition -Grid (ConvertTo-Array (Get-Prop $Visual 'grid'))
    $position = [ordered]@{ x = $pos.x; y = $pos.y; z = $Index; height = $pos.height; width = $pos.width; tabOrder = $Index * 1000 }
    $visual = [ordered]@{ visualType = $cat.PbirType }
    $fields = Get-Prop $Visual 'fields' $null
    if ($cat.IsText) {
        $size = [int](Get-Prop $Visual 'fontSize' 20)
        $visual.objects = [ordered]@{ general = @([ordered]@{ properties = [ordered]@{ paragraphs = @([ordered]@{ textRuns = @([ordered]@{
            value = [string](Get-Prop $Visual 'text' ''); textStyle = [ordered]@{ fontSize = "$($size)pt"; fontWeight = 'bold' } }) }) } }) }
    } else {
        $qs = [ordered]@{}
        foreach ($role in ($cat.Required + $cat.Optional)) {
            $refs = ConvertTo-Array (Get-Prop $fields $role)
            if ($refs.Count -eq 0) { continue }
            $pbirRole = if ($cat.RoleMap.ContainsKey($role)) { $cat.RoleMap[$role] } else { $role }
            $projs = New-Object System.Collections.Generic.List[object]
            foreach ($ref in $refs) { $projs.Add((New-Projection -Ref (Resolve-FieldRef -Model $Model -Ref ([string]$ref)))) }
            $qs[$pbirRole] = [ordered]@{ projections = $projs.ToArray() }
        }
        $query = [ordered]@{ queryState = $qs }
        $sort = Get-Prop $Visual 'sort' $null
        if ($null -ne $sort) {
            $dir = if ([string](Get-Prop $sort 'direction' 'desc') -eq 'asc') { 'Ascending' } else { 'Descending' }
            $query.sortDefinition = [ordered]@{ sort = @([ordered]@{ field = (ConvertTo-EntityExpr -Ref (Resolve-FieldRef -Model $Model -Ref ([string](Get-Prop $sort 'field')))); direction = $dir }); isDefaultSort = $true }
        }
        $visual.query = $query
        $title = Get-Prop $Visual 'title' $null
        if ($title) { $visual.visualContainerObjects = New-TitleObjects -Title ([string]$title) }
    }
    $visual.drillFilterOtherVisuals = $true
    $vc = [ordered]@{ '$schema' = $script:Schemas.visual; name = (New-HexId); position = $position; visual = $visual }
    $filters = New-Object System.Collections.Generic.List[object]
    $topN = Get-Prop $Visual 'topN' $null
    if ($null -ne $topN) {
        $catRefs = ConvertTo-Array (Get-Prop $fields 'Category'); if ($catRefs.Count -eq 0) { $catRefs = ConvertTo-Array (Get-Prop $fields 'Rows') }
        $filters.Add((New-TopNFilter -Model $Model -CategoryField ([string]$catRefs[0]) -ByField ([string](Get-Prop $topN 'by')) -N ([int](Get-Prop $topN 'n'))))
    }
    foreach ($f in (ConvertTo-Array (Get-Prop $Visual 'filters'))) { $filters.Add((New-InFilter -Model $Model -Field ([string](Get-Prop $f 'field')) -Values (Get-Prop $f 'in'))) }
    if ($filters.Count -gt 0) { $vc.filterConfig = [ordered]@{ filters = $filters.ToArray() } }
    return $vc
}

function Write-Report {
    # Writes <Dir>/definition.pbir, <Dir>/definition/**, <Dir>/StaticResources/**. Returns @{ Pages; Visuals }
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Dir, [Parameter(Mandatory)][string]$TemplatesDir, [Parameter(Mandatory)][string]$ModelDirName)
    $def = Join-Path $Dir 'definition'
    ConvertTo-JsonFile -Path (Join-Path $Dir 'definition.pbir') -Object ([ordered]@{ '$schema' = $script:Schemas.pbir; version = '4.0'; datasetReference = [ordered]@{ byPath = [ordered]@{ path = '../' + $ModelDirName } } })
    ConvertTo-JsonFile -Path (Join-Path $def 'version.json') -Object ([ordered]@{ '$schema' = $script:Schemas.version; version = '2.0.0' })
    Write-Utf8File -Path (Join-Path $def 'report.json') -Content ([IO.File]::ReadAllText((Join-Path $TemplatesDir 'report.json'), [Text.Encoding]::UTF8))
    $themeDir = Join-Path $Dir 'StaticResources/SharedResources/BaseThemes'
    Write-Utf8File -Path (Join-Path $themeDir 'Fluent2-CY26SU08.json') -Content ([IO.File]::ReadAllText((Join-Path $TemplatesDir 'theme/Fluent2-CY26SU08.json'), [Text.Encoding]::UTF8))
    $pageIds = New-Object System.Collections.Generic.List[string]
    $visualCount = 0
    foreach ($page in $Model.Pages) {
        $pid = New-HexId; $pageIds.Add($pid)
        $pageDir = Join-Path (Join-Path $def 'pages') $pid
        $pj = [ordered]@{ '$schema' = $script:Schemas.page; name = $pid; displayName = [string](Get-Prop $page 'name'); displayOption = 'FitToPage'; height = 1080; width = 1920 }
        $pf = New-Object System.Collections.Generic.List[object]
        foreach ($f in (ConvertTo-Array (Get-Prop $page 'filters'))) { $pf.Add((New-InFilter -Model $Model -Field ([string](Get-Prop $f 'field')) -Values (Get-Prop $f 'in'))) }
        if ($pf.Count -gt 0) { $pj.filterConfig = [ordered]@{ filters = $pf.ToArray() } }
        ConvertTo-JsonFile -Path (Join-Path $pageDir 'page.json') -Object $pj
        $i = 0
        foreach ($v in (ConvertTo-Array (Get-Prop $page 'visuals'))) {
            $vj = New-VisualJson -Model $Model -Visual $v -Index $i
            ConvertTo-JsonFile -Path (Join-Path (Join-Path (Join-Path $pageDir 'visuals') $vj.name) 'visual.json') -Object $vj
            $i++; $visualCount++
        }
    }
    ConvertTo-JsonFile -Path (Join-Path $def 'pages/pages.json') -Object ([ordered]@{ '$schema' = $script:Schemas.pages; pageOrder = $pageIds.ToArray(); activePageName = $pageIds[0] })
    return @{ Pages = $pageIds.Count; Visuals = $visualCount }
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter pbir
```
Expected: `8 tests, 0 failed`.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/powerbi/scripts/lib/pbir.ps1 .claude/skills/powerbi/scripts/tests/pbir.test.ps1
git commit -m "feat(scripts): PBIR report emitter (visual catalog, grid, sort, TopN/In filters)"
```

---

### Task 11: `build-pbip.ps1` — entry point + end-to-end build tests

**Files:**
- Create: `.claude/skills/powerbi/scripts/build-pbip.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/build.test.ps1`

- [ ] **Step 1: Write the failing test**

```powershell
$buildScript = Join-Path (Split-Path -Parent $libDir) 'build-pbip.ps1'
$flatOut = Join-Path $tmpRoot 'flat'

Test-Case 'build of spec-flat produces a complete PBIP project' {
    $log = (& $buildScript -Spec (Join-Path $fixtures 'spec-flat.json') -RawData $fixtures -OutputDir $flatOut | Out-String)
    Assert-Equal 0 $LASTEXITCODE $log
    Assert-Match 'BUILT .*AppsFlat\.pbip\s+\(tables=2 measures=3 pages=2 visuals=11\)' $log
    foreach ($f in 'AppsFlat.pbip', '.gitignore', 'AppsFlat.Report/definition.pbir', 'AppsFlat.Report/definition/report.json', 'AppsFlat.Report/definition/version.json',
                   'AppsFlat.Report/definition/pages/pages.json', 'AppsFlat.Report/StaticResources/SharedResources/BaseThemes/Fluent2-CY26SU08.json',
                   'AppsFlat.SemanticModel/definition.pbism', 'AppsFlat.SemanticModel/definition/database.tmdl', 'AppsFlat.SemanticModel/definition/model.tmdl',
                   'AppsFlat.SemanticModel/definition/relationships.tmdl', 'AppsFlat.SemanticModel/definition/tables/Apps.tmdl', 'AppsFlat.SemanticModel/definition/tables/Calendar.tmdl') {
        Assert-True (Test-Path (Join-Path $flatOut $f)) "missing $f"
    }
    foreach ($j in (Get-ChildItem -Recurse -File -Path $flatOut | Where-Object { $_.Extension -in '.json', '.pbip', '.pbir', '.pbism' })) { [void]([IO.File]::ReadAllText($j.FullName) | ConvertFrom-Json) }
    $bytes = [IO.File]::ReadAllBytes((Join-Path $flatOut 'AppsFlat.Report/definition/pages/pages.json')); Assert-Equal 0x7B $bytes[0] 'no BOM'
    $tm = [IO.File]::ReadAllText((Join-Path $flatOut 'AppsFlat.SemanticModel/definition/tables/Apps.tmdl')); Assert-True ($tm.Contains("`r`n")) 'CRLF'
    $pbip = [IO.File]::ReadAllText((Join-Path $flatOut 'AppsFlat.pbip')) | ConvertFrom-Json
    Assert-Equal 'AppsFlat.Report' $pbip.artifacts[0].report.path
}
Test-Case 'visual files match the spec roles (page 1) and page 2 has the In filter' {
    $pagesDir = Join-Path $flatOut 'AppsFlat.Report/definition/pages'
    $pages = [IO.File]::ReadAllText((Join-Path $pagesDir 'pages.json')) | ConvertFrom-Json
    $visuals = @(Get-ChildItem -Recurse -Filter visual.json -Path (Join-Path $pagesDir $pages.pageOrder[0]) | ForEach-Object { [IO.File]::ReadAllText($_.FullName) | ConvertFrom-Json })
    Assert-Equal 8 $visuals.Count
    $types = @($visuals | ForEach-Object { $_.visual.visualType } | Sort-Object)
    Assert-Equal 'cardVisual,cardVisual,clusteredBarChart,lineChart,scatterChart,slicer,tableEx,textbox' ($types -join ',')
    $line = $visuals | Where-Object { $_.visual.visualType -eq 'lineChart' }
    Assert-Equal 'Calendar' $line.visual.query.queryState.Category.projections[0].field.Column.Expression.SourceRef.Entity
    $tbl = $visuals | Where-Object { $_.visual.visualType -eq 'tableEx' }
    Assert-Equal 3 $tbl.visual.query.queryState.Values.projections.Count
    $p2 = [IO.File]::ReadAllText((Join-Path $pagesDir ($pages.pageOrder[1] + '/page.json'))) | ConvertFrom-Json
    Assert-Equal "'Games'" $p2.filterConfig.filters[0].filter.Where[0].Condition.In.Values[0][0].Literal.Value
    $v2 = @(Get-ChildItem -Recurse -Filter visual.json -Path (Join-Path $pagesDir $pages.pageOrder[1]) | ForEach-Object { [IO.File]::ReadAllText($_.FullName) | ConvertFrom-Json })
    $pivot = $v2 | Where-Object { $_.visual.visualType -eq 'pivotTable' }
    Assert-Equal 2 $pivot.visual.query.queryState.Values.projections.Count
    Assert-Equal 'Year' $pivot.visual.query.queryState.Columns.projections[0].field.Column.Property
    $gauge = $v2 | Where-Object { $_.visual.visualType -eq 'gauge' }
    Assert-Equal 4 $gauge.visual.query.queryState.MaxValue.projections[0].field.Aggregation.Function
}
Test-Case 'build of spec-star: csv partitions, relationships, Series role' {
    $out = Join-Path $tmpRoot 'star'
    $log = (& $buildScript -Spec (Join-Path $fixtures 'spec-star.json') -RawData $fixtures -OutputDir $out | Out-String)
    Assert-Equal 0 $LASTEXITCODE $log
    Assert-Match 'tables=3 measures=2 pages=1 visuals=3' $log
    $sales = [IO.File]::ReadAllText((Join-Path $out 'SalesStar.SemanticModel/definition/tables/Sales.tmdl'))
    Assert-Match 'Csv\.Document\(File\.Contents\("[^"]*sales_cp949\.csv"\),\[Delimiter=",", Encoding=949, QuoteStyle=QuoteStyle\.Csv\]\)' $sales
    Assert-True (-not $sales.Contains('RenameColumns')) 'no renames in star spec'
    $rels = [IO.File]::ReadAllText((Join-Path $out 'SalesStar.SemanticModel/definition/relationships.tmdl'))
    Assert-Match "fromColumn: Sales\.'지역코드'\r?\n\ttoColumn: Regions\.'지역코드'" $rels
    Assert-Match "toColumn: Calendar\.Date" $rels
    $vis = @(Get-ChildItem -Recurse -Filter visual.json -Path (Join-Path $out 'SalesStar.Report') | ForEach-Object { [IO.File]::ReadAllText($_.FullName) | ConvertFrom-Json })
    $col = $vis | Where-Object { $_.visual.visualType -eq 'clusteredColumnChart' }
    Assert-Equal '제품' $col.visual.query.queryState.Series.projections[0].field.Column.Property
}
Test-Case 'broken spec exits 1 with three ERROR lines and writes nothing' {
    $out = Join-Path $tmpRoot 'broken'
    $log = (& $buildScript -Spec (Join-Path $fixtures 'spec-broken.json') -RawData $fixtures -OutputDir $out | Out-String)
    Assert-Equal 1 $LASTEXITCODE
    Assert-Equal 3 ([regex]::Matches($log, '(?m)^ERROR \d+: ').Count) $log
    Assert-Match '3 error\(s\)' $log
    Assert-True (-not (Test-Path $out)) 'no output on error'
}
Test-Case 'rebuild replaces definition folders but keeps .pbi caches' {
    $cache = Join-Path $flatOut 'AppsFlat.SemanticModel/.pbi/cache.abf'
    New-Item -ItemType Directory -Path (Split-Path -Parent $cache) -Force | Out-Null; [IO.File]::WriteAllText($cache, 'x')
    $stray = Join-Path $flatOut 'AppsFlat.SemanticModel/definition/tables/Stray.tmdl'; [IO.File]::WriteAllText($stray, 'table Stray')
    $log = (& $buildScript -Spec (Join-Path $fixtures 'spec-flat.json') -RawData $fixtures -OutputDir $flatOut | Out-String)
    Assert-Equal 0 $LASTEXITCODE $log
    Assert-True (Test-Path $cache) 'cache kept'
    Assert-True (-not (Test-Path $stray)) 'stray definition removed'
    $log2 = (& $buildScript -Spec (Join-Path $fixtures 'spec-flat.json') -RawData $fixtures -OutputDir $flatOut -Force | Out-String)
    Assert-True (-not (Test-Path $cache)) '-Force removes caches too'
}
```

- [ ] **Step 2: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter build
```
Expected: fails (script not found).

- [ ] **Step 3: Write `scripts/build-pbip.ps1`**

```powershell
<#
.SYNOPSIS  Validates report-spec.json and builds a Power BI Project (.pbip) next to it.
.EXAMPLE   powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-pbip.ps1 -Spec output/Sales/report-spec.json
.PARAMETER RawData    Folder holding the data files (default: <workspace>/rawdata)
.PARAMETER OutputDir  Where to write <Name>.pbip (default: the spec's folder)
.PARAMETER Force      Also delete .pbi caches of a previous build
#>
param([Parameter(Mandatory)][string]$Spec, [string]$RawData = '', [string]$OutputDir = '', [switch]$Force)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib/io.ps1')
. (Join-Path $PSScriptRoot 'lib/spec.ps1')
. (Join-Path $PSScriptRoot 'lib/tmdl.ps1')
. (Join-Path $PSScriptRoot 'lib/pbir.ps1')

$root = Get-WorkspaceRoot -ScriptRoot $PSScriptRoot
if (-not $RawData) { $RawData = Join-Path $root 'rawdata' }
if (-not (Test-Path -LiteralPath $Spec)) { Write-Output "ERROR 1: spec file not found: $Spec"; exit 1 }
$specPath = (Get-Item -LiteralPath $Spec).FullName
try { $specObj = Read-Spec -Path $specPath } catch { Write-Output ("ERROR 1: spec is not valid JSON: " + $_.Exception.Message); exit 1 }

$v = Get-SpecValidation -Spec $specObj -RawDataDir $RawData
if ($v.Errors.Count -gt 0) {
    $i = 0
    foreach ($e in $v.Errors) { $i++; Write-Output ("ERROR {0}: {1}" -f $i, $e) }
    Write-Output ("{0} error(s) in {1} - fix the spec and rebuild" -f $v.Errors.Count, $specPath)
    exit 1
}
foreach ($w in $v.Warnings) { Write-Output ("WARN: " + $w) }
$model = $v.Model
$name = $model.Name
if (-not $OutputDir) { $OutputDir = Split-Path -Parent $specPath }
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
$reportDir = Join-Path $OutputDir ($name + '.Report')
$modelDir  = Join-Path $OutputDir ($name + '.SemanticModel')

if ($Force) { foreach ($d in $reportDir, $modelDir) { if (Test-Path $d) { Remove-Item -Recurse -Force $d } } }
else {
    foreach ($d in (Join-Path $reportDir 'definition'), (Join-Path $reportDir 'StaticResources'), (Join-Path $modelDir 'definition')) { if (Test-Path $d) { Remove-Item -Recurse -Force $d } }
}

$mw = Write-SemanticModel -Model $model -Dir $modelDir -RawDataDir $RawData
$rw = Write-Report -Model $model -Dir $reportDir -TemplatesDir (Join-Path (Split-Path -Parent $PSScriptRoot) 'templates') -ModelDirName ($name + '.SemanticModel')
$pbipPath = Join-Path $OutputDir ($name + '.pbip')
ConvertTo-JsonFile -Path $pbipPath -Object ([ordered]@{
    '$schema' = 'https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json'; version = '1.0'
    artifacts = @([ordered]@{ report = [ordered]@{ path = $name + '.Report' } }); settings = [ordered]@{ enableAutoRecovery = $true } })
Write-Utf8File -Path (Join-Path $OutputDir '.gitignore') -Content "**/.pbi/localSettings.json`n**/.pbi/cache.abf`n"
foreach ($w in $mw.Warnings) { Write-Output ("WARN: " + $w) }
$measureCount = 0; foreach ($k in $model.Tables.Keys) { $measureCount += @($model.Tables[$k].Measures).Count }
Write-Output ("BUILT {0}  (tables={1} measures={2} pages={3} visuals={4})" -f $pbipPath, $model.Tables.Count, $measureCount, $rw.Pages, $rw.Visuals)
exit 0
```

- [ ] **Step 4: Run the whole suite**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1
```
Expected: all test files pass (`… tests, 0 failed`), exit 0.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/powerbi/scripts/build-pbip.ps1 .claude/skills/powerbi/scripts/tests/build.test.ps1
git commit -m "feat(scripts): build-pbip entry point with validation and end-to-end tests"
```

---

### Task 12: `lib/guide.ps1` + `build-guide.ps1` — Korean manual guide with English UI terms

**Files:**
- Create: `.claude/skills/powerbi/scripts/lib/guide.ps1`
- Create: `.claude/skills/powerbi/scripts/build-guide.ps1`
- Create: `.claude/skills/powerbi/scripts/tests/guide.test.ps1`
- Modify: `.claude/skills/powerbi/scripts/profile-data.ps1`, `build-pbip.ps1` (console UTF-8 line)

- [ ] **Step 1: Write the failing test**

```powershell
. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'spec.ps1')
. (Join-Path $libDir 'guide.ps1')
$guideScript = Join-Path (Split-Path -Parent $libDir) 'build-guide.ps1'

Test-Case 'New-ManualGuide: data import section per table with English UI terms' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $g = New-ManualGuide -Model $m -RawDataDir 'C:\work\rawdata'
    Assert-Match '^# AppsFlat — 수동 제작 가이드 \(Power BI Desktop\)' $g
    Assert-Match '\*\*Home > Get Data > Excel Workbook\*\*' $g
    Assert-Match '\*\*Navigator\*\*에서 시트 \*\*Data\*\*' $g
    Assert-Match '\*\*Home > Choose Columns\*\* → 다음 열만 체크: `id`, `track_name`, `size_bytes`' $g
    Assert-Match '`track_name` → `앱 이름`' $g
    Assert-Match '`price` → \*\*Fixed decimal number\*\*' $g
    Assert-Match '`출시일` → \*\*Date\*\*' $g
    Assert-Match '\*\*Home > Close & Apply\*\*' $g
}
Test-Case 'New-ManualGuide: model, calendar, measures' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-star.json'))
    $g = New-ManualGuide -Model $m -RawDataDir 'C:\work\rawdata'
    Assert-Match '\*\*File Origin: 949: Korean\*\*' $g
    Assert-Match '\*\*Home > Manage relationships > New relationship\*\*.*\*\*Sales\*\*.*지역코드.*\*\*Regions\*\*.*\*\*Cardinality: Many to one \(\*:1\)\*\*.*\*\*Cross-filter direction: Single\*\*' $g
    Assert-Match '\*\*Home > New table\*\*' $g
    Assert-Match '```dax\r?\nCalendar = ADDCOLUMNS' $g
    Assert-Match '\*\*Table tools > Mark as date table > Mark as date table\*\*' $g
    Assert-Match '\*\*Column tools > Sort by column\*\* → \*\*MonthNo\*\*' $g
    Assert-Match '\*\*Home > New measure\*\*' $g
    Assert-Match '```dax\r?\n총 매출 = SUM\(Sales\[매출액\]\)\r?\n```' $g
    Assert-Match '\*\*Measure tools > Format\*\*.*`#,0`' $g
}
Test-Case 'New-ManualGuide: report pages with field wells, sort, Top N, size/position' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $g = New-ManualGuide -Model $m -RawDataDir 'C:\work\rawdata'
    Assert-Match '### 페이지 1: `개요`' $g
    Assert-Match '\*\*Insert > Text box\*\*' $g
    Assert-Match '\*\*Card \(new\)\*\*' $g
    Assert-Match '\*\*Data\*\* ← \*\*Apps\*\* 테이블의 측정값 \*\*\[앱 수\]\*\*' $g
    Assert-Match '\*\*Clustered bar chart\*\*' $g
    Assert-Match '\*\*Y-axis\*\* ← \*\*Apps\[장르\]\*\*' $g
    Assert-Match '\*\*X-axis\*\* ← \*\*Apps\*\* 테이블의 측정값 \*\*\[앱 수\]\*\*' $g
    Assert-Match '\*\*More options \(…\)\*\* → \*\*Sort axis\*\* → \*\*\[앱 수\]\*\*, \*\*Sort descending\*\*' $g
    Assert-Match '\*\*Filter type: Top N\*\*, \*\*Show items: Top 3\*\*' $g
    Assert-Match '\*\*X Axis\*\* ← \*\*Apps\[price\]\*\* \(필드 드롭다운 ▾ → \*\*Average\*\*\)' $g
    Assert-Match 'X 40, Y 168, Width 425, Height 235' $g
    Assert-Match '### 페이지 2: `상세`' $g
    Assert-Match '\*\*Filters on this page\*\*.*Apps\[장르\].*\*\*Basic filtering\*\* → `Games`, `Education`' $g
    Assert-Match '\*\*Matrix\*\*' $g
    Assert-Match '\*\*File > Save As\*\*' $g
}
Test-Case 'build-guide.ps1 writes manual-guide.md next to the spec' {
    $dir = Join-Path $tmpRoot 'guide'; New-Item -ItemType Directory -Path $dir -Force | Out-Null
    Copy-Item (Join-Path $fixtures 'spec-flat.json') (Join-Path $dir 'report-spec.json')
    $log = (& $guideScript -Spec (Join-Path $dir 'report-spec.json') -RawData $fixtures | Out-String)
    Assert-Equal 0 $LASTEXITCODE $log
    Assert-Match 'GUIDE .*manual-guide\.md' $log
    $bytes = [IO.File]::ReadAllBytes((Join-Path $dir 'manual-guide.md')); Assert-Equal 0x23 $bytes[0] 'starts with # (no BOM)'
}
```

- [ ] **Step 2: Run to verify it fails**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter guide
```
Expected: fails loading `guide.ps1`.

- [ ] **Step 3: Write `lib/guide.ps1`**

```powershell
# guide.ps1 — manual-guide.md generator. Korean prose, English Power BI Desktop UI terms. Requires io.ps1, spec.ps1, tmdl.ps1.
$ErrorActionPreference = 'Stop'
$script:PqTypeNames = @{ int64 = 'Whole Number'; double = 'Decimal Number'; decimal = 'Fixed decimal number'; text = 'Text'; date = 'Date'; datetime = 'Date/Time'; boolean = 'True/False' }
$script:VisualUiNames = @{ card = 'Card (new)'; multiRowCard = 'Multi-row card'; gauge = 'Gauge'; lineChart = 'Line chart'; areaChart = 'Area chart'
    clusteredColumnChart = 'Clustered column chart'; stackedColumnChart = 'Stacked column chart'; clusteredBarChart = 'Clustered bar chart'; stackedBarChart = 'Stacked bar chart'
    pieChart = 'Pie chart'; donutChart = 'Donut chart'; scatterChart = 'Scatter chart'; tableEx = 'Table'; pivotTable = 'Matrix'; slicer = 'Slicer'; textbox = 'Text box' }
$script:XyWells  = @{ Category = 'X-axis'; Y = 'Y-axis'; Series = 'Legend' }
$script:BarWells = @{ Category = 'Y-axis'; Y = 'X-axis'; Series = 'Legend' }
$script:WellNames = @{
    card = @{ Values = 'Data' }; multiRowCard = @{ Values = 'Fields' }
    gauge = @{ Y = 'Value'; MinValue = 'Minimum value'; MaxValue = 'Maximum value'; TargetValue = 'Target value' }
    lineChart = $script:XyWells; areaChart = $script:XyWells; clusteredColumnChart = $script:XyWells; stackedColumnChart = $script:XyWells
    clusteredBarChart = $script:BarWells; stackedBarChart = $script:BarWells
    pieChart = @{ Category = 'Legend'; Y = 'Values' }; donutChart = @{ Category = 'Legend'; Y = 'Values' }
    scatterChart = @{ Category = 'Values'; X = 'X Axis'; Y = 'Y Axis'; Size = 'Size' }
    tableEx = @{ Values = 'Columns' }; pivotTable = @{ Rows = 'Rows'; Columns = 'Columns'; Values = 'Values' }; slicer = @{ Values = 'Field' }
}
$script:SummarizeUi = @{ none = "Don't summarize"; sum = 'Sum'; average = 'Average'; count = 'Count'; min = 'Minimum'; max = 'Maximum'; distinctCount = 'Count (Distinct)' }
$script:AggUi = @{ sum = 'Sum'; avg = 'Average'; count = 'Count'; countDistinct = 'Count (Distinct)'; min = 'Minimum'; max = 'Maximum' }

function Format-FieldForGuide {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Ref)
    $r = Resolve-FieldRef -Model $Model -Ref $Ref
    switch ($r.Kind) {
        'Measure'     { return "**$($r.Table)** 테이블의 측정값 **[$($r.Name)]**" }
        'Column'      { return "**$($r.Table)[$($r.Name)]**" }
        'Aggregation' { return "**$($r.Table)[$($r.Name)]** (필드 드롭다운 ▾ → **$($script:AggUi[$r.FunctionName])**)" }
    }
}

function New-ManualGuide {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$RawDataDir)
    $l = New-Object System.Collections.Generic.List[string]
    $sourceTables = @($Model.Tables.Keys | Where-Object { -not $Model.Tables[$_].IsCalculated })
    $files = @($sourceTables | ForEach-Object { $Model.Tables[$_].File } | Select-Object -Unique)

    $l.Add("# $($Model.Name) — 수동 제작 가이드 (Power BI Desktop)")
    $l.Add('')
    $l.Add('> 이 문서는 `report-spec.json`에서 자동 생성되었습니다. Claude가 만든 `.pbip`와 같은 결과물을 Power BI Desktop에서 직접 만드는 순서입니다. 메뉴·기능·필드 이름은 Power BI Desktop **영어 UI** 기준입니다.')
    $l.Add('')
    $l.Add('## 0. 준비')
    $l.Add('1. Power BI Desktop을 실행하고 **Blank report**를 선택합니다.')
    $l.Add("2. 데이터 파일 위치: ``$RawDataDir`` — 사용할 파일: " + (($files | ForEach-Object { '`' + $_ + '`' }) -join ', '))
    $l.Add('3. 캔버스 크기(모든 페이지 공통): 캔버스의 빈 곳을 클릭 → **Visualizations** 창의 **Format page** → **Canvas settings** → **Type: Custom**, **Width: 1920**, **Height: 1080**.')
    $l.Add('')

    $l.Add('## 1. 데이터 가져오기 (Get Data → Power Query Editor)')
    $i = 0
    foreach ($tn in $sourceTables) {
        $t = $Model.Tables[$tn]; $i++
        $src = if ($t.IsXlsx) { "``$($t.File)`` / 시트 ``$($t.Sheet)``" } else { "``$($t.File)``" }
        $l.Add("### 1.$i 테이블 ``$tn`` ← $src")
        if ($t.IsXlsx) {
            $l.Add("1. **Home > Get Data > Excel Workbook** → ``$($t.File)`` 선택 → **Open**.")
            $l.Add("2. **Navigator**에서 시트 **$($t.Sheet)**를 체크하고 **Transform Data**를 클릭합니다.")
        } else {
            $origin = if ($t.Encoding -eq 949) { '949: Korean' } else { '65001: Unicode (UTF-8)' }
            $delimName = switch ($t.Delimiter) { ';' { 'Semicolon' } "`t" { 'Tab' } '|' { 'Custom → |' } default { 'Comma' } }
            $l.Add("1. **Home > Get Data > Text/CSV** → ``$($t.File)`` 선택 → **Open**.")
            $l.Add("2. 미리보기 창에서 **File Origin: $origin**, **Delimiter: $delimName** 인지 확인하고 **Transform Data**를 클릭합니다.")
        }
        $l.Add("3. **Power Query Editor**에서 오른쪽 **Query Settings** 창의 **Name**을 ``$tn``으로 바꾼 뒤:")
        if ($t.HeaderRow -gt 1) { $l.Add("   - **Home > Remove Rows > Remove Top Rows** → **Number of rows: $($t.HeaderRow - 1)** → **OK**") }
        $l.Add('   - **Home > Use First Row as Headers** (첫 행이 이미 헤더면 생략)')
        $l.Add('   - **Home > Choose Columns** → 다음 열만 체크: ' + (($t.Columns | ForEach-Object { '`' + $_.Source + '`' }) -join ', '))
        $renames = @($t.Columns | Where-Object { $_.Source -ne $_.Name })
        if ($renames.Count -gt 0) { $l.Add('   - 열 이름 변경(열 헤더 더블클릭): ' + (($renames | ForEach-Object { '`' + $_.Source + '` → `' + $_.Name + '`' }) -join ', ')) }
        $l.Add('   - 데이터 형식(열 헤더 왼쪽의 형식 아이콘 클릭): ' + (($t.Columns | ForEach-Object { '`' + $_.Name + '` → **' + $script:PqTypeNames[$_.Type] + '**' }) -join ', '))
        $l.Add('')
    }
    $l.Add('4. 모든 테이블을 추가한 뒤 **Home > Close & Apply**를 클릭합니다.')
    $l.Add('')

    $l.Add('## 2. 모델 설정 (Model view)')
    $l.Add('왼쪽의 **Model view** 아이콘을 클릭합니다.')
    $l.Add('')
    $l.Add('### 2.1 열 속성 (열을 클릭한 뒤 **Properties** 창)')
    $any = $false
    foreach ($tn in $sourceTables) {
        foreach ($c in $Model.Tables[$tn].Columns) {
            $props = New-Object System.Collections.Generic.List[string]
            if ($c.Hidden) { $props.Add('**Is hidden: On**') }
            if ($c.Type -in 'int64', 'double', 'decimal' -and $c.SummarizeBy -ne 'sum') { $props.Add("**Advanced > Summarize by: $($script:SummarizeUi[$c.SummarizeBy])**") }
            if ($c.Format -and $c.Type -in 'int64', 'double', 'decimal') { $props.Add("**Formatting > Format: Custom** → ``$($c.Format)``") }
            if ($c.SortBy) { $props.Add("**Column tools > Sort by column** → **$($c.SortBy)**") }
            if ($c.Description) { $props.Add("**Description**: $($c.Description)") }
            if ($props.Count -gt 0) { $any = $true; $l.Add("- ``$tn[$($c.Name)]``: " + ($props -join ', ')) }
        }
    }
    if (-not $any) { $l.Add('- 변경할 열 속성이 없습니다.') }
    $l.Add('')
    $userRels = @($Model.Relationships | Where-Object { $_.ToTable -ne 'Calendar' })
    $l.Add('### 2.2 관계 (Relationships)')
    if ($userRels.Count -eq 0) { $l.Add('- 테이블 간 관계가 없습니다 (단일 테이블).') }
    foreach ($r in $userRels) {
        $dir = if ($r.CrossFilter -eq 'both') { 'Both' } else { 'Single' }
        $act = if ($r.Active) { '체크' } else { '체크 해제' }
        $l.Add("- **Home > Manage relationships > New relationship**: From table **$($r.FromTable)** 열 **$($r.FromColumn)** → To table **$($r.ToTable)** 열 **$($r.ToColumn)**, **Cardinality: Many to one (*:1)**, **Cross-filter direction: $dir**, **Make this relationship active** $act → **OK**.")
    }
    $l.Add('')
    if ($null -ne $Model.Calendar) {
        $l.Add('### 2.3 Calendar 테이블 (날짜 테이블)')
        $l.Add('1. **Home > New table** → 수식 입력줄에 아래 DAX를 붙여넣고 Enter:')
        $l.Add('```dax')
        $l.Add('Calendar = ' + (New-CalendarDax -Model $Model))
        $l.Add('```')
        $l.Add('2. **Data** 창에서 **Calendar** 테이블 선택 → **Table tools > Mark as date table > Mark as date table** → **Date column: Date** → **OK**.')
        $l.Add('3. **Calendar[Month]** 열 선택 → **Column tools > Sort by column** → **MonthNo**. **Calendar[MonthNo]**는 **Properties > Is hidden: On**.')
        $n = 4
        foreach ($r in ($Model.Relationships | Where-Object { $_.ToTable -eq 'Calendar' })) {
            $act = if ($r.Active) { '체크' } else { '체크 해제 (비활성 관계)' }
            $l.Add("$n. **Home > Manage relationships > New relationship**: From table **$($r.FromTable)** 열 **$($r.FromColumn)** → To table **Calendar** 열 **Date**, **Cardinality: Many to one (*:1)**, **Cross-filter direction: Single**, **Make this relationship active** $act → **OK**."); $n++
        }
        $l.Add('')
    }

    $l.Add('## 3. 측정값 (Measures)')
    $l.Add('**Report view**로 돌아가 **Data** 창에서 테이블을 선택한 뒤 **Home > New measure**를 클릭하고, 수식 입력줄에 아래 DAX를 붙여넣고 Enter를 누릅니다.')
    $l.Add('')
    $mi = 0
    foreach ($tn in $Model.Tables.Keys) {
        foreach ($m in $Model.Tables[$tn].Measures) {
            $mi++
            $l.Add("### 3.$mi ``$tn`` 테이블 → **New measure**")
            $l.Add('```dax')
            $l.Add("$($m.Name) = $($m.Dax)")
            $l.Add('```')
            if ($m.Format) { $l.Add("- **Measure tools > Format**: **Custom** → ``$($m.Format)``") }
            if ($m.Description) { $l.Add("- 설명: $($m.Description)") }
            $l.Add('')
        }
    }
    if ($mi -eq 0) { $l.Add('- 측정값이 없습니다.'); $l.Add('') }

    $l.Add('## 4. 보고서 페이지 (Report view)')
    $pi = 0
    foreach ($page in $Model.Pages) {
        $pi++
        $pname = [string](Get-Prop $page 'name')
        $l.Add("### 페이지 ${pi}: ``$pname``")
        if ($pi -eq 1) { $l.Add("- 하단의 페이지 탭을 더블클릭 → 이름을 ``$pname``으로 변경.") } else { $l.Add("- 하단의 **+ (New page)** 클릭 → 탭 더블클릭 → 이름을 ``$pname``으로 변경.") }
        foreach ($f in (ConvertTo-Array (Get-Prop $page 'filters'))) {
            $fr = Resolve-FieldRef -Model $Model -Ref ([string](Get-Prop $f 'field'))
            $l.Add("- 페이지 필터: **Filters** 창의 **Filters on this page**에 **$($fr.Table)[$($fr.Name)]**를 드래그 → **Basic filtering** → " + (((ConvertTo-Array (Get-Prop $f 'in')) | ForEach-Object { '`' + $_ + '`' }) -join ', ') + ' 체크.')
        }
        $l.Add('')
        $vi = 0
        foreach ($v in (ConvertTo-Array (Get-Prop $page 'visuals'))) {
            $vi++
            $type = [string](Get-Prop $v 'type'); $ui = $script:VisualUiNames[$type]
            $pos = Get-GridPosition -Grid (ConvertTo-Array (Get-Prop $v 'grid'))
            $title = Get-Prop $v 'title' $null
            $head = "#### 4.$pi.$vi **$ui**"; if ($title) { $head += " — ""$title""" } elseif ($type -eq 'textbox') { $head += " — ""$([string](Get-Prop $v 'text'))""" }
            $l.Add($head)
            if ($type -eq 'textbox') {
                $l.Add("- **Insert > Text box** → 텍스트 입력: ``$([string](Get-Prop $v 'text'))``, 글꼴 크기 **$([int](Get-Prop $v 'fontSize' 20))**, **Bold**.")
            } else {
                $l.Add("- 캔버스 빈 곳 클릭 → **Visualizations** 창에서 **$ui** 클릭.")
                $wells = $script:WellNames[$type]
                $fields = Get-Prop $v 'fields' $null
                foreach ($role in ($script:VisualCatalog[$type].Required + $script:VisualCatalog[$type].Optional)) {
                    $refs = ConvertTo-Array (Get-Prop $fields $role)
                    if ($refs.Count -eq 0) { continue }
                    $l.Add("- **$($wells[$role])** ← " + (($refs | ForEach-Object { Format-FieldForGuide -Model $Model -Ref ([string]$_) }) -join ', '))
                }
                if ($title) { $l.Add("- 제목: **Format visual > General > Title > Text**: ``$title``") }
                $sort = Get-Prop $v 'sort' $null
                if ($null -ne $sort) {
                    $sr = Resolve-FieldRef -Model $Model -Ref ([string](Get-Prop $sort 'field'))
                    $sd = if ([string](Get-Prop $sort 'direction' 'desc') -eq 'asc') { 'Sort ascending' } else { 'Sort descending' }
                    $l.Add("- 정렬: 비주얼 우측 상단 **More options (…)** → **Sort axis** → **[$($sr.Name)]**, **$sd**")
                }
                $topN = Get-Prop $v 'topN' $null
                if ($null -ne $topN) {
                    $catRefs = ConvertTo-Array (Get-Prop $fields 'Category'); if ($catRefs.Count -eq 0) { $catRefs = ConvertTo-Array (Get-Prop $fields 'Rows') }
                    $cr = Resolve-FieldRef -Model $Model -Ref ([string]$catRefs[0])
                    $l.Add("- 상위 N: **Filters** 창 → **Filters on this visual**의 **$($cr.Name)** 카드 펼치기 → **Filter type: Top N**, **Show items: Top $([int](Get-Prop $topN 'n'))**, **By value** ← " + (Format-FieldForGuide -Model $Model -Ref ([string](Get-Prop $topN 'by'))) + ' 드래그 → **Apply filter**')
                }
                foreach ($f in (ConvertTo-Array (Get-Prop $v 'filters'))) {
                    $fr = Resolve-FieldRef -Model $Model -Ref ([string](Get-Prop $f 'field'))
                    $l.Add("- 비주얼 필터: **Filters on this visual**에 **$($fr.Table)[$($fr.Name)]** 드래그 → **Basic filtering** → " + (((ConvertTo-Array (Get-Prop $f 'in')) | ForEach-Object { '`' + $_ + '`' }) -join ', ') + ' 체크')
                }
            }
            $l.Add("- 위치/크기: **Format** > **General > Properties** → **Position**: X $($pos.x), Y $($pos.y), **Size**: Width $($pos.width), Height $($pos.height)")
            $l.Add('')
        }
    }
    $l.Add('## 5. 저장')
    $l.Add("- **File > Save As** → 파일 형식 **Power BI files (*.pbix)** → ``$($Model.Name).pbix``로 저장합니다.")
    $l.Add('')
    return ($l -join "`r`n")
}
```

- [ ] **Step 4: Write `scripts/build-guide.ps1`**

```powershell
<#
.SYNOPSIS  Generates output/<Name>/manual-guide.md (Korean, English UI terms) from report-spec.json.
.EXAMPLE   powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-guide.ps1 -Spec output/Sales/report-spec.json
#>
param([Parameter(Mandatory)][string]$Spec, [string]$RawData = '')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
. (Join-Path $PSScriptRoot 'lib/io.ps1')
. (Join-Path $PSScriptRoot 'lib/spec.ps1')
. (Join-Path $PSScriptRoot 'lib/tmdl.ps1')
. (Join-Path $PSScriptRoot 'lib/guide.ps1')

$root = Get-WorkspaceRoot -ScriptRoot $PSScriptRoot
if (-not $RawData) { $RawData = Join-Path $root 'rawdata' }
if (-not (Test-Path -LiteralPath $Spec)) { Write-Output "ERROR 1: spec file not found: $Spec"; exit 1 }
$specPath = (Get-Item -LiteralPath $Spec).FullName
$specObj = Read-Spec -Path $specPath
$v = Get-SpecValidation -Spec $specObj -RawDataDir $RawData
if ($v.Errors.Count -gt 0) { $i = 0; foreach ($e in $v.Errors) { $i++; Write-Output ("ERROR {0}: {1}" -f $i, $e) }; exit 1 }
$absRaw = (Get-Item -LiteralPath $RawData).FullName
$out = Join-Path (Split-Path -Parent $specPath) 'manual-guide.md'
Write-Utf8File -Path $out -Content (New-ManualGuide -Model $v.Model -RawDataDir $absRaw)
Write-Output ("GUIDE {0}" -f $out)
exit 0
```

- [ ] **Step 5: Add console UTF-8 to the other entry scripts**

In `profile-data.ps1` and `build-pbip.ps1`, insert directly after the `$ErrorActionPreference = 'Stop'` line:

```powershell
[Console]::OutputEncoding = [Text.Encoding]::UTF8
```

- [ ] **Step 6: Run the guide tests and the full suite**

```bash
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1 -Filter guide
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/tests/run-tests.ps1
```
Expected: `4 tests, 0 failed`, then the full suite `0 failed`.

- [ ] **Step 7: Commit**

```bash
git add .claude/skills/powerbi/scripts
git commit -m "feat(scripts): manual-guide generator (Korean prose, English Power BI UI terms)"
```

---

### Task 13: Reference docs, example spec, SKILL.md, README, project settings

**Files:**
- Create: `.claude/skills/powerbi/reference/spec-schema.md`, `reference/visual-catalog.md`, `reference/dax-patterns.md`
- Create: `.claude/skills/powerbi/examples/app-data-spec.json`
- Create: `.claude/skills/powerbi/SKILL.md`
- Create: `README.md`, `.claude/settings.json`

- [ ] **Step 1: Write `reference/spec-schema.md`**

````markdown
# report-spec.json — contract

One JSON file per report at `output/<Name>/report-spec.json`. The builder validates it and emits the PBIP. Keys marked * are required.

```jsonc
{
  "name*": "SalesDashboard",        // letters/digits/_/- ; becomes the .pbip name
  "locale": "ko-KR",                // default en-US
  "autoDateTable": true,            // add a Calendar table for date columns not linked to a date dimension
  "tables*": [{
    "name*": "Sales",               // model table name; no "."; "Calendar" is reserved
    "file*": "sales.csv",           // file inside rawdata/
    "sheet": "Data",                // xlsx only (required for xlsx)
    "headerRow": 1,                 // from the profile: header_row
    "encoding": 65001,              // csv only: 65001 or 949, from the profile
    "delimiter": ",",               // csv only
    "columns*": [{                  // ONLY listed columns are loaded; list the ones the report needs
      "name*": "track_name",        // source header exactly as in the profile
      "rename": "앱 이름",           // model name (optional)
      "type*": "text",              // int64 | double | decimal | text | date | datetime | boolean
      "key": true,                  // unique id column (also sets summarizeBy none)
      "hidden": true,               // hide keys / technical columns
      "format": "#,0",              // numeric format string
      "summarizeBy": "average",     // none|sum|average|count|min|max|distinctCount (default: sum for numbers, else none)
      "description": "Track Name",  // from the dictionary sheet if any
      "sortBy": "MonthNo"           // sort this column by another column of the same table
    }]
  }],
  "relationships": [{ "from*": "Sales.CityKey", "to*": "City.CityKey", "active": true, "crossFilter": "single" }],
  "measures": [{ "table*": "Sales", "name*": "매출", "dax*": "SUM(Sales[Amount])", "format": "#,0", "description": "" }],
  "pages*": [{
    "name*": "개요",
    "filters": [{ "field": "Sales.Region", "in": ["Seoul", "Busan"] }],
    "visuals*": [{
      "type*": "card",              // see visual-catalog.md
      "title": "총 매출",
      "fields": { "Values": ["Sales.매출"] },   // role → list of field refs
      "grid*": [0, 0, 3, 2],        // [col,row,w,h] on a 12x8 grid (1920x1080 page)
      "sort": { "field": "Sales.매출", "direction": "desc" },
      "topN": { "n": 10, "by": "Sales.매출" },   // needs a Category (or Rows) field
      "filters": [{ "field": "Sales.Region", "in": ["Seoul"] }],
      "text": "제목", "fontSize": 24                // textbox only
    }]
  }]
}
```

**Field refs:** `Table.Name` — the builder checks measures first, then columns (use the *renamed* name). Aggregate a raw column with a prefix: `sum:` `avg:` `min:` `max:` `count:` `countDistinct:` (columns only, never measures). Calendar fields when `autoDateTable` produced one: `Calendar.Date`, `Calendar.Year`, `Calendar.Quarter`, `Calendar.Month`, `Calendar.YearMonth`.

**Grid math:** 12 columns × 8 rows, 40 px margin, 20 px gutter. Common shapes: KPI card `[c,r,3,2]`, half-width chart `[c,r,6,4]`, full-width chart `[0,r,12,4]`, slicer `[c,r,3,2]`, title textbox `[0,0,12,1]`.

**DAX inside `dax`:** reference model names: `Sales[매출액]`, `'Table with space'[Col]`, `[OtherMeasure]`. Use `\n` for multi-line.
````

- [ ] **Step 2: Write `reference/visual-catalog.md`**

```markdown
# Visual catalog (v1)

| type | roles (required → optional) | use it for |
|---|---|---|
| `card` | Values(1) | one KPI number (총 매출, 앱 수) |
| `multiRowCard` | Values(n) | several small KPIs in one box |
| `gauge` | Y(1) → MinValue, MaxValue, TargetValue | progress toward a target |
| `lineChart` / `areaChart` | Category(1), Y(n) → Series(1) | trend over time (Category = Calendar.Month / YearMonth / Year) |
| `clusteredColumnChart` / `stackedColumnChart` | Category(1), Y(n) → Series(1) | compare few categories; stacked for composition |
| `clusteredBarChart` / `stackedBarChart` | Category(1), Y(n) → Series(1) | ranking of many categories (add `sort` desc + `topN`) |
| `pieChart` / `donutChart` | Category(1), Y(1) | share of ≤ 6 categories |
| `scatterChart` | Category(1), X(1), Y(1) → Size(1) | correlation; Category = the entity (product, app) |
| `tableEx` | Values(n) | detail rows |
| `pivotTable` | Rows(n), Values(n) → Columns(n) | matrix / cross-tab |
| `slicer` | Values(1) | user filter (category, year) |
| `textbox` | — (`text`, `fontSize`) | page title / notes |

Rules of thumb:
- Y / Values / X / Size take **measures or aggregated columns** (`sum:`), Category / Series / Rows / Columns / slicer take **columns**.
- Page layout default: row 0 textbox title `[0,0,12,1]`; row 1 KPI cards `[0,1,3,2] [3,1,3,2] [6,1,3,2] [9,1,3,2]`; rows 3–7 two charts `[0,3,6,5] [6,3,6,5]` or one wide chart + slicers.
- Max ~8 visuals per page, 1–4 pages. Use `topN` for bar charts over > 15 categories.
- Korean titles for every visual unless the user writes in English.
```

- [ ] **Step 3: Write `reference/dax-patterns.md`**

````markdown
# DAX measure recipes

Replace `T`, `[Amount]`, `[Qty]`, `[Cost]` with model names. `Calendar` exists when `autoDateTable` created it.

| purpose | dax |
|---|---|
| total | `SUM(T[Amount])` |
| row count | `COUNTROWS(T)` |
| distinct count | `DISTINCTCOUNT(T[CustomerId])` |
| average | `AVERAGE(T[Rating])` |
| profit | `[매출] - [원가]` (measures referenced with `[ ]`) |
| margin % | `DIVIDE([이익], [매출])` → format `0.0%` |
| share of total | `DIVIDE([매출], CALCULATE([매출], ALL(T)))` → `0.0%` |
| average per row | `DIVIDE([매출], [건수])` |
| year-to-date | `TOTALYTD([매출], Calendar[Date])` |
| previous year | `CALCULATE([매출], SAMEPERIODLASTYEAR(Calendar[Date]))` |
| YoY % | `DIVIDE([매출] - [전년 매출], [전년 매출])` → `0.0%` |
| running total | `CALCULATE([매출], FILTER(ALL(Calendar[Date]), Calendar[Date] <= MAX(Calendar[Date])))` |
| top-N rank | `RANKX(ALL(T[Product]), [매출])` |
| count where | `CALCULATE(COUNTROWS(T), T[Status] = "Done")` |
| conditional total | `CALCULATE([매출], T[Price] > 0)` |

Format strings: integers `#,0`; currency `$#,0` or `₩#,0`; decimals `0.00`; percent `0.0%`.
Multi-line DAX is fine in the spec (`"dax": "VAR x = ...\nRETURN ..."`).
````

- [ ] **Step 4: Write `examples/app-data-spec.json`**

Copy `scripts/tests/fixtures/spec-flat.json` and change: `"name": "AppStoreDashboard"`, `"file": "App Data.xlsx"`, add columns `{ "name": "sup_devices", "type": "int64" }` and `{ "name": "lang", "type": "int64" }` after `prime_genre`, remove the `release_date` column (the real file has none), set `"autoDateTable": false`, and replace the two Calendar-based visuals (lineChart on page 1; pivotTable's `Columns` on page 2) with: lineChart → `{ "type": "clusteredColumnChart", "title": "장르별 평균 평점", "fields": { "Category": ["Apps.장르"], "Y": ["Apps.평균 평점"] }, "sort": { "field": "Apps.평균 평점", "direction": "desc" }, "grid": [0, 3, 6, 3] }` and pivotTable → `"fields": { "Rows": ["Apps.장르"], "Values": ["Apps.앱 수", "Apps.평균 가격", "Apps.평균 평점"] }`. Verify with the VM: `build-pbip.ps1 -Spec .claude/skills/powerbi/examples/app-data-spec.json -OutputDir <tmp>` must print `BUILT`.

- [ ] **Step 5: Write `SKILL.md`**

````markdown
---
name: powerbi
description: Build a Power BI Desktop project (.pbip) from the CSV/XLSX files in rawdata/ based on the user's description, preferences or mockup image. Use for ANY request about dashboards, reports, charts, KPIs, data analysis, visualization, or Power BI. Talk to the user in Korean.
---

# Power BI report builder

You turn `rawdata/*.csv|xlsx` + the user's wishes into `output/<Name>/<Name>.pbip`. Two PowerShell scripts do the heavy lifting; you only read a compact profile and write one small spec. **The user is non-technical and on a limited token budget: follow the token rules exactly.**

## Token rules (hard)
- NEVER open files in `rawdata/`, `output/*/`, or `templates/` with Read/cat/head. Scripts read them.
- Run the profiler once per session (again only if the user says the data changed).
- Read `reference/*.md` only when the trigger below applies. Do not read `scripts/`.
- Write the spec once with Write; afterwards change it only with small Edits. Do not re-read it.
- At most 2 clarifying questions, only when the answer changes the result. Otherwise decide, and list assumptions at the end.
- On a Power BI error pasted by the user: reason from the error text + spec; fix with an Edit; rebuild. Never regenerate from scratch.
- Keep replies short. No progress narration.

## Workflow
1. **PROFILE** — run:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/profile-data.ps1`
   Output = files → tables (columns, type, nulls, distinct, samples, KEY flags), dictionary sheets, csv `encoding`, `header_row`, and `KEY MATCHES` for star schemas. Exit code 2 = no files → ask the user (Korean) to put files in `rawdata/`.
2. **UNDERSTAND** — read the user's text; view an attached mockup image once and note: pages, visual types, KPIs, layout, colors are ignored (theme is fixed). Map their words to columns from the profile. If a dictionary sheet exists, use its descriptions for column `description` and for choosing Korean names.
3. **DESIGN** — decide:
   - *Schema:* one table candidate → **flat** (no relationships). Several tables with `KEY MATCHES` → **star**: facts = big tables with numeric measures, dims = tables with a KEY; relationships many→one from fact key to dim key. Independent tables with no matches → separate tables, no relationships.
   - *Columns:* list only what the report needs + keys. Drop `(unnamed)`/index columns. `hidden: true` for keys. Types from the profile (`decimal` for money, `double` for ratios/ratings).
   - *Measures:* 4–10. Always a count (`COUNTROWS`) and totals/averages of the main numeric columns; add ratio measures where sensible. Korean names unless data is English. Trigger → read `reference/dax-patterns.md` only for time intelligence / ranking / running totals.
   - *Pages/visuals:* 1–4 pages, ≤ 8 visuals each, following the mockup if given; otherwise: title textbox, 3–4 KPI cards, 2–3 charts, 1–2 slicers on page 1; details (table/matrix) on page 2. Trigger → read `reference/visual-catalog.md` when unsure which visual or which role names to use.
   - `autoDateTable: true` when any date column exists; then use `Calendar.YearMonth` / `Calendar.Month` / `Calendar.Year` as time axes.
4. **WRITE** `output/<Name>/report-spec.json` (Name = short English PascalCase). Format in `reference/spec-schema.md` — read it the first time you write a spec in a session. Keep ≤ 200 lines.
5. **BUILD** — run:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-pbip.ps1 -Spec "output/<Name>/report-spec.json"`
   `ERROR n:` lines name the spec path (e.g. `pages[0].visuals[2].fields.Y[0]`) → Edit that spot, rebuild. `WARN:` lines are informational. Success prints `BUILT … (tables= measures= pages= visuals=)`.
6. **HANDOFF** (Korean, ≤ 12 lines): what was built (pages → visuals, measures), assumptions, then:
   > `output/<Name>/<Name>.pbip` 을 더블클릭해 Power BI Desktop에서 열고 **Refresh**(새로 고침)를 누른 뒤, **File > Save As**로 `.pbix`로 저장하세요. 수정하고 싶은 점을 말씀해 주시면 바로 반영합니다.
7. **GUIDE?** — ask: "같은 리포트를 Power BI Desktop에서 직접 만드는 방법을 정리한 가이드(manual-guide.md)도 만들어 드릴까요?" If yes, run:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-guide.ps1 -Spec "output/<Name>/report-spec.json"`
   and reply with the path `output/<Name>/manual-guide.md` (Korean prose, English menu names). Do not read the guide.

## Iterating
"X를 바꿔줘" → Edit the spec (visual type / fields / grid / measure / page) → rebuild → one-line Korean confirmation. Remind them once that rebuilding overwrites the generated project, so Desktop-side edits should wait until they are happy.

## Quick reference
- Field ref: `Table.Name` (renamed names), aggregation prefix `sum:|avg:|min:|max:|count:|countDistinct:` on columns.
- Roles: card→Values(1); charts→Category,Y[,Series]; scatter→Category,X,Y[,Size]; tableEx→Values; pivotTable→Rows,Values[,Columns]; slicer→Values(1); gauge→Y[,MinValue,MaxValue,TargetValue]; textbox→`text`.
- Grid `[col,row,w,h]` on 12×8. Title `[0,0,12,1]`, KPI `[c,1,3,2]`, chart `[c,3,6,5]`, slicer `[c,r,3,2]`.
- Types: int64 | double | decimal | text | date | datetime | boolean. csv `encoding` 65001 or 949 from the profile.
````

- [ ] **Step 6: Write `README.md` (Korean)**

````markdown
# theo-powerbi — Claude로 Power BI 리포트 만들기

CSV/Excel 파일을 넣고, 원하는 대시보드를 말(또는 그림)로 설명하면 Claude가 Power BI Desktop에서 바로 열리는 프로젝트(`.pbip`)를 만들어 줍니다.

## 준비물 (Windows)
1. **Power BI Desktop** — Microsoft Store 또는 [다운로드](https://powerbi.microsoft.com/desktop/). 설치 후 한 번 실행해 **File > Options and settings > Options > Preview features**에서 **Power BI Project (.pbip) save option**과 **Store reports using enhanced metadata format (PBIR)** 항목이 있으면 켜고 재시작합니다 (없으면 이미 기본 기능입니다).
2. **Claude Code Desktop** — 로그인 후 사용.
3. 이 저장소 — GitHub에서 **Code > Download ZIP** 후 압축 해제 (또는 `git clone`).

## 사용법 (5단계)
1. `rawdata/` 폴더에 `.csv` 또는 `.xlsx` 파일을 넣습니다. (여러 개 가능, Excel은 시트별로 인식)
2. Claude Code Desktop에서 이 폴더를 엽니다.
3. 채팅에 원하는 내용을 씁니다. 예: `rawdata의 판매 데이터로 월별 매출 추이, 지역별 매출 순위, 상위 10개 제품이 보이는 대시보드 만들어줘`. 참고할 화면 캡처가 있으면 이미지를 함께 붙여 넣으세요.
4. Claude가 `output/<이름>/<이름>.pbip`를 만듭니다. 파일을 더블클릭해 Power BI Desktop에서 열고 **Refresh**를 누른 뒤 **File > Save As**로 `.pbix`로 저장합니다.
5. 고치고 싶은 점을 채팅으로 말하면 Claude가 다시 만들어 줍니다. 마지막에 "수동 제작 가이드"를 원하면 `output/<이름>/manual-guide.md`도 받을 수 있습니다.

## 데이터를 바꾸고 싶을 때
- `rawdata/`는 여러분의 폴더입니다. 파일을 자유롭게 넣고, 바꾸고, 지우세요.
- **같은 파일 이름·시트 이름·열 구성**으로 내용만 바뀌었다면 Power BI Desktop에서 **Refresh**만 누르면 됩니다.
- 파일을 추가하거나 열이 바뀌었다면 Claude에게 "데이터 파일을 바꿨어요"라고 말하세요. 다시 분석해서 새로 만들어 줍니다.

## 주의
- Claude가 다시 만들면 `output/<이름>/` 안의 프로젝트가 덮어써집니다. **Power BI Desktop에서 직접 꾸미는 작업은 Claude와의 수정이 끝난 뒤** `.pbix`로 저장한 파일에서 하세요.
- Power BI Desktop에서 오류가 나면 오류 문구를 그대로 복사해 Claude에게 붙여 넣으세요.
- `rawdata/`와 `output/`은 Git에 올라가지 않습니다 (회사 데이터 보호).

## 폴더 구조
```
rawdata/      ← 데이터 파일 넣는 곳
output/       ← 생성된 Power BI 프로젝트
.claude/skills/powerbi/   ← Claude용 스킬 (수정하지 마세요)
```
````

- [ ] **Step 7: Write `.claude/settings.json` (committed; reduces permission prompts for coworkers)**

```json
{
  "permissions": {
    "allow": [
      "Bash(powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/profile-data.ps1*)",
      "Bash(powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-pbip.ps1*)",
      "Bash(powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-guide.ps1*)"
    ]
  }
}
```

- [ ] **Step 8: Verify SKILL.md loads and the example builds**

```bash
head -5 .claude/skills/powerbi/SKILL.md && wc -l .claude/skills/powerbi/SKILL.md   # frontmatter present, < 150 lines
bash /Users/theo/Desktop/Claude/powerbi/dev/vm-ps.sh .claude/skills/powerbi/scripts/build-pbip.ps1 -Spec '\\Mac\Home\Desktop\Claude\powerbi\skills\.claude\skills\powerbi\examples\app-data-spec.json' -OutputDir 'C:\Users\theo\pbi-example'
```
Expected: `BUILT C:\Users\theo\pbi-example\AppStoreDashboard.pbip (tables=1 measures=3 pages=2 visuals=11)`.

- [ ] **Step 9: Commit**

```bash
git add README.md .claude/settings.json .claude/skills/powerbi/SKILL.md .claude/skills/powerbi/reference .claude/skills/powerbi/examples
git commit -m "docs: SKILL.md, Korean README, reference docs, example spec, project permissions"
```

---

### Task 14: Acceptance — open the generated project in Power BI Desktop on the VM

**Files:** none new in the repo (dogfooding produces `output/AppStoreDashboard/` locally, which is gitignored).

- [ ] **Step 1: Mirror the repo to a local Windows path and build there**

```bash
ssh -q -i ~/.ssh/parallels_win11 -o BatchMode=yes theo@10.211.55.3 'robocopy \\Mac\Home\Desktop\Claude\powerbi\skills C:\Users\theo\theo-powerbi /MIR /XD .git output /NFL /NDL /NJH /NJS' ; echo "robocopy exit $?"   # exit codes < 8 are success
ssh -q -i ~/.ssh/parallels_win11 -o BatchMode=yes theo@10.211.55.3 'powershell -NoProfile -ExecutionPolicy Bypass -Command "[Console]::OutputEncoding=[Text.Encoding]::UTF8; cd C:\Users\theo\theo-powerbi; New-Item -ItemType Directory -Force output\AppStoreDashboard | Out-Null; Copy-Item .claude\skills\powerbi\examples\app-data-spec.json output\AppStoreDashboard\report-spec.json; & .claude\skills\powerbi\scripts\build-pbip.ps1 -Spec output\AppStoreDashboard\report-spec.json"'
```
Expected: `BUILT C:\Users\theo\theo-powerbi\output\AppStoreDashboard\AppStoreDashboard.pbip (...)`. The M query now contains `C:\Users\theo\theo-powerbi\rawdata\App Data.xlsx`.

- [ ] **Step 2: Open it in Power BI Desktop**

```bash
ssh -q -i ~/.ssh/parallels_win11 -o BatchMode=yes theo@10.211.55.3 'powershell -NoProfile -Command "Start-Process C:\Users\theo\theo-powerbi\output\AppStoreDashboard\AppStoreDashboard.pbip; Start-Sleep 90; Get-Process PBIDesktop -ErrorAction SilentlyContinue | Select-Object Id, MainWindowTitle | Format-List"'
```
Expected: a `PBIDesktop` process with `MainWindowTitle` containing `AppStoreDashboard`. If the process is gone, Desktop rejected the project: run `Get-ChildItem $env:LOCALAPPDATA\Microsoft\Power BI Desktop\Traces -Filter *.log | Sort-Object LastWriteTime | Select-Object -Last 1 | Get-Content -Tail 60` on the VM to read the error, fix the emitter (usually a TMDL property name or a PBIR object path), add a regression assertion to the relevant test, rebuild, retry.

- [ ] **Step 3: Verify the model through the powerbi-modeling MCP (from this Claude Code session)**

Use `mcp__powerbi-modeling__connection_operations` to list local Power BI Desktop instances and connect to the `AppStoreDashboard` one; then `table_operations` (list) must show `Apps`; `measure_operations` (list) must show `앱 수`, `평균 평점`, `평균 가격`; run `dax_query_operations` with `EVALUATE ROW("n", [앱 수])` — expect 7197. Any refresh error surfaces here as a failed connect/query: read it and fix as in Step 2.

- [ ] **Step 4: Take a screenshot as evidence (optional but cheap)**

```bash
ssh -q -i ~/.ssh/parallels_win11 -o BatchMode=yes theo@10.211.55.3 'powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms,System.Drawing; $b=[Windows.Forms.Screen]::PrimaryScreen.Bounds; $bmp=New-Object Drawing.Bitmap $b.Width,$b.Height; $g=[Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen($b.Location,[Drawing.Point]::Empty,$b.Size); $bmp.Save(\"\\\\Mac\\Home\\Desktop\\Claude\\powerbi\\dev\\acceptance.png\")"'
```
Then Read `/Users/theo/Desktop/Claude/powerbi/dev/acceptance.png` and confirm the report canvas shows the KPI cards and charts with data (not error icons).

- [ ] **Step 5: Close Desktop and record the result**

```bash
ssh -q -i ~/.ssh/parallels_win11 -o BatchMode=yes theo@10.211.55.3 'powershell -NoProfile -Command "Get-Process PBIDesktop -ErrorAction SilentlyContinue | Stop-Process -Force"'
```
Append a short "Acceptance 2026-09-01: opened in Power BI Desktop <version>, refreshed, 7197 rows, 11 visuals rendered" line to the bottom of `docs/superpowers/specs/2026-09-01-powerbi-skill-design.md` and commit: `git commit -am "docs: acceptance result"`.

---

### Task 15: Push to GitHub

- [ ] **Step 1: Try the push**

```bash
cd /Users/theo/Desktop/Claude/powerbi/skills && git status --short && git log --oneline | head -20 && git push -u origin main
```
Expected: `main -> main` pushed. If it fails with 403 (the active `gh` account KUMHWA-DEV has read-only access), try `gh auth switch --user bosangdev-beep && git push -u origin main`; if that also fails, stop and tell the user to run `gh auth login` as `kyuuto3-coder` (or add KUMHWA-DEV as a collaborator on `kyuuto3-coder/theo-powerbi`) and then `git push -u origin main`.

- [ ] **Step 2: Verify on GitHub**

```bash
gh repo view kyuuto3-coder/theo-powerbi --json defaultBranchRef,isEmpty && gh api repos/kyuuto3-coder/theo-powerbi/contents/.claude/skills/powerbi/SKILL.md --jq .size
```
Expected: `isEmpty: false`, default branch `main`, SKILL.md size > 0.

---

## Self-review notes

- **Spec coverage:** §2 layout → Tasks 1, 13; §3 README → Task 13; §4 workflow + token rules → Task 13 (SKILL.md); §5 spec contract → Tasks 8, 13; §6 profiler → Tasks 4–7; §7 modes → Task 8 (Get-SpecModel) + SKILL.md guidance; §8 builder → Tasks 9–11; §9 testing → every task + Task 14; §10 guide → Task 12; §12 distribution → Tasks 1, 15.
- **Known deviation from spec §2:** visual skeletons are an in-code catalog (`lib/spec.ps1`) instead of `templates/visuals/*.json` — already reflected in the spec's layout tree.
- **Names used consistently:** `Get-Prop`, `ConvertTo-Array`, `Write-Utf8File`, `New-HexId`, `New-LineageTag`, `ConvertTo-JsonFile`, `Get-WorkspaceRoot` (io); `Read-XlsxWorkbook`, `ConvertFrom-ColumnLetters` (xlsx); `Get-CsvEncoding`, `Get-CsvDelimiter`, `Read-CsvRows` (csvread); `Get-ValueKind`, `Find-HeaderRow`, `Get-ColumnProfile`, `Test-DictionarySheet`, `Find-KeyMatches` (infer); `Read-Spec`, `Get-SpecModel`, `Resolve-FieldRef`, `Get-SpecValidation`, `Get-GridPosition`, `$script:VisualCatalog` (spec); `New-MQuery`, `New-TableTmdl`, `New-ModelTmdl`, `New-RelationshipsTmdl`, `New-CalendarDax`, `Write-SemanticModel`, `Format-TmdlName` (tmdl); `ConvertTo-EntityExpr`, `New-Projection`, `ConvertTo-FilterLiteral`, `New-InFilter`, `New-TopNFilter`, `New-VisualJson`, `Write-Report` (pbir); `New-ManualGuide`, `Format-FieldForGuide` (guide).
