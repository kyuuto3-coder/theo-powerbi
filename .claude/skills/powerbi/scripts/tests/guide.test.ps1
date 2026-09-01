. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'spec.ps1')
. (Join-Path $libDir 'tmdl.ps1')
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
    Assert-Match 'X 40, Y 168, \*\*Size\*\*: Width 445, Height 235' $g
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
    Assert-Match 'GUIDE .*manual-guide\.html' $log
    Assert-Match 'GUIDE-MD .*manual-guide\.md' $log
    $bytes = [IO.File]::ReadAllBytes((Join-Path $dir 'manual-guide.md')); Assert-Equal 0x23 $bytes[0] 'starts with # (no BOM)'
    $html = [IO.File]::ReadAllText((Join-Path $dir 'manual-guide.html'), [Text.Encoding]::UTF8)
    Assert-Match '^<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">' $html
}

Test-Case 'New-ManualGuide: folder combine and composite key steps' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-advanced.json'))
    $g = New-ManualGuide -Model $m -RawDataDir 'C:\work\rawdata'
    Assert-Match '### 1\.1 테이블 `Orders` ← `orders_\*\.csv` \(여러 파일\)' $g
    Assert-Match '\*\*Home > Get Data > Folder\*\*' $g
    Assert-Match '\*\*Text Filters > Begins With…\*\* `orders_`' $g
    Assert-Match '\*\*Combine Files\*\*' $g
    Assert-Match '\*\*Add Column > Custom Column\*\* → \*\*New column name\*\* `_key_Month_Region Code`, \*\*Custom column formula\*\* `= Text\.Combine\(\{Text\.From\(\[Month\]\), Text\.From\(\[Region Code\]\)\}, "\|"\)`' $g
    Assert-True (-not ($g -match 'Choose Columns\*\* → 다음 열만 체크: [^\n]*_key_')) 'key column is not in Choose Columns'
}

Test-Case 'New-ManualGuideHtml: checkboxes, kbd menus, copy buttons, minimap with one rect per visual, field wells' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $g = New-ManualGuideHtml -Model $m -RawDataDir 'C:\work\rawdata'
    Assert-Match '<title>AppsFlat — 수동 제작 가이드</title>' $g
    Assert-True (([regex]::Matches($g, 'input type="checkbox" data-step="s\d+"')).Count -gt 30) 'many steps'
    Assert-Match '<kbd>Home</kbd><span class="sep">▸</span><kbd>Get Data</kbd><span class="sep">▸</span><kbd>Excel Workbook</kbd>' $g
    Assert-Match '<button class="copy" type="button">복사</button><pre><code>앱 수 = COUNTROWS\(Apps\)</code></pre>' $g
    Assert-Equal 11 ([regex]::Matches($g, '<rect x="\d+" y="\d+" width="\d+" height="\d+" rx="14"')).Count 'one rect per visual (11)'
    Assert-Equal 2 ([regex]::Matches($g, '<svg class="minimap"')).Count 'one minimap per page'
    Assert-Match '<div class="well"><span class="wn">Y-axis</span><b>Apps\[장르\]</b></div>' $g
    Assert-Match '<div class="well"><span class="wn">X-axis</span><b>\[앱 수\]</b> <small>\(Apps 측정값\)</small></div>' $g
    Assert-Match '<span class="wn">X Axis</span><b>Apps\[price\]</b> <small>→ 필드 ▾ <kbd>Average</kbd></small>' $g
    Assert-Match 'Filter type: Top N</kbd> <kbd>Show items: Top 3</kbd>' $g
    Assert-Match 'localStorage' $g
    Assert-True (-not ($g -match '<script[^>]*src=')) 'self-contained (no external scripts)'
    Assert-True (-not ($g -match '&lt;kbd&gt;')) 'menus are not double-escaped'
}
