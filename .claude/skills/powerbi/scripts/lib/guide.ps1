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
    $l.Add("2. 데이터 파일 위치: ``$RawDataDir`` — 사용할 파일: " + (@($files | ForEach-Object { '`' + $_ + '`' }) -join ', '))
    $l.Add('3. 캔버스 크기(모든 페이지 공통): 캔버스의 빈 곳을 클릭 → **Visualizations** 창의 **Format page** → **Canvas settings** → **Type: Custom**, **Width: 1920**, **Height: 1080**.')
    $l.Add('')

    $l.Add('## 1. 데이터 가져오기 (Get Data → Power Query Editor)')
    $i = 0
    foreach ($tn in $sourceTables) {
        $t = $Model.Tables[$tn]; $i++
        $srcCols = @($t.Columns | Where-Object { $null -ne $_.Source }); $compCols = @($t.Columns | Where-Object { $null -eq $_.Source -and $_.Composite })
        $fileLabel = if ($t.FilePattern) { "``$($t.FilePattern)`` (여러 파일)" } else { "``$($t.File)``" }
        $src = if ($t.IsXlsx) { "$fileLabel / 시트 ``$($t.Sheet)``" } else { $fileLabel }
        $l.Add("### 1.$i 테이블 ``$tn`` ← $src")
        if ($t.FilePattern) {
            $pp = ([string]$t.FilePattern).Split('*')
            $l.Add("1. **Home > Get Data > Folder** → 폴더 경로에 ``$RawDataDir`` 입력 → **OK**.")
            $l.Add("2. 파일 목록 창에서 **Transform Data** → **Power Query Editor**에서 **Name** 열의 필터 ▾ → **Text Filters > Begins With…** ``$($pp[0])`` 로 ``$($t.FilePattern)`` 파일만 남깁니다.")
            $sampleNote = if ($t.IsXlsx) { "시트 **$($t.Sheet)** 선택" } else { "**File Origin**/**Delimiter** 확인" }
            $l.Add("3. **Content** 열 머리글의 **Combine Files** 버튼 클릭 → 샘플 파일 창에서 $sampleNote → **OK**. (같은 구조의 파일이 하나의 표로 합쳐지며, 나중에 같은 이름 규칙의 파일을 폴더에 넣으면 **Refresh**만으로 반영됩니다.)")
        } elseif ($t.IsXlsx) {
            $l.Add("1. **Home > Get Data > Excel Workbook** → ``$($t.File)`` 선택 → **Open**.")
            $l.Add("2. **Navigator**에서 시트 **$($t.Sheet)**를 체크하고 **Transform Data**를 클릭합니다.")
        } else {
            $origin = if ($t.Encoding -eq 949) { '949: Korean' } else { '65001: Unicode (UTF-8)' }
            $delimName = switch ($t.Delimiter) { ';' { 'Semicolon' } "`t" { 'Tab' } '|' { 'Custom → |' } default { 'Comma' } }
            $l.Add("1. **Home > Get Data > Text/CSV** → ``$($t.File)`` 선택 → **Open**.")
            $l.Add("2. 미리보기 창에서 **File Origin: $origin**, **Delimiter: $delimName** 인지 확인하고 **Transform Data**를 클릭합니다.")
        }
        $stepNo = if ($t.FilePattern) { 4 } else { 3 }
        $l.Add("$stepNo. **Power Query Editor**에서 오른쪽 **Query Settings** 창의 **Name**을 ``$tn``으로 바꾼 뒤:")
        if ($t.HeaderRow -gt 1) { $l.Add("   - **Home > Remove Rows > Remove Top Rows** → **Number of rows: $($t.HeaderRow - 1)** → **OK**") }
        $l.Add('   - **Home > Use First Row as Headers** (첫 행이 이미 헤더면 생략)')
        $l.Add('   - **Home > Choose Columns** → 다음 열만 체크: ' + (@($srcCols | ForEach-Object { '`' + $_.Source + '`' }) -join ', '))
        $renames = @($srcCols | Where-Object { $_.Source -cne $_.Name })
        if ($renames.Count -gt 0) { $l.Add('   - 열 이름 변경(열 헤더 더블클릭): ' + (@($renames | ForEach-Object { '`' + $_.Source + '` → `' + $_.Name + '`' }) -join ', ')) }
        $l.Add('   - 데이터 형식(열 헤더 왼쪽의 형식 아이콘 클릭): ' + (@($srcCols | ForEach-Object { '`' + $_.Name + '` → **' + $script:PqTypeNames[$_.Type] + '**' }) -join ', '))
        foreach ($c in $compCols) {
            $formula = 'Text.Combine({' + (@($c.Composite | ForEach-Object { 'Text.From([' + $_ + '])' }) -join ', ') + '}, "|")'
            $l.Add("   - 복합 키 열 만들기: **Add Column > Custom Column** → **New column name** ``$($c.Name)``, **Custom column formula** ``= $formula`` → **OK** → 형식 **Text**. (관계는 열 하나끼리만 맺을 수 있어 " + ($c.Composite -join ' + ') + '를 합친 키가 필요합니다.)')
        }
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
        foreach ($r in @($Model.Relationships | Where-Object { $_.ToTable -eq 'Calendar' })) {
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
            $l.Add("- 페이지 필터: **Filters** 창의 **Filters on this page**에 **$($fr.Table)[$($fr.Name)]**를 드래그 → **Basic filtering** → " + (@((ConvertTo-Array (Get-Prop $f 'in')) | ForEach-Object { '`' + $_ + '`' }) -join ', ') + ' 체크.')
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
                    $l.Add("- **$($wells[$role])** ← " + (@($refs | ForEach-Object { Format-FieldForGuide -Model $Model -Ref ([string]$_) }) -join ', '))
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
                    $l.Add("- 비주얼 필터: **Filters on this visual**에 **$($fr.Table)[$($fr.Name)]** 드래그 → **Basic filtering** → " + (@((ConvertTo-Array (Get-Prop $f 'in')) | ForEach-Object { '`' + $_ + '`' }) -join ', ') + ' 체크')
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

# ======================= HTML guide (beginner-friendly) =======================
function ConvertTo-HtmlText { param([AllowEmptyString()][string]$S) return [System.Net.WebUtility]::HtmlEncode($S) }
function Format-MenuHtml {
    # 'Home > Get Data > Excel Workbook' → <kbd>Home</kbd> ▸ <kbd>Get Data</kbd> ▸ ...
    param([string]$Path)
    return (@($Path.Split('>') | ForEach-Object { '<kbd>' + (ConvertTo-HtmlText $_.Trim()) + '</kbd>' }) -join '<span class="sep">▸</span>')
}
function New-CodeHtml { param([string]$Code) return '<div class="code"><button class="copy" type="button">복사</button><pre><code>' + (ConvertTo-HtmlText $Code) + '</code></pre></div>' }
function New-FieldHtml {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Ref)
    $r = Resolve-FieldRef -Model $Model -Ref $Ref
    switch ($r.Kind) {
        'Measure'     { return '<b>[' + (ConvertTo-HtmlText $r.Name) + ']</b> <small>(' + (ConvertTo-HtmlText $r.Table) + ' 측정값)</small>' }
        'Column'      { return '<b>' + (ConvertTo-HtmlText ($r.Table + '[' + $r.Name + ']')) + '</b>' }
        'Aggregation' { return '<b>' + (ConvertTo-HtmlText ($r.Table + '[' + $r.Name + ']')) + '</b> <small>→ 필드 ▾ <kbd>' + $script:AggUi[$r.FunctionName] + '</kbd></small>' }
    }
}
function New-ManualGuideHtml {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$RawDataDir)
    $h = New-Object System.Collections.Generic.List[string]
    $stepNo = 0
    $Step = { param([string]$Inner) $script:__gs++; return '<li class="step"><label><input type="checkbox" data-step="s' + $script:__gs + '"><span>' + $Inner + '</span></label></li>' }
    $script:__gs = 0
    $E = { param([string]$s) ConvertTo-HtmlText $s }
    $sourceTables = @($Model.Tables.Keys | Where-Object { -not $Model.Tables[$_].IsCalculated })
    $files = @($sourceTables | ForEach-Object { $t = $Model.Tables[$_]; if ($t.FilePattern) { $t.FilePattern } else { $t.File } } | Select-Object -Unique)
    $name = & $E $Model.Name

    $h.Add('<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>' + $name + ' — 수동 제작 가이드</title>')
    $h.Add(@'
<style>
:root{--ink:#1f2937;--muted:#6b7280;--line:#e5e7eb;--bg:#f8fafc;--card:#fff;--accent:#0f6cbd;--ok:#16a34a}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.65 "Segoe UI","Malgun Gothic","Apple SD Gothic Neo",sans-serif}
.wrap{max-width:960px;margin:0 auto;padding:28px 20px 80px}
h1{font-size:26px;margin:0 0 6px}h2{font-size:20px;margin:0}h3{font-size:17px;margin:22px 0 8px;color:var(--accent)}
.lead{color:var(--muted);margin:0 0 18px}.pill{display:inline-block;background:#e0f2fe;color:#075985;border-radius:999px;padding:2px 10px;font-size:13px;margin-right:6px}
details{background:var(--card);border:1px solid var(--line);border-radius:12px;margin:14px 0;padding:0}summary{cursor:pointer;padding:14px 18px;list-style:none;display:flex;align-items:center;gap:10px}summary::-webkit-details-marker{display:none}
summary .num{background:var(--accent);color:#fff;border-radius:8px;width:30px;height:30px;display:inline-flex;align-items:center;justify-content:center;font-weight:700;flex:none}
summary .chev{margin-left:auto;color:var(--muted);transition:transform .2s}details[open] summary .chev{transform:rotate(90deg)}
.body{padding:4px 18px 18px}
ol.steps{list-style:none;padding:0;margin:8px 0}.step{padding:8px 10px;border-radius:8px;margin:4px 0}.step:hover{background:#f1f5f9}.step.done span{color:var(--muted);text-decoration:line-through}
.step label{display:flex;gap:12px;align-items:flex-start;cursor:pointer}.step input{width:20px;height:20px;margin-top:4px;flex:none;accent-color:var(--ok)}
kbd{display:inline-block;background:#f3f4f6;border:1px solid #d1d5db;border-bottom-width:2px;border-radius:6px;padding:1px 8px;font:14px/1.5 inherit;color:#111827;white-space:nowrap}.sep{color:var(--muted);margin:0 4px}
code{background:#eef2ff;border-radius:4px;padding:1px 5px;font-family:Consolas,"D2Coding",monospace;font-size:14px}
.code{position:relative;margin:8px 0}.code pre{background:#0b1220;color:#e5e7eb;border-radius:10px;padding:14px 84px 14px 16px;overflow:auto;margin:0;font-size:14px;line-height:1.5;white-space:pre-wrap;word-break:break-all}.code pre code{background:none;color:inherit;padding:0}
.copy{position:absolute;top:8px;right:8px;background:#334155;color:#fff;border:0;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:13px}.copy:hover{background:var(--accent)}
.wells{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:8px;margin:8px 0}.well{border:1px dashed #94a3b8;border-radius:8px;padding:8px 10px;background:#f8fafc}.well .wn{display:block;font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.03em}
.minimap{width:100%;max-width:640px;height:auto;border:1px solid var(--line);border-radius:8px;background:#fff;display:block;margin:10px 0}
.visual{border:1px solid var(--line);border-radius:10px;padding:12px 14px;margin:10px 0;background:#fff}.visual h4{margin:0 0 6px;font-size:16px}.visual .badge{display:inline-block;background:#fef3c7;color:#92400e;border-radius:6px;padding:1px 8px;font-size:12px;margin-right:6px}
.pos{font-size:14px;color:var(--muted)}#progress{position:fixed;right:16px;bottom:16px;background:var(--ink);color:#fff;border-radius:999px;padding:8px 14px;font-size:14px;box-shadow:0 4px 14px rgba(0,0,0,.2)}
.note{background:#fffbeb;border-left:4px solid #f59e0b;padding:10px 14px;border-radius:6px;margin:10px 0}
@media print{details{break-inside:avoid}details:not([open])>*:not(summary){display:block}#progress,.copy{display:none}.step input{display:none}body{background:#fff}}
</style></head>
'@)
    $h.Add('<body data-report="' + $name + '"><div class="wrap">')
    $h.Add('<h1>' + $name + ' <small style="font-weight:400;color:var(--muted)">수동 제작 가이드</small></h1>')
    $h.Add('<p class="lead"><span class="pill">Power BI Desktop</span><span class="pill">영어 메뉴 기준</span> Claude가 만든 <code>' + $name + '.pbip</code>과 같은 리포트를 손으로 만드는 순서입니다. 단계마다 체크하면 진행 상태가 이 브라우저에 저장됩니다. 화면의 메뉴 이름은 <kbd>이렇게</kbd> 표시합니다.</p>')

    # ---- 0 준비
    $h.Add('<details open><summary><span class="num">0</span><h2>준비</h2><span class="chev">▶</span></summary><div class="body"><ol class="steps">')
    $h.Add((& $Step ('Power BI Desktop을 실행하고 ' + (Format-MenuHtml 'Blank report') + '을 선택합니다.')))
    $h.Add((& $Step ('데이터 파일 위치: <code>' + (& $E $RawDataDir) + '</code> — 사용할 파일: ' + (@($files | ForEach-Object { '<code>' + (& $E $_) + '</code>' }) -join ', '))))
    $h.Add((& $Step ('캔버스 크기: 캔버스 빈 곳 클릭 → ' + (Format-MenuHtml 'Visualizations > Format page > Canvas settings') + ' → <kbd>Type: Custom</kbd> <kbd>Width: 1920</kbd> <kbd>Height: 1080</kbd>')))
    $h.Add('</ol></div></details>')

    # ---- 1 데이터
    $h.Add('<details open><summary><span class="num">1</span><h2>데이터 가져오기 <small>(Get Data → Power Query Editor)</small></h2><span class="chev">▶</span></summary><div class="body">')
    $i = 0
    foreach ($tn in $sourceTables) {
        $t = $Model.Tables[$tn]; $i++
        $srcCols = @($t.Columns | Where-Object { $null -ne $_.Source }); $compCols = @($t.Columns | Where-Object { $null -eq $_.Source -and $_.Composite })
        $fileLabel = if ($t.FilePattern) { '<code>' + (& $E $t.FilePattern) + '</code> (여러 파일)' } else { '<code>' + (& $E $t.File) + '</code>' }
        if ($t.IsXlsx) { $fileLabel += ' / 시트 <code>' + (& $E $t.Sheet) + '</code>' }
        $h.Add('<h3>1.' + $i + ' 테이블 <code>' + (& $E $tn) + '</code> ← ' + $fileLabel + '</h3><ol class="steps">')
        if ($t.FilePattern) {
            $pp = ([string]$t.FilePattern).Split('*')
            $h.Add((& $Step ((Format-MenuHtml 'Home > Get Data > Folder') + ' → 폴더 경로에 <code>' + (& $E $RawDataDir) + '</code> 입력 → <kbd>OK</kbd>')))
            $h.Add((& $Step ('파일 목록 창에서 <kbd>Transform Data</kbd> → <kbd>Name</kbd> 열의 필터 ▾ → ' + (Format-MenuHtml 'Text Filters > Begins With…') + ' <code>' + (& $E $pp[0]) + '</code> 로 <code>' + (& $E $t.FilePattern) + '</code> 파일만 남깁니다.')))
            $sampleNote = if ($t.IsXlsx) { '시트 <kbd>' + (& $E $t.Sheet) + '</kbd> 선택' } else { '<kbd>File Origin</kbd> / <kbd>Delimiter</kbd> 확인' }
            $h.Add((& $Step ('<kbd>Content</kbd> 열 머리글의 <kbd>Combine Files</kbd> 버튼 클릭 → 샘플 파일 창에서 ' + $sampleNote + ' → <kbd>OK</kbd>. <small>같은 구조의 파일이 한 표로 합쳐지며, 같은 이름 규칙의 새 파일은 <kbd>Refresh</kbd>만으로 반영됩니다.</small>')))
        } elseif ($t.IsXlsx) {
            $h.Add((& $Step ((Format-MenuHtml 'Home > Get Data > Excel Workbook') + ' → <code>' + (& $E $t.File) + '</code> 선택 → <kbd>Open</kbd>')))
            $h.Add((& $Step ('<kbd>Navigator</kbd>에서 시트 <kbd>' + (& $E $t.Sheet) + '</kbd> 체크 → <kbd>Transform Data</kbd>')))
        } else {
            $origin = if ($t.Encoding -eq 949) { '949: Korean' } else { '65001: Unicode (UTF-8)' }
            $delimName = switch ($t.Delimiter) { ';' { 'Semicolon' } "`t" { 'Tab' } '|' { 'Custom → |' } default { 'Comma' } }
            $h.Add((& $Step ((Format-MenuHtml 'Home > Get Data > Text/CSV') + ' → <code>' + (& $E $t.File) + '</code> 선택 → <kbd>Open</kbd>')))
            $h.Add((& $Step ('미리보기 창에서 <kbd>File Origin: ' + $origin + '</kbd> <kbd>Delimiter: ' + $delimName + '</kbd> 확인 → <kbd>Transform Data</kbd>')))
        }
        $h.Add((& $Step ('<kbd>Power Query Editor</kbd> 오른쪽 <kbd>Query Settings</kbd> 창의 <kbd>Name</kbd>을 <code>' + (& $E $tn) + '</code>으로 변경')))
        if ($t.HeaderRow -gt 1) { $h.Add((& $Step ((Format-MenuHtml 'Home > Remove Rows > Remove Top Rows') + ' → <kbd>Number of rows: ' + ($t.HeaderRow - 1) + '</kbd> → <kbd>OK</kbd>'))) }
        $h.Add((& $Step ((Format-MenuHtml 'Home > Use First Row as Headers') + ' <small>(첫 행이 이미 헤더면 생략)</small>')))
        $h.Add((& $Step ((Format-MenuHtml 'Home > Choose Columns') + ' → 다음 열만 체크: ' + (@($srcCols | ForEach-Object { '<code>' + (& $E $_.Source) + '</code>' }) -join ' '))))
        $renames = @($srcCols | Where-Object { $_.Source -cne $_.Name })
        if ($renames.Count -gt 0) { $h.Add((& $Step ('열 이름 변경(열 헤더 더블클릭): ' + (@($renames | ForEach-Object { '<code>' + (& $E $_.Source) + '</code> → <code>' + (& $E $_.Name) + '</code>' }) -join ', ')))) }
        $h.Add((& $Step ('데이터 형식(열 헤더 왼쪽의 형식 아이콘 클릭): ' + (@($srcCols | ForEach-Object { '<code>' + (& $E $_.Name) + '</code> → <kbd>' + $script:PqTypeNames[$_.Type] + '</kbd>' }) -join ', '))))
        foreach ($c in $compCols) {
            $formula = '= Text.Combine({' + (@($c.Composite | ForEach-Object { 'Text.From([' + $_ + '])' }) -join ', ') + '}, "|")'
            $h.Add((& $Step ('복합 키 열 만들기: ' + (Format-MenuHtml 'Add Column > Custom Column') + ' → <kbd>New column name</kbd> <code>' + (& $E $c.Name) + '</code>, <kbd>Custom column formula</kbd>에 아래 수식 붙여넣기 → <kbd>OK</kbd> → 형식 <kbd>Text</kbd>. <small>관계는 열 하나끼리만 맺을 수 있어 ' + (& $E ($c.Composite -join ' + ')) + '를 합친 키가 필요합니다.</small>' + (New-CodeHtml $formula))))
        }
        $h.Add('</ol>')
    }
    $h.Add('<ol class="steps">' + (& $Step ('모든 테이블을 추가한 뒤 ' + (Format-MenuHtml 'Home > Close & Apply'))) + '</ol></div></details>')

    # ---- 2 모델
    $h.Add('<details open><summary><span class="num">2</span><h2>모델 설정 <small>(Model view)</small></h2><span class="chev">▶</span></summary><div class="body">')
    $h.Add('<p>왼쪽의 <kbd>Model view</kbd> 아이콘을 클릭합니다.</p><h3>2.1 열 속성 <small>(열을 클릭한 뒤 <kbd>Properties</kbd> 창)</small></h3><ol class="steps">')
    $any = $false
    foreach ($tn in $sourceTables) {
        foreach ($c in $Model.Tables[$tn].Columns) {
            $props = New-Object System.Collections.Generic.List[string]
            if ($c.Hidden) { $props.Add('<kbd>Is hidden: On</kbd>') }
            if ($c.Type -in 'int64', 'double', 'decimal' -and $c.SummarizeBy -ne 'sum') { $props.Add((Format-MenuHtml ('Advanced > Summarize by: ' + $script:SummarizeUi[$c.SummarizeBy]))) }
            if ($c.Format -and $c.Type -in 'int64', 'double', 'decimal') { $props.Add((Format-MenuHtml 'Formatting > Format: Custom') + ' → <code>' + (& $E $c.Format) + '</code>') }
            if ($c.SortBy) { $props.Add((Format-MenuHtml 'Column tools > Sort by column') + ' → <kbd>' + (& $E $c.SortBy) + '</kbd>') }
            if ($props.Count -gt 0) { $any = $true; $h.Add((& $Step ('<code>' + (& $E ($tn + '[' + $c.Name + ']')) + '</code>: ' + ($props -join ', ')))) }
        }
    }
    if (-not $any) { $h.Add('<li class="step"><span>변경할 열 속성이 없습니다.</span></li>') }
    $h.Add('</ol><h3>2.2 관계 (Relationships)</h3><ol class="steps">')
    $userRels = @($Model.Relationships | Where-Object { $_.ToTable -ne 'Calendar' })
    if ($userRels.Count -eq 0) { $h.Add('<li class="step"><span>테이블 간 관계가 없습니다 (단일 테이블).</span></li>') }
    foreach ($r in $userRels) {
        $dir = if ($r.CrossFilter -eq 'both') { 'Both' } else { 'Single' }; $act = if ($r.Active) { '체크' } else { '체크 해제' }
        $h.Add((& $Step ((Format-MenuHtml 'Home > Manage relationships > New relationship') + ': From table <kbd>' + (& $E $r.FromTable) + '</kbd> 열 <kbd>' + (& $E $r.FromColumn) + '</kbd> → To table <kbd>' + (& $E $r.ToTable) + '</kbd> 열 <kbd>' + (& $E $r.ToColumn) + '</kbd>, <kbd>Cardinality: Many to one (*:1)</kbd> <kbd>Cross-filter direction: ' + $dir + '</kbd> <kbd>Make this relationship active</kbd> ' + $act + ' → <kbd>OK</kbd>')))
    }
    $h.Add('</ol>')
    if ($null -ne $Model.Calendar) {
        $h.Add('<h3>2.3 Calendar 테이블 (날짜 테이블)</h3><ol class="steps">')
        $h.Add((& $Step ((Format-MenuHtml 'Home > New table') + ' → 수식 입력줄에 아래 DAX를 붙여넣고 Enter' + (New-CodeHtml ('Calendar = ' + (New-CalendarDax -Model $Model))))))
        $h.Add((& $Step ('<kbd>Data</kbd> 창에서 <kbd>Calendar</kbd> 선택 → ' + (Format-MenuHtml 'Table tools > Mark as date table > Mark as date table') + ' → <kbd>Date column: Date</kbd> → <kbd>OK</kbd>')))
        $h.Add((& $Step ('<code>Calendar[Month]</code> 선택 → ' + (Format-MenuHtml 'Column tools > Sort by column') + ' → <kbd>MonthNo</kbd>. <code>Calendar[MonthNo]</code>는 <kbd>Properties</kbd> → <kbd>Is hidden: On</kbd>')))
        foreach ($r in @($Model.Relationships | Where-Object { $_.ToTable -eq 'Calendar' })) {
            $act = if ($r.Active) { '체크' } else { '체크 해제 (비활성 관계)' }
            $h.Add((& $Step ((Format-MenuHtml 'Home > Manage relationships > New relationship') + ': From <kbd>' + (& $E $r.FromTable) + '</kbd> 열 <kbd>' + (& $E $r.FromColumn) + '</kbd> → To <kbd>Calendar</kbd> 열 <kbd>Date</kbd>, <kbd>Many to one (*:1)</kbd> <kbd>Single</kbd> <kbd>Make this relationship active</kbd> ' + $act + ' → <kbd>OK</kbd>')))
        }
        $h.Add('</ol>')
    }
    $h.Add('</div></details>')

    # ---- 3 측정값
    $h.Add('<details open><summary><span class="num">3</span><h2>측정값 <small>(Measures)</small></h2><span class="chev">▶</span></summary><div class="body">')
    $h.Add('<p><kbd>Report view</kbd>로 돌아가 <kbd>Data</kbd> 창에서 테이블을 선택한 뒤 ' + (Format-MenuHtml 'Home > New measure') + '를 클릭하고, 수식 입력줄에 아래 DAX를 붙여넣고 Enter를 누릅니다.</p><ol class="steps">')
    $mi = 0
    foreach ($tn in $Model.Tables.Keys) {
        foreach ($m in $Model.Tables[$tn].Measures) {
            $mi++
            $fmt = if ($m.Format) { ' → ' + (Format-MenuHtml 'Measure tools > Format') + ' <kbd>Custom</kbd> <code>' + (& $E $m.Format) + '</code>' } else { '' }
            $h.Add((& $Step ('<code>' + (& $E $tn) + '</code> 테이블 → <kbd>New measure</kbd>' + $fmt + (New-CodeHtml ($m.Name + ' = ' + $m.Dax)))))
        }
    }
    if ($mi -eq 0) { $h.Add('<li class="step"><span>측정값이 없습니다.</span></li>') }
    $h.Add('</ol></div></details>')

    # ---- 4 페이지
    $h.Add('<details open><summary><span class="num">4</span><h2>보고서 페이지 <small>(Report view)</small></h2><span class="chev">▶</span></summary><div class="body">')
    $pi = 0
    foreach ($page in $Model.Pages) {
        $pi++
        $pname = [string](Get-Prop $page 'name'); $visuals = @(ConvertTo-Array (Get-Prop $page 'visuals'))
        $h.Add('<h3>페이지 ' + $pi + ': <code>' + (& $E $pname) + '</code></h3>')
        # minimap
        $svg = New-Object System.Collections.Generic.List[string]
        $svg.Add('<svg class="minimap" viewBox="0 0 1920 1080" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="1920" height="1080" fill="#fff"/>')
        $vi = 0
        foreach ($v in $visuals) {
            $vi++; $pos = Get-GridPosition -Grid (ConvertTo-Array (Get-Prop $v 'grid')); $type = [string](Get-Prop $v 'type')
            $label = '' + $vi + '. ' + $script:VisualUiNames[$type]
            $svg.Add('<rect x="' + $pos.x + '" y="' + $pos.y + '" width="' + $pos.width + '" height="' + $pos.height + '" rx="14" fill="#e0f2fe" stroke="#0f6cbd" stroke-width="4"/>')
            $svg.Add('<text x="' + ($pos.x + 24) + '" y="' + ($pos.y + 58) + '" font-size="40" font-family="Segoe UI, sans-serif" fill="#0c4a6e">' + (& $E $label) + '</text>')
        }
        $svg.Add('</svg>')
        $h.Add('<p class="pos">아래 그림은 이 페이지의 완성 배치입니다(1920×1080). 번호는 아래 단계 번호와 같습니다.</p>' + ($svg -join ''))
        $h.Add('<ol class="steps">')
        if ($pi -eq 1) { $h.Add((& $Step ('하단의 페이지 탭 더블클릭 → 이름을 <code>' + (& $E $pname) + '</code>으로 변경'))) } else { $h.Add((& $Step ('하단의 <kbd>+</kbd> (New page) 클릭 → 탭 더블클릭 → 이름을 <code>' + (& $E $pname) + '</code>으로 변경'))) }
        foreach ($f in (ConvertTo-Array (Get-Prop $page 'filters'))) {
            $fr = Resolve-FieldRef -Model $Model -Ref ([string](Get-Prop $f 'field'))
            $h.Add((& $Step ('페이지 필터: <kbd>Filters</kbd> 창의 <kbd>Filters on this page</kbd>에 <b>' + (& $E ($fr.Table + '[' + $fr.Name + ']')) + '</b> 드래그 → <kbd>Basic filtering</kbd> → ' + (@((ConvertTo-Array (Get-Prop $f 'in')) | ForEach-Object { '<code>' + (& $E ([string]$_)) + '</code>' }) -join ' ') + ' 체크')))
        }
        $h.Add('</ol>')
        $vi = 0
        foreach ($v in $visuals) {
            $vi++
            $type = [string](Get-Prop $v 'type'); $ui = $script:VisualUiNames[$type]
            $pos = Get-GridPosition -Grid (ConvertTo-Array (Get-Prop $v 'grid'))
            $title = Get-Prop $v 'title' $null
            $head = '' + $vi + '. <span class="badge">' + (& $E $ui) + '</span>'; if ($title) { $head += ' ' + (& $E $title) } elseif ($type -eq 'textbox') { $head += ' ' + (& $E ([string](Get-Prop $v 'text'))) }
            $h.Add('<div class="visual"><h4>' + $head + '</h4><ol class="steps">')
            if ($type -eq 'textbox') {
                $h.Add((& $Step ((Format-MenuHtml 'Insert > Text box') + ' → 텍스트 입력: <code>' + (& $E ([string](Get-Prop $v 'text'))) + '</code>, 글꼴 크기 <kbd>' + [int](Get-Prop $v 'fontSize' 20) + '</kbd> <kbd>Bold</kbd>')))
            } else {
                $h.Add((& $Step ('캔버스 빈 곳 클릭 → <kbd>Visualizations</kbd> 창에서 <kbd>' + (& $E $ui) + '</kbd> 클릭')))
                $wells = $script:WellNames[$type]; $fields = Get-Prop $v 'fields' $null
                $wh = New-Object System.Collections.Generic.List[string]
                foreach ($role in ($script:VisualCatalog[$type].Required + $script:VisualCatalog[$type].Optional)) {
                    $refs = ConvertTo-Array (Get-Prop $fields $role); if ($refs.Count -eq 0) { continue }
                    $wh.Add('<div class="well"><span class="wn">' + (& $E $wells[$role]) + '</span>' + (@($refs | ForEach-Object { New-FieldHtml -Model $Model -Ref ([string]$_) }) -join '<br>') + '</div>')
                }
                $h.Add((& $Step ('<kbd>Data</kbd> 창에서 필드를 아래 상자(<kbd>Visualizations</kbd> 창의 필드 영역)로 드래그:<div class="wells">' + ($wh -join '') + '</div>')))
                if ($title) { $h.Add((& $Step ('제목: ' + (Format-MenuHtml 'Format visual > General > Title > Text') + ' <code>' + (& $E $title) + '</code>'))) }
                $sort = Get-Prop $v 'sort' $null
                if ($null -ne $sort) {
                    $sr = Resolve-FieldRef -Model $Model -Ref ([string](Get-Prop $sort 'field')); $sd = if ([string](Get-Prop $sort 'direction' 'desc') -eq 'asc') { 'Sort ascending' } else { 'Sort descending' }
                    $h.Add((& $Step ('정렬: 비주얼 우측 상단 <kbd>More options (…)</kbd> → <kbd>Sort axis</kbd> → <kbd>' + (& $E $sr.Name) + '</kbd> → <kbd>' + $sd + '</kbd>')))
                }
                $topN = Get-Prop $v 'topN' $null
                if ($null -ne $topN) {
                    $catRefs = ConvertTo-Array (Get-Prop $fields 'Category'); if ($catRefs.Count -eq 0) { $catRefs = ConvertTo-Array (Get-Prop $fields 'Rows') }
                    $cr = Resolve-FieldRef -Model $Model -Ref ([string]$catRefs[0])
                    $h.Add((& $Step ('상위 N: <kbd>Filters</kbd> 창 → <kbd>Filters on this visual</kbd>의 <kbd>' + (& $E $cr.Name) + '</kbd> 카드 펼치기 → <kbd>Filter type: Top N</kbd> <kbd>Show items: Top ' + [int](Get-Prop $topN 'n') + '</kbd> <kbd>By value</kbd> ← ' + (New-FieldHtml -Model $Model -Ref ([string](Get-Prop $topN 'by'))) + ' 드래그 → <kbd>Apply filter</kbd>')))
                }
                foreach ($f in (ConvertTo-Array (Get-Prop $v 'filters'))) {
                    $fr = Resolve-FieldRef -Model $Model -Ref ([string](Get-Prop $f 'field'))
                    $h.Add((& $Step ('비주얼 필터: <kbd>Filters on this visual</kbd>에 <b>' + (& $E ($fr.Table + '[' + $fr.Name + ']')) + '</b> 드래그 → <kbd>Basic filtering</kbd> → ' + (@((ConvertTo-Array (Get-Prop $f 'in')) | ForEach-Object { '<code>' + (& $E ([string]$_)) + '</code>' }) -join ' ') + ' 체크')))
                }
            }
            $h.Add((& $Step ('위치/크기: ' + (Format-MenuHtml 'Format > General > Properties') + ' → <kbd>Position</kbd> X <code>' + $pos.x + '</code> Y <code>' + $pos.y + '</code>, <kbd>Size</kbd> Width <code>' + $pos.width + '</code> Height <code>' + $pos.height + '</code>')))
            $h.Add('</ol></div>')
        }
    }
    $h.Add('</div></details>')

    # ---- 5 저장
    $h.Add('<details open><summary><span class="num">5</span><h2>저장</h2><span class="chev">▶</span></summary><div class="body"><ol class="steps">')
    $h.Add((& $Step ((Format-MenuHtml 'File > Save As') + ' → 파일 형식 <kbd>Power BI files (*.pbix)</kbd> → <code>' + $name + '.pbix</code>로 저장')))
    $h.Add('</ol><div class="note">막히면 Claude 채팅에 화면의 오류 문구를 그대로 붙여 넣고 물어보세요. 이 문서는 <code>report-spec.json</code>에서 자동 생성되었습니다.</div></div></details>')
    $h.Add('<div id="progress"></div>')
    $h.Add(@'
<script>
(function(){var key='pbi-guide-'+document.body.getAttribute('data-report');var saved={};try{saved=JSON.parse(localStorage.getItem(key)||'{}')}catch(e){}
var boxes=Array.prototype.slice.call(document.querySelectorAll('input[data-step]'));
function step(b){var n=b;while(n&&!(n.classList&&n.classList.contains('step')))n=n.parentNode;return n}
function progress(){var d=0;boxes.forEach(function(b){if(b.checked)d++});var p=document.getElementById('progress');if(p)p.textContent=d+' / '+boxes.length+' 단계 완료'}
boxes.forEach(function(b){var id=b.getAttribute('data-step');if(saved[id]){b.checked=true;step(b).classList.add('done')}
b.addEventListener('change',function(){saved[id]=b.checked;step(b).classList.toggle('done',b.checked);try{localStorage.setItem(key,JSON.stringify(saved))}catch(e){}progress()})});
progress();
Array.prototype.slice.call(document.querySelectorAll('button.copy')).forEach(function(btn){btn.addEventListener('click',function(){var t=btn.parentNode.querySelector('code').textContent;
function ok(){btn.textContent='복사됨';setTimeout(function(){btn.textContent='복사'},1500)}
if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(t).then(ok,function(){})}else{var ta=document.createElement('textarea');ta.value=t;document.body.appendChild(ta);ta.select();try{document.execCommand('copy');ok()}catch(e){}document.body.removeChild(ta)}})});
})();
</script></div></body></html>
'@)
    return ($h -join "`r`n")
}
