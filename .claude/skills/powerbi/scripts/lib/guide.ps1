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
