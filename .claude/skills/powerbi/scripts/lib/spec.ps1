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
    if ($null -ne $agg -and -not $script:AggFunctions.Contains($agg)) { throw "unknown aggregation prefix '${agg}:' in '$Ref' (allowed: $($script:AggFunctions.Keys -join ', '))" }
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
            if ($seenCols.ContainsKey($final)) { $errors.Add("${cp}: duplicate column name '$final'") }; $seenCols[$final] = $true
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
            try { $r = Resolve-FieldRef -Model $model -Ref $ref; if ($r.Kind -ne 'Column') { $errors.Add("$p.${side}: must reference a column") } }
            catch { $errors.Add("$p.${side}: " + $_.Exception.Message) }
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
            foreach ($req in $cat.Required) { if ($roleNames -notcontains $req -or (ConvertTo-Array (Get-Prop $fields $req)).Count -eq 0) { $errors.Add("$vp.fields.${req}: required for $type") } }
            foreach ($role in $roleNames) {
                if (($cat.Required + $cat.Optional) -notcontains $role) { $errors.Add("$vp.fields.${role}: not a role of $type (roles: $(($cat.Required + $cat.Optional) -join ', '))"); continue }
                $refs = ConvertTo-Array (Get-Prop $fields $role)
                if ($cat.Max.ContainsKey($role) -and $refs.Count -gt $cat.Max[$role]) { $errors.Add("$vp.fields.${role}: at most $($cat.Max[$role]) field(s)") }
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
