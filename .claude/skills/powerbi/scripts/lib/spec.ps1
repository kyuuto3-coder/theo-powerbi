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

function Resolve-KeySide {
    # One side of a relationship. Single ref → its table/column. Several refs (composite key) → adds a hidden text column '_key_<parts>' to the table (once) and returns it.
    param($Tables, [string[]]$Refs, [bool]$IsKey)
    $parts = New-Object System.Collections.Generic.List[string]; $table = $null
    foreach ($ref in $Refs) {
        $i = $ref.IndexOf('.'); if ($i -lt 1) { return $null }
        $tn = $ref.Substring(0, $i); $cn = $ref.Substring($i + 1)
        if ($null -eq $table) { $table = $tn } elseif ($table -ne $tn) { return $null }
        $parts.Add($cn)
    }
    if (-not $Tables.Contains($table)) { return $null }
    if ($parts.Count -eq 1) { return @{ Table = $table; Column = $parts[0] } }
    $keyName = '_key_' + ($parts -join '_')
    $exists = $false; foreach ($c in $Tables[$table].Columns) { if ($c.Name -eq $keyName) { $exists = $true } }
    if (-not $exists) {
        $Tables[$table].Columns = @($Tables[$table].Columns) + @(@{ Source = $null; Name = $keyName; Type = 'text'; Key = $IsKey; Hidden = $true; Format = $null; SummarizeBy = 'none'
                                                                     Description = ('Composite key: ' + ($parts -join ' + ')); SortBy = $null; Composite = $parts.ToArray() })
    }
    return @{ Table = $table; Column = $keyName }
}

function Get-SpecModel {
    # Builds the normalized model. Tolerant: invalid entries are skipped here and reported by Get-SpecValidation.
    param([Parameter(Mandatory)]$Spec)
    $tables = [ordered]@{}
    foreach ($t in (ConvertTo-Array (Get-Prop $Spec 'tables'))) {
        $name = [string](Get-Prop $t 'name' '')
        if (-not $name) { continue }
        if ([string](Get-Prop $t 'from' '')) { continue }   # summary tables are built in a second pass, after their source exists
        $file = [string](Get-Prop $t 'file' '')
        $cols = New-Object System.Collections.Generic.List[object]
        foreach ($c in (ConvertTo-Array (Get-Prop $t 'columns'))) {
            $src = [string](Get-Prop $c 'name' ''); if (-not $src) { continue }
            $final = [string](Get-Prop $c 'rename' $src); if (-not $final) { $final = $src }
            $type = [string](Get-Prop $c 'type' 'text')
            $key = [bool](Get-Prop $c 'key' $false)
            $vmRaw = Get-Prop $c 'valueMap' $null
            $vm = [ordered]@{}
            if ($null -ne $vmRaw) {
                if ($vmRaw -is [System.Collections.IDictionary]) { foreach ($k in $vmRaw.Keys) { $vm[[string]$k] = [string]$vmRaw[$k] } }
                else { foreach ($pp in $vmRaw.PSObject.Properties) { $vm[[string]$pp.Name] = [string]$pp.Value } }
            }
            $cols.Add(@{ Source = $src; Name = $final; Type = $type; Key = $key; Hidden = [bool](Get-Prop $c 'hidden' $false)
                         Format = (Get-Prop $c 'format' $null); SummarizeBy = [string](Get-Prop $c 'summarizeBy' (Get-DefaultSummarize $type $key))
                         Description = (Get-Prop $c 'description' $null); SortBy = (Get-Prop $c 'sortBy' $null)
                         Trim = [bool](Get-Prop $c 'trim' $false); Case = (Get-Prop $c 'case' $null)
                         NullValues = @(ConvertTo-Array (Get-Prop $c 'nullValues') | ForEach-Object { [string]$_ })
                         DateFormats = @(ConvertTo-Array (Get-Prop $c 'dateFormats') | ForEach-Object { [string]$_ })
                         ValueMap = $vm; NumberClean = [bool](Get-Prop $c 'numberClean' $false); FillDown = [bool](Get-Prop $c 'fillDown' $false)
                         ValidRange = @(ConvertTo-Array (Get-Prop $c 'validRange')) })
        }
        $derived = New-Object System.Collections.Generic.List[object]
        foreach ($d in (ConvertTo-Array (Get-Prop $t 'derived'))) {
            $dnm = [string](Get-Prop $d 'name' ''); if (-not $dnm) { continue }
            $dty = [string](Get-Prop $d 'type' 'double')
            $derived.Add(@{ Name = $dnm; Expr = [string](Get-Prop $d 'expr' ''); Type = $dty })
            $cols.Add(@{ Source = $null; Name = $dnm; Type = $dty; Key = $false; Hidden = [bool](Get-Prop $d 'hidden' $false)
                         Format = (Get-Prop $d 'format' $null); SummarizeBy = [string](Get-Prop $d 'summarizeBy' (Get-DefaultSummarize $dty $false))
                         Description = (Get-Prop $d 'description' $null); SortBy = $null })
        }
        $pattern = [string](Get-Prop $t 'filePattern' '')
        $cleanRaw = Get-Prop $t 'clean' $null
        $dbRaw = Get-Prop $cleanRaw 'dedupeBy' $null
        $dedupeBy = $null
        if ($null -ne $dbRaw) { $dedupeBy = @{ KeyCols = @(ConvertTo-Array (Get-Prop $dbRaw 'keys') | ForEach-Object { [string]$_ }); Keep = [string](Get-Prop $dbRaw 'keep' 'first'); OrderBy = [string](Get-Prop $dbRaw 'orderBy' '') } }
        $tables[$name] = @{ Name = $name; File = $file; FilePattern = $pattern; Sheet = (Get-Prop $t 'sheet' $null); HeaderRow = [int](Get-Prop $t 'headerRow' 1)
                            Encoding = [int](Get-Prop $t 'encoding' 65001); Delimiter = [string](Get-Prop $t 'delimiter' ',')
                            IsXlsx = (($file + $pattern) -match '(?i)\.xls[xm]$'); Columns = $cols.ToArray(); Measures = @(); IsCalculated = $false
                            Derived = $derived.ToArray(); From = ''; GroupBy = @(); Aggs = @()
                            Clean = @{ RemoveBlankRows = [bool](Get-Prop $cleanRaw 'removeBlankRows' $false); DropDuplicates = [bool](Get-Prop $cleanRaw 'dropDuplicates' $false); ErrorsToNull = [bool](Get-Prop $cleanRaw 'errorsToNull' $false); DedupeBy = $dedupeBy } }
    }
    # second pass: summary tables (from/groupBy/aggregations) - their columns are derived from the source table
    foreach ($t in (ConvertTo-Array (Get-Prop $Spec 'tables'))) {
        $from = [string](Get-Prop $t 'from' ''); if (-not $from) { continue }
        $name = [string](Get-Prop $t 'name' ''); if (-not $name -or -not $tables.Contains($from)) { continue }
        $srcT = $tables[$from]
        $gcols = New-Object System.Collections.Generic.List[object]
        $groupBy = @(ConvertTo-Array (Get-Prop $t 'groupBy') | ForEach-Object { [string]$_ })
        foreach ($gn in $groupBy) {
            $sc = $srcT.Columns | Where-Object { $_.Name -eq $gn } | Select-Object -First 1
            $gty = if ($sc) { $sc.Type } else { 'text' }
            $gcols.Add(@{ Source = $null; Name = $gn; Type = $gty; Key = $false; Hidden = $false; Format = $null; SummarizeBy = 'none'; Description = $null; SortBy = $null })
        }
        $aggs = New-Object System.Collections.Generic.List[object]
        foreach ($a in (ConvertTo-Array (Get-Prop $t 'aggregations'))) {
            $an = [string](Get-Prop $a 'name' ''); if (-not $an) { continue }
            $agg = [string](Get-Prop $a 'agg' 'count'); $acol = [string](Get-Prop $a 'column' '')
            $sc = $srcT.Columns | Where-Object { $_.Name -eq $acol } | Select-Object -First 1
            $aty = if ($agg -in 'count', 'countDistinct') { 'int64' } elseif ($agg -eq 'average') { 'double' } elseif ($sc) { $sc.Type } else { 'double' }
            $aggs.Add(@{ Name = $an; Agg = $agg; Column = $acol; Type = $aty })
            $gcols.Add(@{ Source = $null; Name = $an; Type = $aty; Key = $false; Hidden = $false; Format = (Get-Prop $a 'format' $null); SummarizeBy = [string](Get-DefaultSummarize $aty $false); Description = $null; SortBy = $null })
        }
        $tables[$name] = @{ Name = $name; File = $null; FilePattern = ''; Sheet = $null; HeaderRow = 1; Encoding = 65001; Delimiter = ','; IsXlsx = $false
                            Columns = $gcols.ToArray(); Measures = @(); IsCalculated = $false
                            Derived = @(); From = $from; GroupBy = $groupBy; Aggs = $aggs.ToArray()
                            Clean = @{ RemoveBlankRows = $false; DropDuplicates = $false; ErrorsToNull = $false; DedupeBy = $null } }
    }
    foreach ($m in (ConvertTo-Array (Get-Prop $Spec 'measures'))) {
        $tn = [string](Get-Prop $m 'table' '')
        if (-not $tables.Contains($tn)) { continue }
        $tables[$tn].Measures = @($tables[$tn].Measures) + @(@{ Name = [string](Get-Prop $m 'name' ''); Dax = [string](Get-Prop $m 'dax' ''); Format = (Get-Prop $m 'format' $null); Description = (Get-Prop $m 'description' $null) })
    }
    $rels = New-Object System.Collections.Generic.List[object]
    foreach ($r in (ConvertTo-Array (Get-Prop $Spec 'relationships'))) {
        $fromRefs = @(ConvertTo-Array (Get-Prop $r 'from') | ForEach-Object { [string]$_ }); $toRefs = @(ConvertTo-Array (Get-Prop $r 'to') | ForEach-Object { [string]$_ })
        if ($fromRefs.Count -eq 0 -or $toRefs.Count -eq 0 -or $fromRefs.Count -ne $toRefs.Count) { continue }
        $fromSide = Resolve-KeySide -Tables $tables -Refs $fromRefs -IsKey $false
        $toSide = Resolve-KeySide -Tables $tables -Refs $toRefs -IsKey $true
        if ($null -eq $fromSide -or $null -eq $toSide) { continue }
        $rels.Add(@{ FromTable = $fromSide.Table; FromColumn = $fromSide.Column; ToTable = $toSide.Table; ToColumn = $toSide.Column
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
        if ([string](Get-Prop $t 'from' '')) { continue }   # summary tables are validated after the model is built
        $file = [string](Get-Prop $t 'file' ''); $pattern = [string](Get-Prop $t 'filePattern' '')
        if (-not $file -and -not $pattern) { $errors.Add("$p.file: required (or filePattern)") }
        elseif ($file -and $pattern) { $errors.Add("${p}: use either file or filePattern, not both") }
        elseif ($pattern) {
            if (($pattern.Split('*').Count - 1) -ne 1) { $errors.Add("$p.filePattern: must contain exactly one * (e.g. orders_*.csv)") }
            elseif (@(Get-ChildItem -LiteralPath $RawDataDir -File -Filter $pattern -ErrorAction SilentlyContinue).Count -eq 0) { $errors.Add("$p.filePattern: no file in $RawDataDir matches '$pattern'") }
            elseif ($pattern -match '(?i)\.xls[xm]$' -and -not (Get-Prop $t 'sheet' $null)) { $errors.Add("$p.sheet: required for Excel files") }
        }
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
            $case = Get-Prop $c 'case' $null
            if ($null -ne $case -and @('upper', 'lower') -notcontains [string]$case) { $errors.Add("$cp.case: '$case' is not one of upper|lower") }
            if (@(ConvertTo-Array (Get-Prop $c 'dateFormats')).Count -gt 0 -and $type -notin 'date', 'datetime') { $errors.Add("$cp.dateFormats: only valid for date/datetime columns") }
            $vr = @(ConvertTo-Array (Get-Prop $c 'validRange'))
            if ($vr.Count -gt 0) {
                if ($vr.Count -ne 2) { $errors.Add("$cp.validRange: must be [min, max]") }
                elseif ($type -notin 'int64', 'double', 'decimal') { $errors.Add("$cp.validRange: only valid for numeric columns") }
                elseif ([double]$vr[0] -gt [double]$vr[1]) { $errors.Add("$cp.validRange: min > max") }
            }
        }
        $derRaw = ConvertTo-Array (Get-Prop $t 'derived')
        for ($di = 0; $di -lt $derRaw.Count; $di++) {
            $d = $derRaw[$di]; $dp = "$p.derived[$di]"
            $dnm = [string](Get-Prop $d 'name' '')
            if (-not $dnm) { $errors.Add("$dp.name: required"); continue }
            if ($seenCols.ContainsKey($dnm)) { $errors.Add("$dp.name: '$dnm' collides with a column name") }; $seenCols[$dnm] = $true
            if (-not [string](Get-Prop $d 'expr' '')) { $errors.Add("$dp.expr: required (an M expression over renamed columns, e.g. [Adults] + [Children])") }
            $dty = [string](Get-Prop $d 'type' 'double')
            if ($script:AllowedTypes -notcontains $dty) { $errors.Add("$dp.type: '$dty' is not one of $($script:AllowedTypes -join '|')") }
        }
        $cleanRaw = Get-Prop $t 'clean' $null
        if ($null -ne $cleanRaw) {
            $cleanProps = if ($cleanRaw -is [System.Collections.IDictionary]) { @($cleanRaw.Keys) } else { @($cleanRaw.PSObject.Properties.Name) }
            foreach ($prop in $cleanProps) { if ($prop -notin 'removeBlankRows', 'dropDuplicates', 'errorsToNull', 'dedupeBy') { $errors.Add("$p.clean.${prop}: unknown option (use removeBlankRows|dropDuplicates|errorsToNull|dedupeBy)") } }
            $dbRaw = Get-Prop $cleanRaw 'dedupeBy' $null
            if ($null -ne $dbRaw) {
                $dbKeys = @(ConvertTo-Array (Get-Prop $dbRaw 'keys') | ForEach-Object { [string]$_ })
                if ($dbKeys.Count -eq 0) { $errors.Add("$p.clean.dedupeBy.keys: required (renamed column names)") }
                foreach ($dk in $dbKeys) { if (-not $seenCols.ContainsKey($dk)) { $errors.Add("$p.clean.dedupeBy.keys: unknown column '$dk'") } }
                $keep = [string](Get-Prop $dbRaw 'keep' 'first')
                if ($keep -notin 'first', 'last', 'mostComplete') { $errors.Add("$p.clean.dedupeBy.keep: '$keep' is not one of first|last|mostComplete") }
                $ob = [string](Get-Prop $dbRaw 'orderBy' '')
                if ($keep -in 'first', 'last') {
                    if (-not $ob) { $errors.Add("$p.clean.dedupeBy.orderBy: required for keep=first|last (the column that decides which row wins)") }
                    elseif (-not $seenCols.ContainsKey($ob)) { $errors.Add("$p.clean.dedupeBy.orderBy: unknown column '$ob'") }
                }
            }
        }
    }
    $model = Get-SpecModel -Spec $Spec
    # summary tables (from/groupBy/aggregations)
    for ($ti = 0; $ti -lt $tablesRaw.Count; $ti++) {
        $t = $tablesRaw[$ti]; $p = "tables[$ti]"
        $from = [string](Get-Prop $t 'from' ''); if (-not $from) { continue }
        foreach ($bad in 'file', 'filePattern', 'columns', 'clean', 'derived') { if ($null -ne (Get-Prop $t $bad $null)) { $errors.Add("$p.${bad}: not allowed on a summary table (it inherits the cleaned source)") } }
        if (-not $model.Tables.Contains($from)) { $errors.Add("$p.from: unknown table '$from'"); continue }
        if ($model.Tables[$from].From) { $errors.Add("$p.from: '$from' is itself a summary table; group a file table instead"); continue }
        $srcCols = @{}; foreach ($c in $model.Tables[$from].Columns) { $srcCols[$c.Name] = $true }
        $groupBy = @(ConvertTo-Array (Get-Prop $t 'groupBy') | ForEach-Object { [string]$_ })
        if ($groupBy.Count -eq 0) { $errors.Add("$p.groupBy: at least one grouping column is required") }
        foreach ($g in $groupBy) { if (-not $srcCols.ContainsKey($g)) { $errors.Add("$p.groupBy: unknown column '$g' in '$from' (use renamed names)") } }
        $aggsRaw = ConvertTo-Array (Get-Prop $t 'aggregations')
        if ($aggsRaw.Count -eq 0) { $errors.Add("$p.aggregations: at least one aggregation is required") }
        for ($ai = 0; $ai -lt $aggsRaw.Count; $ai++) {
            $a = $aggsRaw[$ai]; $ap = "$p.aggregations[$ai]"
            if (-not [string](Get-Prop $a 'name' '')) { $errors.Add("$ap.name: required") }
            $agg = [string](Get-Prop $a 'agg' 'count')
            if ($agg -notin 'count', 'countDistinct', 'sum', 'average', 'min', 'max') { $errors.Add("$ap.agg: '$agg' is not one of count|countDistinct|sum|average|min|max") }
            $acol = [string](Get-Prop $a 'column' '')
            if ($agg -ne 'count') {
                if (-not $acol) { $errors.Add("$ap.column: required for agg=$agg") }
                elseif (-not $srcCols.ContainsKey($acol)) { $errors.Add("$ap.column: unknown column '$acol' in '$from' (use renamed names)") }
            }
        }
    }
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
        $fromRefs = @(ConvertTo-Array (Get-Prop $relsRaw[$ri] 'from') | ForEach-Object { [string]$_ }); $toRefs = @(ConvertTo-Array (Get-Prop $relsRaw[$ri] 'to') | ForEach-Object { [string]$_ })
        if ($fromRefs.Count -eq 0) { $errors.Add("$p.from: required") }; if ($toRefs.Count -eq 0) { $errors.Add("$p.to: required") }
        if ($fromRefs.Count -gt 0 -and $toRefs.Count -gt 0 -and $fromRefs.Count -ne $toRefs.Count) { $errors.Add("${p}: from and to must have the same number of columns (composite keys pair up position by position)") }
        $sides = @{ from = $fromRefs; to = $toRefs }
        foreach ($side in 'from', 'to') {
            $tablesSeen = @{}
            for ($k = 0; $k -lt $sides[$side].Count; $k++) {
                $ref = $sides[$side][$k]
                try {
                    $r = Resolve-FieldRef -Model $model -Ref $ref
                    if ($r.Kind -ne 'Column') { $errors.Add("$p.${side}[$k]: must reference a column") }
                    $tablesSeen[$r.Table] = $true
                } catch { $errors.Add("$p.${side}[$k]: " + $_.Exception.Message) }
            }
            if ($tablesSeen.Count -gt 1) { $errors.Add("$p.${side}: all columns of a composite key must belong to the same table") }
        }
        if ($fromRefs.Count -eq $toRefs.Count) {
            for ($k = 0; $k -lt $fromRefs.Count; $k++) {
                try { $a = Resolve-FieldRef -Model $model -Ref $fromRefs[$k]; $b = Resolve-FieldRef -Model $model -Ref $toRefs[$k]
                      $ga = if ($a.Type -in 'int64','double','decimal') { 'number' } elseif ($a.Type -in 'date','datetime') { 'date' } else { $a.Type }
                      $gb = if ($b.Type -in 'int64','double','decimal') { 'number' } elseif ($b.Type -in 'date','datetime') { 'date' } else { $b.Type }
                      if ($ga -ne $gb -and $a.Type -and $b.Type) { $warnings.Add("${p}: '$($fromRefs[$k])' ($($a.Type)) and '$($toRefs[$k])' ($($b.Type)) have different types; values may not match") } } catch { }
            }
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
