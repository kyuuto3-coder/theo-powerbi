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

function Get-MReader {
    # M expression that turns a binary ($Binary) into the raw sheet/csv table for this table's source type.
    param($Table, [string]$Binary)
    if ($Table.IsXlsx) { return 'Excel.Workbook(' + $Binary + ', null, true){[Item=' + (ConvertTo-MString ([string]$Table.Sheet)) + ',Kind="Sheet"]}[Data]' }
    $delim = if ($Table.Delimiter -eq "`t") { '#(tab)' } else { ([string]$Table.Delimiter).Replace('"', '""') }
    return 'Csv.Document(' + $Binary + ',[Delimiter="' + $delim + '", Encoding=' + $Table.Encoding + ', QuoteStyle=QuoteStyle.Csv])'
}

function New-MQuery {
    # Returns the M query as an array of lines (unindented).
    # Pipeline: Source → (Skip) → Promoted → Selected → (NoBlankRows) → (Renamed) → (Cleaned) → (FilledDown) → (FixedDates) → Typed → (NoErrors) → (ValidRange) → (Deduped) → (DedupedByKey) → (Derived) → (Grouped) → (composite keys)
    param([Parameter(Mandatory)]$Table, $Model = $null)
    $outTable = $Table
    if ($outTable.From -and $Model) { $Table = $Model.Tables[$outTable.From] }   # summary table: build the source table's full pipeline, then group it
    $steps = New-Object System.Collections.Generic.List[object]   # @( @{ Name; Expr } )
    $src = @($Table.Columns | Where-Object { $null -ne $_.Source })
    $comp = @($Table.Columns | Where-Object { $null -eq $_.Source -and $_.Composite })
    $skip = $Table.HeaderRow - 1
    if ($Table.FilePattern) {
        $parts = ([string]$Table.FilePattern).Split('*')
        $steps.Add(@{ Name = 'Files'; Expr = 'Folder.Files(DataFolder)' })
        $steps.Add(@{ Name = 'Matched'; Expr = 'Table.SelectRows(Files, each Text.StartsWith([Name], ' + (ConvertTo-MString $parts[0]) + ') and Text.EndsWith([Name], ' + (ConvertTo-MString $parts[1]) + '))' })
        $perFile = Get-MReader -Table $Table -Binary '[Content]'
        if ($skip -gt 0) { $perFile = 'Table.Skip(' + $perFile + ', ' + $skip + ')' }
        $steps.Add(@{ Name = 'Loaded'; Expr = 'Table.AddColumn(Matched, "Data", each Table.PromoteHeaders(' + $perFile + ', [PromoteAllScalars=true]))' })
        $steps.Add(@{ Name = 'Promoted'; Expr = 'Table.Combine(Loaded[Data])' })
    } else {
        $fileRef = 'File.Contents(DataFolder & ' + (ConvertTo-MString ('\' + [string]$Table.File)) + ')'
        $steps.Add(@{ Name = 'Source'; Expr = (Get-MReader -Table $Table -Binary $fileRef) })
        $prev = 'Source'
        if ($skip -gt 0) { $steps.Add(@{ Name = 'Skipped'; Expr = 'Table.Skip(Source, ' + $skip + ')' }); $prev = 'Skipped' }
        $steps.Add(@{ Name = 'Promoted'; Expr = 'Table.PromoteHeaders(' + $prev + ', [PromoteAllScalars=true])' })
    }
    $sel = @($src | ForEach-Object { ConvertTo-MString $_.Source }) -join ', '
    $steps.Add(@{ Name = 'Selected'; Expr = 'Table.SelectColumns(Promoted, {' + $sel + '})' })
    $prev = 'Selected'
    $clean = $Table.Clean
    if ($clean -and $clean.RemoveBlankRows) {
        $steps.Add(@{ Name = 'NoBlankRows'; Expr = 'Table.SelectRows(' + $prev + ', each not List.AllTrue(List.Transform(Record.FieldValues(_), each _ = null or Text.Trim(Text.From(_)) = "")))' }); $prev = 'NoBlankRows'
    }
    # -cne: M is case-sensitive, so a case-only rename (price -> Price) must still be emitted
    $renames = @($src | Where-Object { $_.Source -cne $_.Name } | ForEach-Object { '{' + (ConvertTo-MString $_.Source) + ', ' + (ConvertTo-MString $_.Name) + '}' })
    if ($renames.Count -gt 0) { $steps.Add(@{ Name = 'Renamed'; Expr = 'Table.RenameColumns(' + $prev + ', {' + ($renames -join ', ') + '})' }); $prev = 'Renamed' }
    # per-column value cleaning (trim / case / nullValues / valueMap / numberClean) - runs before typing, while every value is still text
    $txtClean = @($src | Where-Object { $_.Trim -or $_.Case -or $_.NumberClean -or $_.FillDown -or (@($_.NullValues).Count -gt 0) -or ($_.ValueMap -and $_.ValueMap.Count -gt 0) })
    if ($txtClean.Count -gt 0) {
        $ops = @($txtClean | ForEach-Object {
            $base = 'Text.From(_)'
            if ($_.Trim -or $_.NumberClean) { $base = 'Text.Trim(' + $base + ')' }
            if ($_.Case -eq 'upper') { $base = 'Text.Upper(' + $base + ')' } elseif ($_.Case -eq 'lower') { $base = 'Text.Lower(' + $base + ')' }
            $lets = New-Object System.Collections.Generic.List[string]
            $lets.Add('t0 = ' + $base); $cur = 't0'
            if ($_.FillDown) { $lets.Add('t1 = if ' + $cur + ' = "" then null else ' + $cur); $cur = 't1' }   # FillDown fills nulls, so empties must become null first
            if (@($_.NullValues).Count -gt 0) { $lets.Add('t2 = if ' + $cur + ' = null then null else if List.Contains({' + (@($_.NullValues | ForEach-Object { ConvertTo-MString ([string]$_) }) -join ', ') + '}, ' + $cur + ') then null else ' + $cur); $cur = 't2' }
            if ($_.ValueMap -and $_.ValueMap.Count -gt 0) {
                $rec = @($_.ValueMap.GetEnumerator() | ForEach-Object { '#' + (ConvertTo-MString ([string]$_.Key)) + ' = ' + (ConvertTo-MString ([string]$_.Value)) }) -join ', '
                $lets.Add('t3 = if ' + $cur + ' = null then null else Record.FieldOrDefault([' + $rec + '], ' + $cur + ', ' + $cur + ')'); $cur = 't3'
            }
            if ($_.NumberClean) {
                $lets.Add('t4 = if ' + $cur + ' = null then null else let s = ' + $cur + ', neg = Text.StartsWith(s, "(") and Text.EndsWith(s, ")"), c = Text.Remove(s, {"₩", "$", "€", "£", "¥", ",", "(", ")", " "}), p = Text.EndsWith(c, "%"), v = try Number.FromText(if p then Text.TrimEnd(c, "%") else c) otherwise null in if v = null then null else (if neg then -v else v) * (if p then 0.01 else 1)'); $cur = 't4'
            }
            '{' + (ConvertTo-MString $_.Name) + ', each if _ = null then null else let ' + ($lets -join ', ') + ' in ' + $cur + '}'
        })
        $steps.Add(@{ Name = 'Cleaned'; Expr = 'Table.TransformColumns(' + $prev + ', {' + ($ops -join ', ') + '})' }); $prev = 'Cleaned'
    }
    $fillCols = @($src | Where-Object { $_.FillDown })
    if ($fillCols.Count -gt 0) { $steps.Add(@{ Name = 'FilledDown'; Expr = 'Table.FillDown(' + $prev + ', {' + (@($fillCols | ForEach-Object { ConvertTo-MString $_.Name }) -join ', ') + '})' }); $prev = 'FilledDown' }
    # mixed date formats: try the default parse, then each declared format; anything else becomes null
    $dateClean = @($src | Where-Object { @($_.DateFormats).Count -gt 0 })
    if ($dateClean.Count -gt 0) {
        $ops = @($dateClean | ForEach-Object {
            $fn = if ($_.Type -eq 'datetime') { 'DateTime.FromText' } else { 'Date.FromText' }
            $expr = 'try ' + $fn + '(s) otherwise '
            foreach ($f in @($_.DateFormats)) { $expr += 'try ' + $fn + '(s, [Format=' + (ConvertTo-MString ([string]$f)) + ']) otherwise ' }
            $expr += 'null'
            '{' + (ConvertTo-MString $_.Name) + ', each if _ = null then null else let s = Text.Trim(Text.From(_)) in if s = "" then null else ' + $expr + '}'
        })
        $steps.Add(@{ Name = 'FixedDates'; Expr = 'Table.TransformColumns(' + $prev + ', {' + ($ops -join ', ') + '})' }); $prev = 'FixedDates'
    }
    $types = @($src | ForEach-Object { '{' + (ConvertTo-MString $_.Name) + ', ' + (Get-MType $_.Type) + '}' }) -join ', '
    $steps.Add(@{ Name = 'Typed'; Expr = 'Table.TransformColumnTypes(' + $prev + ', {' + $types + '})' })
    $prev = 'Typed'
    if ($clean -and $clean.ErrorsToNull) {
        $pairs = @($src | ForEach-Object { '{' + (ConvertTo-MString $_.Name) + ', null}' }) -join ', '
        $steps.Add(@{ Name = 'NoErrors'; Expr = 'Table.ReplaceErrorValues(' + $prev + ', {' + $pairs + '})' }); $prev = 'NoErrors'
    }
    # out-of-range numeric values become null (post-typing)
    $rangeCols = @($src | Where-Object { @($_.ValidRange).Count -eq 2 })
    if ($rangeCols.Count -gt 0) {
        $ops = @($rangeCols | ForEach-Object {
            $lo = ([double]$_.ValidRange[0]).ToString('R', [Globalization.CultureInfo]::InvariantCulture); $hi = ([double]$_.ValidRange[1]).ToString('R', [Globalization.CultureInfo]::InvariantCulture)
            '{' + (ConvertTo-MString $_.Name) + ', each if _ = null then null else if _ < ' + $lo + ' or _ > ' + $hi + ' then null else _}'
        })
        $steps.Add(@{ Name = 'ValidRange'; Expr = 'Table.TransformColumns(' + $prev + ', {' + ($ops -join ', ') + '})' }); $prev = 'ValidRange'
    }
    if ($clean -and $clean.DropDuplicates) {
        $steps.Add(@{ Name = 'Deduped'; Expr = 'Table.Distinct(' + $prev + ')' }); $prev = 'Deduped'
    }
    # key-level dedupe: rows that share a key but differ are resolved by rule, not dropped blindly
    if ($clean -and $clean.DedupeBy) {
        $db = $clean.DedupeBy
        $keyList = '{' + (@($db.KeyCols | ForEach-Object { ConvertTo-MString ([string]$_) }) -join ', ') + '}'
        if ($db.Keep -eq 'mostComplete') {
            $steps.Add(@{ Name = 'Scored'; Expr = 'Table.AddColumn(' + $prev + ', "_nonNulls", each List.NonNullCount(Record.FieldValues(_)), Int64.Type)' })
            $steps.Add(@{ Name = 'DedupedByKey0'; Expr = 'Table.Distinct(Table.Buffer(Table.Sort(Scored, {{"_nonNulls", Order.Descending}})), ' + $keyList + ')' })
            $steps.Add(@{ Name = 'DedupedByKey'; Expr = 'Table.RemoveColumns(DedupedByKey0, {"_nonNulls"})' }); $prev = 'DedupedByKey'
        } else {
            $dir = if ($db.Keep -eq 'last') { 'Order.Descending' } else { 'Order.Ascending' }
            $steps.Add(@{ Name = 'DedupedByKey'; Expr = 'Table.Distinct(Table.Buffer(Table.Sort(' + $prev + ', {{' + (ConvertTo-MString ([string]$db.OrderBy)) + ', ' + $dir + '}})), ' + $keyList + ')' }); $prev = 'DedupedByKey'
        }
    }
    # derived columns (M expressions over the typed, cleaned data)
    $dn = 0
    foreach ($d in @($Table.Derived)) {
        $dn++
        $steps.Add(@{ Name = "Derived$dn"; Expr = 'Table.AddColumn(' + $prev + ', ' + (ConvertTo-MString $d.Name) + ', each ' + $d.Expr + ', ' + (Get-MType $d.Type) + ')' }); $prev = "Derived$dn"
    }
    # summary table: group the source pipeline to the declared granularity
    if ($outTable.From) {
        $keys = @($outTable.GroupBy | ForEach-Object { ConvertTo-MString ([string]$_) }) -join ', '
        $aggList = @($outTable.Aggs | ForEach-Object {
            $fld = '[#' + (ConvertTo-MString ([string]$_.Column)) + ']'
            $inner = switch ([string]$_.Agg) {
                'count'         { 'Table.RowCount(_)' }
                'countDistinct' { 'List.Count(List.Distinct(' + $fld + '))' }
                'sum'           { 'List.Sum(' + $fld + ')' }
                'average'       { 'List.Average(' + $fld + ')' }
                'min'           { 'List.Min(' + $fld + ')' }
                'max'           { 'List.Max(' + $fld + ')' }
            }
            '{' + (ConvertTo-MString $_.Name) + ', each ' + $inner + ', ' + (Get-MType $_.Type) + '}'
        }) -join ', '
        $steps.Add(@{ Name = 'Grouped'; Expr = 'Table.Group(' + $prev + ', {' + $keys + '}, {' + $aggList + '})' }); $prev = 'Grouped'
        $comp = @()   # composite key columns belong to the source table, not the summary
    }
    $n = 0
    foreach ($c in $comp) {
        $n++
        $combine = 'Text.Combine({' + (@($c.Composite | ForEach-Object { 'Text.From([' + $_ + '])' }) -join ', ') + '}, "|")'
        $steps.Add(@{ Name = "WithKey$n"; Expr = 'Table.AddColumn(' + $prev + ', ' + (ConvertTo-MString $c.Name) + ', each ' + $combine + ', type text)' }); $prev = "WithKey$n"
    }
    $l = New-Object System.Collections.Generic.List[string]
    $l.Add('let')
    for ($i = 0; $i -lt $steps.Count; $i++) { $comma = if ($i -lt $steps.Count - 1) { ',' } else { '' }; $l.Add('    ' + $steps[$i].Name + ' = ' + $steps[$i].Expr + $comma) }
    $l.Add('in')
    $l.Add('    ' + $prev)
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
    param([Parameter(Mandatory)]$Table, [Parameter(Mandatory)]$Model)
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
        foreach ($line in (New-MQuery -Table $Table -Model $Model)) { $l.Add("`t`t`t`t" + $line) }
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
    $l.Add('ref expression DataFolder')
    $l.Add('')
    return ($l -join "`r`n")
}

function New-ExpressionsTmdl {
    # Shared M parameter DataFolder = folder holding the data files. Users change it via Transform data > Edit parameters when the project moves.
    param([Parameter(Mandatory)][string]$DataFolder)
    $l = New-Object System.Collections.Generic.List[string]
    $l.Add('expression DataFolder = ' + (ConvertTo-MString $DataFolder) + ' meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]')
    $l.Add("`tlineageTag: " + (New-LineageTag))
    $l.Add('')
    $l.Add("`tannotation PBI_ResultType = Text")
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
    $absRaw = (Get-Item -LiteralPath $RawDataDir).FullName.TrimEnd([char]92, [char]47)
    if ($absRaw -notmatch '^([A-Za-z]:\\|\\\\)') { $warnings.Add("data folder '$absRaw' is not a Windows path - open the project on Windows and set the DataFolder parameter (Transform data > Edit parameters)") }
    Write-Utf8File -Path (Join-Path $def 'expressions.tmdl') -Content (New-ExpressionsTmdl -DataFolder $absRaw)
    foreach ($n in $Model.Tables.Keys) {
        $t = $Model.Tables[$n]
        Write-Utf8File -Path (Join-Path (Join-Path $def 'tables') ($n + '.tmdl')) -Content (New-TableTmdl -Table $t -Model $Model)
    }
    return @{ Warnings = $warnings.ToArray(); DataFolder = $absRaw }
}
