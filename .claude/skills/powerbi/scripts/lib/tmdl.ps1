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
