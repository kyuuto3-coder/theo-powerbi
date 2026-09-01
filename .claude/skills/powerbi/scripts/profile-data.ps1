<#
.SYNOPSIS  Profiles every .csv/.xlsx in rawdata/ into a compact text summary for Claude.
.EXAMPLE   powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/profile-data.ps1
#>
param([string]$DataFolder = '', [int]$SampleRows = 10000, [int]$MaxColumns = 200, [int]$FullReadLimit = 400000)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
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

function Get-FullValues {
    # Re-reads the whole file (up to FullReadLimit rows) and returns the key-string set of one column. Used for unique-key columns of tables larger than the sample.
    param([string]$Path, [string]$Ext, [int]$CodePage, [string]$Delimiter, [string]$Sheet, [int]$HeaderIndex, [int]$ColIndex)
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($Ext -eq '.csv') {
        $r = Read-CsvRows -Path $Path -CodePage $CodePage -Delimiter $Delimiter -MaxRows ($FullReadLimit + $HeaderIndex + 1)
        for ($i = $HeaderIndex + 1; $i -lt $r.Rows.Count; $i++) { $row = $r.Rows[$i]; if ($ColIndex -lt $row.Count) { $k = ConvertTo-KeyString $row[$ColIndex]; if ($null -ne $k) { [void]$set.Add($k) } } }
    } else {
        $wb = Read-XlsxWorkbook -Path $Path -MaxRows ($FullReadLimit + 40)
        $sh = $wb.Sheets | Where-Object { $_.Name -eq $Sheet } | Select-Object -First 1
        if ($sh) { for ($i = $HeaderIndex + 1; $i -lt $sh.Rows.Count; $i++) { $cells = $sh.Rows[$i].Cells; if ($cells.ContainsKey($ColIndex + 1)) { $k = ConvertTo-KeyString $cells[$ColIndex + 1]; if ($null -ne $k) { [void]$set.Add($k) } } } }
    }
    return $set
}

function Get-CompositeCandidates {
    # Low-cardinality dimension-like columns (ids, codes, months, categories) that could form a composite key.
    param($T)
    $c = @($T.Columns | Where-Object { $_.Name -ne '(unnamed)' -and $_.Profile.Distinct -ge 2 -and -not $_.Profile.IsUnique -and $_.Profile.Type -in 'int64', 'text', 'date' -and ($_.Profile.Type -ne 'int64' -or $_.Profile.Distinct -le [Math]::Max(2, $_.Profile.NonEmpty * 0.5)) } | Sort-Object { -$_.Profile.Distinct } | Select-Object -First 10 | Sort-Object { $_.Index })
    return [int[]]@($c | ForEach-Object { $_.Index })
}

function Register-Table {
    # Builds the matching record for Find-KeyMatches / Find-SameStructureGroups and prints the composite key line when there is no single key.
    param($T, [string]$Name, [string]$File, $Sheet, [bool]$HasSingleKey, [string]$Ext, [int]$CodePage, [string]$Delimiter, [string]$Path)
    $dataRows = Get-DataRows $T
    $cols = New-Object System.Collections.Generic.List[object]
    foreach ($c in $T.Columns) {
        $vals = $c.Profile.Values; $unique = $c.Profile.IsUnique; $min = $c.Profile.Min; $max = $c.Profile.Max
        if ($unique -and $T.DataRows -gt $T.SampleRows -and $c.Name -ne '(unnamed)') {
            $vals = Get-FullValues -Path $Path -Ext $Ext -CodePage $CodePage -Delimiter $Delimiter -Sheet ([string]$Sheet) -HeaderIndex $T.HeaderIndex -ColIndex $c.Index
            $unique = ($vals.Count -ge [Math]::Min($T.DataRows, $FullReadLimit))
            if ($c.Profile.Type -in 'int64', 'double', 'decimal') { $min = $null; $max = $null; foreach ($v in $vals) { $d = [double]::Parse($v, [Globalization.CultureInfo]::InvariantCulture); if ($null -eq $min -or $d -lt $min) { $min = $d }; if ($null -eq $max -or $d -gt $max) { $max = $d } } }
        }
        $cols.Add(@{ Name = $c.Name; Index = $c.Index; Type = $c.Profile.Type; IsUnique = $unique; Values = $vals; Min = $min; Max = $max })
    }
    $composite = $null
    if (-not $HasSingleKey) {
        $cand = Get-CompositeCandidates $T
        if ($cand.Count -ge 2) {
            $ck = Find-CompositeKey -Rows $dataRows -Candidates $cand
            if ($null -ne $ck) {
                $names = @($ck | ForEach-Object { $i = $_; ($T.Columns | Where-Object { $_.Index -eq $i } | Select-Object -First 1).Name })
                $set = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($tp in (Get-TupleStrings -Rows $dataRows -Indices $ck)) { [void]$set.Add($tp) }
                $composite = @{ Indices = $ck; Names = $names; Values = $set }
                $lines.Add('    KEY (composite): ' + ($names -join ' + ') + '   <- no single unique column; rows are unique on this combination')
            }
        }
    }
    $tables.Add(@{ Name = $Name; File = $File; Sheet = $Sheet; Rows = $dataRows; RowCount = $T.DataRows; Columns = $cols.ToArray(); CompositeKey = $composite })
}

function Get-DataRows { param($T) if ($T.Grid.Count -le $T.HeaderIndex + 1) { return @() }; return @($T.Grid[($T.HeaderIndex + 1)..($T.Grid.Count - 1)]) }

function Get-GridProfile {
    # $Grid: object[] rows (absolute column indices, 0-based). Returns table profile hashtable or $null when no header.
    param($Grid, [int]$TotalDataRows, [string]$Label)
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
        $hasHeader = ($j -lt $header.Count -and (Get-ValueKind $header[$j]) -ne 'empty')
        if ($p.NonEmpty -eq 0 -and -not $hasHeader) { continue }   # fully empty column
        $name = if ($hasHeader) { ([string]$header[$j]).Trim() } else { '(unnamed)' }
        $cols.Add(@{ Index = $j; Name = $name; Profile = $p })
    }
    return @{ Label = $Label; HeaderIndex = $h; FirstCol = $firstCol; Columns = $cols.ToArray(); DataRows = $TotalDataRows; SampleRows = ($Grid.Count - $h - 1); Grid = $Grid }
}

function Add-TableLines {
    param($T, [string]$ColPrefixMode)   # 'letters' for xlsx, 'index' for csv
    $script:LastHadKey = $false
    $lines.Add(("    {0,-4} {1,-24} {2,-9} {3,-6} {4,-8} {5,-28} {6}" -f 'col', 'name', 'type', 'nulls', 'distinct', 'sample', 'note'))
    $shown = 0
    $firstUniqueInt = -1   # fallback KEY: a unique integer only when it is the first real column (typical id position)
    foreach ($c in $T.Columns) { if ($c.Name -eq '(unnamed)') { continue }; if ($c.Profile.IsUnique -and $c.Profile.Type -eq 'int64') { $firstUniqueInt = $c.Index }; break }
    foreach ($c in $T.Columns) {
        if ($shown -ge $MaxColumns) { $lines.Add(("    +{0} more columns" -f ($T.Columns.Count - $shown))); break }
        $shown++
        $p = $c.Profile
        $colLabel = if ($ColPrefixMode -eq 'letters') { ConvertTo-ColumnLetters ($c.Index + 1) } else { [string]($c.Index + 1) }
        $notes = New-Object System.Collections.Generic.List[string]
        if ($null -ne $p.Min -and $null -ne $p.Max) { $notes.Add(("min={0} max={1}" -f $p.Min, $p.Max)) }
        if ($c.Name -eq '(unnamed)') { $notes.Add('index? drop') }
        elseif ($p.IsUnique -and (($p.Type -in 'int64', 'text' -and ($c.Index -eq $firstUniqueInt -or $c.Name -match '(?i)(^id$|id$|_id|key$|code$|코드$|번호$|키$|no$)')) -or $p.Type -eq 'date')) {
            if ($T.SampleRows -ge $T.DataRows) { $notes.Add('KEY') } else { $notes.Add('KEY?') } }
        $sample = (@($p.Samples | ForEach-Object { if ($p.Type -eq 'text') { '"' + (Format-Cell $_) + '"' } else { Format-Cell $_ } }) -join ', ')
        $lines.Add(("    {0,-4} {1,-24} {2,-9} {3,-6} {4,-8} {5,-28} {6}" -f $colLabel, (Format-Cell $c.Name), $p.Type, ("{0}%" -f $p.NullPct), $p.Distinct, $sample, ($notes -join ' ')))
        if ($notes -contains 'KEY' -or $notes -contains 'KEY?') { $script:LastHadKey = $true }
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
        $t = Get-GridProfile -Grid $grid -TotalDataRows $dataRows -Label $f.Name
        $delimLabel = if ($delim -eq "`t") { '\t' } else { $delim }
        if ($null -eq $t) { $lines.Add(("FILE ""{0}""  encoding={1}  -> no header row found (skipped)" -f $f.Name, $cp)); continue }
        $lines.Add(("FILE ""{0}""  encoding={1}  delimiter=""{2}""  rows={3}  header_row={4}  -> table candidate" -f $f.Name, $cp, $delimLabel, $t.DataRows, ($t.HeaderIndex + 1)))
        Add-TableLines -T $t -ColPrefixMode 'index'
        Register-Table -T $t -Name ([IO.Path]::GetFileNameWithoutExtension($f.Name)) -File $f.Name -Sheet $null -HasSingleKey $script:LastHadKey -Ext '.csv' -CodePage $cp -Delimiter $delim -Path $f.FullName
    }
    else {
        $wb = Read-XlsxWorkbook -Path $f.FullName -MaxRows ($SampleRows + 40)
        $lines.Add(("FILE ""{0}""" -f $f.Name))
        $profiles = @{}
        foreach ($s in $wb.Sheets) {
            $grid = @()
            foreach ($row in $s.Rows) { $arr = New-Object object[] ([Math]::Max(1, $s.LastCol)); foreach ($k in $row.Cells.Keys) { $arr[$k - 1] = $row.Cells[$k] }; $grid += ,$arr }
            if ($grid.Count -eq 0) { $profiles[$s.Name] = $null; continue }
            $h = Find-HeaderRow -Grid $grid
            $absHeaderRow = if ($h -ge 0) { $s.Rows[$h].R } else { 0 }
            $dataRows = if ($h -ge 0) { $s.LastRow - $absHeaderRow } else { 0 }
            $t = Get-GridProfile -Grid $grid -TotalDataRows $dataRows -Label $s.Name
            if ($null -ne $t) { $t.HeaderRowRelative = $absHeaderRow - $s.FirstRow + 1 }
            $profiles[$s.Name] = $t
        }
        $allHeaders = @(); foreach ($k in $profiles.Keys) { if ($profiles[$k]) { $allHeaders += @($profiles[$k].Columns | ForEach-Object { $_.Name }) } }
        foreach ($s in $wb.Sheets) {
            $t = $profiles[$s.Name]
            if ($null -eq $t) { $lines.Add(("  SHEET ""{0}""  -> empty (skipped)" -f $s.Name)); continue }
            $headers = @($t.Columns | ForEach-Object { $_.Name })
            $others = @($allHeaders | Where-Object { $headers -notcontains $_ })
            $dataRowsArr = Get-DataRows $t
            $firstColVals = @(); if ($t.Columns.Count -gt 0) { $ci = $t.Columns[0].Index; foreach ($row in $dataRowsArr) { if ($ci -lt $row.Count -and $null -ne $row[$ci]) { $firstColVals += [string]$row[$ci] } } }
            if (Test-DictionarySheet -Headers $headers -DataRowCount $t.DataRows -OtherHeaders $others -FirstColumnValues $firstColVals) {
                $lines.Add(("  SHEET ""{0}""  rows={1}  -> dictionary (not loaded)" -f $s.Name, $t.DataRows))
                $pairs = New-Object System.Collections.Generic.List[string]
                $c0 = $t.Columns[0].Index; $c1 = if ($t.Columns.Count -gt 1) { $t.Columns[1].Index } else { -1 }
                foreach ($row in $dataRowsArr) {
                    $k = if ($c0 -lt $row.Count) { [string]$row[$c0] } else { '' }
                    $v = if ($c1 -ge 0 -and $c1 -lt $row.Count) { [string]$row[$c1] } else { '' }
                    if ($k.Trim()) { $pairs.Add(("{0}={1}" -f $k.Trim(), $v.Trim())) }
                }
                $lines.Add('    ' + ($pairs -join '; '))
                continue
            }
            $lines.Add(("  SHEET ""{0}""  rows={1}  header_row={2}  first_col={3}  -> table candidate" -f $s.Name, $t.DataRows, $t.HeaderRowRelative, (ConvertTo-ColumnLetters ($t.FirstCol + 1))))
            Add-TableLines -T $t -ColPrefixMode 'letters'
            Register-Table -T $t -Name $s.Name -File $f.Name -Sheet $s.Name -HasSingleKey $script:LastHadKey -Ext $f.Extension -CodePage 0 -Delimiter '' -Path $f.FullName
        }
    }
}
if ($tables.Count -ge 2) {
    $groups = @(Find-SameStructureGroups -Tables $tables.ToArray())
    if ($groups.Count -gt 0) {
        $lines.Add('GROUPS (identical columns -> load as ONE table with "filePattern"; new files matching the pattern are picked up on Refresh):')
        foreach ($g in $groups) { $sheetNote = if ($g.Sheet) { '  sheet "' + $g.Sheet + '"' } else { '' }; $lines.Add('    filePattern "' + $g.Pattern + '"' + $sheetNote + '  <- ' + ($g.Files -join ', ')) }
    }
    $km = Find-KeyMatches -Tables $tables.ToArray() -IncludeWeak
    $lines.Add('KEY MATCHES (many-side column -> unique key; % = share of sampled values found on the key side; >= 90% is a safe relationship):')
    if ($km.Strong.Count -eq 0) { $lines.Add('    none - treat tables as independent (flat) unless the user says otherwise') }
    foreach ($m in $km.Strong) { $lines.Add('    ' + $m) }
    if ($km.Weak.Count -gt 0) {
        $lines.Add('POSSIBLE MATCHES (values agree but the column names are unrelated - ask the user before using):')
        foreach ($m in $km.Weak) { $lines.Add('    ' + $m) }
    }
}
$lines.Add('NOTE header_row is relative to the sheet''s used range (what Power BI sees); for csv it is the physical line number.')
$lines | ForEach-Object { $_ }
exit 0
