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
