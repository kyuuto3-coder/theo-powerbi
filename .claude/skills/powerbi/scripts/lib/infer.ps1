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
              IsUnique = ($nonEmpty -gt 0 -and $counts.empty -eq 0 -and $distinct.Count -eq $nonEmpty); NonEmpty = $nonEmpty; Values = $distinct }
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

function ConvertTo-KeyString {
    # Canonical string for cross-file value comparison: numbers via invariant 'R', dates as yyyy-MM-dd, text trimmed. $null for empty.
    param($V)
    $k = Get-ValueKind $V
    switch ($k) {
        'empty'  { return $null }
        'number' { $d = if ($V -is [string]) { ConvertTo-NumberValue $V } else { [double]$V }; return $d.ToString('R', $script:Inv) }
        'date'   { $dt = if ($V -is [datetime]) { $V } else { ConvertTo-DateValue $V }; return $dt.ToString('yyyy-MM-dd') }
        default  { return ([string]$V).Trim() }
    }
}

function Test-IdLikeName { param([string]$Name) return ($Name -match '(?i)(^id$|id$|_id|key$|code$|코드$|번호$|키$|no$|^sku|^pk_|_pk$)') }

function Get-SubjectTokens {
    # Lower-case word tokens of a column name minus id/key noise words. Latin tokens only when -LatinOnly.
    param([string]$Name)
    $stop = @('id', 'key', 'code', 'no', 'num', 'number', 'nbr', 'pk', 'fk', 'wwi', 'sk', 'fact', 'dim', '코드', '번호', '키', '아이디', '식별자')
    $raw = [regex]::Split($Name.ToLowerInvariant(), '[^\p{L}\p{N}]+') | Where-Object { $_ -ne '' }
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($t in $raw) {
        $t2 = $t -replace '(id|key|code|no|번호|코드|키)$', ''   # customerid → customer, 지역코드 → 지역
        if ($t2 -eq '' ) { $t2 = $t }
        if ($stop -notcontains $t2 -and $stop -notcontains $t) { $tokens.Add($t2) }
    }
    return $tokens.ToArray()
}

function Test-NameTokensCompatible {
    # $true unless both names carry Latin subject words and none of them share a prefix (>= 3 chars): 'Stock Item Key' vs 'Employee Key' → $false; 'cust_no' vs 'customer_id' → $true; Korean vs English → $true (not judged).
    param([string]$A, [string]$B)
    $ta = @(Get-SubjectTokens $A | Where-Object { $_ -match '^[a-z0-9]+$' }); $tb = @(Get-SubjectTokens $B | Where-Object { $_ -match '^[a-z0-9]+$' })
    if ($ta.Count -eq 0 -or $tb.Count -eq 0) { return $true }
    foreach ($x in $ta) { foreach ($y in $tb) {
        $n = [Math]::Min([Math]::Min($x.Length, $y.Length), 3)
        if ($n -ge 3 -and $x.Substring(0, $n) -eq $y.Substring(0, $n)) { return $true }
        if ($x -eq $y) { return $true }
    } }
    return $false
}

function Get-TypeGroup { param([string]$Type) if ($Type -in 'int64', 'double', 'decimal') { return 'number' }; if ($Type -in 'date', 'datetime') { return 'date' }; return $Type }

function Get-TupleStrings {
    # Joined key strings for each row over the given column indices (rows with any empty part are skipped).
    param([Parameter(Mandatory)]$Rows, [Parameter(Mandatory)][int[]]$Indices)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($row in $Rows) {
        $parts = New-Object System.Collections.Generic.List[string]; $ok = $true
        foreach ($i in $Indices) { $v = if ($i -lt $row.Count) { ConvertTo-KeyString $row[$i] } else { $null }; if ($null -eq $v) { $ok = $false; break }; $parts.Add($v) }
        if ($ok) { $out.Add(($parts -join '|')) }
    }
    return $out.ToArray()
}

function Find-CompositeKey {
    # Smallest combination (2..MaxParts) of candidate column indices that is unique over the sampled rows. Returns int[] or $null.
    param([Parameter(Mandatory)]$Rows, [Parameter(Mandatory)][int[]]$Candidates, [int]$MaxParts = 3)
    $n = $Candidates.Count
    if ($Rows.Count -lt 2 -or $n -lt 2) { return $null }
    $combos = New-Object System.Collections.Generic.List[object]
    for ($a = 0; $a -lt $n; $a++) { for ($b = $a + 1; $b -lt $n; $b++) { $combos.Add([int[]]@($Candidates[$a], $Candidates[$b])) } }
    if ($MaxParts -ge 3) { for ($a = 0; $a -lt $n; $a++) { for ($b = $a + 1; $b -lt $n; $b++) { for ($c = $b + 1; $c -lt $n; $c++) { $combos.Add([int[]]@($Candidates[$a], $Candidates[$b], $Candidates[$c])) } } } }
    foreach ($combo in $combos) {
        $tuples = Get-TupleStrings -Rows $Rows -Indices $combo
        if ($tuples.Count -lt [Math]::Max(2, [int]($Rows.Count * 0.9))) { continue }
        $set = New-Object 'System.Collections.Generic.HashSet[string]'
        $dup = $false
        foreach ($t in $tuples) { if (-not $set.Add($t)) { $dup = $true; break } }
        if (-not $dup) { return $combo }
    }
    return $null
}

function Get-Containment {
    # Percent (0-100) of the sampled values that exist in the key set.
    param([Parameter(Mandatory)][AllowEmptyCollection()]$Sample, [Parameter(Mandatory)]$Set)
    $n = 0; $hit = 0
    foreach ($v in $Sample) { if ($null -eq $v) { continue }; $n++; if ($Set.Contains($v)) { $hit++ } }
    if ($n -eq 0) { return 0 }
    return [int][Math]::Round(100.0 * $hit / $n)
}

function Find-KeyMatches {
    # $Tables: @( @{ Name; RowCount; Rows (sampled data rows); Columns = @( @{ Name; Index; Type; IsUnique; Values (HashSet); Min; Max } ); CompositeKey = $null | @{ Indices; Names; Values } } )
    # Returns strings "many.col -> one.key  (NN% of M sampled values found[; names differ])". With -IncludeWeak returns @{ Strong; Weak } where Weak are
    # value-supported guesses that failed a soft rule (narrow range / unrelated names) and must be confirmed by the user.
    param([Parameter(Mandatory)]$Tables, [int]$MinPercent = 90, [switch]$IncludeWeak)
    $records = New-Object System.Collections.Generic.List[object]
    $composites = New-Object System.Collections.Generic.List[string]
    for ($a = 0; $a -lt $Tables.Count; $a++) {
        $ta = $Tables[$a]
        for ($b = 0; $b -lt $Tables.Count; $b++) {
            if ($a -eq $b) { continue }
            $tb = $Tables[$b]
            foreach ($ca in $ta.Columns) {
                if ($ca.IsUnique -or $ca.Values.Count -lt 2 -or $ca.Name -eq '(unnamed)') { continue }
                foreach ($cb in $tb.Columns) {
                    if (-not $cb.IsUnique -or $cb.Name -eq '(unnamed)') { continue }
                    if ((Get-TypeGroup $ca.Type) -ne (Get-TypeGroup $cb.Type)) { continue }
                    $sameName = ((Get-NormalizedName $ca.Name) -eq (Get-NormalizedName $cb.Name))
                    $reasons = New-Object System.Collections.Generic.List[string]; $coverage = 100
                    if (-not $sameName) {
                        # Different names: value overlap alone is weak evidence. Hard rules reject; soft rules demote to POSSIBLE with a reason.
                        if ($ca.Values.Count -lt 3) { continue }
                        if ($ta.RowCount -lt $tb.RowCount) { continue }
                        if (-not (((Test-IdLikeName $ca.Name) -and (Test-IdLikeName $cb.Name)) -or ($cb.Values.Count -ge 20 -and $ca.Values.Count -ge 5))) { continue }
                        if ((Get-TypeGroup $ca.Type) -eq 'number' -and $null -ne $ca.Min -and $null -ne $cb.Min) {
                            $keyRange = [double]$cb.Max - [double]$cb.Min; $colRange = [double]$ca.Max - [double]$ca.Min
                            $dense = ($keyRange -gt 0 -and $cb.Values.Count -ge 0.9 * ($keyRange + 1))
                            if ($dense -and -not (Test-IdLikeName $ca.Name)) { continue }   # dense 1..N keys "contain" every small integer column
                            if ($keyRange -gt 0) { $coverage = [int](100 * $colRange / $keyRange) }
                            if ($coverage -lt 50) { $reasons.Add(('values cover only {0}% of the key range' -f $coverage)) }
                        }
                        if (-not (Test-NameTokensCompatible $ca.Name $cb.Name)) { $reasons.Add('names unrelated') }
                    }
                    $pct = Get-Containment -Sample @($ca.Values) -Set $cb.Values
                    if (-not ($pct -ge $MinPercent -or ($sameName -and $pct -ge 50))) { continue }
                    $weak = ($reasons.Count -gt 0)
                    $line = if ($weak) { "{0}.{1} -> {2}.{3}  ({4}% of {5} sampled values found; {6} - confirm with the user)" -f $ta.Name, $ca.Name, $tb.Name, $cb.Name, $pct, $ca.Values.Count, ($reasons -join '; ') }
                            else { "{0}.{1} -> {2}.{3}  ({4}% of {5} sampled values found{6})" -f $ta.Name, $ca.Name, $tb.Name, $cb.Name, $pct, $ca.Values.Count, $(if ($sameName) { '' } else { '; names differ' }) }
                    $records.Add(@{ TableA = $ta.Name; ColA = $ca.Name; TableB = $tb.Name; RowsA = $ta.RowCount; RowsB = $tb.RowCount; SameName = $sameName; Weak = $weak; Coverage = $coverage; Line = $line })
                }
            }
            if ($null -ne $tb.CompositeKey) {
                $idx = New-Object System.Collections.Generic.List[int]; $names = New-Object System.Collections.Generic.List[string]; $ok = $true
                for ($p = 0; $p -lt $tb.CompositeKey.Names.Count; $p++) {
                    $partName = $tb.CompositeKey.Names[$p]; $partCol = $tb.Columns | Where-Object { $_.Name -eq $partName } | Select-Object -First 1
                    $match = $ta.Columns | Where-Object { (Get-NormalizedName $_.Name) -eq (Get-NormalizedName $partName) } | Select-Object -First 1
                    if ($null -eq $match -and $null -ne $partCol) {
                        $best = $null; $bestPct = 0
                        foreach ($ca in $ta.Columns) { if ((Get-TypeGroup $ca.Type) -ne (Get-TypeGroup $partCol.Type) -or $idx.Contains($ca.Index)) { continue }; $pp = Get-Containment -Sample @($ca.Values) -Set $partCol.Values; if ($pp -gt $bestPct) { $bestPct = $pp; $best = $ca } }
                        if ($bestPct -ge $MinPercent) { $match = $best }
                    }
                    if ($null -eq $match) { $ok = $false; break }
                    $idx.Add($match.Index); $names.Add($match.Name)
                }
                if ($ok) {
                    $tuples = Get-TupleStrings -Rows $ta.Rows -Indices $idx.ToArray()
                    $pct = Get-Containment -Sample $tuples -Set $tb.CompositeKey.Values
                    if ($pct -ge $MinPercent) { $composites.Add(("{0}.({1}) -> {2}.({3})  ({4}% of {5} sampled rows found; composite key)" -f $ta.Name, ($names -join '+'), $tb.Name, ($tb.CompositeKey.Names -join '+'), $pct, $tuples.Count)) }
                }
            }
        }
    }
    # --- post filters
    # 1) a column that matches a same-named key somewhere gets no different-name guesses
    $sameNameCols = @{}; foreach ($r in $records) { if ($r.SameName) { $sameNameCols[$r.TableA + '|' + $r.ColA] = $true } }
    $records = @($records | Where-Object { $_.SameName -or -not $sameNameCols.ContainsKey($_.TableA + '|' + $_.ColA) })
    # 2) a column with a strong match gets no weak guesses; weak guesses only from fact-like tables (have a strong match, >= 2x the key table), coverage >= 10%, top 2 per column
    $strongCols = @{}; $strongTables = @{}
    foreach ($r in $records) { if (-not $r.Weak) { $strongCols[$r.TableA + '|' + $r.ColA] = $true; $strongTables[$r.TableA] = $true } }
    $strong = @($records | Where-Object { -not $_.Weak } | ForEach-Object { $_.Line })
    $weakRecs = @($records | Where-Object { $_.Weak -and -not $strongCols.ContainsKey($_.TableA + '|' + $_.ColA) -and $strongTables.ContainsKey($_.TableA) -and $_.RowsA -ge 2 * $_.RowsB -and $_.Coverage -ge 10 })
    $weak = New-Object System.Collections.Generic.List[string]
    foreach ($grp in ($weakRecs | Group-Object { $_.TableA + '|' + $_.ColA })) { foreach ($r in ($grp.Group | Sort-Object { -$_.Coverage } | Select-Object -First 2)) { $weak.Add($r.Line) } }
    $out = @($strong) + @($composites)
    if ($IncludeWeak) { return @{ Strong = $out; Weak = $weak.ToArray() } }
    return $out
}

function Find-SameStructureGroups {
    # $Tables: @( @{ Name; File; Sheet; Columns } ). Groups csv files / same-named sheets whose header signature is identical. Returns @( @{ Files; Sheet; Pattern } ).
    param([Parameter(Mandatory)]$Tables)
    $groups = @{}
    foreach ($t in $Tables) {
        $sig = (($t.Columns | ForEach-Object { Get-NormalizedName $_.Name }) -join '|') + '||' + [string]$t.Sheet
        if (-not $groups.ContainsKey($sig)) { $groups[$sig] = New-Object System.Collections.Generic.List[object] }
        $groups[$sig].Add($t)
    }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($sig in $groups.Keys) {
        $members = @($groups[$sig] | Sort-Object { $_.File })
        $files = @($members | ForEach-Object { $_.File } | Select-Object -Unique)
        if ($files.Count -lt 2) { continue }
        $prefix = $files[0]; $suffix = $files[0]
        foreach ($f in $files) {
            while ($prefix.Length -gt 0 -and -not $f.StartsWith($prefix)) { $prefix = $prefix.Substring(0, $prefix.Length - 1) }
            while ($suffix.Length -gt 0 -and -not $f.EndsWith($suffix)) { $suffix = $suffix.Substring(1) }
        }
        if ($prefix.Length + $suffix.Length -ge $files[0].Length) { $suffix = [IO.Path]::GetExtension($files[0]); $prefix = '' }
        $out.Add(@{ Files = $files; Sheet = $members[0].Sheet; Pattern = ($prefix + '*' + $suffix) })
    }
    return $out.ToArray()
}
