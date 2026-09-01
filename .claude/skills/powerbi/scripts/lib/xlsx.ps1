# xlsx.ps1 — minimal .xlsx reader (zip + XML). No Excel, no modules.
# Read-XlsxWorkbook returns @{ Sheets = @( @{ Name; FirstRow; FirstCol; LastRow; LastCol; Rows = @( @{ R=<absRow>; Cells=@{ <colIndex> = <value> } } ) } ) }
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function ConvertFrom-ColumnLetters {
    param([Parameter(Mandatory)][string]$Letters)
    $n = 0
    foreach ($ch in $Letters.ToCharArray()) { $n = $n * 26 + ([int][char]::ToUpper($ch) - 64) }
    return $n
}

function Get-XmlText {
    param($Node)
    if ($null -eq $Node) { return '' }
    if ($Node -is [string]) { return $Node }
    $t = $Node.'#text'
    if ($null -ne $t) { return [string]$t }
    return [string]$Node.InnerText
}

function Read-ZipEntryText {
    param($Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { $entry = $Zip.GetEntry($Name.TrimStart('/')) }
    if ($null -eq $entry) { return $null }
    $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

function Test-DateFormatCode {
    param([string]$Code)
    if (-not $Code) { return $false }
    $clean = $Code -replace '"[^"]*"', '' -replace '\[[^\]]*\]', '' -replace '\\.', ''
    return ($clean -match '(?i)[ymdh]')
}

function Read-XlsxStyles {
    param($Zip)
    $isDate = New-Object System.Collections.Generic.List[bool]
    $xml = Read-ZipEntryText -Zip $Zip -Name 'xl/styles.xml'
    if (-not $xml) { return $isDate }
    $doc = [xml]$xml
    $custom = @{}
    if ($doc.styleSheet.numFmts -and $doc.styleSheet.numFmts.numFmt) {
        foreach ($nf in $doc.styleSheet.numFmts.numFmt) { $custom[[int]$nf.numFmtId] = [string]$nf.formatCode }
    }
    $builtin = @(14,15,16,17,18,19,20,21,22,27,28,29,30,31,32,33,34,35,36,45,46,47,50,51,52,53,54,55,56,57,58)
    if ($doc.styleSheet.cellXfs -and $doc.styleSheet.cellXfs.xf) {
        foreach ($xf in $doc.styleSheet.cellXfs.xf) {
            $id = 0; if ($xf.numFmtId) { $id = [int]$xf.numFmtId }
            $d = $false
            if ($builtin -contains $id) { $d = $true }
            elseif ($custom.ContainsKey($id)) { $d = Test-DateFormatCode $custom[$id] }
            $isDate.Add($d)
        }
    }
    return $isDate
}

function Read-XlsxSharedStrings {
    param($Zip)
    $list = New-Object System.Collections.Generic.List[string]
    $xml = Read-ZipEntryText -Zip $Zip -Name 'xl/sharedStrings.xml'
    if (-not $xml) { return $list }
    $doc = [xml]$xml
    foreach ($si in $doc.sst.si) {
        $text = ''
        if ($null -ne $si.t) { $text = Get-XmlText $si.t }
        elseif ($null -ne $si.r) { foreach ($run in $si.r) { $text += (Get-XmlText $run.t) } }
        $list.Add($text)
    }
    return $list
}

function ConvertFrom-CellRef {
    # "AB12" → @{ Col=28; Row=12 }
    param([string]$Ref)
    if ($Ref -match '^([A-Z]+)(\d+)$') { return @{ Col = (ConvertFrom-ColumnLetters $Matches[1]); Row = [int]$Matches[2] } }
    throw "bad cell ref '$Ref'"
}

function Read-XlsxSheet {
    param($Zip, [string]$EntryName, [string]$SheetName, $Shared, $IsDateStyle, [int]$MaxRows)
    $doc = [xml](Read-ZipEntryText -Zip $Zip -Name $EntryName)
    $firstRow = 1; $firstCol = 1; $lastRow = 0; $lastCol = 0
    $dim = $null
    if ($doc.worksheet.dimension) { $dim = [string]$doc.worksheet.dimension.ref }
    if ($dim) {
        $parts = $dim.Split(':')
        $a = ConvertFrom-CellRef $parts[0]; $firstRow = $a.Row; $firstCol = $a.Col
        if ($parts.Count -gt 1) { $b = ConvertFrom-CellRef $parts[1]; $lastRow = $b.Row; $lastCol = $b.Col } else { $lastRow = $a.Row; $lastCol = $a.Col }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    $sheetData = $doc.worksheet.sheetData
    if ($sheetData -and $sheetData.row) {
        foreach ($row in $sheetData.row) {
            if ($rows.Count -ge $MaxRows) { break }
            $r = [int]$row.r
            $cells = @{}
            if ($row.c) {
                foreach ($c in $row.c) {
                    $ref = ConvertFrom-CellRef ([string]$c.r)
                    $t = [string]$c.t
                    $v = $null
                    if ($null -ne $c.v) { $v = Get-XmlText $c.v }
                    $val = $null
                    switch ($t) {
                        's'         { if ($v -ne '') { $val = $Shared[[int]$v] } }
                        'str'       { $val = $v }
                        'inlineStr' { if ($c.is) { $val = Get-XmlText $c.is.t } }
                        'b'         { $val = ($v -eq '1') }
                        'e'         { $val = $null }
                        default {
                            if ($v -ne $null -and $v -ne '') {
                                $d = [double]::Parse($v, [Globalization.CultureInfo]::InvariantCulture)
                                $s = -1; if ($c.s) { $s = [int]$c.s }
                                if ($s -ge 0 -and $s -lt $IsDateStyle.Count -and $IsDateStyle[$s]) { $val = [DateTime]::FromOADate($d) } else { $val = $d }
                            }
                        }
                    }
                    if ($null -ne $val -and -not ($val -is [string] -and $val -eq '')) { $cells[$ref.Col] = $val }
                    if ($ref.Col -gt $lastCol) { $lastCol = $ref.Col }
                }
            }
            $rows.Add(@{ R = $r; Cells = $cells })
            if ($r -gt $lastRow) { $lastRow = $r }
        }
    }
    return @{ Name = $SheetName; FirstRow = $firstRow; FirstCol = $firstCol; LastRow = $lastRow; LastCol = $lastCol; Rows = $rows.ToArray() }
}

function Read-XlsxWorkbook {
    param([Parameter(Mandatory)][string]$Path, [int]$MaxRows = 2001)
    $full = (Get-Item -LiteralPath $Path).FullName
    $zip = [System.IO.Compression.ZipFile]::OpenRead($full)
    try {
        $shared = Read-XlsxSharedStrings -Zip $zip
        $isDate = Read-XlsxStyles -Zip $zip
        $wbDoc = [xml](Read-ZipEntryText -Zip $zip -Name 'xl/workbook.xml')
        $relDoc = [xml](Read-ZipEntryText -Zip $zip -Name 'xl/_rels/workbook.xml.rels')
        $targets = @{}
        foreach ($rel in $relDoc.Relationships.Relationship) {
            $target = [string]$rel.Target
            if ($target.StartsWith('/')) { $target = $target.TrimStart('/') } else { $target = 'xl/' + $target }
            $targets[[string]$rel.Id] = $target
        }
        $sheets = New-Object System.Collections.Generic.List[object]
        $rns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        foreach ($s in $wbDoc.workbook.sheets.sheet) {
            $rid = $s.GetAttribute('id', $rns)
            $sheets.Add((Read-XlsxSheet -Zip $zip -EntryName $targets[$rid] -SheetName ([string]$s.name) -Shared $shared -IsDateStyle $isDate -MaxRows $MaxRows))
        }
        return @{ Sheets = $sheets.ToArray() }
    } finally { $zip.Dispose() }
}
