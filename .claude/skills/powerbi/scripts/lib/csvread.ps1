# csvread.ps1 — encoding + delimiter detection and a bounded CSV reader.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic

function Get-CsvEncoding {
    # Returns 65001 (UTF-8) or 949 (CP949 / EUC-KR family). Heuristic: BOM → UTF-8; strict UTF-8 decode of the first 64 KB; else 949.
    param([Parameter(Mandatory)][string]$Path)
    $fs = [IO.File]::OpenRead($Path)
    try {
        $buf = New-Object byte[] 65536
        $n = $fs.Read($buf, 0, $buf.Length)
    } finally { $fs.Dispose() }
    if ($n -ge 3 -and $buf[0] -eq 0xEF -and $buf[1] -eq 0xBB -and $buf[2] -eq 0xBF) { return 65001 }
    $strict = New-Object System.Text.UTF8Encoding($false, $true)
    foreach ($trim in 0, 1, 2, 3) {
        $len = $n - $trim
        if ($len -le 0) { break }
        try { [void]$strict.GetString($buf, 0, $len); return 65001 } catch { }
    }
    return 949
}

function Get-CsvDelimiter {
    param([Parameter(Mandatory)][string]$Path, [int]$CodePage = 65001)
    $sr = New-Object IO.StreamReader($Path, [Text.Encoding]::GetEncoding($CodePage))
    try { $line = $sr.ReadLine() } finally { $sr.Dispose() }
    if ($null -eq $line) { return ',' }
    $best = ','; $bestCount = -1
    foreach ($d in ',', ';', "`t", '|') {
        $count = ($line.ToCharArray() | Where-Object { $_ -eq $d }).Count
        if ($count -gt $bestCount) { $best = $d; $bestCount = $count }
    }
    return $best
}

function Read-CsvRows {
    # Returns @{ Rows = @( string[] ... ) (first MaxRows physical records incl. header); TotalLines = number of non-empty lines in the file }
    param([Parameter(Mandatory)][string]$Path, [int]$CodePage = 65001, [string]$Delimiter = ',', [int]$MaxRows = 2001)
    $enc = [Text.Encoding]::GetEncoding($CodePage)
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, $enc)
    $rows = New-Object System.Collections.Generic.List[object]
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters($Delimiter)
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        while (-not $parser.EndOfData -and $rows.Count -lt $MaxRows) {
            $fields = $parser.ReadFields()
            $rows.Add([string[]]$fields)
        }
    } finally { $parser.Close() }
    $total = 0
    $sr = New-Object IO.StreamReader($Path, $enc)
    try { while ($null -ne ($l = $sr.ReadLine())) { if ($l.Trim().Length -gt 0) { $total++ } } } finally { $sr.Dispose() }
    return @{ Rows = $rows.ToArray(); TotalLines = $total }
}
