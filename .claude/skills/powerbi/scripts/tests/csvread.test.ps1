. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'csvread.ps1')

Test-Case 'Get-CsvEncoding detects CP949 and UTF-8' {
    Assert-Equal 949   (Get-CsvEncoding -Path (Join-Path $fixtures 'sales_cp949.csv'))
    Assert-Equal 65001 (Get-CsvEncoding -Path (Join-Path $fixtures 'regions.csv'))
}
Test-Case 'Get-CsvEncoding detects UTF-8 BOM' {
    $p = Join-Path $tmpRoot 'bom.csv'
    [IO.File]::WriteAllText($p, "a,b`n1,2`n", (New-Object System.Text.UTF8Encoding($true)))
    Assert-Equal 65001 (Get-CsvEncoding -Path $p)
}
Test-Case 'Get-CsvDelimiter picks the most frequent candidate on the header line' {
    Assert-Equal ',' (Get-CsvDelimiter -Path (Join-Path $fixtures 'regions.csv') -CodePage 65001)
    $p = Join-Path $tmpRoot 'semi.csv'; [IO.File]::WriteAllText($p, "a;b;c`n1;2;3`n")
    Assert-Equal ';' (Get-CsvDelimiter -Path $p -CodePage 65001)
}
Test-Case 'Read-CsvRows decodes Korean headers from CP949 and counts all rows' {
    $r = Read-CsvRows -Path (Join-Path $fixtures 'sales_cp949.csv') -CodePage 949 -Delimiter ',' -MaxRows 10
    Assert-Equal '주문일' $r.Rows[0][0]
    Assert-Equal '매출액' $r.Rows[0][4]
    Assert-Equal 10 $r.Rows.Count
    Assert-Equal 31 $r.TotalLines
}
Test-Case 'Read-CsvRows honours quoted fields with embedded delimiters' {
    $p = Join-Path $tmpRoot 'q.csv'; [IO.File]::WriteAllText($p, "name,note`n""Doe, John"",""says """"hi""""""`n")
    $r = Read-CsvRows -Path $p -CodePage 65001 -Delimiter ',' -MaxRows 10
    Assert-Equal 'Doe, John' $r.Rows[1][0]
    Assert-Equal 'says "hi"' $r.Rows[1][1]
}
