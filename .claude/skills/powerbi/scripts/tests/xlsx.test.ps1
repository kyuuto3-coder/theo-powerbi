. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'xlsx.ps1')

Test-Case 'ConvertFrom-ColumnLetters maps A, Z, AA, AB' {
    Assert-Equal 1 (ConvertFrom-ColumnLetters 'A'); Assert-Equal 26 (ConvertFrom-ColumnLetters 'Z')
    Assert-Equal 27 (ConvertFrom-ColumnLetters 'AA'); Assert-Equal 28 (ConvertFrom-ColumnLetters 'AB')
}
Test-Case 'Read-XlsxWorkbook lists sheets in workbook order with dimensions' {
    $wb = Read-XlsxWorkbook -Path (Join-Path $fixtures 'apps.xlsx')
    Assert-Equal 2 $wb.Sheets.Count
    Assert-Equal 'Data' $wb.Sheets[0].Name
    Assert-Equal 'Data Defintion' $wb.Sheets[1].Name
    Assert-Equal 1 $wb.Sheets[0].FirstRow; Assert-Equal 1 $wb.Sheets[0].FirstCol
    Assert-Equal 41 $wb.Sheets[0].LastRow; Assert-Equal 9 $wb.Sheets[0].LastCol
    Assert-Equal 2 $wb.Sheets[1].FirstRow; Assert-Equal 2 $wb.Sheets[1].FirstCol
}
Test-Case 'cells carry typed values: shared strings, numbers, dates' {
    $wb = Read-XlsxWorkbook -Path (Join-Path $fixtures 'apps.xlsx')
    $data = $wb.Sheets[0]
    $hdr = $data.Rows[0]
    Assert-Equal 1 $hdr.R
    Assert-True ($null -eq $hdr.Cells[1]) 'A1 is empty'
    Assert-Equal 'id' $hdr.Cells[2]
    $r2 = $data.Rows[1]
    Assert-Equal 0 $r2.Cells[1]
    Assert-True ($r2.Cells[2] -is [double]) 'id is numeric'
    Assert-True ($r2.Cells[9] -is [datetime]) "release_date should be datetime, got $($r2.Cells[9].GetType().Name)"
    Assert-True ($r2.Cells[9].Year -ge 2016) 'date decoded from serial'
}
Test-Case 'MaxRows limits rows read but LastRow still reports the sheet size' {
    $wb = Read-XlsxWorkbook -Path (Join-Path $fixtures 'apps.xlsx') -MaxRows 5
    Assert-Equal 5 $wb.Sheets[0].Rows.Count
    Assert-Equal 41 $wb.Sheets[0].LastRow
}
