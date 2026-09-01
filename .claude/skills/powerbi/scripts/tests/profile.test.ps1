$profileScript = Join-Path (Split-Path -Parent $libDir) 'profile-data.ps1'

Test-Case 'profile of fixtures lists both files, dictionary sheet, encoding and key match' {
    $out = (& $profileScript -DataFolder $fixtures | Out-String)
    Assert-Match 'PROFILE .*\(3 files\)' $out
    Assert-Match 'FILE "apps.xlsx"' $out
    Assert-Match 'SHEET "Data"\s+rows=40\s+header_row=1\s+first_col=B\s+-> table candidate' $out
    Assert-Match '\(unnamed\)\s+int64' $out
    Assert-Match 'id\s+int64\s+0%\s+40\s+281000000, 281000001\s+min=281000000 max=281000039 KEY' $out
    Assert-Match 'release_date\s+date\s+0%' $out
    Assert-Match 'SHEET "Data Defintion"\s+rows=8\s+-> dictionary \(not loaded\)' $out
    Assert-Match 'id=App ID; track_name=Track Name' $out
    Assert-Match 'FILE "sales_cp949.csv"\s+encoding=949\s+delimiter=","\s+rows=30\s+header_row=1' $out
    Assert-Match '주문일\s+date' $out
    Assert-Match '매출액\s+decimal' $out
    Assert-Match 'KEY MATCHES.*\n.*sales_cp949.지역코드 -> regions.지역코드' $out
}
Test-Case 'profile of an empty folder exits 2 with a clear message' {
    $empty = Join-Path $tmpRoot 'empty'; New-Item -ItemType Directory -Path $empty -Force | Out-Null
    $out = (& $profileScript -DataFolder $empty 2>&1 | Out-String)
    Assert-Match 'no \.csv/\.xlsx files' $out
    Assert-Equal 2 $LASTEXITCODE
}
