$profileScript = Join-Path (Split-Path -Parent $libDir) 'profile-data.ps1'

Test-Case 'profile of fixtures lists both files, dictionary sheet, encoding and key match' {
    $out = (& $profileScript -DataFolder $fixtures | Out-String)
    Assert-Match 'PROFILE .*\(8 files\)' $out
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
    Assert-Match 'sales_cp949\.지역코드 -> regions\.지역코드  \(100% of 4 sampled values found\)' $out
    Assert-Match 'orders_2025-01\.cust_no -> customers\.고객ID  \(100% of \d+ sampled values found; names differ\)' $out
    Assert-Match 'FILE "targets\.csv"[\s\S]*?KEY \(composite\): 월 \+ 지역코드' $out
    Assert-Match 'sales_monthly\.\(월\+지역코드\) -> targets\.\(월\+지역코드\)  \(100% of \d+ sampled rows found; composite key\)' $out
    Assert-Match 'GROUPS \(identical columns' $out
    Assert-Match 'filePattern "orders_2025-0\*\.csv"  <- orders_2025-01\.csv, orders_2025-02\.csv' $out
    Assert-True (-not ($out -match 'sales_monthly[\s\S]*?KEY \(composite\)[\s\S]*?FILE "targets')) 'no spurious composite key on the fact table'
}
Test-Case 'profile of an empty folder exits 2 with a clear message' {
    $empty = Join-Path $tmpRoot 'empty'; New-Item -ItemType Directory -Path $empty -Force | Out-Null
    $out = (& $profileScript -DataFolder $empty 2>&1 | Out-String)
    Assert-Match 'no \.csv/\.xlsx files' $out
    Assert-Equal 2 $LASTEXITCODE
}
