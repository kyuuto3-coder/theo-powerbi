. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'infer.ps1')

Test-Case 'Get-ValueKind classifies typed and string values' {
    Assert-Equal 'empty'  (Get-ValueKind $null)
    Assert-Equal 'empty'  (Get-ValueKind '  ')
    Assert-Equal 'number' (Get-ValueKind 3.5)
    Assert-Equal 'number' (Get-ValueKind '1,234.50')
    Assert-Equal 'date'   (Get-ValueKind ([datetime]'2025-01-02'))
    Assert-Equal 'date'   (Get-ValueKind '2025-01-02')
    Assert-Equal 'date'   (Get-ValueKind '2025/1/2 13:45')
    Assert-Equal 'bool'   (Get-ValueKind 'TRUE')
    Assert-Equal 'text'   (Get-ValueKind 'Games')
}
Test-Case 'Find-HeaderRow finds first mostly-text row with ≥2 cells' {
    $grid = @(); $grid += ,@($null, $null, $null); $grid += ,@($null, 'Name', 'Description'); $grid += ,@($null, 'id', 'App ID')
    Assert-Equal 1 (Find-HeaderRow -Grid $grid)
    $grid2 = @(); $grid2 += ,@('id', 'name', 'price'); $grid2 += ,@(1, 'a', 2.5)
    Assert-Equal 0 (Find-HeaderRow -Grid $grid2)
}
Test-Case 'Get-ColumnProfile infers int64 / decimal / double / date / text with stats' {
    $ints  = Get-ColumnProfile -Values @(1, 2, 3, 3)
    Assert-Equal 'int64' $ints.Type; Assert-Equal 3 $ints.Distinct; Assert-Equal 1 $ints.Min; Assert-Equal 3 $ints.Max
    $dec   = Get-ColumnProfile -Values @('0.99', '2.5', $null)
    Assert-Equal 'decimal' $dec.Type; Assert-Equal 33 $dec.NullPct
    $dbl   = Get-ColumnProfile -Values @(0.123456, 1.5)
    Assert-Equal 'double' $dbl.Type
    $dates = Get-ColumnProfile -Values @('2025-01-01', '2025-02-01')
    Assert-Equal 'date' $dates.Type; Assert-Equal '2025-01-01' $dates.Min
    $text  = Get-ColumnProfile -Values @('a', 'b', 'a')
    Assert-Equal 'text' $text.Type; Assert-Equal 2 $text.Distinct; Assert-Equal 'a' $text.Samples[0]
}
Test-Case 'Get-ColumnProfile flags unique non-null columns as key candidates' {
    $k = Get-ColumnProfile -Values @(10, 11, 12)
    Assert-True $k.IsUnique 'unique ints'
    $nk = Get-ColumnProfile -Values @(10, 10, 12)
    Assert-True (-not $nk.IsUnique) 'duplicates'
}
Test-Case 'Test-DictionarySheet recognises Name/Description sheets' {
    Assert-True (Test-DictionarySheet -Headers @('Name', 'Description') -DataRowCount 8 -OtherHeaders @())
    Assert-True (Test-DictionarySheet -Headers @('항목', '설명') -DataRowCount 8 -OtherHeaders @())
    Assert-True (Test-DictionarySheet -Headers @('col', 'meaning') -DataRowCount 5 -OtherHeaders @('id','track_name','price') -FirstColumnValues @('id','track_name','price','x'))
    Assert-True (-not (Test-DictionarySheet -Headers @('id', 'name', 'price') -DataRowCount 500 -OtherHeaders @()))
}
Test-Case 'Find-KeyMatches pairs same-named columns where one side is unique' {
    $t1 = @{ Name = 'sales';   Columns = @( @{ Name = '지역코드'; IsUnique = $false }, @{ Name = '수량'; IsUnique = $false } ) }
    $t2 = @{ Name = 'regions'; Columns = @( @{ Name = '지역코드'; IsUnique = $true } ) }
    $m = @(Find-KeyMatches -Tables @($t1, $t2))
    Assert-Equal 1 $m.Count
    Assert-Equal 'sales.지역코드 -> regions.지역코드' $m[0]
}
