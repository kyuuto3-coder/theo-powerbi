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
    $v12 = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($v in '1', '2') { [void]$v12.Add($v) }
    $v1234 = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($v in '1', '2', '3', '4') { [void]$v1234.Add($v) }
    $qty = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($v in '7', '8', '9') { [void]$qty.Add($v) }
    $t1 = @{ Name = 'sales';   Rows = @(); RowCount = 30; Columns = @( @{ Name = '지역코드'; Index = 0; Type = 'int64'; IsUnique = $false; Values = $v12 }, @{ Name = '수량'; Index = 1; Type = 'int64'; IsUnique = $false; Values = $qty } ); CompositeKey = $null }
    $t2 = @{ Name = 'regions'; Rows = @(); RowCount = 4; Columns = @( @{ Name = '지역코드'; Index = 0; Type = 'int64'; IsUnique = $true; Values = $v1234 } ); CompositeKey = $null }
    $m = @(Find-KeyMatches -Tables @($t1, $t2))
    Assert-Equal 1 $m.Count
    Assert-Equal 'sales.지역코드 -> regions.지역코드  (100% of 2 sampled values found)' $m[0]
}

Test-Case 'ConvertTo-KeyString canonicalises numbers, dates, text' {
    Assert-Equal '1001' (ConvertTo-KeyString '1,001'); Assert-Equal '1001' (ConvertTo-KeyString 1001.0)
    Assert-Equal '2025-01-05' (ConvertTo-KeyString '2025/1/5'); Assert-Equal 'A' (ConvertTo-KeyString ' A ')
    Assert-True ($null -eq (ConvertTo-KeyString '')) 'empty → null'
}
Test-Case 'Find-CompositeKey finds the smallest unique column combination' {
    $rows = @(); foreach ($m in 1, 2) { foreach ($r in 1, 2, 3) { $rows += ,@($m, $r, 'x', 10) } }
    $k = Find-CompositeKey -Rows $rows -Candidates @(0, 1, 2)
    Assert-Equal '0,1' ($k -join ',')
    $rows2 = @(); $rows2 += ,@(1, 1, 'a'); $rows2 += ,@(1, 1, 'a')
    Assert-True ($null -eq (Find-CompositeKey -Rows $rows2 -Candidates @(0, 1, 2))) 'no unique combo'
}
Test-Case 'Get-Containment and value-based Find-KeyMatches with different names and composite keys' {
    $set = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($v in '1001', '1002', '1003') { [void]$set.Add($v) }
    Assert-Equal 100 (Get-Containment -Sample @('1001', '1003') -Set $set)
    Assert-Equal 50 (Get-Containment -Sample @('1001', '9') -Set $set)
    $custVals = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($v in '1001', '1002', '1003') { [void]$custVals.Add($v) }
    $orderVals = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($v in '1001', '1002', '1003') { [void]$orderVals.Add($v) }
    $orders = @{ Name = 'orders'; Rows = @(); RowCount = 50; Columns = @( @{ Name = 'cust_no'; Index = 0; Type = 'int64'; IsUnique = $false; Values = $orderVals; Min = 1001; Max = 1003 } ); CompositeKey = $null }
    $custs  = @{ Name = 'customers'; Rows = @(); RowCount = 20; Columns = @( @{ Name = '고객ID'; Index = 0; Type = 'int64'; IsUnique = $true; Values = $custVals; Min = 1001; Max = 1003 } ); CompositeKey = $null }
    $m = @(Find-KeyMatches -Tables @($orders, $custs))
    Assert-Equal 1 $m.Count; Assert-Match '^orders\.cust_no -> customers\.고객ID  \(100% of 3 sampled values found; names differ\)$' $m[0]
    # composite: targets unique on (월,지역코드); sales rows reference them
    $tRows = @(); foreach ($mo in 1, 2) { foreach ($r in 1, 2) { $tRows += ,@($mo, $r, 100) } }
    $tk = Find-CompositeKey -Rows $tRows -Candidates @(0, 1)
    $tkVals = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($s in (Get-TupleStrings -Rows $tRows -Indices $tk)) { [void]$tkVals.Add($s) }
    $mVals = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($v in '1', '2') { [void]$mVals.Add($v) }
    $rVals = New-Object 'System.Collections.Generic.HashSet[string]'; foreach ($v in '1', '2') { [void]$rVals.Add($v) }
    $targets = @{ Name = 'targets'; Rows = $tRows; RowCount = 4; Columns = @( @{ Name = '월'; Index = 0; Type = 'int64'; IsUnique = $false; Values = $mVals }, @{ Name = '지역코드'; Index = 1; Type = 'int64'; IsUnique = $false; Values = $rVals } ); CompositeKey = @{ Indices = $tk; Names = @('월', '지역코드'); Values = $tkVals } }
    $sRows = @(); $sRows += ,@(1, 2, 5); $sRows += ,@(2, 1, 7); $sRows += ,@(2, 2, 9)
    $sales = @{ Name = 'sales'; Rows = $sRows; RowCount = 3; Columns = @( @{ Name = '월'; Index = 0; Type = 'int64'; IsUnique = $false; Values = $mVals }, @{ Name = '지역코드'; Index = 1; Type = 'int64'; IsUnique = $false; Values = $rVals } ); CompositeKey = $null }
    $m2 = @(Find-KeyMatches -Tables @($sales, $targets))
    Assert-True (($m2 -join "`n") -match 'sales\.\(월\+지역코드\) -> targets\.\(월\+지역코드\)  \(100% of 3 sampled rows found; composite key\)') ($m2 -join ' | ')
}
Test-Case 'Find-SameStructureGroups groups identical headers into a file pattern' {
    $cols = @( @{ Name = 'order_id' }, @{ Name = 'amount' } )
    $g = @(Find-SameStructureGroups -Tables @( @{ Name = 'orders_2025-01'; File = 'orders_2025-01.csv'; Sheet = $null; Columns = $cols }, @{ Name = 'orders_2025-02'; File = 'orders_2025-02.csv'; Sheet = $null; Columns = $cols }, @{ Name = 'x'; File = 'x.csv'; Sheet = $null; Columns = @( @{ Name = 'a' } ) } ))
    Assert-Equal 1 $g.Count; Assert-Equal 'orders_2025-0*.csv' $g[0].Pattern; Assert-Equal 2 $g[0].Files.Count
}

Test-Case 'Find-KeyMatches rejects small integers that merely fall inside a dense surrogate key range' {
    $saleKeys = New-Object 'System.Collections.Generic.HashSet[string]'; 1..2000 | ForEach-Object { [void]$saleKeys.Add([string]$_) }
    $days = New-Object 'System.Collections.Generic.HashSet[string]'; 1..31 | ForEach-Object { [void]$days.Add([string]$_) }
    $dim  = @{ Name = 'DimDate'; Rows = @(); RowCount = 1460; Columns = @( @{ Name = 'Day Number'; Index = 0; Type = 'int64'; IsUnique = $false; Values = $days; Min = 1; Max = 31 } ); CompositeKey = $null }
    $fact = @{ Name = 'FactSale'; Rows = @(); RowCount = 228265; Columns = @( @{ Name = 'Sale Key'; Index = 0; Type = 'int64'; IsUnique = $true; Values = $saleKeys; Min = 1; Max = 2000 } ); CompositeKey = $null }
    Assert-Equal 0 (@(Find-KeyMatches -Tables @($dim, $fact))).Count
}

Test-Case 'Find-KeyMatches: a column with a same-name key gets no different-name guesses; dense keys need id-like names' {
    $cust = New-Object 'System.Collections.Generic.HashSet[string]'; 1..400 | ForEach-Object { [void]$cust.Add([string]$_) }
    $item = New-Object 'System.Collections.Generic.HashSet[string]'; 1..300 | ForEach-Object { [void]$item.Add([string]$_) }
    $fk   = New-Object 'System.Collections.Generic.HashSet[string]'; 1..250 | ForEach-Object { [void]$fk.Add([string]$_) }
    $qty  = New-Object 'System.Collections.Generic.HashSet[string]'; 1..350 | ForEach-Object { [void]$qty.Add([string]$_) }
    $fact = @{ Name = 'FactSale'; Rows = @(); RowCount = 100000; CompositeKey = $null; Columns = @(
        @{ Name = 'Stock Item Key'; Index = 0; Type = 'int64'; IsUnique = $false; Values = $fk; Min = 1; Max = 250 },
        @{ Name = 'Quantity'; Index = 1; Type = 'int64'; IsUnique = $false; Values = $qty; Min = 1; Max = 350 } ) }
    $dCust = @{ Name = 'DimCustomer'; Rows = @(); RowCount = 400; CompositeKey = $null; Columns = @( @{ Name = 'Customer Key'; Index = 0; Type = 'int64'; IsUnique = $true; Values = $cust; Min = 1; Max = 400 } ) }
    $dItem = @{ Name = 'DimStockItem'; Rows = @(); RowCount = 300; CompositeKey = $null; Columns = @( @{ Name = 'Stock Item Key'; Index = 0; Type = 'int64'; IsUnique = $true; Values = $item; Min = 1; Max = 300 } ) }
    $m = @(Find-KeyMatches -Tables @($fact, $dCust, $dItem))
    Assert-Equal 1 $m.Count ($m -join ' | ')
    Assert-Match '^FactSale\.Stock Item Key -> DimStockItem\.Stock Item Key' $m[0]
}

Test-Case 'Test-NameTokensCompatible judges subject words, not id/key noise, and skips cross-script names' {
    Assert-True (Test-NameTokensCompatible 'cust_no' 'customer_id') 'cust ~ customer'
    Assert-True (Test-NameTokensCompatible 'Bill To Customer Key' 'Customer Key') 'shared customer'
    Assert-True (Test-NameTokensCompatible 'Invoice Date Key' 'Date') 'shared date'
    Assert-True (Test-NameTokensCompatible 'cust_no' '고객ID') 'cross-script not judged'
    Assert-True (-not (Test-NameTokensCompatible 'WWI Stock Item ID' 'Employee Key')) 'stock item vs employee'
    Assert-True (-not (Test-NameTokensCompatible 'Latest Recorded Population' 'Sale Key')) 'population vs sale'
}
Test-Case 'Find-KeyMatches -IncludeWeak lists value-strong but name-unrelated pairs separately' {
    $emp = New-Object 'System.Collections.Generic.HashSet[string]'; 1..212 | ForEach-Object { [void]$emp.Add([string]$_) }
    $sp  = New-Object 'System.Collections.Generic.HashSet[string]'; 5, 60, 120, 180, 200 | ForEach-Object { [void]$sp.Add([string]$_) }
    $fact = @{ Name = 'FactSale'; Rows = @(); RowCount = 100000; CompositeKey = $null; Columns = @( @{ Name = 'Salesperson Key'; Index = 0; Type = 'int64'; IsUnique = $false; Values = $sp; Min = 5; Max = 200 } ) }
    $dim  = @{ Name = 'DimEmployee'; Rows = @(); RowCount = 212; CompositeKey = $null; Columns = @( @{ Name = 'Employee Key'; Index = 0; Type = 'int64'; IsUnique = $true; Values = $emp; Min = 1; Max = 212 } ) }
    $r = Find-KeyMatches -Tables @($fact, $dim) -IncludeWeak
    Assert-Equal 0 $r.Strong.Count; Assert-Equal 1 $r.Weak.Count
    Assert-Match '^FactSale\.Salesperson Key -> DimEmployee\.Employee Key  \(100% of 5 sampled values found; names unrelated' $r.Weak[0]
}
