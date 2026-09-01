$buildScript = Join-Path (Split-Path -Parent $libDir) 'build-pbip.ps1'
$flatOut = Join-Path $tmpRoot 'flat'

Test-Case 'build of spec-flat produces a complete PBIP project' {
    $log = (& $buildScript -Spec (Join-Path $fixtures 'spec-flat.json') -RawData $fixtures -OutputDir $flatOut | Out-String)
    Assert-Equal 0 $LASTEXITCODE $log
    Assert-Match 'BUILT .*AppsFlat\.pbip\s+\(tables=2 measures=3 pages=2 visuals=11\)' $log
    foreach ($f in 'AppsFlat.pbip', '.gitignore', 'AppsFlat.Report/definition.pbir', 'AppsFlat.Report/definition/report.json', 'AppsFlat.Report/definition/version.json',
                   'AppsFlat.Report/definition/pages/pages.json', 'AppsFlat.Report/StaticResources/SharedResources/BaseThemes/Fluent2-CY26SU08.json',
                   'AppsFlat.SemanticModel/definition.pbism', 'AppsFlat.SemanticModel/definition/database.tmdl', 'AppsFlat.SemanticModel/definition/model.tmdl',
                   'AppsFlat.SemanticModel/definition/relationships.tmdl', 'AppsFlat.SemanticModel/definition/expressions.tmdl', 'AppsFlat.SemanticModel/definition/tables/Apps.tmdl', 'AppsFlat.SemanticModel/definition/tables/Calendar.tmdl') {
        Assert-True (Test-Path (Join-Path $flatOut $f)) "missing $f"
    }
    foreach ($j in (Get-ChildItem -Recurse -File -Path $flatOut | Where-Object { $_.Extension -in '.json', '.pbip', '.pbir', '.pbism' })) { [void]([IO.File]::ReadAllText($j.FullName) | ConvertFrom-Json) }
    $bytes = [IO.File]::ReadAllBytes((Join-Path $flatOut 'AppsFlat.Report/definition/pages/pages.json')); Assert-Equal 0x7B $bytes[0] 'no BOM'
    $tm = [IO.File]::ReadAllText((Join-Path $flatOut 'AppsFlat.SemanticModel/definition/tables/Apps.tmdl')); Assert-True ($tm.Contains("`r`n")) 'CRLF'
    $pbip = [IO.File]::ReadAllText((Join-Path $flatOut 'AppsFlat.pbip')) | ConvertFrom-Json
    Assert-Equal 'AppsFlat.Report' $pbip.artifacts[0].report.path
}
Test-Case 'visual files match the spec roles (page 1) and page 2 has the In filter' {
    $pagesDir = Join-Path $flatOut 'AppsFlat.Report/definition/pages'
    $pages = [IO.File]::ReadAllText((Join-Path $pagesDir 'pages.json')) | ConvertFrom-Json
    $visuals = @(Get-ChildItem -Recurse -Filter visual.json -Path (Join-Path $pagesDir $pages.pageOrder[0]) | ForEach-Object { [IO.File]::ReadAllText($_.FullName) | ConvertFrom-Json })
    Assert-Equal 8 $visuals.Count
    $types = @($visuals | ForEach-Object { $_.visual.visualType } | Sort-Object)
    Assert-Equal 'cardVisual,cardVisual,clusteredBarChart,lineChart,scatterChart,slicer,tableEx,textbox' ($types -join ',')
    $line = $visuals | Where-Object { $_.visual.visualType -eq 'lineChart' }
    Assert-Equal 'Calendar' $line.visual.query.queryState.Category.projections[0].field.Column.Expression.SourceRef.Entity
    $tbl = $visuals | Where-Object { $_.visual.visualType -eq 'tableEx' }
    Assert-Equal 3 $tbl.visual.query.queryState.Values.projections.Count
    $p2 = [IO.File]::ReadAllText((Join-Path $pagesDir ($pages.pageOrder[1] + '/page.json'))) | ConvertFrom-Json
    Assert-Equal "'Games'" $p2.filterConfig.filters[0].filter.Where[0].Condition.In.Values[0][0].Literal.Value
    $v2 = @(Get-ChildItem -Recurse -Filter visual.json -Path (Join-Path $pagesDir $pages.pageOrder[1]) | ForEach-Object { [IO.File]::ReadAllText($_.FullName) | ConvertFrom-Json })
    $pivot = $v2 | Where-Object { $_.visual.visualType -eq 'pivotTable' }
    Assert-Equal 2 $pivot.visual.query.queryState.Values.projections.Count
    Assert-Equal 'Year' $pivot.visual.query.queryState.Columns.projections[0].field.Column.Property
    $gauge = $v2 | Where-Object { $_.visual.visualType -eq 'gauge' }
    Assert-Equal 4 $gauge.visual.query.queryState.MaxValue.projections[0].field.Aggregation.Function
}
Test-Case 'build of spec-star: csv partitions, relationships, Series role' {
    $out = Join-Path $tmpRoot 'star'
    $log = (& $buildScript -Spec (Join-Path $fixtures 'spec-star.json') -RawData $fixtures -OutputDir $out | Out-String)
    Assert-Equal 0 $LASTEXITCODE $log
    Assert-Match 'tables=3 measures=2 pages=1 visuals=3' $log
    $sales = [IO.File]::ReadAllText((Join-Path $out 'SalesStar.SemanticModel/definition/tables/Sales.tmdl'))
    Assert-Match 'Csv\.Document\(File\.Contents\(DataFolder & "\\sales_cp949\.csv"\),\[Delimiter=",", Encoding=949, QuoteStyle=QuoteStyle\.Csv\]\)' $sales
    Assert-True (-not $sales.Contains('RenameColumns')) 'no renames in star spec'
    $rels = [IO.File]::ReadAllText((Join-Path $out 'SalesStar.SemanticModel/definition/relationships.tmdl'))
    Assert-Match "fromColumn: Sales\.'지역코드'\r?\n\ttoColumn: Regions\.'지역코드'" $rels
    Assert-Match "toColumn: Calendar\.Date" $rels
    $vis = @(Get-ChildItem -Recurse -Filter visual.json -Path (Join-Path $out 'SalesStar.Report') | ForEach-Object { [IO.File]::ReadAllText($_.FullName) | ConvertFrom-Json })
    $col = $vis | Where-Object { $_.visual.visualType -eq 'clusteredColumnChart' }
    Assert-Equal '제품' $col.visual.query.queryState.Series.projections[0].field.Column.Property
}
Test-Case 'broken spec exits 1 with three ERROR lines and writes nothing' {
    $out = Join-Path $tmpRoot 'broken'
    $log = (& $buildScript -Spec (Join-Path $fixtures 'spec-broken.json') -RawData $fixtures -OutputDir $out | Out-String)
    Assert-Equal 1 $LASTEXITCODE
    Assert-Equal 3 ([regex]::Matches($log, '(?m)^ERROR \d+: ').Count) $log
    Assert-Match '3 error\(s\)' $log
    Assert-True (-not (Test-Path $out)) 'no output on error'
}
Test-Case 'rebuild replaces definition folders but keeps .pbi caches' {
    $cache = Join-Path $flatOut 'AppsFlat.SemanticModel/.pbi/cache.abf'
    New-Item -ItemType Directory -Path (Split-Path -Parent $cache) -Force | Out-Null; [IO.File]::WriteAllText($cache, 'x')
    $stray = Join-Path $flatOut 'AppsFlat.SemanticModel/definition/tables/Stray.tmdl'; [IO.File]::WriteAllText($stray, 'table Stray')
    $log = (& $buildScript -Spec (Join-Path $fixtures 'spec-flat.json') -RawData $fixtures -OutputDir $flatOut | Out-String)
    Assert-Equal 0 $LASTEXITCODE $log
    Assert-True (Test-Path $cache) 'cache kept'
    Assert-True (-not (Test-Path $stray)) 'stray definition removed'
    $log2 = (& $buildScript -Spec (Join-Path $fixtures 'spec-flat.json') -RawData $fixtures -OutputDir $flatOut -Force | Out-String)
    Assert-True (-not (Test-Path $cache)) '-Force removes caches too'
}
