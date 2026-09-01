. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'spec.ps1')

Test-Case 'Get-SpecModel builds tables with final column names, measures, and a Calendar table' {
    $spec = Read-Spec -Path (Join-Path $fixtures 'spec-flat.json')
    $m = Get-SpecModel -Spec $spec
    Assert-Equal 'AppsFlat' $m.Name
    Assert-True ($m.Tables.Contains('Apps')) 'Apps table'
    Assert-Equal '앱 이름' $m.Tables['Apps'].Columns[1].Name
    Assert-Equal 'track_name' $m.Tables['Apps'].Columns[1].Source
    Assert-Equal 'none' $m.Tables['Apps'].Columns[0].SummarizeBy
    Assert-Equal 'sum' $m.Tables['Apps'].Columns[2].SummarizeBy
    Assert-Equal 3 $m.Tables['Apps'].Measures.Count
    Assert-True ($m.Tables.Contains('Calendar')) 'Calendar generated'
    Assert-Equal 1 $m.Relationships.Count
    Assert-Equal 'Apps' $m.Relationships[0].FromTable; Assert-Equal '출시일' $m.Relationships[0].FromColumn
    Assert-Equal 'Calendar' $m.Relationships[0].ToTable; Assert-Equal 'Date' $m.Relationships[0].ToColumn
}
Test-Case 'Resolve-FieldRef distinguishes measure, column, aggregation' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $a = Resolve-FieldRef -Model $m -Ref 'Apps.앱 수';         Assert-Equal 'Measure' $a.Kind
    $b = Resolve-FieldRef -Model $m -Ref 'Apps.장르';          Assert-Equal 'Column' $b.Kind; Assert-Equal 'Apps' $b.Table
    $c = Resolve-FieldRef -Model $m -Ref 'sum:Apps.rating_count_tot'; Assert-Equal 'Aggregation' $c.Kind; Assert-Equal 0 $c.Function
    $d = Resolve-FieldRef -Model $m -Ref 'countDistinct:Apps.장르'; Assert-Equal 2 $d.Function
    $e = Resolve-FieldRef -Model $m -Ref 'Calendar.YearMonth'; Assert-Equal 'Column' $e.Kind
    Assert-Throws { Resolve-FieldRef -Model $m -Ref 'Apps.nope' } 'unknown field'
    Assert-Throws { Resolve-FieldRef -Model $m -Ref 'sum:Apps.앱 수' } 'aggregation prefix on a measure'
    Assert-Throws { Resolve-FieldRef -Model $m -Ref 'noDot' } 'Table.Name'
}
Test-Case 'Get-SpecValidation passes the good fixtures' {
    foreach ($f in 'spec-flat.json', 'spec-star.json') {
        $v = Get-SpecValidation -Spec (Read-Spec -Path (Join-Path $fixtures $f)) -RawDataDir $fixtures
        Assert-Equal 0 $v.Errors.Count ("errors in " + $f + ": " + ($v.Errors -join ' | '))
    }
}
Test-Case 'Get-SpecValidation reports the three deliberate errors with spec paths' {
    $v = Get-SpecValidation -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-broken.json')) -RawDataDir $fixtures
    Assert-Equal 3 $v.Errors.Count ($v.Errors -join ' | ')
    Assert-Match "tables\[0\]\.columns\[1\]\.type.*money" ($v.Errors -join "`n")
    Assert-Match "pages\[0\]\.visuals\[0\]\.fields\.Values\[0\].*unknown field" ($v.Errors -join "`n")
    Assert-Match "pages\[0\]\.visuals\[1\]\.fields\.Bogus.*not a role of lineChart" ($v.Errors -join "`n")
}
Test-Case 'Get-SpecValidation catches missing file, bad grid, duplicate names' {
    $spec = Read-Spec -Path (Join-Path $fixtures 'spec-flat.json')
    $spec.tables[0].file = 'missing.xlsx'
    $spec.pages[0].visuals[1].grid = @(10, 0, 4, 2)
    $spec.measures[1].name = '앱 수'
    $v = Get-SpecValidation -Spec $spec -RawDataDir $fixtures
    $all = $v.Errors -join "`n"
    Assert-Match 'tables\[0\]\.file.*missing\.xlsx' $all
    Assert-Match 'pages\[0\]\.visuals\[1\]\.grid.*exceeds' $all
    Assert-Match 'measures\[1\]\.name.*duplicate' $all
}
Test-Case 'Get-GridPosition converts the 12x8 grid into pixels' {
    $p = Get-GridPosition -Grid @(0, 0, 3, 2)
    Assert-Equal 40 $p.x; Assert-Equal 40 $p.y; Assert-Equal 445 $p.width; Assert-Equal 235 $p.height
    $q = Get-GridPosition -Grid @(6, 1, 6, 4)
    Assert-Equal 970 $q.x; Assert-Equal 168 $q.y; Assert-Equal 910 $q.width; Assert-Equal 490 $q.height
}
