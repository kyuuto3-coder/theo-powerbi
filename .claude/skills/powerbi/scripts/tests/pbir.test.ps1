. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'spec.ps1')
. (Join-Path $libDir 'pbir.ps1')
$flat = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))

Test-Case 'ConvertTo-EntityExpr renders Column / Measure / Aggregation' {
    $c = ConvertTo-EntityExpr -Ref (Resolve-FieldRef -Model $flat -Ref 'Apps.장르')
    Assert-Equal 'Apps' $c.Column.Expression.SourceRef.Entity; Assert-Equal '장르' $c.Column.Property
    $m = ConvertTo-EntityExpr -Ref (Resolve-FieldRef -Model $flat -Ref 'Apps.앱 수')
    Assert-Equal '앱 수' $m.Measure.Property
    $a = ConvertTo-EntityExpr -Ref (Resolve-FieldRef -Model $flat -Ref 'avg:Apps.price')
    Assert-Equal 1 $a.Aggregation.Function; Assert-Equal 'price' $a.Aggregation.Expression.Column.Property
}
Test-Case 'New-Projection sets queryRef and nativeQueryRef' {
    $p = New-Projection -Ref (Resolve-FieldRef -Model $flat -Ref 'sum:Apps.rating_count_tot')
    Assert-Equal 'Sum(Apps.rating_count_tot)' $p.queryRef; Assert-Equal 'Sum of rating_count_tot' $p.nativeQueryRef
    $q = New-Projection -Ref (Resolve-FieldRef -Model $flat -Ref 'Apps.앱 수')
    Assert-Equal 'Apps.앱 수' $q.queryRef; Assert-Equal '앱 수' $q.nativeQueryRef
}
Test-Case 'ConvertTo-FilterLiteral formats strings, ints, decimals, bools' {
    Assert-Equal "'Games'" (ConvertTo-FilterLiteral 'Games')
    Assert-Equal "'It''s'" (ConvertTo-FilterLiteral "It's")
    Assert-Equal '10L' (ConvertTo-FilterLiteral 10)
    Assert-Equal '2.5D' (ConvertTo-FilterLiteral 2.5)
    Assert-Equal '2.5D' (ConvertTo-FilterLiteral ([decimal]2.5))
    Assert-Equal 'true' (ConvertTo-FilterLiteral $true)
}
Test-Case 'New-VisualJson: card with title' {
    $v = $flat.Pages[0].visuals[1]
    $j = New-VisualJson -Model $flat -Visual $v -Index 1
    Assert-Match '^[0-9a-f]{20}$' $j.name
    Assert-Equal 'cardVisual' $j.visual.visualType
    Assert-Equal '앱 수' $j.visual.query.queryState.Data.projections[0].field.Measure.Property
    Assert-Equal "'앱 수'" $j.visual.visualContainerObjects.title[0].properties.text.expr.Literal.Value
    Assert-Equal 40 $j.position.x; Assert-Equal 168 $j.position.y; Assert-Equal 445 $j.position.width; Assert-Equal 235 $j.position.height; Assert-Equal 1000 $j.position.tabOrder
    Assert-True $j.visual.drillFilterOtherVisuals 'drill flag'
}
Test-Case 'New-VisualJson: bar chart with sort and TopN filter' {
    $j = New-VisualJson -Model $flat -Visual $flat.Pages[0].visuals[3] -Index 3
    Assert-Equal 'clusteredBarChart' $j.visual.visualType
    Assert-Equal '장르' $j.visual.query.queryState.Category.projections[0].field.Column.Property
    Assert-Equal 'Descending' $j.visual.query.sortDefinition.sort[0].direction
    Assert-Equal '앱 수' $j.visual.query.sortDefinition.sort[0].field.Measure.Property
    Assert-True $j.visual.query.sortDefinition.isDefaultSort 'default sort'
    $f = $j.filterConfig.filters[0]
    Assert-Equal 'TopN' $f.type; Assert-Equal '장르' $f.field.Column.Property
    Assert-Equal 2 $f.filter.Version; Assert-Equal 'Apps' $f.filter.From[0].Entity; Assert-Equal 't' $f.filter.From[0].Name
    $top = $f.filter.Where[0].Condition.TopN
    Assert-Equal 3 $top.Count; Assert-Equal 't' $top.Expression.Column.Expression.SourceRef.Source
    Assert-Equal 2 $top.OrderBy[0].Direction; Assert-Equal '앱 수' $top.OrderBy[0].Expression.Measure.Property
}
Test-Case 'New-VisualJson: scatter aggregations and textbox' {
    $s = New-VisualJson -Model $flat -Visual $flat.Pages[0].visuals[5] -Index 5
    Assert-Equal 1 $s.visual.query.queryState.X.projections[0].field.Aggregation.Function
    Assert-Equal 0 $s.visual.query.queryState.Size.projections[0].field.Aggregation.Function
    $t = New-VisualJson -Model $flat -Visual $flat.Pages[0].visuals[0] -Index 0
    Assert-Equal 'textbox' $t.visual.visualType
    $run = $t.visual.objects.general[0].properties.paragraphs[0].textRuns[0]
    Assert-Equal '앱스토어 대시보드' $run.value; Assert-Equal '24pt' $run.textStyle.fontSize
    Assert-True ($null -eq $t.visual.query) 'textbox has no query'
}
Test-Case 'New-InFilter builds a Categorical In filter with nested value arrays' {
    $f = New-InFilter -Model $flat -Field 'Apps.장르' -Values @('Games', 'Education')
    Assert-Equal 'Categorical' $f.type
    $in = $f.filter.Where[0].Condition.In
    Assert-Equal '장르' $in.Expressions[0].Column.Property
    Assert-Equal 2 $in.Values.Count
    Assert-Equal "'Games'" $in.Values[0][0].Literal.Value
    $json = $f | ConvertTo-Json -Depth 32
    Assert-Match '"Values":\s*\[\s*\[\s*\{' $json
}
Test-Case 'Write-Report writes pages, visuals, static resources and metadata' {
    $dir = Join-Path $tmpRoot 'rep'
    $r = Write-Report -Model $flat -Dir $dir -TemplatesDir (Join-Path (Split-Path -Parent (Split-Path -Parent $libDir)) 'templates') -ModelDirName 'AppsFlat.SemanticModel'
    Assert-Equal 2 $r.Pages; Assert-Equal 11 $r.Visuals
    $pbir = [IO.File]::ReadAllText((Join-Path $dir 'definition.pbir')) | ConvertFrom-Json
    Assert-Equal '../AppsFlat.SemanticModel' $pbir.datasetReference.byPath.path
    $pages = [IO.File]::ReadAllText((Join-Path $dir 'definition/pages/pages.json')) | ConvertFrom-Json
    Assert-Equal 2 $pages.pageOrder.Count; Assert-Equal $pages.pageOrder[0] $pages.activePageName
    $p2 = [IO.File]::ReadAllText((Join-Path $dir ('definition/pages/' + $pages.pageOrder[1] + '/page.json'))) | ConvertFrom-Json
    Assert-Equal '상세' $p2.displayName; Assert-Equal 1920 $p2.width; Assert-Equal 'FitToPage' $p2.displayOption
    Assert-Equal 'Categorical' $p2.filterConfig.filters[0].type
    Assert-Equal 8 (Get-ChildItem (Join-Path $dir ('definition/pages/' + $pages.pageOrder[0] + '/visuals')) -Directory).Count
    Assert-True (Test-Path (Join-Path $dir 'StaticResources/SharedResources/BaseThemes/Fluent2-CY26SU08.json')) 'theme'
    Assert-True (Test-Path (Join-Path $dir 'definition/report.json')) 'report.json'
    Assert-Equal '2.0.0' (([IO.File]::ReadAllText((Join-Path $dir 'definition/version.json')) | ConvertFrom-Json).version)
}
