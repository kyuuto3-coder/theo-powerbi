. (Join-Path $libDir 'io.ps1')
. (Join-Path $libDir 'spec.ps1')
. (Join-Path $libDir 'tmdl.ps1')

Test-Case 'Format-TmdlName quotes anything that is not a plain identifier' {
    Assert-Equal 'Apps' (Format-TmdlName 'Apps')
    Assert-Equal "'앱 이름'" (Format-TmdlName '앱 이름')
    Assert-Equal "'It''s'" (Format-TmdlName "It's")
}
Test-Case 'New-MQuery for xlsx: sheet navigation, select, rename, types' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $q = (New-MQuery -Table $m.Tables['Apps']) -join "`n"
    Assert-Match 'Source = Excel\.Workbook\(File\.Contents\(DataFolder & "\\apps\.xlsx"\), null, true\),' $q
    Assert-Match 'Sheet = Source\{\[Item="Data",Kind="Sheet"\]\}\[Data\],' $q
    Assert-Match 'Promoted = Table\.PromoteHeaders\(Sheet, \[PromoteAllScalars=true\]\),' $q
    Assert-Match 'Selected = Table\.SelectColumns\(Promoted, \{"id", "track_name", "size_bytes"' $q
    Assert-Match 'Renamed = Table\.RenameColumns\(Selected, \{\{"track_name", "앱 이름"\}, \{"prime_genre", "장르"\}, \{"release_date", "출시일"\}\}\),' $q
    Assert-Match 'Typed = Table\.TransformColumnTypes\(Renamed, \{\{"id", Int64\.Type\}, \{"앱 이름", type text\}, \{"size_bytes", Int64\.Type\}, \{"price", Currency\.Type\}, \{"rating_count_tot", Int64\.Type\}, \{"user_rating", type number\}, \{"장르", type text\}, \{"출시일", type date\}\}\)' $q
    Assert-True (-not $q.Contains('Skipped')) 'headerRow 1 → no skip'
}
Test-Case 'New-MQuery for csv: encoding, delimiter, header skip, no rename step' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-star.json'))
    $t = $m.Tables['Sales']; $t.HeaderRow = 3
    $q = (New-MQuery -Table $t) -join "`n"
    Assert-Match 'Csv\.Document\(File\.Contents\(DataFolder & "\\sales_cp949\.csv"\),\[Delimiter=",", Encoding=949, QuoteStyle=QuoteStyle\.Csv\]\),' $q
    Assert-Match 'Skipped = Table\.Skip\(Source, 2\),' $q
    Assert-Match 'Promoted = Table\.PromoteHeaders\(Skipped' $q
    Assert-Match 'Typed = Table\.TransformColumnTypes\(Selected, ' $q
    Assert-True (-not $q.Contains('RenameColumns')) 'no renames'
}
Test-Case 'New-TableTmdl emits columns, measures, partition' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $txt = New-TableTmdl -Table $m.Tables['Apps'] -Model $m
    Assert-Match "(?m)^table Apps\r?\n\tlineageTag: [0-9a-f-]{36}" $txt
    Assert-Match "(?m)^\t/// App ID\r?\n\tcolumn id\r?\n\t\tdataType: int64\r?\n\t\tisHidden\r?\n\t\tisKey\r?\n\t\tlineageTag: .*\r?\n\t\tsummarizeBy: none\r?\n\t\tsourceColumn: id" $txt
    Assert-Match "(?m)^\tcolumn '출시일'\r?\n\t\tdataType: dateTime\r?\n\t\tformatString: yyyy-MM-dd\r?\n(.*\r?\n)+?\t\tannotation UnderlyingDateTimeDataType = Date" $txt
    Assert-Match "(?m)^\tcolumn user_rating\r?\n\t\tdataType: double\r?\n\t\tformatString: 0\.0\r?\n\t\tlineageTag: .*\r?\n\t\tsummarizeBy: average" $txt
    Assert-Match "(?m)^\tmeasure '앱 수' = COUNTROWS\(Apps\)\r?\n\t\tformatString: #,0\r?\n\t\tlineageTag: " $txt
    Assert-Match "formatString: \\\`$#,0\.00" $txt
    Assert-Match "(?m)^\tpartition Apps = m\r?\n\t\tmode: import\r?\n\t\tsource =\r?\n\t\t\t\tlet\r?\n\t\t\t\t    Source = Excel" $txt
}
Test-Case 'multi-line DAX measures are indented under the measure line' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $m.Tables['Apps'].Measures = @(@{ Name = 'X'; Dax = "VAR a = 1`nRETURN a"; Format = $null; Description = 'two lines' })
    $txt = New-TableTmdl -Table $m.Tables['Apps'] -Model $m
    Assert-Match "(?m)^\t/// two lines\r?\n\tmeasure X =\r?\n\t\t\tVAR a = 1\r?\n\t\t\tRETURN a\r?\n\t\tlineageTag: " $txt
}
Test-Case 'Calendar table is a calculated table marked as date table' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-star.json'))
    $txt = New-TableTmdl -Table $m.Tables['Calendar'] -Model $m
    Assert-Match "(?m)^table Calendar\r?\n\tlineageTag: .*\r?\n\tdataCategory: Time" $txt
    Assert-Match "(?m)^\tcolumn Date\r?\n\t\tdataType: dateTime\r?\n\t\tisKey\r?\n\t\tformatString: yyyy-MM-dd\r?\n\t\tlineageTag: .*\r?\n\t\tsummarizeBy: none\r?\n\t\tisNameInferred\r?\n\t\tisDataTypeInferred\r?\n\t\tsourceColumn: \[Date\]" $txt
    Assert-Match "(?m)^\tcolumn Month\r?\n(.*\r?\n)+?\t\tsortByColumn: MonthNo" $txt
    Assert-Match "(?m)^\tpartition Calendar = calculated\r?\n\t\tmode: import\r?\n\t\tsource = ADDCOLUMNS \( CALENDAR \( DATE \( YEAR \( MIN \( 'Sales'\[주문일\] \) \), 1, 1 \), DATE \( YEAR \( MAX \( 'Sales'\[주문일\] \) \), 12, 31 \) \), ""Year"", YEAR \( \[Date\] \), ""Quarter"", ""Q"" & FORMAT \( \[Date\], ""q"" \), ""MonthNo"", MONTH \( \[Date\] \), ""Month"", FORMAT \( \[Date\], ""MMM"" \), ""YearMonth"", FORMAT \( \[Date\], ""yyyy-MM"" \) \)" $txt
}
Test-Case 'New-ModelTmdl and New-RelationshipsTmdl' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-star.json'))
    $mt = New-ModelTmdl -Model $m
    Assert-Match "(?m)^model Model\r?\n\tculture: ko-KR\r?\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\r?\n\tsourceQueryCulture: ko-KR\r?\n\tdataAccessOptions\r?\n\t\tlegacyRedirects\r?\n\t\treturnErrorValuesAsNull" $mt
    Assert-Match 'annotation __PBI_TimeIntelligenceEnabled = 0' $mt
    Assert-Match 'annotation PBI_QueryOrder = \["Sales","Regions"\]' $mt
    Assert-Match "(?m)^ref table Sales\r?\n^ref table Regions\r?\n^ref table Calendar\r?\n\r?\n^ref expression DataFolder" $mt
    $rt = New-RelationshipsTmdl -Model $m
    Assert-Match "(?m)^relationship Sales_지역코드_Regions\r?\n\tfromColumn: Sales\.'지역코드'\r?\n\ttoColumn: Regions\.'지역코드'\r?\n" $rt
    Assert-Match "(?m)^relationship Sales_주문일_Calendar\r?\n\tfromColumn: Sales\.'주문일'\r?\n\ttoColumn: Calendar\.Date" $rt
    $m.Relationships[0].Active = $false; $m.Relationships[0].CrossFilter = 'both'
    $rt2 = New-RelationshipsTmdl -Model $m
    Assert-Match "(?m)^relationship Sales_지역코드_Regions\r?\n\tisActive: false\r?\n\tcrossFilteringBehavior: bothDirections\r?\n\tfromColumn" $rt2
}
Test-Case 'Write-SemanticModel writes the full folder' {
    $m = Get-SpecModel -Spec (Read-Spec -Path (Join-Path $fixtures 'spec-flat.json'))
    $dir = Join-Path $tmpRoot 'sm'
    $r = Write-SemanticModel -Model $m -Dir $dir -RawDataDir $fixtures
    foreach ($f in 'definition.pbism', 'definition/database.tmdl', 'definition/model.tmdl', 'definition/relationships.tmdl', 'definition/expressions.tmdl', 'definition/tables/Apps.tmdl', 'definition/tables/Calendar.tmdl') { Assert-True (Test-Path (Join-Path $dir $f)) "missing $f" }
    $apps = [IO.File]::ReadAllText((Join-Path $dir 'definition/tables/Apps.tmdl'))
    Assert-Match 'File\.Contents\(DataFolder & "\\apps\.xlsx"\)' $apps
    $expr = [IO.File]::ReadAllText((Join-Path $dir 'definition/expressions.tmdl'))
    Assert-Match '^expression DataFolder = "(\\\\|[A-Za-z]:\\)[^"]*fixtures" meta \[IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true\]' $expr
    Assert-Equal $r.DataFolder ($expr -replace '(?s)^expression DataFolder = "([^"]*)".*$', '$1')
    Assert-Equal 0 $r.Warnings.Count ($r.Warnings -join ' | ')
}
