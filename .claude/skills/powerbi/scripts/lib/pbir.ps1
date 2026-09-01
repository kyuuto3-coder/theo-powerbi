# pbir.ps1 — PBIR report emitter. Requires io.ps1 and spec.ps1 dot-sourced first.
$ErrorActionPreference = 'Stop'
$script:SchemaBase = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition'
$script:Schemas = @{
    visual  = "$script:SchemaBase/visualContainer/2.12.0/schema.json"
    page    = "$script:SchemaBase/page/2.1.0/schema.json"
    pages   = "$script:SchemaBase/pagesMetadata/1.1.0/schema.json"
    version = "$script:SchemaBase/versionMetadata/1.0.0/schema.json"
    pbir    = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json'
}
$script:AggLabels = @{ sum = 'Sum'; avg = 'Avg'; countDistinct = 'CountDistinct'; min = 'Min'; max = 'Max'; count = 'Count' }

function ConvertTo-EntityExpr {
    param([Parameter(Mandatory)]$Ref)
    $src = [ordered]@{ Expression = [ordered]@{ SourceRef = [ordered]@{ Entity = $Ref.Table } }; Property = $Ref.Name }
    switch ($Ref.Kind) {
        'Measure'     { return [ordered]@{ Measure = $src } }
        'Column'      { return [ordered]@{ Column = $src } }
        'Aggregation' { return [ordered]@{ Aggregation = [ordered]@{ Expression = [ordered]@{ Column = $src }; Function = $Ref.Function } } }
    }
    throw "unsupported field kind $($Ref.Kind)"
}

function ConvertTo-SourceExpr {
    # Same as ConvertTo-EntityExpr but with SourceRef.Source = alias (used inside filter Where clauses).
    param([Parameter(Mandatory)]$Ref, [Parameter(Mandatory)][string]$Alias)
    $src = [ordered]@{ Expression = [ordered]@{ SourceRef = [ordered]@{ Source = $Alias } }; Property = $Ref.Name }
    switch ($Ref.Kind) {
        'Measure'     { return [ordered]@{ Measure = $src } }
        'Column'      { return [ordered]@{ Column = $src } }
        'Aggregation' { return [ordered]@{ Aggregation = [ordered]@{ Expression = [ordered]@{ Column = $src }; Function = $Ref.Function } } }
    }
    throw "unsupported field kind $($Ref.Kind)"
}

function New-Projection {
    param([Parameter(Mandatory)]$Ref)
    if ($Ref.Kind -eq 'Aggregation') {
        $label = $script:AggLabels[$Ref.FunctionName]
        return [ordered]@{ field = (ConvertTo-EntityExpr -Ref $Ref); queryRef = "$label($($Ref.Table).$($Ref.Name))"; nativeQueryRef = "$label of $($Ref.Name)" }
    }
    return [ordered]@{ field = (ConvertTo-EntityExpr -Ref $Ref); queryRef = "$($Ref.Table).$($Ref.Name)"; nativeQueryRef = $Ref.Name }
}

function ConvertTo-FilterLiteral {
    param($V)
    if ($V -is [bool]) { return $V.ToString().ToLowerInvariant() }
    if ($V -is [int] -or $V -is [long] -or $V -is [int16] -or $V -is [byte]) { return ([long]$V).ToString() + 'L' }
    if ($V -is [double] -or $V -is [decimal] -or $V -is [single]) {
        $d = [double]$V
        if ($d -eq [Math]::Floor($d) -and [Math]::Abs($d) -lt 9e15) { return ([long]$d).ToString() + 'L' }
        return $d.ToString('R', [Globalization.CultureInfo]::InvariantCulture) + 'D'
    }
    return "'" + ([string]$V).Replace("'", "''") + "'"
}

function New-InFilter {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Field, [Parameter(Mandatory)]$Values)
    $r = Resolve-FieldRef -Model $Model -Ref $Field
    $vals = New-Object System.Collections.Generic.List[object]
    foreach ($v in (ConvertTo-Array $Values)) { $vals.Add([object[]]@([ordered]@{ Literal = [ordered]@{ Value = (ConvertTo-FilterLiteral $v) } })) }
    return [ordered]@{
        name = (New-HexId); field = (ConvertTo-EntityExpr -Ref $r); type = 'Categorical'
        filter = [ordered]@{
            Version = 2
            From = @([ordered]@{ Name = 't'; Entity = $r.Table; Type = 0 })
            Where = @([ordered]@{ Condition = [ordered]@{ In = [ordered]@{ Expressions = @((ConvertTo-SourceExpr -Ref $r -Alias 't')); Values = $vals.ToArray() } } })
        }
    }
}

function New-TopNFilter {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$CategoryField, [Parameter(Mandatory)][string]$ByField, [Parameter(Mandatory)][int]$N)
    $c = Resolve-FieldRef -Model $Model -Ref $CategoryField
    $b = Resolve-FieldRef -Model $Model -Ref $ByField
    $from = New-Object System.Collections.Generic.List[object]
    $from.Add([ordered]@{ Name = 't'; Entity = $c.Table; Type = 0 })
    $bAlias = 't'
    if ($b.Table -ne $c.Table) { $from.Add([ordered]@{ Name = 'm'; Entity = $b.Table; Type = 0 }); $bAlias = 'm' }
    return [ordered]@{
        name = (New-HexId); field = (ConvertTo-EntityExpr -Ref $c); type = 'TopN'
        filter = [ordered]@{
            Version = 2
            From = $from.ToArray()
            Where = @([ordered]@{ Condition = [ordered]@{ TopN = [ordered]@{
                Expression = (ConvertTo-SourceExpr -Ref $c -Alias 't'); Count = $N
                OrderBy = @([ordered]@{ Expression = (ConvertTo-SourceExpr -Ref $b -Alias $bAlias); Direction = 2 }) } } })
        }
    }
}

function New-TitleObjects {
    param([string]$Title)
    return [ordered]@{ title = @([ordered]@{ properties = [ordered]@{
        show = [ordered]@{ expr = [ordered]@{ Literal = [ordered]@{ Value = 'true' } } }
        text = [ordered]@{ expr = [ordered]@{ Literal = [ordered]@{ Value = "'" + $Title.Replace("'", "''") + "'" } } } } }) }
}

function New-VisualJson {
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)]$Visual, [Parameter(Mandatory)][int]$Index)
    $type = [string](Get-Prop $Visual 'type')
    $cat = $script:VisualCatalog[$type]
    $pos = Get-GridPosition -Grid (ConvertTo-Array (Get-Prop $Visual 'grid'))
    $position = [ordered]@{ x = $pos.x; y = $pos.y; z = $Index; height = $pos.height; width = $pos.width; tabOrder = $Index * 1000 }
    $vis = [ordered]@{ visualType = $cat.PbirType }
    $fields = Get-Prop $Visual 'fields' $null
    if ($cat.IsText) {
        $size = [int](Get-Prop $Visual 'fontSize' 20)
        $vis.objects = [ordered]@{ general = @([ordered]@{ properties = [ordered]@{ paragraphs = @([ordered]@{ textRuns = @([ordered]@{
            value = [string](Get-Prop $Visual 'text' ''); textStyle = [ordered]@{ fontSize = "$($size)pt"; fontWeight = 'bold' } }) }) } }) }
    } else {
        $qs = [ordered]@{}
        foreach ($role in ($cat.Required + $cat.Optional)) {
            $refs = ConvertTo-Array (Get-Prop $fields $role)
            if ($refs.Count -eq 0) { continue }
            $pbirRole = if ($cat.RoleMap.ContainsKey($role)) { $cat.RoleMap[$role] } else { $role }
            $projs = New-Object System.Collections.Generic.List[object]
            foreach ($ref in $refs) { $projs.Add((New-Projection -Ref (Resolve-FieldRef -Model $Model -Ref ([string]$ref)))) }
            $qs[$pbirRole] = [ordered]@{ projections = $projs.ToArray() }
        }
        $query = [ordered]@{ queryState = $qs }
        $sort = Get-Prop $Visual 'sort' $null
        if ($null -ne $sort) {
            $dir = if ([string](Get-Prop $sort 'direction' 'desc') -eq 'asc') { 'Ascending' } else { 'Descending' }
            $query.sortDefinition = [ordered]@{ sort = @([ordered]@{ field = (ConvertTo-EntityExpr -Ref (Resolve-FieldRef -Model $Model -Ref ([string](Get-Prop $sort 'field')))); direction = $dir }); isDefaultSort = $true }
        }
        $vis.query = $query
        $title = Get-Prop $Visual 'title' $null
        if ($title) { $vis.visualContainerObjects = New-TitleObjects -Title ([string]$title) }
    }
    $vis.drillFilterOtherVisuals = $true
    $vc = [ordered]@{ '$schema' = $script:Schemas.visual; name = (New-HexId); position = $position; visual = $vis }
    $filters = New-Object System.Collections.Generic.List[object]
    $topN = Get-Prop $Visual 'topN' $null
    if ($null -ne $topN) {
        $catRefs = ConvertTo-Array (Get-Prop $fields 'Category'); if ($catRefs.Count -eq 0) { $catRefs = ConvertTo-Array (Get-Prop $fields 'Rows') }
        $filters.Add((New-TopNFilter -Model $Model -CategoryField ([string]$catRefs[0]) -ByField ([string](Get-Prop $topN 'by')) -N ([int](Get-Prop $topN 'n'))))
    }
    foreach ($f in (ConvertTo-Array (Get-Prop $Visual 'filters'))) { $filters.Add((New-InFilter -Model $Model -Field ([string](Get-Prop $f 'field')) -Values (Get-Prop $f 'in'))) }
    if ($filters.Count -gt 0) { $vc.filterConfig = [ordered]@{ filters = $filters.ToArray() } }
    return $vc
}

function Write-Report {
    # Writes <Dir>/definition.pbir, <Dir>/definition/**, <Dir>/StaticResources/**. Returns @{ Pages; Visuals }
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Dir, [Parameter(Mandatory)][string]$TemplatesDir, [Parameter(Mandatory)][string]$ModelDirName)
    $def = Join-Path $Dir 'definition'
    ConvertTo-JsonFile -Path (Join-Path $Dir 'definition.pbir') -Object ([ordered]@{ '$schema' = $script:Schemas.pbir; version = '4.0'; datasetReference = [ordered]@{ byPath = [ordered]@{ path = '../' + $ModelDirName } } })
    ConvertTo-JsonFile -Path (Join-Path $def 'version.json') -Object ([ordered]@{ '$schema' = $script:Schemas.version; version = '2.0.0' })
    Write-Utf8File -Path (Join-Path $def 'report.json') -Content ([IO.File]::ReadAllText((Join-Path $TemplatesDir 'report.json'), [Text.Encoding]::UTF8))
    $themeDir = Join-Path $Dir 'StaticResources/SharedResources/BaseThemes'
    Write-Utf8File -Path (Join-Path $themeDir 'Fluent2-CY26SU08.json') -Content ([IO.File]::ReadAllText((Join-Path $TemplatesDir 'theme/Fluent2-CY26SU08.json'), [Text.Encoding]::UTF8))
    $pageIds = New-Object System.Collections.Generic.List[string]
    $visualCount = 0
    foreach ($page in $Model.Pages) {
        $pageId = New-HexId; $pageIds.Add($pageId)
        $pageDir = Join-Path (Join-Path $def 'pages') $pageId
        $pj = [ordered]@{ '$schema' = $script:Schemas.page; name = $pageId; displayName = [string](Get-Prop $page 'name'); displayOption = 'FitToPage'; height = 1080; width = 1920 }
        $pf = New-Object System.Collections.Generic.List[object]
        foreach ($f in (ConvertTo-Array (Get-Prop $page 'filters'))) { $pf.Add((New-InFilter -Model $Model -Field ([string](Get-Prop $f 'field')) -Values (Get-Prop $f 'in'))) }
        if ($pf.Count -gt 0) { $pj.filterConfig = [ordered]@{ filters = $pf.ToArray() } }
        ConvertTo-JsonFile -Path (Join-Path $pageDir 'page.json') -Object $pj
        $i = 0
        foreach ($v in (ConvertTo-Array (Get-Prop $page 'visuals'))) {
            $vj = New-VisualJson -Model $Model -Visual $v -Index $i
            ConvertTo-JsonFile -Path (Join-Path (Join-Path (Join-Path $pageDir 'visuals') $vj.name) 'visual.json') -Object $vj
            $i++; $visualCount++
        }
    }
    ConvertTo-JsonFile -Path (Join-Path $def 'pages/pages.json') -Object ([ordered]@{ '$schema' = $script:Schemas.pages; pageOrder = $pageIds.ToArray(); activePageName = $pageIds[0] })
    return @{ Pages = $pageIds.Count; Visuals = $visualCount }
}
