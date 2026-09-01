# io.ps1 — file writing, identifiers, paths. Dot-source from scripts.
$ErrorActionPreference = 'Stop'

function Write-Utf8File {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][AllowEmptyString()][string]$Content)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $text = $Content -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($false)))
}

function New-HexId { -join ((1..20) | ForEach-Object { '{0:x}' -f (Get-Random -Maximum 16) }) }

function New-LineageTag { [guid]::NewGuid().ToString() }

function ConvertTo-JsonFile {
    param([Parameter(Mandatory)]$Object, [Parameter(Mandatory)][string]$Path)
    Write-Utf8File -Path $Path -Content ($Object | ConvertTo-Json -Depth 64)
}

function Get-WorkspaceRoot {
    # <root>/.claude/skills/powerbi/scripts → <root>
    param([Parameter(Mandatory)][string]$ScriptRoot)
    [System.IO.Path]::GetFullPath((Join-Path $ScriptRoot '../../../..'))
}

function Get-Prop {
    # Safe property read for PSCustomObject/hashtable; returns $Default when missing or null.
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } else { return $Default } }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

function ConvertTo-Array {
    # Always returns object[] (empty for $null), never unrolls.
    param($Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Array]) { return $Value }
    return ,@($Value)
}
