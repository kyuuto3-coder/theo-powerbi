<#
.SYNOPSIS  Validates report-spec.json and builds a Power BI Project (.pbip) next to it.
.EXAMPLE   powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-pbip.ps1 -Spec output/Sales/report-spec.json
.PARAMETER RawData    Folder holding the data files (default: <workspace>/rawdata)
.PARAMETER OutputDir  Where to write <Name>.pbip (default: the spec's folder)
.PARAMETER Force      Also delete .pbi caches of a previous build
#>
param([Parameter(Mandatory)][string]$Spec, [string]$RawData = '', [string]$OutputDir = '', [switch]$Force)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
. (Join-Path $PSScriptRoot 'lib/io.ps1')
. (Join-Path $PSScriptRoot 'lib/spec.ps1')
. (Join-Path $PSScriptRoot 'lib/tmdl.ps1')
. (Join-Path $PSScriptRoot 'lib/pbir.ps1')

$root = Get-WorkspaceRoot -ScriptRoot $PSScriptRoot
if (-not $RawData) { $RawData = Join-Path $root 'rawdata' }
if (-not (Test-Path -LiteralPath $Spec)) { Write-Output "ERROR 1: spec file not found: $Spec"; exit 1 }
$specPath = (Get-Item -LiteralPath $Spec).FullName
try { $specObj = Read-Spec -Path $specPath } catch { Write-Output ("ERROR 1: spec is not valid JSON: " + $_.Exception.Message); exit 1 }

$v = Get-SpecValidation -Spec $specObj -RawDataDir $RawData
if ($v.Errors.Count -gt 0) {
    $i = 0
    foreach ($e in $v.Errors) { $i++; Write-Output ("ERROR {0}: {1}" -f $i, $e) }
    Write-Output ("{0} error(s) in {1} - fix the spec and rebuild" -f $v.Errors.Count, $specPath)
    exit 1
}
foreach ($w in $v.Warnings) { Write-Output ("WARN: " + $w) }
$model = $v.Model
$name = $model.Name
if (-not $OutputDir) { $OutputDir = Split-Path -Parent $specPath }
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
$reportDir = Join-Path $OutputDir ($name + '.Report')
$modelDir  = Join-Path $OutputDir ($name + '.SemanticModel')

if ($Force) { foreach ($d in $reportDir, $modelDir) { if (Test-Path $d) { Remove-Item -Recurse -Force $d } } }
else {
    foreach ($d in (Join-Path $reportDir 'definition'), (Join-Path $reportDir 'StaticResources'), (Join-Path $modelDir 'definition')) { if (Test-Path $d) { Remove-Item -Recurse -Force $d } }
}

$mw = Write-SemanticModel -Model $model -Dir $modelDir -RawDataDir $RawData
$rw = Write-Report -Model $model -Dir $reportDir -TemplatesDir (Join-Path (Split-Path -Parent $PSScriptRoot) 'templates') -ModelDirName ($name + '.SemanticModel')
$pbipPath = Join-Path $OutputDir ($name + '.pbip')
ConvertTo-JsonFile -Path $pbipPath -Object ([ordered]@{
    '$schema' = 'https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json'; version = '1.0'
    artifacts = @([ordered]@{ report = [ordered]@{ path = $name + '.Report' } }); settings = [ordered]@{ enableAutoRecovery = $true } })
Write-Utf8File -Path (Join-Path $OutputDir '.gitignore') -Content "**/.pbi/localSettings.json`n**/.pbi/cache.abf`n"
foreach ($w in $mw.Warnings) { Write-Output ("WARN: " + $w) }
$measureCount = 0; foreach ($k in $model.Tables.Keys) { $measureCount += @($model.Tables[$k].Measures).Count }
Write-Output ("BUILT {0}  (tables={1} measures={2} pages={3} visuals={4})" -f $pbipPath, $model.Tables.Count, $measureCount, $rw.Pages, $rw.Visuals)
exit 0
