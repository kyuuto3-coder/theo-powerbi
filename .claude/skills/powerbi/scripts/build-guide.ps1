<#
.SYNOPSIS  Generates output/<Name>/manual-guide.md (Korean, English UI terms) from report-spec.json.
.EXAMPLE   powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/powerbi/scripts/build-guide.ps1 -Spec output/Sales/report-spec.json
#>
param([Parameter(Mandatory)][string]$Spec, [string]$RawData = '')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
. (Join-Path $PSScriptRoot 'lib/io.ps1')
. (Join-Path $PSScriptRoot 'lib/spec.ps1')
. (Join-Path $PSScriptRoot 'lib/tmdl.ps1')
. (Join-Path $PSScriptRoot 'lib/guide.ps1')

$root = Get-WorkspaceRoot -ScriptRoot $PSScriptRoot
if (-not $RawData) { $RawData = Join-Path $root 'rawdata' }
if (-not (Test-Path -LiteralPath $Spec)) { Write-Output "ERROR 1: spec file not found: $Spec"; exit 1 }
$specPath = (Get-Item -LiteralPath $Spec).FullName
$specObj = Read-Spec -Path $specPath
$v = Get-SpecValidation -Spec $specObj -RawDataDir $RawData
if ($v.Errors.Count -gt 0) { $i = 0; foreach ($e in $v.Errors) { $i++; Write-Output ("ERROR {0}: {1}" -f $i, $e) }; exit 1 }
$absRaw = (Get-Item -LiteralPath $RawData).FullName
$dir = Split-Path -Parent $specPath
$outMd = Join-Path $dir 'manual-guide.md'; $outHtml = Join-Path $dir 'manual-guide.html'
Write-Utf8File -Path $outMd -Content (New-ManualGuide -Model $v.Model -RawDataDir $absRaw)
Write-Utf8File -Path $outHtml -Content (New-ManualGuideHtml -Model $v.Model -RawDataDir $absRaw)
Write-Output ("GUIDE {0}" -f $outHtml)
Write-Output ("GUIDE-MD {0}" -f $outMd)
exit 0
