# Minimal test harness (no Pester dependency). Runs every *.test.ps1 in this folder.
param([string]$Filter = '*')
$ErrorActionPreference = 'Stop'
$script:TestsRun = 0; $script:TestsFailed = 0; $script:Failures = New-Object System.Collections.Generic.List[string]

function Test-Case {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][scriptblock]$Body)
    $script:TestsRun++
    try { & $Body; Write-Host ("  PASS  " + $Name) }
    catch { $script:TestsFailed++; $script:Failures.Add("$Name :: $($_.Exception.Message)"); Write-Host ("  FAIL  " + $Name + "`n        " + $_.Exception.Message) -ForegroundColor Red }
}
function Assert-True  { param($Condition, [string]$Message = 'expected true')  if (-not $Condition) { throw $Message } }
function Assert-Equal { param($Expected, $Actual, [string]$Message = '')
    if ("$Expected" -ne "$Actual") { throw ("expected [" + $Expected + "] got [" + $Actual + "] " + $Message) } }
function Assert-Match { param([string]$Pattern, [string]$Text, [string]$Message = '')
    if ($Text -notmatch $Pattern) { throw ("expected match /" + $Pattern + "/ in: " + $Text.Substring(0, [Math]::Min(300, $Text.Length)) + " " + $Message) } }
function Assert-Throws { param([scriptblock]$Body, [string]$Pattern = '.')
    $threw = $false; try { & $Body } catch { $threw = $true; if ($_.Exception.Message -notmatch $Pattern) { throw ("threw but message [" + $_.Exception.Message + "] !~ /" + $Pattern + "/") } }
    if (-not $threw) { throw 'expected an exception' } }

$fixtures = Join-Path $PSScriptRoot 'fixtures'
$libDir   = Join-Path (Split-Path -Parent $PSScriptRoot) 'lib'
$tmpRoot  = Join-Path ([IO.Path]::GetTempPath()) ('pbi-skill-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null

foreach ($file in Get-ChildItem -Path $PSScriptRoot -Filter "$Filter.test.ps1" | Sort-Object Name) {
    Write-Host ("== " + $file.Name)
    . $file.FullName
}
Write-Host ""
Write-Host ("{0} tests, {1} failed" -f $script:TestsRun, $script:TestsFailed)
foreach ($f in $script:Failures) { Write-Host ("  - " + $f) -ForegroundColor Red }
Remove-Item -Recurse -Force $tmpRoot -ErrorAction SilentlyContinue
if ($script:TestsFailed -gt 0) { exit 1 } else { exit 0 }
