. (Join-Path $libDir 'io.ps1')

Test-Case 'Write-Utf8File writes UTF-8 without BOM and CRLF' {
    $p = Join-Path $tmpRoot 'a\b\c.txt'
    Write-Utf8File -Path $p -Content "line1`nline2 한글"
    $bytes = [IO.File]::ReadAllBytes($p)
    Assert-True ($bytes[0] -ne 0xEF) 'must not start with BOM'
    $text = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
    Assert-Equal "line1`r`nline2 한글" $text
}
Test-Case 'New-HexId returns 20 lowercase hex chars, unique' {
    $a = New-HexId; $b = New-HexId
    Assert-Match '^[0-9a-f]{20}$' $a
    Assert-True ($a -ne $b) 'ids must differ'
}
Test-Case 'ConvertTo-JsonFile round-trips ordered objects' {
    $p = Join-Path $tmpRoot 'obj.json'
    ConvertTo-JsonFile -Object ([ordered]@{ b = 1; a = @('x') }) -Path $p
    $raw = [IO.File]::ReadAllText($p)
    Assert-True ($raw.IndexOf('"b"') -lt $raw.IndexOf('"a"')) 'key order preserved'
    $back = $raw | ConvertFrom-Json
    Assert-Equal 'x' $back.a[0]
}
Test-Case 'Get-WorkspaceRoot resolves four levels above scripts/' {
    $root = Get-WorkspaceRoot -ScriptRoot (Split-Path -Parent $libDir)
    Assert-True (Test-Path (Join-Path $root 'CLAUDE.md')) "root should contain CLAUDE.md, got $root"
}
