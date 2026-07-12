# Test-Syntax.ps1 - 自我驗證所有 .ps1 / .psm1 是否語法正確
$ErrorActionPreference = 'Stop'
if ($PSScriptRoot) {
    $here = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $here = Split-Path -LiteralPath $MyInvocation.MyCommand.Path -Parent
} else {
    $here = (Get-Location).Path
}
$files = @()
$files += Get-ChildItem -LiteralPath $here -Recurse -Include *.ps1, *.psm1 -File
$failed = 0
foreach ($f in $files) {
    $errors = $null
    $tokens = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        Write-Host ("[FAIL] {0}" -f $f.FullName) -ForegroundColor Red
        foreach ($e in $errors) {
            Write-Host ("       Line {0}: {1}" -f $e.Extent.StartLineNumber, $e.Message) -ForegroundColor Red
        }
        $failed++
    } else {
        Write-Host ("[OK]   {0}" -f $f.FullName) -ForegroundColor Green
    }
}
Write-Host ''
if ($failed -gt 0) {
    Write-Host ("語法檢查失敗: {0} 個檔案" -f $failed) -ForegroundColor Red
    exit 1
} else {
    Write-Host ("全部通過 ({0} 檔案)" -f $files.Count) -ForegroundColor Green
    exit 0
}
