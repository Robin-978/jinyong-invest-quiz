# Start-CoWorkTracker.ps1
# 啟動協同工作紀錄系統 (PowerShell 5.1 / .NET / WinForms)
[CmdletBinding()]
param(
    [string]$DataRoot
)

$ErrorActionPreference = 'Stop'
if ($PSScriptRoot) {
    $here = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $here = Split-Path -LiteralPath $MyInvocation.MyCommand.Path -Parent
} else {
    $here = (Get-Location).Path
}

# 載入模組
$modules = @('Core', 'Security', 'Logging', 'Backup', 'Users', 'Projects')
foreach ($m in $modules) {
    Import-Module -Name (Join-Path $here ("Modules\$m.psm1")) -Force -DisableNameChecking
}

# 初始化
Initialize-CwtCore -RootPath $here
if ($PSBoundParameters.ContainsKey('DataRoot') -and -not [string]::IsNullOrWhiteSpace($DataRoot)) {
    $cfgPath = Join-Path $here 'config.json'
    $cfg = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $cfg.DataRoot = $DataRoot
    ($cfg | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $cfgPath -Encoding UTF8
    Initialize-CwtCore -RootPath $here
}

$paths = Get-CwtPaths
Initialize-CwtMasterKey -KeyPath $paths.MasterKey | Out-Null
Initialize-CwtDefaultAdmin | Out-Null

# 啟動時自動備份
try {
    $cfg = Get-CwtConfig
    if ($cfg.AutoBackupOnStart) { Invoke-CwtBackup -Tag 'startup' | Out-Null }
} catch {
    Write-CwtLog -Action 'backup.startup.fail' -Detail $_.Exception.Message -Level 'Warning'
}

# 載入 UI
foreach ($u in @('LoginForm','RegisterForm','MainForm','ProjectForm','AdminForm')) {
    . (Join-Path $here ("UI\$u.ps1"))
}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

Write-CwtLog -Action 'app.start' -Detail "host=$env:COMPUTERNAME user=$env:USERNAME"

try {
    $user = Show-CwtLoginForm
    if ($null -ne $user) {
        Show-CwtMainForm -User $user
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("發生未預期錯誤:`r`n$($_.Exception.Message)", '錯誤', 'OK', 'Error') | Out-Null
    Write-CwtLog -Action 'app.error' -Detail $_.Exception.Message -Level 'Error'
} finally {
    Write-CwtLog -Action 'app.exit'
}
