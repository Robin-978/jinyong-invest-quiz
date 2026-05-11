# Backup.psm1 - backup/restore Data + Logs as ZIP
Set-StrictMode -Version Latest

function Invoke-CwtBackup {
    param([string]$Tag = '')
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $paths = Get-CwtPaths
    $cfg = Get-CwtConfig
    $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $name = if ([string]::IsNullOrWhiteSpace($Tag)) { "backup_$ts.zip" } else { "backup_${ts}_$Tag.zip" }
    $dest = Join-Path $paths.Backups $name
    if (Test-Path $dest) { Remove-Item -LiteralPath $dest -Force }

    $staging = Join-Path $paths.Backups ("_stage_$ts")
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    try {
        Copy-Item -LiteralPath $paths.Data -Destination (Join-Path $staging 'Data') -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item -LiteralPath $paths.Logs -Destination (Join-Path $staging 'Logs') -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path $paths.Uploads) {
            Copy-Item -LiteralPath $paths.Uploads -Destination (Join-Path $staging 'Uploads') -Recurse -Force -ErrorAction SilentlyContinue
        }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $dest)
    } finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }

    $keep = [int]$cfg.BackupKeepCount
    if ($keep -gt 0) {
        $all = @(Get-ChildItem -LiteralPath $paths.Backups -Filter 'backup_*.zip' -File | Sort-Object LastWriteTime -Descending)
        if ($all.Count -gt $keep) {
            $all | Select-Object -Skip $keep | ForEach-Object {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    }
    return $dest
}

function Get-CwtBackupList {
    $paths = Get-CwtPaths
    return @(Get-ChildItem -LiteralPath $paths.Backups -Filter 'backup_*.zip' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending)
}

function Restore-CwtBackup {
    param([Parameter(Mandatory)][string]$ZipPath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    if (-not (Test-Path $ZipPath)) { throw "備份檔不存在: $ZipPath" }
    $paths = Get-CwtPaths
    $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $preZip = Invoke-CwtBackup -Tag "before_restore_$ts"

    $extract = Join-Path $paths.Backups ("_restore_$ts")
    if (Test-Path $extract) { Remove-Item -LiteralPath $extract -Recurse -Force }
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $extract)
    try {
        $srcData = Join-Path $extract 'Data'
        $srcLogs = Join-Path $extract 'Logs'
        $srcUp   = Join-Path $extract 'Uploads'
        if (Test-Path $srcData) {
            Get-ChildItem -LiteralPath $paths.Data -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item -LiteralPath (Join-Path $srcData '*') -Destination $paths.Data -Recurse -Force
        }
        if (Test-Path $srcLogs) {
            Get-ChildItem -LiteralPath $paths.Logs -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item -LiteralPath (Join-Path $srcLogs '*') -Destination $paths.Logs -Recurse -Force
        }
        if (Test-Path $srcUp) {
            if (-not (Test-Path $paths.Uploads)) { New-Item -ItemType Directory -Path $paths.Uploads -Force | Out-Null }
            Get-ChildItem -LiteralPath $paths.Uploads -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item -LiteralPath (Join-Path $srcUp '*') -Destination $paths.Uploads -Recurse -Force
        }
    } finally {
        Remove-Item -LiteralPath $extract -Recurse -Force -ErrorAction SilentlyContinue
    }
    return $preZip
}

Export-ModuleMember -Function Invoke-CwtBackup, Get-CwtBackupList, Restore-CwtBackup
