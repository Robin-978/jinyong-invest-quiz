# Logging.psm1 - encrypted append-only log
# Each line = AES-encrypted JSON entry. PowerShell 5.1 compatible.
Set-StrictMode -Version Latest

function Get-CwtLogFile {
    $paths = Get-CwtPaths
    $today = (Get-Date).ToString('yyyyMMdd')
    return (Join-Path $paths.Logs ("log_$today.enc"))
}

function Write-CwtLog {
    param(
        [Parameter(Mandatory)][string]$Action,
        [string]$ActorId = '',
        [string]$ActorName = '',
        [string]$Target = '',
        [string]$Detail = '',
        [ValidateSet('Info','Warning','Error','Security')]
        [string]$Level = 'Info'
    )
    try {
        $paths = Get-CwtPaths
        $key = Get-CwtMasterKey -KeyPath $paths.MasterKey
        $entry = [ordered]@{
            ts        = Get-CwtUtcNow
            level     = $Level
            action    = $Action
            actorId   = $ActorId
            actorName = $ActorName
            target    = $Target
            detail    = $Detail
            host      = $env:COMPUTERNAME
            user      = $env:USERNAME
        }
        $json = ($entry | ConvertTo-Json -Compress -Depth 10)
        $cipher = Protect-CwtString -Plain $json -Key $key
        $logFile = Get-CwtLogFile
        $line = $cipher + "`r`n"
        $cfg = Get-CwtConfig
        $retry = [int]$cfg.FileLockRetry
        $delay = [int]$cfg.FileLockDelayMs
        for ($i = 0; $i -lt $retry; $i++) {
            try {
                $fs = [System.IO.File]::Open($logFile, 'Append', 'Write', 'Read')
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($line)
                    $fs.Write($bytes, 0, $bytes.Length)
                    $fs.Flush($true)
                } finally { $fs.Dispose() }
                return
            } catch [System.IO.IOException] {
                Start-Sleep -Milliseconds $delay
            }
        }
    } catch {
        # 紀錄失敗時不影響主程式; 寫入 fallback 純文字檔
        try {
            $fb = Join-Path (Get-CwtPaths).Logs 'log_errors.txt'
            $msg = "{0} LOG_FAIL: {1}" -f (Get-CwtUtcNow), $_.Exception.Message
            Add-Content -LiteralPath $fb -Value $msg -Encoding UTF8
        } catch { }
    }
}

function Read-CwtLogs {
    param(
        [datetime]$From,
        [datetime]$To,
        [string]$ActorId,
        [string]$Action,
        [string]$Level
    )
    $paths = Get-CwtPaths
    $key = Get-CwtMasterKey -KeyPath $paths.MasterKey
    $files = @(Get-ChildItem -LiteralPath $paths.Logs -Filter 'log_*.enc' -File -ErrorAction SilentlyContinue)
    $results = New-Object System.Collections.ArrayList
    foreach ($f in $files) {
        $lines = @()
        try { $lines = Get-Content -LiteralPath $f.FullName -Encoding UTF8 -ErrorAction Stop } catch { continue }
        foreach ($ln in $lines) {
            if ([string]::IsNullOrWhiteSpace($ln)) { continue }
            try {
                $plain = Unprotect-CwtString -CipherBase64 $ln -Key $key
                if ([string]::IsNullOrWhiteSpace($plain)) { continue }
                $obj = $plain | ConvertFrom-Json
                if ($null -eq $obj) { continue }
                if ($From) {
                    $ts = [datetime]::Parse($obj.ts).ToUniversalTime()
                    if ($ts -lt $From.ToUniversalTime()) { continue }
                }
                if ($To) {
                    $ts = [datetime]::Parse($obj.ts).ToUniversalTime()
                    if ($ts -gt $To.ToUniversalTime()) { continue }
                }
                if ($ActorId -and $obj.actorId -ne $ActorId) { continue }
                if ($Action  -and $obj.action  -notlike "*$Action*") { continue }
                if ($Level   -and $obj.level   -ne $Level) { continue }
                [void]$results.Add($obj)
            } catch { continue }
        }
    }
    return @($results)
}

function Export-CwtLogsCsv {
    param(
        [Parameter(Mandatory)][string]$Path,
        [datetime]$From,
        [datetime]$To,
        [string]$ActorId,
        [string]$Action,
        [string]$Level
    )
    $entries = Read-CwtLogs -From $From -To $To -ActorId $ActorId -Action $Action -Level $Level
    $entries | Select-Object ts, level, action, actorId, actorName, target, detail, host, user |
        Export-Csv -LiteralPath $Path -Encoding UTF8 -NoTypeInformation
}

Export-ModuleMember -Function Write-CwtLog, Read-CwtLogs, Export-CwtLogsCsv, Get-CwtLogFile
