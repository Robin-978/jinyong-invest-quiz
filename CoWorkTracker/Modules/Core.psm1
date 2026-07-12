# Core.psm1 - JSON store, file locking, paths, config
# PowerShell 5.1 compatible
Set-StrictMode -Version Latest

$script:Config = $null
$script:Paths  = $null

function Initialize-CwtCore {
    param(
        [Parameter(Mandatory)][string]$RootPath
    )

    $configPath = Join-Path $RootPath 'config.json'
    if (-not (Test-Path $configPath)) {
        throw "找不到設定檔: $configPath"
    }
    $cfg = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

    if ([string]::IsNullOrWhiteSpace($cfg.DataRoot)) {
        $dataRoot = $RootPath
    } else {
        $dataRoot = $cfg.DataRoot
    }

    $paths = [pscustomobject]@{
        Root        = $RootPath
        DataRoot    = $dataRoot
        Data        = Join-Path $dataRoot $cfg.DataFolder
        Logs        = Join-Path $dataRoot $cfg.LogsFolder
        Backups     = Join-Path $dataRoot $cfg.BackupsFolder
        Uploads     = Join-Path $dataRoot $cfg.UploadsFolder
        Users       = Join-Path (Join-Path $dataRoot $cfg.DataFolder) 'users.json'
        Projects    = Join-Path (Join-Path $dataRoot $cfg.DataFolder) 'projects.json'
        Memberships = Join-Path (Join-Path $dataRoot $cfg.DataFolder) 'memberships.json'
        JoinReqs    = Join-Path (Join-Path $dataRoot $cfg.DataFolder) 'join_requests.json'
        MasterKey   = Join-Path $dataRoot $cfg.MasterKeyFile
    }

    foreach ($p in @($paths.Data, $paths.Logs, $paths.Backups, $paths.Uploads)) {
        if (-not (Test-Path $p)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }
    }

    foreach ($f in @($paths.Users, $paths.Projects, $paths.Memberships, $paths.JoinReqs)) {
        if (-not (Test-Path $f)) {
            Set-Content -LiteralPath $f -Value '[]' -Encoding UTF8
        }
    }

    $script:Config = $cfg
    $script:Paths  = $paths
}

function Get-CwtConfig { return $script:Config }
function Get-CwtPaths  { return $script:Paths }

function Read-CwtJson {
    param([Parameter(Mandatory)][string]$Path)
    $cfg = $script:Config
    if ($null -eq $cfg) { throw '尚未初始化 Core' }
    $retry = [int]$cfg.FileLockRetry
    $delay = [int]$cfg.FileLockDelayMs
    for ($i = 0; $i -lt $retry; $i++) {
        try {
            $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read')
            try {
                $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                try {
                    $txt = $sr.ReadToEnd()
                } finally { $sr.Dispose() }
            } finally { $fs.Dispose() }
            if ([string]::IsNullOrWhiteSpace($txt)) { return }
            $obj = $txt | ConvertFrom-Json
            if ($null -eq $obj) { return }
            foreach ($item in @($obj)) { Write-Output $item }
            return
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds $delay
        }
    }
    throw "讀取檔案逾時: $Path"
}

function Write-CwtJson {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()]$Data
    )
    $cfg = $script:Config
    if ($null -eq $cfg) { throw '尚未初始化 Core' }
    if ($null -eq $Data) { $Data = @() }

    $arr = @($Data)
    $json = $arr | ConvertTo-Json -Depth 25
    if ($null -eq $json -or $json.Length -eq 0) { $json = '[]' }
    if ($arr.Count -eq 1 -and $json.TrimStart().StartsWith('{')) {
        $json = "[`r`n$json`r`n]"
    }

    $retry = [int]$cfg.FileLockRetry
    $delay = [int]$cfg.FileLockDelayMs
    $tmp = $Path + '.tmp'
    for ($i = 0; $i -lt $retry; $i++) {
        try {
            $fs = [System.IO.File]::Open($tmp, 'Create', 'Write', 'None')
            try {
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
                $fs.Write($bytes, 0, $bytes.Length)
                $fs.Flush($true)
            } finally { $fs.Dispose() }
            if (Test-Path $Path) {
                $bak = $Path + '.bak'
                Copy-Item -LiteralPath $Path -Destination $bak -Force
            }
            Move-Item -LiteralPath $tmp -Destination $Path -Force
            return
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds $delay
        }
    }
    throw "寫入檔案逾時: $Path"
}

function New-CwtId {
    return [guid]::NewGuid().ToString('N')
}

function Get-CwtUtcNow {
    return (Get-Date).ToUniversalTime().ToString('o')
}

function ConvertTo-CwtSafeString {
    param([AllowNull()]$Value, [string]$Default = '')
    if ($null -eq $Value) { return $Default }
    $s = [string]$Value
    if ([string]::IsNullOrEmpty($s)) { return $Default }
    return $s
}

Export-ModuleMember -Function Initialize-CwtCore, Get-CwtConfig, Get-CwtPaths,
    Read-CwtJson, Write-CwtJson, New-CwtId, Get-CwtUtcNow, ConvertTo-CwtSafeString
