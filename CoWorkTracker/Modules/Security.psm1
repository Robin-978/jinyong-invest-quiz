# Security.psm1 - PBKDF2 password hashing, AES-256-CBC encryption
# PowerShell 5.1 compatible
Set-StrictMode -Version Latest

function New-CwtSalt {
    param([int]$Bytes = 32)
    $b = New-Object byte[] $Bytes
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($b) } finally { $rng.Dispose() }
    return [Convert]::ToBase64String($b)
}

function Get-CwtPasswordHash {
    param(
        [Parameter(Mandatory)][string]$Password,
        [Parameter(Mandatory)][string]$SaltBase64,
        [int]$Iterations = 100000
    )
    if ([string]::IsNullOrEmpty($Password)) { throw '密碼不可為空' }
    $salt = [Convert]::FromBase64String($SaltBase64)
    $kdf = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $salt, $Iterations)
    try {
        $bytes = $kdf.GetBytes(32)
        return [Convert]::ToBase64String($bytes)
    } finally { $kdf.Dispose() }
}

function Test-CwtPassword {
    param(
        [Parameter(Mandatory)][string]$Password,
        [Parameter(Mandatory)][string]$SaltBase64,
        [Parameter(Mandatory)][string]$HashBase64,
        [int]$Iterations = 100000
    )
    if ([string]::IsNullOrEmpty($Password)) { return $false }
    if ([string]::IsNullOrEmpty($SaltBase64)) { return $false }
    if ([string]::IsNullOrEmpty($HashBase64)) { return $false }
    $calc = Get-CwtPasswordHash -Password $Password -SaltBase64 $SaltBase64 -Iterations $Iterations
    $a = [Convert]::FromBase64String($HashBase64)
    $b = [Convert]::FromBase64String($calc)
    if ($a.Length -ne $b.Length) { return $false }
    $diff = 0
    for ($i = 0; $i -lt $a.Length; $i++) { $diff = $diff -bor ($a[$i] -bxor $b[$i]) }
    return ($diff -eq 0)
}

function Initialize-CwtMasterKey {
    param([Parameter(Mandatory)][string]$KeyPath)
    if (Test-Path $KeyPath) { return }
    $dir = [System.IO.Path]::GetDirectoryName($KeyPath)
    if (-not [string]::IsNullOrEmpty($dir) -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $key = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($key) } finally { $rng.Dispose() }
    [System.IO.File]::WriteAllBytes($KeyPath, $key)
}

function Get-CwtMasterKey {
    param([Parameter(Mandatory)][string]$KeyPath)
    if (-not (Test-Path $KeyPath)) {
        Initialize-CwtMasterKey -KeyPath $KeyPath
    }
    return [System.IO.File]::ReadAllBytes($KeyPath)
}

function Protect-CwtString {
    param(
        [Parameter(Mandatory)][string]$Plain,
        [Parameter(Mandatory)][byte[]]$Key
    )
    if ($null -eq $Plain) { $Plain = '' }
    $aes = [System.Security.Cryptography.Aes]::Create()
    try {
        $aes.KeySize = 256
        $aes.BlockSize = 128
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.Key = $Key
        $aes.GenerateIV()
        $iv = $aes.IV
        $enc = $aes.CreateEncryptor()
        try {
            $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($Plain)
            $cipher = $enc.TransformFinalBlock($plainBytes, 0, $plainBytes.Length)
        } finally { $enc.Dispose() }
        $out = New-Object byte[] ($iv.Length + $cipher.Length)
        [Array]::Copy($iv, 0, $out, 0, $iv.Length)
        [Array]::Copy($cipher, 0, $out, $iv.Length, $cipher.Length)
        return [Convert]::ToBase64String($out)
    } finally { $aes.Dispose() }
}

function Unprotect-CwtString {
    param(
        [Parameter(Mandatory)][string]$CipherBase64,
        [Parameter(Mandatory)][byte[]]$Key
    )
    if ([string]::IsNullOrEmpty($CipherBase64)) { return '' }
    $blob = [Convert]::FromBase64String($CipherBase64)
    if ($blob.Length -le 16) { return '' }
    $iv = New-Object byte[] 16
    [Array]::Copy($blob, 0, $iv, 0, 16)
    $cipher = New-Object byte[] ($blob.Length - 16)
    [Array]::Copy($blob, 16, $cipher, 0, $cipher.Length)
    $aes = [System.Security.Cryptography.Aes]::Create()
    try {
        $aes.KeySize = 256
        $aes.BlockSize = 128
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.Key = $Key
        $aes.IV = $iv
        $dec = $aes.CreateDecryptor()
        try {
            $plainBytes = $dec.TransformFinalBlock($cipher, 0, $cipher.Length)
            return [System.Text.Encoding]::UTF8.GetString($plainBytes)
        } finally { $dec.Dispose() }
    } finally { $aes.Dispose() }
}

function Get-CwtFileHash {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return '' }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        try {
            $h = $sha.ComputeHash($fs)
            return [BitConverter]::ToString($h).Replace('-', '').ToLowerInvariant()
        } finally { $fs.Dispose() }
    } finally { $sha.Dispose() }
}

Export-ModuleMember -Function New-CwtSalt, Get-CwtPasswordHash, Test-CwtPassword,
    Initialize-CwtMasterKey, Get-CwtMasterKey, Protect-CwtString, Unprotect-CwtString,
    Get-CwtFileHash
