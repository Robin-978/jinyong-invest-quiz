# Users.psm1 - user CRUD and authentication
Set-StrictMode -Version Latest

function Get-CwtUsers {
    $p = Get-CwtPaths
    return @(Read-CwtJson -Path $p.Users)
}

function Save-CwtUsers {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Users)
    $p = Get-CwtPaths
    Write-CwtJson -Path $p.Users -Data $Users
}

function Get-CwtUserByName {
    param([Parameter(Mandatory)][string]$Username)
    $users = Get-CwtUsers
    foreach ($u in $users) {
        if ($null -ne $u -and $u.username -eq $Username) { return $u }
    }
    return $null
}

function Get-CwtUserById {
    param([Parameter(Mandatory)][string]$Id)
    if ([string]::IsNullOrEmpty($Id)) { return $null }
    $users = Get-CwtUsers
    foreach ($u in $users) {
        if ($null -ne $u -and $u.id -eq $Id) { return $u }
    }
    return $null
}

function New-CwtUser {
    param(
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][string]$Password,
        [string]$DisplayName = '',
        [string]$Email = '',
        [ValidateSet('Admin','User')][string]$Role = 'User',
        [bool]$Active = $true
    )
    if ([string]::IsNullOrWhiteSpace($Username)) { throw '帳號不可為空' }
    if ([string]::IsNullOrWhiteSpace($Password)) { throw '密碼不可為空' }
    if (Get-CwtUserByName -Username $Username) { throw "帳號已存在: $Username" }

    $cfg = Get-CwtConfig
    $salt = New-CwtSalt
    $hash = Get-CwtPasswordHash -Password $Password -SaltBase64 $salt -Iterations ([int]$cfg.PBKDF2Iterations)

    $user = [pscustomobject]@{
        id           = New-CwtId
        username     = $Username
        displayName  = (ConvertTo-CwtSafeString $DisplayName $Username)
        email        = (ConvertTo-CwtSafeString $Email '')
        role         = $Role
        passwordHash = $hash
        salt         = $salt
        iterations   = [int]$cfg.PBKDF2Iterations
        active       = [bool]$Active
        createdAt    = Get-CwtUtcNow
        updatedAt    = Get-CwtUtcNow
    }
    $users = @(Get-CwtUsers)
    $users += $user
    Save-CwtUsers -Users $users
    Write-CwtLog -Action 'user.create' -ActorName $Username -Target $user.id -Detail "role=$Role" -Level 'Security'
    return $user
}

function Update-CwtUser {
    param(
        [Parameter(Mandatory)][string]$Id,
        [hashtable]$Set
    )
    if ($null -eq $Set) { return $null }
    $users = @(Get-CwtUsers)
    $updated = $null
    for ($i = 0; $i -lt $users.Count; $i++) {
        if ($users[$i].id -eq $Id) {
            foreach ($k in $Set.Keys) {
                if ($users[$i].PSObject.Properties.Name -contains $k) {
                    $users[$i].$k = $Set[$k]
                } else {
                    $users[$i] | Add-Member -NotePropertyName $k -NotePropertyValue $Set[$k]
                }
            }
            $users[$i].updatedAt = Get-CwtUtcNow
            $updated = $users[$i]
            break
        }
    }
    if ($null -ne $updated) {
        Save-CwtUsers -Users $users
        Write-CwtLog -Action 'user.update' -Target $Id -Detail (($Set.Keys) -join ',') -Level 'Security'
    }
    return $updated
}

function Set-CwtUserPassword {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$NewPassword
    )
    if ([string]::IsNullOrWhiteSpace($NewPassword)) { throw '密碼不可為空' }
    $cfg = Get-CwtConfig
    $salt = New-CwtSalt
    $hash = Get-CwtPasswordHash -Password $NewPassword -SaltBase64 $salt -Iterations ([int]$cfg.PBKDF2Iterations)
    return Update-CwtUser -Id $Id -Set @{ salt = $salt; passwordHash = $hash; iterations = [int]$cfg.PBKDF2Iterations }
}

function Test-CwtLogin {
    param(
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][string]$Password
    )
    $u = Get-CwtUserByName -Username $Username
    if ($null -eq $u) { return $null }
    if (-not $u.active) { return $null }
    $iter = 100000
    if ($u.PSObject.Properties.Name -contains 'iterations') { $iter = [int]$u.iterations }
    if (Test-CwtPassword -Password $Password -SaltBase64 $u.salt -HashBase64 $u.passwordHash -Iterations $iter) {
        return $u
    }
    return $null
}

function Initialize-CwtDefaultAdmin {
    $users = @(Get-CwtUsers)
    if ($users.Count -gt 0) { return $null }
    return (New-CwtUser -Username 'admin' -Password 'admin@123' -DisplayName '系統管理員' -Role 'Admin')
}

Export-ModuleMember -Function Get-CwtUsers, Save-CwtUsers, Get-CwtUserByName, Get-CwtUserById,
    New-CwtUser, Update-CwtUser, Set-CwtUserPassword, Test-CwtLogin, Initialize-CwtDefaultAdmin
