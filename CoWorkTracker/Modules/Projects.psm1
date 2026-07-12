# Projects.psm1 - project, phase, membership, join request management
Set-StrictMode -Version Latest

function Get-CwtProjects {
    $p = Get-CwtPaths
    return @(Read-CwtJson -Path $p.Projects)
}
function Save-CwtProjects {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Projects)
    Write-CwtJson -Path (Get-CwtPaths).Projects -Data $Projects
}
function Get-CwtMemberships {
    return @(Read-CwtJson -Path (Get-CwtPaths).Memberships)
}
function Save-CwtMemberships {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data)
    Write-CwtJson -Path (Get-CwtPaths).Memberships -Data $Data
}
function Get-CwtJoinRequests {
    return @(Read-CwtJson -Path (Get-CwtPaths).JoinReqs)
}
function Save-CwtJoinRequests {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data)
    Write-CwtJson -Path (Get-CwtPaths).JoinReqs -Data $Data
}

function Get-CwtProjectById {
    param([Parameter(Mandatory)][string]$Id)
    foreach ($p in (Get-CwtProjects)) {
        if ($null -ne $p -and $p.id -eq $Id) { return $p }
    }
    return $null
}

function Get-CwtProjectByCode {
    param([Parameter(Mandatory)][string]$Code)
    foreach ($p in (Get-CwtProjects)) {
        if ($null -ne $p -and $p.code -eq $Code) { return $p }
    }
    return $null
}

function New-CwtProjectCode {
    param([string]$Prefix = 'PRJ')
    $today = (Get-Date).ToString('yyyyMMdd')
    $projects = @(Get-CwtProjects)
    $max = 0
    foreach ($p in $projects) {
        if ($null -eq $p -or $null -eq $p.code) { continue }
        if ($p.code -match "^$Prefix-$today-(\d{3,})$") {
            $n = [int]$matches[1]
            if ($n -gt $max) { $max = $n }
        }
    }
    $next = $max + 1
    return ('{0}-{1}-{2:D3}' -f $Prefix, $today, $next)
}

function New-CwtProject {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$OwnerId,
        [string]$Code = '',
        [string]$Description = '',
        [ValidateSet('Public','Internal','Confidential','TopSecret')][string]$Confidentiality = 'Internal',
        [datetime]$StartDate = (Get-Date),
        [Nullable[datetime]]$Deadline = $null,
        [ValidateSet('Planning','InProgress','OnHold','Completed','Cancelled')][string]$Status = 'Planning'
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { throw '專案名稱不可為空' }
    if ([string]::IsNullOrWhiteSpace($OwnerId)) { throw '必須指定發起人' }
    if ([string]::IsNullOrWhiteSpace($Code)) { $Code = New-CwtProjectCode }
    if (Get-CwtProjectByCode -Code $Code) { throw "專案編碼已存在: $Code" }

    $deadlineStr = ''
    if ($Deadline) { $deadlineStr = $Deadline.ToUniversalTime().ToString('o') }

    $proj = [pscustomobject]@{
        id              = New-CwtId
        code            = $Code
        name            = $Name
        description     = (ConvertTo-CwtSafeString $Description)
        confidentiality = $Confidentiality
        ownerId         = $OwnerId
        startDate       = $StartDate.ToUniversalTime().ToString('o')
        deadline        = $deadlineStr
        status          = $Status
        phases          = @()
        createdAt       = Get-CwtUtcNow
        updatedAt       = Get-CwtUtcNow
    }
    $list = @(Get-CwtProjects)
    $list += $proj
    Save-CwtProjects -Projects $list

    # 發起人自動為 Owner 權限
    Set-CwtMembership -ProjectId $proj.id -UserId $OwnerId -Permission 'Owner' -GrantedBy $OwnerId | Out-Null

    Write-CwtLog -Action 'project.create' -ActorId $OwnerId -Target $proj.id -Detail $proj.code -Level 'Info'
    return $proj
}

function Update-CwtProject {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][hashtable]$Set,
        [string]$ActorId = ''
    )
    $list = @(Get-CwtProjects)
    $updated = $null
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].id -eq $Id) {
            foreach ($k in $Set.Keys) {
                if ($list[$i].PSObject.Properties.Name -contains $k) {
                    $list[$i].$k = $Set[$k]
                } else {
                    $list[$i] | Add-Member -NotePropertyName $k -NotePropertyValue $Set[$k]
                }
            }
            $list[$i].updatedAt = Get-CwtUtcNow
            $updated = $list[$i]
            break
        }
    }
    if ($null -ne $updated) {
        Save-CwtProjects -Projects $list
        Write-CwtLog -Action 'project.update' -ActorId $ActorId -Target $Id -Detail (($Set.Keys) -join ',')
    }
    return $updated
}

function Remove-CwtProject {
    param(
        [Parameter(Mandatory)][string]$Id,
        [string]$ActorId = ''
    )
    $list = @(Get-CwtProjects | Where-Object { $_.id -ne $Id })
    Save-CwtProjects -Projects $list
    $mem = @(Get-CwtMemberships | Where-Object { $_.projectId -ne $Id })
    Save-CwtMemberships -Data $mem
    $jr = @(Get-CwtJoinRequests | Where-Object { $_.projectId -ne $Id })
    Save-CwtJoinRequests -Data $jr
    Write-CwtLog -Action 'project.delete' -ActorId $ActorId -Target $Id -Level 'Warning'
}

# ------- Memberships -------

function Get-CwtMembership {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$UserId
    )
    foreach ($m in (Get-CwtMemberships)) {
        if ($null -ne $m -and $m.projectId -eq $ProjectId -and $m.userId -eq $UserId) { return $m }
    }
    return $null
}

function Get-CwtProjectMembers {
    param([Parameter(Mandatory)][string]$ProjectId)
    return @(Get-CwtMemberships | Where-Object { $_.projectId -eq $ProjectId })
}

function Get-CwtUserProjects {
    param([Parameter(Mandatory)][string]$UserId)
    $mids = @(Get-CwtMemberships | Where-Object { $_.userId -eq $UserId } | ForEach-Object { $_.projectId })
    return @(Get-CwtProjects | Where-Object { $mids -contains $_.id })
}

function Set-CwtMembership {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][ValidateSet('Owner','Editor','Viewer')][string]$Permission,
        [string]$GrantedBy = ''
    )
    $list = @(Get-CwtMemberships)
    $found = $false
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].projectId -eq $ProjectId -and $list[$i].userId -eq $UserId) {
            $list[$i].permission = $Permission
            $list[$i].grantedBy = $GrantedBy
            $list[$i].grantedAt = Get-CwtUtcNow
            $found = $true
            break
        }
    }
    if (-not $found) {
        $list += [pscustomobject]@{
            projectId  = $ProjectId
            userId     = $UserId
            permission = $Permission
            grantedBy  = $GrantedBy
            grantedAt  = Get-CwtUtcNow
        }
    }
    Save-CwtMemberships -Data $list
    Write-CwtLog -Action 'membership.set' -ActorId $GrantedBy -Target "$ProjectId/$UserId" -Detail $Permission -Level 'Security'
}

function Remove-CwtMembership {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$UserId,
        [string]$ActorId = ''
    )
    $list = @(Get-CwtMemberships | Where-Object { -not ($_.projectId -eq $ProjectId -and $_.userId -eq $UserId) })
    Save-CwtMemberships -Data $list
    Write-CwtLog -Action 'membership.remove' -ActorId $ActorId -Target "$ProjectId/$UserId" -Level 'Security'
}

function Test-CwtCanEdit {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)]$User
    )
    if ($null -eq $User) { return $false }
    if ($User.role -eq 'Admin') { return $true }
    $m = Get-CwtMembership -ProjectId $ProjectId -UserId $User.id
    if ($null -eq $m) { return $false }
    return ($m.permission -eq 'Owner' -or $m.permission -eq 'Editor')
}
function Test-CwtCanView {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)]$User
    )
    if ($null -eq $User) { return $false }
    if ($User.role -eq 'Admin') { return $true }
    return ($null -ne (Get-CwtMembership -ProjectId $ProjectId -UserId $User.id))
}
function Test-CwtCanManage {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)]$User
    )
    if ($null -eq $User) { return $false }
    if ($User.role -eq 'Admin') { return $true }
    $m = Get-CwtMembership -ProjectId $ProjectId -UserId $User.id
    if ($null -eq $m) { return $false }
    return ($m.permission -eq 'Owner')
}

# ------- Join Requests -------

function New-CwtJoinRequest {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$UserId,
        [ValidateSet('Editor','Viewer')][string]$RequestedPermission = 'Viewer',
        [string]$Message = ''
    )
    $existing = Get-CwtMembership -ProjectId $ProjectId -UserId $UserId
    if ($null -ne $existing) { throw '您已是該專案成員' }
    $pending = @(Get-CwtJoinRequests | Where-Object {
        $_.projectId -eq $ProjectId -and $_.userId -eq $UserId -and $_.status -eq 'Pending'
    })
    if ($pending.Count -gt 0) { throw '已有待審核的申請' }

    $req = [pscustomobject]@{
        id                  = New-CwtId
        projectId           = $ProjectId
        userId              = $UserId
        requestedPermission = $RequestedPermission
        message             = (ConvertTo-CwtSafeString $Message)
        status              = 'Pending'
        requestedAt         = Get-CwtUtcNow
        decidedBy           = ''
        decidedAt           = ''
    }
    $list = @(Get-CwtJoinRequests)
    $list += $req
    Save-CwtJoinRequests -Data $list
    Write-CwtLog -Action 'project.join.request' -ActorId $UserId -Target $ProjectId -Detail $RequestedPermission
    return $req
}

function Resolve-CwtJoinRequest {
    param(
        [Parameter(Mandatory)][string]$RequestId,
        [Parameter(Mandatory)][ValidateSet('Approved','Rejected')][string]$Decision,
        [Parameter(Mandatory)][string]$DeciderId,
        [ValidateSet('Owner','Editor','Viewer')][string]$GrantedPermission = ''
    )
    $list = @(Get-CwtJoinRequests)
    $target = $null
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].id -eq $RequestId) {
            $list[$i].status = $Decision
            $list[$i].decidedBy = $DeciderId
            $list[$i].decidedAt = Get-CwtUtcNow
            $target = $list[$i]
            break
        }
    }
    if ($null -eq $target) { throw '找不到申請紀錄' }
    Save-CwtJoinRequests -Data $list
    if ($Decision -eq 'Approved') {
        $perm = if ([string]::IsNullOrWhiteSpace($GrantedPermission)) { $target.requestedPermission } else { $GrantedPermission }
        Set-CwtMembership -ProjectId $target.projectId -UserId $target.userId -Permission $perm -GrantedBy $DeciderId
    }
    Write-CwtLog -Action "project.join.$($Decision.ToLower())" -ActorId $DeciderId -Target "$($target.projectId)/$($target.userId)" -Level 'Security'
    return $target
}

# ------- Phases -------

function Add-CwtPhase {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$Name,
        [string]$Kpi = '',
        [string]$LeaderId = '',
        [string[]]$CollaboratorIds = @(),
        [Nullable[datetime]]$StartDate = $null,
        [Nullable[datetime]]$DueDate = $null,
        [string]$ActorId = ''
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { throw '階段名稱不可為空' }
    $list = @(Get-CwtProjects)
    $proj = $null
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].id -eq $ProjectId) {
            $startStr = ''; if ($StartDate) { $startStr = $StartDate.ToUniversalTime().ToString('o') }
            $dueStr   = ''; if ($DueDate)   { $dueStr   = $DueDate.ToUniversalTime().ToString('o') }
            $phase = [pscustomobject]@{
                id              = New-CwtId
                name            = $Name
                kpi             = (ConvertTo-CwtSafeString $Kpi)
                leaderId        = (ConvertTo-CwtSafeString $LeaderId)
                collaboratorIds = @($CollaboratorIds)
                startDate       = $startStr
                dueDate         = $dueStr
                progress        = 0
                progressNote    = ''
                status          = 'NotStarted'
                attachments     = @()
                updates         = @()
                createdAt       = Get-CwtUtcNow
                updatedAt       = Get-CwtUtcNow
            }
            $phases = @()
            if ($list[$i].PSObject.Properties.Name -contains 'phases' -and $null -ne $list[$i].phases) {
                $phases = @($list[$i].phases)
            }
            $phases += $phase
            $list[$i].phases = $phases
            $list[$i].updatedAt = Get-CwtUtcNow
            $proj = $list[$i]
            break
        }
    }
    if ($null -eq $proj) { throw '找不到專案' }
    Save-CwtProjects -Projects $list
    Write-CwtLog -Action 'phase.add' -ActorId $ActorId -Target "$ProjectId/$Name"
    return $proj
}

function Update-CwtPhase {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$PhaseId,
        [Parameter(Mandatory)][hashtable]$Set,
        [string]$ActorId = ''
    )
    $list = @(Get-CwtProjects)
    $updated = $null
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].id -ne $ProjectId) { continue }
        $phases = @()
        if ($list[$i].PSObject.Properties.Name -contains 'phases' -and $null -ne $list[$i].phases) {
            $phases = @($list[$i].phases)
        }
        for ($j = 0; $j -lt $phases.Count; $j++) {
            if ($phases[$j].id -eq $PhaseId) {
                foreach ($k in $Set.Keys) {
                    if ($phases[$j].PSObject.Properties.Name -contains $k) {
                        $phases[$j].$k = $Set[$k]
                    } else {
                        $phases[$j] | Add-Member -NotePropertyName $k -NotePropertyValue $Set[$k]
                    }
                }
                $phases[$j].updatedAt = Get-CwtUtcNow
                if ($Set.ContainsKey('progress')) {
                    $upd = @()
                    if ($phases[$j].PSObject.Properties.Name -contains 'updates' -and $null -ne $phases[$j].updates) {
                        $upd = @($phases[$j].updates)
                    }
                    $note = ''
                    if ($Set.ContainsKey('progressNote')) { $note = [string]$Set['progressNote'] }
                    $upd += [pscustomobject]@{
                        at       = Get-CwtUtcNow
                        by       = $ActorId
                        progress = [int]$Set['progress']
                        note     = $note
                    }
                    $phases[$j].updates = $upd
                }
                $updated = $phases[$j]
                break
            }
        }
        $list[$i].phases = $phases
        $list[$i].updatedAt = Get-CwtUtcNow
        break
    }
    if ($null -eq $updated) { throw '找不到階段' }
    Save-CwtProjects -Projects $list
    Write-CwtLog -Action 'phase.update' -ActorId $ActorId -Target "$ProjectId/$PhaseId" -Detail (($Set.Keys) -join ',')
    return $updated
}

function Remove-CwtPhase {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$PhaseId,
        [string]$ActorId = ''
    )
    $list = @(Get-CwtProjects)
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].id -ne $ProjectId) { continue }
        $phases = @()
        if ($list[$i].PSObject.Properties.Name -contains 'phases' -and $null -ne $list[$i].phases) {
            $phases = @($list[$i].phases | Where-Object { $_.id -ne $PhaseId })
        }
        $list[$i].phases = $phases
        $list[$i].updatedAt = Get-CwtUtcNow
        break
    }
    Save-CwtProjects -Projects $list
    Write-CwtLog -Action 'phase.remove' -ActorId $ActorId -Target "$ProjectId/$PhaseId" -Level 'Warning'
}

function Add-CwtPhaseAttachment {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$PhaseId,
        [Parameter(Mandatory)][string]$SourcePath,
        [string]$ActorId = ''
    )
    if (-not (Test-Path $SourcePath)) { throw "檔案不存在: $SourcePath" }
    $paths = Get-CwtPaths
    $proj = Get-CwtProjectById -Id $ProjectId
    if ($null -eq $proj) { throw '找不到專案' }
    $destDir = Join-Path (Join-Path $paths.Uploads $proj.code) $PhaseId
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    $fileName = [System.IO.Path]::GetFileName($SourcePath)
    $stamp = (Get-Date).ToString('yyyyMMddHHmmss')
    $destName = "$stamp`_$fileName"
    $destPath = Join-Path $destDir $destName
    Copy-Item -LiteralPath $SourcePath -Destination $destPath -Force

    $att = [pscustomobject]@{
        id         = New-CwtId
        name       = $fileName
        path       = $destPath
        size       = (Get-Item -LiteralPath $destPath).Length
        sha256     = Get-CwtFileHash -Path $destPath
        uploadedBy = $ActorId
        uploadedAt = Get-CwtUtcNow
    }

    $list = @(Get-CwtProjects)
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].id -ne $ProjectId) { continue }
        $phases = @()
        if ($list[$i].PSObject.Properties.Name -contains 'phases' -and $null -ne $list[$i].phases) {
            $phases = @($list[$i].phases)
        }
        for ($j = 0; $j -lt $phases.Count; $j++) {
            if ($phases[$j].id -eq $PhaseId) {
                $atts = @()
                if ($phases[$j].PSObject.Properties.Name -contains 'attachments' -and $null -ne $phases[$j].attachments) {
                    $atts = @($phases[$j].attachments)
                }
                $atts += $att
                $phases[$j].attachments = $atts
                $phases[$j].updatedAt = Get-CwtUtcNow
                break
            }
        }
        $list[$i].phases = $phases
        $list[$i].updatedAt = Get-CwtUtcNow
        break
    }
    Save-CwtProjects -Projects $list
    Write-CwtLog -Action 'phase.attach' -ActorId $ActorId -Target "$ProjectId/$PhaseId" -Detail $fileName
    return $att
}

function Get-CwtProjectProgress {
    param([Parameter(Mandatory)]$Project)
    if ($null -eq $Project) { return 0 }
    $phases = @()
    if ($Project.PSObject.Properties.Name -contains 'phases' -and $null -ne $Project.phases) {
        $phases = @($Project.phases)
    }
    if ($phases.Count -eq 0) { return 0 }
    $sum = 0
    foreach ($p in $phases) {
        $v = 0
        if ($p.PSObject.Properties.Name -contains 'progress' -and $null -ne $p.progress) {
            $v = [int]$p.progress
        }
        $sum += $v
    }
    return [int]([math]::Floor($sum / $phases.Count))
}

Export-ModuleMember -Function Get-CwtProjects, Save-CwtProjects, Get-CwtProjectById, Get-CwtProjectByCode,
    New-CwtProjectCode, New-CwtProject, Update-CwtProject, Remove-CwtProject,
    Get-CwtMemberships, Save-CwtMemberships, Get-CwtMembership, Get-CwtProjectMembers, Get-CwtUserProjects,
    Set-CwtMembership, Remove-CwtMembership, Test-CwtCanEdit, Test-CwtCanView, Test-CwtCanManage,
    Get-CwtJoinRequests, Save-CwtJoinRequests, New-CwtJoinRequest, Resolve-CwtJoinRequest,
    Add-CwtPhase, Update-CwtPhase, Remove-CwtPhase, Add-CwtPhaseAttachment, Get-CwtProjectProgress
