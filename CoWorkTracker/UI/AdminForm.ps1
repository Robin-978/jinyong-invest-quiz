# AdminForm.ps1 - 管理頁面 (人員、專案、Log)
function Show-CwtAdminForm {
    param([Parameter(Mandatory)]$CurrentUser)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    if ($CurrentUser.role -ne 'Admin') {
        [System.Windows.Forms.MessageBox]::Show('僅管理員可進入管理頁面', '權限不足', 'OK', 'Warning') | Out-Null
        return
    }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "管理頁面 - $($CurrentUser.displayName)"
    $f.Size = New-Object System.Drawing.Size(1050, 700); $f.StartPosition = 'CenterParent'

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'
    $f.Controls.Add($tabs)

    # --- 人員管理 ---
    $tabU = New-Object System.Windows.Forms.TabPage; $tabU.Text = '人員管理'
    $tabs.TabPages.Add($tabU)
    $gridU = New-Object System.Windows.Forms.DataGridView
    $gridU.Dock = 'Fill'; $gridU.ReadOnly = $true; $gridU.AllowUserToAddRows = $false
    $gridU.AllowUserToDeleteRows = $false; $gridU.SelectionMode = 'FullRowSelect'
    $gridU.MultiSelect = $false; $gridU.RowHeadersVisible = $false
    $gridU.AutoSizeColumnsMode = 'Fill'
    [void]$gridU.Columns.Add('username', '帳號')
    [void]$gridU.Columns.Add('displayName', '顯示名稱')
    [void]$gridU.Columns.Add('email', 'Email')
    [void]$gridU.Columns.Add('role', '角色')
    [void]$gridU.Columns.Add('active', '啟用')
    [void]$gridU.Columns.Add('createdAt', '建立時間')
    $tabU.Controls.Add($gridU)
    $pnlU = New-Object System.Windows.Forms.Panel
    $pnlU.Dock = 'Top'; $pnlU.Height = 45
    $tabU.Controls.Add($pnlU)
    $btnAddU = New-Object System.Windows.Forms.Button
    $btnAddU.Text = '＋ 新增使用者'; $btnAddU.Location = New-Object System.Drawing.Point(10, 8); $btnAddU.Size = New-Object System.Drawing.Size(120, 30)
    $pnlU.Controls.Add($btnAddU)
    $btnRole = New-Object System.Windows.Forms.Button
    $btnRole.Text = '切換管理員'; $btnRole.Location = New-Object System.Drawing.Point(140, 8); $btnRole.Size = New-Object System.Drawing.Size(110, 30)
    $pnlU.Controls.Add($btnRole)
    $btnToggle = New-Object System.Windows.Forms.Button
    $btnToggle.Text = '啟用/停用'; $btnToggle.Location = New-Object System.Drawing.Point(260, 8); $btnToggle.Size = New-Object System.Drawing.Size(100, 30)
    $pnlU.Controls.Add($btnToggle)
    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = '重設密碼'; $btnReset.Location = New-Object System.Drawing.Point(370, 8); $btnReset.Size = New-Object System.Drawing.Size(100, 30)
    $pnlU.Controls.Add($btnReset)

    $script:UserRows = @()
    function script:Reload-Users {
        $gridU.Rows.Clear()
        $list = @(Get-CwtUsers | Sort-Object username)
        $script:UserRows = $list
        foreach ($u in $list) {
            $tm = ''
            if (-not [string]::IsNullOrEmpty($u.createdAt)) {
                try { $tm = ([datetime]$u.createdAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { $tm = $u.createdAt }
            }
            [void]$gridU.Rows.Add($u.username, $u.displayName, $u.email, $u.role, $u.active, $tm)
        }
    }

    $btnAddU.Add_Click({
        $r = Show-CwtRegisterForm
        if ($null -ne $r) { Reload-Users }
    })
    $btnRole.Add_Click({
        if ($gridU.SelectedRows.Count -eq 0) { return }
        $idx = $gridU.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:UserRows.Count) { return }
        $u = $script:UserRows[$idx]
        if ($u.id -eq $CurrentUser.id) {
            [System.Windows.Forms.MessageBox]::Show('無法變更自己的角色', '提示', 'OK', 'Warning') | Out-Null
            return
        }
        $newRole = if ($u.role -eq 'Admin') { 'User' } else { 'Admin' }
        Update-CwtUser -Id $u.id -Set @{ role = $newRole } | Out-Null
        Reload-Users
    })
    $btnToggle.Add_Click({
        if ($gridU.SelectedRows.Count -eq 0) { return }
        $idx = $gridU.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:UserRows.Count) { return }
        $u = $script:UserRows[$idx]
        if ($u.id -eq $CurrentUser.id) {
            [System.Windows.Forms.MessageBox]::Show('無法停用自己', '提示', 'OK', 'Warning') | Out-Null
            return
        }
        Update-CwtUser -Id $u.id -Set @{ active = -not [bool]$u.active } | Out-Null
        Reload-Users
    })
    $btnReset.Add_Click({
        if ($gridU.SelectedRows.Count -eq 0) { return }
        $idx = $gridU.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:UserRows.Count) { return }
        $u = $script:UserRows[$idx]
        $pwd = Show-CwtInputBox -Title '重設密碼' -Prompt "為 $($u.username) 設定新密碼（至少 8 字元）:" -IsPassword
        if ([string]::IsNullOrEmpty($pwd)) { return }
        if ($pwd.Length -lt 8) {
            [System.Windows.Forms.MessageBox]::Show('密碼至少 8 字元', '提示', 'OK', 'Warning') | Out-Null
            return
        }
        Set-CwtUserPassword -Id $u.id -NewPassword $pwd | Out-Null
        Write-CwtLog -Action 'user.resetPassword' -ActorId $CurrentUser.id -Target $u.id -Level 'Security'
        [System.Windows.Forms.MessageBox]::Show('密碼已重設', '完成', 'OK', 'Information') | Out-Null
    })

    # --- 專案管理 ---
    $tabP = New-Object System.Windows.Forms.TabPage; $tabP.Text = '專案管理'
    $tabs.TabPages.Add($tabP)
    $gridP = New-Object System.Windows.Forms.DataGridView
    $gridP.Dock = 'Fill'; $gridP.ReadOnly = $true; $gridP.AllowUserToAddRows = $false
    $gridP.AllowUserToDeleteRows = $false; $gridP.SelectionMode = 'FullRowSelect'
    $gridP.MultiSelect = $false; $gridP.RowHeadersVisible = $false
    $gridP.AutoSizeColumnsMode = 'Fill'
    [void]$gridP.Columns.Add('code', '專案編碼')
    [void]$gridP.Columns.Add('name', '專案名稱')
    [void]$gridP.Columns.Add('owner', '發起人')
    [void]$gridP.Columns.Add('status', '狀態')
    [void]$gridP.Columns.Add('conf', '機密等級')
    [void]$gridP.Columns.Add('progress', '進度')
    [void]$gridP.Columns.Add('members', '成員數')
    [void]$gridP.Columns.Add('createdAt', '建立時間')
    $tabP.Controls.Add($gridP)
    $pnlP = New-Object System.Windows.Forms.Panel
    $pnlP.Dock = 'Top'; $pnlP.Height = 45
    $tabP.Controls.Add($pnlP)
    $btnOpenP = New-Object System.Windows.Forms.Button
    $btnOpenP.Text = '開啟專案'; $btnOpenP.Location = New-Object System.Drawing.Point(10, 8); $btnOpenP.Size = New-Object System.Drawing.Size(100, 30)
    $pnlP.Controls.Add($btnOpenP)
    $btnDelP = New-Object System.Windows.Forms.Button
    $btnDelP.Text = '刪除專案'; $btnDelP.Location = New-Object System.Drawing.Point(120, 8); $btnDelP.Size = New-Object System.Drawing.Size(100, 30)
    $pnlP.Controls.Add($btnDelP)
    $btnTransfer = New-Object System.Windows.Forms.Button
    $btnTransfer.Text = '變更發起人'; $btnTransfer.Location = New-Object System.Drawing.Point(230, 8); $btnTransfer.Size = New-Object System.Drawing.Size(110, 30)
    $pnlP.Controls.Add($btnTransfer)

    $script:ProjRows = @()
    function script:Reload-Projects {
        $gridP.Rows.Clear()
        $list = @(Get-CwtProjects | Sort-Object code)
        $script:ProjRows = $list
        foreach ($p in $list) {
            $o = Get-CwtUserById -Id $p.ownerId
            $oName = if ($null -ne $o) { $o.displayName } else { '(未知)' }
            $mems = @(Get-CwtProjectMembers -ProjectId $p.id)
            $tm = ''
            if (-not [string]::IsNullOrEmpty($p.createdAt)) {
                try { $tm = ([datetime]$p.createdAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { $tm = $p.createdAt }
            }
            $prog = Get-CwtProjectProgress -Project $p
            [void]$gridP.Rows.Add($p.code, $p.name, $oName, $p.status, $p.confidentiality, "$prog %", $mems.Count, $tm)
        }
    }
    $btnOpenP.Add_Click({
        if ($gridP.SelectedRows.Count -eq 0) { return }
        $idx = $gridP.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:ProjRows.Count) { return }
        Show-CwtProjectDetailForm -CurrentUser $CurrentUser -ProjectId $script:ProjRows[$idx].id
        Reload-Projects
    })
    $btnDelP.Add_Click({
        if ($gridP.SelectedRows.Count -eq 0) { return }
        $idx = $gridP.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:ProjRows.Count) { return }
        $p = $script:ProjRows[$idx]
        $r = [System.Windows.Forms.MessageBox]::Show("確定刪除專案「$($p.code) $($p.name)」？此操作不可復原。", '確認', 'YesNo', 'Warning')
        if ($r -eq 'Yes') {
            Remove-CwtProject -Id $p.id -ActorId $CurrentUser.id
            Reload-Projects
        }
    })
    $btnTransfer.Add_Click({
        if ($gridP.SelectedRows.Count -eq 0) { return }
        $idx = $gridP.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:ProjRows.Count) { return }
        $p = $script:ProjRows[$idx]
        $sel = Show-CwtUserPicker -Title '選擇新發起人'
        if ($null -eq $sel) { return }
        Update-CwtProject -Id $p.id -Set @{ ownerId = $sel.id } -ActorId $CurrentUser.id | Out-Null
        Set-CwtMembership -ProjectId $p.id -UserId $sel.id -Permission 'Owner' -GrantedBy $CurrentUser.id
        Reload-Projects
    })

    # --- Log 查詢 ---
    $tabL = New-Object System.Windows.Forms.TabPage; $tabL.Text = 'Log 查詢 / 匯出'
    $tabs.TabPages.Add($tabL)
    $pnlLF = New-Object System.Windows.Forms.Panel
    $pnlLF.Dock = 'Top'; $pnlLF.Height = 90
    $tabL.Controls.Add($pnlLF)
    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Text = '起始日期:'; $lblFrom.Location = New-Object System.Drawing.Point(10, 12); $lblFrom.Size = New-Object System.Drawing.Size(70, 25)
    $pnlLF.Controls.Add($lblFrom)
    $dtFrom = New-Object System.Windows.Forms.DateTimePicker
    $dtFrom.Format = 'Short'; $dtFrom.Value = (Get-Date).AddDays(-30)
    $dtFrom.Location = New-Object System.Drawing.Point(80, 9); $dtFrom.Size = New-Object System.Drawing.Size(120, 25)
    $pnlLF.Controls.Add($dtFrom)
    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = '結束日期:'; $lblTo.Location = New-Object System.Drawing.Point(215, 12); $lblTo.Size = New-Object System.Drawing.Size(70, 25)
    $pnlLF.Controls.Add($lblTo)
    $dtTo = New-Object System.Windows.Forms.DateTimePicker
    $dtTo.Format = 'Short'; $dtTo.Value = (Get-Date).AddDays(1)
    $dtTo.Location = New-Object System.Drawing.Point(285, 9); $dtTo.Size = New-Object System.Drawing.Size(120, 25)
    $pnlLF.Controls.Add($dtTo)
    $lblA = New-Object System.Windows.Forms.Label
    $lblA.Text = '動作 (模糊):'; $lblA.Location = New-Object System.Drawing.Point(420, 12); $lblA.Size = New-Object System.Drawing.Size(80, 25)
    $pnlLF.Controls.Add($lblA)
    $txtA = New-Object System.Windows.Forms.TextBox
    $txtA.Location = New-Object System.Drawing.Point(500, 9); $txtA.Size = New-Object System.Drawing.Size(150, 25)
    $pnlLF.Controls.Add($txtA)
    $lblLv = New-Object System.Windows.Forms.Label
    $lblLv.Text = '等級:'; $lblLv.Location = New-Object System.Drawing.Point(665, 12); $lblLv.Size = New-Object System.Drawing.Size(40, 25)
    $pnlLF.Controls.Add($lblLv)
    $cmbLv = New-Object System.Windows.Forms.ComboBox; $cmbLv.DropDownStyle = 'DropDownList'
    [void]$cmbLv.Items.AddRange(@('(全部)', 'Info', 'Warning', 'Error', 'Security'))
    $cmbLv.SelectedIndex = 0
    $cmbLv.Location = New-Object System.Drawing.Point(705, 9); $cmbLv.Size = New-Object System.Drawing.Size(100, 25)
    $pnlLF.Controls.Add($cmbLv)

    $btnQ = New-Object System.Windows.Forms.Button
    $btnQ.Text = '查詢'; $btnQ.Location = New-Object System.Drawing.Point(10, 50); $btnQ.Size = New-Object System.Drawing.Size(100, 30)
    $pnlLF.Controls.Add($btnQ)
    $btnExp = New-Object System.Windows.Forms.Button
    $btnExp.Text = '匯出 CSV (解密)'; $btnExp.Location = New-Object System.Drawing.Point(120, 50); $btnExp.Size = New-Object System.Drawing.Size(150, 30)
    $pnlLF.Controls.Add($btnExp)
    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Text = '立即備份'; $btnBackup.Location = New-Object System.Drawing.Point(280, 50); $btnBackup.Size = New-Object System.Drawing.Size(110, 30)
    $pnlLF.Controls.Add($btnBackup)
    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Text = '還原備份…'; $btnRestore.Location = New-Object System.Drawing.Point(400, 50); $btnRestore.Size = New-Object System.Drawing.Size(110, 30)
    $pnlLF.Controls.Add($btnRestore)

    $gridL = New-Object System.Windows.Forms.DataGridView
    $gridL.Dock = 'Fill'; $gridL.ReadOnly = $true; $gridL.AllowUserToAddRows = $false
    $gridL.AllowUserToDeleteRows = $false; $gridL.SelectionMode = 'FullRowSelect'
    $gridL.MultiSelect = $false; $gridL.RowHeadersVisible = $false
    $gridL.AutoSizeColumnsMode = 'Fill'
    [void]$gridL.Columns.Add('ts', '時間')
    [void]$gridL.Columns.Add('level', '等級')
    [void]$gridL.Columns.Add('action', '動作')
    [void]$gridL.Columns.Add('actorName', '操作者')
    [void]$gridL.Columns.Add('target', '對象')
    [void]$gridL.Columns.Add('detail', '說明')
    [void]$gridL.Columns.Add('host', '主機')
    [void]$gridL.Columns.Add('user', 'OS 帳號')
    $tabL.Controls.Add($gridL)

    function script:Do-LogQuery {
        $gridL.Rows.Clear()
        $lvl = $null
        if ($cmbLv.SelectedItem -ne '(全部)') { $lvl = $cmbLv.SelectedItem }
        $entries = Read-CwtLogs -From $dtFrom.Value.Date -To $dtTo.Value.Date.AddDays(1) `
            -Action $txtA.Text -Level $lvl
        $sorted = @($entries | Sort-Object ts -Descending)
        foreach ($e in $sorted) {
            $t = $e.ts
            try { $t = ([datetime]$e.ts).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') } catch { }
            [void]$gridL.Rows.Add($t, $e.level, $e.action, $e.actorName, $e.target, $e.detail, $e.host, $e.user)
        }
    }
    $btnQ.Add_Click({ Do-LogQuery })
    $btnExp.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'CSV (*.csv)|*.csv'
        $sfd.FileName = "logs_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
        if ($sfd.ShowDialog() -eq 'OK') {
            $lvl = $null
            if ($cmbLv.SelectedItem -ne '(全部)') { $lvl = $cmbLv.SelectedItem }
            Export-CwtLogsCsv -Path $sfd.FileName -From $dtFrom.Value.Date -To $dtTo.Value.Date.AddDays(1) -Action $txtA.Text -Level $lvl
            Write-CwtLog -Action 'log.export' -ActorId $CurrentUser.id -ActorName $CurrentUser.username -Target $sfd.FileName -Level 'Security'
            [System.Windows.Forms.MessageBox]::Show("已匯出至: $($sfd.FileName)", '完成', 'OK', 'Information') | Out-Null
        }
    })
    $btnBackup.Add_Click({
        try {
            $z = Invoke-CwtBackup -Tag 'manual'
            Write-CwtLog -Action 'backup.manual' -ActorId $CurrentUser.id -Target $z
            [System.Windows.Forms.MessageBox]::Show("備份已建立: $z", '完成', 'OK', 'Information') | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '錯誤', 'OK', 'Error') | Out-Null
        }
    })
    $btnRestore.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'ZIP (*.zip)|*.zip'
        $ofd.InitialDirectory = (Get-CwtPaths).Backups
        if ($ofd.ShowDialog() -eq 'OK') {
            $r = [System.Windows.Forms.MessageBox]::Show('還原前會自動備份目前資料，確定執行嗎？', '確認', 'YesNo', 'Warning')
            if ($r -eq 'Yes') {
                try {
                    $pre = Restore-CwtBackup -ZipPath $ofd.FileName
                    Write-CwtLog -Action 'backup.restore' -ActorId $CurrentUser.id -Target $ofd.FileName -Level 'Security'
                    [System.Windows.Forms.MessageBox]::Show("還原完成。還原前的備份: $pre`r`n建議重新啟動程式。", '完成', 'OK', 'Information') | Out-Null
                } catch {
                    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '錯誤', 'OK', 'Error') | Out-Null
                }
            }
        }
    })

    $tabs.Add_SelectedIndexChanged({
        switch ($tabs.SelectedTab.Text) {
            '人員管理' { Reload-Users }
            '專案管理' { Reload-Projects }
        }
    })

    $f.Add_Shown({ Reload-Users; Reload-Projects })
    $null = $f.ShowDialog()
}

function Show-CwtInputBox {
    param(
        [string]$Title = '輸入',
        [string]$Prompt = '請輸入:',
        [switch]$IsPassword
    )
    Add-Type -AssemblyName System.Windows.Forms

    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title; $f.Size = New-Object System.Drawing.Size(420, 180); $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Prompt; $lbl.Location = New-Object System.Drawing.Point(15, 15); $lbl.Size = New-Object System.Drawing.Size(380, 30)
    $f.Controls.Add($lbl)
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(15, 50); $txt.Size = New-Object System.Drawing.Size(380, 25)
    if ($IsPassword) { $txt.UseSystemPasswordChar = $true }
    $f.Controls.Add($txt)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = '確定'; $ok.Location = New-Object System.Drawing.Point(195, 95); $ok.Size = New-Object System.Drawing.Size(95, 30)
    $f.Controls.Add($ok)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = '取消'; $cancel.Location = New-Object System.Drawing.Point(300, 95); $cancel.Size = New-Object System.Drawing.Size(95, 30)
    $cancel.DialogResult = 'Cancel'
    $f.Controls.Add($cancel)

    $script:InputResult = $null
    $ok.Add_Click({ $script:InputResult = $txt.Text; $f.DialogResult = 'OK'; $f.Close() })
    $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $null = $f.ShowDialog()
    return $script:InputResult
}

function Show-CwtUserPicker {
    param([string]$Title = '選擇使用者')
    Add-Type -AssemblyName System.Windows.Forms
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title; $f.Size = New-Object System.Drawing.Size(420, 440); $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false

    $lb = New-Object System.Windows.Forms.ListBox
    $lb.Location = New-Object System.Drawing.Point(10, 10); $lb.Size = New-Object System.Drawing.Size(385, 330)
    $f.Controls.Add($lb)
    $map = @{}
    foreach ($u in @(Get-CwtUsers | Where-Object { $_.active } | Sort-Object username)) {
        $label = "$($u.username) - $($u.displayName)"
        [void]$lb.Items.Add($label)
        $map[$label] = $u
    }
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = '選擇'; $ok.Location = New-Object System.Drawing.Point(195, 355); $ok.Size = New-Object System.Drawing.Size(95, 30)
    $f.Controls.Add($ok)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = '取消'; $cancel.Location = New-Object System.Drawing.Point(300, 355); $cancel.Size = New-Object System.Drawing.Size(95, 30)
    $cancel.DialogResult = 'Cancel'
    $f.Controls.Add($cancel)

    $script:PickResult = $null
    $ok.Add_Click({
        if ($null -ne $lb.SelectedItem -and $map.ContainsKey($lb.SelectedItem)) {
            $script:PickResult = $map[$lb.SelectedItem]
            $f.DialogResult = 'OK'; $f.Close()
        }
    })
    $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $null = $f.ShowDialog()
    return $script:PickResult
}
