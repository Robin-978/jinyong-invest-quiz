# MainForm.ps1 - 專案列表主頁
function Show-CwtMainForm {
    param([Parameter(Mandatory)]$User)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "協同工作紀錄系統 - 主頁 [$($User.displayName) / $($User.role)]"
    $form.Size = New-Object System.Drawing.Size(1000, 640)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(900, 560)

    # 上方工具列
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = 'Top'
    $pnlTop.Height = 50
    $form.Controls.Add($pnlTop)

    $btnNewProject = New-Object System.Windows.Forms.Button
    $btnNewProject.Text = '＋ 申請新專案'
    $btnNewProject.Location = New-Object System.Drawing.Point(10, 10); $btnNewProject.Size = New-Object System.Drawing.Size(120, 30)
    $pnlTop.Controls.Add($btnNewProject)

    $btnJoin = New-Object System.Windows.Forms.Button
    $btnJoin.Text = '申請參與專案'
    $btnJoin.Location = New-Object System.Drawing.Point(140, 10); $btnJoin.Size = New-Object System.Drawing.Size(120, 30)
    $pnlTop.Controls.Add($btnJoin)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Text = '開啟專案'
    $btnOpen.Location = New-Object System.Drawing.Point(270, 10); $btnOpen.Size = New-Object System.Drawing.Size(100, 30)
    $pnlTop.Controls.Add($btnOpen)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = '重新整理'
    $btnRefresh.Location = New-Object System.Drawing.Point(380, 10); $btnRefresh.Size = New-Object System.Drawing.Size(90, 30)
    $pnlTop.Controls.Add($btnRefresh)

    $btnAdmin = New-Object System.Windows.Forms.Button
    $btnAdmin.Text = '管理頁面'
    $btnAdmin.Location = New-Object System.Drawing.Point(480, 10); $btnAdmin.Size = New-Object System.Drawing.Size(90, 30)
    $btnAdmin.Visible = ($User.role -eq 'Admin')
    $pnlTop.Controls.Add($btnAdmin)

    $btnChangePwd = New-Object System.Windows.Forms.Button
    $btnChangePwd.Text = '修改密碼'
    $btnChangePwd.Location = New-Object System.Drawing.Point(580, 10); $btnChangePwd.Size = New-Object System.Drawing.Size(90, 30)
    $pnlTop.Controls.Add($btnChangePwd)

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = '檢視:'; $lblFilter.Location = New-Object System.Drawing.Point(700, 15); $lblFilter.Size = New-Object System.Drawing.Size(45, 25)
    $pnlTop.Controls.Add($lblFilter)
    $cmbFilter = New-Object System.Windows.Forms.ComboBox
    $cmbFilter.DropDownStyle = 'DropDownList'
    [void]$cmbFilter.Items.AddRange(@('我的專案', '全部可見專案'))
    $cmbFilter.SelectedIndex = 0
    $cmbFilter.Location = New-Object System.Drawing.Point(750, 12); $cmbFilter.Size = New-Object System.Drawing.Size(140, 25)
    $pnlTop.Controls.Add($cmbFilter)

    # 列表
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.RowHeadersVisible = $false
    $form.Controls.Add($grid)

    # 狀態列
    $status = New-Object System.Windows.Forms.StatusStrip
    $statusLbl = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLbl.Text = '就緒'
    [void]$status.Items.Add($statusLbl)
    $form.Controls.Add($status)

    $script:CurrentRows = @()

    function script:Refresh-Grid {
        $grid.Columns.Clear()
        $grid.Rows.Clear()
        $colCode = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colCode.HeaderText = '專案編碼'; $colCode.Name = 'code'; $colCode.FillWeight = 18
        [void]$grid.Columns.Add($colCode)
        $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colName.HeaderText = '專案名稱'; $colName.Name = 'name'; $colName.FillWeight = 24
        [void]$grid.Columns.Add($colName)
        $colOwner = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colOwner.HeaderText = '發起人'; $colOwner.Name = 'owner'; $colOwner.FillWeight = 12
        [void]$grid.Columns.Add($colOwner)
        $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colStatus.HeaderText = '狀態'; $colStatus.Name = 'status'; $colStatus.FillWeight = 10
        [void]$grid.Columns.Add($colStatus)
        $colConf = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colConf.HeaderText = '機密等級'; $colConf.Name = 'conf'; $colConf.FillWeight = 10
        [void]$grid.Columns.Add($colConf)
        $colProg = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colProg.HeaderText = '進度'; $colProg.Name = 'progress'; $colProg.FillWeight = 8
        [void]$grid.Columns.Add($colProg)
        $colPerm = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colPerm.HeaderText = '我的權限'; $colPerm.Name = 'perm'; $colPerm.FillWeight = 10
        [void]$grid.Columns.Add($colPerm)
        $colDue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colDue.HeaderText = '期限'; $colDue.Name = 'due'; $colDue.FillWeight = 12
        [void]$grid.Columns.Add($colDue)

        $all = @(Get-CwtProjects)
        $mode = $cmbFilter.SelectedItem
        $list = @()
        if ($mode -eq '我的專案') {
            $list = @(Get-CwtUserProjects -UserId $User.id)
            if ($User.role -eq 'Admin') { $list = $all }
        } else {
            if ($User.role -eq 'Admin') {
                $list = $all
            } else {
                foreach ($p in $all) {
                    if ($p.confidentiality -eq 'Public') { $list += $p; continue }
                    if ($null -ne (Get-CwtMembership -ProjectId $p.id -UserId $User.id)) { $list += $p }
                }
            }
        }
        $script:CurrentRows = $list
        foreach ($p in $list) {
            $owner = Get-CwtUserById -Id $p.ownerId
            $ownerName = if ($null -ne $owner) { $owner.displayName } else { '(未知)' }
            $mem = Get-CwtMembership -ProjectId $p.id -UserId $User.id
            $perm = if ($null -ne $mem) { $mem.permission } else { '' }
            if ($User.role -eq 'Admin' -and [string]::IsNullOrEmpty($perm)) { $perm = '(Admin)' }
            $prog = Get-CwtProjectProgress -Project $p
            $due = ''
            if ($p.PSObject.Properties.Name -contains 'deadline' -and -not [string]::IsNullOrEmpty($p.deadline)) {
                try { $due = ([datetime]$p.deadline).ToLocalTime().ToString('yyyy-MM-dd') } catch { $due = $p.deadline }
            }
            [void]$grid.Rows.Add($p.code, $p.name, $ownerName, $p.status, $p.confidentiality, "$prog %", $perm, $due)
        }
        $statusLbl.Text = "顯示 $($list.Count) 個專案"
    }

    $btnRefresh.Add_Click({ Refresh-Grid })
    $cmbFilter.Add_SelectedIndexChanged({ Refresh-Grid })

    $btnNewProject.Add_Click({
        $r = Show-CwtProjectEditForm -CurrentUser $User -Project $null
        if ($r) { Refresh-Grid }
    })

    $btnJoin.Add_Click({
        Show-CwtJoinProjectForm -CurrentUser $User
        Refresh-Grid
    })

    $openSelected = {
        if ($grid.SelectedRows.Count -eq 0) { return }
        $idx = $grid.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:CurrentRows.Count) { return }
        $proj = $script:CurrentRows[$idx]
        if (-not (Test-CwtCanView -ProjectId $proj.id -User $User)) {
            [System.Windows.Forms.MessageBox]::Show('您沒有檢視此專案的權限', '權限不足', 'OK', 'Warning') | Out-Null
            return
        }
        Show-CwtProjectDetailForm -CurrentUser $User -ProjectId $proj.id
        Refresh-Grid
    }
    $btnOpen.Add_Click($openSelected)
    $grid.Add_CellDoubleClick($openSelected)

    $btnAdmin.Add_Click({
        Show-CwtAdminForm -CurrentUser $User
        Refresh-Grid
    })

    $btnChangePwd.Add_Click({
        Show-CwtChangePasswordForm -CurrentUser $User
    })

    $form.Add_Shown({ Refresh-Grid })
    $form.Add_FormClosed({
        Write-CwtLog -Action 'logout' -ActorId $User.id -ActorName $User.username -Level 'Security'
    })
    $null = $form.ShowDialog()
}

function Show-CwtChangePasswordForm {
    param([Parameter(Mandatory)]$CurrentUser)
    Add-Type -AssemblyName System.Windows.Forms

    $f = New-Object System.Windows.Forms.Form
    $f.Text = '修改密碼'; $f.Size = New-Object System.Drawing.Size(380, 280); $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false

    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = '舊密碼:'; $lbl1.Location = New-Object System.Drawing.Point(20, 25); $lbl1.Size = New-Object System.Drawing.Size(80, 25)
    $f.Controls.Add($lbl1)
    $t1 = New-Object System.Windows.Forms.TextBox; $t1.UseSystemPasswordChar = $true
    $t1.Location = New-Object System.Drawing.Point(110, 22); $t1.Size = New-Object System.Drawing.Size(230, 25)
    $f.Controls.Add($t1)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text = '新密碼:'; $lbl2.Location = New-Object System.Drawing.Point(20, 60); $lbl2.Size = New-Object System.Drawing.Size(80, 25)
    $f.Controls.Add($lbl2)
    $t2 = New-Object System.Windows.Forms.TextBox; $t2.UseSystemPasswordChar = $true
    $t2.Location = New-Object System.Drawing.Point(110, 57); $t2.Size = New-Object System.Drawing.Size(230, 25)
    $f.Controls.Add($t2)

    $lbl3 = New-Object System.Windows.Forms.Label
    $lbl3.Text = '確認新密碼:'; $lbl3.Location = New-Object System.Drawing.Point(20, 95); $lbl3.Size = New-Object System.Drawing.Size(85, 25)
    $f.Controls.Add($lbl3)
    $t3 = New-Object System.Windows.Forms.TextBox; $t3.UseSystemPasswordChar = $true
    $t3.Location = New-Object System.Drawing.Point(110, 92); $t3.Size = New-Object System.Drawing.Size(230, 25)
    $f.Controls.Add($t3)

    $msg = New-Object System.Windows.Forms.Label
    $msg.Location = New-Object System.Drawing.Point(20, 130); $msg.Size = New-Object System.Drawing.Size(320, 50)
    $msg.ForeColor = [System.Drawing.Color]::Firebrick
    $f.Controls.Add($msg)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = '確定'; $ok.Location = New-Object System.Drawing.Point(120, 190); $ok.Size = New-Object System.Drawing.Size(100, 32)
    $f.Controls.Add($ok)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = '取消'; $cancel.Location = New-Object System.Drawing.Point(235, 190); $cancel.Size = New-Object System.Drawing.Size(100, 32)
    $cancel.DialogResult = 'Cancel'
    $f.Controls.Add($cancel)

    $ok.Add_Click({
        $msg.Text = ''
        if ([string]::IsNullOrEmpty($t1.Text)) { $msg.Text = '請輸入舊密碼'; return }
        $verify = Test-CwtLogin -Username $CurrentUser.username -Password $t1.Text
        if ($null -eq $verify) { $msg.Text = '舊密碼錯誤'; return }
        if ($t2.Text.Length -lt 8) { $msg.Text = '新密碼至少 8 個字元'; return }
        if ($t2.Text -ne $t3.Text) { $msg.Text = '兩次新密碼不一致'; return }
        try {
            Set-CwtUserPassword -Id $CurrentUser.id -NewPassword $t2.Text | Out-Null
            Write-CwtLog -Action 'user.changePassword' -ActorId $CurrentUser.id -ActorName $CurrentUser.username -Level 'Security'
            [System.Windows.Forms.MessageBox]::Show('密碼已更新', '成功', 'OK', 'Information') | Out-Null
            $f.DialogResult = 'OK'; $f.Close()
        } catch { $msg.Text = $_.Exception.Message }
    })

    $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $null = $f.ShowDialog()
}

function Show-CwtJoinProjectForm {
    param([Parameter(Mandatory)]$CurrentUser)
    Add-Type -AssemblyName System.Windows.Forms

    $f = New-Object System.Windows.Forms.Form
    $f.Text = '申請參與專案'; $f.Size = New-Object System.Drawing.Size(720, 480); $f.StartPosition = 'CenterParent'

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(10, 10); $grid.Size = New-Object System.Drawing.Size(685, 340)
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $false; $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    [void]$grid.Columns.Add('code', '專案編碼')
    [void]$grid.Columns.Add('name', '專案名稱')
    [void]$grid.Columns.Add('owner', '發起人')
    [void]$grid.Columns.Add('conf', '機密等級')
    $f.Controls.Add($grid)

    $lblPerm = New-Object System.Windows.Forms.Label
    $lblPerm.Text = '申請權限:'; $lblPerm.Location = New-Object System.Drawing.Point(10, 365); $lblPerm.Size = New-Object System.Drawing.Size(70, 25)
    $f.Controls.Add($lblPerm)
    $cmbPerm = New-Object System.Windows.Forms.ComboBox
    $cmbPerm.DropDownStyle = 'DropDownList'
    [void]$cmbPerm.Items.AddRange(@('Viewer (唯讀)', 'Editor (協作編輯)'))
    $cmbPerm.SelectedIndex = 0
    $cmbPerm.Location = New-Object System.Drawing.Point(85, 362); $cmbPerm.Size = New-Object System.Drawing.Size(160, 25)
    $f.Controls.Add($cmbPerm)

    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text = '訊息:'; $lblMsg.Location = New-Object System.Drawing.Point(260, 365); $lblMsg.Size = New-Object System.Drawing.Size(40, 25)
    $f.Controls.Add($lblMsg)
    $txtMsg = New-Object System.Windows.Forms.TextBox
    $txtMsg.Location = New-Object System.Drawing.Point(300, 362); $txtMsg.Size = New-Object System.Drawing.Size(395, 25)
    $f.Controls.Add($txtMsg)

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = '送出申請'; $btnApply.Location = New-Object System.Drawing.Point(10, 400); $btnApply.Size = New-Object System.Drawing.Size(110, 32)
    $f.Controls.Add($btnApply)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = '關閉'; $btnClose.Location = New-Object System.Drawing.Point(595, 400); $btnClose.Size = New-Object System.Drawing.Size(100, 32)
    $btnClose.DialogResult = 'Cancel'
    $f.Controls.Add($btnClose)

    $script:JoinRows = @()
    function script:Reload-Join {
        $grid.Rows.Clear()
        $all = @(Get-CwtProjects)
        $list = @()
        foreach ($p in $all) {
            if ($p.ownerId -eq $CurrentUser.id) { continue }
            $m = Get-CwtMembership -ProjectId $p.id -UserId $CurrentUser.id
            if ($null -ne $m) { continue }
            if ($p.confidentiality -eq 'TopSecret') { continue }
            $list += $p
        }
        $script:JoinRows = $list
        foreach ($p in $list) {
            $o = Get-CwtUserById -Id $p.ownerId
            $oName = if ($null -ne $o) { $o.displayName } else { '(未知)' }
            [void]$grid.Rows.Add($p.code, $p.name, $oName, $p.confidentiality)
        }
    }

    $btnApply.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('請選擇要申請的專案', '提示', 'OK', 'Information') | Out-Null
            return
        }
        $idx = $grid.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:JoinRows.Count) { return }
        $proj = $script:JoinRows[$idx]
        $perm = if ($cmbPerm.SelectedIndex -eq 1) { 'Editor' } else { 'Viewer' }
        try {
            New-CwtJoinRequest -ProjectId $proj.id -UserId $CurrentUser.id -RequestedPermission $perm -Message $txtMsg.Text | Out-Null
            [System.Windows.Forms.MessageBox]::Show('已送出申請，等待專案發起人核可', '完成', 'OK', 'Information') | Out-Null
            Reload-Join
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '錯誤', 'OK', 'Error') | Out-Null
        }
    })

    $f.Add_Shown({ Reload-Join })
    $f.CancelButton = $btnClose
    $null = $f.ShowDialog()
}
