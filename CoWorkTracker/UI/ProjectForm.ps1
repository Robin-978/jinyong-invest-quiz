# ProjectForm.ps1 - 專案編輯/詳細頁
function Show-CwtProjectEditForm {
    param(
        [Parameter(Mandatory)]$CurrentUser,
        $Project = $null
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $isNew = ($null -eq $Project)
    if (-not $isNew) {
        if (-not (Test-CwtCanManage -ProjectId $Project.id -User $CurrentUser)) {
            [System.Windows.Forms.MessageBox]::Show('僅發起人或管理員可編輯此專案', '權限不足', 'OK', 'Warning') | Out-Null
            return $null
        }
    }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = if ($isNew) { '申請新專案' } else { "編輯專案 - $($Project.code)" }
    $f.Size = New-Object System.Drawing.Size(560, 520); $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false

    $y = 20
    $lblCode = New-Object System.Windows.Forms.Label
    $lblCode.Text = '專案編碼:'; $lblCode.Location = New-Object System.Drawing.Point(20, $y); $lblCode.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblCode)
    $txtCode = New-Object System.Windows.Forms.TextBox
    $txtCode.Location = New-Object System.Drawing.Point(120, ($y - 3)); $txtCode.Size = New-Object System.Drawing.Size(200, 25)
    if ($isNew) { $txtCode.Text = (New-CwtProjectCode) }
    else { $txtCode.Text = $Project.code; $txtCode.ReadOnly = $true }
    $f.Controls.Add($txtCode)

    $y += 35
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = '專案名稱 *:'; $lblName.Location = New-Object System.Drawing.Point(20, $y); $lblName.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblName)
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(120, ($y - 3)); $txtName.Size = New-Object System.Drawing.Size(400, 25)
    if (-not $isNew) { $txtName.Text = $Project.name }
    $f.Controls.Add($txtName)

    $y += 35
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = '專案說明:'; $lblDesc.Location = New-Object System.Drawing.Point(20, $y); $lblDesc.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblDesc)
    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Multiline = $true; $txtDesc.ScrollBars = 'Vertical'
    $txtDesc.Location = New-Object System.Drawing.Point(120, ($y - 3)); $txtDesc.Size = New-Object System.Drawing.Size(400, 100)
    if (-not $isNew) { $txtDesc.Text = $Project.description }
    $f.Controls.Add($txtDesc)

    $y += 110
    $lblConf = New-Object System.Windows.Forms.Label
    $lblConf.Text = '機密等級:'; $lblConf.Location = New-Object System.Drawing.Point(20, $y); $lblConf.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblConf)
    $cmbConf = New-Object System.Windows.Forms.ComboBox; $cmbConf.DropDownStyle = 'DropDownList'
    [void]$cmbConf.Items.AddRange(@('Public', 'Internal', 'Confidential', 'TopSecret'))
    $cmbConf.Location = New-Object System.Drawing.Point(120, ($y - 3)); $cmbConf.Size = New-Object System.Drawing.Size(150, 25)
    if ($isNew) { $cmbConf.SelectedItem = 'Internal' } else { $cmbConf.SelectedItem = $Project.confidentiality }
    $f.Controls.Add($cmbConf)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = '專案狀態:'; $lblStatus.Location = New-Object System.Drawing.Point(290, $y); $lblStatus.Size = New-Object System.Drawing.Size(70, 25)
    $f.Controls.Add($lblStatus)
    $cmbStatus = New-Object System.Windows.Forms.ComboBox; $cmbStatus.DropDownStyle = 'DropDownList'
    [void]$cmbStatus.Items.AddRange(@('Planning','InProgress','OnHold','Completed','Cancelled'))
    $cmbStatus.Location = New-Object System.Drawing.Point(365, ($y - 3)); $cmbStatus.Size = New-Object System.Drawing.Size(155, 25)
    if ($isNew) { $cmbStatus.SelectedItem = 'Planning' } else { $cmbStatus.SelectedItem = $Project.status }
    $f.Controls.Add($cmbStatus)

    $y += 35
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = '開始日期:'; $lblStart.Location = New-Object System.Drawing.Point(20, $y); $lblStart.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblStart)
    $dtStart = New-Object System.Windows.Forms.DateTimePicker
    $dtStart.Format = 'Short'; $dtStart.Location = New-Object System.Drawing.Point(120, ($y - 3)); $dtStart.Size = New-Object System.Drawing.Size(150, 25)
    if (-not $isNew -and -not [string]::IsNullOrEmpty($Project.startDate)) {
        try { $dtStart.Value = ([datetime]$Project.startDate).ToLocalTime() } catch { }
    }
    $f.Controls.Add($dtStart)

    $lblDue = New-Object System.Windows.Forms.Label
    $lblDue.Text = '截止日期:'; $lblDue.Location = New-Object System.Drawing.Point(290, $y); $lblDue.Size = New-Object System.Drawing.Size(70, 25)
    $f.Controls.Add($lblDue)
    $dtDue = New-Object System.Windows.Forms.DateTimePicker
    $dtDue.Format = 'Short'; $dtDue.ShowCheckBox = $true; $dtDue.Checked = $false
    $dtDue.Location = New-Object System.Drawing.Point(365, ($y - 3)); $dtDue.Size = New-Object System.Drawing.Size(155, 25)
    if (-not $isNew -and -not [string]::IsNullOrEmpty($Project.deadline)) {
        try { $dtDue.Value = ([datetime]$Project.deadline).ToLocalTime(); $dtDue.Checked = $true } catch { }
    }
    $f.Controls.Add($dtDue)

    $y += 45
    $lblTip = New-Object System.Windows.Forms.Label
    $lblTip.Text = '建立後可在「成員」分頁設定參與人員與權限。'
    $lblTip.Location = New-Object System.Drawing.Point(20, $y); $lblTip.Size = New-Object System.Drawing.Size(500, 25)
    $lblTip.ForeColor = [System.Drawing.Color]::Gray
    $f.Controls.Add($lblTip)

    $msg = New-Object System.Windows.Forms.Label
    $msg.Location = New-Object System.Drawing.Point(20, ($y + 30)); $msg.Size = New-Object System.Drawing.Size(500, 30)
    $msg.ForeColor = [System.Drawing.Color]::Firebrick
    $f.Controls.Add($msg)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = if ($isNew) { '建立專案' } else { '儲存變更' }
    $btnOk.Location = New-Object System.Drawing.Point(285, 435); $btnOk.Size = New-Object System.Drawing.Size(110, 32)
    $f.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '取消'; $btnCancel.Location = New-Object System.Drawing.Point(410, 435); $btnCancel.Size = New-Object System.Drawing.Size(110, 32)
    $btnCancel.DialogResult = 'Cancel'
    $f.Controls.Add($btnCancel)

    $script:EditResult = $null
    $btnOk.Add_Click({
        $msg.Text = ''
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) { $msg.Text = '請輸入專案名稱'; return }
        if ([string]::IsNullOrWhiteSpace($txtCode.Text)) { $msg.Text = '請輸入專案編碼'; return }
        $deadline = $null
        if ($dtDue.Checked) { $deadline = $dtDue.Value }
        try {
            if ($isNew) {
                $p = New-CwtProject -Name $txtName.Text -OwnerId $CurrentUser.id -Code $txtCode.Text `
                    -Description $txtDesc.Text -Confidentiality $cmbConf.SelectedItem `
                    -StartDate $dtStart.Value -Deadline $deadline -Status $cmbStatus.SelectedItem
                $script:EditResult = $p
            } else {
                $set = @{
                    name = $txtName.Text
                    description = $txtDesc.Text
                    confidentiality = $cmbConf.SelectedItem
                    status = $cmbStatus.SelectedItem
                    startDate = $dtStart.Value.ToUniversalTime().ToString('o')
                    deadline = if ($dtDue.Checked) { $dtDue.Value.ToUniversalTime().ToString('o') } else { '' }
                }
                $p = Update-CwtProject -Id $Project.id -Set $set -ActorId $CurrentUser.id
                $script:EditResult = $p
            }
            $f.DialogResult = 'OK'; $f.Close()
        } catch { $msg.Text = $_.Exception.Message }
    })

    $f.AcceptButton = $btnOk; $f.CancelButton = $btnCancel
    $null = $f.ShowDialog()
    return $script:EditResult
}

function Show-CwtProjectDetailForm {
    param(
        [Parameter(Mandatory)]$CurrentUser,
        [Parameter(Mandatory)][string]$ProjectId
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $proj = Get-CwtProjectById -Id $ProjectId
    if ($null -eq $proj) {
        [System.Windows.Forms.MessageBox]::Show('找不到專案', '錯誤', 'OK', 'Error') | Out-Null
        return
    }
    Write-CwtLog -Action 'project.open' -ActorId $CurrentUser.id -ActorName $CurrentUser.username -Target $ProjectId

    $canEdit = Test-CwtCanEdit -ProjectId $ProjectId -User $CurrentUser
    $canManage = Test-CwtCanManage -ProjectId $ProjectId -User $CurrentUser

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "專案: $($proj.code) - $($proj.name)"
    $f.Size = New-Object System.Drawing.Size(1050, 700); $f.StartPosition = 'CenterParent'
    $f.MinimumSize = New-Object System.Drawing.Size(900, 600)

    # 上方：基本資訊與進度條
    $pnlInfo = New-Object System.Windows.Forms.Panel
    $pnlInfo.Dock = 'Top'; $pnlInfo.Height = 110
    $f.Controls.Add($pnlInfo)

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 12, [System.Drawing.FontStyle]::Bold)
    $lblHeader.Location = New-Object System.Drawing.Point(10, 8); $lblHeader.Size = New-Object System.Drawing.Size(1000, 25)
    $pnlInfo.Controls.Add($lblHeader)

    $lblMeta = New-Object System.Windows.Forms.Label
    $lblMeta.Location = New-Object System.Drawing.Point(10, 35); $lblMeta.Size = New-Object System.Drawing.Size(1000, 40)
    $pnlInfo.Controls.Add($lblMeta)

    $progBar = New-Object System.Windows.Forms.ProgressBar
    $progBar.Location = New-Object System.Drawing.Point(10, 80); $progBar.Size = New-Object System.Drawing.Size(800, 22)
    $progBar.Minimum = 0; $progBar.Maximum = 100
    $pnlInfo.Controls.Add($progBar)
    $lblProgVal = New-Object System.Windows.Forms.Label
    $lblProgVal.Location = New-Object System.Drawing.Point(820, 80); $lblProgVal.Size = New-Object System.Drawing.Size(180, 22)
    $pnlInfo.Controls.Add($lblProgVal)

    # Tab control
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'
    $f.Controls.Add($tabs)
    $tabs.BringToFront()

    # Tab: 階段
    $tabPhases = New-Object System.Windows.Forms.TabPage; $tabPhases.Text = '階段 / KPI / 進度'
    $tabs.TabPages.Add($tabPhases)

    $gridPhases = New-Object System.Windows.Forms.DataGridView
    $gridPhases.Dock = 'Top'; $gridPhases.Height = 260
    $gridPhases.ReadOnly = $true; $gridPhases.AllowUserToAddRows = $false; $gridPhases.AllowUserToDeleteRows = $false
    $gridPhases.SelectionMode = 'FullRowSelect'; $gridPhases.MultiSelect = $false; $gridPhases.RowHeadersVisible = $false
    $gridPhases.AutoSizeColumnsMode = 'Fill'
    [void]$gridPhases.Columns.Add('name', '階段名稱')
    [void]$gridPhases.Columns.Add('kpi', 'KPI')
    [void]$gridPhases.Columns.Add('leader', '負責人')
    [void]$gridPhases.Columns.Add('collab', '協同人員')
    [void]$gridPhases.Columns.Add('status', '狀態')
    [void]$gridPhases.Columns.Add('progress', '進度')
    [void]$gridPhases.Columns.Add('due', '截止日')
    $tabPhases.Controls.Add($gridPhases)

    $pnlPhaseBtn = New-Object System.Windows.Forms.Panel
    $pnlPhaseBtn.Dock = 'Top'; $pnlPhaseBtn.Height = 45
    $tabPhases.Controls.Add($pnlPhaseBtn)

    $btnAddPhase = New-Object System.Windows.Forms.Button
    $btnAddPhase.Text = '＋ 新增階段'; $btnAddPhase.Location = New-Object System.Drawing.Point(10, 8); $btnAddPhase.Size = New-Object System.Drawing.Size(110, 30)
    $btnAddPhase.Enabled = $canEdit
    $pnlPhaseBtn.Controls.Add($btnAddPhase)

    $btnEditPhase = New-Object System.Windows.Forms.Button
    $btnEditPhase.Text = '編輯階段'; $btnEditPhase.Location = New-Object System.Drawing.Point(130, 8); $btnEditPhase.Size = New-Object System.Drawing.Size(100, 30)
    $btnEditPhase.Enabled = $canEdit
    $pnlPhaseBtn.Controls.Add($btnEditPhase)

    $btnDelPhase = New-Object System.Windows.Forms.Button
    $btnDelPhase.Text = '刪除階段'; $btnDelPhase.Location = New-Object System.Drawing.Point(240, 8); $btnDelPhase.Size = New-Object System.Drawing.Size(100, 30)
    $btnDelPhase.Enabled = $canManage
    $pnlPhaseBtn.Controls.Add($btnDelPhase)

    $btnUploadFile = New-Object System.Windows.Forms.Button
    $btnUploadFile.Text = '上傳附件'; $btnUploadFile.Location = New-Object System.Drawing.Point(350, 8); $btnUploadFile.Size = New-Object System.Drawing.Size(100, 30)
    $btnUploadFile.Enabled = $canEdit
    $pnlPhaseBtn.Controls.Add($btnUploadFile)

    $btnOpenFolder = New-Object System.Windows.Forms.Button
    $btnOpenFolder.Text = '開啟附件資料夾'; $btnOpenFolder.Location = New-Object System.Drawing.Point(460, 8); $btnOpenFolder.Size = New-Object System.Drawing.Size(120, 30)
    $pnlPhaseBtn.Controls.Add($btnOpenFolder)

    # 進度視覺化（簡易 bar chart）
    $picChart = New-Object System.Windows.Forms.PictureBox
    $picChart.Dock = 'Fill'; $picChart.BackColor = [System.Drawing.Color]::White
    $picChart.SizeMode = 'Normal'
    $tabPhases.Controls.Add($picChart)
    $picChart.BringToFront()

    # Tab: 成員
    $tabMembers = New-Object System.Windows.Forms.TabPage; $tabMembers.Text = '成員與權限'
    $tabs.TabPages.Add($tabMembers)

    $gridMembers = New-Object System.Windows.Forms.DataGridView
    $gridMembers.Dock = 'Fill'
    $gridMembers.ReadOnly = $true; $gridMembers.AllowUserToAddRows = $false; $gridMembers.AllowUserToDeleteRows = $false
    $gridMembers.SelectionMode = 'FullRowSelect'; $gridMembers.MultiSelect = $false; $gridMembers.RowHeadersVisible = $false
    $gridMembers.AutoSizeColumnsMode = 'Fill'
    [void]$gridMembers.Columns.Add('user', '帳號')
    [void]$gridMembers.Columns.Add('name', '顯示名稱')
    [void]$gridMembers.Columns.Add('perm', '權限')
    [void]$gridMembers.Columns.Add('granted', '授權時間')
    $tabMembers.Controls.Add($gridMembers)

    $pnlMemBtn = New-Object System.Windows.Forms.Panel
    $pnlMemBtn.Dock = 'Top'; $pnlMemBtn.Height = 45
    $tabMembers.Controls.Add($pnlMemBtn)

    $btnAddMem = New-Object System.Windows.Forms.Button
    $btnAddMem.Text = '＋ 加入成員'; $btnAddMem.Location = New-Object System.Drawing.Point(10, 8); $btnAddMem.Size = New-Object System.Drawing.Size(110, 30)
    $btnAddMem.Enabled = $canManage
    $pnlMemBtn.Controls.Add($btnAddMem)

    $btnSetMem = New-Object System.Windows.Forms.Button
    $btnSetMem.Text = '變更權限'; $btnSetMem.Location = New-Object System.Drawing.Point(130, 8); $btnSetMem.Size = New-Object System.Drawing.Size(100, 30)
    $btnSetMem.Enabled = $canManage
    $pnlMemBtn.Controls.Add($btnSetMem)

    $btnRemMem = New-Object System.Windows.Forms.Button
    $btnRemMem.Text = '移除成員'; $btnRemMem.Location = New-Object System.Drawing.Point(240, 8); $btnRemMem.Size = New-Object System.Drawing.Size(100, 30)
    $btnRemMem.Enabled = $canManage
    $pnlMemBtn.Controls.Add($btnRemMem)

    # Tab: 加入申請
    $tabReqs = New-Object System.Windows.Forms.TabPage; $tabReqs.Text = '加入申請'
    $tabs.TabPages.Add($tabReqs)
    $gridReqs = New-Object System.Windows.Forms.DataGridView
    $gridReqs.Dock = 'Fill'
    $gridReqs.ReadOnly = $true; $gridReqs.AllowUserToAddRows = $false; $gridReqs.AllowUserToDeleteRows = $false
    $gridReqs.SelectionMode = 'FullRowSelect'; $gridReqs.MultiSelect = $false; $gridReqs.RowHeadersVisible = $false
    $gridReqs.AutoSizeColumnsMode = 'Fill'
    [void]$gridReqs.Columns.Add('user', '申請人')
    [void]$gridReqs.Columns.Add('perm', '申請權限')
    [void]$gridReqs.Columns.Add('status', '狀態')
    [void]$gridReqs.Columns.Add('at', '申請時間')
    [void]$gridReqs.Columns.Add('msg', '訊息')
    $tabReqs.Controls.Add($gridReqs)
    $pnlReqBtn = New-Object System.Windows.Forms.Panel
    $pnlReqBtn.Dock = 'Top'; $pnlReqBtn.Height = 45
    $tabReqs.Controls.Add($pnlReqBtn)
    $btnAppr = New-Object System.Windows.Forms.Button
    $btnAppr.Text = '核可'; $btnAppr.Location = New-Object System.Drawing.Point(10, 8); $btnAppr.Size = New-Object System.Drawing.Size(90, 30)
    $btnAppr.Enabled = $canManage
    $pnlReqBtn.Controls.Add($btnAppr)
    $btnRej = New-Object System.Windows.Forms.Button
    $btnRej.Text = '拒絕'; $btnRej.Location = New-Object System.Drawing.Point(110, 8); $btnRej.Size = New-Object System.Drawing.Size(90, 30)
    $btnRej.Enabled = $canManage
    $pnlReqBtn.Controls.Add($btnRej)

    # Tab: 附件
    $tabFiles = New-Object System.Windows.Forms.TabPage; $tabFiles.Text = '附件總覽'
    $tabs.TabPages.Add($tabFiles)
    $gridFiles = New-Object System.Windows.Forms.DataGridView
    $gridFiles.Dock = 'Fill'
    $gridFiles.ReadOnly = $true; $gridFiles.AllowUserToAddRows = $false; $gridFiles.AllowUserToDeleteRows = $false
    $gridFiles.SelectionMode = 'FullRowSelect'; $gridFiles.MultiSelect = $false; $gridFiles.RowHeadersVisible = $false
    $gridFiles.AutoSizeColumnsMode = 'Fill'
    [void]$gridFiles.Columns.Add('phase', '所屬階段')
    [void]$gridFiles.Columns.Add('name', '檔案名稱')
    [void]$gridFiles.Columns.Add('size', '大小')
    [void]$gridFiles.Columns.Add('user', '上傳者')
    [void]$gridFiles.Columns.Add('at', '上傳時間')
    [void]$gridFiles.Columns.Add('path', '路徑')
    $tabFiles.Controls.Add($gridFiles)

    # 編輯/設定按鈕
    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Dock = 'Bottom'; $pnlBottom.Height = 40
    $f.Controls.Add($pnlBottom)
    $btnEditProj = New-Object System.Windows.Forms.Button
    $btnEditProj.Text = '編輯專案設定'; $btnEditProj.Location = New-Object System.Drawing.Point(10, 5); $btnEditProj.Size = New-Object System.Drawing.Size(130, 30)
    $btnEditProj.Enabled = $canManage
    $pnlBottom.Controls.Add($btnEditProj)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = '關閉'; $btnClose.Location = New-Object System.Drawing.Point(900, 5); $btnClose.Size = New-Object System.Drawing.Size(100, 30)
    $btnClose.Anchor = 'Top, Right'
    $btnClose.DialogResult = 'Cancel'
    $pnlBottom.Controls.Add($btnClose)

    $script:CurrentProject = $proj
    $script:PhaseRows = @()
    $script:MemberRows = @()
    $script:ReqRows = @()

    function script:Draw-Chart {
        param($Project)
        $w = [Math]::Max(200, $picChart.ClientSize.Width - 20)
        $h = [Math]::Max(140, $picChart.ClientSize.Height - 20)
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.Clear([System.Drawing.Color]::White)
            $titleFont = New-Object System.Drawing.Font('Microsoft JhengHei', 10, [System.Drawing.FontStyle]::Bold)
            $g.DrawString('各階段進度圖', $titleFont, [System.Drawing.Brushes]::Black, 5, 2)
            $phases = @()
            if ($Project.PSObject.Properties.Name -contains 'phases' -and $null -ne $Project.phases) {
                $phases = @($Project.phases)
            }
            if ($phases.Count -eq 0) {
                $g.DrawString('(尚未建立階段)', $titleFont, [System.Drawing.Brushes]::Gray, 20, 40)
            } else {
                $barH = 22; $gap = 8; $top = 30; $leftLabel = 140; $maxBarW = $w - $leftLabel - 80
                $font = New-Object System.Drawing.Font('Microsoft JhengHei', 9)
                $fill = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(70,130,200))
                $bg = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(230,230,230))
                $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Gray)
                for ($i = 0; $i -lt $phases.Count; $i++) {
                    $p = $phases[$i]
                    $y = $top + $i * ($barH + $gap)
                    $name = $p.name
                    if ($name.Length -gt 12) { $name = $name.Substring(0, 12) + '…' }
                    $g.DrawString($name, $font, [System.Drawing.Brushes]::Black, 5, $y + 3)
                    $g.FillRectangle($bg, $leftLabel, $y, $maxBarW, $barH)
                    $prog = 0
                    if ($p.PSObject.Properties.Name -contains 'progress' -and $null -ne $p.progress) { $prog = [int]$p.progress }
                    if ($prog -lt 0) { $prog = 0 }; if ($prog -gt 100) { $prog = 100 }
                    $bw = [int]([math]::Floor($maxBarW * $prog / 100))
                    if ($bw -gt 0) { $g.FillRectangle($fill, $leftLabel, $y, $bw, $barH) }
                    $g.DrawRectangle($pen, $leftLabel, $y, $maxBarW, $barH)
                    $g.DrawString("$prog %", $font, [System.Drawing.Brushes]::Black, $leftLabel + $maxBarW + 6, $y + 3)
                }
            }
        } finally { $g.Dispose() }
        if ($null -ne $picChart.Image) { $picChart.Image.Dispose() }
        $picChart.Image = $bmp
    }

    function script:Reload-All {
        $p = Get-CwtProjectById -Id $ProjectId
        if ($null -eq $p) { return }
        $script:CurrentProject = $p
        $owner = Get-CwtUserById -Id $p.ownerId
        $ownerName = if ($null -ne $owner) { $owner.displayName } else { '(未知)' }
        $lblHeader.Text = "[$($p.code)] $($p.name)"
        $due = if ([string]::IsNullOrEmpty($p.deadline)) { '(無)' } else {
            try { ([datetime]$p.deadline).ToLocalTime().ToString('yyyy-MM-dd') } catch { $p.deadline }
        }
        $lblMeta.Text = "發起人: $ownerName    狀態: $($p.status)    機密等級: $($p.confidentiality)    截止: $due`r`n說明: $($p.description)"
        $prog = Get-CwtProjectProgress -Project $p
        $progBar.Value = $prog
        $lblProgVal.Text = "整體完成度: $prog %"

        $gridPhases.Rows.Clear()
        $phases = @()
        if ($p.PSObject.Properties.Name -contains 'phases' -and $null -ne $p.phases) { $phases = @($p.phases) }
        $script:PhaseRows = $phases
        foreach ($ph in $phases) {
            $leader = Get-CwtUserById -Id $ph.leaderId
            $leaderName = if ($null -ne $leader) { $leader.displayName } else { '' }
            $collabNames = @()
            if ($ph.PSObject.Properties.Name -contains 'collaboratorIds' -and $null -ne $ph.collaboratorIds) {
                foreach ($cid in @($ph.collaboratorIds)) {
                    $cu = Get-CwtUserById -Id $cid
                    if ($null -ne $cu) { $collabNames += $cu.displayName }
                }
            }
            $dueStr = ''
            if (-not [string]::IsNullOrEmpty($ph.dueDate)) {
                try { $dueStr = ([datetime]$ph.dueDate).ToLocalTime().ToString('yyyy-MM-dd') } catch { $dueStr = $ph.dueDate }
            }
            [void]$gridPhases.Rows.Add($ph.name, $ph.kpi, $leaderName, ($collabNames -join ', '), $ph.status, "$([int]$ph.progress) %", $dueStr)
        }
        Draw-Chart -Project $p

        # Members
        $gridMembers.Rows.Clear()
        $members = @(Get-CwtProjectMembers -ProjectId $ProjectId)
        $script:MemberRows = $members
        foreach ($m in $members) {
            $mu = Get-CwtUserById -Id $m.userId
            $uName = if ($null -ne $mu) { $mu.username } else { '(未知)' }
            $dName = if ($null -ne $mu) { $mu.displayName } else { '' }
            $tm = ''
            if (-not [string]::IsNullOrEmpty($m.grantedAt)) {
                try { $tm = ([datetime]$m.grantedAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { $tm = $m.grantedAt }
            }
            [void]$gridMembers.Rows.Add($uName, $dName, $m.permission, $tm)
        }

        # Join requests
        $gridReqs.Rows.Clear()
        $reqs = @(Get-CwtJoinRequests | Where-Object { $_.projectId -eq $ProjectId })
        $script:ReqRows = $reqs
        foreach ($r in $reqs) {
            $ru = Get-CwtUserById -Id $r.userId
            $uName = if ($null -ne $ru) { "$($ru.username) ($($ru.displayName))" } else { '(未知)' }
            $tm = ''
            if (-not [string]::IsNullOrEmpty($r.requestedAt)) {
                try { $tm = ([datetime]$r.requestedAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { $tm = $r.requestedAt }
            }
            [void]$gridReqs.Rows.Add($uName, $r.requestedPermission, $r.status, $tm, $r.message)
        }

        # Files
        $gridFiles.Rows.Clear()
        foreach ($ph in $phases) {
            $atts = @()
            if ($ph.PSObject.Properties.Name -contains 'attachments' -and $null -ne $ph.attachments) {
                $atts = @($ph.attachments)
            }
            foreach ($a in $atts) {
                $uu = Get-CwtUserById -Id $a.uploadedBy
                $uName = if ($null -ne $uu) { $uu.displayName } else { '' }
                $tm = ''
                if (-not [string]::IsNullOrEmpty($a.uploadedAt)) {
                    try { $tm = ([datetime]$a.uploadedAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { $tm = $a.uploadedAt }
                }
                $kb = '{0:N1} KB' -f (([double]$a.size) / 1024.0)
                [void]$gridFiles.Rows.Add($ph.name, $a.name, $kb, $uName, $tm, $a.path)
            }
        }
    }

    $picChart.Add_Resize({
        if ($null -ne $script:CurrentProject) { Draw-Chart -Project $script:CurrentProject }
    })

    $btnAddPhase.Add_Click({
        $r = Show-CwtPhaseEditForm -CurrentUser $CurrentUser -ProjectId $ProjectId -Phase $null
        if ($r) { Reload-All }
    })
    $btnEditPhase.Add_Click({
        if ($gridPhases.SelectedRows.Count -eq 0) { return }
        $idx = $gridPhases.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:PhaseRows.Count) { return }
        $r = Show-CwtPhaseEditForm -CurrentUser $CurrentUser -ProjectId $ProjectId -Phase $script:PhaseRows[$idx]
        if ($r) { Reload-All }
    })
    $btnDelPhase.Add_Click({
        if ($gridPhases.SelectedRows.Count -eq 0) { return }
        $idx = $gridPhases.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:PhaseRows.Count) { return }
        $ph = $script:PhaseRows[$idx]
        $r = [System.Windows.Forms.MessageBox]::Show("確定刪除階段「$($ph.name)」？", '確認', 'YesNo', 'Warning')
        if ($r -eq 'Yes') {
            Remove-CwtPhase -ProjectId $ProjectId -PhaseId $ph.id -ActorId $CurrentUser.id
            Reload-All
        }
    })

    $btnUploadFile.Add_Click({
        if ($gridPhases.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('請先選擇要上傳附件的階段', '提示', 'OK', 'Information') | Out-Null
            return
        }
        $idx = $gridPhases.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:PhaseRows.Count) { return }
        $ph = $script:PhaseRows[$idx]
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Multiselect = $true
        $ofd.Filter = '所有檔案 (*.*)|*.*|圖片 (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp'
        if ($ofd.ShowDialog() -eq 'OK') {
            foreach ($fp in $ofd.FileNames) {
                try {
                    Add-CwtPhaseAttachment -ProjectId $ProjectId -PhaseId $ph.id -SourcePath $fp -ActorId $CurrentUser.id | Out-Null
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("上傳失敗: $($_.Exception.Message)", '錯誤', 'OK', 'Error') | Out-Null
                }
            }
            Reload-All
        }
    })

    $btnOpenFolder.Add_Click({
        $paths = Get-CwtPaths
        $dir = Join-Path $paths.Uploads $script:CurrentProject.code
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        try { Start-Process -FilePath 'explorer.exe' -ArgumentList $dir } catch { }
    })

    $btnEditProj.Add_Click({
        $r = Show-CwtProjectEditForm -CurrentUser $CurrentUser -Project $script:CurrentProject
        if ($r) { Reload-All }
    })

    $btnAddMem.Add_Click({
        Show-CwtMemberPickerForm -ProjectId $ProjectId -ActorId $CurrentUser.id
        Reload-All
    })
    $btnSetMem.Add_Click({
        if ($gridMembers.SelectedRows.Count -eq 0) { return }
        $idx = $gridMembers.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:MemberRows.Count) { return }
        $m = $script:MemberRows[$idx]
        $perm = Read-CwtPermissionInput -Current $m.permission
        if ($perm) {
            Set-CwtMembership -ProjectId $ProjectId -UserId $m.userId -Permission $perm -GrantedBy $CurrentUser.id
            Reload-All
        }
    })
    $btnRemMem.Add_Click({
        if ($gridMembers.SelectedRows.Count -eq 0) { return }
        $idx = $gridMembers.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:MemberRows.Count) { return }
        $m = $script:MemberRows[$idx]
        if ($m.userId -eq $script:CurrentProject.ownerId) {
            [System.Windows.Forms.MessageBox]::Show('無法移除發起人', '提示', 'OK', 'Warning') | Out-Null
            return
        }
        $r = [System.Windows.Forms.MessageBox]::Show('確定移除此成員？', '確認', 'YesNo', 'Warning')
        if ($r -eq 'Yes') {
            Remove-CwtMembership -ProjectId $ProjectId -UserId $m.userId -ActorId $CurrentUser.id
            Reload-All
        }
    })

    $btnAppr.Add_Click({
        if ($gridReqs.SelectedRows.Count -eq 0) { return }
        $idx = $gridReqs.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:ReqRows.Count) { return }
        $rq = $script:ReqRows[$idx]
        if ($rq.status -ne 'Pending') { return }
        try {
            Resolve-CwtJoinRequest -RequestId $rq.id -Decision 'Approved' -DeciderId $CurrentUser.id | Out-Null
            Reload-All
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '錯誤', 'OK', 'Error') | Out-Null
        }
    })
    $btnRej.Add_Click({
        if ($gridReqs.SelectedRows.Count -eq 0) { return }
        $idx = $gridReqs.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:ReqRows.Count) { return }
        $rq = $script:ReqRows[$idx]
        if ($rq.status -ne 'Pending') { return }
        try {
            Resolve-CwtJoinRequest -RequestId $rq.id -Decision 'Rejected' -DeciderId $CurrentUser.id | Out-Null
            Reload-All
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '錯誤', 'OK', 'Error') | Out-Null
        }
    })

    $f.Add_Shown({ Reload-All })
    $f.CancelButton = $btnClose
    $null = $f.ShowDialog()
}

function Show-CwtPhaseEditForm {
    param(
        [Parameter(Mandatory)]$CurrentUser,
        [Parameter(Mandatory)][string]$ProjectId,
        $Phase = $null
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $isNew = ($null -eq $Phase)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = if ($isNew) { '新增階段' } else { "編輯階段 - $($Phase.name)" }
    $f.Size = New-Object System.Drawing.Size(560, 540); $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false

    $y = 20
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = '階段名稱 *:'; $lblName.Location = New-Object System.Drawing.Point(20, $y); $lblName.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblName)
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(120, ($y - 3)); $txtName.Size = New-Object System.Drawing.Size(400, 25)
    if (-not $isNew) { $txtName.Text = $Phase.name }
    $f.Controls.Add($txtName)

    $y += 35
    $lblKpi = New-Object System.Windows.Forms.Label
    $lblKpi.Text = 'KPI:'; $lblKpi.Location = New-Object System.Drawing.Point(20, $y); $lblKpi.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblKpi)
    $txtKpi = New-Object System.Windows.Forms.TextBox
    $txtKpi.Multiline = $true; $txtKpi.ScrollBars = 'Vertical'
    $txtKpi.Location = New-Object System.Drawing.Point(120, ($y - 3)); $txtKpi.Size = New-Object System.Drawing.Size(400, 60)
    if (-not $isNew) { $txtKpi.Text = $Phase.kpi }
    $f.Controls.Add($txtKpi)

    $y += 70
    $members = @(Get-CwtProjectMembers -ProjectId $ProjectId)
    $memberItems = @()
    $memberMap = @{}
    foreach ($m in $members) {
        $u = Get-CwtUserById -Id $m.userId
        if ($null -ne $u) {
            $label = "$($u.username) - $($u.displayName)"
            $memberItems += $label
            $memberMap[$label] = $u.id
        }
    }

    $lblLeader = New-Object System.Windows.Forms.Label
    $lblLeader.Text = '負責人:'; $lblLeader.Location = New-Object System.Drawing.Point(20, $y); $lblLeader.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblLeader)
    $cmbLeader = New-Object System.Windows.Forms.ComboBox; $cmbLeader.DropDownStyle = 'DropDownList'
    [void]$cmbLeader.Items.Add('(未指定)')
    foreach ($it in $memberItems) { [void]$cmbLeader.Items.Add($it) }
    $cmbLeader.SelectedIndex = 0
    if (-not $isNew -and -not [string]::IsNullOrEmpty($Phase.leaderId)) {
        $lu = Get-CwtUserById -Id $Phase.leaderId
        if ($null -ne $lu) {
            $key = "$($lu.username) - $($lu.displayName)"
            $idx = $cmbLeader.Items.IndexOf($key)
            if ($idx -ge 0) { $cmbLeader.SelectedIndex = $idx }
        }
    }
    $cmbLeader.Location = New-Object System.Drawing.Point(120, ($y - 3)); $cmbLeader.Size = New-Object System.Drawing.Size(400, 25)
    $f.Controls.Add($cmbLeader)

    $y += 35
    $lblCol = New-Object System.Windows.Forms.Label
    $lblCol.Text = '協同人員:'; $lblCol.Location = New-Object System.Drawing.Point(20, $y); $lblCol.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblCol)
    $clbCol = New-Object System.Windows.Forms.CheckedListBox
    $clbCol.Location = New-Object System.Drawing.Point(120, ($y - 3)); $clbCol.Size = New-Object System.Drawing.Size(400, 110)
    $clbCol.CheckOnClick = $true
    foreach ($it in $memberItems) { [void]$clbCol.Items.Add($it) }
    if (-not $isNew -and $Phase.PSObject.Properties.Name -contains 'collaboratorIds' -and $null -ne $Phase.collaboratorIds) {
        $existing = @($Phase.collaboratorIds)
        for ($i = 0; $i -lt $clbCol.Items.Count; $i++) {
            $key = $clbCol.Items[$i]
            if ($memberMap.ContainsKey($key) -and $existing -contains $memberMap[$key]) {
                $clbCol.SetItemChecked($i, $true)
            }
        }
    }
    $f.Controls.Add($clbCol)

    $y += 120
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = '開始日期:'; $lblStart.Location = New-Object System.Drawing.Point(20, $y); $lblStart.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblStart)
    $dtStart = New-Object System.Windows.Forms.DateTimePicker
    $dtStart.Format = 'Short'; $dtStart.ShowCheckBox = $true; $dtStart.Checked = $false
    $dtStart.Location = New-Object System.Drawing.Point(120, ($y - 3)); $dtStart.Size = New-Object System.Drawing.Size(150, 25)
    if (-not $isNew -and -not [string]::IsNullOrEmpty($Phase.startDate)) {
        try { $dtStart.Value = ([datetime]$Phase.startDate).ToLocalTime(); $dtStart.Checked = $true } catch { }
    }
    $f.Controls.Add($dtStart)
    $lblDue = New-Object System.Windows.Forms.Label
    $lblDue.Text = '截止日期:'; $lblDue.Location = New-Object System.Drawing.Point(290, $y); $lblDue.Size = New-Object System.Drawing.Size(70, 25)
    $f.Controls.Add($lblDue)
    $dtDue = New-Object System.Windows.Forms.DateTimePicker
    $dtDue.Format = 'Short'; $dtDue.ShowCheckBox = $true; $dtDue.Checked = $false
    $dtDue.Location = New-Object System.Drawing.Point(365, ($y - 3)); $dtDue.Size = New-Object System.Drawing.Size(155, 25)
    if (-not $isNew -and -not [string]::IsNullOrEmpty($Phase.dueDate)) {
        try { $dtDue.Value = ([datetime]$Phase.dueDate).ToLocalTime(); $dtDue.Checked = $true } catch { }
    }
    $f.Controls.Add($dtDue)

    $y += 35
    $lblSt = New-Object System.Windows.Forms.Label
    $lblSt.Text = '狀態:'; $lblSt.Location = New-Object System.Drawing.Point(20, $y); $lblSt.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblSt)
    $cmbSt = New-Object System.Windows.Forms.ComboBox; $cmbSt.DropDownStyle = 'DropDownList'
    [void]$cmbSt.Items.AddRange(@('NotStarted','InProgress','Done','Blocked'))
    $cmbSt.Location = New-Object System.Drawing.Point(120, ($y - 3)); $cmbSt.Size = New-Object System.Drawing.Size(150, 25)
    if ($isNew) { $cmbSt.SelectedItem = 'NotStarted' } else { $cmbSt.SelectedItem = $Phase.status }
    $f.Controls.Add($cmbSt)

    $lblPg = New-Object System.Windows.Forms.Label
    $lblPg.Text = '進度 %:'; $lblPg.Location = New-Object System.Drawing.Point(290, $y); $lblPg.Size = New-Object System.Drawing.Size(70, 25)
    $f.Controls.Add($lblPg)
    $numPg = New-Object System.Windows.Forms.NumericUpDown
    $numPg.Minimum = 0; $numPg.Maximum = 100
    $numPg.Location = New-Object System.Drawing.Point(365, ($y - 3)); $numPg.Size = New-Object System.Drawing.Size(80, 25)
    if (-not $isNew) { $numPg.Value = [int]$Phase.progress }
    $f.Controls.Add($numPg)

    $y += 35
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = '進度說明:'; $lblNote.Location = New-Object System.Drawing.Point(20, $y); $lblNote.Size = New-Object System.Drawing.Size(90, 25)
    $f.Controls.Add($lblNote)
    $txtNote = New-Object System.Windows.Forms.TextBox
    $txtNote.Multiline = $true; $txtNote.ScrollBars = 'Vertical'
    $txtNote.Location = New-Object System.Drawing.Point(120, ($y - 3)); $txtNote.Size = New-Object System.Drawing.Size(400, 60)
    if (-not $isNew) { $txtNote.Text = $Phase.progressNote }
    $f.Controls.Add($txtNote)

    $msg = New-Object System.Windows.Forms.Label
    $msg.Location = New-Object System.Drawing.Point(20, 460); $msg.Size = New-Object System.Drawing.Size(400, 25)
    $msg.ForeColor = [System.Drawing.Color]::Firebrick
    $f.Controls.Add($msg)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = if ($isNew) { '建立階段' } else { '儲存' }
    $btnOk.Location = New-Object System.Drawing.Point(300, 455); $btnOk.Size = New-Object System.Drawing.Size(100, 32)
    $f.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '取消'; $btnCancel.Location = New-Object System.Drawing.Point(410, 455); $btnCancel.Size = New-Object System.Drawing.Size(110, 32)
    $btnCancel.DialogResult = 'Cancel'
    $f.Controls.Add($btnCancel)

    $script:PhaseResult = $null
    $btnOk.Add_Click({
        $msg.Text = ''
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) { $msg.Text = '請輸入階段名稱'; return }
        $leaderId = ''
        if ($cmbLeader.SelectedIndex -gt 0) {
            $key = $cmbLeader.SelectedItem
            if ($memberMap.ContainsKey($key)) { $leaderId = $memberMap[$key] }
        }
        $collabIds = @()
        for ($i = 0; $i -lt $clbCol.Items.Count; $i++) {
            if ($clbCol.GetItemChecked($i)) {
                $key = $clbCol.Items[$i]
                if ($memberMap.ContainsKey($key)) { $collabIds += $memberMap[$key] }
            }
        }
        $startStr = ''; if ($dtStart.Checked) { $startStr = $dtStart.Value.ToUniversalTime().ToString('o') }
        $dueStr = ''; if ($dtDue.Checked) { $dueStr = $dtDue.Value.ToUniversalTime().ToString('o') }
        try {
            if ($isNew) {
                $start = $null; if ($dtStart.Checked) { $start = $dtStart.Value }
                $due = $null; if ($dtDue.Checked) { $due = $dtDue.Value }
                Add-CwtPhase -ProjectId $ProjectId -Name $txtName.Text -Kpi $txtKpi.Text `
                    -LeaderId $leaderId -CollaboratorIds $collabIds `
                    -StartDate $start -DueDate $due -ActorId $CurrentUser.id | Out-Null
                # 設進度與狀態
                $proj = Get-CwtProjectById -Id $ProjectId
                $newPhase = @($proj.phases) | Where-Object { $_.name -eq $txtName.Text } | Select-Object -Last 1
                if ($null -ne $newPhase) {
                    Update-CwtPhase -ProjectId $ProjectId -PhaseId $newPhase.id -ActorId $CurrentUser.id -Set @{
                        progress = [int]$numPg.Value
                        progressNote = $txtNote.Text
                        status = $cmbSt.SelectedItem
                    } | Out-Null
                }
            } else {
                Update-CwtPhase -ProjectId $ProjectId -PhaseId $Phase.id -ActorId $CurrentUser.id -Set @{
                    name = $txtName.Text
                    kpi = $txtKpi.Text
                    leaderId = $leaderId
                    collaboratorIds = $collabIds
                    startDate = $startStr
                    dueDate = $dueStr
                    status = $cmbSt.SelectedItem
                    progress = [int]$numPg.Value
                    progressNote = $txtNote.Text
                } | Out-Null
            }
            $script:PhaseResult = $true
            $f.DialogResult = 'OK'; $f.Close()
        } catch { $msg.Text = $_.Exception.Message }
    })

    $f.AcceptButton = $btnOk; $f.CancelButton = $btnCancel
    $null = $f.ShowDialog()
    return $script:PhaseResult
}

function Show-CwtMemberPickerForm {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$ActorId
    )
    Add-Type -AssemblyName System.Windows.Forms

    $f = New-Object System.Windows.Forms.Form
    $f.Text = '加入成員'; $f.Size = New-Object System.Drawing.Size(500, 480); $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = '選擇使用者:'; $lbl.Location = New-Object System.Drawing.Point(15, 15); $lbl.Size = New-Object System.Drawing.Size(120, 25)
    $f.Controls.Add($lbl)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(15, 45); $clb.Size = New-Object System.Drawing.Size(460, 280)
    $clb.CheckOnClick = $true
    $f.Controls.Add($clb)
    $map = @{}
    $users = @(Get-CwtUsers | Where-Object { $_.active })
    $existing = @(Get-CwtProjectMembers -ProjectId $ProjectId | ForEach-Object { $_.userId })
    foreach ($u in $users) {
        if ($existing -contains $u.id) { continue }
        $label = "$($u.username) - $($u.displayName)"
        [void]$clb.Items.Add($label)
        $map[$label] = $u.id
    }

    $lblPerm = New-Object System.Windows.Forms.Label
    $lblPerm.Text = '權限:'; $lblPerm.Location = New-Object System.Drawing.Point(15, 340); $lblPerm.Size = New-Object System.Drawing.Size(60, 25)
    $f.Controls.Add($lblPerm)
    $cmbPerm = New-Object System.Windows.Forms.ComboBox; $cmbPerm.DropDownStyle = 'DropDownList'
    [void]$cmbPerm.Items.AddRange(@('Viewer','Editor','Owner'))
    $cmbPerm.SelectedItem = 'Viewer'
    $cmbPerm.Location = New-Object System.Drawing.Point(80, 337); $cmbPerm.Size = New-Object System.Drawing.Size(150, 25)
    $f.Controls.Add($cmbPerm)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = '加入'; $btnOk.Location = New-Object System.Drawing.Point(260, 395); $btnOk.Size = New-Object System.Drawing.Size(100, 32)
    $f.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '關閉'; $btnCancel.Location = New-Object System.Drawing.Point(375, 395); $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
    $btnCancel.DialogResult = 'Cancel'
    $f.Controls.Add($btnCancel)

    $btnOk.Add_Click({
        $sel = @()
        for ($i = 0; $i -lt $clb.Items.Count; $i++) {
            if ($clb.GetItemChecked($i)) { $sel += $clb.Items[$i] }
        }
        if ($sel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('請至少選一位使用者', '提示', 'OK', 'Information') | Out-Null
            return
        }
        foreach ($k in $sel) {
            if ($map.ContainsKey($k)) {
                Set-CwtMembership -ProjectId $ProjectId -UserId $map[$k] -Permission $cmbPerm.SelectedItem -GrantedBy $ActorId
            }
        }
        $f.DialogResult = 'OK'; $f.Close()
    })

    $f.CancelButton = $btnCancel
    $null = $f.ShowDialog()
}

function Read-CwtPermissionInput {
    param([string]$Current = 'Viewer')
    Add-Type -AssemblyName System.Windows.Forms
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '變更權限'; $f.Size = New-Object System.Drawing.Size(320, 180); $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = '新權限:'; $lbl.Location = New-Object System.Drawing.Point(20, 25); $lbl.Size = New-Object System.Drawing.Size(70, 25)
    $f.Controls.Add($lbl)
    $cmb = New-Object System.Windows.Forms.ComboBox; $cmb.DropDownStyle = 'DropDownList'
    [void]$cmb.Items.AddRange(@('Viewer','Editor','Owner'))
    $cmb.SelectedItem = $Current
    $cmb.Location = New-Object System.Drawing.Point(95, 22); $cmb.Size = New-Object System.Drawing.Size(180, 25)
    $f.Controls.Add($cmb)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = '確定'; $ok.Location = New-Object System.Drawing.Point(70, 85); $ok.Size = New-Object System.Drawing.Size(90, 32)
    $f.Controls.Add($ok)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = '取消'; $cancel.Location = New-Object System.Drawing.Point(180, 85); $cancel.Size = New-Object System.Drawing.Size(90, 32)
    $cancel.DialogResult = 'Cancel'
    $f.Controls.Add($cancel)

    $script:PermPick = $null
    $ok.Add_Click({ $script:PermPick = $cmb.SelectedItem; $f.DialogResult = 'OK'; $f.Close() })
    $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $null = $f.ShowDialog()
    return $script:PermPick
}
