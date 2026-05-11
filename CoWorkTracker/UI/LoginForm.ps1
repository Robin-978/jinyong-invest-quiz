# LoginForm.ps1
function Show-CwtLoginForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = '協同工作紀錄系統 - 登入'
    $form.Size = New-Object System.Drawing.Size(420, 280)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = '協同工作紀錄系統'
    $lblTitle.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(380, 30)
    $lblTitle.TextAlign = 'MiddleCenter'
    $form.Controls.Add($lblTitle)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = '帳號:'
    $lblUser.Location = New-Object System.Drawing.Point(40, 65)
    $lblUser.Size = New-Object System.Drawing.Size(80, 25)
    $form.Controls.Add($lblUser)

    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(130, 62)
    $txtUser.Size = New-Object System.Drawing.Size(240, 25)
    $form.Controls.Add($txtUser)

    $lblPwd = New-Object System.Windows.Forms.Label
    $lblPwd.Text = '密碼:'
    $lblPwd.Location = New-Object System.Drawing.Point(40, 100)
    $lblPwd.Size = New-Object System.Drawing.Size(80, 25)
    $form.Controls.Add($lblPwd)

    $txtPwd = New-Object System.Windows.Forms.TextBox
    $txtPwd.Location = New-Object System.Drawing.Point(130, 97)
    $txtPwd.Size = New-Object System.Drawing.Size(240, 25)
    $txtPwd.UseSystemPasswordChar = $true
    $form.Controls.Add($txtPwd)

    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Location = New-Object System.Drawing.Point(40, 130)
    $lblMsg.Size = New-Object System.Drawing.Size(330, 25)
    $lblMsg.ForeColor = [System.Drawing.Color]::Firebrick
    $form.Controls.Add($lblMsg)

    $btnLogin = New-Object System.Windows.Forms.Button
    $btnLogin.Text = '登入'
    $btnLogin.Location = New-Object System.Drawing.Point(130, 165)
    $btnLogin.Size = New-Object System.Drawing.Size(110, 32)
    $form.Controls.Add($btnLogin)

    $btnRegister = New-Object System.Windows.Forms.Button
    $btnRegister.Text = '新使用者註冊'
    $btnRegister.Location = New-Object System.Drawing.Point(255, 165)
    $btnRegister.Size = New-Object System.Drawing.Size(115, 32)
    $form.Controls.Add($btnRegister)

    $lblTip = New-Object System.Windows.Forms.Label
    $lblTip.Text = '預設管理員: admin / admin@123 (首次登入請立即修改密碼)'
    $lblTip.Location = New-Object System.Drawing.Point(20, 210)
    $lblTip.Size = New-Object System.Drawing.Size(380, 20)
    $lblTip.ForeColor = [System.Drawing.Color]::Gray
    $lblTip.TextAlign = 'MiddleCenter'
    $form.Controls.Add($lblTip)

    $script:LoggedInUser = $null

    $btnLogin.Add_Click({
        $lblMsg.Text = ''
        $u = $txtUser.Text.Trim()
        $p = $txtPwd.Text
        if ([string]::IsNullOrWhiteSpace($u) -or [string]::IsNullOrWhiteSpace($p)) {
            $lblMsg.Text = '請輸入帳號與密碼'
            return
        }
        try {
            $user = Test-CwtLogin -Username $u -Password $p
            if ($null -eq $user) {
                $lblMsg.Text = '帳號或密碼錯誤，或帳號未啟用'
                Write-CwtLog -Action 'login.fail' -ActorName $u -Level 'Security'
                return
            }
            Write-CwtLog -Action 'login.success' -ActorId $user.id -ActorName $user.username -Level 'Security'
            $script:LoggedInUser = $user
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        } catch {
            $lblMsg.Text = $_.Exception.Message
        }
    })

    $btnRegister.Add_Click({
        $reg = Show-CwtRegisterForm
        if ($null -ne $reg) {
            $txtUser.Text = $reg.username
            $txtPwd.Text = ''
            $lblMsg.ForeColor = [System.Drawing.Color]::DarkGreen
            $lblMsg.Text = "註冊成功，請使用 $($reg.username) 登入"
        }
    })

    $form.AcceptButton = $btnLogin
    $null = $form.ShowDialog()
    return $script:LoggedInUser
}
