# RegisterForm.ps1
function Show-CwtRegisterForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = '新使用者註冊'
    $form.Size = New-Object System.Drawing.Size(440, 360)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = '帳號 *:'; $lbl1.Location = New-Object System.Drawing.Point(20, 25); $lbl1.Size = New-Object System.Drawing.Size(90, 25)
    $form.Controls.Add($lbl1)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(120, 22); $txtUser.Size = New-Object System.Drawing.Size(280, 25)
    $form.Controls.Add($txtUser)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text = '顯示名稱:'; $lbl2.Location = New-Object System.Drawing.Point(20, 60); $lbl2.Size = New-Object System.Drawing.Size(90, 25)
    $form.Controls.Add($lbl2)
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(120, 57); $txtName.Size = New-Object System.Drawing.Size(280, 25)
    $form.Controls.Add($txtName)

    $lbl3 = New-Object System.Windows.Forms.Label
    $lbl3.Text = 'Email:'; $lbl3.Location = New-Object System.Drawing.Point(20, 95); $lbl3.Size = New-Object System.Drawing.Size(90, 25)
    $form.Controls.Add($lbl3)
    $txtEmail = New-Object System.Windows.Forms.TextBox
    $txtEmail.Location = New-Object System.Drawing.Point(120, 92); $txtEmail.Size = New-Object System.Drawing.Size(280, 25)
    $form.Controls.Add($txtEmail)

    $lbl4 = New-Object System.Windows.Forms.Label
    $lbl4.Text = '密碼 *:'; $lbl4.Location = New-Object System.Drawing.Point(20, 130); $lbl4.Size = New-Object System.Drawing.Size(90, 25)
    $form.Controls.Add($lbl4)
    $txtPwd1 = New-Object System.Windows.Forms.TextBox
    $txtPwd1.Location = New-Object System.Drawing.Point(120, 127); $txtPwd1.Size = New-Object System.Drawing.Size(280, 25)
    $txtPwd1.UseSystemPasswordChar = $true
    $form.Controls.Add($txtPwd1)

    $lbl5 = New-Object System.Windows.Forms.Label
    $lbl5.Text = '確認密碼 *:'; $lbl5.Location = New-Object System.Drawing.Point(20, 165); $lbl5.Size = New-Object System.Drawing.Size(95, 25)
    $form.Controls.Add($lbl5)
    $txtPwd2 = New-Object System.Windows.Forms.TextBox
    $txtPwd2.Location = New-Object System.Drawing.Point(120, 162); $txtPwd2.Size = New-Object System.Drawing.Size(280, 25)
    $txtPwd2.UseSystemPasswordChar = $true
    $form.Controls.Add($txtPwd2)

    $lblRule = New-Object System.Windows.Forms.Label
    $lblRule.Text = '* 密碼至少 8 個字元，建議包含英數字混合。新使用者預設為一般使用者。'
    $lblRule.Location = New-Object System.Drawing.Point(20, 200); $lblRule.Size = New-Object System.Drawing.Size(380, 40)
    $lblRule.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($lblRule)

    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Location = New-Object System.Drawing.Point(20, 240); $lblMsg.Size = New-Object System.Drawing.Size(380, 25)
    $lblMsg.ForeColor = [System.Drawing.Color]::Firebrick
    $form.Controls.Add($lblMsg)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = '建立帳號'
    $btnOk.Location = New-Object System.Drawing.Point(140, 280); $btnOk.Size = New-Object System.Drawing.Size(110, 32)
    $form.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '取消'
    $btnCancel.Location = New-Object System.Drawing.Point(265, 280); $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)

    $script:RegResult = $null

    $btnOk.Add_Click({
        $lblMsg.Text = ''
        $u = $txtUser.Text.Trim()
        $p1 = $txtPwd1.Text
        $p2 = $txtPwd2.Text
        if ([string]::IsNullOrWhiteSpace($u)) { $lblMsg.Text = '請輸入帳號'; return }
        if ($u -notmatch '^[A-Za-z0-9_\.\-]{3,32}$') { $lblMsg.Text = '帳號僅可使用英數及 _ . -，3~32 字元'; return }
        if ([string]::IsNullOrEmpty($p1)) { $lblMsg.Text = '請輸入密碼'; return }
        if ($p1.Length -lt 8) { $lblMsg.Text = '密碼至少 8 個字元'; return }
        if ($p1 -ne $p2) { $lblMsg.Text = '兩次密碼不一致'; return }
        try {
            $newUser = New-CwtUser -Username $u -Password $p1 -DisplayName $txtName.Text -Email $txtEmail.Text -Role 'User'
            $script:RegResult = $newUser
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        } catch {
            $lblMsg.Text = $_.Exception.Message
        }
    })

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel
    $null = $form.ShowDialog()
    return $script:RegResult
}
