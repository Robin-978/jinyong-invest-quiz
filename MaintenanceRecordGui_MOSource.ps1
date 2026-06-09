# PATCH_SUMMARY 2026-06-09 MO Source PDF One-Row-Per-Source Fix
# - 修正 PDF/Excel 匯出時，MO Source 只有一筆資料會被拆成直向 12 列（一欄一個欄位值）的問題。
# - 根因：Get-MoSourceRowsForExport 以 ",$rows" 慣用法回傳完整列集合，呼叫端卻又 Get-CollectionItem 0，
#   反而取到第一列，使單一 source 的 12 欄被當成 12 列輸出。
# - 修正四處呼叫端，直接使用回傳的列集合，確保「一列代表一個 source、12 欄橫向」。
# - 不改維修紀錄核心邏輯、欄位、鎖定、稽核與草稿資料結構。

# PATCH_SUMMARY 2026-06-09 PDF Landscape Orientation
# - 匯出 PDF 改為橫式 (landscape) 列印，而非直式 (portrait)，讓 12 欄 MO Source 表格完整顯示。
# - Excel COM 匯出：PageSetup.Orientation 由 1 (xlPortrait) 改為 2 (xlLandscape)，FitToPagesWide 仍為 1。
# - 影像式 PDF 退回路徑：MediaBox 固定為 792x612 橫式，頁面內容維持原始比例置中，避免拉伸或裁切。
# - 不改維修紀錄核心邏輯、欄位、鎖定、稽核與草稿資料結構。

# PATCH_SUMMARY 2026-06-08 MO Source PDF ASCII Header Fix
# - 修正 MO Source PDF 表格內中文欄位標籤在部分 Windows/GDI+ PDF 影像輸出環境顯示空白的問題。
# - 僅調整 PDF 的 MO Source 欄位標籤為 ASCII 英文短標籤，避免中文字型 fallback 或小欄寬造成無法顯示。
# - GUI、草稿、JSON、CSV、TXT 仍保留原本中文欄位名稱與完整 12 欄資料。
# - 不改維修核心邏輯、鎖定、稽核、草稿資料結構。

# PATCH_SUMMARY 2026-06-08 MO Source PDF Header Label Fix
# - 修正 MO Source PDF 表頭只顯示 Source、其餘標籤空白或未帶入的問題。
# - 低風險補強：PDF 若偵測到 MO Source 有資料，會額外插入一列「標籤資料列」，讓 Source/設定瓶壓/位置/重量Off...成為表格資料的一部分。
# - 調整 Draw-PdfRow 參數型別，避免 PowerShell 對 [object[]] 參數繫結造成陣列被錯誤拆解或只取第一格。
# - 保留草稿橫向一列一筆、AI JSON/CSV/TXT、稽核、鎖定與維修核心邏輯不變。
# - 相容 Windows PowerShell 2.0-5.1 與 .NET Framework；未使用 PowerShell 7-only 語法。

# PATCH_SUMMARY 2026-06-08 MO Source Row Mapping Fix
# - 修正讀取草稿時 `.ToArray()` 在 Windows PowerShell 2.0-5.1 可能把 row 視為 System.String 而失敗的錯誤。
# - 修正 MO Source 草稿讀回與 PDF 橫向表格欄位對位，避免一列資料被拆到 Source 欄直式顯示。
# - 維持 MO Source 在 PDF 與暫存草稿為橫向一列一筆。
# - 保留舊版草稿讀取相容，不改維修紀錄核心邏輯。

# PATCH_SUMMARY 2026-06-08 mo-source-fields-01
# - 低風險修改：統一 MO SOURCE 更換紀錄欄位；PDF 與暫存草稿均維持一列一筆、12 欄橫向格式。
# - 新增/同步欄位：Source、設定瓶壓、位置、重量Off、重量On、重量T/W、來料瓶壓、實際瓶壓、效率Off、效率On、退庫數量(瓶)、備註。
# - 修正 MO Source DataGridView 欄位 Name 重複問題，避免欄位新增或讀取時發生衝突。
# - 保留原本維修紀錄核心邏輯、鎖定流程、稽核流程與既有 PDF 影像式輸出架構。
# - 相容 Windows PowerShell 2.0-5.1 與 .NET Framework；未使用 PowerShell 7-only 語法。

# PATCH_SUMMARY 2026-06-04
# - 修正 PDF 匯出可能空白：新增 Convert-ToPdfImageBytes，避免 PowerShell 將 JPEG byte[] 包成巢狀 object[]。
# - 強化 Write-ImagePdf：加入 JPEG SOI 檢查、/ProcSet [/PDF /ImageC]、以實際 ASCII byte length 寫入 Content stream。
# - 保留原本 GUI、欄位、稽核、鎖定與 AI 匯出流程，不改維修紀錄核心邏輯。
# - 相容 Windows PowerShell 2.0-5.1 與 .NET Framework；未使用 PowerShell 7-only 語法。

# PATCH_SUMMARY 2026-06-08 Auto AI Output On Record Lock
# - 維修紀錄完成鎖定或重新鎖定時，自動更新 AI 預覽並儲存 JSON / CSV / TXT，避免忘記手動輸出。
# - AI輸出分頁仍保留手動「更新AI預覽」與「儲存 JSON / CSV / TXT」，並改用同一套共用函式。
# - 自動輸出失敗時不維持假鎖定狀態；管制修改重新鎖定失敗時會回到管制修改中。
# - 保留既有稽核、鎖定、PDF、草稿與 MO Source 資料結構。

# PATCH_SUMMARY 2026-06-09 MO Source PM-Before Controls
# - 在「紀錄/MO source」分頁新增匯出維修前PDF、維修項目完成鎖定、管制修改維修項目與狀態顯示。
# - 新增按鈕共用既有維修項目鎖定流程，不建立第二套狀態，避免稽核與鎖定狀態分裂。
# - 維修前 PDF 改為保留目前 MO Source 表格內容，維修紀錄區仍輸出空白列供維修後填寫。

# PATCH_SUMMARY 2026-06-09 MO Source PDF Horizontal Row Fix
# - 修正 MO Source 匯出 PDF 時資料被拆成直向一欄的問題。
# - PDF 版面改為與 GUI 相同資料粒度：一筆 MO Source 更換紀錄輸出為一列，12 個欄位橫向排列。

# PATCH_SUMMARY 2026-06-09 MO Source PDF Header Alignment Fix
# - MO Source PDF 表格改用專用逐欄繪製，避免通用列函式在小欄寬下出現表頭位移或裁切。
# - 縮短 PDF 表頭為 Src/SetP/Pos/OffW/OnW/TW/InP/ActP/OffE/OnE/Ret/Note，資料仍維持 12 欄對位。

# PATCH_SUMMARY 2026-06-09 Lock Validation Confirm Autosave
# - 新增鎖定前必填檢查與摘要確認，降低誤鎖與漏填風險。
# - 新增未完成資料自動保存到 drafts\autosave，不寫稽核 Log，避免中斷時資料遺失。

# PATCH_SUMMARY 2026-06-09 Readonly Completed PDF Preview Attachments
# - 新增完成紀錄唯讀檢視：從查詢結果選取完成紀錄後讀取 TXT/CSV 路徑並顯示，不覆蓋目前表單。
# - PDF 匯出完成後可直接開啟預覽。
# - 基本資料新增附件/照片路徑欄位，納入草稿、AI JSON/CSV/TXT。

# PATCH_SUMMARY 2026-06-09 Prelock Checklist Attachment Check Reprint PDF
# - 新增完成前檢查清單：鎖定前列出附件、CBr4、烘烤日期、維修紀錄與 MO Source 提醒。
# - 附件/照片路徑可手動檢查檔案或資料夾是否存在。
# - 查詢/匯出分頁可從已完成紀錄搜尋並開啟完整維修 PDF，供預覽與補印。

# PATCH_SUMMARY 2026-06-09 MO Source PDF Single Row Guard
# - 修正 PDF 繪製時若 MO Source 只有一筆資料，PowerShell 將單列 12 欄誤當成 12 列輸出。
# - 新增 PDF 專用正規化，強制 MO Source 一筆資料維持橫向一列。

# PATCH_SUMMARY 2026-06-09 PDF Chinese Log Font MO Width
# - 維修紀錄區改用專用中文字型與較高列高，避免中文在小字級 PDF 影像輸出時被裁切或無法顯示。
# - MO Source PDF Note 欄縮小，其他欄位加寬，並縮小 MO Source 表格字體避免重疊。

# PATCH_SUMMARY 2026-06-09 Excel First PDF Export
# - PDF 匯出優先建立同名 Excel 檔，整理欄寬、列高與列印版面後由 Excel 轉存 PDF。
# - 若電腦未安裝 Excel 或 COM 匯出失敗，自動退回既有影像式 PDF 匯出。

# PATCH_SUMMARY 2026-06-09 Excel PDF Missing Item Rows Chinese MO Header
# - Excel PDF 補回特氣切換、烘烤部件/日期、CBr4 與兩列注意事項。
# - MO Source 表頭改回中文，並維持一筆資料橫向一列顯示。

# PATCH_SUMMARY 2026-06-09 MO Source Export Direct Grid Rows
# - PDF/Excel 匯出 MO Source 時直接從 GUI 表格逐列逐欄讀值，避免中間陣列傳參被拆成直欄。

param(
    [string]$OutputDirectory = "\\Tycba6\tyc_odd_pde\10-EPI_Growth_log\OtherFunctions\MaintenaceRecord\MaintenanceRecord"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# PATCH_SUMMARY v6-pdf-font-01
# - 低風險修改：僅調整 PDF 圖像輸出字型大小與文字框內距。
# - 不變動：GUI、欄位資料、鎖定流程、稽核紀錄、JSON/CSV/TXT 匯出邏輯。
# - 目的：修正 PDF 轉出後字體過大、表格內文字被上下裁切或顯示不完整。
# - 相容性：維持 Windows PowerShell 2.0-5.1 與 .NET Framework 可用語法。

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

function Resolve-FullPath {
    param([string]$Path)
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

function Escape-Json {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    $s = [string]$Value
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $s.Length; $i++) {
        $ch = $s[$i]
        switch ($ch) {
            '"' { [void]$sb.Append('\"') }
            '\' { [void]$sb.Append('\\') }
            "`b" { [void]$sb.Append('\b') }
            "`f" { [void]$sb.Append('\f') }
            "`n" { [void]$sb.Append('\n') }
            "`r" { [void]$sb.Append('\r') }
            "`t" { [void]$sb.Append('\t') }
            default {
                $code = [int][char]$ch
                if ($code -lt 32) {
                    [void]$sb.Append('\u' + $code.ToString('x4'))
                } else {
                    [void]$sb.Append($ch)
                }
            }
        }
    }
    return $sb.ToString()
}

function Quote-Csv {
    param([object]$Value)
    if ($null -eq $Value) { return '""' }
    $s = [string]$Value
    return '"' + $s.Replace('"', '""') + '"'
}

function Clean-FilePart {
    param([string]$Value)
    if ($null -eq $Value -or $Value.Trim().Length -eq 0) { return "NA" }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $chars = $Value.Trim().ToCharArray()
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $chars) {
        if ([Array]::IndexOf($invalid, $c) -ge 0) {
            [void]$sb.Append("_")
        } else {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString().Replace(" ", "_")
}

function Normalize-DateText {
    param([string]$Text)
    if ($null -eq $Text -or $Text.Trim().Length -eq 0) { return "" }
    $dt = New-Object System.DateTime
    if ([System.DateTime]::TryParse($Text.Trim(), [ref]$dt)) {
        return $dt.ToString("yyyy-MM-dd")
    }
    return $Text.Trim()
}

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$W)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($W, 23)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    return $label
}

function New-TextBox {
    param([string]$Text, [int]$X, [int]$Y, [int]$W)
    $box = New-Object System.Windows.Forms.TextBox
    $box.Text = $Text
    $box.Location = New-Object System.Drawing.Point($X, $Y)
    $box.Size = New-Object System.Drawing.Size($W, 23)
    return $box
}

function New-MultiTextBox {
    param([string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
    $box = New-Object System.Windows.Forms.TextBox
    $box.Text = $Text
    $box.Location = New-Object System.Drawing.Point($X, $Y)
    $box.Size = New-Object System.Drawing.Size($W, $H)
    $box.Multiline = $true
    $box.ScrollBars = "Vertical"
    return $box
}

function New-DataGrid {
    param([object[]]$Columns, [int]$X, [int]$Y, [int]$W, [int]$H)
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point($X, $Y)
    $grid.Size = New-Object System.Drawing.Size($W, $H)
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToDeleteRows = $true
    $grid.AutoSizeRowsMode = "None"
    $grid.RowHeadersWidth = 38
    $grid.SelectionMode = "CellSelect"
    $grid.MultiSelect = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = "FixedSingle"
    foreach ($colDef in $Columns) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = [string]$colDef[0]
        $col.Name = [string]$colDef[1]
        $col.Width = [int]$colDef[2]
        [void]$grid.Columns.Add($col)
    }
    return $grid
}

function Get-CollectionCount {
    param($Value)
    if ($null -eq $Value) { return 0 }
    if ($Value -is [string]) { return 1 }
    if ($Value -is [System.Array]) { return $Value.Length }
    if ($Value -is [System.Collections.ICollection]) { return $Value.Count }
    return 1
}

function Get-CollectionItem {
    param($Value, [int]$Index)
    if ($null -eq $Value) { return $null }
    if ($Index -lt 0) { return $null }
    if ($Value -is [System.Array]) {
        if ($Index -ge $Value.Length) { return $null }
        return $Value[$Index]
    }
    if ($Value -is [System.Collections.IList]) {
        if ($Index -ge $Value.Count) { return $null }
        return $Value[$Index]
    }
    if ($Index -eq 0) { return $Value }
    return $null
}

function Get-SafeRowCell {
    param($Row, [int]$Index)
    if ($null -eq $Row) { return "" }
    $count = Get-CollectionCount $Row
    if ($Index -lt 0 -or $Index -ge $count) { return "" }
    $value = Get-CollectionItem $Row $Index
    if ($null -eq $value) { return "" }
    return [string]$value
}

function Get-MoSourceColumnCount {
    return 12
}

function Get-MoSourceColumnDefs {
    return @(
        @("Source", "Source", 80),
        @("設定瓶壓", "SetPressure", 70),
        @("位置", "Position", 60),
        @("重量Off", "WeightOff", 70),
        @("重量On", "WeightOn", 70),
        @("重量T/W", "WeightTW", 70),
        @("來料瓶壓", "IncomingPressure", 75),
        @("實際瓶壓", "ActualPressure", 75),
        @("效率Off", "EfficiencyOff", 70),
        @("效率On", "EfficiencyOn", 70),
        @("退庫數量(瓶)", "ReturnBottleCount", 85),
        @("備註", "Note", 160)
    )
}

function Get-MoSourceKeyDefs {
    return @(
        @("source", "Source", "source"),
        @("set_pressure", "設定瓶壓", "set_pressure"),
        @("position", "位置", "position"),
        @("weight_off", "重量Off", "weight_off"),
        @("weight_on", "重量On", "weight_on"),
        @("weight_tw", "重量T/W", "weight_tw"),
        @("incoming_pressure", "來料瓶壓", "incoming_pressure"),
        @("actual_pressure", "實際瓶壓", "actual_pressure"),
        @("efficiency_off", "效率Off", "efficiency_off"),
        @("efficiency_on", "效率On", "efficiency_on"),
        @("return_bottle_count", "退庫數量(瓶)", "return_bottle_count"),
        @("note", "備註", "note")
    )
}

function Get-MoSourceRowText {
    param($Row)
    $parts = New-Object System.Collections.ArrayList
    $defs = Get-MoSourceKeyDefs
    for ($i = 0; $i -lt (Get-CollectionCount $defs); $i++) {
        $def = Get-CollectionItem $defs $i
        $label = Get-SafeRowCell $def 1
        $value = Get-SafeRowCell $Row $i
        if ($value.Length -gt 0) {
            [void]$parts.Add($label + "=" + $value)
        }
    }
    return ($parts -join "; ")
}

function Get-MoSourceHeaderTexts {
    $headers = New-Object System.Collections.ArrayList
    $defs = Get-MoSourceColumnDefs
    for ($i = 0; $i -lt (Get-CollectionCount $defs); $i++) {
        $def = Get-CollectionItem $defs $i
        [void]$headers.Add((Get-SafeRowCell $def 0))
    }
    return $headers
}

function Get-MoSourcePdfHeaderTexts {
    # PDF raster output in some Windows/GDI+ environments may fail to render
    # Chinese text in very narrow MO Source cells. Use ASCII short labels only
    # for the PDF table so the labels remain visible and aligned.
    $headers = New-Object System.Collections.ArrayList
    foreach ($h in @(
        "Source",
        "SetP",
        "Pos",
        "WtOff",
        "WtOn",
        "WtTW",
        "InP",
        "ActP",
        "EffOff",
        "EffOn",
        "Return",
        "Note"
    )) {
        [void]$headers.Add($h)
    }
    return $headers
}

function Test-HasMoSourceData {
    param($Rows)
    $rowCount = Get-CollectionCount $Rows
    for ($r = 0; $r -lt $rowCount; $r++) {
        $row = Get-CollectionItem $Rows $r
        for ($c = 0; $c -lt (Get-MoSourceColumnCount); $c++) {
            if ((Get-SafeRowCell $row $c).Length -gt 0) { return $true }
        }
    }
    return $false
}

function Convert-RowToObjectArray {
    param($Row, [int]$ColumnCount)
    $arr = New-Object 'object[]' $ColumnCount
    for ($i = 0; $i -lt $ColumnCount; $i++) {
        $arr[$i] = Get-SafeRowCell $Row $i
    }
    return ,$arr
}

function Add-MoSourceGridRow {
    param($Grid, $Row)
    $idx = $Grid.Rows.Add()
    for ($c = 0; $c -lt (Get-MoSourceColumnCount) -and $c -lt $Grid.Columns.Count; $c++) {
        $Grid.Rows[$idx].Cells[$c].Value = Get-SafeRowCell $Row $c
    }
}

function Get-ControlText {
    param($Control)
    if ($null -eq $Control) { return "" }
    $value = $Control.Text
    if ($null -eq $value) { return "" }
    return [string]$value
}

function Get-ControlTrim {
    param($Control)
    return (Get-ControlText $Control).Trim()
}

function Get-SelectedText {
    param($Control)
    if ($null -eq $Control) { return "" }
    if ($null -eq $Control.SelectedItem) { return "" }
    return [string]$Control.SelectedItem
}

function Add-GridRow {
    param($Grid, [object[]]$Values)
    $idx = $Grid.Rows.Add()
    $valueCount = Get-CollectionCount $Values
    for ($i = 0; $i -lt $valueCount -and $i -lt $Grid.Columns.Count; $i++) {
        $Grid.Rows[$idx].Cells[$i].Value = (Get-CollectionItem $Values $i)
    }
}

function Get-GridCellText {
    param($Grid, [int]$RowIndex, [int]$ColIndex)
    $v = $Grid.Rows[$RowIndex].Cells[$ColIndex].Value
    if ($null -eq $v) { return "" }
    return ([string]$v).Trim()
}

function Get-GridRows {
    param($Grid)
    $rows = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $Grid.Rows.Count; $i++) {
        $row = $Grid.Rows[$i]
        if ($row.IsNewRow) { continue }
        $values = New-Object System.Collections.ArrayList
        $hasValue = $false
        for ($c = 0; $c -lt $Grid.Columns.Count; $c++) {
            $text = Get-GridCellText $Grid $i $c
            if ($text.Length -gt 0) { $hasValue = $true }
            [void]$values.Add($text)
        }
        if ($hasValue) {
            [void]$rows.Add($values)
        }
    }
    return $rows
}

function Normalize-TableRows {
    param($Rows, [int]$ColumnCount, [int]$MinimumRows)
    $normalized = New-Object System.Collections.ArrayList
    $rowCount = Get-CollectionCount $Rows
    for ($i = 0; $i -lt $rowCount; $i++) {
        $sourceRow = Get-CollectionItem $Rows $i
        $values = New-Object System.Collections.ArrayList
        for ($c = 0; $c -lt $ColumnCount; $c++) {
            [void]$values.Add((Get-SafeRowCell $sourceRow $c))
        }
        [void]$normalized.Add($values)
    }
    while ((Get-CollectionCount $normalized) -lt $MinimumRows) {
        $blank = New-Object System.Collections.ArrayList
        for ($c = 0; $c -lt $ColumnCount; $c++) {
            [void]$blank.Add("")
        }
        [void]$normalized.Add($blank)
    }
    return $normalized
}

function Normalize-MoSourceRowsForPdf {
    param($Rows, [int]$MinimumRows)
    $normalized = New-Object System.Collections.ArrayList
    $columnCount = Get-MoSourceColumnCount
    $rowCount = Get-CollectionCount $Rows
    if ($rowCount -gt 0) {
        $first = Get-CollectionItem $Rows 0
        $firstIsRow = $false
        if ($first -is [System.Array]) { $firstIsRow = $true }
        if ($first -is [System.Collections.IList]) { $firstIsRow = $true }
        if (-not $firstIsRow) {
            $oneRow = New-Object System.Collections.ArrayList
            for ($c = 0; $c -lt $columnCount; $c++) {
                [void]$oneRow.Add((Get-SafeRowCell $Rows $c))
            }
            [void]$normalized.Add($oneRow)
        } else {
            for ($i = 0; $i -lt $rowCount; $i++) {
                $sourceRow = Get-CollectionItem $Rows $i
                $values = New-Object System.Collections.ArrayList
                for ($c = 0; $c -lt $columnCount; $c++) {
                    [void]$values.Add((Get-SafeRowCell $sourceRow $c))
                }
                [void]$normalized.Add($values)
            }
        }
    }
    while ((Get-CollectionCount $normalized) -lt $MinimumRows) {
        $blank = New-Object System.Collections.ArrayList
        for ($c = 0; $c -lt $columnCount; $c++) {
            [void]$blank.Add("")
        }
        [void]$normalized.Add($blank)
    }
    return ,$normalized
}

function Get-MoSourceRowsForExport {
    param([int]$MinimumRows)
    $rows = New-Object System.Collections.ArrayList
    if ($null -ne $script:gridMoSource) {
        for ($r = 0; $r -lt $script:gridMoSource.Rows.Count; $r++) {
            $gridRow = $script:gridMoSource.Rows[$r]
            if ($gridRow.IsNewRow) { continue }
            $values = New-Object System.Collections.ArrayList
            $hasValue = $false
            for ($c = 0; $c -lt (Get-MoSourceColumnCount); $c++) {
                $text = ""
                if ($c -lt $script:gridMoSource.Columns.Count) {
                    $value = $gridRow.Cells[$c].Value
                    if ($null -ne $value) { $text = [string]$value }
                }
                if ($text.Trim().Length -gt 0) { $hasValue = $true }
                [void]$values.Add($text)
            }
            if ($hasValue) { [void]$rows.Add($values) }
        }
    }
    while ((Get-CollectionCount $rows) -lt $MinimumRows) {
        $blank = New-Object System.Collections.ArrayList
        for ($c = 0; $c -lt (Get-MoSourceColumnCount); $c++) {
            [void]$blank.Add("")
        }
        [void]$rows.Add($blank)
    }
    return ,$rows
}

function Add-DefaultDiskRows {
    if ($null -eq $script:gridDiskSerials) { return }
    $script:gridDiskSerials.Rows.Clear()
    foreach ($pos in @("A", "B", "C", "D", "E", "F", "G")) {
        Add-GridRow $script:gridDiskSerials @($pos, "")
    }
}

function Get-DiskSerialRows {
    if ($null -eq $script:gridDiskSerials) { return (New-Object System.Collections.ArrayList) }
    return Normalize-TableRows (Get-GridRows $script:gridDiskSerials) 2 0
}

function Get-DiskSerialSummary {
    $parts = New-Object System.Collections.ArrayList
    $rows = Get-DiskSerialRows
    for ($i = 0; $i -lt (Get-CollectionCount $rows); $i++) {
        $row = Get-CollectionItem $rows $i
        $position = Get-SafeRowCell $row 0
        $serial = Get-SafeRowCell $row 1
        if ($position.Length -gt 0 -or $serial.Length -gt 0) {
            [void]$parts.Add(($position + ": " + $serial).Trim())
        }
    }
    return ($parts -join "; ")
}

function Set-DiskSerialRowsFromText {
    param([string]$Text)
    if ($null -eq $script:gridDiskSerials) { return }
    $script:gridDiskSerials.Rows.Clear()
    if ($null -eq $Text -or $Text.Trim().Length -eq 0) {
        Add-DefaultDiskRows
        return
    }
    $items = $Text.Split(@(";"), [System.StringSplitOptions]::RemoveEmptyEntries)
    foreach ($item in $items) {
        $part = $item.Trim()
        if ($part.Length -eq 0) { continue }
        $idx = $part.IndexOf(":")
        if ($idx -ge 0) {
            Add-GridRow $script:gridDiskSerials @($part.Substring(0, $idx).Trim(), $part.Substring($idx + 1).Trim())
        } else {
            Add-GridRow $script:gridDiskSerials @("", $part)
        }
    }
    if ($script:gridDiskSerials.Rows.Count -eq 1) {
        Add-DefaultDiskRows
    }
}

function New-ComboBox {
    param([string[]]$Items, [string]$Text, [int]$X, [int]$Y, [int]$W)
    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = "DropDownList"
    [void]$combo.Items.AddRange($Items)
    $combo.Location = New-Object System.Drawing.Point($X, $Y)
    $combo.Size = New-Object System.Drawing.Size($W, 23)
    if ($Text.Length -gt 0 -and $combo.Items.Contains($Text)) {
        $combo.SelectedItem = $Text
    } elseif ($combo.Items.Count -gt 0) {
        $combo.SelectedIndex = 0
    }
    return $combo
}

function New-Field {
    param($Rows, [string]$Section, [string]$Field, [string]$Value, [string]$Unit, [string]$Note, [string]$Key)
    $row = New-Object PSObject
    Add-Member -InputObject $row -MemberType NoteProperty -Name section -Value $Section
    Add-Member -InputObject $row -MemberType NoteProperty -Name field -Value $Field
    Add-Member -InputObject $row -MemberType NoteProperty -Name value -Value $Value
    Add-Member -InputObject $row -MemberType NoteProperty -Name unit -Value $Unit
    Add-Member -InputObject $row -MemberType NoteProperty -Name note -Value $Note
    Add-Member -InputObject $row -MemberType NoteProperty -Name ai_key -Value $Key
    [void]$Rows.Add($row)
}

function Prompt-AuditReason {
    param([string]$Title)
    $reason = [Microsoft.VisualBasic.Interaction]::InputBox("請輸入本次管制修改原因。此原因會寫入稽核紀錄。", $Title, "")
    if ($null -eq $reason) { return "" }
    return $reason.Trim()
}

function Get-AuditUser {
    $u = [System.Environment]::UserName
    $d = [System.Environment]::UserDomainName
    if ($d -ne $null -and $d.Length -gt 0) { return $d + "\" + $u }
    return $u
}

function Get-AuditPath {
    $outDir = Resolve-FullPath (Get-ControlText $script:txtOutputDir)
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir | Out-Null
    }
    $auditDir = Join-Path $outDir "audit"
    if (-not (Test-Path $auditDir)) {
        New-Item -ItemType Directory -Path $auditDir | Out-Null
    }
    return (Join-Path $auditDir ("maintenance_audit_" + (Get-Date -Format "yyyyMMdd") + ".csv"))
}

function Write-Audit {
    param([string]$Action, [string]$Section, [string]$Field, [string]$OldValue, [string]$NewValue, [string]$Reason)
    $path = Get-AuditPath
    if (-not (Test-Path $path)) {
        $header = ((Quote-Csv "timestamp"), (Quote-Csv "user"), (Quote-Csv "machine"), (Quote-Csv "record_id"), (Quote-Csv "action"), (Quote-Csv "section"), (Quote-Csv "field"), (Quote-Csv "old_value"), (Quote-Csv "new_value"), (Quote-Csv "reason")) -join ","
        [System.IO.File]::AppendAllText($path, $header + "`r`n", [System.Text.Encoding]::UTF8)
    }
    $line = @(
        (Quote-Csv (Get-Date -Format "yyyy-MM-ddTHH:mm:sszzz")),
        (Quote-Csv (Get-AuditUser)),
        (Quote-Csv ([System.Environment]::MachineName)),
        (Quote-Csv (Get-RecordBaseName)),
        (Quote-Csv $Action),
        (Quote-Csv $Section),
        (Quote-Csv $Field),
        (Quote-Csv $OldValue),
        (Quote-Csv $NewValue),
        (Quote-Csv $Reason)
    ) -join ","
    [System.IO.File]::AppendAllText($path, $line + "`r`n", [System.Text.Encoding]::UTF8)
}

function New-SnapshotEntry {
    param([string]$Key, [string]$Value)
    $row = New-Object PSObject
    Add-Member -InputObject $row -MemberType NoteProperty -Name key -Value $Key
    Add-Member -InputObject $row -MemberType NoteProperty -Name value -Value $Value
    return $row
}

function Get-ItemsSnapshot {
    $rows = New-Object System.Collections.ArrayList
    $itemRows = Get-MaintenanceItemRows
    foreach ($r in $itemRows) {
        [void]$rows.Add((New-SnapshotEntry ("item." + [string]$r[0]) ([string]$r[1] + "|" + [string]$r[2])))
    }
    return $rows
}

function Get-RecordsSnapshot {
    $rows = New-Object System.Collections.ArrayList
    $logRows = Normalize-TableRows (Get-GridRows $script:gridMaintenanceLog) 3 0
    for ($i = 0; $i -lt (Get-CollectionCount $logRows); $i++) {
        $logRow = Get-CollectionItem $logRows $i
        [void]$rows.Add((New-SnapshotEntry ("maintenance_log." + ($i + 1)) ((Get-SafeRowCell $logRow 0) + "|" + (Get-SafeRowCell $logRow 1) + "|" + (Get-SafeRowCell $logRow 2))))
    }
    $moRowsWrapper = Get-MoSourceRowsForExport 0
    $moRows = $moRowsWrapper
    for ($i = 0; $i -lt (Get-CollectionCount $moRows); $i++) {
        $moRow = Get-CollectionItem $moRows $i
        [void]$rows.Add((New-SnapshotEntry ("mo_source." + ($i + 1)) (Get-MoSourceRowText $moRow)))
    }
    return $rows
}

function Get-SnapshotValue {
    param($Snapshot, [string]$Key)
    foreach ($row in $Snapshot) {
        if ($row.key -eq $Key) { return [string]$row.value }
    }
    return ""
}

function Write-SnapshotDiffAudit {
    param([string]$Section, $OldSnapshot, $NewSnapshot, [string]$Reason)
    $keys = New-Object System.Collections.ArrayList
    foreach ($row in $OldSnapshot) {
        if (-not $keys.Contains($row.key)) { [void]$keys.Add($row.key) }
    }
    foreach ($row in $NewSnapshot) {
        if (-not $keys.Contains($row.key)) { [void]$keys.Add($row.key) }
    }
    $changeCount = 0
    foreach ($key in $keys) {
        $oldValue = Get-SnapshotValue $OldSnapshot $key
        $newValue = Get-SnapshotValue $NewSnapshot $key
        if ($oldValue -ne $newValue) {
            Write-Audit "CHANGE" $Section $key $oldValue $newValue $Reason
            $changeCount++
        }
    }
    if ($changeCount -eq 0) {
        Write-Audit "NO_CHANGE" $Section "" "" "" $Reason
    }
}

function Get-CheckboxValue {
    param($CheckBox)
    if ($CheckBox.Checked) { return "V" }
    return ""
}

function Get-PmTopic {
    $parts = New-Object System.Collections.ArrayList
    if ($script:chkDisk.Checked) { [void]$parts.Add("Disk") }
    if ($script:chkCover.Checked) { [void]$parts.Add("Cover") }
    if ($script:chkStar.Checked) { [void]$parts.Add("Star") }
    if ($script:chkCeiling.Checked) { [void]$parts.Add("Ceiling") }
    if ($parts.Count -eq 0) { return "" }
    return "Change " + ($parts -join "/")
}

function Refresh-PmTopic {
    $script:txtPmTopic.Text = Get-PmTopic
}

function Get-RecordDateCode {
    $pmDate = Normalize-DateText $script:txtPmDate.Text
    $datePart = Clean-FilePart $pmDate
    if ($datePart -ne "NA") { $datePart = $datePart.Replace("-", "") }
    return ("MR-" + $datePart)
}

function Get-FullRecordCode {
    $dateCode = Get-RecordDateCode
    $toolPart = Clean-FilePart $script:txtToolName.Text
    $runPart = Clean-FilePart $script:txtRunId.Text
    return ($dateCode + "-" + $toolPart + "-" + $runPart)
}

function Get-RecordCode {
    if ($script:recordCodeFinalized -and $script:finalRecordCode.Length -gt 0) {
        return $script:finalRecordCode
    }
    return Get-RecordDateCode
}

function Refresh-RecordCode {
    if ($null -ne $script:txtRecordCode) {
        $script:txtRecordCode.Text = Get-RecordCode
    }
}

function Finalize-RecordCode {
    $oldCode = Get-RecordCode
    $script:finalRecordCode = Get-FullRecordCode
    $script:recordCodeFinalized = $true
    Refresh-RecordCode
    if ($oldCode -ne $script:finalRecordCode) {
        Write-Audit "FINALIZE_RECORD_CODE" "record" "record_code" $oldCode $script:finalRecordCode "維修項目完成鎖定後產生完整編碼"
    }
}

function New-MaintenanceRecord {
    $script:itemEditReason = ""
    $script:recordEditReason = ""
    $script:itemBaselineSnapshot = New-Object System.Collections.ArrayList
    $script:recordBaselineSnapshot = New-Object System.Collections.ArrayList
    $script:itemLocked = $false
    $script:recordLocked = $false
    $script:recordCodeFinalized = $false
    $script:finalRecordCode = ""

    Set-ItemsLocked $false
    Set-RecordsLocked $false

    $script:cmbPmCategory.SelectedItem = "Scheduled PM"
    $script:txtBatch.Text = ""
    $script:txtProduct.Text = ""
    $script:txtRunId.Text =""
    $script:txtMeetingPeople.Text = ""
    $script:txtPmDate.Text = (Get-Date -Format "yyyy-MM-dd")
    $script:cmbShift.SelectedItem = ""
    $script:txtToolName.Text = ""
    $script:txtMaintPeople.Text = ""
    $script:txtMaintDate.Text = ""
    $script:txtAttachmentPaths.Text = ""

    $script:chkDisk.Checked = $false
    $script:chkCover.Checked = $false
    $script:chkStar.Checked = $false
    $script:txtCoverSerial.Text = ""
    $script:txtStarSerial.Text = ""
    $script:chkCeiling.Checked = $false
    $script:txtCeiling.Text = ""
    Add-DefaultDiskRows
    $script:chkPullDown.Checked = $false
    $script:txtPullDown.Text = ""
    $script:chkSusceptor.Checked = $false
    $script:txtSusceptor.Text = ""
    $script:chkCollectorRing.Checked = $false
    $script:txtCollectorRing.Text = ""
    $script:chkExhaustTube.Checked = $false
    $script:txtExhaustTube.Text = ""
    $script:chkLightpipe.Checked = $false
    $script:txtLightpipe.Text = ""
    $script:chkThrottleValve.Checked = $false
    $script:txtThrottleValve.Text = ""
    $script:chkMaintPump.Checked = $false
    $script:txtMaintPump.Text = ""
    $script:chkCv25.Checked = $false
    $script:txtCv25.Text = ""
    $script:chkGloveboxPump.Checked = $false
    $script:txtGloveboxPump.Text = ""
    $script:chkCt1000.Checked = $false
    $script:txtCt1000.Text = ""
    $script:chkDorPump.Checked = $false
    $script:txtDorPump.Text = ""
    $script:chkYTube.Checked = $false
    $script:txtYTube.Text = ""
    $script:chkPTrap.Checked = $false
    $script:txtPTrap.Text = ""
    $script:chkReactorValve.Checked = $false
    $script:txtReactorValve.Text = ""
    $script:chkExhaustPipe.Checked = $false
    $script:txtExhaustPipe.Text = ""
    $script:chkDorORing.Checked = $false
    $script:txtDorORing.Text = ""
    $script:cmbSpecialGas.SelectedIndex = 0
    $script:txtSusceptorBakeDate.Text = ""
    $script:txtDiskBakeDate.Text = ""
    $script:txtCeilingBakeDate.Text = ""
    $script:txtCbr4Bottle1.Text = ""
    $script:txtCbr4Bottle2.Text = ""
    $script:txtOtherItems.Text = ""

    $script:gridMaintenanceLog.Rows.Clear()
    $script:gridMoSource.Rows.Clear()
    $script:txtPreview.Text = ""

    Refresh-PmTopic
    Refresh-RecordCode
    Write-Audit "NEW_RECORD" "record" "record_code" "" (Get-RecordCode) "新建紀錄"
}

function Build-Fields {
    $rows = New-Object System.Collections.ArrayList
    $pmDate = Normalize-DateText $script:txtPmDate.Text
    $maintDate = Normalize-DateText $script:txtMaintDate.Text
    $topic = Get-PmTopic

    New-Field $rows "maintenance_header" "維修紀錄編碼" (Get-RecordCode) "" "PM日期、機台名稱、RunID 組成" "record_code"
    New-Field $rows "maintenance_header" "機台名稱" $script:txtToolName.Text.Trim() "" "" "tool_name"
    New-Field $rows "maintenance_header" "PM主題" $topic "" "由 GUI 維修項目自動組成" "pm_topic"
    New-Field $rows "maintenance_header" "PM類別" ([string]$script:cmbPmCategory.SelectedItem) "" "" "pm_category"
    New-Field $rows "maintenance_header" "Batch" $script:txtBatch.Text.Trim() "" "" "batch"
    New-Field $rows "maintenance_header" "產品" $script:txtProduct.Text.Trim() "" "" "product"
    New-Field $rows "maintenance_header" "RunID" $script:txtRunId.Text.Trim() "" "" "run_id"
    New-Field $rows "maintenance_header" "PM會議人員" $script:txtMeetingPeople.Text.Trim() "" "" "pm_meeting_people"
    New-Field $rows "maintenance_header" "PM日期" $pmDate "" "ISO date preferred" "pm_date"
    New-Field $rows "maintenance_header" "班別" ([string]$script:cmbShift.SelectedItem) "" "" "shift"
    New-Field $rows "maintenance_attachment" "附件/照片路徑" $script:txtAttachmentPaths.Text.Trim() "" "可填內網路徑、資料夾或多個檔案路徑" "attachment_paths"
    New-Field $rows "maintenance_confirm" "維修人員" $script:txtMaintPeople.Text.Trim() "" "" "maintenance_people"
    New-Field $rows "maintenance_confirm" "維修日期" $maintDate "" "ISO date preferred" "maintenance_date"

    New-Field $rows "maintenance_item" "Star" (Get-CheckboxValue $script:chkStar) "" "" "item_star"
    New-Field $rows "maintenance_item_serial" "Star 序號" $script:txtStarSerial.Text.Trim() "" "" "item_star_serial"
    New-Field $rows "maintenance_item" "Cover" (Get-CheckboxValue $script:chkCover) "" "" "item_cover"
    New-Field $rows "maintenance_item_serial" "Cover 序號" $script:txtCoverSerial.Text.Trim() "" "" "item_cover_serial"
    New-Field $rows "maintenance_item" "Ceiling" (Get-CheckboxValue $script:chkCeiling) "" "" "item_ceiling"
    New-Field $rows "maintenance_item_serial" "Ceiling 序號" $script:txtCeiling.Text.Trim() "" "" "item_ceiling_serial"
    New-Field $rows "maintenance_item" "Pull down" (Get-CheckboxValue $script:chkPullDown) "" "" "item_pull_down"
    New-Field $rows "maintenance_item_serial" "Pull down 序號" $script:txtPullDown.Text.Trim() "" "" "item_pull_down_serial"
    New-Field $rows "maintenance_item" "Susceptor" (Get-CheckboxValue $script:chkSusceptor) "" "" "item_susceptor"
    New-Field $rows "maintenance_item_serial" "Susceptor 序號" $script:txtSusceptor.Text.Trim() "" "" "item_susceptor_serial"
    New-Field $rows "maintenance_item" "Collector ring" (Get-CheckboxValue $script:chkCollectorRing) "" "" "item_collector_ring"
    New-Field $rows "maintenance_item_serial" "Collector ring 序號" $script:txtCollectorRing.Text.Trim() "" "" "item_collector_ring_serial"
    New-Field $rows "maintenance_item" "Disk/SN" (Get-CheckboxValue $script:chkDisk) "" "" "item_disk_sn"
    $diskRows = Get-DiskSerialRows
    for ($i = 0; $i -lt (Get-CollectionCount $diskRows); $i++) {
        $diskRow = Get-CollectionItem $diskRows $i
        $position = Get-SafeRowCell $diskRow 0
        $serial = Get-SafeRowCell $diskRow 1
        New-Field $rows "maintenance_item_serial" ("Disk/SN " + ($i + 1)) $serial "" ("位置=" + $position) ("item_disk_sn_" + ($i + 1))
    }
    New-Field $rows "maintenance_item" "Exhaust tube" (Get-CheckboxValue $script:chkExhaustTube) "" "" "item_exhaust_tube"
    New-Field $rows "maintenance_item_serial" "Exhaust tube 序號" $script:txtExhaustTube.Text.Trim() "" "" "item_exhaust_tube_serial"
    New-Field $rows "maintenance_item" "Lightpipe" (Get-CheckboxValue $script:chkLightpipe) "" "" "item_lightpipe"
    New-Field $rows "maintenance_item_serial" "Lightpipe 序號" $script:txtLightpipe.Text.Trim() "" "" "item_lightpipe_serial"
    New-Field $rows "maintenance_item" "Throttle Valve" (Get-CheckboxValue $script:chkThrottleValve) "" "" "item_throttle_valve"
    New-Field $rows "maintenance_item_serial" "Throttle Valve 序號" $script:txtThrottleValve.Text.Trim() "" "" "item_throttle_valve_serial"
    New-Field $rows "maintenance_item" "Maint pump" (Get-CheckboxValue $script:chkMaintPump) "" "" "item_maint_pump"
    New-Field $rows "maintenance_item_serial" "Maint pump 序號" $script:txtMaintPump.Text.Trim() "" "" "item_maint_pump_serial"
    New-Field $rows "maintenance_item" "25mbar C.V." (Get-CheckboxValue $script:chkCv25) "" "" "item_25mbar_cv"
    New-Field $rows "maintenance_item_serial" "25mbar C.V. 序號" $script:txtCv25.Text.Trim() "" "" "item_25mbar_cv_serial"
    New-Field $rows "maintenance_item" "Glovebox pump" (Get-CheckboxValue $script:chkGloveboxPump) "" "" "item_glovebox_pump"
    New-Field $rows "maintenance_item_serial" "Glovebox pump 序號" $script:txtGloveboxPump.Text.Trim() "" "" "item_glovebox_pump_serial"
    New-Field $rows "maintenance_item" "CT1000" (Get-CheckboxValue $script:chkCt1000) "" "" "item_ct1000"
    New-Field $rows "maintenance_item_serial" "CT1000 序號" $script:txtCt1000.Text.Trim() "" "" "item_ct1000_serial"
    New-Field $rows "maintenance_item" "DOR pump" (Get-CheckboxValue $script:chkDorPump) "" "" "item_dor_pump"
    New-Field $rows "maintenance_item_serial" "DOR pump 序號" $script:txtDorPump.Text.Trim() "" "" "item_dor_pump_serial"
    New-Field $rows "maintenance_item" "Y-tube" (Get-CheckboxValue $script:chkYTube) "" "" "item_y_tube"
    New-Field $rows "maintenance_item_serial" "Y-tube 序號" $script:txtYTube.Text.Trim() "" "" "item_y_tube_serial"
    New-Field $rows "maintenance_item" "P-trap" (Get-CheckboxValue $script:chkPTrap) "" "" "item_p_trap"
    New-Field $rows "maintenance_item_serial" "P-trap 序號" $script:txtPTrap.Text.Trim() "" "" "item_p_trap_serial"
    New-Field $rows "maintenance_item" "Reactor valve" (Get-CheckboxValue $script:chkReactorValve) "" "" "item_reactor_valve"
    New-Field $rows "maintenance_item_serial" "Reactor valve 序號" $script:txtReactorValve.Text.Trim() "" "" "item_reactor_valve_serial"
    New-Field $rows "maintenance_item" "Exhaust pipe" (Get-CheckboxValue $script:chkExhaustPipe) "" "" "item_exhaust_pipe"
    New-Field $rows "maintenance_item_serial" "Exhaust pipe 序號" $script:txtExhaustPipe.Text.Trim() "" "" "item_exhaust_pipe_serial"
    New-Field $rows "maintenance_item" "DOR O-ring" (Get-CheckboxValue $script:chkDorORing) "" "" "item_dor_o_ring"
    New-Field $rows "maintenance_item_serial" "DOR O-ring 序號" $script:txtDorORing.Text.Trim() "" "" "item_dor_o_ring_serial"
    New-Field $rows "maintenance_item" "特氣是否切換" ([string]$script:cmbSpecialGas.SelectedItem) "" "" "item_special_gas_switched"
    New-Field $rows "maintenance_item" "Susceptor烘烤日期" (Normalize-DateText $script:txtSusceptorBakeDate.Text) "" "" "item_susceptor_bake_date"
    New-Field $rows "maintenance_item" "Disk烘烤日期" (Normalize-DateText $script:txtDiskBakeDate.Text) "" "" "item_disk_bake_date"
    New-Field $rows "maintenance_item" "Ceiling(G4)烘烤日期" (Normalize-DateText $script:txtCeilingBakeDate.Text) "" "" "item_ceiling_g4_bake_date"
    New-Field $rows "maintenance_item" "CBr4第一瓶重量(g)" $script:txtCbr4Bottle1.Text.Trim() "g" "" "item_cbr4_bottle_1_weight_g"
    New-Field $rows "maintenance_item" "CBr4第二瓶重量(g)" $script:txtCbr4Bottle2.Text.Trim() "g" "" "item_cbr4_bottle_2_weight_g"
    New-Field $rows "maintenance_note" "注意事項" "確認下一 Batch 如 CBr4 第一瓶重量會 < 滿瓶的 30%，需下機執行敲打作業" "" "" "note_cbr4_bottle_1"
    New-Field $rows "maintenance_note" "注意事項" "MFC default 設定值若小於 10 sccm，請一律設定為 10 sccm" "" "" "note_mfc_default"
    New-Field $rows "maintenance_item" "其他維修項目" $script:txtOtherItems.Text.Trim() "" "" "item_other"

    $moRows = Normalize-TableRows (Get-GridRows $script:gridMoSource) (Get-MoSourceColumnCount) 0
    $moDefs = Get-MoSourceKeyDefs
    for ($i = 0; $i -lt (Get-CollectionCount $moRows); $i++) {
        $moRow = Get-CollectionItem $moRows $i
        $rowNo = ($i + 1)
        New-Field $rows "mo_source" ("MO Source " + $rowNo + " 摘要") (Get-MoSourceRowText $moRow) "" "MO SOURCE 更換紀錄整列摘要" ("mo_source_" + $rowNo + "_summary")
        for ($c = 0; $c -lt (Get-CollectionCount $moDefs); $c++) {
            $def = Get-CollectionItem $moDefs $c
            $fieldKey = Get-SafeRowCell $def 0
            $fieldLabel = Get-SafeRowCell $def 1
            New-Field $rows "mo_source" ("MO Source " + $rowNo + " " + $fieldLabel) (Get-SafeRowCell $moRow $c) "" "MO SOURCE 更換紀錄欄位" ("mo_source_" + $rowNo + "_" + $fieldKey)
        }
    }

    $logRows = Normalize-TableRows (Get-GridRows $script:gridMaintenanceLog) 3 0
    for ($i = 0; $i -lt (Get-CollectionCount $logRows); $i++) {
        $logRow = Get-CollectionItem $logRows $i
        $date = Normalize-DateText (Get-SafeRowCell $logRow 0)
        $content = Get-SafeRowCell $logRow 1
        $note = Get-SafeRowCell $logRow 2
        New-Field $rows "maintenance_log" ([string]($i + 1)) $content "" ($date + " " + $note).Trim() ("maintenance_log_" + ($i + 1))
    }
    return $rows
}

function Write-RecordFiles {
    param($Rows)
    Ensure-NoOpenControlledEdit
    Ensure-FinalLocked
    $outDir = Resolve-FullPath (Get-ControlText $script:txtOutputDir)
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir | Out-Null
    }

    $base = Get-RecordBaseName
    $jsonPath = Join-Path $outDir ($base + "_ai.json")
    $csvPath = Join-Path $outDir ($base + "_ai_fields.csv")
    $txtPath = Join-Path $outDir ($base + ".txt")

    $jsonLines = New-Object System.Collections.ArrayList
    [void]$jsonLines.Add("{")
    [void]$jsonLines.Add('  "schema_version": "mocvd-maintenance-record/v1",')
    [void]$jsonLines.Add('  "generated_from": "MaintenanceRecordGui.ps1",')
    [void]$jsonLines.Add('  "generated_at_local": "' + (Get-Date -Format "yyyy-MM-ddTHH:mm:sszzz") + '",')
    [void]$jsonLines.Add('  "record_id": "' + (Escape-Json $base) + '",')
    [void]$jsonLines.Add('  "fields": [')
    $rowCount = Get-CollectionCount $Rows
    for ($i = 0; $i -lt $rowCount; $i++) {
        $r = $Rows[$i]
        $suffix = ","
        if ($i -eq ($rowCount - 1)) { $suffix = "" }
        [void]$jsonLines.Add("    {")
        [void]$jsonLines.Add('      "section": "' + (Escape-Json $r.section) + '",')
        [void]$jsonLines.Add('      "field": "' + (Escape-Json $r.field) + '",')
        [void]$jsonLines.Add('      "value": "' + (Escape-Json $r.value) + '",')
        [void]$jsonLines.Add('      "unit": "' + (Escape-Json $r.unit) + '",')
        [void]$jsonLines.Add('      "note": "' + (Escape-Json $r.note) + '",')
        [void]$jsonLines.Add('      "ai_key": "' + (Escape-Json $r.ai_key) + '"')
        [void]$jsonLines.Add("    }$suffix")
    }
    [void]$jsonLines.Add("  ]")
    [void]$jsonLines.Add("}")

    $csvLines = New-Object System.Collections.ArrayList
    [void]$csvLines.Add(((Quote-Csv "section"), (Quote-Csv "field"), (Quote-Csv "value"), (Quote-Csv "unit"), (Quote-Csv "note"), (Quote-Csv "ai_key")) -join ",")
    foreach ($r in $Rows) {
        [void]$csvLines.Add(((Quote-Csv $r.section), (Quote-Csv $r.field), (Quote-Csv $r.value), (Quote-Csv $r.unit), (Quote-Csv $r.note), (Quote-Csv $r.ai_key)) -join ",")
    }

    $txtLines = New-Object System.Collections.ArrayList
    [void]$txtLines.Add("===維修紀錄表===")
    [void]$txtLines.Add("維修紀錄編碼:" + (Get-RecordCode))
    [void]$txtLines.Add("機台名稱:" + $script:txtToolName.Text.Trim())
    [void]$txtLines.Add("PM主題:" + (Get-PmTopic))
    [void]$txtLines.Add("PM類別:" + [string]$script:cmbPmCategory.SelectedItem)
    [void]$txtLines.Add("Batch:" + $script:txtBatch.Text.Trim())
    [void]$txtLines.Add("產品:" + $script:txtProduct.Text.Trim())
    [void]$txtLines.Add("RunID:" + $script:txtRunId.Text.Trim())
    [void]$txtLines.Add("PM會議人員:" + $script:txtMeetingPeople.Text.Trim())
    [void]$txtLines.Add("PM日期/班別:" + (Normalize-DateText $script:txtPmDate.Text) + " (" + [string]$script:cmbShift.SelectedItem + ")")
    [void]$txtLines.Add("附件/照片路徑:" + $script:txtAttachmentPaths.Text.Trim())
    [void]$txtLines.Add("")
    [void]$txtLines.Add("===AI欄位===")
    foreach ($r in $Rows) {
        [void]$txtLines.Add($r.ai_key + "=" + $r.value)
    }

    [System.IO.File]::WriteAllText($jsonPath, ($jsonLines -join "`r`n"), [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($csvPath, ($csvLines -join "`r`n"), [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($txtPath, ($txtLines -join "`r`n"), [System.Text.Encoding]::UTF8)
    Write-CompletedRecordIndex $jsonPath $csvPath $txtPath
    Write-Audit "EXPORT_AI" "record" "" "" (($jsonPath + "; " + $csvPath + "; " + $txtPath)) ""
    return @($jsonPath, $csvPath, $txtPath)
}

function Assert-AiOutputRequiredFields {
    if ($script:txtBatch.Text.Trim().Length -eq 0) {
        throw "Batch 不可空白，不能儲存 AI 格式。"
    }
    if ($script:txtRunId.Text.Trim().Length -eq 0) {
        throw "RunID 不可空白，不能儲存 AI 格式。"
    }
    if ($script:txtToolName.Text.Trim().Length -eq 0) {
        throw "機台名稱不可空白，不能儲存 AI 格式。"
    }
}

function Update-AiPreview {
    Refresh-PmTopic
    $rows = Build-Fields
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("schema_version=mocvd-maintenance-record/v1")
    [void]$lines.Add("field_count=" + (Get-CollectionCount $rows))
    [void]$lines.Add("")
    foreach ($r in $rows) {
        [void]$lines.Add($r.section + " | " + $r.ai_key + " | " + $r.field + " = " + $r.value)
    }
    $script:txtPreview.Text = $lines -join "`r`n"
    return $rows
}

function Save-AiOutput {
    Assert-AiOutputRequiredFields
    $rows = Update-AiPreview
    $paths = Write-RecordFiles $rows
    $script:txtPreview.Text = "已輸出：" + "`r`n" + ($paths -join "`r`n") + "`r`n`r`n" + $script:txtPreview.Text
    return $paths
}

function Get-RecordBaseName {
    return "維修紀錄_" + (Get-RecordCode)
}

function Get-CompletedIndexPath {
    $outDir = Resolve-FullPath $script:txtOutputDir.Text
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir | Out-Null
    }
    return (Join-Path $outDir "maintenance_records_index.csv")
}

function Write-CompletedRecordIndex {
    param([string]$JsonPath, [string]$CsvPath, [string]$TxtPath)
    $path = Get-CompletedIndexPath
    if (-not (Test-Path $path)) {
        $header = ((Quote-Csv "record_code"), (Quote-Csv "pm_date"), (Quote-Csv "tool_name"), (Quote-Csv "batch"), (Quote-Csv "run_id"), (Quote-Csv "product"), (Quote-Csv "pm_topic"), (Quote-Csv "json_path"), (Quote-Csv "csv_path"), (Quote-Csv "txt_path"), (Quote-Csv "completed_at"), (Quote-Csv "user")) -join ","
        [System.IO.File]::AppendAllText($path, $header + "`r`n", [System.Text.Encoding]::UTF8)
    }
    $line = @(
        (Quote-Csv (Get-RecordCode)),
        (Quote-Csv (Normalize-DateText $script:txtPmDate.Text)),
        (Quote-Csv $script:txtToolName.Text.Trim()),
        (Quote-Csv $script:txtBatch.Text.Trim()),
        (Quote-Csv $script:txtRunId.Text.Trim()),
        (Quote-Csv $script:txtProduct.Text.Trim()),
        (Quote-Csv (Get-PmTopic)),
        (Quote-Csv $JsonPath),
        (Quote-Csv $CsvPath),
        (Quote-Csv $TxtPath),
        (Quote-Csv (Get-Date -Format "yyyy-MM-ddTHH:mm:sszzz")),
        (Quote-Csv (Get-AuditUser))
    ) -join ","
    [System.IO.File]::AppendAllText($path, $line + "`r`n", [System.Text.Encoding]::UTF8)
}

function Clear-GridColumnsAndRows {
    param($Grid)
    $Grid.Rows.Clear()
    $Grid.Columns.Clear()
}

function Add-GridColumnsFromHeaders {
    param($Grid, [object[]]$Headers)
    foreach ($h in $Headers) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = [string]$h
        $col.Name = [string]$h
        $col.Width = 140
        [void]$Grid.Columns.Add($col)
    }
}

function Read-CsvFileRows {
    param([string]$Path)
    $result = New-Object System.Collections.ArrayList
    if (-not (Test-Path $Path)) { return $result }
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [System.Text.Encoding]::UTF8)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(",")
        while (-not $parser.EndOfData) {
            [void]$result.Add($parser.ReadFields())
        }
    }
    finally {
        $parser.Close()
    }
    return $result
}

function Load-CsvFilesToGrid {
    param($Grid, [string[]]$Paths, [string]$Filter)
    Clear-GridColumnsAndRows $Grid
    $headers = $null
    $filterText = ""
    if ($null -ne $Filter) { $filterText = $Filter.Trim().ToLowerInvariant() }
    foreach ($path in $Paths) {
        $rows = Read-CsvFileRows $path
        if ((Get-CollectionCount $rows) -eq 0) { continue }
        if ($null -eq $headers) {
            $headers = $rows[0]
            Add-GridColumnsFromHeaders $Grid $headers
        }
        for ($i = 1; $i -lt (Get-CollectionCount $rows); $i++) {
            $lineText = ($rows[$i] -join " ").ToLowerInvariant()
            if ($filterText.Length -gt 0 -and -not $lineText.Contains($filterText)) { continue }
            $idx = $Grid.Rows.Add()
            $rowValues = $rows[$i]
            for ($c = 0; $c -lt (Get-CollectionCount $rowValues) -and $c -lt $Grid.Columns.Count; $c++) {
                $Grid.Rows[$idx].Cells[$c].Value = (Get-CollectionItem $rowValues $c)
            }
        }
    }
}

function Export-GridToCsv {
    param($Grid, [string]$Path)
    $lines = New-Object System.Collections.ArrayList
    $headers = New-Object System.Collections.ArrayList
    for ($c = 0; $c -lt $Grid.Columns.Count; $c++) {
        [void]$headers.Add((Quote-Csv $Grid.Columns[$c].HeaderText))
    }
    [void]$lines.Add(($headers -join ","))
    for ($r = 0; $r -lt $Grid.Rows.Count; $r++) {
        if ($Grid.Rows[$r].IsNewRow) { continue }
        $cells = New-Object System.Collections.ArrayList
        for ($c = 0; $c -lt $Grid.Columns.Count; $c++) {
            $value = $Grid.Rows[$r].Cells[$c].Value
            [void]$cells.Add((Quote-Csv $value))
        }
        [void]$lines.Add(($cells -join ","))
    }
    [System.IO.File]::WriteAllText($Path, ($lines -join "`r`n"), [System.Text.Encoding]::UTF8)
}

function Get-GridSelectedRowIndex {
    param($Grid)
    if ($null -eq $Grid) { return -1 }
    if ($Grid.SelectedRows.Count -gt 0) {
        return $Grid.SelectedRows[0].Index
    }
    if ($Grid.CurrentCell -ne $null) {
        return $Grid.CurrentCell.RowIndex
    }
    return -1
}

function Get-GridCellTextByColumn {
    param($Grid, [int]$RowIndex, [string]$ColumnName)
    if ($null -eq $Grid -or $RowIndex -lt 0 -or $RowIndex -ge $Grid.Rows.Count) { return "" }
    if (-not $Grid.Columns.Contains($ColumnName)) { return "" }
    $value = $Grid.Rows[$RowIndex].Cells[$ColumnName].Value
    if ($null -eq $value) { return "" }
    return [string]$value
}

function Open-CompletedRecordReadOnlyByRow {
    param([int]$RowIndex)
    if ($RowIndex -lt 0 -or $RowIndex -ge $script:gridCompletedRecords.Rows.Count) {
        throw "請先選取一筆已完成維修紀錄。"
    }
    if ($script:gridCompletedRecords.Rows[$RowIndex].IsNewRow) {
        throw "請先選取一筆已完成維修紀錄。"
    }
    $recordCode = Get-GridCellTextByColumn $script:gridCompletedRecords $RowIndex "record_code"
    $txtPath = Get-GridCellTextByColumn $script:gridCompletedRecords $RowIndex "txt_path"
    $csvPath = Get-GridCellTextByColumn $script:gridCompletedRecords $RowIndex "csv_path"
    $jsonPath = Get-GridCellTextByColumn $script:gridCompletedRecords $RowIndex "json_path"
    $readPath = ""
    if ($txtPath.Length -gt 0 -and (Test-Path $txtPath)) {
        $readPath = $txtPath
    } elseif ($csvPath.Length -gt 0 -and (Test-Path $csvPath)) {
        $readPath = $csvPath
    } elseif ($jsonPath.Length -gt 0 -and (Test-Path $jsonPath)) {
        $readPath = $jsonPath
    } else {
        throw ("找不到可讀取的完成紀錄檔案：" + "`r`nTXT: " + $txtPath + "`r`nCSV: " + $csvPath + "`r`nJSON: " + $jsonPath)
    }

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("=== 完成紀錄唯讀檢視 ===")
    [void]$lines.Add("維修紀錄編碼: " + $recordCode)
    [void]$lines.Add("讀取檔案: " + $readPath)
    [void]$lines.Add("")
    [void]$lines.AddRange([System.IO.File]::ReadAllLines($readPath, [System.Text.Encoding]::UTF8))
    $script:txtPreview.Text = ($lines -join "`r`n")
    $script:txtPreview.ReadOnly = $true
    $tabs.SelectedTab = $tabOutput
    Write-Audit "VIEW_COMPLETED_READONLY" "completed_record" "record_code" "" $recordCode $readPath
}

function Open-SelectedCompletedRecordReadOnly {
    $rowIndex = Get-GridSelectedRowIndex $script:gridCompletedRecords
    Open-CompletedRecordReadOnlyByRow $rowIndex
}

function Get-CompletedRecordSourceDirectory {
    param([int]$RowIndex)
    foreach ($colName in @("txt_path", "csv_path", "json_path")) {
        $path = Get-GridCellTextByColumn $script:gridCompletedRecords $RowIndex $colName
        if ($path.Length -gt 0) {
            $dir = [System.IO.Path]::GetDirectoryName($path)
            if ($dir.Length -gt 0 -and (Test-Path $dir)) { return $dir }
        }
    }
    return (Resolve-FullPath $script:txtOutputDir.Text)
}

function Find-CompletedRecordPdfByRow {
    param([int]$RowIndex)
    if ($RowIndex -lt 0 -or $RowIndex -ge $script:gridCompletedRecords.Rows.Count) {
        throw "請先選取一筆已完成維修紀錄。"
    }
    if ($script:gridCompletedRecords.Rows[$RowIndex].IsNewRow) {
        throw "請先選取一筆已完成維修紀錄。"
    }
    $recordCode = Get-GridCellTextByColumn $script:gridCompletedRecords $RowIndex "record_code"
    $pdfPath = Get-GridCellTextByColumn $script:gridCompletedRecords $RowIndex "pdf_path"
    if ($pdfPath.Length -gt 0 -and (Test-Path $pdfPath)) { return $pdfPath }

    $dir = Get-CompletedRecordSourceDirectory $RowIndex
    $pattern = "維修紀錄_" + $recordCode + "*完整維修紀錄.pdf"
    $files = Get-ChildItem -Path $dir -Filter $pattern | Where-Object { -not $_.PSIsContainer } | Sort-Object LastWriteTime -Descending
    foreach ($file in $files) {
        return $file.FullName
    }

    $fallbackPattern = "*" + $recordCode + "*.pdf"
    $fallbackFiles = Get-ChildItem -Path $dir -Filter $fallbackPattern | Where-Object { -not $_.PSIsContainer } | Sort-Object LastWriteTime -Descending
    foreach ($file in $fallbackFiles) {
        return $file.FullName
    }
    throw ("找不到此完成紀錄的 PDF。" + "`r`n紀錄編碼：" + $recordCode + "`r`n搜尋資料夾：" + $dir + "`r`n請確認該筆紀錄已匯出完整維修 PDF。")
}

function Open-SelectedCompletedRecordPdf {
    $rowIndex = Get-GridSelectedRowIndex $script:gridCompletedRecords
    $path = Find-CompletedRecordPdfByRow $rowIndex
    Open-FilePreview $path
    Write-Audit "OPEN_COMPLETED_PDF" "completed_record" "pdf_path" "" $path ""
}

function Get-DraftDirectory {
    $outDir = Resolve-FullPath $script:txtOutputDir.Text
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir | Out-Null
    }
    $draftDir = Join-Path $outDir "drafts"
    if (-not (Test-Path $draftDir)) {
        New-Item -ItemType Directory -Path $draftDir | Out-Null
    }
    return $draftDir
}

function Get-DefaultDraftPath {
    $draftDir = Get-DraftDirectory
    return (Join-Path $draftDir ((Get-RecordBaseName) + "_draft.csv"))
}

function Get-AutoDraftDirectory {
    $draftDir = Get-DraftDirectory
    $autoDir = Join-Path $draftDir "autosave"
    if (-not (Test-Path $autoDir)) {
        New-Item -ItemType Directory -Path $autoDir | Out-Null
    }
    return $autoDir
}

function Get-AutoDraftPath {
    $autoDir = Get-AutoDraftDirectory
    return (Join-Path $autoDir ((Get-RecordBaseName) + "_autosave.draft.csv"))
}

function Get-DraftFileRows {
    $rows = New-Object System.Collections.ArrayList
    $draftDir = Get-DraftDirectory
    if (-not (Test-Path $draftDir)) { return $rows }
    $files = Get-ChildItem -Path $draftDir -Filter "*draft.csv" | Where-Object { -not $_.PSIsContainer } | Sort-Object LastWriteTime -Descending
    foreach ($file in $files) {
        $kb = [Math]::Round(($file.Length / 1024), 1)
        [void]$rows.Add(@($file.Name, $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $kb.ToString(), $file.FullName))
    }
    return $rows
}

function Refresh-DraftList {
    if ($null -eq $script:gridDraftList) { return }
    $script:gridDraftList.Rows.Clear()
    $rows = Get-DraftFileRows
    for ($i = 0; $i -lt (Get-CollectionCount $rows); $i++) {
        $row = Get-CollectionItem $rows $i
        Add-GridRow $script:gridDraftList @((Get-SafeRowCell $row 0), (Get-SafeRowCell $row 1), (Get-SafeRowCell $row 2), (Get-SafeRowCell $row 3))
    }
    if ($null -ne $script:lblDraftListStatus) {
        $countText = (Get-CollectionCount $rows).ToString()
        $script:lblDraftListStatus.Text = "草稿列表：共 " + $countText + " 筆"
        if ((Get-CollectionCount $rows) -eq 0) {
            $script:lblDraftListStatus.Text = "草稿列表：共 0 筆，資料夾：" + (Get-DraftDirectory)
        }
    }
}

function Get-SelectedDraftPath {
    if ($null -eq $script:gridDraftList) { return "" }
    if ($script:gridDraftList.SelectedCells.Count -eq 0) { return "" }
    $rowIndex = $script:gridDraftList.SelectedCells[0].RowIndex
    if ($rowIndex -lt 0 -or $rowIndex -ge $script:gridDraftList.Rows.Count) { return "" }
    $row = $script:gridDraftList.Rows[$rowIndex]
    if ($row.IsNewRow) { return "" }
    $value = $row.Cells["Path"].Value
    if ($null -eq $value) { return "" }
    return ([string]$value).Trim()
}

function Open-SelectedDraft {
    $path = Get-SelectedDraftPath
    if ($path.Length -eq 0) {
        throw "請先在草稿列表選取一筆草稿。"
    }
    $answer = [System.Windows.Forms.MessageBox]::Show("讀取草稿會覆蓋目前畫面內容。是否繼續？", "讀取草稿", "YesNo", "Warning")
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    Load-DraftFile $path
    $tabs.SelectedTab = $tabBasic
    Refresh-DraftList
    [System.Windows.Forms.MessageBox]::Show(("草稿已讀取：" + "`r`n" + $path), "完成", "OK", "Information") | Out-Null
}

function Delete-SelectedDraft {
    $path = Get-SelectedDraftPath
    if ($path.Length -eq 0) {
        throw "請先在草稿列表選取要刪除的草稿。"
    }
    if (-not (Test-Path $path)) {
        Refresh-DraftList
        throw "草稿檔已不存在：" + $path
    }
    $answer = [System.Windows.Forms.MessageBox]::Show(("確定刪除此草稿？`r`n`r`n" + $path), "刪除草稿", "YesNo", "Warning")
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    Remove-Item -Path $path -Force
    if ($script:lastDraftPath -eq $path) { $script:lastDraftPath = "" }
    Write-Audit "DELETE_DRAFT" "draft" "" $path "" "刪除不再使用的草稿"
    Refresh-DraftList
}

function Add-DraftLine {
    param($Lines, [string]$Section, [string]$Key, [string]$Value)
    [void]$Lines.Add(((Quote-Csv $Section), (Quote-Csv $Key), (Quote-Csv $Value)) -join ",")
}

function Add-DraftMoSourceRow {
    param($Lines, [int]$RowIndex, $MoRow)
    $cells = New-Object System.Collections.ArrayList
    [void]$cells.Add((Quote-Csv "mo_source_row"))
    [void]$cells.Add((Quote-Csv ("row" + $RowIndex)))
    [void]$cells.Add((Quote-Csv ""))
    for ($c = 0; $c -lt (Get-MoSourceColumnCount); $c++) {
        [void]$cells.Add((Quote-Csv (Get-SafeRowCell $MoRow $c)))
    }
    [void]$Lines.Add(($cells -join ","))
}

function Set-ComboValue {
    param($Combo, [string]$Value)
    if ($null -eq $Value) { return }
    if ($Combo.Items.Contains($Value)) {
        $Combo.SelectedItem = $Value
    } elseif ($Combo.Items.Count -gt 0) {
        $Combo.SelectedIndex = 0
    }
}

function Save-DraftFile {
    param([string]$Path, [bool]$WriteAudit)
    if ($PSBoundParameters.ContainsKey("WriteAudit") -eq $false) { $WriteAudit = $true }
    Refresh-PmTopic
    Refresh-RecordCode
    $lines = New-Object System.Collections.ArrayList
    # 草稿採寬表頭，讓 MO SOURCE 更換紀錄可以一列一筆橫向保存；前三欄保留給原本 key/value 草稿格式。
    [void]$lines.Add(((Quote-Csv "section"), (Quote-Csv "key"), (Quote-Csv "value"), (Quote-Csv "Source"), (Quote-Csv "設定瓶壓"), (Quote-Csv "位置"), (Quote-Csv "重量Off"), (Quote-Csv "重量On"), (Quote-Csv "重量T/W"), (Quote-Csv "來料瓶壓"), (Quote-Csv "實際瓶壓"), (Quote-Csv "效率Off"), (Quote-Csv "效率On"), (Quote-Csv "退庫數量(瓶)"), (Quote-Csv "備註")) -join ",")

    Add-DraftLine $lines "state" "item_locked" ([string]$script:itemLocked)
    Add-DraftLine $lines "state" "record_locked" ([string]$script:recordLocked)
    Add-DraftLine $lines "state" "record_code_finalized" ([string]$script:recordCodeFinalized)
    Add-DraftLine $lines "state" "final_record_code" $script:finalRecordCode
    Add-DraftLine $lines "state" "item_edit_reason" $script:itemEditReason
    Add-DraftLine $lines "state" "record_edit_reason" $script:recordEditReason

    Add-DraftLine $lines "basic" "pm_category" ([string]$script:cmbPmCategory.SelectedItem)
    Add-DraftLine $lines "basic" "batch" $script:txtBatch.Text
    Add-DraftLine $lines "basic" "product" $script:txtProduct.Text
    Add-DraftLine $lines "basic" "run_id" $script:txtRunId.Text
    Add-DraftLine $lines "basic" "meeting_people" $script:txtMeetingPeople.Text
    Add-DraftLine $lines "basic" "pm_date" $script:txtPmDate.Text
    Add-DraftLine $lines "basic" "shift" ([string]$script:cmbShift.SelectedItem)
    Add-DraftLine $lines "basic" "tool_name" $script:txtToolName.Text
    Add-DraftLine $lines "basic" "attachment_paths" $script:txtAttachmentPaths.Text
    Add-DraftLine $lines "basic" "maint_people" $script:txtMaintPeople.Text
    Add-DraftLine $lines "basic" "maint_date" $script:txtMaintDate.Text

    Add-DraftLine $lines "items" "disk_checked" ([string]$script:chkDisk.Checked)
    Add-DraftLine $lines "items" "cover_checked" ([string]$script:chkCover.Checked)
    Add-DraftLine $lines "items" "cover_serial" $script:txtCoverSerial.Text
    Add-DraftLine $lines "items" "star_checked" ([string]$script:chkStar.Checked)
    Add-DraftLine $lines "items" "star_serial" $script:txtStarSerial.Text
    Add-DraftLine $lines "items" "ceiling_checked" ([string]$script:chkCeiling.Checked)
    Add-DraftLine $lines "items" "ceiling" $script:txtCeiling.Text
    Add-DraftLine $lines "items" "disk_serials" (Get-DiskSerialSummary)
    $diskRows = Get-DiskSerialRows
    for ($i = 0; $i -lt (Get-CollectionCount $diskRows); $i++) {
        $diskRow = Get-CollectionItem $diskRows $i
        Add-DraftLine $lines "disk_serial" ("row" + $i + ".position") (Get-SafeRowCell $diskRow 0)
        Add-DraftLine $lines "disk_serial" ("row" + $i + ".serial") (Get-SafeRowCell $diskRow 1)
    }
    Add-DraftLine $lines "items" "pull_down_checked" ([string]$script:chkPullDown.Checked)
    Add-DraftLine $lines "items" "pull_down" $script:txtPullDown.Text
    Add-DraftLine $lines "items" "susceptor_checked" ([string]$script:chkSusceptor.Checked)
    Add-DraftLine $lines "items" "susceptor" $script:txtSusceptor.Text
    Add-DraftLine $lines "items" "collector_ring_checked" ([string]$script:chkCollectorRing.Checked)
    Add-DraftLine $lines "items" "collector_ring" $script:txtCollectorRing.Text
    Add-DraftLine $lines "items" "exhaust_tube_checked" ([string]$script:chkExhaustTube.Checked)
    Add-DraftLine $lines "items" "exhaust_tube" $script:txtExhaustTube.Text
    Add-DraftLine $lines "items" "lightpipe_checked" ([string]$script:chkLightpipe.Checked)
    Add-DraftLine $lines "items" "lightpipe" $script:txtLightpipe.Text
    Add-DraftLine $lines "items" "throttle_valve_checked" ([string]$script:chkThrottleValve.Checked)
    Add-DraftLine $lines "items" "throttle_valve" $script:txtThrottleValve.Text
    Add-DraftLine $lines "items" "maint_pump_checked" ([string]$script:chkMaintPump.Checked)
    Add-DraftLine $lines "items" "maint_pump" $script:txtMaintPump.Text
    Add-DraftLine $lines "items" "cv25_checked" ([string]$script:chkCv25.Checked)
    Add-DraftLine $lines "items" "cv25" $script:txtCv25.Text
    Add-DraftLine $lines "items" "glovebox_pump_checked" ([string]$script:chkGloveboxPump.Checked)
    Add-DraftLine $lines "items" "glovebox_pump" $script:txtGloveboxPump.Text
    Add-DraftLine $lines "items" "ct1000_checked" ([string]$script:chkCt1000.Checked)
    Add-DraftLine $lines "items" "ct1000" $script:txtCt1000.Text
    Add-DraftLine $lines "items" "dor_pump_checked" ([string]$script:chkDorPump.Checked)
    Add-DraftLine $lines "items" "dor_pump" $script:txtDorPump.Text
    Add-DraftLine $lines "items" "y_tube_checked" ([string]$script:chkYTube.Checked)
    Add-DraftLine $lines "items" "y_tube" $script:txtYTube.Text
    Add-DraftLine $lines "items" "p_trap_checked" ([string]$script:chkPTrap.Checked)
    Add-DraftLine $lines "items" "p_trap" $script:txtPTrap.Text
    Add-DraftLine $lines "items" "reactor_valve_checked" ([string]$script:chkReactorValve.Checked)
    Add-DraftLine $lines "items" "reactor_valve" $script:txtReactorValve.Text
    Add-DraftLine $lines "items" "exhaust_pipe_checked" ([string]$script:chkExhaustPipe.Checked)
    Add-DraftLine $lines "items" "exhaust_pipe" $script:txtExhaustPipe.Text
    Add-DraftLine $lines "items" "dor_o_ring_checked" ([string]$script:chkDorORing.Checked)
    Add-DraftLine $lines "items" "dor_o_ring" $script:txtDorORing.Text
    Add-DraftLine $lines "items" "special_gas" ([string]$script:cmbSpecialGas.SelectedItem)
    Add-DraftLine $lines "items" "susceptor_bake_date" $script:txtSusceptorBakeDate.Text
    Add-DraftLine $lines "items" "disk_bake_date" $script:txtDiskBakeDate.Text
    Add-DraftLine $lines "items" "ceiling_bake_date" $script:txtCeilingBakeDate.Text
    Add-DraftLine $lines "items" "cbr4_bottle1" $script:txtCbr4Bottle1.Text
    Add-DraftLine $lines "items" "cbr4_bottle2" $script:txtCbr4Bottle2.Text
    Add-DraftLine $lines "items" "other_items" $script:txtOtherItems.Text

    $logRows = Normalize-TableRows (Get-GridRows $script:gridMaintenanceLog) 3 0
    for ($i = 0; $i -lt (Get-CollectionCount $logRows); $i++) {
        $logRow = Get-CollectionItem $logRows $i
        Add-DraftLine $lines "maintenance_log" ("row" + $i + ".date") (Get-SafeRowCell $logRow 0)
        Add-DraftLine $lines "maintenance_log" ("row" + $i + ".content") (Get-SafeRowCell $logRow 1)
        Add-DraftLine $lines "maintenance_log" ("row" + $i + ".note") (Get-SafeRowCell $logRow 2)
    }

    $moRows = Normalize-TableRows (Get-GridRows $script:gridMoSource) (Get-MoSourceColumnCount) 0
    for ($i = 0; $i -lt (Get-CollectionCount $moRows); $i++) {
        $moRow = $moRows[$i]
        Add-DraftMoSourceRow -Lines $lines -RowIndex $i -MoRow $moRow
    }

    [System.IO.File]::WriteAllText($Path, ($lines -join "`r`n"), [System.Text.Encoding]::UTF8)
    $script:lastDraftPath = $Path
    if ($WriteAudit) {
        Write-Audit "SAVE_DRAFT" "draft" "" "" $Path "未完成維修紀錄暫存"
    }
}

function Load-DraftFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        throw "找不到草稿檔：" + $Path
    }
    $rows = Read-CsvFileRows $Path
    if ((Get-CollectionCount $rows) -lt 2) {
        throw "草稿檔沒有可讀取內容。"
    }
    $map = @{}
    $moSourceDraftRows = New-Object System.Collections.ArrayList
    for ($i = 1; $i -lt (Get-CollectionCount $rows); $i++) {
        $r = $rows[$i]
        if ((Get-CollectionCount $r) -lt 3) { continue }
        $sectionName = [string]$r[0]
        $rowKey = [string]$r[1]
        $map[($sectionName + "." + $rowKey)] = [string]$r[2]

        if ($sectionName -eq "mo_source_row") {
            $values = New-Object System.Collections.ArrayList
            $hasMoValue = $false
            for ($c = 0; $c -lt (Get-MoSourceColumnCount); $c++) {
                $sourceIndex = $c + 3
                $value = ""
                if ((Get-CollectionCount $r) -gt $sourceIndex) { $value = [string](Get-CollectionItem $r $sourceIndex) }
                if ($value.Length -gt 0) { $hasMoValue = $true }
                [void]$values.Add($value)
            }
            if ($hasMoValue) {
                $rowArray = New-Object 'object[]' (Get-MoSourceColumnCount)
                for ($vc = 0; $vc -lt (Get-MoSourceColumnCount); $vc++) {
                    $rowArray[$vc] = Get-SafeRowCell $values $vc
                }
                [void]$moSourceDraftRows.Add($rowArray)
            }
        }
    }

    $script:itemLocked = $false
    $script:recordLocked = $false
    Set-ItemsLocked $false
    Set-RecordsLocked $false

    Set-ComboValue $script:cmbPmCategory $map["basic.pm_category"]
    $script:txtBatch.Text = $map["basic.batch"]
    $script:txtProduct.Text = $map["basic.product"]
    $script:txtRunId.Text = $map["basic.run_id"]
    $script:txtMeetingPeople.Text = $map["basic.meeting_people"]
    $script:txtPmDate.Text = $map["basic.pm_date"]
    Set-ComboValue $script:cmbShift $map["basic.shift"]
    $script:txtToolName.Text = $map["basic.tool_name"]
    $script:txtAttachmentPaths.Text = $map["basic.attachment_paths"]
    $script:txtMaintPeople.Text = $map["basic.maint_people"]
    $script:txtMaintDate.Text = $map["basic.maint_date"]

    $script:chkDisk.Checked = ([string]$map["items.disk_checked"] -eq "True")
    $script:chkCover.Checked = ([string]$map["items.cover_checked"] -eq "True")
    $script:txtCoverSerial.Text = $map["items.cover_serial"]
    $script:chkStar.Checked = ([string]$map["items.star_checked"] -eq "True")
    $script:txtStarSerial.Text = $map["items.star_serial"]
    $script:chkCeiling.Checked = ([string]$map["items.ceiling_checked"] -eq "True")
    $script:txtCeiling.Text = $map["items.ceiling"]
    $script:gridDiskSerials.Rows.Clear()
    $loadedDiskRows = $false
    for ($i = 0; $i -lt 200; $i++) {
        $positionKey = "disk_serial.row" + $i + ".position"
        $serialKey = "disk_serial.row" + $i + ".serial"
        if (-not $map.ContainsKey($positionKey) -and -not $map.ContainsKey($serialKey)) { continue }
        Add-GridRow $script:gridDiskSerials @($map[$positionKey], $map[$serialKey])
        $loadedDiskRows = $true
    }
    if (-not $loadedDiskRows) {
        Set-DiskSerialRowsFromText $map["items.disk_serials"]
    }
    $script:chkPullDown.Checked = ([string]$map["items.pull_down_checked"] -eq "True")
    $script:txtPullDown.Text = $map["items.pull_down"]
    $script:chkSusceptor.Checked = ([string]$map["items.susceptor_checked"] -eq "True")
    $script:txtSusceptor.Text = $map["items.susceptor"]
    $script:chkCollectorRing.Checked = ([string]$map["items.collector_ring_checked"] -eq "True")
    $script:txtCollectorRing.Text = $map["items.collector_ring"]
    $script:chkExhaustTube.Checked = ([string]$map["items.exhaust_tube_checked"] -eq "True")
    $script:txtExhaustTube.Text = $map["items.exhaust_tube"]
    $script:chkLightpipe.Checked = ([string]$map["items.lightpipe_checked"] -eq "True")
    $script:txtLightpipe.Text = $map["items.lightpipe"]
    $script:chkThrottleValve.Checked = ([string]$map["items.throttle_valve_checked"] -eq "True")
    $script:txtThrottleValve.Text = $map["items.throttle_valve"]
    $script:chkMaintPump.Checked = ([string]$map["items.maint_pump_checked"] -eq "True")
    $script:txtMaintPump.Text = $map["items.maint_pump"]
    $script:chkCv25.Checked = ([string]$map["items.cv25_checked"] -eq "True")
    $script:txtCv25.Text = $map["items.cv25"]
    $script:chkGloveboxPump.Checked = ([string]$map["items.glovebox_pump_checked"] -eq "True")
    $script:txtGloveboxPump.Text = $map["items.glovebox_pump"]
    $script:chkCt1000.Checked = ([string]$map["items.ct1000_checked"] -eq "True")
    $script:txtCt1000.Text = $map["items.ct1000"]
    $script:chkDorPump.Checked = ([string]$map["items.dor_pump_checked"] -eq "True")
    $script:txtDorPump.Text = $map["items.dor_pump"]
    $script:chkYTube.Checked = ([string]$map["items.y_tube_checked"] -eq "True")
    $script:txtYTube.Text = $map["items.y_tube"]
    $script:chkPTrap.Checked = ([string]$map["items.p_trap_checked"] -eq "True")
    $script:txtPTrap.Text = $map["items.p_trap"]
    $script:chkReactorValve.Checked = ([string]$map["items.reactor_valve_checked"] -eq "True")
    $script:txtReactorValve.Text = $map["items.reactor_valve"]
    $script:chkExhaustPipe.Checked = ([string]$map["items.exhaust_pipe_checked"] -eq "True")
    $script:txtExhaustPipe.Text = $map["items.exhaust_pipe"]
    $script:chkDorORing.Checked = ([string]$map["items.dor_o_ring_checked"] -eq "True")
    $script:txtDorORing.Text = $map["items.dor_o_ring"]
    Set-ComboValue $script:cmbSpecialGas $map["items.special_gas"]
    $script:txtSusceptorBakeDate.Text = $map["items.susceptor_bake_date"]
    if ($script:txtSusceptorBakeDate.Text.Length -eq 0) { $script:txtSusceptorBakeDate.Text = $map["items.gas_susceptor"] }
    $script:txtDiskBakeDate.Text = $map["items.disk_bake_date"]
    if ($script:txtDiskBakeDate.Text.Length -eq 0) { $script:txtDiskBakeDate.Text = $map["items.gas_disk"] }
    $script:txtCeilingBakeDate.Text = $map["items.ceiling_bake_date"]
    $script:txtCbr4Bottle1.Text = $map["items.cbr4_bottle1"]
    $script:txtCbr4Bottle2.Text = $map["items.cbr4_bottle2"]
    $script:txtOtherItems.Text = $map["items.other_items"]

    $script:gridMaintenanceLog.Rows.Clear()
    for ($i = 0; $i -lt 200; $i++) {
        $dateKey = "maintenance_log.row" + $i + ".date"
        $contentKey = "maintenance_log.row" + $i + ".content"
        $noteKey = "maintenance_log.row" + $i + ".note"
        if (-not $map.ContainsKey($dateKey) -and -not $map.ContainsKey($contentKey) -and -not $map.ContainsKey($noteKey)) { continue }
        Add-GridRow $script:gridMaintenanceLog @($map[$dateKey], $map[$contentKey], $map[$noteKey])
    }

    $script:gridMoSource.Rows.Clear()
    if ((Get-CollectionCount $moSourceDraftRows) -gt 0) {
        for ($i = 0; $i -lt (Get-CollectionCount $moSourceDraftRows); $i++) {
            $moRow = $moSourceDraftRows[$i]
            Add-MoSourceGridRow -Grid $script:gridMoSource -Row $moRow
        }
    } else {
        # Backward compatibility: 讀取舊版直式草稿 mo_source.rowN.field。
        $moDefs = Get-MoSourceKeyDefs
        for ($i = 0; $i -lt 200; $i++) {
            $values = New-Object System.Collections.ArrayList
            $hasMoValue = $false
            for ($c = 0; $c -lt (Get-CollectionCount $moDefs); $c++) {
                $def = Get-CollectionItem $moDefs $c
                $key = "mo_source.row" + $i + "." + (Get-SafeRowCell $def 2)
                $value = ""
                if ($map.ContainsKey($key)) { $value = $map[$key] }
                if ($value.Length -gt 0) { $hasMoValue = $true }
                [void]$values.Add($value)
            }
            # Backward compatibility: 更舊草稿只有 source / pressure / note 三欄。
            $oldSourceKey = "mo_source.row" + $i + ".source"
            $oldPressureKey = "mo_source.row" + $i + ".pressure"
            $oldNoteKey = "mo_source.row" + $i + ".note"
            if ($map.ContainsKey($oldSourceKey) -and $values[0].Length -eq 0) {
                $values[0] = $map[$oldSourceKey]
                if ($values[0].Length -gt 0) { $hasMoValue = $true }
            }
            if ($map.ContainsKey($oldPressureKey) -and $values[1].Length -eq 0) {
                $values[1] = $map[$oldPressureKey]
                if ($values[1].Length -gt 0) { $hasMoValue = $true }
            }
            if ($map.ContainsKey($oldNoteKey) -and $values[11].Length -eq 0) {
                $values[11] = $map[$oldNoteKey]
                if ($values[11].Length -gt 0) { $hasMoValue = $true }
            }
            if ($hasMoValue) { Add-MoSourceGridRow -Grid $script:gridMoSource -Row $values }
        }
    }

    $script:itemEditReason = $map["state.item_edit_reason"]
    $script:recordEditReason = $map["state.record_edit_reason"]
    $script:recordCodeFinalized = ([string]$map["state.record_code_finalized"] -eq "True")
    $script:finalRecordCode = $map["state.final_record_code"]
    $script:itemLocked = ([string]$map["state.item_locked"] -eq "True")
    $script:recordLocked = ([string]$map["state.record_locked"] -eq "True")
    $script:itemBaselineSnapshot = Get-ItemsSnapshot
    $script:recordBaselineSnapshot = Get-RecordsSnapshot
    Set-ItemsLocked $script:itemLocked
    Set-RecordsLocked $script:recordLocked
    Refresh-PmTopic
    Refresh-RecordCode
    $script:lastDraftPath = $Path
    Write-Audit "LOAD_DRAFT" "draft" "" "" $Path "讀取未完成維修紀錄草稿"
}

function Add-ValidationError {
    param($Errors, [string]$Text)
    if ($Text.Length -gt 0) { [void]$Errors.Add($Text) }
}

function Assert-BasicRequiredFields {
    $errors = New-Object System.Collections.ArrayList
    if ((Get-ControlTrim $script:txtToolName).Length -eq 0) { Add-ValidationError $errors "機台名稱不可空白。" }
    if ((Get-ControlTrim $script:txtBatch).Length -eq 0) { Add-ValidationError $errors "Batch 不可空白。" }
    if ((Get-ControlTrim $script:txtRunId).Length -eq 0) { Add-ValidationError $errors "RunID 不可空白。" }
    if ((Normalize-DateText (Get-ControlText $script:txtPmDate)).Length -eq 0) { Add-ValidationError $errors "PM日期不可空白。" }
    if ((Get-CollectionCount $errors) -gt 0) {
        throw ("必填欄位不足：" + "`r`n- " + ($errors -join "`r`n- "))
    }
}

function Assert-ItemLockReady {
    Assert-BasicRequiredFields
    $errors = New-Object System.Collections.ArrayList
    $hasItem = $false
    foreach ($chk in @(
        $script:chkDisk, $script:chkCover, $script:chkStar, $script:chkCeiling,
        $script:chkPullDown, $script:chkSusceptor, $script:chkCollectorRing,
        $script:chkExhaustTube, $script:chkLightpipe, $script:chkThrottleValve,
        $script:chkMaintPump, $script:chkCv25, $script:chkGloveboxPump,
        $script:chkCt1000, $script:chkDorPump, $script:chkYTube, $script:chkPTrap,
        $script:chkReactorValve, $script:chkExhaustPipe, $script:chkDorORing
    )) {
        if ($null -ne $chk -and $chk.Checked) { $hasItem = $true }
    }
    if (-not $hasItem -and (Get-ControlTrim $script:txtOtherItems).Length -eq 0) {
        Add-ValidationError $errors "至少需勾選一項維修項目，或填寫其他維修項目。"
    }
    if ($script:chkCover.Checked -and (Get-ControlTrim $script:txtCoverSerial).Length -eq 0) { Add-ValidationError $errors "Cover 已勾選，請填寫 Cover 序號。" }
    if ($script:chkStar.Checked -and (Get-ControlTrim $script:txtStarSerial).Length -eq 0) { Add-ValidationError $errors "Star 已勾選，請填寫 Star 序號。" }
    if ($script:chkCeiling.Checked -and (Get-ControlTrim $script:txtCeiling).Length -eq 0) { Add-ValidationError $errors "Ceiling 已勾選，請填寫 Ceiling 序號。" }
    if ($script:chkDisk.Checked) {
        $diskRows = Get-DiskSerialRows
        $hasDiskSerial = $false
        for ($i = 0; $i -lt (Get-CollectionCount $diskRows); $i++) {
            $diskRow = Get-CollectionItem $diskRows $i
            if ((Get-SafeRowCell $diskRow 0).Length -gt 0 -and (Get-SafeRowCell $diskRow 1).Length -gt 0) {
                $hasDiskSerial = $true
            }
        }
        if (-not $hasDiskSerial) { Add-ValidationError $errors "Disk/SN 已勾選，請至少填寫一筆 Disk 位置與序號。" }
    }
    if ((Get-CollectionCount $errors) -gt 0) {
        throw ("維修項目鎖定前檢查未通過：" + "`r`n- " + ($errors -join "`r`n- "))
    }
}

function Assert-RecordLockReady {
    Assert-BasicRequiredFields
    $errors = New-Object System.Collections.ArrayList
    if (-not $script:itemLocked) { Add-ValidationError $errors "維修項目尚未完成鎖定。" }
    if ((Get-ControlTrim $script:txtMaintPeople).Length -eq 0) { Add-ValidationError $errors "維修人員不可空白。" }
    if ((Normalize-DateText (Get-ControlText $script:txtMaintDate)).Length -eq 0) { Add-ValidationError $errors "維修日期不可空白。" }

    $moRows = Normalize-TableRows (Get-GridRows $script:gridMoSource) (Get-MoSourceColumnCount) 0
    for ($i = 0; $i -lt (Get-CollectionCount $moRows); $i++) {
        $moRow = Get-CollectionItem $moRows $i
        $hasValue = $false
        for ($c = 0; $c -lt (Get-MoSourceColumnCount); $c++) {
            if ((Get-SafeRowCell $moRow $c).Length -gt 0) { $hasValue = $true }
        }
        if ($hasValue -and (Get-SafeRowCell $moRow 0).Length -eq 0) {
            Add-ValidationError $errors ("MO Source 第 " + ($i + 1).ToString() + " 列已有資料，Source 不可空白。")
        }
    }
    if ((Get-CollectionCount $errors) -gt 0) {
        throw ("維修紀錄鎖定前檢查未通過：" + "`r`n- " + ($errors -join "`r`n- "))
    }
}

function Get-AttachmentPathList {
    $result = New-Object System.Collections.ArrayList
    $text = Get-ControlText $script:txtAttachmentPaths
    if ($text.Trim().Length -eq 0) { return $result }
    $parts = $text -split "[;`r`n]+"
    foreach ($part in $parts) {
        $p = [string]$part
        $p = $p.Trim()
        $p = $p.Trim('"')
        if ($p.Length -gt 0) { [void]$result.Add($p) }
    }
    return $result
}

function Get-MissingAttachmentPaths {
    $missing = New-Object System.Collections.ArrayList
    $paths = Get-AttachmentPathList
    for ($i = 0; $i -lt (Get-CollectionCount $paths); $i++) {
        $path = Get-CollectionItem $paths $i
        if (-not (Test-Path $path)) {
            [void]$missing.Add($path)
        }
    }
    return $missing
}

function Show-AttachmentPathCheck {
    $paths = Get-AttachmentPathList
    if ((Get-CollectionCount $paths) -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("附件/照片路徑目前空白。", "附件檢查", "OK", "Information") | Out-Null
        return
    }
    $missing = Get-MissingAttachmentPaths
    if ((Get-CollectionCount $missing) -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(("附件/照片路徑檢查通過，共 " + (Get-CollectionCount $paths).ToString() + " 筆。"), "附件檢查", "OK", "Information") | Out-Null
        Write-Audit "CHECK_ATTACHMENTS" "attachment" "attachment_paths" "" "PASS" ((Get-CollectionCount $paths).ToString() + " paths")
    } else {
        [System.Windows.Forms.MessageBox]::Show(("以下附件/照片路徑不存在：" + "`r`n- " + ($missing -join "`r`n- ")), "附件檢查", "OK", "Warning") | Out-Null
        Write-Audit "CHECK_ATTACHMENTS" "attachment" "attachment_paths" "" "MISSING" ($missing -join "; ")
    }
}

function Add-ChecklistWarning {
    param($Warnings, [string]$Text)
    if ($Text.Length -gt 0) { [void]$Warnings.Add($Text) }
}

function Get-PrelockChecklistWarnings {
    param([string]$LockKind)
    $warnings = New-Object System.Collections.ArrayList
    $attachmentPaths = Get-AttachmentPathList
    if ((Get-CollectionCount $attachmentPaths) -eq 0) {
        Add-ChecklistWarning $warnings "附件/照片路徑空白；若本次有照片或附件，請先補上路徑。"
    } else {
        $missing = Get-MissingAttachmentPaths
        for ($i = 0; $i -lt (Get-CollectionCount $missing); $i++) {
            Add-ChecklistWarning $warnings ("附件/照片路徑不存在：" + (Get-CollectionItem $missing $i))
        }
    }

    if ((Get-ControlTrim $script:txtCbr4Bottle1).Length -eq 0) { Add-ChecklistWarning $warnings "CBr4 第一瓶重量空白。" }
    if ((Get-ControlTrim $script:txtCbr4Bottle2).Length -eq 0) { Add-ChecklistWarning $warnings "CBr4 第二瓶重量空白。" }
    if ($script:chkSusceptor.Checked -and (Get-ControlTrim $script:txtSusceptorBakeDate).Length -eq 0) { Add-ChecklistWarning $warnings "Susceptor 已勾選，Susceptor 烘烤日期空白。" }
    if ($script:chkDisk.Checked -and (Get-ControlTrim $script:txtDiskBakeDate).Length -eq 0) { Add-ChecklistWarning $warnings "Disk/SN 已勾選，Disk 烘烤日期空白。" }
    if ($script:chkCeiling.Checked -and (Get-ControlTrim $script:txtCeilingBakeDate).Length -eq 0) { Add-ChecklistWarning $warnings "Ceiling 已勾選，Ceiling(G4) 烘烤日期空白。" }

    if ($LockKind -ne "維修項目") {
        $logRows = Normalize-TableRows (Get-GridRows $script:gridMaintenanceLog) 3 0
        $moRows = Normalize-TableRows (Get-GridRows $script:gridMoSource) (Get-MoSourceColumnCount) 0
        if ((Get-CollectionCount $logRows) -eq 0) { Add-ChecklistWarning $warnings "維修紀錄表格目前沒有資料列。" }
        if ((Get-CollectionCount $moRows) -eq 0) { Add-ChecklistWarning $warnings "MO Source 表格目前沒有資料列；若本次沒有更換，請確認可接受空白。" }
    }
    return $warnings
}

function Confirm-PrelockChecklist {
    param([string]$LockKind)
    $warnings = Get-PrelockChecklistWarnings $LockKind
    $msg = "完成前檢查清單：" + "`r`n`r`n" +
        "鎖定項目：" + $LockKind + "`r`n" +
        "必填檢查：通過" + "`r`n" +
        "附件筆數：" + (Get-CollectionCount (Get-AttachmentPathList)).ToString() + "`r`n"
    if ((Get-CollectionCount $warnings) -eq 0) {
        $msg = $msg + "`r`n未發現提醒事項。"
    } else {
        $msg = $msg + "`r`n提醒事項：" + "`r`n- " + ($warnings -join "`r`n- ")
    }
    $msg = $msg + "`r`n`r`n是否繼續下一步鎖定確認？"
    $answer = [System.Windows.Forms.MessageBox]::Show($msg, "完成前檢查清單", "OKCancel", "Warning")
    Write-Audit "PRELOCK_CHECKLIST" "record" $LockKind "" ((Get-CollectionCount $warnings).ToString() + " warnings") ($warnings -join "; ")
    return ($answer -eq [System.Windows.Forms.DialogResult]::OK)
}

function Get-CheckedItemSummary {
    $items = New-Object System.Collections.ArrayList
    foreach ($pair in @(
        @($script:chkDisk, "Disk/SN"), @($script:chkCover, "Cover"), @($script:chkStar, "Star"), @($script:chkCeiling, "Ceiling"),
        @($script:chkPullDown, "Pull down"), @($script:chkSusceptor, "Susceptor"), @($script:chkCollectorRing, "Collector ring"),
        @($script:chkExhaustTube, "Exhaust tube"), @($script:chkLightpipe, "Lightpipe"), @($script:chkThrottleValve, "Throttle Valve"),
        @($script:chkMaintPump, "Maint pump"), @($script:chkCv25, "25mbar C.V."), @($script:chkGloveboxPump, "Glovebox pump"),
        @($script:chkCt1000, "CT1000"), @($script:chkDorPump, "DOR pump"), @($script:chkYTube, "Y-tube"), @($script:chkPTrap, "P-trap"),
        @($script:chkReactorValve, "Reactor valve"), @($script:chkExhaustPipe, "Exhaust pipe"), @($script:chkDorORing, "DOR O-ring")
    )) {
        $chk = Get-CollectionItem $pair 0
        if ($null -ne $chk -and $chk.Checked) { [void]$items.Add((Get-SafeRowCell $pair 1)) }
    }
    if ((Get-CollectionCount $items) -eq 0) { return "未勾選" }
    return ($items -join ", ")
}

function Confirm-LockSummary {
    param([string]$LockKind)
    Refresh-PmTopic
    $diskCount = Get-CollectionCount (Get-DiskSerialRows)
    $moCount = Get-CollectionCount (Normalize-TableRows (Get-GridRows $script:gridMoSource) (Get-MoSourceColumnCount) 0)
    $logCount = Get-CollectionCount (Normalize-TableRows (Get-GridRows $script:gridMaintenanceLog) 3 0)
    $msg = "請確認即將鎖定的內容：" + "`r`n`r`n" +
        "鎖定項目：" + $LockKind + "`r`n" +
        "紀錄編碼：" + (Get-RecordCode) + "`r`n" +
        "機台名稱：" + (Get-ControlTrim $script:txtToolName) + "`r`n" +
        "Batch：" + (Get-ControlTrim $script:txtBatch) + "`r`n" +
        "RunID：" + (Get-ControlTrim $script:txtRunId) + "`r`n" +
        "PM日期：" + (Normalize-DateText (Get-ControlText $script:txtPmDate)) + "`r`n" +
        "PM主題：" + (Get-PmTopic) + "`r`n" +
        "維修項目：" + (Get-CheckedItemSummary) + "`r`n" +
        "Disk 序號筆數：" + $diskCount.ToString() + "`r`n" +
        "MO Source 筆數：" + $moCount.ToString() + "`r`n" +
        "維修紀錄筆數：" + $logCount.ToString() + "`r`n`r`n" +
        "確認後將鎖定，是否繼續？"
    $answer = [System.Windows.Forms.MessageBox]::Show($msg, "鎖定前確認", "YesNo", "Warning")
    return ($answer -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Save-AutoDraft {
    if ($script:itemLocked -and $script:recordLocked) { return }
    try {
        $path = Get-AutoDraftPath
        Save-DraftFile $path $false
        if ($null -ne $script:lblDraftListStatus) {
            $script:lblDraftListStatus.Text = "自動保存：" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
    } catch {
        if ($null -ne $script:lblDraftListStatus) {
            $script:lblDraftListStatus.Text = "自動保存失敗：" + $_.Exception.Message
        }
    }
}

function Get-AuditLogPaths {
    $outDir = Resolve-FullPath $script:txtOutputDir.Text
    $auditDir = Join-Path $outDir "audit"
    $paths = New-Object System.Collections.ArrayList
    if (Test-Path $auditDir) {
        $files = Get-ChildItem -Path $auditDir -Filter "maintenance_audit_*.csv" | Where-Object { -not $_.PSIsContainer }
        foreach ($f in $files) { [void]$paths.Add($f.FullName) }
    }
    return $paths
}

function Get-MaintenanceItemRows {
    $rows = New-Object System.Collections.ArrayList
    [void]$rows.Add(@("Disk/SN", (Get-CheckboxValue $script:chkDisk), (Get-DiskSerialSummary)))
    $diskRows = Get-DiskSerialRows
    for ($i = 0; $i -lt (Get-CollectionCount $diskRows); $i++) {
        $diskRow = Get-CollectionItem $diskRows $i
        [void]$rows.Add(@("Disk " + (Get-SafeRowCell $diskRow 0), "", (Get-SafeRowCell $diskRow 1)))
    }
    [void]$rows.Add(@("Cover", (Get-CheckboxValue $script:chkCover), $script:txtCoverSerial.Text.Trim()))
    [void]$rows.Add(@("Star", (Get-CheckboxValue $script:chkStar), $script:txtStarSerial.Text.Trim()))
    [void]$rows.Add(@("Ceiling", (Get-CheckboxValue $script:chkCeiling), $script:txtCeiling.Text.Trim()))
    [void]$rows.Add(@("Pull down", (Get-CheckboxValue $script:chkPullDown), $script:txtPullDown.Text.Trim()))
    [void]$rows.Add(@("Susceptor", (Get-CheckboxValue $script:chkSusceptor), $script:txtSusceptor.Text.Trim()))
    [void]$rows.Add(@("Collector ring", (Get-CheckboxValue $script:chkCollectorRing), $script:txtCollectorRing.Text.Trim()))
    [void]$rows.Add(@("Exhaust tube", (Get-CheckboxValue $script:chkExhaustTube), $script:txtExhaustTube.Text.Trim()))
    [void]$rows.Add(@("Lightpipe", (Get-CheckboxValue $script:chkLightpipe), $script:txtLightpipe.Text.Trim()))
    [void]$rows.Add(@("Throttle Valve", (Get-CheckboxValue $script:chkThrottleValve), $script:txtThrottleValve.Text.Trim()))
    [void]$rows.Add(@("Maint pump", (Get-CheckboxValue $script:chkMaintPump), $script:txtMaintPump.Text.Trim()))
    [void]$rows.Add(@("25mbar C.V.", (Get-CheckboxValue $script:chkCv25), $script:txtCv25.Text.Trim()))
    [void]$rows.Add(@("Glovebox pump", (Get-CheckboxValue $script:chkGloveboxPump), $script:txtGloveboxPump.Text.Trim()))
    [void]$rows.Add(@("CT1000", (Get-CheckboxValue $script:chkCt1000), $script:txtCt1000.Text.Trim()))
    [void]$rows.Add(@("DOR pump", (Get-CheckboxValue $script:chkDorPump), $script:txtDorPump.Text.Trim()))
    [void]$rows.Add(@("Y-tube", (Get-CheckboxValue $script:chkYTube), $script:txtYTube.Text.Trim()))
    [void]$rows.Add(@("P-trap", (Get-CheckboxValue $script:chkPTrap), $script:txtPTrap.Text.Trim()))
    [void]$rows.Add(@("Reactor valve", (Get-CheckboxValue $script:chkReactorValve), $script:txtReactorValve.Text.Trim()))
    [void]$rows.Add(@("Exhaust pipe", (Get-CheckboxValue $script:chkExhaustPipe), $script:txtExhaustPipe.Text.Trim()))
    [void]$rows.Add(@("DOR O-ring", (Get-CheckboxValue $script:chkDorORing), $script:txtDorORing.Text.Trim()))
    [void]$rows.Add(@("特氣是否切換", [string]$script:cmbSpecialGas.SelectedItem, ""))
    [void]$rows.Add(@("Susceptor烘烤日期", (Normalize-DateText $script:txtSusceptorBakeDate.Text), ""))
    [void]$rows.Add(@("Disk烘烤日期", (Normalize-DateText $script:txtDiskBakeDate.Text), ""))
    [void]$rows.Add(@("Ceiling(G4)烘烤日期", (Normalize-DateText $script:txtCeilingBakeDate.Text), ""))
    [void]$rows.Add(@("CBr4第一瓶重量(g)", $script:txtCbr4Bottle1.Text.Trim(), ""))
    [void]$rows.Add(@("CBr4第二瓶重量(g)", $script:txtCbr4Bottle2.Text.Trim(), ""))
    [void]$rows.Add(@("注意事項", "CBr4", "確認下一 Batch 如 CBr4 第一瓶重量會 < 滿瓶的 30%，需下機執行敲打作業"))
    [void]$rows.Add(@("注意事項", "MFC", "MFC default 設定值若小於 10 sccm，請一律設定為 10 sccm"))
    [void]$rows.Add(@("其他維修項目", $script:txtOtherItems.Text.Trim(), ""))
    return $rows
}

function Set-TextControlLock {
    param($Control, [bool]$Locked)
    if ($Control -is [System.Windows.Forms.TextBox]) {
        $Control.ReadOnly = $Locked
    } elseif ($Control -is [System.Windows.Forms.DataGridView]) {
        $Control.ReadOnly = $Locked
        $Control.AllowUserToAddRows = (-not $Locked)
        $Control.AllowUserToDeleteRows = (-not $Locked)
    } else {
        $Control.Enabled = (-not $Locked)
    }
}

function Set-ItemsLocked {
    param([bool]$Locked)
    foreach ($ctrl in $script:itemEditControls) {
        Set-TextControlLock $ctrl $Locked
    }
    $script:btnLockItems.Enabled = (-not $Locked)
    $script:btnLockItems.Text = "維修項目完成鎖定"
    if ($Locked) { $script:btnLockItems.Text = "維修項目已鎖定" }
    $script:btnUnlockItems.Enabled = $Locked
    $script:lblItemLockStatus.Text = "維修項目狀態：編輯中"
    if ($Locked) { $script:lblItemLockStatus.Text = "維修項目狀態：已鎖定" }
    if ($null -ne $script:btnLockItemsRecords) {
        $script:btnLockItemsRecords.Enabled = (-not $Locked)
        $script:btnLockItemsRecords.Text = "維修項目完成鎖定"
        if ($Locked) { $script:btnLockItemsRecords.Text = "維修項目已鎖定" }
    }
    if ($null -ne $script:btnUnlockItemsRecords) {
        $script:btnUnlockItemsRecords.Enabled = $Locked
    }
    if ($null -ne $script:lblItemLockStatusRecords) {
        $script:lblItemLockStatusRecords.Text = "維修項目狀態：編輯中"
        if ($Locked) { $script:lblItemLockStatusRecords.Text = "維修項目狀態：已鎖定" }
    }
}

function Set-RecordsLocked {
    param([bool]$Locked)
    $script:gridMaintenanceLog.ReadOnly = $Locked
    $script:gridMaintenanceLog.AllowUserToAddRows = (-not $Locked)
    $script:gridMaintenanceLog.AllowUserToDeleteRows = (-not $Locked)
    $script:gridMoSource.ReadOnly = $Locked
    $script:gridMoSource.AllowUserToAddRows = (-not $Locked)
    $script:gridMoSource.AllowUserToDeleteRows = (-not $Locked)
    $script:btnAddLog.Enabled = (-not $Locked)
    $script:btnDelLog.Enabled = (-not $Locked)
    $script:btnAddMo.Enabled = (-not $Locked)
    $script:btnDelMo.Enabled = (-not $Locked)
    $script:btnLockRecords.Enabled = (-not $Locked)
    $script:btnLockRecords.Text = "維修紀錄完成鎖定"
    if ($Locked) { $script:btnLockRecords.Text = "維修紀錄已鎖定" }
    $script:btnUnlockRecords.Enabled = $Locked
    $script:lblRecordLockStatus.Text = "維修紀錄狀態：編輯中"
    if ($Locked) { $script:lblRecordLockStatus.Text = "維修紀錄狀態：已鎖定" }
}

function Ensure-NoOpenControlledEdit {
    if ($script:itemEditReason.Length -gt 0) {
        throw "維修項目正在管制修改中，請先重新鎖定後再輸出。"
    }
    if ($script:recordEditReason.Length -gt 0) {
        throw "維修紀錄正在管制修改中，請先重新鎖定後再輸出。"
    }
}

function Ensure-FinalLocked {
    if (-not $script:itemLocked) {
        throw "維修項目尚未完成鎖定，不能輸出正式維修紀錄。"
    }
    if (-not $script:recordLocked) {
        throw "維修紀錄尚未完成鎖定，不能輸出正式維修紀錄。"
    }
}

function Draw-PdfCell {
    param($G, $Pen, $Brush, $Font, [float]$X, [float]$Y, [float]$W, [float]$H, [string]$Text, [bool]$FillHeader, [bool]$AllowWrap)
    if ($null -eq $G) { throw "PDF繪圖物件為空，無法輸出。" }
    if ($null -eq $Pen) { throw "PDF線條物件為空，無法輸出。" }
    if ($null -eq $Brush) { throw "PDF文字筆刷為空，無法輸出。" }
    if ($null -eq $Font) { throw "PDF字型物件為空，無法輸出。" }
    if ($null -eq $Text) { $Text = "" }
    $rect = New-Object System.Drawing.RectangleF($X, $Y, $W, $H)
    if ($FillHeader) {
        $fill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(230, 238, 247))
        try {
            $G.FillRectangle($fill, $rect)
        } finally {
            if ($null -ne $fill) { $fill.Dispose() }
        }
    }
    $G.DrawRectangle($Pen, $X, $Y, $W, $H)
    $sf = New-Object System.Drawing.StringFormat
    try {
        $sf.Alignment = [System.Drawing.StringAlignment]::Near
        if ($AllowWrap) {
            # PATCH v6-pdf-diskwrap-01: 允許 Disk/SN 序號/內容欄在儲存格內自動換行。
            $sf.Trimming = [System.Drawing.StringTrimming]::Word
            $sf.FormatFlags = 0
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Near
        } else {
            $sf.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
            $sf.FormatFlags = [System.Drawing.StringFormatFlags]::LineLimit
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        }
        # PATCH v7-pdf-layout-02: 非換行欄位垂直置中，避免直式 PDF 小列高文字貼線或被裁切。
        $textRect = New-Object System.Drawing.RectangleF(($X + 3), $Y, ($W - 6), $H)
        $G.DrawString($Text, $Font, $Brush, $textRect, $sf)
    } finally {
        if ($null -ne $sf) { $sf.Dispose() }
    }
}

function Draw-PdfRow {
    # PATCH 2026-06-08: Widths / Texts 不宣告為 [object[]]，避免 Windows PowerShell
    # 在部分呼叫情境把陣列拆解或只取第一個元素，導致 MO Source 表頭只剩 Source。
    param($G, $Pen, $Brush, $Font, [float]$X, [float]$Y, $Widths, [float]$H, $Texts, [bool]$FillHeader)
    $xx = $X
    $widthCount = Get-CollectionCount $Widths
    $textCount = Get-CollectionCount $Texts
    for ($i = 0; $i -lt $widthCount; $i++) {
        $txt = ""
        $textValue = Get-CollectionItem $Texts $i
        if ($i -lt $textCount -and $null -ne $textValue) { $txt = [string]$textValue }
        $widthValue = Get-CollectionItem $Widths $i
        Draw-PdfCell $G $Pen $Brush $Font $xx $Y ([float]$widthValue) $H $txt $FillHeader $false
        $xx += [float]$widthValue
    }
}

function Get-PdfItemCell {
    param($ItemRows, [string]$Name, [int]$Index)
    $rowCount = Get-CollectionCount $ItemRows
    for ($i = 0; $i -lt $rowCount; $i++) {
        $row = Get-CollectionItem $ItemRows $i
        if ((Get-SafeRowCell $row 0) -eq $Name) {
            return (Get-SafeRowCell $row $Index)
        }
    }
    return ""
}


function Get-DiskSerialPdfSerial {
    # PATCH 2026-06-08 Disk/SN PDF serial fix
    # 直接從 GUI 的 Disk 位置/序號表格讀值，避免先轉成 ItemRows 後因 PowerShell 陣列繫結造成序號欄遺失。
    param([string]$Position)
    if ($null -eq $Position) { return "" }
    $target = $Position.Trim().ToUpperInvariant()
    if ($target.Length -eq 0) { return "" }

    $rows = Get-DiskSerialRows
    $rowCount = Get-CollectionCount $rows
    for ($i = 0; $i -lt $rowCount; $i++) {
        $row = Get-CollectionItem -Value $rows -Index $i
        $pos = (Get-SafeRowCell -Row $row -Index 0).Trim().ToUpperInvariant()
        if ($pos -eq $target) {
            return (Get-SafeRowCell -Row $row -Index 1)
        }
    }
    return ""
}


function Draw-MoSourcePdfSection {
    # PATCH 2026-06-09 MO Source PDF Horizontal Row Fix
    # - PDF must match the GUI data grain: one MO Source change record per row.
    # - Keep all 12 fields on the same horizontal row instead of splitting values into a vertical block.
    param($G, $Pen, $Brush, $Font, $HeadFont, [float]$X, [float]$Y, [float]$W, $MoRows, [string]$Mode, [int]$PageIndex)

    $moFont = New-Object System.Drawing.Font("Microsoft JhengHei", 5)
    $moHeadFont = New-Object System.Drawing.Font("Microsoft JhengHei", 5.5, [System.Drawing.FontStyle]::Bold)
    # Get-MoSourceRowsForExport 透過 ",$rows" 已回傳完整列集合（每列 12 欄）。
    # 不可再 Get-CollectionItem 0，否則只有一筆 source 時會被誤取成第一列，
    # 導致 12 個欄位被當成 12 列直向輸出。
    $safeMoRows = Get-MoSourceRowsForExport 0
    $drawRows = New-Object System.Collections.ArrayList
    $startMo = 0
    $maxMo = Get-CollectionCount $safeMoRows
    if ($Mode -eq "Full" -and $PageIndex -eq 0 -and $maxMo -gt 8) { $maxMo = 8 }
    if ($Mode -eq "Full" -and $PageIndex -eq 1) { $startMo = 8; $maxMo = Get-CollectionCount $safeMoRows }
    for ($i = $startMo; $i -lt $maxMo; $i++) {
        $row = Get-CollectionItem $safeMoRows $i
        [void]$drawRows.Add($row)
    }

    if ((Get-CollectionCount $drawRows) -eq 0) {
        $drawRows = Normalize-TableRows $null (Get-MoSourceColumnCount) 3
    }

    try {
        Draw-PdfRow $G $Pen $Brush $moHeadFont $X $Y @($W) 16 @("MO Source Change Record") $true
        $Y += 16

        $widths = @(60, 65, 55, 65, 65, 65, 65, 65, 65, 65, 80, 110)
        $headers = @("Source", "設定瓶壓", "位置", "重量Off", "重量On", "重量T/W", "來料瓶壓", "實際瓶壓", "效率Off", "效率On", "退庫數量", "備註")
        $xx = $X
        for ($c = 0; $c -lt (Get-MoSourceColumnCount); $c++) {
            Draw-PdfCell $G $Pen $Brush $moHeadFont $xx $Y ([float](Get-CollectionItem $widths $c)) 18 ([string](Get-CollectionItem $headers $c)) $true $false
            $xx += [float](Get-CollectionItem $widths $c)
        }
        $Y += 18
        for ($r = 0; $r -lt (Get-CollectionCount $drawRows); $r++) {
            $row = Get-CollectionItem $drawRows $r
            $xx = $X
            for ($c = 0; $c -lt (Get-MoSourceColumnCount); $c++) {
                Draw-PdfCell $G $Pen $Brush $moFont $xx $Y ([float](Get-CollectionItem $widths $c)) 18 (Get-SafeRowCell $row $c) $false $false
                $xx += [float](Get-CollectionItem $widths $c)
            }
            $Y += 18
        }
    } finally {
        if ($null -ne $moFont) { $moFont.Dispose() }
        if ($null -ne $moHeadFont) { $moHeadFont.Dispose() }
    }

    return $Y
}

function Draw-MaintenanceItemsPdfSection {
    param($G, $Pen, $Brush, $Font, $HeadFont, [float]$X, [float]$Y, [float]$W, $ItemRows)
    $rh = 17
    Draw-PdfRow $G $Pen $Brush $HeadFont $X $Y @($W) $rh @("維修項目") $true
    $Y += $rh

    $leftW = 540
    $diskW = 260
    $gap = 10
    $diskX = $X + $leftW + $gap
    Draw-PdfRow $G $Pen $Brush $HeadFont $X $Y @(75, 45, 145, 75, 45, 155) $rh @("項目", "值", "序號/內容", "項目", "值", "序號/內容") $true
    Draw-PdfRow $G $Pen $Brush $HeadFont $diskX $Y @(70, 190) $rh @("Disk/SN", "序號/內容") $true
    $Y += $rh

    $pairs = @(
        @("Star", "Cover"),
        @("Ceiling", "Pull down"),
        @("Susceptor", "Collector ring"),
        @("Lightpipe", "Exhaust tube"),
        @("Throttle Valve", "25mbar C.V."),
        @("Maint pump", "CT1000"),
        @("Glovebox pump", "Y-tube"),
        @("DOR pump", "Reactor valve"),
        @("P-trap", "DOR O-ring"),
        @("Exhaust pipe", "")
    )
    $diskPositions = @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")
    for ($i = 0; $i -lt 12; $i++) {
        if ($i -lt (Get-CollectionCount $pairs)) {
            $pair = Get-CollectionItem $pairs $i
            $leftName = Get-SafeRowCell $pair 0
            $rightName = Get-SafeRowCell $pair 1
            Draw-PdfRow $G $Pen $Brush $Font $X $Y @(75, 45, 145, 75, 45, 155) $rh @(
                $leftName,
                (Get-PdfItemCell $ItemRows $leftName 1),
                (Get-PdfItemCell $ItemRows $leftName 2),
                $rightName,
                (Get-PdfItemCell $ItemRows $rightName 1),
                (Get-PdfItemCell $ItemRows $rightName 2)
            ) $false
        } else {
            Draw-PdfRow $G $Pen $Brush $Font $X $Y @(75, 45, 145, 75, 45, 155) $rh @("", "", "", "", "", "") $false
        }
        $diskPos = [string](Get-CollectionItem $diskPositions $i)
        $diskSerial = Get-DiskSerialPdfSerial -Position $diskPos
        Draw-PdfRow $G $Pen $Brush $Font $diskX $Y @(70, 190) $rh @($diskPos, $diskSerial) $false
        $Y += $rh
    }

    $Y += 4
    Draw-PdfRow $G $Pen $Brush $Font $X $Y @(90, 130, 130, 130, 130, 220) $rh @("特氣切換", (Get-PdfItemCell $ItemRows "特氣是否切換" 1), "", "", "", "") $false
    $Y += $rh
    Draw-PdfRow $G $Pen $Brush $HeadFont $X $Y @(90, 130, 130, 130, 130, 220) $rh @("烘烤部件", "Susceptor", "Disk", "Ceiling", "", "") $false
    $Y += $rh
    Draw-PdfRow $G $Pen $Brush $Font $X $Y @(90, 130, 130, 130, 130, 220) $rh @("烘烤日期", (Get-PdfItemCell $ItemRows "Susceptor烘烤日期" 1), (Get-PdfItemCell $ItemRows "Disk烘烤日期" 1), (Get-PdfItemCell $ItemRows "Ceiling(G4)烘烤日期" 1), "", "") $false
    $Y += $rh
    Draw-PdfRow $G $Pen $Brush $Font $X $Y @(90, 130, 130, 130, 130, 220) $rh @("CBr4", "第一瓶重量", (Get-PdfItemCell $ItemRows "CBr4第一瓶重量(g)" 1), "第二瓶重量", (Get-PdfItemCell $ItemRows "CBr4第二瓶重量(g)" 1), "") $false
    $Y += $rh
    Draw-PdfRow $G $Pen $Brush $Font $X $Y @(90, 740) $rh @("", "注意事項：確認下一 Batch 如 CBr4 第一瓶重量會 < 滿瓶的 30%，需下機執行敲打作業") $false
    $Y += $rh
    Draw-PdfRow $G $Pen $Brush $Font $X $Y @(90, 740) $rh @("", "MFC default 設定值若小於 10 sccm，請一律設定為 10 sccm") $false
    $Y += $rh
    Draw-PdfRow $G $Pen $Brush $Font $X $Y @(90, 740) 20 @("其他維修項目", (Get-PdfItemCell $ItemRows "其他維修項目" 1)) $false
    $Y += 24
    return $Y
}

function Draw-MaintenancePdfPage {
    param($G, $Bounds, [string]$Mode, [int]$PageIndex, $LogRows, $MoRows, $ItemRows)
    $x = [float]$Bounds.Left
    $y = [float]$Bounds.Top
    $w = [float]$Bounds.Width
    $pen = $null
    $brush = $null
    $titleFont = $null
    $headFont = $null
    $font = $null
    $logFont = $null
    $logHeadFont = $null
    try {
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(128, 128, 128), 1)
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
        # PATCH v7-pdf-layout-02: 直式 PDF 縮小字級並增加列高，避免標題與表格內容互相壓到。
        $titleFont = New-Object System.Drawing.Font("Microsoft JhengHei", 13, [System.Drawing.FontStyle]::Bold)
        $headFont = New-Object System.Drawing.Font("Microsoft JhengHei", 7, [System.Drawing.FontStyle]::Bold)
        $font = New-Object System.Drawing.Font("Microsoft JhengHei", 6)
        $logHeadFont = New-Object System.Drawing.Font("Microsoft JhengHei UI", 7.5, [System.Drawing.FontStyle]::Bold)
        $logFont = New-Object System.Drawing.Font("Microsoft JhengHei UI", 7)

        $title = "MOCVD 維修前檢查表"
        if ($Mode -eq "Full") { $title = "MOCVD 完整維修紀錄" }
        if ($Mode -eq "Full" -and $PageIndex -gt 0) { $title = $title + " - 第 " + ($PageIndex + 1).ToString() + " 頁" }
        $G.DrawString($title, $titleFont, $brush, $x, $y)
        $y += 36

        Draw-PdfRow $G $pen $brush $font $x $y @(85, 745) 20 @("PM主題", (Get-PmTopic)) $false
        $y += 20
        Draw-PdfRow $G $pen $brush $font $x $y @(85, 300, 100, 345) 20 @("紀錄編碼", (Get-RecordCode), "機台名稱", (Get-ControlTrim $script:txtToolName)) $false
        $y += 24

        Draw-PdfRow $G $pen $brush $headFont $x $y @($w) 18 @("基本資料") $true
        $y += 18
        Draw-PdfRow $G $pen $brush $font $x $y @(70, 190, 70, 160, 70, 160) 20 @("Batch", (Get-ControlTrim $script:txtBatch), "產品", (Get-ControlTrim $script:txtProduct), "PM類別", (Get-SelectedText $script:cmbPmCategory)) $false
        $y += 20
        Draw-PdfRow $G $pen $brush $font $x $y @(70, 190, 70, 160, 70, 160) 20 @("PM日期", (Normalize-DateText (Get-ControlText $script:txtPmDate)), "班別", (Get-SelectedText $script:cmbShift), "RunID", (Get-ControlTrim $script:txtRunId)) $false
        $y += 20
        Draw-PdfRow $G $pen $brush $font $x $y @(70, 190, 70, 160, 70, 160) 20 @("會議人員", (Get-ControlTrim $script:txtMeetingPeople), "維修人員", (Get-ControlTrim $script:txtMaintPeople), "維修日期", (Normalize-DateText (Get-ControlText $script:txtMaintDate))) $false
        $y += 24

        $y = Draw-MaintenanceItemsPdfSection $G $pen $brush $font $headFont $x $y $w $ItemRows

        $logRowsForPage = Normalize-TableRows $LogRows 3 3
        $moRowsForPage = Normalize-TableRows $MoRows (Get-MoSourceColumnCount) 3
        if ($Mode -eq "Pre") {
            $logRowsForPage = Normalize-TableRows $null 3 5
            $moRowsForPage = Normalize-TableRows $MoRows (Get-MoSourceColumnCount) 5
        }

        Draw-PdfRow $G $pen $brush $logHeadFont $x $y @($w) 18 @("維修紀錄") $true
        $y += 18
        Draw-PdfRow $G $pen $brush $logHeadFont $x $y @(90, 360, 380) 18 @("日期", "內容", "備註") $true
        $y += 18
        $startLog = 0
        $maxLog = Get-CollectionCount $logRowsForPage
        if ($Mode -eq "Full" -and $PageIndex -eq 0 -and $maxLog -gt 8) { $maxLog = 8 }
        if ($Mode -eq "Full" -and $PageIndex -eq 1) { $startLog = 8; $maxLog = Get-CollectionCount $logRowsForPage }
        for ($i = $startLog; $i -lt $maxLog; $i++) {
            $logRow = Get-CollectionItem $logRowsForPage $i
            Draw-PdfRow $G $pen $brush $logFont $x $y @(90, 360, 380) 18 @((Get-SafeRowCell $logRow 0), (Get-SafeRowCell $logRow 1), (Get-SafeRowCell $logRow 2)) $false
            $y += 18
        }
        $y += 6

        $y = Draw-MoSourcePdfSection $G $pen $brush $font $headFont $x $y $w $moRowsForPage $Mode $PageIndex
    } finally {
        if ($null -ne $font) { $font.Dispose() }
        if ($null -ne $headFont) { $headFont.Dispose() }
        if ($null -ne $titleFont) { $titleFont.Dispose() }
        if ($null -ne $logFont) { $logFont.Dispose() }
        if ($null -ne $logHeadFont) { $logHeadFont.Dispose() }
        if ($null -ne $brush) { $brush.Dispose() }
        if ($null -ne $pen) { $pen.Dispose() }
    }
}

function New-MaintenancePdfPageImage {
    param([string]$Mode, [int]$PageIndex, $LogRows, $MoRows, $ItemRows)
    $bmp = $null
    $g = $null
    $ms = $null
    try {
        $bmp = New-Object System.Drawing.Bitmap(900, 1400)
        $bmp.SetResolution(150, 150)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
        $bounds = New-Object System.Drawing.Rectangle(35, 35, 830, 1290)
        Draw-MaintenancePdfPage $g $bounds $Mode $PageIndex $LogRows $MoRows $ItemRows
        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        return ,$ms.ToArray()
    } finally {
        if ($null -ne $ms) { $ms.Dispose() }
        if ($null -ne $g) { $g.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
    }
}

function Write-AsciiBytes {
    param($Stream, [string]$Text)
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($Text)
    $Stream.Write($bytes, 0, $bytes.Length)
}

function Convert-ToPdfImageBytes {
    param($Value)
    if ($null -eq $Value) {
        throw "PDF 頁面影像為空，無法輸出。"
    }

    $current = $Value

    while (($current -is [System.Array]) -and -not ($current -is [byte[]])) {
        if ($current.Length -eq 0) {
            throw "PDF 頁面影像陣列為空，無法輸出。"
        }
        if ($current.Length -eq 1) {
            $current = $current[0]
        } else {
            break
        }
    }

    if ($current -is [byte[]]) {
        if ($current.Length -lt 4) {
            throw "PDF 頁面影像資料太短，可能產生空白 PDF。"
        }
        return ,$current
    }

    if ($current -is [System.Collections.ICollection]) {
        $count = $current.Count
        if ($count -lt 4) {
            throw "PDF 頁面影像資料太短，可能產生空白 PDF。"
        }
        $bytes = New-Object byte[] $count
        $i = 0
        foreach ($item in $current) {
            $bytes[$i] = [System.Convert]::ToByte($item)
            $i++
        }
        return ,$bytes
    }

    throw ("PDF 頁面影像格式錯誤：" + $current.GetType().FullName)
}

function Write-ImagePdf {
    param([string]$Path, [object[]]$PageImages, [int]$ImageWidth, [int]$ImageHeight)
    $pageCount = Get-CollectionCount $PageImages
    if ($pageCount -eq 0) { throw "沒有可輸出的 PDF 頁面。" }
    # 影像式 PDF 退回路徑同樣固定為橫式 (landscape) Letter 版面，與 Excel 匯出一致。
    $pageWidth = 792
    $pageHeight = 612

    # 頁面內容影像維持原始比例置中，避免拉伸變形或被裁切。
    $scaleW = $pageWidth / [double]$ImageWidth
    $scaleH = $pageHeight / [double]$ImageHeight
    $scale = $scaleW
    if ($scaleH -lt $scale) { $scale = $scaleH }
    $drawWidth = [double]$ImageWidth * $scale
    $drawHeight = [double]$ImageHeight * $scale
    $drawOffsetX = ($pageWidth - $drawWidth) / 2.0
    $drawOffsetY = ($pageHeight - $drawHeight) / 2.0

    $fs = $null
    try {
        $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
        $offsets = New-Object System.Collections.ArrayList

        Write-AsciiBytes $fs "%PDF-1.4`n"

        [void]$offsets.Add($fs.Position)
        Write-AsciiBytes $fs "1 0 obj`n<< /Type /Catalog /Pages 2 0 R >>`nendobj`n"

        $kids = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt $pageCount; $i++) {
            $pageObj = 3 + ($i * 3)
            [void]$kids.Append($pageObj.ToString() + " 0 R ")
        }

        [void]$offsets.Add($fs.Position)
        Write-AsciiBytes $fs ("2 0 obj`n<< /Type /Pages /Kids [ " + $kids.ToString() + "] /Count " + $pageCount.ToString() + " >>`nendobj`n")

        for ($i = 0; $i -lt $pageCount; $i++) {
            $pageObj = 3 + ($i * 3)
            $imageObj = $pageObj + 1
            $contentObj = $pageObj + 2
            $imageName = "Im" + ($i + 1).ToString()

            $rawImage = Get-CollectionItem $PageImages $i
            $img = Convert-ToPdfImageBytes $rawImage
            if ($null -eq $img) { throw "PDF 頁面影像為空。" }
            if ($img.Length -lt 4) { throw "PDF 頁面影像資料太短，無法輸出。" }

            if (-not ($img[0] -eq 255 -and $img[1] -eq 216)) {
                throw "PDF 頁面影像不是 JPEG 格式，可能造成 PDF 空白。"
            }

            [void]$offsets.Add($fs.Position)
            Write-AsciiBytes $fs ($pageObj.ToString() + " 0 obj`n")
            Write-AsciiBytes $fs ("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + $pageWidth.ToString() + " " + $pageHeight.ToString() + "] ")
            Write-AsciiBytes $fs ("/Resources << /ProcSet [/PDF /ImageC] /XObject << /" + $imageName + " " + $imageObj.ToString() + " 0 R >> >> ")
            Write-AsciiBytes $fs ("/Contents " + $contentObj.ToString() + " 0 R >>`nendobj`n")

            [void]$offsets.Add($fs.Position)
            Write-AsciiBytes $fs ($imageObj.ToString() + " 0 obj`n")
            Write-AsciiBytes $fs ("<< /Type /XObject /Subtype /Image /Width " + $ImageWidth.ToString() + " /Height " + $ImageHeight.ToString())
            Write-AsciiBytes $fs (" /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + $img.Length.ToString() + " >>`nstream`n")
            $fs.Write($img, 0, $img.Length)
            Write-AsciiBytes $fs "`nendstream`nendobj`n"

            $inv = [System.Globalization.CultureInfo]::InvariantCulture
            $content = "q`n" + $drawWidth.ToString("0.##", $inv) + " 0 0 " + $drawHeight.ToString("0.##", $inv) + " " + $drawOffsetX.ToString("0.##", $inv) + " " + $drawOffsetY.ToString("0.##", $inv) + " cm`n/" + $imageName + " Do`nQ"
            $contentBytes = [System.Text.Encoding]::ASCII.GetBytes($content)
            [void]$offsets.Add($fs.Position)
            Write-AsciiBytes $fs ($contentObj.ToString() + " 0 obj`n<< /Length " + $contentBytes.Length.ToString() + " >>`nstream`n")
            $fs.Write($contentBytes, 0, $contentBytes.Length)
            Write-AsciiBytes $fs "`nendstream`nendobj`n"
        }

        $xref = $fs.Position
        $objectCount = 2 + ($pageCount * 3)
        Write-AsciiBytes $fs ("xref`n0 " + ($objectCount + 1).ToString() + "`n")
        Write-AsciiBytes $fs "0000000000 65535 f `n"
        for ($i = 0; $i -lt (Get-CollectionCount $offsets); $i++) {
            $offsetValue = [int64](Get-CollectionItem $offsets $i)
            Write-AsciiBytes $fs ($offsetValue.ToString("0000000000") + " 00000 n `n")
        }
        Write-AsciiBytes $fs ("trailer`n<< /Size " + ($objectCount + 1).ToString() + " /Root 1 0 R >>`nstartxref`n" + $xref.ToString() + "`n%%EOF`n")
    } finally {
        if ($null -ne $fs) { $fs.Dispose() }
    }
}

function Set-ExcelCell {
    param($Sheet, [int]$Row, [int]$Col, [string]$Text, [bool]$Bold, [double]$Size)
    $cell = $Sheet.Cells.Item($Row, $Col)
    $cell.Value2 = $Text
    $cell.Font.Name = "Microsoft JhengHei"
    $cell.Font.Size = $Size
    $cell.Font.Bold = $Bold
    $cell.WrapText = $true
    $cell.VerticalAlignment = -4108
}

function Merge-ExcelCells {
    param($Sheet, [int]$Row, [int]$Col1, [int]$Col2, [string]$Text, [bool]$Bold, [double]$Size, [bool]$FillHeader)
    $range = $Sheet.Range($Sheet.Cells.Item($Row, $Col1), $Sheet.Cells.Item($Row, $Col2))
    $range.Merge() | Out-Null
    $range.Value2 = $Text
    $range.Font.Name = "Microsoft JhengHei"
    $range.Font.Size = $Size
    $range.Font.Bold = $Bold
    $range.WrapText = $true
    $range.VerticalAlignment = -4108
    if ($FillHeader) { $range.Interior.Color = 16247773 }
}

function Set-ExcelBorder {
    param($Sheet, [int]$Row1, [int]$Col1, [int]$Row2, [int]$Col2)
    $range = $Sheet.Range($Sheet.Cells.Item($Row1, $Col1), $Sheet.Cells.Item($Row2, $Col2))
    $range.Borders.LineStyle = 1
    $range.Borders.Color = 12632256
    $range.Borders.Weight = 2
}

function Write-MaintenancePdfViaExcel {
    param([string]$Mode, [string]$PdfPath, $LogRows, $MoRows, $ItemRows)
    $excel = $null
    $wb = $null
    $ws = $null
    $xlsxPath = [System.IO.Path]::ChangeExtension($PdfPath, ".xlsx")
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = "Maintenance"

        for ($c = 1; $c -le 12; $c++) { $ws.Columns.Item($c).ColumnWidth = 12 }
        $ws.Columns.Item(1).ColumnWidth = 9
        $ws.Columns.Item(3).ColumnWidth = 16
        $ws.Columns.Item(4).ColumnWidth = 16
        $ws.Columns.Item(7).ColumnWidth = 16
        $ws.Columns.Item(8).ColumnWidth = 16
        $ws.Columns.Item(11).ColumnWidth = 14
        $ws.Columns.Item(12).ColumnWidth = 14

        $row = 1
        $title = "MOCVD 維修前檢查表"
        if ($Mode -eq "Full") { $title = "MOCVD 完整維修紀錄" }
        Merge-ExcelCells $ws $row 1 12 $title $true 16 $false
        $ws.Rows.Item($row).RowHeight = 28
        $row++

        Merge-ExcelCells $ws $row 1 12 "基本資料" $true 10 $true
        $row++
        Set-ExcelCell $ws $row 1 "PM主題" $true 9
        Merge-ExcelCells $ws $row 2 12 (Get-PmTopic) $false 9 $false
        $row++
        Set-ExcelCell $ws $row 1 "紀錄編碼" $true 9
        Merge-ExcelCells $ws $row 2 4 (Get-RecordCode) $false 9 $false
        Set-ExcelCell $ws $row 5 "機台名稱" $true 9
        Merge-ExcelCells $ws $row 6 8 (Get-ControlTrim $script:txtToolName) $false 9 $false
        Set-ExcelCell $ws $row 9 "Batch" $true 9
        Merge-ExcelCells $ws $row 10 12 (Get-ControlTrim $script:txtBatch) $false 9 $false
        $row++
        Set-ExcelCell $ws $row 1 "產品" $true 9
        Merge-ExcelCells $ws $row 2 4 (Get-ControlTrim $script:txtProduct) $false 9 $false
        Set-ExcelCell $ws $row 5 "PM類別" $true 9
        Merge-ExcelCells $ws $row 6 8 (Get-SelectedText $script:cmbPmCategory) $false 9 $false
        Set-ExcelCell $ws $row 9 "PM日期" $true 9
        Merge-ExcelCells $ws $row 10 12 (Normalize-DateText (Get-ControlText $script:txtPmDate)) $false 9 $false
        $row++
        Set-ExcelCell $ws $row 1 "RunID" $true 9
        Merge-ExcelCells $ws $row 2 4 (Get-ControlTrim $script:txtRunId) $false 9 $false
        Set-ExcelCell $ws $row 5 "維修人員" $true 9
        Merge-ExcelCells $ws $row 6 8 (Get-ControlTrim $script:txtMaintPeople) $false 9 $false
        Set-ExcelCell $ws $row 9 "維修日期" $true 9
        Merge-ExcelCells $ws $row 10 12 (Normalize-DateText (Get-ControlText $script:txtMaintDate)) $false 9 $false
        Set-ExcelBorder $ws 2 1 $row 12
        $row += 2

        Merge-ExcelCells $ws $row 1 12 "維修項目" $true 10 $true
        $itemStart = $row
        $row++
        foreach ($c in 1..12) { Set-ExcelCell $ws $row $c "" $true 8 }
        Set-ExcelCell $ws $row 1 "項目" $true 8
        Set-ExcelCell $ws $row 2 "值" $true 8
        Merge-ExcelCells $ws $row 3 4 "序號/內容" $true 8 $true
        Set-ExcelCell $ws $row 5 "項目" $true 8
        Set-ExcelCell $ws $row 6 "值" $true 8
        Merge-ExcelCells $ws $row 7 8 "序號/內容" $true 8 $true
        Set-ExcelCell $ws $row 9 "Disk/SN" $true 8
        Merge-ExcelCells $ws $row 10 12 "序號/內容" $true 8 $true
        $row++
        $pairs = @(
            @("Star", "Cover"), @("Ceiling", "Pull down"), @("Susceptor", "Collector ring"),
            @("Lightpipe", "Exhaust tube"), @("Throttle Valve", "25mbar C.V."), @("Maint pump", "CT1000"),
            @("Glovebox pump", "Y-tube"), @("DOR pump", "Reactor valve"), @("P-trap", "DOR O-ring"),
            @("Exhaust pipe", "")
        )
        $diskPositions = @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")
        for ($i = 0; $i -lt 12; $i++) {
            $pair = @("", "")
            if ($i -lt (Get-CollectionCount $pairs)) { $pair = Get-CollectionItem $pairs $i }
            $leftName = Get-SafeRowCell $pair 0
            $rightName = Get-SafeRowCell $pair 1
            Set-ExcelCell $ws $row 1 $leftName $false 8
            Set-ExcelCell $ws $row 2 (Get-PdfItemCell $ItemRows $leftName 1) $false 8
            Merge-ExcelCells $ws $row 3 4 (Get-PdfItemCell $ItemRows $leftName 2) $false 8 $false
            Set-ExcelCell $ws $row 5 $rightName $false 8
            Set-ExcelCell $ws $row 6 (Get-PdfItemCell $ItemRows $rightName 1) $false 8
            Merge-ExcelCells $ws $row 7 8 (Get-PdfItemCell $ItemRows $rightName 2) $false 8 $false
            $diskPos = [string](Get-CollectionItem $diskPositions $i)
            Set-ExcelCell $ws $row 9 $diskPos $false 8
            Merge-ExcelCells $ws $row 10 12 (Get-DiskSerialPdfSerial -Position $diskPos) $false 8 $false
            $row++
        }
        Set-ExcelCell $ws $row 1 "特氣切換" $false 8
        Merge-ExcelCells $ws $row 2 4 (Get-PdfItemCell $ItemRows "特氣是否切換" 1) $false 8 $false
        Merge-ExcelCells $ws $row 5 12 "" $false 8 $false
        $row++
        Set-ExcelCell $ws $row 1 "烘烤部件" $true 8
        Set-ExcelCell $ws $row 2 "Susceptor" $true 8
        Set-ExcelCell $ws $row 3 "Disk" $true 8
        Set-ExcelCell $ws $row 4 "Ceiling" $true 8
        Merge-ExcelCells $ws $row 5 12 "" $false 8 $false
        $row++
        Set-ExcelCell $ws $row 1 "烘烤日期" $false 8
        Set-ExcelCell $ws $row 2 (Get-PdfItemCell $ItemRows "Susceptor烘烤日期" 1) $false 8
        Set-ExcelCell $ws $row 3 (Get-PdfItemCell $ItemRows "Disk烘烤日期" 1) $false 8
        Set-ExcelCell $ws $row 4 (Get-PdfItemCell $ItemRows "Ceiling(G4)烘烤日期" 1) $false 8
        Merge-ExcelCells $ws $row 5 12 "" $false 8 $false
        $row++
        Set-ExcelCell $ws $row 1 "CBr4" $false 8
        Set-ExcelCell $ws $row 2 "第一瓶重量" $false 8
        Set-ExcelCell $ws $row 3 (Get-PdfItemCell $ItemRows "CBr4第一瓶重量(g)" 1) $false 8
        Set-ExcelCell $ws $row 4 "第二瓶重量" $false 8
        Set-ExcelCell $ws $row 5 (Get-PdfItemCell $ItemRows "CBr4第二瓶重量(g)" 1) $false 8
        Merge-ExcelCells $ws $row 6 12 "" $false 8 $false
        $row++
        Merge-ExcelCells $ws $row 1 12 "注意事項：確認下一 Batch 如 CBr4 第一瓶重量會 < 滿瓶的 30%，需下機執行敲打作業" $false 8 $false
        $row++
        Merge-ExcelCells $ws $row 1 12 "MFC default 設定值若小於 10 sccm，請一律設定為 10 sccm" $false 8 $false
        $row++
        Set-ExcelCell $ws $row 1 "其他維修項目" $false 8
        Merge-ExcelCells $ws $row 2 12 (Get-PdfItemCell $ItemRows "其他維修項目" 1) $false 8 $false
        Set-ExcelBorder $ws $itemStart 1 $row 12
        $row += 2

        Merge-ExcelCells $ws $row 1 12 "維修紀錄" $true 10 $true
        $logStart = $row
        $row++
        Set-ExcelCell $ws $row 1 "日期" $true 9
        Merge-ExcelCells $ws $row 2 7 "內容" $true 9 $true
        Merge-ExcelCells $ws $row 8 12 "備註" $true 9 $true
        $row++
        $excelLogRows = Normalize-TableRows $LogRows 3 5
        if ($Mode -eq "Full") { $excelLogRows = Normalize-TableRows $LogRows 3 8 }
        for ($i = 0; $i -lt (Get-CollectionCount $excelLogRows); $i++) {
            $logRow = Get-CollectionItem $excelLogRows $i
            Set-ExcelCell $ws $row 1 (Get-SafeRowCell $logRow 0) $false 9
            Merge-ExcelCells $ws $row 2 7 (Get-SafeRowCell $logRow 1) $false 9 $false
            Merge-ExcelCells $ws $row 8 12 (Get-SafeRowCell $logRow 2) $false 9 $false
            $ws.Rows.Item($row).RowHeight = 22
            $row++
        }
        Set-ExcelBorder $ws $logStart 1 ($row - 1) 12
        $row++

        Merge-ExcelCells $ws $row 1 12 "MO Source Change Record" $true 10 $true
        $moStart = $row
        $row++
        $moHeaders = @("Source", "設定瓶壓", "位置", "重量Off", "重量On", "重量T/W", "來料瓶壓", "實際瓶壓", "效率Off", "效率On", "退庫數量", "備註")
        for ($c = 0; $c -lt 12; $c++) { Set-ExcelCell $ws $row ($c + 1) ([string](Get-CollectionItem $moHeaders $c)) $true 8 }
        $row++
        # Get-MoSourceRowsForExport 已回傳完整列集合（每列代表一個 source、12 欄）。
        # 直接使用，不可再取 index 0，否則只有一筆 source 時 12 欄會被拆成直向 12 列。
        $excelMoRows = Get-MoSourceRowsForExport 5
        if ($Mode -eq "Full") { $excelMoRows = Get-MoSourceRowsForExport 8 }
        for ($i = 0; $i -lt (Get-CollectionCount $excelMoRows); $i++) {
            $moRow = Get-CollectionItem $excelMoRows $i
            for ($c = 0; $c -lt 12; $c++) {
                Set-ExcelCell $ws $row ($c + 1) (Get-SafeRowCell $moRow $c) $false 8
            }
            $ws.Rows.Item($row).RowHeight = 20
            $row++
        }
        Set-ExcelBorder $ws $moStart 1 ($row - 1) 12

        $used = $ws.UsedRange
        $used.HorizontalAlignment = -4131
        $used.VerticalAlignment = -4108
        $used.WrapText = $true
        # 匯出 PDF 以橫式 (landscape) 列印，避免 12 欄 MO Source 表格被壓縮或截斷。xlLandscape = 2。
        $ws.PageSetup.Orientation = 2
        $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1
        if ($Mode -eq "Full") { $ws.PageSetup.FitToPagesTall = 2 } else { $ws.PageSetup.FitToPagesTall = 1 }
        $ws.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
        $ws.PageSetup.RightMargin = $excel.InchesToPoints(0.25)
        $ws.PageSetup.TopMargin = $excel.InchesToPoints(0.25)
        $ws.PageSetup.BottomMargin = $excel.InchesToPoints(0.25)
        $ws.PageSetup.PrintArea = $used.Address()

        $wb.SaveAs($xlsxPath, 51)
        $wb.ExportAsFixedFormat(0, $PdfPath)
        Write-Audit "EXPORT_PDF_EXCEL" $Mode "" "" ($PdfPath + "; " + $xlsxPath) "Excel COM"
        return $PdfPath
    } finally {
        if ($null -ne $wb) { $wb.Close($false) | Out-Null }
        if ($null -ne $excel) { $excel.Quit() | Out-Null }
        if ($null -ne $ws) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null }
        if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
        if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Write-MaintenancePdf {
    param([string]$Mode)
    Ensure-NoOpenControlledEdit
    if ($Mode -eq "Full") {
        Ensure-FinalLocked
    }
    $outDir = Resolve-FullPath (Get-ControlText $script:txtOutputDir)
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir | Out-Null
    }
    $suffix = "_維修前使用.pdf"
    if ($Mode -eq "Full") { $suffix = "_完整維修紀錄.pdf" }
    $pdfPath = Join-Path $outDir ((Get-RecordBaseName) + $suffix)

    $logRows = Normalize-TableRows (Get-GridRows $script:gridMaintenanceLog) 3 0
    $moRowsWrapper = Get-MoSourceRowsForExport 0
    $moRows = $moRowsWrapper
    $itemRows = Get-MaintenanceItemRows

    try {
        return (Write-MaintenancePdfViaExcel $Mode $pdfPath $logRows $moRows $itemRows)
    } catch {
        Write-Audit "EXPORT_PDF_EXCEL_FALLBACK" $Mode "" "" $pdfPath $_.Exception.Message
    }

    $pageImages = New-Object System.Collections.ArrayList
    [void]$pageImages.Add((New-MaintenancePdfPageImage $Mode 0 $logRows $moRows $itemRows))
    if ($Mode -eq "Full" -and ((Get-CollectionCount $logRows) -gt 8 -or (Get-CollectionCount $moRows) -gt 8)) {
        [void]$pageImages.Add((New-MaintenancePdfPageImage $Mode 1 $logRows $moRows $itemRows))
    }
    Write-Audit "EXPORT_PDF" $Mode "" "" $pdfPath ""
    Write-ImagePdf $pdfPath $pageImages 900 1400
    return $pdfPath
}

function Open-FilePreview {
    param([string]$Path)
    if ($null -eq $Path -or $Path.Trim().Length -eq 0) { return }
    if (-not (Test-Path $Path)) {
        throw ("找不到檔案：" + $Path)
    }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Path
    $psi.UseShellExecute = $true
    [System.Diagnostics.Process]::Start($psi) | Out-Null
}

function Ask-OpenPdfPreview {
    param([string]$Path)
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ("PDF 已輸出：" + "`r`n" + $Path + "`r`n`r`n是否立即開啟預覽？"),
        "PDF 匯出完成",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
        Open-FilePreview $Path
    }
}

$script:itemLocked = $false
$script:recordLocked = $false
$script:itemEditReason = ""
$script:recordEditReason = ""
$script:itemBaselineSnapshot = New-Object System.Collections.ArrayList
$script:recordBaselineSnapshot = New-Object System.Collections.ArrayList
$script:itemEditControls = New-Object System.Collections.ArrayList
$script:recordCodeFinalized = $false
$script:finalRecordCode = ""
$script:lastDraftPath = ""
$script:txtRecordCode = $null
$script:txtToolName = $null
$script:txtAttachmentPaths = $null
$script:gridDraftList = $null
$script:lblDraftListStatus = $null
$script:gridDiskSerials = $null
$script:txtSusceptorBakeDate = $null
$script:txtDiskBakeDate = $null
$script:txtItemNotes = $null
$script:btnLockItemsRecords = $null
$script:btnUnlockItemsRecords = $null
$script:lblItemLockStatusRecords = $null
$script:autoSaveTimer = $null

$form = New-Object System.Windows.Forms.Form
$form.Text = "MOCVD 維修紀錄 GUI"
$form.Size = New-Object System.Drawing.Size(1120, 780)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft JhengHei", 9)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 700)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = "Fill"
$form.Controls.Add($tabs)

$tabBasic = New-Object System.Windows.Forms.TabPage
$tabBasic.Text = "基本資料"
$tabItems = New-Object System.Windows.Forms.TabPage
$tabItems.Text = "維修項目"
$tabRecords = New-Object System.Windows.Forms.TabPage
$tabRecords.Text = "紀錄/MO source"
$tabOutput = New-Object System.Windows.Forms.TabPage
$tabOutput.Text = "AI輸出"
$tabQuery = New-Object System.Windows.Forms.TabPage
$tabQuery.Text = "查詢/匯出"
[void]$tabs.TabPages.Add($tabBasic)
[void]$tabs.TabPages.Add($tabItems)
[void]$tabs.TabPages.Add($tabRecords)
[void]$tabs.TabPages.Add($tabOutput)
[void]$tabs.TabPages.Add($tabQuery)

$panelBasic = New-Object System.Windows.Forms.Panel
$panelBasic.Dock = "Fill"
$panelBasic.AutoScroll = $true
$tabBasic.Controls.Add($panelBasic)

$panelBasic.Controls.Add((New-Label "PM主題" 20 24 130))
$script:txtPmTopic = New-TextBox "Change Disk/Cover/Star/Ceiling" 160 24 540
$script:txtPmTopic.ReadOnly = $true
$panelBasic.Controls.Add($script:txtPmTopic)
$panelBasic.Controls.Add((New-Label "PM類別" 20 58 130))
$script:cmbPmCategory = New-ComboBox @("Scheduled PM", "Unscheduled PM") "" 160 58 220
$panelBasic.Controls.Add($script:cmbPmCategory)
$panelBasic.Controls.Add((New-Label "Batch" 20 92 130))
$script:txtBatch = New-TextBox "" 160 92 220
$panelBasic.Controls.Add($script:txtBatch)
$panelBasic.Controls.Add((New-Label "產品" 420 92 110))
$script:txtProduct = New-TextBox "" 540 92 220
$panelBasic.Controls.Add($script:txtProduct)
$panelBasic.Controls.Add((New-Label "RunID" 20 126 130))
$script:txtRunId = New-TextBox "" 160 126 220
$panelBasic.Controls.Add($script:txtRunId)
$panelBasic.Controls.Add((New-Label "PM會議人員" 420 126 110))
$script:txtMeetingPeople = New-TextBox "" 540 126 220
$panelBasic.Controls.Add($script:txtMeetingPeople)
$panelBasic.Controls.Add((New-Label "PM日期" 20 160 130))
$script:txtPmDate = New-TextBox "" 160 160 220
$panelBasic.Controls.Add($script:txtPmDate)
$panelBasic.Controls.Add((New-Label "班別" 420 160 110))
$script:cmbShift = New-ComboBox @("日班", "夜班", "中班", "早班") "" 540 160 160
$panelBasic.Controls.Add($script:cmbShift)
$panelBasic.Controls.Add((New-Label "機台名稱" 20 206 130))
$script:txtToolName = New-TextBox "" 160 206 220
$panelBasic.Controls.Add($script:txtToolName)
$panelBasic.Controls.Add((New-Label "維修紀錄編碼" 420 206 110))
$script:txtRecordCode = New-TextBox "" 540 206 300
$script:txtRecordCode.ReadOnly = $true
$panelBasic.Controls.Add($script:txtRecordCode)
$panelBasic.Controls.Add((New-Label "維修人員" 20 240 130))
$script:txtMaintPeople = New-TextBox "" 160 240 220
$panelBasic.Controls.Add($script:txtMaintPeople)
$panelBasic.Controls.Add((New-Label "維修日期" 420 240 110))
$script:txtMaintDate = New-TextBox "" 540 240 220
$panelBasic.Controls.Add($script:txtMaintDate)
$panelBasic.Controls.Add((New-Label "附件/照片路徑" 20 274 130))
$script:txtAttachmentPaths = New-TextBox "" 160 274 620
$panelBasic.Controls.Add($script:txtAttachmentPaths)
$btnBrowseAttach = New-Object System.Windows.Forms.Button
$btnBrowseAttach.Text = "選擇"
$btnBrowseAttach.Location = New-Object System.Drawing.Point(800, 272)
$btnBrowseAttach.Size = New-Object System.Drawing.Size(80, 28)
$panelBasic.Controls.Add($btnBrowseAttach)
$btnCheckAttach = New-Object System.Windows.Forms.Button
$btnCheckAttach.Text = "檢查路徑"
$btnCheckAttach.Location = New-Object System.Drawing.Point(890, 272)
$btnCheckAttach.Size = New-Object System.Drawing.Size(100, 28)
$panelBasic.Controls.Add($btnCheckAttach)
$btnNewRecord = New-Object System.Windows.Forms.Button
$btnNewRecord.Text = "新建紀錄"
$btnNewRecord.Location = New-Object System.Drawing.Point(480, 58)
$btnNewRecord.Size = New-Object System.Drawing.Size(110, 30)
$panelBasic.Controls.Add($btnNewRecord)
$btnSaveDraft = New-Object System.Windows.Forms.Button
$btnSaveDraft.Text = "暫存草稿"
$btnSaveDraft.Location = New-Object System.Drawing.Point(600, 58)
$btnSaveDraft.Size = New-Object System.Drawing.Size(110, 30)
$panelBasic.Controls.Add($btnSaveDraft)
$btnLoadDraft = New-Object System.Windows.Forms.Button
$btnLoadDraft.Text = "讀取草稿"
$btnLoadDraft.Location = New-Object System.Drawing.Point(700, 58)
$btnLoadDraft.Size = New-Object System.Drawing.Size(110, 30)
$panelBasic.Controls.Add($btnLoadDraft)

$notice = New-MultiTextBox "日期請用 yyyy-MM-dd。`r`n新建紀錄時先產生 MR-日期；維修項目完成鎖定後才更新為 MR-日期-機台名稱-RunID。`r`n未完成維修項目可用暫存草稿保存，讀取草稿後可繼續編輯。`r`nPM主題由維修項目中的 Disk/Cover/Star/Ceiling 自動產生。請勿填入帳號、密碼、內網連線資訊或個資。" 20 320 1010 60
$notice.ReadOnly = $true
$panelBasic.Controls.Add($notice)
$panelBasic.Controls.Add((New-Label "草稿列表" 20 392 120))
$script:lblDraftListStatus = New-Label "草稿列表：尚未載入" 120 392 260
$panelBasic.Controls.Add($script:lblDraftListStatus)
$btnDeleteDraft = New-Object System.Windows.Forms.Button
$btnDeleteDraft.Text = "刪除選取草稿"
$btnDeleteDraft.Location = New-Object System.Drawing.Point(880, 388)
$btnDeleteDraft.Size = New-Object System.Drawing.Size(150, 30)
$panelBasic.Controls.Add($btnDeleteDraft)
$script:gridDraftList = New-DataGrid @(
    @("檔名", "FileName", 420),
    @("最後修改", "ModifiedTime", 180),
    @("大小(KB)", "SizeKb", 90),
    @("完整路徑", "Path", 300)
) 20 428 1010 170
$script:gridDraftList.AllowUserToAddRows = $false
$script:gridDraftList.AllowUserToDeleteRows = $false
$script:gridDraftList.ReadOnly = $true
$script:gridDraftList.RowHeadersWidth = 24
$script:gridDraftList.Columns["Path"].Visible = $false
$panelBasic.Controls.Add($script:gridDraftList)
$script:txtPmDate.Add_TextChanged({ Refresh-RecordCode })
$script:txtToolName.Add_TextChanged({ Refresh-RecordCode })
$script:txtRunId.Add_TextChanged({ Refresh-RecordCode })

$panelItems = New-Object System.Windows.Forms.Panel
$panelItems.Dock = "Fill"
$panelItems.AutoScroll = $true
$tabItems.Controls.Add($panelItems)

$script:chkDisk = New-Object System.Windows.Forms.CheckBox
$script:chkDisk.Text = "Disk/SN"
$script:chkDisk.Checked = $false
$script:chkDisk.Location = New-Object System.Drawing.Point(20, 94)
$script:chkDisk.Size = New-Object System.Drawing.Size(100, 24)
$panelItems.Controls.Add($script:chkDisk)
$script:chkCover = New-Object System.Windows.Forms.CheckBox
$script:chkCover.Text = "Cover"
$script:chkCover.Checked = $false
$script:chkCover.Location = New-Object System.Drawing.Point(20, 24)
$script:chkCover.Size = New-Object System.Drawing.Size(100, 24)
$panelItems.Controls.Add($script:chkCover)
$script:txtCoverSerial = New-TextBox "" 180 24 220
$panelItems.Controls.Add($script:txtCoverSerial)
$script:chkStar = New-Object System.Windows.Forms.CheckBox
$script:chkStar.Text = "Star"
$script:chkStar.Checked = $false
$script:chkStar.Location = New-Object System.Drawing.Point(520, 24)
$script:chkStar.Size = New-Object System.Drawing.Size(80, 24)
$panelItems.Controls.Add($script:chkStar)
$script:txtStarSerial = New-TextBox "" 680 24 220
$panelItems.Controls.Add($script:txtStarSerial)

$script:chkCeiling = New-Object System.Windows.Forms.CheckBox
$script:chkCeiling.Text = "Ceiling"
$script:chkCeiling.Checked = $false
$script:chkCeiling.Location = New-Object System.Drawing.Point(20, 54)
$script:chkCeiling.Size = New-Object System.Drawing.Size(150, 24)
$panelItems.Controls.Add($script:chkCeiling)
$script:txtCeiling = New-TextBox "" 180 54 220
$panelItems.Controls.Add($script:txtCeiling)
$panelItems.Controls.Add((New-Label "Disk 位置/序號" 180 80 180))
$script:gridDiskSerials = New-DataGrid @(
    @("位置", "Position", 120),
    @("序號", "SerialNumber", 560)
) 180 104 780 86
$script:gridDiskSerials.RowHeadersWidth = 24
$panelItems.Controls.Add($script:gridDiskSerials)
Add-DefaultDiskRows
$btnAddDisk = New-Object System.Windows.Forms.Button
$btnAddDisk.Text = "增加 Disk 列"
$btnAddDisk.Location = New-Object System.Drawing.Point(970, 104)
$btnAddDisk.Size = New-Object System.Drawing.Size(110, 28)
$panelItems.Controls.Add($btnAddDisk)
$btnDelDisk = New-Object System.Windows.Forms.Button
$btnDelDisk.Text = "刪除 Disk 列"
$btnDelDisk.Location = New-Object System.Drawing.Point(970, 144)
$btnDelDisk.Size = New-Object System.Drawing.Size(110, 28)
$panelItems.Controls.Add($btnDelDisk)

$y = 200
$names = @(
    @("Pull down", "chkPullDown", "txtPullDown"), @("Susceptor", "chkSusceptor", "txtSusceptor"), @("Collector ring", "chkCollectorRing", "txtCollectorRing"),
    @("Exhaust tube", "chkExhaustTube", "txtExhaustTube"), @("Lightpipe", "chkLightpipe", "txtLightpipe"), @("Throttle Valve", "chkThrottleValve", "txtThrottleValve"),
    @("Maint pump", "chkMaintPump", "txtMaintPump"), @("25mbar C.V.", "chkCv25", "txtCv25"), @("Glovebox pump", "chkGloveboxPump", "txtGloveboxPump"),
    @("CT1000", "chkCt1000", "txtCt1000"), @("DOR pump", "chkDorPump", "txtDorPump"), @("Y-tube", "chkYTube", "txtYTube"),
    @("P-trap", "chkPTrap", "txtPTrap"), @("Reactor valve", "chkReactorValve", "txtReactorValve"), @("Exhaust pipe", "chkExhaustPipe", "txtExhaustPipe"),
    @("DOR O-ring", "chkDorORing", "txtDorORing")
)
for ($i = 0; $i -lt $names.Count; $i++) {
    $col = $i % 2
    $row = [int][Math]::Floor($i / 2)
    $xLabel = 20 + ($col * 500)
    $xBox = 180 + ($col * 500)
    $yy = $y + ($row * 28)
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $names[$i][0]
    $chk.Location = New-Object System.Drawing.Point($xLabel, $yy)
    $chk.Size = New-Object System.Drawing.Size(150, 24)
    Set-Variable -Name ($names[$i][1]) -Scope Script -Value $chk
    $panelItems.Controls.Add($chk)
    Set-Variable -Name ($names[$i][2]) -Scope Script -Value (New-TextBox "" $xBox $yy 200)
    $panelItems.Controls.Add((Get-Variable -Name ($names[$i][2]) -Scope Script).Value)
}

$y2 = $y + 8 * 26 + 20
$panelItems.Controls.Add((New-Label "特氣是否切換" 20 $y2 150))
$script:cmbSpecialGas = New-ComboBox @("", "Y", "N", "NA") "" 180 $y2 120
$panelItems.Controls.Add($script:cmbSpecialGas)
$y2 += 26
$panelItems.Controls.Add((New-Label "Susceptor烘烤日期" 20 $y2 150))
$script:txtSusceptorBakeDate = New-TextBox "" 180 $y2 160
$panelItems.Controls.Add($script:txtSusceptorBakeDate)
$panelItems.Controls.Add((New-Label "Disk烘烤日期" 360 $y2 130))
$script:txtDiskBakeDate = New-TextBox "" 500 $y2 160
$panelItems.Controls.Add($script:txtDiskBakeDate)
$panelItems.Controls.Add((New-Label "Ceiling(G4)烘烤日期" 680 $y2 150))
$script:txtCeilingBakeDate = New-TextBox "" 840 $y2 160
$panelItems.Controls.Add($script:txtCeilingBakeDate)

$y2 += 26
$panelItems.Controls.Add((New-Label "CBr4第一瓶重量(g)" 20 $y2 150))
$script:txtCbr4Bottle1 = New-TextBox " / " 180 $y2 120
$panelItems.Controls.Add($script:txtCbr4Bottle1)
$panelItems.Controls.Add((New-Label "CBr4第二瓶重量(g)" 360 $y2 150))
$script:txtCbr4Bottle2 = New-TextBox " / " 520 $y2 120
$panelItems.Controls.Add($script:txtCbr4Bottle2)

$y2 += 38
$script:txtItemNotes = New-MultiTextBox "注意事項：確認下一 Batch 如 CBr4 第一瓶重量會 < 滿瓶的 30%，需下機執行敲打作業`r`nMFC default 設定值若小於 10 sccm，請一律設定為 10 sccm" 20 $y2 940 50
$script:txtItemNotes.ReadOnly = $true
$panelItems.Controls.Add($script:txtItemNotes)

$y2 += 62
$panelItems.Controls.Add((New-Label "其他維修項目" 20 $y2 150))
$script:txtOtherItems = New-MultiTextBox "" 180 $y2 780 60
$panelItems.Controls.Add($script:txtOtherItems)

$script:chkDisk.Add_CheckedChanged({ Refresh-PmTopic })
$script:chkCover.Add_CheckedChanged({ Refresh-PmTopic })
$script:chkStar.Add_CheckedChanged({ Refresh-PmTopic })
$script:chkCeiling.Add_CheckedChanged({ Refresh-PmTopic })
$btnAddDisk.Add_Click({
    Add-GridRow $script:gridDiskSerials @("", "")
})
$btnDelDisk.Add_Click({
    if ($script:gridDiskSerials.SelectedCells.Count -gt 0) {
        $rowIndex = $script:gridDiskSerials.SelectedCells[0].RowIndex
        if (-not $script:gridDiskSerials.Rows[$rowIndex].IsNewRow) {
            $script:gridDiskSerials.Rows.RemoveAt($rowIndex)
        }
    }
})

$btnPrePdf = New-Object System.Windows.Forms.Button
$btnPrePdf.Text = "匯出維修前PDF"
$btnPrePdf.Location = New-Object System.Drawing.Point(20, ($y2 + 86))
$btnPrePdf.Size = New-Object System.Drawing.Size(150, 32)
$panelItems.Controls.Add($btnPrePdf)
$script:btnLockItems = New-Object System.Windows.Forms.Button
$script:btnLockItems.Text = "維修項目完成鎖定"
$script:btnLockItems.Location = New-Object System.Drawing.Point(190, ($y2 + 86))
$script:btnLockItems.Size = New-Object System.Drawing.Size(150, 32)
$panelItems.Controls.Add($script:btnLockItems)
$script:btnUnlockItems = New-Object System.Windows.Forms.Button
$script:btnUnlockItems.Text = "管制修改維修項目"
$script:btnUnlockItems.Location = New-Object System.Drawing.Point(360, ($y2 + 86))
$script:btnUnlockItems.Size = New-Object System.Drawing.Size(150, 32)
$script:btnUnlockItems.Enabled = $false
$panelItems.Controls.Add($script:btnUnlockItems)
$script:lblItemLockStatus = New-Label "維修項目狀態：編輯中" 530 ($y2 + 90) 260
$panelItems.Controls.Add($script:lblItemLockStatus)

foreach ($ctrl in @(
    $script:chkDisk, $script:chkCover, $script:txtCoverSerial, $script:chkStar, $script:txtStarSerial,
    $script:chkCeiling, $script:txtCeiling, $script:gridDiskSerials, $btnAddDisk, $btnDelDisk,
    $script:chkPullDown, $script:txtPullDown, $script:chkSusceptor, $script:txtSusceptor,
    $script:chkCollectorRing, $script:txtCollectorRing, $script:chkExhaustTube, $script:txtExhaustTube,
    $script:chkLightpipe, $script:txtLightpipe, $script:chkThrottleValve, $script:txtThrottleValve,
    $script:chkMaintPump, $script:txtMaintPump, $script:chkCv25, $script:txtCv25,
    $script:chkGloveboxPump, $script:txtGloveboxPump, $script:chkCt1000, $script:txtCt1000,
    $script:chkDorPump, $script:txtDorPump, $script:chkYTube, $script:txtYTube,
    $script:chkPTrap, $script:txtPTrap, $script:chkReactorValve, $script:txtReactorValve,
    $script:chkExhaustPipe, $script:txtExhaustPipe, $script:chkDorORing, $script:txtDorORing,
    $script:cmbSpecialGas, $script:txtSusceptorBakeDate, $script:txtDiskBakeDate, $script:txtCeilingBakeDate,
    $script:txtCbr4Bottle1, $script:txtCbr4Bottle2, $script:txtOtherItems
)) {
    [void]$script:itemEditControls.Add($ctrl)
}

$panelRecords = New-Object System.Windows.Forms.Panel
$panelRecords.Dock = "Fill"
$panelRecords.AutoScroll = $true
$tabRecords.Controls.Add($panelRecords)

$panelRecords.Controls.Add((New-Label "維修紀錄" 20 20 160))
$script:gridMaintenanceLog = New-DataGrid @(
    @("日期", "Date", 120),
    @("內容", "Content", 430),
    @("備註", "Note", 430)
) 20 52 1010 180
$panelRecords.Controls.Add($script:gridMaintenanceLog)
Add-GridRow $script:gridMaintenanceLog @("-", "", "")
$btnAddLog = New-Object System.Windows.Forms.Button
$btnAddLog.Text = "增加維修紀錄列"
$btnAddLog.Location = New-Object System.Drawing.Point(20, 244)
$btnAddLog.Size = New-Object System.Drawing.Size(140, 30)
$panelRecords.Controls.Add($btnAddLog)
$btnDelLog = New-Object System.Windows.Forms.Button
$btnDelLog.Text = "刪除選取列"
$btnDelLog.Location = New-Object System.Drawing.Point(176, 244)
$btnDelLog.Size = New-Object System.Drawing.Size(110, 30)
$panelRecords.Controls.Add($btnDelLog)
$btnPrePdfRecords = New-Object System.Windows.Forms.Button
$btnPrePdfRecords.Text = "匯出維修前PDF"
$btnPrePdfRecords.Location = New-Object System.Drawing.Point(320, 244)
$btnPrePdfRecords.Size = New-Object System.Drawing.Size(150, 30)
$panelRecords.Controls.Add($btnPrePdfRecords)
$script:btnLockItemsRecords = New-Object System.Windows.Forms.Button
$script:btnLockItemsRecords.Text = "維修項目完成鎖定"
$script:btnLockItemsRecords.Location = New-Object System.Drawing.Point(490, 244)
$script:btnLockItemsRecords.Size = New-Object System.Drawing.Size(150, 30)
$panelRecords.Controls.Add($script:btnLockItemsRecords)
$script:btnUnlockItemsRecords = New-Object System.Windows.Forms.Button
$script:btnUnlockItemsRecords.Text = "管制修改維修項目"
$script:btnUnlockItemsRecords.Location = New-Object System.Drawing.Point(660, 244)
$script:btnUnlockItemsRecords.Size = New-Object System.Drawing.Size(150, 30)
$script:btnUnlockItemsRecords.Enabled = $false
$panelRecords.Controls.Add($script:btnUnlockItemsRecords)
$script:lblItemLockStatusRecords = New-Label "維修項目狀態：編輯中" 830 248 220
$panelRecords.Controls.Add($script:lblItemLockStatusRecords)

$baseY = 300
$panelRecords.Controls.Add((New-Label "MO source更換紀錄" 20 $baseY 180))
$script:gridMoSource = New-DataGrid @(
    @("Source", "Source", 80),
    @("設定瓶壓", "SetPressure", 70),
    @("位置", "Position", 60),
    @("重量Off", "WeightOff", 70),
    @("重量On", "WeightOn", 70),
    @("重量T/W", "WeightTW", 70),
    @("來料瓶壓", "IncomingPressure", 75),
    @("實際瓶壓", "ActualPressure", 75),
    @("效率Off", "EfficiencyOff", 70),
    @("效率On", "EfficiencyOn", 70),
    @("退庫數量(瓶)", "ReturnBottleCount", 85),
    @("備註", "Note", 160)
) 20 ($baseY + 32) 1010 180
$panelRecords.Controls.Add($script:gridMoSource)
$btnAddMo = New-Object System.Windows.Forms.Button
$btnAddMo.Text = "增加 MO source 列"
$btnAddMo.Location = New-Object System.Drawing.Point(20, ($baseY + 224))
$btnAddMo.Size = New-Object System.Drawing.Size(150, 30)
$panelRecords.Controls.Add($btnAddMo)
$btnDelMo = New-Object System.Windows.Forms.Button
$btnDelMo.Text = "刪除選取列"
$btnDelMo.Location = New-Object System.Drawing.Point(186, ($baseY + 224))
$btnDelMo.Size = New-Object System.Drawing.Size(110, 30)
$panelRecords.Controls.Add($btnDelMo)
$btnFullPdf = New-Object System.Windows.Forms.Button
$btnFullPdf.Text = "匯出完整維修PDF"
$btnFullPdf.Location = New-Object System.Drawing.Point(320, ($baseY + 224))
$btnFullPdf.Size = New-Object System.Drawing.Size(160, 30)
$panelRecords.Controls.Add($btnFullPdf)
$script:btnLockRecords = New-Object System.Windows.Forms.Button
$script:btnLockRecords.Text = "維修紀錄完成鎖定"
$script:btnLockRecords.Location = New-Object System.Drawing.Point(500, ($baseY + 224))
$script:btnLockRecords.Size = New-Object System.Drawing.Size(150, 30)
$panelRecords.Controls.Add($script:btnLockRecords)
$script:btnUnlockRecords = New-Object System.Windows.Forms.Button
$script:btnUnlockRecords.Text = "管制修改維修紀錄"
$script:btnUnlockRecords.Location = New-Object System.Drawing.Point(670, ($baseY + 224))
$script:btnUnlockRecords.Size = New-Object System.Drawing.Size(150, 30)
$script:btnUnlockRecords.Enabled = $false
$panelRecords.Controls.Add($script:btnUnlockRecords)
$script:lblRecordLockStatus = New-Label "維修紀錄狀態：編輯中" 840 ($baseY + 228) 220
$panelRecords.Controls.Add($script:lblRecordLockStatus)

$panelOutput = New-Object System.Windows.Forms.Panel
$panelOutput.Dock = "Fill"
$panelOutput.AutoScroll = $true
$tabOutput.Controls.Add($panelOutput)

$panelOutput.Controls.Add((New-Label "輸出資料夾" 20 24 120))
$script:txtOutputDir = New-TextBox (Resolve-FullPath $OutputDirectory) 150 24 720
$panelOutput.Controls.Add($script:txtOutputDir)
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "選擇"
$btnBrowse.Location = New-Object System.Drawing.Point(890, 22)
$btnBrowse.Size = New-Object System.Drawing.Size(80, 28)
$panelOutput.Controls.Add($btnBrowse)

$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = "更新AI預覽"
$btnPreview.Location = New-Object System.Drawing.Point(20, 68)
$btnPreview.Size = New-Object System.Drawing.Size(120, 32)
$panelOutput.Controls.Add($btnPreview)
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "儲存 JSON / CSV / TXT"
$btnSave.Location = New-Object System.Drawing.Point(160, 68)
$btnSave.Size = New-Object System.Drawing.Size(170, 32)
$panelOutput.Controls.Add($btnSave)

$script:txtPreview = New-MultiTextBox "" 20 118 960 480
$script:txtPreview.Font = New-Object System.Drawing.Font("Consolas", 9)
$panelOutput.Controls.Add($script:txtPreview)

$panelQuery = New-Object System.Windows.Forms.Panel
$panelQuery.Dock = "Fill"
$panelQuery.AutoScroll = $true
$tabQuery.Controls.Add($panelQuery)
$panelQuery.Controls.Add((New-Label "查詢關鍵字" 20 20 100))
$script:txtQueryFilter = New-TextBox "" 130 20 360
$panelQuery.Controls.Add($script:txtQueryFilter)
$btnRefreshCompleted = New-Object System.Windows.Forms.Button
$btnRefreshCompleted.Text = "查詢完成紀錄"
$btnRefreshCompleted.Location = New-Object System.Drawing.Point(510, 18)
$btnRefreshCompleted.Size = New-Object System.Drawing.Size(120, 28)
$panelQuery.Controls.Add($btnRefreshCompleted)
$btnExportCompleted = New-Object System.Windows.Forms.Button
$btnExportCompleted.Text = "匯出完成紀錄查詢"
$btnExportCompleted.Location = New-Object System.Drawing.Point(650, 18)
$btnExportCompleted.Size = New-Object System.Drawing.Size(150, 28)
$panelQuery.Controls.Add($btnExportCompleted)
$btnRefreshAudit = New-Object System.Windows.Forms.Button
$btnRefreshAudit.Text = "查詢 Log"
$btnRefreshAudit.Location = New-Object System.Drawing.Point(820, 18)
$btnRefreshAudit.Size = New-Object System.Drawing.Size(90, 28)
$panelQuery.Controls.Add($btnRefreshAudit)
$btnExportAudit = New-Object System.Windows.Forms.Button
$btnExportAudit.Text = "匯出 Log 查詢"
$btnExportAudit.Location = New-Object System.Drawing.Point(930, 18)
$btnExportAudit.Size = New-Object System.Drawing.Size(120, 28)
$panelQuery.Controls.Add($btnExportAudit)
$panelQuery.Controls.Add((New-Label "已完成維修紀錄" 20 58 180))
$btnOpenCompletedReadonly = New-Object System.Windows.Forms.Button
$btnOpenCompletedReadonly.Text = "唯讀開啟完成紀錄"
$btnOpenCompletedReadonly.Location = New-Object System.Drawing.Point(210, 54)
$btnOpenCompletedReadonly.Size = New-Object System.Drawing.Size(150, 28)
$panelQuery.Controls.Add($btnOpenCompletedReadonly)
$btnOpenCompletedPdf = New-Object System.Windows.Forms.Button
$btnOpenCompletedPdf.Text = "開啟PDF補印"
$btnOpenCompletedPdf.Location = New-Object System.Drawing.Point(380, 54)
$btnOpenCompletedPdf.Size = New-Object System.Drawing.Size(130, 28)
$panelQuery.Controls.Add($btnOpenCompletedPdf)
$script:gridCompletedRecords = New-DataGrid @() 20 86 1030 240
$script:gridCompletedRecords.ReadOnly = $true
$script:gridCompletedRecords.AllowUserToAddRows = $false
$script:gridCompletedRecords.AllowUserToDeleteRows = $false
$panelQuery.Controls.Add($script:gridCompletedRecords)
$panelQuery.Controls.Add((New-Label "Log / 稽核紀錄" 20 344 180))
$script:gridAuditLogs = New-DataGrid @() 20 372 1030 260
$script:gridAuditLogs.ReadOnly = $true
$script:gridAuditLogs.AllowUserToAddRows = $false
$script:gridAuditLogs.AllowUserToDeleteRows = $false
$panelQuery.Controls.Add($script:gridAuditLogs)

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "選擇內網維修紀錄資料夾"
    if (Test-Path $script:txtOutputDir.Text) {
        $dlg.SelectedPath = $script:txtOutputDir.Text
    }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:txtOutputDir.Text = $dlg.SelectedPath
    }
})

$btnBrowseAttach.Add_Click({
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = "選擇附件或照片"
        $dlg.Filter = "所有檔案 (*.*)|*.*"
        $dlg.Multiselect = $true
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $existing = $script:txtAttachmentPaths.Text.Trim()
        $selected = ($dlg.FileNames -join "; ")
        if ($existing.Length -gt 0) {
            $script:txtAttachmentPaths.Text = $existing + "; " + $selected
        } else {
            $script:txtAttachmentPaths.Text = $selected
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "選擇附件失敗", "OK", "Error") | Out-Null
    }
})

$btnCheckAttach.Add_Click({
    try {
        Show-AttachmentPathCheck
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "附件檢查失敗", "OK", "Error") | Out-Null
    }
})

$btnRefreshCompleted.Add_Click({
    try {
        $indexPath = Get-CompletedIndexPath
        Load-CsvFilesToGrid $script:gridCompletedRecords @($indexPath) $script:txtQueryFilter.Text
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "查詢完成紀錄失敗", "OK", "Error") | Out-Null
    }
})

$btnOpenCompletedReadonly.Add_Click({
    try {
        Open-SelectedCompletedRecordReadOnly
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "唯讀開啟失敗", "OK", "Error") | Out-Null
    }
})

$btnOpenCompletedPdf.Add_Click({
    try {
        Open-SelectedCompletedRecordPdf
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "開啟 PDF 失敗", "OK", "Error") | Out-Null
    }
})

$script:gridCompletedRecords.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    try {
        Open-CompletedRecordReadOnlyByRow $e.RowIndex
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "唯讀開啟失敗", "OK", "Error") | Out-Null
    }
})

$btnExportCompleted.Add_Click({
    try {
        $outDir = Resolve-FullPath $script:txtOutputDir.Text
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
        $path = Join-Path $outDir ("completed_records_query_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
        Export-GridToCsv $script:gridCompletedRecords $path
        Write-Audit "EXPORT_QUERY" "completed_records" "" "" $path $script:txtQueryFilter.Text
        [System.Windows.Forms.MessageBox]::Show(("完成紀錄查詢已匯出：" + "`r`n" + $path), "完成", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "匯出完成紀錄查詢失敗", "OK", "Error") | Out-Null
    }
})

$btnRefreshAudit.Add_Click({
    try {
        $paths = Get-AuditLogPaths
        $arr = @()
        foreach ($p in $paths) { $arr += [string]$p }
        Load-CsvFilesToGrid $script:gridAuditLogs $arr $script:txtQueryFilter.Text
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "查詢 Log 失敗", "OK", "Error") | Out-Null
    }
})

$btnExportAudit.Add_Click({
    try {
        $outDir = Resolve-FullPath $script:txtOutputDir.Text
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
        $path = Join-Path $outDir ("audit_logs_query_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
        Export-GridToCsv $script:gridAuditLogs $path
        Write-Audit "EXPORT_QUERY" "audit_logs" "" "" $path $script:txtQueryFilter.Text
        [System.Windows.Forms.MessageBox]::Show(("Log 查詢已匯出：" + "`r`n" + $path), "完成", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "匯出 Log 查詢失敗", "OK", "Error") | Out-Null
    }
})

$btnNewRecord.Add_Click({
    try {
        $answer = [System.Windows.Forms.MessageBox]::Show("新建紀錄會清空目前畫面所有輸入，並重設鎖定狀態。是否繼續？", "新建紀錄", "YesNo", "Warning")
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        New-MaintenanceRecord
        $tabs.SelectedTab = $tabBasic
        [System.Windows.Forms.MessageBox]::Show(("已新建紀錄：" + "`r`n" + (Get-RecordCode)), "完成", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "新建紀錄失敗", "OK", "Error") | Out-Null
    }
})

$btnSaveDraft.Add_Click({
    try {
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Title = "暫存未完成維修紀錄"
        $dlg.Filter = "維修紀錄草稿 (*.draft.csv)|*.draft.csv|CSV (*.csv)|*.csv"
        $dlg.InitialDirectory = Get-DraftDirectory
        $dlg.FileName = [System.IO.Path]::GetFileName((Get-DefaultDraftPath))
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Save-DraftFile $dlg.FileName
        Refresh-DraftList
        [System.Windows.Forms.MessageBox]::Show(("草稿已暫存：" + "`r`n" + $dlg.FileName), "完成", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "暫存草稿失敗", "OK", "Error") | Out-Null
    }
})

$btnLoadDraft.Add_Click({
    try {
        Refresh-DraftList
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "讀取草稿失敗", "OK", "Error") | Out-Null
    }
})

$script:gridDraftList.Add_CellDoubleClick({
    param($sender, $e)
    try {
        if ($e.RowIndex -lt 0) { return }
        Open-SelectedDraft
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "讀取草稿失敗", "OK", "Error") | Out-Null
    }
})

$btnDeleteDraft.Add_Click({
    try {
        Delete-SelectedDraft
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "刪除草稿失敗", "OK", "Error") | Out-Null
    }
})

$btnAddLog.Add_Click({
    Add-GridRow $script:gridMaintenanceLog @("", "", "")
})

$btnDelLog.Add_Click({
    if ($script:gridMaintenanceLog.SelectedCells.Count -gt 0) {
        $rowIndex = $script:gridMaintenanceLog.SelectedCells[0].RowIndex
        if (-not $script:gridMaintenanceLog.Rows[$rowIndex].IsNewRow) {
            $script:gridMaintenanceLog.Rows.RemoveAt($rowIndex)
        }
    }
})

$btnAddMo.Add_Click({
    Add-MoSourceGridRow -Grid $script:gridMoSource -Row @("", "", "", "", "", "", "", "", "", "", "", "")
})

$btnDelMo.Add_Click({
    if ($script:gridMoSource.SelectedCells.Count -gt 0) {
        $rowIndex = $script:gridMoSource.SelectedCells[0].RowIndex
        if (-not $script:gridMoSource.Rows[$rowIndex].IsNewRow) {
            $script:gridMoSource.Rows.RemoveAt($rowIndex)
        }
    }
})

$btnPrePdfRecords.Add_Click({
    $btnPrePdf.PerformClick()
})

$script:btnLockItemsRecords.Add_Click({
    $script:btnLockItems.PerformClick()
})

$script:btnUnlockItemsRecords.Add_Click({
    $script:btnUnlockItems.PerformClick()
})

$script:btnLockItems.Add_Click({
    try {
        Refresh-PmTopic
        Assert-ItemLockReady
        if (-not (Confirm-PrelockChecklist "維修項目")) { return }
        if (-not (Confirm-LockSummary "維修項目")) { return }
        $newSnapshot = Get-ItemsSnapshot
        if ($script:itemEditReason.Length -gt 0) {
            Write-SnapshotDiffAudit "maintenance_item" $script:itemBaselineSnapshot $newSnapshot $script:itemEditReason
            Write-Audit "RELOCK" "maintenance_item" "" "" "" $script:itemEditReason
            $script:itemEditReason = ""
        } else {
            Write-Audit "FINALIZE" "maintenance_item" "" "" "" "維修項目完成鎖定"
        }
        Finalize-RecordCode
        $script:itemBaselineSnapshot = $newSnapshot
        $script:itemLocked = $true
        Set-ItemsLocked $true
        [System.Windows.Forms.MessageBox]::Show("維修項目已鎖定。後續若需修改，請使用管制修改並留下原因。", "完成", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "鎖定失敗", "OK", "Error") | Out-Null
    }
})

$script:btnUnlockItems.Add_Click({
    try {
        $reason = Prompt-AuditReason "管制修改維修項目"
        if ($reason.Length -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("未輸入原因，維修項目維持鎖定。", "未解鎖", "OK", "Warning") | Out-Null
            return
        }
        $script:itemBaselineSnapshot = Get-ItemsSnapshot
        $script:itemEditReason = $reason
        $script:itemLocked = $false
        Write-Audit "UNLOCK_FOR_CHANGE" "maintenance_item" "" "" "" $reason
        Set-ItemsLocked $false
        $script:btnLockItems.Text = "重新鎖定維修項目"
        $script:lblItemLockStatus.Text = "維修項目狀態：管制修改中"
        if ($null -ne $script:btnLockItemsRecords) { $script:btnLockItemsRecords.Text = "重新鎖定維修項目" }
        if ($null -ne $script:lblItemLockStatusRecords) { $script:lblItemLockStatusRecords.Text = "維修項目狀態：管制修改中" }
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "解鎖失敗", "OK", "Error") | Out-Null
    }
})

$script:btnLockRecords.Add_Click({
    try {
        Assert-RecordLockReady
        if (-not (Confirm-PrelockChecklist "維修紀錄 / MO Source")) { return }
        if (-not (Confirm-LockSummary "維修紀錄 / MO Source")) { return }
        $newSnapshot = Get-RecordsSnapshot
        $oldRecordBaselineSnapshot = $script:recordBaselineSnapshot
        $pendingRecordEditReason = $script:recordEditReason
        $script:recordEditReason = ""
        $autoSavedPaths = $null
        $script:recordLocked = $true
        Set-RecordsLocked $true
        try {
            $autoSavedPaths = Save-AiOutput
        } catch {
            $script:recordLocked = $false
            $script:recordEditReason = $pendingRecordEditReason
            Set-RecordsLocked $false
            if ($pendingRecordEditReason.Length -gt 0) {
                $script:btnLockRecords.Text = "重新鎖定維修紀錄"
                $script:lblRecordLockStatus.Text = "維修紀錄狀態：管制修改中"
            }
            throw
        }
        if ($pendingRecordEditReason.Length -gt 0) {
            Write-SnapshotDiffAudit "maintenance_record" $oldRecordBaselineSnapshot $newSnapshot $pendingRecordEditReason
            Write-Audit "RELOCK" "maintenance_record" "" "" "" $pendingRecordEditReason
        } else {
            Write-Audit "FINALIZE" "maintenance_record" "" "" "" "維修紀錄完成鎖定並自動輸出 AI 格式"
        }
        $script:recordBaselineSnapshot = $newSnapshot
        [System.Windows.Forms.MessageBox]::Show(("維修紀錄與 MO source 已鎖定，並已自動更新 AI 預覽與儲存 JSON / CSV / TXT。" + "`r`n`r`n" + ($autoSavedPaths -join "`r`n") + "`r`n`r`n後續若需修改，請使用管制修改並留下原因。"), "完成", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "鎖定失敗", "OK", "Error") | Out-Null
    }
})

$script:btnUnlockRecords.Add_Click({
    try {
        $reason = Prompt-AuditReason "管制修改維修紀錄"
        if ($reason.Length -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("未輸入原因，維修紀錄維持鎖定。", "未解鎖", "OK", "Warning") | Out-Null
            return
        }
        $script:recordBaselineSnapshot = Get-RecordsSnapshot
        $script:recordEditReason = $reason
        $script:recordLocked = $false
        Write-Audit "UNLOCK_FOR_CHANGE" "maintenance_record" "" "" "" $reason
        Set-RecordsLocked $false
        $script:btnLockRecords.Text = "重新鎖定維修紀錄"
        $script:lblRecordLockStatus.Text = "維修紀錄狀態：管制修改中"
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "解鎖失敗", "OK", "Error") | Out-Null
    }
})

$btnPrePdf.Add_Click({
    try {
        Refresh-PmTopic
        $path = Write-MaintenancePdf "Pre"
        Ask-OpenPdfPreview $path
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "PDF 匯出失敗", "OK", "Error") | Out-Null
    }
})

$btnFullPdf.Add_Click({
    try {
        Refresh-PmTopic
        $path = Write-MaintenancePdf "Full"
        Ask-OpenPdfPreview $path
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "PDF 匯出失敗", "OK", "Error") | Out-Null
    }
})

$btnPreview.Add_Click({
    try {
        [void](Update-AiPreview)
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "更新AI預覽失敗", "OK", "Error") | Out-Null
    }
})

$btnSave.Add_Click({
    try {
        $paths = Save-AiOutput
        [System.Windows.Forms.MessageBox]::Show("維修紀錄與 AI 格式已輸出。", "完成", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "儲存失敗", "OK", "Error") | Out-Null
    }
})

$form.Add_FormClosing({
    param($sender, $e)
    try {
        if ($null -ne $script:autoSaveTimer) { $script:autoSaveTimer.Stop() }
        if ($script:itemLocked -and $script:recordLocked) { return }
        $answer = [System.Windows.Forms.MessageBox]::Show("目前維修紀錄尚未全部完成鎖定。離開前是否暫存草稿？`r`n`r`n選 Yes：暫存後離開`r`n選 No：不暫存直接離開`r`n選 Cancel：取消離開", "離開前暫存", "YesNoCancel", "Warning")
        if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) {
            $e.Cancel = $true
            if ($null -ne $script:autoSaveTimer) { $script:autoSaveTimer.Start() }
            return
        }
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            $path = $script:lastDraftPath
            if ($null -eq $path -or $path.Trim().Length -eq 0) {
                $path = Get-DefaultDraftPath
            }
            Save-DraftFile $path
        }
    } catch {
        $e.Cancel = $true
        if ($null -ne $script:autoSaveTimer) { $script:autoSaveTimer.Start() }
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "離開前暫存失敗", "OK", "Error") | Out-Null
    }
})

$script:autoSaveTimer = New-Object System.Windows.Forms.Timer
$script:autoSaveTimer.Interval = 300000
$script:autoSaveTimer.Add_Tick({
    Save-AutoDraft
})
$script:autoSaveTimer.Start()

Refresh-PmTopic
Refresh-RecordCode
[void]$form.ShowDialog()
