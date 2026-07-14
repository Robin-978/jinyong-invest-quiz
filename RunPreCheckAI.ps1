#requires -version 5.1
<#
==============================================================================
 RunPreCheckAI.ps1
 Run 前 AI 製程風險診斷助手 (MOCVD / EPI Pre-Run Risk Diagnosis Assistant)
 Version : 1.0.0
 Target  : Windows PowerShell 5.1 + .NET Framework WinForms (免安裝、可共享資料夾部署)
==============================================================================
 依據文件: 磊晶製程工程師_AI導入SOP_Run前診斷.md v1.0

 定位 (SOP 1.1 / 4.x):
   - AI 只做: 整理、比對、提醒、建議、留痕
   - AI 不做: 不開 Run、不 Hold、不改 Recipe、不放行產品、不宣稱根因、
             不取代 SPC / MES / QMS 與工程師覆判
   - 本系統為離線規則引擎 + 模板摘要, 不外送任何資料 (零機密外洩風險)
   - Layer 3 LLM 介面僅預留, 預設停用

 資安 (SOP 7.x):
   - 僅處理去識別化資料 (RunID_Alias / ToolAlias / ProductFamily ...)
   - 內建 RawImport -> Import 去識別化轉換工具, 對照表僅存本機受控資料夾
   - 覆判留痕僅記錄 RoleCode, 不記錄真實姓名

 使用方式:
   一般啟動 : 雙擊 啟動.cmd
   自我測試 : powershell -NoProfile -ExecutionPolicy Bypass -File RunPreCheckAI.ps1 -SelfTest

 Changelog:
   1.0.15 新增: 資料表管理介面 — 設定頁「資料表管理」按鈕 (Admin/Engineer):
          RunSummary / ToolStability / Maintenance / ProductDrift /
          MeasurementTrend / HistoryCases 六張 Data\Import 資料表的
          檢視/編輯對話框:
          (1) 表格直接編輯; 底部空白列輸入即新增; 選列 Delete 刪除;
          (2) 儲存採原子寫檔, 寫檔前自動備份至 Backup\BeforeWrite,
              並寫入稽核留痕 (TABLE_EDITED + RoleCode);
          (3) 標頭以檔案現況為準 (保留自訂欄位), 檔案缺漏用範本標頭;
              未儲存切換/關閉會提醒; 六張表標頭定義集中共用;
          SelfTest 增至 129 項。
   1.0.14 修正 (依現場回饋): Log 分組鍵改以 StepLabel 為主 —
          StepLabel 才是真正對應到實際磊晶 Layer 的欄位 (同一 Layer 可
          橫跨多個 Step 編號; 不同 Run 的 Step 編號可能位移):
          (1) Get-StepCodeMeanTable 分組鍵優先順序改為
              StepLabel -> Step -> Stepcode_s -> StepCode_SP
              (C# 快速引擎與純 PS 退回模式一致; 無 StepLabel 的舊 Log
              自動退回原行為); StepLabel 不列入數值比對欄;
          (2) 比對原因與差異排行顯示文字改「Step/Layer」;
          SelfTest 增至 127 項 (StepLabel 分組/跨 Step 合併/退回模式)。
   1.0.13 修正/新增 (依現場回饋 3 項):
          (1) 趨勢圖停在舊 Run — Get-TrendChartData 改為 MeasurementTrend
              表與量測總檔目錄「聯集」: 表未更新 (如檔案被 Excel 鎖住致
              補列失敗) 時, 圖仍即時涵蓋量測總檔內全部 Run (取最近 N 筆);
              排序改以 Run 碼數字為主 (皆為數字時), 不受日期格式影響;
          (2) Dop1 變化量大掩蓋其他警訊 — 新增設定 LogCompareExcludeParams
              (萬用字元樣式, 分號分隔; 預設 "Dop1.*" 依現場指示):
              Log 比對時排除指定參數不列入, 報告註記排除清單以留痕;
              「編輯設定」對話框可直接修改 (亦可加 Hyd1.*;*.dp_SP 等);
          (3) 主診斷報告補「Log 比對差異排行」小節 (原僅前 Run Log 診斷
              對話框報告有排行, 主報告的「詳見差異排行」無處可看);
          SelfTest 增至 124 項。
   1.0.12 新增 (依現場回饋): 趨勢圖無 PL 資料時改畫 Rs —
          部分產品不量 PL (AGA/IGA 空白), 原趨勢圖 PL 線恆為空。
          (1) 新增 Get-TrendChartData (資料與 UI 分離, 可自我測試):
              以 MeasurementTrend 為主, 由量測總檔合併每 Run 的 Rs 平均
              (LEHI_RS); 該機台完全無 PL 資料時, 主線自動改畫 Rs,
              Y 軸標題與圖例同步切換; 表無資料時直接以量測總檔 Run 作圖;
          (2) 排序改 RunDate + Run 碼數字 (同日期不再亂序);
              缺值畫空點 (不再以 0 畫成假平線), 全空數列不顯示;
          SelfTest 增至 119 項。
   1.0.11 修正 (依現場回饋): 資料表 RunID_Alias 鍵格式容錯 —
          工程師於 ToolStability / Maintenance / ProductDrift 填入
          「機台+Run 碼」(如 MAT06261176) 時, 系統以「Run 碼」(261176)
          查表對不上, 導致已填資料仍報「無前一 Run 穩定度資料」。
          Find-RowByRun 增加 ToolAlias 參數, 同時接受: Run 碼、
          機台+Run 碼、機台_Run 碼、機台-Run 碼 (不分大小寫、去空白);
          欄位說明檔註明接受格式。SelfTest 增至 117 項。
   1.0.10 新增: 資料表自動補列 (Update-DerivedTables) — 回應「資料表都只有
          表頭」回饋:
          (1) RunSummary 由 Log 檔名目錄自動補列 (RunID_Alias/ToolAlias/
              ProductFamily/PreviousProductFamily/RunStartTime=檔案時間);
              MeasurementTrend 由量測總檔自動補列 (RunDate/PL 偏移);
              僅新增缺少的 RunID_Alias 列, 人工維護列一律不動;
              啟動與「重載資料」時自動執行;
          (2) 診斷整合與資料來源解耦: Run 已建入 RunSummary 後, 秒級 Log
              比對與量測總檔整合仍自動執行 (原本僅檔名模式才有);
              「由 Log 檔名建立」註記僅在 RunSummary 查無該 Run 時出現;
          (3) Run 下拉選單顯示「Run 碼 | 產品」(表格與檔名目錄一致),
              同一 Run 不重複列出;
          (4) ToolStability / Maintenance / HistoryCases 之 alarm / PM /
              案例資料不在 Log 或量測檔內, 仍需 MES 匯出或人工維護
              (未填時該面向依既有規則保守呈現), 欄位說明檔已註明;
          SelfTest 增至 113 項。
   1.0.9  修正/新增 (依現場回饋):
          (1) 修正 PS 5.1 StrictMode 下去識別化轉換報「找不到屬性 'Count'」—
              Import-CsvSafe 等函式回傳單列/空集合被解開為純量或 $null,
              呼叫端一律以 @() 包裝; 全檔盤點修正 4 處;
          (2) 去識別化轉換自動略過機台秒級 Log 檔 (MAT 檔名慣例, Log 比對
              直接讀取免轉換) 與量測總檔 (已為去識別後資料); 單一檔案失敗
              不中斷其餘檔案, 結果逐檔列出;
          (3) 啟動時自動建立六張資料表空白範本 (Data\Import, 僅標頭,
              行為中性) 與 Data\資料表欄位說明.txt, 方便建立
              RunSummary / ToolStability / Maintenance / ProductDrift /
              MeasurementTrend / HistoryCases; 既有檔案一律不動;
          (4) 新增門檻 LogCompareMinBaseAbs (預設 0=不啟用): Log 比對時
              基準平均絕對值低於此者略過 — 供濾除近零參數 (如 Dop 微量
              流量) 造成的數千 % 相對差異噪音, 由治理小組視機台調整;
          SelfTest 增至 109 項。
   1.0.8  新增: 量測總檔 (Measurement Master) 支援 — Import/RawImport 內含
          STRUCTURE,REACTOR,RUN_NO,POS_NO,PF,... 標頭之 CSV/TXT 自動偵測
          (可含表頭說明前言, 自動跳過並定位真實標頭):
          (1) Import-MeasurementMasterRows / Get-MeasurementRunCatalog —
              每片 wafer 資料依 機台(REACTOR 尾碼數字)+RUN_NO 彙總:
              Pass/Fail 統計、FAIL_NO 代碼、Rs/均勻性/Haze/LPD3+4/PL 平均;
          (2) Log 檔名目錄模式診斷時, ProductDrift 查無資料改以量測總檔為
              產品特性飄移來源 (取 RunId <= 本 Run 的最近量測 Run, 符合
              N-2 資料時效): Rs 偏差% 餵入既有 RsWarnPct 規則、均勻性 STD
              增幅 -> Worse、LPD3+4/Area_total 增幅 -> AOI Up、PL 波長偏移
              與連續同方向 Run 數餵入既有 PL 規則; PF 出現 Fail -> 高風險;
          (3) 新增門檻 MeasUnifWorsenPct/MeasDefectUpPct/PLShiftMinNm;
          SelfTest 增至 104 項 (前言跳過、Run 彙總、來源選取、Fail 高風險、
          Rs 偏差規則)。
   1.0.7  新增: 機台 Log 檔名目錄 — Import/RawImport 內的秒級 Log 檔
          (檔名慣例 <產品>.xxx..._MAT<機台2碼><Run碼>_<流水號>.csv, 如
          H01B9N.PRODUCT.B9N.P02069_M06_P(NGR)_MAT06261175_17155.csv):
          (1) Get-RunLogFileInfo / Get-RunLogCatalog 由檔名解析 機台(MAT06)/
              Run(261175)/產品(H01B9N), 診斷頁機台與 Run 下拉選單自動帶入
              (原本僅讀 RunSummary.csv, Input 只放 Log 檔時選單為空白);
          (2) Invoke-RunDiagnosis 支援 Log 檔名建立的 Run: RunSummary 查無此
              Run 時改由檔名目錄補基本資料, 並自動以同機台前一 Run Log 做
              StepCode 平均差異比對, 結果併入機台穩定度; 表格缺漏面向依既有
              規則以保守等級呈現並列於資料不足處;
          SelfTest 增至 99 項 (檔名解析、目錄掃描、Log Run 端到端診斷)。
   1.0.6  效能/修正: 去識別化轉換 (RawImport -> Import) 大幅加速並確保對照表落地 —
          (1) Get-DeidentAlias 類別編號改用計數快取 (原版每新增一個 alias 即全表
              掃描, 數萬個新 ID 為 O(n^2), 大檔需數十分鐘 -> 數秒);
          (2) Convert-RawImportFile 輸出改 StringBuilder + 行內 CSV 跳脫
              (移除陣列 += 與逐儲存格函式呼叫);
          (3) Add-LinesSafe 改批次附加 (原版逐行 Add-Content 每行開關檔案);
          (4) 寫檔順序修正: 先寫 DeidentMap.csv 對照表、再寫去識別化輸出檔,
              且對照表檔一定會建立 — 修正中途中斷時「有輸出檔、無對照表」
              導致 alias 無法回溯的問題;
          (5) 設定頁轉換按鈕執行中顯示等待游標, 完成訊息含耗時;
          SelfTest 增至 93 項 (編號接續、重複轉換不重複建檔、輸出一致)。
   1.0.5  效能: 前 Run Log 診斷 / StepCode 平均計算大幅加速 (大檔由 20+ 分鐘降至數秒) —
          (1) Get-StepCodeMeanTable 改用內嵌 C# 串流彙總引擎 (Add-Type, 離線編譯,
              免安裝; 逐字元 CSV 解析, 支援引號欄位), 取代 Import-Csv + 逐儲存格
              PowerShell 函式呼叫 (原每格 ~50us 開銷 x 數百萬格 = 數十分鐘);
          (2) C# 編譯失敗時自動退回純 PowerShell 串流模式 (StreamReader + 行內
              TryParse, 仍比原版快 10 倍以上), 行為與輸出完全相同;
          (3) Compare-StepCodeMeanTables 差異明細改用 List 收集 (避免陣列 += 的
              O(n^2) 複製); 診斷報告與視窗顯示比對耗時;
          (4) 診斷對話框執行中顯示等待游標與狀態, 避免誤認為當機;
          SelfTest 增至 90 項 (含快速引擎/退回模式一致性、引號欄位、空白行)。
   1.0.4  新增: 前 Run Log 分層診斷 — 因資料時效限制 (Run 前僅能取得上一 Run [N-1]
          機台秒級 Log; PL/XRD/Thickness/Rs/AOI 量測最快只到上上 Run [N-2]):
          (1) Get-StepCodeMeanTable / Compare-StepCodeMeanTables — N-1 vs N-2 秒級
              Log StepCode 平均差異比對, 並支援 Golden Run 基準線 (Data\Golden);
          (2) Invoke-PreRunLogDiagnosis — 產出前 Run Log 診斷報告 (Reports\LogPreCheck_*);
          (3) Invoke-RunDiagnosis 量測資料 (ProductDrift) 改以 N-2 為主要來源,
              報告註明量測來源 Run 與資料時效; 新增門檻 LogDeltaWarnPct/LogDeltaHighPct/
              GoldenDeltaWarnPct/GoldenDeltaHighPct/LogCompareMinSec/LogCompareTopN;
          (4) 診斷頁新增「前 Run Log 診斷」對話框; SelfTest 增至 85 項。
   1.0.3  新增: 以秒計 Run Log 檔 StepCode 平均值計算 (Export-StepCodeMeanReport);
          設定頁新增「Log 秒檔 StepCode 平均」按鈕, 輸出至 Reports\StepCode_Mean_*.csv;
          SelfTest 增加 StepCode 平均測試 (共 63 項)。
   1.0.2  新增: 設定頁「編輯設定」對話框 (僅 Admin) — 資料來源模式 (CSV/SQL)、
          SQL 啟用與連線字串、規則門檻值即時編輯, 存檔後立即生效並留痕;
          SelfTest 增加 config 儲存 / 回讀 / 缺欄位回填測試 (共 56 項)。
   1.0.1  修正: 事件閉包內 $script: 路徑變數解析為空 (LiteralPath 空字串錯誤,
          覆判紀錄頁載入失敗); Write-FileSafe/Add-LinesSafe 移除多餘 GetNewClosure;
          Save-HumanReview 未產報告時 Split-Path 空字串防呆。
          新增: 設定頁使用者管理 (Admin 新增帳號 / 停用啟用, 含最後一位 Admin 保護)。
   1.0.0  首版: CSV 匯入 + SQL 介面預留 + 去識別化轉換 + 規則引擎 +
          相似案例 + 模板診斷報告 + 燈號儀表板 + 趨勢圖 + 工程師覆判留痕
==============================================================================
#>
param(
    [switch]$SelfTest
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:IsGuiMode = (-not $SelfTest)

if ($script:IsGuiMode) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic
    [System.Windows.Forms.Application]::EnableVisualStyles()
}

$script:ChartingAvailable = $false
if ($script:IsGuiMode) {
    try {
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction Stop
        $script:ChartingAvailable = $true
    } catch { $script:ChartingAvailable = $false }
}

# ============================================================
# Globals
# ============================================================
$script:AppName   = "Run 前 AI 製程風險診斷助手"
$script:AIVersion = "1.0.15"

function Set-AppRootPaths {
    # 集中設定所有路徑; SelfTest 模式會改指到暫存資料夾, 不污染正式資料
    param([string]$Root)
    $script:AppRoot         = $Root
    $script:DataRoot        = Join-Path $script:AppRoot "Data"
    $script:ImportRoot      = Join-Path $script:DataRoot "Import"
    $script:RawImportRoot   = Join-Path $script:DataRoot "RawImport"
    $script:GoldenRoot      = Join-Path $script:DataRoot "Golden"
    $script:RecordRoot      = Join-Path $script:DataRoot "records"
    $script:LockRoot        = Join-Path $script:DataRoot "locks"
    $script:PendingRoot     = Join-Path $script:DataRoot "pending"
    $script:ReportRoot      = Join-Path $script:AppRoot "Reports"
    $script:BackupRoot      = Join-Path $script:AppRoot "Backup"
    $script:BackupBeforeWriteRoot = Join-Path $script:BackupRoot "BeforeWrite"
    $script:LogRoot         = Join-Path $script:AppRoot "Logs"
    $script:AppLogPath      = Join-Path $script:LogRoot "app.log"
    $script:ErrorLogPath    = Join-Path $script:LogRoot "error.log"
    $script:ConfigPath      = Join-Path $script:DataRoot "config.json"
    $script:UsersPath       = Join-Path $script:DataRoot "users.csv"
    $script:DeidentMapPath  = Join-Path $script:DataRoot "DeidentMap.csv"
    $script:HumanReviewPath = Join-Path $script:RecordRoot "HumanReview.csv"
}

$script:StartupRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($script:StartupRoot)) { $script:StartupRoot = (Get-Location).Path }
Set-AppRootPaths -Root $script:StartupRoot

# 共享狀態 (Pitfall 1: 事件閉包中不可依賴 $script: 變數, 一律用 hashtable)
$script:AppState = @{
    CurrentUser   = $null
    IsAdmin       = $false
    Config        = $null
    LastDiagnosis = $null
    LastReportPath = ""
    LogCatalog    = @()   # v1.0.7: Log 檔名目錄快取 (機台/Run/產品 由檔名解析)
}

# ============================================================
# Bootstrap
# ============================================================
function Initialize-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Initialize-AppFolders {
    foreach ($p in @($script:DataRoot, $script:ImportRoot, $script:RawImportRoot,
                     $script:GoldenRoot,
                     $script:RecordRoot, $script:LockRoot, $script:PendingRoot,
                     $script:ReportRoot, $script:BackupRoot,
                     $script:BackupBeforeWriteRoot, $script:LogRoot)) {
        Initialize-Directory $p
    }
}

# ============================================================
# Logging / audit
# ============================================================
function Write-AppLog {
    param([string]$Message)
    try {
        $line = "{0} [INFO] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
        Add-Content -LiteralPath $script:AppLogPath -Value $line -Encoding UTF8
    } catch { }
}

function Write-ErrorLog {
    param([string]$Message)
    try {
        $line = "{0} [ERROR] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
        Add-Content -LiteralPath $script:ErrorLogPath -Value $line -Encoding UTF8
    } catch { }
}

function Write-AuditLog {
    # 稽核留痕: 誰(RoleCode) 在何時 做了什麼
    param([string]$Action, [string]$Detail = "")
    $who = "SYSTEM"
    try {
        if ($null -ne $script:AppState.CurrentUser) {
            $who = [string](Get-ObjectPropertyValueSafe -Object $script:AppState.CurrentUser -PropertyName "RoleCode" -DefaultValue "UNKNOWN")
        }
    } catch { }
    Write-AppLog ("[AUDIT] {0} | {1} | {2}" -f $who, $Action, $Detail)
}

function Show-Info {
    param([string]$M)
    if ($script:IsGuiMode) { [System.Windows.Forms.MessageBox]::Show($M, "訊息", 'OK', 'Information') | Out-Null }
    else { Write-Host ("[INFO] " + $M) }
}
function Show-Warn {
    param([string]$M)
    if ($script:IsGuiMode) { [System.Windows.Forms.MessageBox]::Show($M, "提醒", 'OK', 'Warning') | Out-Null }
    else { Write-Host ("[WARN] " + $M) }
}
function Show-ErrorMessage {
    param([string]$M)
    if ($script:IsGuiMode) { [System.Windows.Forms.MessageBox]::Show($M, "錯誤", 'OK', 'Error') | Out-Null }
    else { Write-Host ("[ERROR] " + $M) }
}

# ============================================================
# Safe property / value helpers (Pitfall 3: StrictMode)
# ============================================================
function Get-ObjectPropertyValueSafe {
    param([object]$Object, [string]$PropertyName, [object]$DefaultValue = "")
    if ($null -eq $Object) { return $DefaultValue }
    try {
        if ($Object -is [hashtable]) {
            if ($Object.ContainsKey($PropertyName) -and $null -ne $Object[$PropertyName]) { return $Object[$PropertyName] }
            return $DefaultValue
        }
        $prop = $Object.PSObject.Properties[$PropertyName]
        if ($null -ne $prop -and $null -ne $prop.Value) { return $prop.Value }
    } catch { return $DefaultValue }
    return $DefaultValue
}

function Get-FieldDouble {
    # 取欄位並轉為 double; 轉不動時回傳預設值並可記為資料缺漏
    param([object]$Object, [string]$PropertyName, [double]$DefaultValue = 0.0)
    $raw = Get-ObjectPropertyValueSafe -Object $Object -PropertyName $PropertyName -DefaultValue ""
    $s = ([string]$raw).Trim()
    if ([string]::IsNullOrEmpty($s)) { return $DefaultValue }
    $out = 0.0
    $style = [System.Globalization.NumberStyles]::Float
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if ([double]::TryParse($s, $style, $culture, [ref]$out)) { return $out }
    return $DefaultValue
}

function Get-FieldInt {
    param([object]$Object, [string]$PropertyName, [int]$DefaultValue = 0)
    return [int][math]::Round((Get-FieldDouble -Object $Object -PropertyName $PropertyName -DefaultValue ([double]$DefaultValue)))
}

function Get-FieldString {
    param([object]$Object, [string]$PropertyName, [string]$DefaultValue = "")
    return ([string](Get-ObjectPropertyValueSafe -Object $Object -PropertyName $PropertyName -DefaultValue $DefaultValue)).Trim()
}

function Test-FieldMissing {
    param([object]$Object, [string]$PropertyName)
    $raw = Get-ObjectPropertyValueSafe -Object $Object -PropertyName $PropertyName -DefaultValue ""
    return ([string]::IsNullOrEmpty(([string]$raw).Trim()))
}

function Test-TruthyFlag {
    param([object]$Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return [bool]$Value }
    $s = ([string]$Value).Trim().ToLower()
    return ($s -eq "true" -or $s -eq "1" -or $s -eq "yes" -or $s -eq "y")
}

function ConvertTo-SafeFileName {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "empty" }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $r = $Text
    foreach ($c in $invalid) { $r = $r.Replace([string]$c, "_") }
    return $r
}

# ============================================================
# UTF-8 file I/O (no BOM)
# ============================================================
function Read-AllTextUtf8 {
    param([string]$Path)
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}
function Write-AllTextUtf8 {
    param([string]$Path, [string]$Text)
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
}

# ============================================================
# CSV helpers
# ============================================================
function ConvertTo-CsvValue {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    $s = [string]$Value
    $s = $s.Replace('"', '""')
    if ($s.Contains(",") -or $s.Contains("`r") -or $s.Contains("`n") -or $s.Contains('"')) {
        return '"' + $s + '"'
    }
    return $s
}
function Join-CsvLine {
    param([object[]]$Values)
    $parts = @()
    foreach ($v in $Values) { $parts += (ConvertTo-CsvValue $v) }
    return ($parts -join ",")
}

function Import-CsvSafe {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    try { return @(Import-Csv -LiteralPath $Path -Encoding UTF8) }
    catch {
        Write-ErrorLog ("Import-CsvSafe failed: {0} : {1}" -f $Path, $_.Exception.Message)
        return @()
    }
}

# ============================================================
# Atomic file ops (Pitfall 6: 共享資料夾需原子寫檔)
# ============================================================
function Invoke-WithFileLock {
    param([string]$LockName, [scriptblock]$ScriptBlock, [int]$TimeoutSeconds = 20)
    $safe = ConvertTo-SafeFileName $LockName
    $lockPath = Join-Path $script:LockRoot ($safe + ".lock")
    $start = Get-Date
    $stream = $null
    while ($true) {
        try {
            $stream = New-Object System.IO.FileStream(
                $lockPath,
                [System.IO.FileMode]::OpenOrCreate,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None
            )
            break
        } catch {
            Start-Sleep -Milliseconds 300
            if (((Get-Date) - $start).TotalSeconds -gt $TimeoutSeconds) {
                throw ("檔案鎖定逾時: " + $LockName)
            }
        }
    }
    try { & $ScriptBlock }
    finally {
        if ($null -ne $stream) { $stream.Close(); $stream.Dispose() }
        try { Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue } catch { }
    }
}

function Backup-FileBeforeWrite {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        $name = Split-Path -Leaf $Path
        $safe = ConvertTo-SafeFileName $name
        $bk = "{0}_{1}.bak" -f (Get-Date -Format "yyyyMMdd_HHmmss"), $safe
        Copy-Item -LiteralPath $Path -Destination (Join-Path $script:BackupBeforeWriteRoot $bk) -Force
    }
}

function Write-FileSafe {
    param([string]$Path, [string]$Text)
    Invoke-WithFileLock -LockName (Split-Path -Leaf $Path) -ScriptBlock {
        try {
            Backup-FileBeforeWrite -Path $Path
            $tmp = $Path + ".tmp_" + ([System.Guid]::NewGuid().ToString("N"))
            Write-AllTextUtf8 -Path $tmp -Text $Text
            Move-Item -LiteralPath $tmp -Destination $Path -Force
        } catch {
            $pname = "{0}_{1}.pending.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"), ([System.Guid]::NewGuid().ToString("N").Substring(0, 8))
            Write-AllTextUtf8 -Path (Join-Path $script:PendingRoot $pname) -Text $Text
            Write-ErrorLog ("Write-FileSafe failed: " + $_.Exception.Message)
            throw
        }
    }
}

function Add-LinesSafe {
    param([string]$Path, [string[]]$Lines)
    Invoke-WithFileLock -LockName (Split-Path -Leaf $Path) -ScriptBlock {
        try {
            Backup-FileBeforeWrite -Path $Path
            $tmp = $Path + ".tmp_" + ([System.Guid]::NewGuid().ToString("N"))
            $existing = ""
            if (Test-Path -LiteralPath $Path) {
                Copy-Item -LiteralPath $Path -Destination $tmp -Force
                $existing = [System.IO.File]::ReadAllText($tmp, [System.Text.Encoding]::UTF8)
            } else {
                Write-AllTextUtf8 -Path $tmp -Text ""
            }
            # 既有內容若無結尾換行, 先補上, 避免新行黏在最後一行之後
            if ($existing.Length -gt 0 -and -not $existing.EndsWith("`n")) {
                [System.IO.File]::AppendAllText($tmp, [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
            }
            # v1.0.6: 批次一次附加 (原版逐行 Add-Content 每行開關檔案, 數萬行極慢)
            $sbApp = New-Object System.Text.StringBuilder
            foreach ($line in $Lines) { [void]$sbApp.Append($line); [void]$sbApp.Append([Environment]::NewLine) }
            [System.IO.File]::AppendAllText($tmp, $sbApp.ToString(), (New-Object System.Text.UTF8Encoding($false)))
            Move-Item -LiteralPath $tmp -Destination $Path -Force
        } catch {
            $pname = "{0}_{1}.pending.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"), ([System.Guid]::NewGuid().ToString("N").Substring(0, 8))
            $Lines | Set-Content -LiteralPath (Join-Path $script:PendingRoot $pname) -Encoding UTF8
            Write-ErrorLog ("Add-LinesSafe failed: " + $_.Exception.Message)
            throw
        }
    }
}

# ============================================================
# 閉包安全的路徑取得函式 (Pitfall 1)
# WinForms 事件閉包 (GetNewClosure) 內讀 $script: 變數會解析為空,
# 但「函式呼叫」在閉包內正常 (函式於 script scope 執行)。
# 事件處理器內一律用這些函式取路徑, 禁止直接讀 $script:XxxPath。
# ============================================================
function Get-HumanReviewCsvPath { return $script:HumanReviewPath }
function Get-ReportRootPath     { return $script:ReportRoot }
function Get-ImportRootPath     { return $script:ImportRoot }
function Get-RawImportRootPath  { return $script:RawImportRoot }
function Get-GoldenRootPath     { return $script:GoldenRoot }
function Get-LogRootPath        { return $script:LogRoot }
function Get-DeidentMapPath     { return $script:DeidentMapPath }
function Get-AIVersion          { return $script:AIVersion }

# ============================================================
# Config (含規則門檻, 可由治理小組調整; SOP 14.1)
# ============================================================
function Get-DefaultConfig {
    return @{
        AIVersion  = $script:AIVersion
        DataSource = "CSV"          # CSV | SQL (SQL 為預留介面)
        Sql = @{
            Enabled          = $false
            ConnectionString = ""   # 建議使用 Integrated Security, 禁止明碼帳密
            Queries = @{
                RunSummary       = ""
                ToolStability    = ""
                Maintenance      = ""
                ProductDrift     = ""
                MeasurementTrend = ""
                HistoryCases     = ""
            }
        }
        Llm = @{
            Enabled = $false        # Layer 3 LLM 介面預留, 預設停用 (資安: 不外送)
            Note    = "啟用前須通過資安審核, 且僅允許送出去識別化相對值資料"
        }
        # Log 比對排除參數 (v1.0.13): 萬用字元樣式, 分號/逗號分隔。
        # 不是所有 Device 參數都需比對 — 依現場指示預設排除 Dop1 (近零噪音
        # 掩蓋其他警訊); 可於「編輯設定」加入如 Hyd1.*;*.dp_SP;*.dp_MV
        LogCompareExcludeParams = "Dop1.*"
        Thresholds = @{
            StabilityScoreMedium   = 85     # 任一穩定分數低於此 -> Medium
            StabilityScoreHigh     = 70     # 任一穩定分數低於此 -> High
            AlarmCountMedium       = 5      # 前一 Run alarm 數 >= 此 -> Medium
            VacuumRecoveryWarnMin  = 20     # 真空恢復時間(分) > 此 -> Medium
            PMTargetDays           = 21     # PM 目標週期(天)
            PMWarnRatio            = 0.8    # DaysAfterPM > 週期*此比率 -> Medium
            SameAlarmHighCount     = 3      # 同一 alarm 7 天內 >= 此 -> High
            PostPmWatchRuns        = 3      # PM 後前 N Run 屬觀察期
            PLConsecutiveRuns      = 3      # PL 連續同方向偏移 Run 數 -> Medium
            ThicknessWarnPct       = 2.0    # |厚度偏差%| > 此 -> Medium
            ThicknessHighPct       = 3.0    # |厚度偏差%| > 此 -> High
            RsWarnPct              = 2.0    # |Rs 偏差%| > 此 -> Medium
            TrendRunCount          = 10     # 趨勢圖顯示最近 N Run
            SimilarCaseMax         = 5      # 相似案例最多顯示筆數
            LogDeltaWarnPct        = 3.0    # N-1 vs N-2 Log StepCode 平均 |差異%| > 此 -> Medium
            LogDeltaHighPct        = 8.0    # N-1 vs N-2 Log StepCode 平均 |差異%| > 此 -> High
            GoldenDeltaWarnPct     = 5.0    # N-1 vs Golden |差異%| > 此 -> Medium
            GoldenDeltaHighPct     = 10.0   # N-1 vs Golden |差異%| > 此 -> High
            LogCompareMinSec       = 5      # Step 秒數低於此不列入比對 (避免短 step 噪音)
            LogCompareTopN         = 10     # 報告列出差異最大的前 N 筆
            LogCompareMinBaseAbs   = 0.0    # Log 比對: 基準平均絕對值低於此者略過 (0=不啟用; 濾近零參數噪音)
            MeasUnifWorsenPct      = 20.0   # 量測總檔: 均勻性 STD 增幅% > 此 -> 趨勢 Worse
            MeasDefectUpPct        = 30.0   # 量測總檔: 缺陷 (LPD3+4/Area_total) 平均增幅% > 此 -> 趨勢 Up
            PLShiftMinNm           = 0.5    # 量測總檔: PL 波長偏移絕對值 >= 此 (nm) 才計入連續偏移
        }
    }
}

function ConvertTo-HashtableFromPsObject {
    # ConvertFrom-Json 回傳 PSCustomObject, 轉為巢狀 hashtable 方便安全存取
    param([object]$Object)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Object.Keys) { $h[[string]$k] = ConvertTo-HashtableFromPsObject $Object[$k] }
        return $h
    }
    if ($Object -is [System.Array]) {
        $list = @()
        foreach ($i in $Object) { $list += ,(ConvertTo-HashtableFromPsObject $i) }
        return $list
    }
    if ($Object -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $Object.PSObject.Properties) { $h[$p.Name] = ConvertTo-HashtableFromPsObject $p.Value }
        return $h
    }
    return $Object
}

function Merge-ConfigDefaults {
    # 舊 config 缺欄位時以預設值回填 (Pitfall 3: schema backfill)
    param([hashtable]$Config, [hashtable]$Defaults)
    foreach ($k in $Defaults.Keys) {
        if (-not $Config.ContainsKey($k) -or $null -eq $Config[$k]) {
            $Config[$k] = $Defaults[$k]
        } elseif ($Defaults[$k] -is [hashtable] -and $Config[$k] -is [hashtable]) {
            Merge-ConfigDefaults -Config $Config[$k] -Defaults $Defaults[$k] | Out-Null
        }
    }
    return $Config
}

function Initialize-Config {
    $defaults = Get-DefaultConfig
    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        $json = (New-Object PSObject -Property $defaults) | ConvertTo-Json -Depth 10
        Write-FileSafe -Path $script:ConfigPath -Text $json
        $script:AppState.Config = $defaults
        Write-AppLog "config.json 不存在, 已建立預設設定。"
        return
    }
    try {
        $raw = Read-AllTextUtf8 -Path $script:ConfigPath
        $obj = $raw | ConvertFrom-Json
        $cfg = ConvertTo-HashtableFromPsObject $obj
        if ($null -eq $cfg) { $cfg = @{} }
        $script:AppState.Config = Merge-ConfigDefaults -Config $cfg -Defaults $defaults
    } catch {
        Write-ErrorLog ("讀取 config.json 失敗, 改用預設值: " + $_.Exception.Message)
        $script:AppState.Config = $defaults
    }
}

function Save-AppConfig {
    # 將目前 AppState.Config 寫回 config.json (原子寫檔 + 寫前備份 + 稽核留痕)
    $cfg = $script:AppState.Config
    if ($null -eq $cfg) { throw "設定尚未載入。" }
    $json = ConvertTo-Json -InputObject $cfg -Depth 10
    Write-FileSafe -Path $script:ConfigPath -Text $json
    Write-AuditLog -Action "CONFIG_SAVED" -Detail "config.json updated"
}

function Get-LogCompareExcludePatterns {
    # Log 比對排除參數樣式清單 (config.LogCompareExcludeParams; 分號/逗號分隔)
    $cfg = $script:AppState.Config
    $s = [string](Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "LogCompareExcludeParams" -DefaultValue "")
    $pats = @()
    foreach ($p in $s.Split([char[]]@(';', ','))) {
        $t = $p.Trim()
        if ($t.Length -gt 0) { $pats += $t }
    }
    return $pats
}

function Get-Threshold {
    param([string]$Name, [double]$DefaultValue = 0.0)
    $cfg = $script:AppState.Config
    if ($null -eq $cfg) { return $DefaultValue }
    $th = Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "Thresholds" -DefaultValue $null
    if ($null -eq $th) { return $DefaultValue }
    $v = Get-ObjectPropertyValueSafe -Object $th -PropertyName $Name -DefaultValue $DefaultValue
    try { return [double]$v } catch { return $DefaultValue }
}

# ============================================================
# Users (登入/角色; 覆判留痕僅記 RoleCode, 符合去識別化)
# ============================================================
function Initialize-UsersCsv {
    if (-not (Test-Path -LiteralPath $script:UsersPath)) {
        $lines = @()
        $lines += "UserId,DisplayName,RoleCode,Role,IsActive"
        $lines += "admin,系統管理員,ADM_01,Admin,1"
        $lines += "eng01,製程工程師1,ENG_01,Engineer,1"
        $lines += "eng02,設備工程師1,EQP_01,Engineer,1"
        $lines += "view01,檢視者1,VIEW_01,Viewer,1"
        Write-FileSafe -Path $script:UsersPath -Text ($lines -join [Environment]::NewLine)
        Write-AppLog "users.csv 不存在, 已建立預設帳號 (admin/eng01/eng02/view01)。"
    }
}

function Find-UserById {
    param([string]$UserId)
    if ([string]::IsNullOrEmpty($UserId)) { return $null }
    $rows = @(Import-CsvSafe -Path $script:UsersPath)
    foreach ($u in $rows) {
        $uid = Get-FieldString -Object $u -PropertyName "UserId"
        $act = Get-FieldString -Object $u -PropertyName "IsActive" -DefaultValue "0"
        if ($uid.ToUpper() -eq $UserId.Trim().ToUpper() -and $act -eq "1") { return $u }
    }
    return $null
}

function Get-AllUsers {
    return @(Import-CsvSafe -Path $script:UsersPath)
}

function Add-UserAccount {
    <#
      新增使用者 (僅 Admin 於 GUI 呼叫)。
      RoleCode 依角色自動編號 (ADM_xx / ENG_xx / VIEW_xx), 留痕僅記 RoleCode。
      回傳新 RoleCode。
    #>
    param([string]$UserId, [string]$DisplayName, [string]$Role)
    $UserId = ([string]$UserId).Trim()
    $DisplayName = ([string]$DisplayName).Trim()
    if ([string]::IsNullOrEmpty($UserId)) { throw "工號不可為空。" }
    if ($UserId -notmatch '^[A-Za-z0-9_\-\.]+$') { throw "工號僅允許英數字與 _ - . 符號。" }
    $validRoles = @("Admin", "Engineer", "Viewer")
    if ($validRoles -notcontains $Role) { throw ("角色必須為: " + ($validRoles -join " / ")) }
    if ([string]::IsNullOrEmpty($DisplayName)) { $DisplayName = $UserId }

    Initialize-UsersCsv
    $rows = @(Get-AllUsers)
    foreach ($u in $rows) {
        if ((Get-FieldString -Object $u -PropertyName "UserId").ToUpper() -eq $UserId.ToUpper()) {
            throw ("工號 {0} 已存在 (含停用帳號)。" -f $UserId)
        }
    }
    $prefixMap = @{ Admin = "ADM"; Engineer = "ENG"; Viewer = "VIEW" }
    $prefix = [string]$prefixMap[$Role]
    $maxN = 0
    foreach ($u in $rows) {
        $rc = Get-FieldString -Object $u -PropertyName "RoleCode"
        if ($rc -match ('^' + $prefix + '_(\d+)$')) {
            $n = [int]$matches[1]
            if ($n -gt $maxN) { $maxN = $n }
        }
    }
    $roleCode = "{0}_{1:D2}" -f $prefix, ($maxN + 1)
    $line = Join-CsvLine @($UserId, $DisplayName, $roleCode, $Role, "1")
    Add-LinesSafe -Path $script:UsersPath -Lines @($line)
    Write-AuditLog -Action "USER_ADDED" -Detail ("{0} role={1}" -f $roleCode, $Role)
    return $roleCode
}

function Set-UserActive {
    # 停用 / 啟用帳號; 保護最後一位可用 Admin 不可停用
    param([string]$UserId, [bool]$Active)
    $rows = @(Get-AllUsers)
    if ($rows.Count -eq 0) { throw "users.csv 無資料。" }
    $found = $false
    foreach ($u in $rows) {
        if ((Get-FieldString -Object $u -PropertyName "UserId").ToUpper() -eq $UserId.Trim().ToUpper()) { $found = $true }
    }
    if (-not $found) { throw ("找不到工號: " + $UserId) }

    if (-not $Active) {
        $activeAdmins = 0
        $targetIsAdmin = $false
        foreach ($u in $rows) {
            $isActive = ((Get-FieldString -Object $u -PropertyName "IsActive" -DefaultValue "0") -eq "1")
            $isAdmin = ((Get-FieldString -Object $u -PropertyName "Role") -eq "Admin")
            $isTarget = ((Get-FieldString -Object $u -PropertyName "UserId").ToUpper() -eq $UserId.Trim().ToUpper())
            if ($isActive -and $isAdmin) { $activeAdmins++ }
            if ($isTarget -and $isAdmin) { $targetIsAdmin = $true }
        }
        if ($targetIsAdmin -and $activeAdmins -le 1) { throw "不可停用最後一位可用的 Admin 帳號。" }
    }

    $lines = @("UserId,DisplayName,RoleCode,Role,IsActive")
    foreach ($u in $rows) {
        $uid = Get-FieldString -Object $u -PropertyName "UserId"
        $act = Get-FieldString -Object $u -PropertyName "IsActive" -DefaultValue "0"
        if ($uid.ToUpper() -eq $UserId.Trim().ToUpper()) {
            if ($Active) { $act = "1" } else { $act = "0" }
        }
        $lines += (Join-CsvLine @(
            $uid,
            (Get-FieldString -Object $u -PropertyName "DisplayName"),
            (Get-FieldString -Object $u -PropertyName "RoleCode"),
            (Get-FieldString -Object $u -PropertyName "Role"),
            $act))
    }
    Write-FileSafe -Path $script:UsersPath -Text ($lines -join [Environment]::NewLine)
    Write-AuditLog -Action "USER_ACTIVE_CHANGED" -Detail ("{0} active={1}" -f $UserId, $Active)
}

# ============================================================
# 去識別化 (SOP 7.2): RawImport -> Import 轉換 + 對照表
# ============================================================
$script:DeidentColumnRules = @(
    @{ RawColumn = "RunID";     AliasColumn = "RunID_Alias";           Prefix = "RUN"  },
    @{ RawColumn = "ToolID";    AliasColumn = "ToolAlias";             Prefix = "TOOL" },
    @{ RawColumn = "ChamberID"; AliasColumn = "ChamberAlias";          Prefix = "CH"   },
    @{ RawColumn = "LotID";     AliasColumn = "LotAlias";              Prefix = "LOT"  },
    @{ RawColumn = "ProductID"; AliasColumn = "ProductFamily";         Prefix = "PF"   },
    @{ RawColumn = "RecipeName";AliasColumn = "RecipeFamily";          Prefix = "RF"   },
    @{ RawColumn = "PrevProductID"; AliasColumn = "PreviousProductFamily"; Prefix = "PF" },
    @{ RawColumn = "Customer";  AliasColumn = "CustomerGroup";         Prefix = "CG"   },
    @{ RawColumn = "Engineer";  AliasColumn = "RoleCode";              Prefix = "ROLE" }
)

function Import-DeidentMap {
    param([string]$MapPath)
    $map = @{}
    foreach ($r in (Import-CsvSafe -Path $MapPath)) {
        $cat = Get-FieldString -Object $r -PropertyName "Category"
        $raw = Get-FieldString -Object $r -PropertyName "RawValue"
        $ali = Get-FieldString -Object $r -PropertyName "Alias"
        if (-not [string]::IsNullOrEmpty($cat) -and -not [string]::IsNullOrEmpty($raw)) {
            $map[($cat + "|" + $raw)] = $ali
        }
    }
    return $map
}

function Get-DeidentAlias {
    <#
      取得或新建 alias; State.NewRows 為 hashtable 回寫槽 (Pitfall 7.5)。
      Category 一律用 Prefix (如 PF / TOOL), 使 ProductID 與 PrevProductID
      等不同欄位共用同一別名空間: 同一原始值必得同一 alias, 不同值必不撞名。
      v1.0.6 效能: 類別既有數量改用 State.Counters 快取 (首次使用該類別時
      掃描一次, 之後 O(1) 遞增)。原版每新增一個 alias 即全表掃描, 大檔
      (數萬個新 RunID/LotID) 為 O(n^2), 是轉換耗時數十分鐘的主因。
      編號結果與原版完全相同。
    #>
    param([hashtable]$Map, [string]$Category, [string]$RawValue, [string]$Prefix, [hashtable]$State)
    $key = $Category + "|" + $RawValue
    if ($Map.ContainsKey($key)) { return [string]$Map[$key] }

    if (-not $State.ContainsKey("Counters")) { $State["Counters"] = @{} }
    $counters = $State["Counters"]
    if (-not $counters.ContainsKey($Category)) {
        $cnt = 0
        $catPrefix = $Category + "|"
        foreach ($k in $Map.Keys) { if ($k.StartsWith($catPrefix)) { $cnt++ } }
        $counters[$Category] = $cnt
    }
    $count = [int]$counters[$Category] + 1
    $counters[$Category] = $count

    if (-not $State.ContainsKey("Stamp")) {
        $State["Stamp"] = @{
            Date = (Get-Date -Format "yyyyMMdd")
            Time = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
    }
    $alias = "{0}_{1:D3}" -f $Prefix, $count
    if ($Category -eq "RUN") {
        $alias = "RUN_{0}_{1:D3}" -f ([string]$State["Stamp"].Date), $count
    }
    $Map[$key] = $alias
    $newRow = Join-CsvLine @($Category, $RawValue, $alias, [string]$State["Stamp"].Time)
    if ($State.NewRows -is [System.Collections.Generic.List[string]]) { [void]$State.NewRows.Add($newRow) }
    else { $State.NewRows += ,$newRow }
    return $alias
}

function Convert-RawImportFile {
    <#
      將 RawImport 內含真實 ID 的 CSV 轉為去識別化 CSV 寫入 Import。
      - 依 $script:DeidentColumnRules 將欄位改名並以 alias 取代值
      - 其餘欄位原樣保留 (應僅含相對值/分數, 不得含絕對機密參數)
      - 對照表 DeidentMap.csv 僅存本機受控資料夾
    #>
    param([string]$RawPath, [string]$OutPath, [string]$MapPath)
    # v1.0.9: 必以 @() 包裝 — 單列 CSV 回傳會被解開為單一物件,
    # PS 5.1 StrictMode 下 .Count 會報「找不到屬性 'Count'」
    $rows = @(Import-CsvSafe -Path $RawPath)
    if ($rows.Count -eq 0) { throw ("RawImport 檔案無資料: " + $RawPath) }
    $map = Import-DeidentMap -MapPath $MapPath
    $mapState = @{ NewRows = (New-Object System.Collections.Generic.List[string]); Counters = @{} }

    $first = $rows[0]
    $rawCols = @()
    foreach ($p in $first.PSObject.Properties) { $rawCols += $p.Name }

    # 以欄位「位置」對應規則 (Import-Csv 每列屬性順序與標頭一致)
    $outCols = @()
    $ruleByIdx = @{}
    for ($i = 0; $i -lt $rawCols.Count; $i++) {
        $matched = $null
        foreach ($rule in $script:DeidentColumnRules) {
            if ($rule.RawColumn -eq $rawCols[$i]) { $matched = $rule; break }
        }
        if ($null -ne $matched) {
            $ruleByIdx[$i] = $matched
            $outCols += $matched.AliasColumn
        } else {
            $outCols += $rawCols[$i]
        }
    }

    # v1.0.6: StringBuilder + 行內 CSV 跳脫。原版以陣列 += 累積輸出行 (O(n^2)
    # 複製) 且每儲存格經 2~3 層函式呼叫, 大檔需數十分鐘; 改寫後為數秒。
    $nl = [Environment]::NewLine
    $sbOut = New-Object System.Text.StringBuilder
    [void]$sbOut.Append((Join-CsvLine $outCols))
    foreach ($r in $rows) {
        [void]$sbOut.Append($nl)
        $i = 0
        foreach ($p in $r.PSObject.Properties) {
            if ($i -gt 0) { [void]$sbOut.Append(',') }
            $v = [string]$p.Value
            if ($v.Length -gt 0) { $v = $v.Trim() }
            if ($v.Length -gt 0 -and $ruleByIdx.ContainsKey($i)) {
                $rule = $ruleByIdx[$i]
                # 已有 alias 直接查表 (免函式呼叫); 僅新值才進 Get-DeidentAlias
                $mk = ([string]$rule.Prefix) + "|" + $v
                if ($map.ContainsKey($mk)) { $v = [string]$map[$mk] }
                else { $v = Get-DeidentAlias -Map $map -Category $rule.Prefix -RawValue $v -Prefix $rule.Prefix -State $mapState }
            }
            # 行內跳脫, 規則與 ConvertTo-CsvValue 相同 (避免逐儲存格函式呼叫)
            if ($v.IndexOf(',') -ge 0 -or $v.IndexOf('"') -ge 0 -or $v.IndexOf("`r") -ge 0 -or $v.IndexOf("`n") -ge 0) {
                $v = '"' + $v.Replace('"', '""') + '"'
            }
            [void]$sbOut.Append($v)
            $i++
        }
    }

    # v1.0.6 順序修正: 先落地對照表、再寫去識別化輸出檔 —
    # 原版順序相反, 中途中斷會產生「有輸出檔、無對照表」而使 alias 無法回溯;
    # 且對照表檔一律建立 (即使本次無新 alias), 便於確認去識別化已受控留存。
    if (-not (Test-Path -LiteralPath $MapPath)) {
        Write-FileSafe -Path $MapPath -Text "Category,RawValue,Alias,CreatedAt"
    }
    if ($mapState.NewRows.Count -gt 0) {
        Add-LinesSafe -Path $MapPath -Lines $mapState.NewRows
    }
    Write-FileSafe -Path $OutPath -Text $sbOut.ToString()
    return $rows.Count
}

function Convert-AllRawImports {
    <#
      將 Data\RawImport 下的 CSV 轉換到 Data\Import (同檔名)。
      v1.0.9: 機台秒級 Log 檔 (MAT 檔名慣例; Log 比對直接讀取) 與
      量測總檔 (依格式定義已為去識別後資料) 不含去識別化規則欄位,
      自動略過; 單一檔案失敗不中斷其餘檔案, 結果逐檔列出。
    #>
    $files = @(Get-ChildItem -LiteralPath $script:RawImportRoot -Filter "*.csv" -ErrorAction SilentlyContinue)
    if ($files.Count -eq 0) { return @() }
    $results = @()
    foreach ($f in $files) {
        if ($null -ne (Get-RunLogFileInfo -FileName $f.Name)) {
            $results += ("{0} -> 機台 Log 檔, 免轉換 (Log 比對直接讀取)" -f $f.Name)
            continue
        }
        if (Test-MeasurementMasterFile -Path $f.FullName) {
            $results += ("{0} -> 量測總檔 (已為去識別後資料), 免轉換" -f $f.Name)
            continue
        }
        try {
            $out = Join-Path $script:ImportRoot $f.Name
            $n = Convert-RawImportFile -RawPath $f.FullName -OutPath $out -MapPath $script:DeidentMapPath
            $results += ("{0} -> {1} 筆已去識別化" -f $f.Name, $n)
            Write-AuditLog -Action "DEIDENT_CONVERT" -Detail ("{0} rows={1}" -f $f.Name, $n)
        } catch {
            $results += ("{0} -> 轉換失敗: {1}" -f $f.Name, $_.Exception.Message)
            Write-ErrorLog ("Convert-AllRawImports: {0} : {1}" -f $f.Name, $_.Exception.Message)
        }
    }
    return $results
}

# ============================================================
# 機台 Log 檔名目錄 (v1.0.7)
# 檔名慣例: <產品>.xxx..._MAT<機台2碼><Run碼>_<流水號>.csv
#   例: H01B9N.PRODUCT.B9N.P02069_M06_P(NGR)_MAT06261175_17155.csv
#       -> 機台 MAT06 / Run 261175 / 產品 H01B9N
# 供診斷頁下拉選單帶入與 Log 對 Log 自動比對 (RunSummary 未建檔時)。
# ============================================================
function Get-RunLogFileInfo {
    # 解析單一 Log 檔名; 不符合慣例回傳 $null (資料表 CSV 自然被排除)
    param([string]$FileName)
    if ([string]::IsNullOrEmpty($FileName)) { return $null }
    $m = [regex]::Match($FileName, '(?i)MAT(\d{2})(\d{4,})(?:_\d+)?\.csv$')
    if (-not $m.Success) { return $null }
    $prod = ""
    $dot = $FileName.IndexOf('.')
    if ($dot -gt 0) { $prod = $FileName.Substring(0, $dot) }
    # 檔名第一段即含 MAT 碼者視為無產品前綴
    if ($prod -match '(?i)MAT\d') { $prod = "" }
    return @{
        Tool    = ("MAT" + $m.Groups[1].Value)
        RunId   = $m.Groups[2].Value
        Product = $prod
    }
}

function Get-RunLogCatalog {
    <#
      掃描 Import 與 RawImport 資料夾內符合檔名慣例的機台秒級 Log 檔,
      回傳陣列: @{ Tool; RunId; Product; Path; FileName }
      同一 機台|Run 以先找到者為準 (Import 優先於 RawImport)。
    #>
    $entries = @{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($root in @((Get-ImportRootPath), (Get-RawImportRootPath))) {
        $files = @(Get-ChildItem -LiteralPath $root -Filter "*.csv" -ErrorAction SilentlyContinue)
        foreach ($f in $files) {
            $info = Get-RunLogFileInfo -FileName $f.Name
            if ($null -eq $info) { continue }
            $key = ([string]$info.Tool) + "|" + ([string]$info.RunId)
            if ($entries.ContainsKey($key)) { continue }
            $info["Path"] = $f.FullName
            $info["FileName"] = $f.Name
            $entries[$key] = $info
            [void]$order.Add($key)
        }
    }
    $list = @()
    foreach ($k in $order) { $list += $entries[$k] }
    return $list
}

# ============================================================
# 資料表範本 (v1.0.9)
# 於 Data\Import 建立六張資料表的空白範本 (僅標頭)。空表與無檔案的
# 診斷行為完全相同 (行為中性), 目的是讓工程師知道要填哪些欄位。
# 已存在的檔案一律不動。
# ============================================================
# 六張資料表標頭定義 (v1.0.15: 集中定義, 範本建立與資料表管理介面共用)
$script:ImportTableHeaders = @{
    "RunSummary.csv"       = "RunID_Alias,ToolAlias,ChamberAlias,ProductFamily,RecipeFamily,RecipeVersionGroup,PreviousProductFamily,RunStartTime,RunEndTime,RunResult,OperatorShift,RecipeVersionChanged,RecipeChangeNote"
    "ToolStability.csv"    = "RunID_Alias,ToolAlias,TempStabilityScore,PressureStabilityScore,MFCStabilityScore,RotationStabilityScore,VacuumRecoveryTimeMin,AlarmCountLastRun,CriticalAlarmCount,InterlockEventCount,CriticalAlarmWithin24h,RecentAlarmCodes"
    "Maintenance.csv"      = "RunID_Alias,ToolAlias,DaysAfterPM,RunsAfterPM,DaysAfterPartChange,RecentSameAlarmCount7d,MTBFTrend,RecentCorrectiveMaintenance,GoldenRunVerified"
    "ProductDrift.csv"     = "RunID_Alias,ToolAlias,ProductFamily,PLPeakShiftNm,ConsecutivePLShiftRuns,PLIntensityTrend,XRDPeakShiftDeg,ThicknessDeltaPct,UniformityTrend,RsDeltaPct,AOIDefectTrend,OverSpcWarning,OverSpcControl"
    "MeasurementTrend.csv" = "RunID_Alias,ToolAlias,RunDate,AlarmCount,PLPeakShiftNm,ThicknessDeltaPct"
    "HistoryCases.csv"     = "CaseID,ToolAlias,ProductFamily,RiskCategory,Keywords,Summary,Action,Outcome,CaseDate"
}
function Get-ImportTableHeaders { return $script:ImportTableHeaders }

function Initialize-ImportTableTemplates {
    $templates = Get-ImportTableHeaders
    $created = @()
    foreach ($name in $templates.Keys) {
        $p = Join-Path $script:ImportRoot $name
        if (-not (Test-Path -LiteralPath $p)) {
            Write-AllTextUtf8 -Path $p -Text ([string]$templates[$name])
            $created += $name
        }
    }
    $doc = Join-Path $script:DataRoot "資料表欄位說明.txt"
    if (-not (Test-Path -LiteralPath $doc)) {
        $nl = [Environment]::NewLine
        $help = @(
            "資料表欄位說明 (Data\Import\*.csv; 以 RunID_Alias 為鍵, 一列一 Run)",
            "RunID_Alias 接受「Run 碼」(如 261176) 或「機台+Run 碼」(如 MAT06261176 /",
            "MAT06_261176 / MAT06-261176), 不分大小寫。",
            "所有 ID 應為去識別化代號 (可用設定頁 RawImport -> Import 轉換工具產生)。",
            "機台秒級 Log 檔與量測總檔直接放入 Import / RawImport 即可, 不必建表。",
            "",
            "[自動 vs 手動]",
            "  自動補列: RunSummary (由 Log 檔名) 與 MeasurementTrend (由量測總檔)",
            "            於啟動與「重載資料」時自動增列缺少的 Run; 人工維護列不會被覆蓋。",
            "  自動彙總: ProductDrift 可不填 — 量測總檔存在時系統自動計算飄移。",
            "  需手動 / MES 匯出: ToolStability (alarm 與穩定度分數)、Maintenance (PM / 換件)、",
            "            HistoryCases (歷史案例) — 這些資料不在 Log 或量測檔內, 系統無從產生;",
            "            未填時該面向以「保守中風險 + 資料不足」呈現。",
            "",
            "[RunSummary.csv] Run 基本資料 (機台/Run 下拉選單來源之一)",
            "  RunID_Alias=Run 代號  ToolAlias=機台代號  ChamberAlias=腔體  ProductFamily=產品族",
            "  RecipeFamily=Recipe 族  RecipeVersionGroup=Recipe 版本群  PreviousProductFamily=前一 Run 產品族",
            "  RunStartTime/RunEndTime=起訖 (yyyy-MM-dd HH:mm:ss)  RunResult=Normal|Abnormal|Abort|Hold|Planned",
            "  OperatorShift=班別  RecipeVersionChanged=0|1  RecipeChangeNote=變更說明 (含 ECN)",
            "",
            "[ToolStability.csv] 前一 Run 機台穩定度 (以該 Run 的 RunID_Alias 記錄)",
            "  TempStabilityScore/PressureStabilityScore/MFCStabilityScore/RotationStabilityScore=0~100 分",
            "  VacuumRecoveryTimeMin=真空恢復分鐘  AlarmCountLastRun=Alarm 總數  CriticalAlarmCount=Critical Alarm 數",
            "  InterlockEventCount=Interlock 次數  CriticalAlarmWithin24h=0|1  RecentAlarmCodes=代碼(分號分隔)",
            "",
            "[Maintenance.csv] 維修 / PM 特徵",
            "  DaysAfterPM=距上次 PM 天數  RunsAfterPM=PM 後第幾 Run  DaysAfterPartChange=距換件天數",
            "  RecentSameAlarmCount7d=同一 Alarm 7 天內次數  MTBFTrend=Up|Flat|Down",
            "  RecentCorrectiveMaintenance=Yes|No  GoldenRunVerified=Yes|No|NA",
            "",
            "[ProductDrift.csv] 產品特性飄移 (量測總檔存在時可不填, 系統自動彙總)",
            "  PLPeakShiftNm=PL 峰值偏移 nm  ConsecutivePLShiftRuns=連續同方向偏移 Run 數",
            "  PLIntensityTrend=Up|Flat|Down  XRDPeakShiftDeg=XRD 偏移度  ThicknessDeltaPct=厚度偏差%",
            "  UniformityTrend=Better|Flat|Worse  RsDeltaPct=Rs 偏差%  AOIDefectTrend=Up|Flat|Down",
            "  OverSpcWarning/OverSpcControl=0|1 (超出 SPC warning/control limit)",
            "",
            "[MeasurementTrend.csv] 趨勢圖資料 (診斷頁右側圖表)",
            "  RunDate=yyyy-MM-dd  AlarmCount=Alarm 數  PLPeakShiftNm / ThicknessDeltaPct 同上",
            "",
            "[HistoryCases.csv] 歷史案例 (相似案例推薦來源)",
            "  CaseID=案例代號  RiskCategory=Stability|Maintenance|Drift  Keywords=關鍵字(分號分隔)",
            "  Summary=摘要  Action=處置  Outcome=結果  CaseDate=yyyy-MM-dd"
        ) -join $nl
        Write-AllTextUtf8 -Path $doc -Text $help
    }
    if ($created.Count -gt 0) {
        Write-AppLog ("已建立資料表空白範本: " + ($created -join ", "))
    }
    return $created
}

# ============================================================
# 量測總檔 (Measurement Master, v1.0.8)
# 標頭: STRUCTURE,REACTOR,RUN_NO,POS_NO,PF,FAIL_NO,...,LEHI_RS,...
#   STRUCTURE=產品名 / REACTOR=EQPID (如 Tool06) / RUN_NO=RUNID /
#   POS_NO=擺放位置 / PF=最終判定 Pass|Fail / FAIL_NO=Fail 代碼
#   量測欄: ODD1,ODD2,Area_Cnt,Area_total,Haze avg (表面缺陷),
#           LEHI_RS,LEHI_UNIF_STD (Rs 與均勻性), LPD3+4 (表面缺陷),
#           AGA_WL_AVG,IGA_WL_* (PL, 非所有產品皆量測)
# 檔案可含表頭說明前言 (# 註解與欄位片段行), 解析時自動定位真實標頭。
# 機台對應: REACTOR 尾碼數字 == Log 檔名 MAT 尾碼數字 (Tool06 <-> MAT06)。
# ============================================================
$script:MeasurementMetricCols = @(
    "LEHI_RS", "LEHI_UNIF_STD", "Haze avg", "LPD3+4",
    "Area_Cnt", "Area_total", "ODD1", "ODD2", "R2R_W2W",
    "AGA_WL_AVG", "IGA_WL_AVG", "IGA_WL_MAX", "IGA_WL_STD", "IGA_INT", "IGA_FWHM")

function Get-ToolNumber {
    # 取機台識別尾碼數字 (MAT06 / Tool06 -> 6); 無數字回傳 -1
    param([string]$Name)
    if (-not [string]::IsNullOrEmpty($Name) -and $Name -match '(\d+)\s*$') { return [int]$matches[1] }
    return -1
}

function Import-MeasurementMasterRows {
    <#
      解析量測總檔為 hashtable 列陣列 (key = 標頭欄名, 值已 Trim)。
      真實標頭 = 以 STRUCTURE,REACTOR 開頭且欄位含 RUN_NO 與 PF 的行
      (前言中的欄位片段行如 "STRUCTURE,REACTOR,RUN_NO,POS_NO," 無 PF, 不會誤判);
      標頭之後的空白行與 # 註解行略過; 短列缺欄視為空值。
    #>
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    $hdrIdx = -1
    $header = $null
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $ln = $lines[$i].Trim()
        if ($ln -notmatch '^(?i)STRUCTURE\s*,\s*REACTOR') { continue }
        $names = @()
        foreach ($x in (Split-CsvLineFields -Line $ln)) { $names += $x.Trim() }
        if ($names -contains "RUN_NO" -and $names -contains "PF") { $hdrIdx = $i; $header = $names; break }
    }
    if ($hdrIdx -lt 0) { return @() }

    $list = New-Object System.Collections.Generic.List[object]
    for ($i = $hdrIdx + 1; $i -lt $lines.Length; $i++) {
        $ln = $lines[$i]
        if ([string]::IsNullOrEmpty($ln) -or $ln.Trim().Length -eq 0) { continue }
        if ($ln.TrimStart().StartsWith("#")) { continue }
        $f = Split-CsvLineFields -Line $ln
        $h = @{}
        $n = [math]::Min($header.Count, $f.Length)
        for ($j = 0; $j -lt $n; $j++) { $h[$header[$j]] = $f[$j].Trim() }
        [void]$list.Add($h)
    }
    # 注意: 這裡不可用 @($list) — List[object] 經 @() 轉換會觸發
    # PSToObjectArrayBinder 的 "Argument types do not match" 錯誤
    return $list.ToArray()
}

function Test-MeasurementMasterFile {
    # 以標頭快速判斷檔案是否為量測總檔 (前 100 行內出現真實標頭)
    param([string]$Path)
    try {
        $sr = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
        try {
            for ($i = 0; $i -lt 100; $i++) {
                $ln = $sr.ReadLine()
                if ($null -eq $ln) { break }
                if ($ln -match '^(?i)STRUCTURE\s*,\s*REACTOR' -and $ln -match 'RUN_NO' -and $ln -match '(^|,)\s*PF\s*(,|$)') {
                    return $true
                }
            }
        } finally { $sr.Dispose() }
    } catch {
        Write-ErrorLog ("Test-MeasurementMasterFile: {0} : {1}" -f $Path, $_.Exception.Message)
    }
    return $false
}

function Find-MeasurementMasterFiles {
    # 掃描 Import / RawImport 內的 *.csv / *.txt, 以標頭快速判斷是否為量測總檔
    $found = @()
    foreach ($root in @((Get-ImportRootPath), (Get-RawImportRootPath))) {
        foreach ($pat in @("*.csv", "*.txt")) {
            $files = @(Get-ChildItem -LiteralPath $root -Filter $pat -ErrorAction SilentlyContinue)
            foreach ($f in $files) {
                if (Test-MeasurementMasterFile -Path $f.FullName) { $found += $f.FullName }
            }
        }
    }
    return $found
}

function Get-MeasurementRunCatalog {
    <#
      量測總檔 -> 每 Run 彙總 (每列 = 一片 wafer, 依 機台尾碼+RUN_NO 分組):
      @{ ToolNum; ToolName; RunId; Structure; GDate; WaferCount;
         PassCount; FailCount; FailCodes; Mean(各量測欄, 無值為 $null); SourceFile }
    #>
    $groups = @{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($path in @(Find-MeasurementMasterFiles)) {
        foreach ($r in @(Import-MeasurementMasterRows -Path $path)) {
            $reactor = Get-FieldString -Object $r -PropertyName "REACTOR"
            $runNo   = Get-FieldString -Object $r -PropertyName "RUN_NO"
            if ([string]::IsNullOrEmpty($reactor) -or $runNo -notmatch '^\d+$') { continue }
            $tn = Get-ToolNumber -Name $reactor
            if ($tn -lt 0) { continue }
            $key = [string]$tn + "|" + $runNo
            if (-not $groups.ContainsKey($key)) {
                $g = @{
                    ToolNum = $tn; ToolName = $reactor; RunId = $runNo
                    Structure = (Get-FieldString -Object $r -PropertyName "STRUCTURE")
                    GDate     = (Get-FieldString -Object $r -PropertyName "G_DATE")
                    WaferCount = 0; PassCount = 0; FailCount = 0; FailCodes = @()
                    Sum = @{}; Cnt = @{}; Mean = @{}
                    SourceFile = (Split-Path -Leaf $path)
                }
                foreach ($c in $script:MeasurementMetricCols) { $g.Sum[$c] = 0.0; $g.Cnt[$c] = 0 }
                $groups[$key] = $g
                [void]$order.Add($key)
            }
            $g = $groups[$key]
            $g.WaferCount = $g.WaferCount + 1
            $pf = (Get-FieldString -Object $r -PropertyName "PF").ToUpper()
            if ($pf -eq "PASS") { $g.PassCount = $g.PassCount + 1 }
            elseif ($pf -eq "FAIL") {
                $g.FailCount = $g.FailCount + 1
                $fc = Get-FieldString -Object $r -PropertyName "FAIL_NO"
                if (-not [string]::IsNullOrEmpty($fc) -and (@($g.FailCodes) -notcontains $fc)) { $g.FailCodes += $fc }
            }
            foreach ($c in $script:MeasurementMetricCols) {
                $v = ConvertTo-DoubleInvariant -Text (Get-FieldString -Object $r -PropertyName $c)
                if ($null -ne $v) {
                    $g.Sum[$c] = $g.Sum[$c] + $v
                    $g.Cnt[$c] = $g.Cnt[$c] + 1
                }
            }
        }
    }
    $list = @()
    foreach ($k in $order) {
        $g = $groups[$k]
        foreach ($c in $script:MeasurementMetricCols) {
            if ($g.Cnt[$c] -gt 0) { $g.Mean[$c] = [double]($g.Sum[$c] / $g.Cnt[$c]) }
            else { $g.Mean[$c] = $null }
        }
        $list += $g
    }
    return $list
}

function Get-MeasurementDriftInfo {
    <#
      依機台 (MAT06 / Tool06 皆取尾碼數字對應) 找 RunId <= UpToRunId 的
      最近量測 Run (符合 N-2 資料時效), 與再前一量測 Run 比較, 產出:
      - DriftRow: 餵給 Get-DriftRisk 的欄位 (RsDeltaPct / PLPeakShiftNm /
        ConsecutivePLShiftRuns / UniformityTrend / AOIDefectTrend)
      - ValueSummary: 主要量測值摘要字串 (本 Run vs 前 Run)
      回傳 $null (無資料) 或 @{ Current; Previous; DriftRow; ValueSummary }
      僅做整理 / 比對, 不判定放行 (PF 為量測系統之最終判定, 僅轉述)。
    #>
    param([string]$ToolAlias, [long]$UpToRunId)
    $tn = Get-ToolNumber -Name $ToolAlias
    if ($tn -lt 0) { return $null }
    $mine = @()
    foreach ($m in @(Get-MeasurementRunCatalog)) {
        if ([int]$m.ToolNum -ne $tn) { continue }
        if ([long]$m.RunId -gt $UpToRunId) { continue }
        $mine += $m
    }
    if ($mine.Count -eq 0) { return $null }
    $mine = @($mine | Sort-Object -Property @{ Expression = { [long]$_.RunId }; Descending = $true })
    $cur = $mine[0]
    $prev = $null
    if ($mine.Count -ge 2) { $prev = $mine[1] }

    $unifWorsen = Get-Threshold -Name "MeasUnifWorsenPct" -DefaultValue 20.0
    $defUp      = Get-Threshold -Name "MeasDefectUpPct"   -DefaultValue 30.0
    $plMin      = Get-Threshold -Name "PLShiftMinNm"      -DefaultValue 0.5

    $row = @{ RunID_Alias = [string]$cur.RunId }
    if ($null -ne $prev) {
        $rc = $cur.Mean["LEHI_RS"]; $rp = $prev.Mean["LEHI_RS"]
        if ($null -ne $rc -and $null -ne $rp -and [math]::Abs([double]$rp) -gt 1e-9) {
            $row["RsDeltaPct"] = [math]::Round((([double]$rc - [double]$rp) / [math]::Abs([double]$rp)) * 100.0, 2)
        }
        $uc = $cur.Mean["LEHI_UNIF_STD"]; $up = $prev.Mean["LEHI_UNIF_STD"]
        if ($null -ne $uc -and $null -ne $up -and [double]$up -gt 1e-9) {
            if (((([double]$uc - [double]$up) / [double]$up) * 100.0) -gt $unifWorsen) { $row["UniformityTrend"] = "Worse" }
        }
        $defCol = ""
        foreach ($c in @("LPD3+4", "Area_total")) {
            if ($null -ne $cur.Mean[$c] -and $null -ne $prev.Mean[$c]) { $defCol = $c; break }
        }
        if (-not [string]::IsNullOrEmpty($defCol)) {
            $dc = [double]$cur.Mean[$defCol]; $dp = [double]$prev.Mean[$defCol]
            if ($dp -gt 1e-9 -and ((($dc - $dp) / $dp) * 100.0) -gt $defUp) { $row["AOIDefectTrend"] = "Up" }
        }
        # PL 波長偏移與連續同方向偏移 Run 數 (非所有產品皆量測 PL)
        $plCol = ""
        foreach ($c in @("IGA_WL_AVG", "AGA_WL_AVG")) {
            if ($null -ne $cur.Mean[$c] -and $null -ne $prev.Mean[$c]) { $plCol = $c; break }
        }
        if (-not [string]::IsNullOrEmpty($plCol)) {
            $shift = [double]$cur.Mean[$plCol] - [double]$prev.Mean[$plCol]
            $row["PLPeakShiftNm"] = [math]::Round($shift, 3)
            $consec = 0
            if ([math]::Abs($shift) -ge $plMin) {
                $sign = [math]::Sign($shift)
                $consec = 1
                for ($i = 1; $i -lt ($mine.Count - 1); $i++) {
                    $a = $mine[$i].Mean[$plCol]; $b = $mine[$i + 1].Mean[$plCol]
                    if ($null -eq $a -or $null -eq $b) { break }
                    $d = [double]$a - [double]$b
                    if ([math]::Abs($d) -lt $plMin -or [math]::Sign($d) -ne $sign) { break }
                    $consec++
                }
            }
            $row["ConsecutivePLShiftRuns"] = $consec
        }
    }

    # 主要量測值摘要 (本 Run vs 前一量測 Run)
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $vals = @()
    foreach ($def in @(
        @("LEHI_RS", "Rs", "0.###"),
        @("LEHI_UNIF_STD", "均勻性STD", "0.###"),
        @("Haze avg", "Haze", "0.####"),
        @("LPD3+4", "LPD3+4", "0.#"),
        @("IGA_WL_AVG", "PL波長", "0.##"))) {
        $c = [string]$def[0]
        $cv = "-"
        if ($null -ne $cur.Mean[$c]) { $cv = ([double]$cur.Mean[$c]).ToString([string]$def[2], $inv) }
        $pv = "-"
        if ($null -ne $prev -and $null -ne $prev.Mean[$c]) { $pv = ([double]$prev.Mean[$c]).ToString([string]$def[2], $inv) }
        if ($cv -ne "-" -or $pv -ne "-") { $vals += ("{0} {1} (前 {2})" -f [string]$def[1], $cv, $pv) }
    }

    return @{
        Current = $cur; Previous = $prev
        DriftRow = $row
        ValueSummary = ($vals -join "; ")
    }
}

# ============================================================
# 資料表自動補列 (v1.0.10)
# RunSummary / MeasurementTrend 的內容可由 Log 檔名目錄與量測總檔推得,
# 自動補列缺少的 RunID_Alias (只增列, 人工維護的既有列一律不動)。
# ToolStability / Maintenance / HistoryCases 之 alarm / PM / 案例資料
# 不存在於 Log 或量測檔, 需由 MES / 設備系統匯出或人工維護。
# ============================================================
function Update-DerivedTables {
    Initialize-ImportTableTemplates | Out-Null
    $inv = [System.Globalization.CultureInfo]::InvariantCulture

    # --- RunSummary: 由 Log 檔名目錄補列 ---
    $rsPath = Join-Path $script:ImportRoot "RunSummary.csv"
    $existing = @{}
    foreach ($r in @(Import-CsvSafe -Path $rsPath)) {
        $k = Get-FieldString -Object $r -PropertyName "RunID_Alias"
        if (-not [string]::IsNullOrEmpty($k)) { $existing[$k] = $true }
    }
    $cat = @(Get-RunLogCatalog)
    $cat = @($cat | Sort-Object -Property @{ Expression = { [string]$_.Tool } }, @{ Expression = { [long]$_.RunId } })
    $rsLines = New-Object System.Collections.Generic.List[string]
    $prevProdByTool = @{}
    $rsAdded = 0
    foreach ($e in $cat) {
        $tool = [string]$e.Tool
        $prevProd = ""
        if ($prevProdByTool.ContainsKey($tool)) { $prevProd = [string]$prevProdByTool[$tool] }
        $prevProdByTool[$tool] = [string]$e.Product
        if ($existing.ContainsKey([string]$e.RunId)) { continue }
        $ts = ""
        try { $ts = ([System.IO.File]::GetLastWriteTime([string]$e.Path)).ToString("yyyy-MM-dd HH:mm:ss") } catch { }
        # 欄序: RunID_Alias,ToolAlias,ChamberAlias,ProductFamily,RecipeFamily,RecipeVersionGroup,
        #       PreviousProductFamily,RunStartTime,RunEndTime,RunResult,OperatorShift,
        #       RecipeVersionChanged,RecipeChangeNote
        # RunStartTime 以 Log 檔案時間近似 (僅供排序); RunResult 未知留空 (不臆測)
        $rsLines.Add((Join-CsvLine @([string]$e.RunId, $tool, "", [string]$e.Product, "", "", $prevProd, $ts, "", "", "", "0", "")))
        $existing[[string]$e.RunId] = $true
        $rsAdded++
    }
    if ($rsLines.Count -gt 0) {
        Add-LinesSafe -Path $rsPath -Lines $rsLines
        Write-AppLog ("RunSummary 自動補列 {0} 筆 (由 Log 檔名目錄)。" -f $rsAdded)
    }

    # --- MeasurementTrend: 由量測總檔補列 (PL 偏移 = 對前一量測 Run) ---
    $mtPath = Join-Path $script:ImportRoot "MeasurementTrend.csv"
    $mtExisting = @{}
    foreach ($r in @(Import-CsvSafe -Path $mtPath)) {
        $k = Get-FieldString -Object $r -PropertyName "RunID_Alias"
        if (-not [string]::IsNullOrEmpty($k)) { $mtExisting[$k] = $true }
    }
    $meas = @(Get-MeasurementRunCatalog)
    $meas = @($meas | Sort-Object -Property @{ Expression = { [int]$_.ToolNum } }, @{ Expression = { [long]$_.RunId } })
    $mtLines = New-Object System.Collections.Generic.List[string]
    $prevPlByTool = @{}
    $mtAdded = 0
    foreach ($m in $meas) {
        $toolAlias = "MAT{0:D2}" -f [int]$m.ToolNum
        $pl = $m.Mean["IGA_WL_AVG"]
        if ($null -eq $pl) { $pl = $m.Mean["AGA_WL_AVG"] }
        $shiftTxt = ""
        if ($prevPlByTool.ContainsKey($toolAlias) -and $null -ne $pl -and $null -ne $prevPlByTool[$toolAlias]) {
            $shiftTxt = ([double]$pl - [double]$prevPlByTool[$toolAlias]).ToString("0.###", $inv)
        }
        $prevPlByTool[$toolAlias] = $pl
        if ($mtExisting.ContainsKey([string]$m.RunId)) { continue }
        # 欄序: RunID_Alias,ToolAlias,RunDate,AlarmCount,PLPeakShiftNm,ThicknessDeltaPct
        $mtLines.Add((Join-CsvLine @([string]$m.RunId, $toolAlias, [string]$m.GDate, "", $shiftTxt, "")))
        $mtExisting[[string]$m.RunId] = $true
        $mtAdded++
    }
    if ($mtLines.Count -gt 0) {
        Add-LinesSafe -Path $mtPath -Lines $mtLines
        Write-AppLog ("MeasurementTrend 自動補列 {0} 筆 (由量測總檔)。" -f $mtAdded)
    }

    return @{ RunSummaryAdded = $rsAdded; MeasurementTrendAdded = $mtAdded }
}

# ============================================================
# StepCode 平均值 (以秒計 Run Log -> 每個 StepCode 各參數平均)
# ============================================================
function ConvertTo-DoubleInvariant {
    <# 以 InvariantCulture 解析數字; 失敗回傳 $null (避免地區設定小數點差異) #>
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    $v = 0.0
    $ok = [double]::TryParse($Text.Trim(),
            [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v)
    if ($ok) { return $v }
    return $null
}

# ------------------------------------------------------------
# 快速彙總引擎 (v1.0.5 效能修正)
# 原版以 Import-Csv + 逐儲存格呼叫 PowerShell 函式計算平均,
# 在 PS 5.1 每格約 50~100 微秒; 秒級 Log (數萬列 x 數十~數百欄)
# 為數百萬格, 單檔即需數分鐘, 前 Run 診斷 (3 檔) 可達 20 分鐘以上。
# 改為內嵌 C# 串流彙總 (Add-Type 離線編譯, 免安裝):
#   - 逐字元 CSV 解析 (支援引號欄位 / 引號內逗號與換行, RFC4180)
#   - 單次掃描累加 Sum/Cnt, 不建立每列物件
# 編譯失敗時退回純 PowerShell 串流模式 (仍遠快於原版)。
# ------------------------------------------------------------
$script:StepMeanAggregatorSource = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace RunPreCheckAI
{
    public class StepGroupAgg
    {
        public string Key;
        public string Stepcode;
        public int N;
        public double[] Sum;
        public int[] Cnt;
    }

    public class StepMeanAggResult
    {
        public string GroupColumn;
        public string[] NumCols;
        public int RowCount;
        public List<string> Order;
        public Dictionary<string, StepGroupAgg> Groups;

        public StepMeanAggResult()
        {
            Order = new List<string>();
            Groups = new Dictionary<string, StepGroupAgg>(StringComparer.Ordinal);
        }
    }

    public static class StepMeanAggregator
    {
        // RFC4180 CSV 紀錄讀取: 支援引號欄位 (含逗號 / 雙引號 / 換行);
        // 讀到串流結尾回傳 null。sb / fields 由呼叫端重複使用以減少配置。
        private static List<string> ReadRecord(TextReader reader, StringBuilder sb, List<string> fields)
        {
            int c = reader.Read();
            if (c < 0) { return null; }
            fields.Clear();
            sb.Length = 0;
            bool inQuotes = false;
            while (true)
            {
                if (c < 0) { fields.Add(sb.ToString()); return fields; }
                char ch = (char)c;
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (reader.Peek() == '"') { sb.Append('"'); reader.Read(); }
                        else { inQuotes = false; }
                    }
                    else { sb.Append(ch); }
                }
                else
                {
                    if (ch == '"' && sb.Length == 0) { inQuotes = true; }
                    else if (ch == ',') { fields.Add(sb.ToString()); sb.Length = 0; }
                    else if (ch == '\n') { fields.Add(sb.ToString()); return fields; }
                    else if (ch == '\r')
                    {
                        if (reader.Peek() == '\n') { reader.Read(); }
                        fields.Add(sb.ToString());
                        return fields;
                    }
                    else { sb.Append(ch); }
                }
                c = reader.Read();
            }
        }

        private static bool IsBlankRecord(List<string> rec)
        {
            return (rec.Count == 1 && rec[0].Length == 0);
        }

        // 錯誤代碼 (由 PowerShell 端轉為原版中文訊息):
        //   NO_DATA / NO_GROUP_COLUMN / NO_NUMERIC_COLUMNS
        public static StepMeanAggResult Aggregate(string path, string[] groupCandidates, string[] metaCols)
        {
            StepMeanAggResult result = new StepMeanAggResult();
            StringBuilder sb = new StringBuilder(256);
            List<string> buf = new List<string>(64);
            CultureInfo inv = CultureInfo.InvariantCulture;

            using (StreamReader sr = new StreamReader(path, Encoding.UTF8, true))
            {
                List<string> header = ReadRecord(sr, sb, buf);
                while (header != null && IsBlankRecord(header)) { header = ReadRecord(sr, sb, buf); }
                if (header == null) { throw new InvalidDataException("NO_DATA"); }
                string[] cols = header.ToArray();

                // 先確認至少有一筆資料列 (與原版錯誤順序一致: 無資料優先於缺欄位)
                List<string> rec = ReadRecord(sr, sb, buf);
                while (rec != null && IsBlankRecord(rec)) { rec = ReadRecord(sr, sb, buf); }
                if (rec == null) { throw new InvalidDataException("NO_DATA"); }

                int groupIdx = -1;
                string groupCol = null;
                for (int gi = 0; gi < groupCandidates.Length && groupIdx < 0; gi++)
                {
                    for (int i = 0; i < cols.Length; i++)
                    {
                        if (string.Equals(cols[i], groupCandidates[gi], StringComparison.Ordinal))
                        {
                            groupIdx = i;
                            groupCol = cols[i];
                            break;
                        }
                    }
                }
                if (groupIdx < 0) { throw new InvalidDataException("NO_GROUP_COLUMN"); }

                int stepcodeIdx = -1;
                for (int i = 0; i < cols.Length; i++)
                {
                    if (string.Equals(cols[i], "Stepcode_s", StringComparison.Ordinal)) { stepcodeIdx = i; break; }
                }

                List<int> numIdxList = new List<int>();
                List<string> numNames = new List<string>();
                for (int i = 0; i < cols.Length; i++)
                {
                    bool isMeta = false;
                    for (int m = 0; m < metaCols.Length; m++)
                    {
                        if (string.Equals(cols[i], metaCols[m], StringComparison.Ordinal)) { isMeta = true; break; }
                    }
                    if (!isMeta)
                    {
                        numIdxList.Add(i);
                        numNames.Add(cols[i]);
                    }
                }
                if (numIdxList.Count == 0) { throw new InvalidDataException("NO_NUMERIC_COLUMNS"); }

                result.GroupColumn = groupCol;
                result.NumCols = numNames.ToArray();
                int[] numIdx = numIdxList.ToArray();
                int nNum = numIdx.Length;
                int rowCount = 0;

                while (rec != null)
                {
                    if (!IsBlankRecord(rec))
                    {
                        rowCount++;
                        string key = "";
                        if (groupIdx < rec.Count) { key = rec[groupIdx]; }
                        if (key.Length == 0) { key = "(blank)"; }

                        StepGroupAgg g;
                        if (!result.Groups.TryGetValue(key, out g))
                        {
                            g = new StepGroupAgg();
                            g.Key = key;
                            g.Stepcode = "";
                            g.N = 0;
                            g.Sum = new double[nNum];
                            g.Cnt = new int[nNum];
                            if (stepcodeIdx >= 0)
                            {
                                string sc = "";
                                if (stepcodeIdx < rec.Count) { sc = rec[stepcodeIdx]; }
                                if (sc.Length == 0) { sc = "(blank)"; }
                                g.Stepcode = sc;
                            }
                            result.Groups.Add(key, g);
                            result.Order.Add(key);
                        }
                        g.N++;

                        for (int j = 0; j < nNum; j++)
                        {
                            int idx = numIdx[j];
                            if (idx >= rec.Count) { continue; }
                            string s = rec[idx];
                            if (s.Length == 0) { continue; }
                            double v;
                            if (double.TryParse(s.Trim(), NumberStyles.Float, inv, out v))
                            {
                                g.Sum[j] += v;
                                g.Cnt[j]++;
                            }
                        }
                    }
                    rec = ReadRecord(sr, sb, buf);
                }
                result.RowCount = rowCount;
            }
            return result;
        }
    }
}
'@

$script:FastLogAggChecked = $false
$script:FastLogAggReady   = $false

function Initialize-FastLogAggregator {
    # 首次使用時編譯 C# 彙總引擎 (約 1 秒, 之後常駐); 失敗僅記 log, 自動退回純 PS 模式
    if ($script:FastLogAggChecked) { return }
    $script:FastLogAggChecked = $true
    try {
        if ($null -eq ("RunPreCheckAI.StepMeanAggregator" -as [type])) {
            Add-Type -TypeDefinition $script:StepMeanAggregatorSource -ErrorAction Stop
        }
        $script:FastLogAggReady = $true
        Write-AppLog "StepMean 快速彙總引擎 (C#) 已載入。"
    } catch {
        $script:FastLogAggReady = $false
        Write-AppLog ("StepMean 快速彙總引擎編譯失敗, 改用純 PowerShell 串流模式: " + $_.Exception.Message)
    }
}

function Split-CsvLineFields {
    <#
      單行 CSV 欄位切割 (純 PS 退回模式用)。
      無引號時走 String.Split 快速路徑; 有引號時逐字元解析 (支援 "" 跳脫)。
      不支援引號內換行 (機台秒級 Log 不會出現; C# 主路徑有完整支援)。
    #>
    param([string]$Line)
    if ($Line.IndexOf('"') -lt 0) { return ,($Line.Split(',')) }
    $fields = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $inQ = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($inQ) {
            if ($ch -eq '"') {
                if (($i + 1) -lt $Line.Length -and $Line[$i + 1] -eq '"') { [void]$sb.Append('"'); $i++ }
                else { $inQ = $false }
            } else { [void]$sb.Append($ch) }
        } else {
            if ($ch -eq '"' -and $sb.Length -eq 0) { $inQ = $true }
            elseif ($ch -eq ',') { $fields.Add($sb.ToString()); [void]$sb.Clear() }
            else { [void]$sb.Append($ch) }
        }
    }
    $fields.Add($sb.ToString())
    return ,($fields.ToArray())
}

function Get-StepCodeMeanTableFallback {
    <#
      純 PowerShell 串流彙總 (C# 引擎不可用時的退回模式)。
      StreamReader 逐行讀取 + 行內 [double]::TryParse + 索引陣列累加,
      不建立每列 PSCustomObject、不做逐儲存格函式呼叫。
      行為 / 錯誤訊息與快速路徑及原版一致。
    #>
    param([string]$LogPath)

    $inv   = [System.Globalization.CultureInfo]::InvariantCulture
    $style = [System.Globalization.NumberStyles]::Float
    $sr = New-Object System.IO.StreamReader($LogPath, [System.Text.Encoding]::UTF8, $true)
    try {
        $headerLine = $sr.ReadLine()
        while ($null -ne $headerLine -and $headerLine.Length -eq 0) { $headerLine = $sr.ReadLine() }
        if ($null -eq $headerLine) { throw ("Log 檔無資料: " + $LogPath) }
        $cols = Split-CsvLineFields -Line $headerLine

        # 先確認至少有一筆資料列 (與原版錯誤順序一致)
        $firstLine = $sr.ReadLine()
        while ($null -ne $firstLine -and $firstLine.Length -eq 0) { $firstLine = $sr.ReadLine() }
        if ($null -eq $firstLine) { throw ("Log 檔無資料: " + $LogPath) }

        $groupIdx = -1
        $groupCol = ""
        # v1.0.14: StepLabel 優先 (對應真實 Layer)
        foreach ($cand in @("StepLabel", "Step", "Stepcode_s", "StepCode_SP")) {
            for ($i = 0; $i -lt $cols.Length; $i++) {
                if ($cols[$i] -ceq $cand) { $groupIdx = $i; $groupCol = $cols[$i]; break }
            }
            if ($groupIdx -ge 0) { break }
        }
        if ($groupIdx -lt 0) {
            throw "Log 檔缺少分組欄位 (StepLabel / Step / Stepcode_s / StepCode_SP), 無法計算 StepCode 平均。"
        }
        $stepcodeIdx = -1
        for ($i = 0; $i -lt $cols.Length; $i++) {
            if ($cols[$i] -ceq "Stepcode_s") { $stepcodeIdx = $i; break }
        }

        $metaCols = @("Timestamp", "StepLabel", "Step", "Stepcode_s")
        $numIdx = New-Object System.Collections.Generic.List[int]
        $numCols = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $cols.Length; $i++) {
            if ($metaCols -cnotcontains $cols[$i]) { [void]$numIdx.Add($i); [void]$numCols.Add($cols[$i]) }
        }
        if ($numCols.Count -eq 0) { throw "Log 檔無可計算的數值欄位。" }
        $nNum = $numIdx.Count

        $groups = @{}
        $order = New-Object System.Collections.Generic.List[string]
        $rowCount = 0
        $line = $firstLine
        $v = 0.0
        while ($null -ne $line) {
            if ($line.Length -gt 0) {
                $rowCount++
                $f = Split-CsvLineFields -Line $line
                $key = ""
                if ($groupIdx -lt $f.Length) { $key = $f[$groupIdx] }
                if ($key.Length -eq 0) { $key = "(blank)" }
                $g = $null
                if (-not $groups.ContainsKey($key)) {
                    $g = @{ Key = $key; Stepcode = ""; N = 0
                            Sum = (New-Object 'double[]' $nNum); Cnt = (New-Object 'int[]' $nNum) }
                    if ($stepcodeIdx -ge 0) {
                        $sc = ""
                        if ($stepcodeIdx -lt $f.Length) { $sc = $f[$stepcodeIdx] }
                        if ($sc.Length -eq 0) { $sc = "(blank)" }
                        $g.Stepcode = $sc
                    }
                    $groups[$key] = $g
                    [void]$order.Add($key)
                } else {
                    $g = $groups[$key]
                }
                $g.N = $g.N + 1
                $sum = $g.Sum; $cnt = $g.Cnt
                for ($j = 0; $j -lt $nNum; $j++) {
                    $idx = $numIdx[$j]
                    if ($idx -ge $f.Length) { continue }
                    $s = $f[$idx]
                    if ($s.Length -eq 0) { continue }
                    if ([double]::TryParse($s.Trim(), $style, $inv, [ref]$v)) {
                        $sum[$j] += $v
                        $cnt[$j] = $cnt[$j] + 1
                    }
                }
            }
            $line = $sr.ReadLine()
        }
    } finally {
        $sr.Dispose()
    }

    # Sum/Cnt -> Mean (無可解析值時為 $null)
    foreach ($key in $order) {
        $g = $groups[$key]
        $mean = @{}
        for ($j = 0; $j -lt $nNum; $j++) {
            if ($g.Cnt[$j] -gt 0) { $mean[$numCols[$j]] = [double]($g.Sum[$j] / $g.Cnt[$j]) }
            else { $mean[$numCols[$j]] = $null }
        }
        $g.Mean = $mean
        $g.Remove("Sum"); $g.Remove("Cnt")
    }

    return @{
        GroupColumn = $groupCol
        Order       = @($order)
        Groups      = $groups
        NumCols     = @($numCols)
        RowCount    = $rowCount
        SourceName  = (Split-Path -Leaf $LogPath)
    }
}

function Get-StepCodeMeanTable {
    <#
      將以秒計的 Run Log CSV 依 StepCode 分組, 計算每個數值欄位的平均值。
      - 分組鍵優先順序: StepLabel -> Step -> Stepcode_s -> StepCode_SP
        (v1.0.14: StepLabel 才是對應真實磊晶 Layer 的欄位, 同一 Layer 可
        橫跨多個 Step 編號; 無 StepLabel 的舊 Log 依序退回)
      - 分組鍵空白的列歸入 "(blank)" 群組 (常見於 Log 缺 Stepcode 標記)
      - 非數值 / 空白儲存格略過不計; 該群組全數無法解析時平均為 $null
      - 僅做數學平均, 不判定好壞 (AI 權限邊界: 整理/比對, 不放行)
      v1.0.5: 改用 C# 串流彙總引擎 (大檔由數分鐘降至數秒);
              編譯失敗時自動退回純 PowerShell 串流模式, 結果相同。
      回傳 hashtable:
        GroupColumn / Order (key 順序) / NumCols / RowCount / SourceName
        Groups: key -> @{ Key; Stepcode; N; Mean = @{ col -> [double] 或 $null } }
    #>
    param([string]$LogPath)

    if (-not (Test-Path -LiteralPath $LogPath)) { throw ("找不到 Log 檔: " + $LogPath) }

    Initialize-FastLogAggregator
    if (-not $script:FastLogAggReady) {
        return Get-StepCodeMeanTableFallback -LogPath $LogPath
    }

    $agg = $null
    try {
        # v1.0.14: StepLabel 優先 (對應真實 Layer); StepLabel 亦列入 meta 欄不做數值比對
        $agg = [RunPreCheckAI.StepMeanAggregator]::Aggregate(
            $LogPath,
            [string[]]@("StepLabel", "Step", "Stepcode_s", "StepCode_SP"),
            [string[]]@("Timestamp", "StepLabel", "Step", "Stepcode_s"))
    } catch {
        $ex = $_.Exception
        while ($null -ne $ex -and -not ($ex -is [System.IO.InvalidDataException]) -and $null -ne $ex.InnerException) {
            $ex = $ex.InnerException
        }
        if ($ex -is [System.IO.InvalidDataException]) {
            switch ($ex.Message) {
                "NO_DATA"            { throw ("Log 檔無資料: " + $LogPath) }
                "NO_GROUP_COLUMN"    { throw "Log 檔缺少分組欄位 (StepLabel / Step / Stepcode_s / StepCode_SP), 無法計算 StepCode 平均。" }
                "NO_NUMERIC_COLUMNS" { throw "Log 檔無可計算的數值欄位。" }
                default              { throw }
            }
        }
        throw
    }

    # C# 結果 -> 與原版相同形狀的 hashtable (群組數少, 轉換成本可忽略)
    $numCols = @($agg.NumCols)
    $groups = @{}
    foreach ($key in $agg.Order) {
        $ga = $agg.Groups[$key]
        $mean = @{}
        for ($j = 0; $j -lt $numCols.Count; $j++) {
            if ($ga.Cnt[$j] -gt 0) { $mean[$numCols[$j]] = [double]($ga.Sum[$j] / $ga.Cnt[$j]) }
            else { $mean[$numCols[$j]] = $null }
        }
        $groups[$key] = @{ Key = $ga.Key; Stepcode = $ga.Stepcode; N = $ga.N; Mean = $mean }
    }

    return @{
        GroupColumn = $agg.GroupColumn
        Order       = @($agg.Order)
        Groups      = $groups
        NumCols     = $numCols
        RowCount    = $agg.RowCount
        SourceName  = (Split-Path -Leaf $LogPath)
    }
}

function Export-StepCodeMeanReport {
    <#
      StepCode 平均輸出成 CSV (介面與輸出格式與 v1.0.3 相同, 內部改用 Get-StepCodeMeanTable)。
      - 輸出欄位: <分組鍵>, Stepcode_s, N_sec (該 Step 秒數), 各參數平均
      - 以 Write-FileSafe 原子寫檔至 Reports\StepCode_Mean_<原檔名>.csv
      回傳 hashtable: OutPath / GroupCount / RowCount / GroupColumn
    #>
    param([string]$LogPath, [string]$OutPath = "")

    $tbl = Get-StepCodeMeanTable -LogPath $LogPath

    if ([string]::IsNullOrEmpty($OutPath)) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($LogPath)
        $OutPath = Join-Path (Get-ReportRootPath) ("StepCode_Mean_" + (ConvertTo-SafeFileName $base) + ".csv")
    }

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $lines = @()
    $lines += (Join-CsvLine (@($tbl.GroupColumn, "Stepcode_s", "N_sec") + $tbl.NumCols))
    foreach ($key in $tbl.Order) {
        $g = $tbl.Groups[$key]
        $vals = @($g.Key, $g.Stepcode, [string]$g.N)
        foreach ($c in $tbl.NumCols) {
            if ($null -ne $g.Mean[$c]) {
                $vals += ([double]$g.Mean[$c]).ToString("0.####", $inv)
            } else {
                $vals += ""
            }
        }
        $lines += (Join-CsvLine $vals)
    }
    Write-FileSafe -Path $OutPath -Text ($lines -join [Environment]::NewLine)
    Write-AuditLog -Action "STEPMEAN_EXPORT" -Detail ("{0} groups={1} rows={2}" -f (Split-Path -Leaf $LogPath), $tbl.Order.Count, $tbl.RowCount)

    return @{ OutPath = $OutPath; GroupCount = $tbl.Order.Count; RowCount = $tbl.RowCount; GroupColumn = $tbl.GroupColumn }
}

# ============================================================
# 前 Run Log 分層診斷 (v1.0.4; 資料時效限制)
#   Run 前可取得: N-1 (上一 Run) 機台秒級 Log
#   Run 前不可得: N-1 的 PL / XRD / Thickness / Rs / AOI 量測 (最快 N-2)
#   -> N-1 以 Log 對 Log 比對 (vs N-2 Log 與 Golden 基準線) 提前預警
# ============================================================
function Compare-StepCodeMeanTables {
    <#
      比對兩份 StepCode 平均表 (Get-StepCodeMeanTable 輸出)。
      - 以分組鍵對齊 Step; 僅比對兩邊皆有、且秒數皆 >= MinSec 的 Step
      - 相對差異% = (Test - Base) / |Base| * 100; |Base| < 1e-9 之儲存格略過
      - |差異%| > HighPct -> 等級 3; > WarnPct -> 等級 1 (與其他規則同語意)
      - Step 結構不一致 (單邊獨有 Step) -> 至少等級 1 (可能 recipe / 段落不同)
      - 僅做比對與提醒, 不判定好壞與放行
      回傳 hashtable: Level / Reasons / TopDeltas / ExceedCount / ComparedCells /
                      SkippedSteps / OnlyInBase / OnlyInTest / BaseName / TestName
    #>
    param(
        [hashtable]$BaseTable, [hashtable]$TestTable,
        [string]$BaseName, [string]$TestName,
        [double]$WarnPct, [double]$HighPct,
        [int]$MinSec = 5, [int]$TopN = 10,
        [double]$MinBaseAbs = 0.0,   # v1.0.9: 基準平均絕對值低於此者略過 (濾近零參數噪音)
        [string[]]$ExcludeParams = @()   # v1.0.13: 排除參數樣式 (萬用字元; 如 Dop1.*)
    )

    $r = @{
        Level = 0; Reasons = @(); TopDeltas = @(); ExceedCount = 0
        ComparedCells = 0; SkippedSteps = @(); OnlyInBase = @(); OnlyInTest = @()
        BaseName = $BaseName; TestName = $TestName
    }

    # 共同數值欄位 (v1.0.13: 依設定排除指定參數 — 不是所有 Device 參數都需比對)
    $commonCols = @()
    $excludedCols = @()
    foreach ($c in $TestTable.NumCols) {
        if (@($BaseTable.NumCols) -notcontains $c) { continue }
        $skip = $false
        foreach ($pat in $ExcludeParams) {
            if ($c -like $pat) { $skip = $true; break }
        }
        if ($skip) { $excludedCols += $c } else { $commonCols += $c }
    }
    $r["ExcludedParams"] = $excludedCols
    if ($commonCols.Count -eq 0) {
        Add-RiskReason -Result $r -Level 1 -Reason ("{0} 與 {1} 無共同參數欄位, 無法比對 (請確認 Log 欄位格式)。" -f $TestName, $BaseName)
        return $r
    }

    # Step 結構比對
    foreach ($k in $TestTable.Order) { if (-not $BaseTable.Groups.ContainsKey($k)) { $r.OnlyInTest += $k } }
    foreach ($k in $BaseTable.Order) { if (-not $TestTable.Groups.ContainsKey($k)) { $r.OnlyInBase += $k } }
    if ($r.OnlyInTest.Count -gt 0 -or $r.OnlyInBase.Count -gt 0) {
        Add-RiskReason -Result $r -Level 1 -Reason ("Step 結構不一致: {0} 獨有 [{1}]; {2} 獨有 [{3}] (可能 recipe / 段落不同, 需人工確認)。" -f `
            $TestName, ($r.OnlyInTest -join ";"), $BaseName, ($r.OnlyInBase -join ";"))
    }

    # 逐 Step / 逐參數比對 (v1.0.5: List 收集, 避免陣列 += 逐筆複製)
    $details = New-Object System.Collections.Generic.List[object]
    foreach ($k in $TestTable.Order) {
        if (-not $BaseTable.Groups.ContainsKey($k)) { continue }
        $gt = $TestTable.Groups[$k]
        $gb = $BaseTable.Groups[$k]
        if ($gt.N -lt $MinSec -or $gb.N -lt $MinSec) {
            $r.SkippedSteps += $k
            continue
        }
        foreach ($c in $commonCols) {
            $bv = $gb.Mean[$c]
            $tv = $gt.Mean[$c]
            if ($null -eq $bv -or $null -eq $tv) { continue }
            if ([math]::Abs([double]$bv) -lt [math]::Max(1e-9, $MinBaseAbs)) { continue }
            $r.ComparedCells = $r.ComparedCells + 1
            $deltaPct = ([double]$tv - [double]$bv) / [math]::Abs([double]$bv) * 100.0
            $lv = 0
            if ([math]::Abs($deltaPct) -gt $HighPct)     { $lv = 3 }
            elseif ([math]::Abs($deltaPct) -gt $WarnPct) { $lv = 1 }
            if ($lv -gt 0 -or [math]::Abs($deltaPct) -gt 0) {
                $details.Add(@{
                    StepKey = $k; Stepcode = [string]$gt.Stepcode; Param = $c
                    BaseMean = [double]$bv; TestMean = [double]$tv
                    DeltaPct = [double]$deltaPct; Level = $lv
                })
            }
            if ($lv -gt 0) {
                $r.ExceedCount = $r.ExceedCount + 1
                if ($lv -gt $r.Level) { $r.Level = $lv }
            }
        }
    }

    # 依 |差異%| 由大到小取前 N 筆
    $sorted = @($details | Sort-Object -Property @{ Expression = { [math]::Abs([double]$_.DeltaPct) }; Descending = $true })
    if ($sorted.Count -gt $TopN) { $sorted = @($sorted[0..($TopN - 1)]) }
    $r.TopDeltas = $sorted

    # 摘要理由 (最多列 5 筆超標)
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $listed = 0
    foreach ($d in $sorted) {
        if ($d.Level -le 0) { continue }
        if ($listed -ge 5) { break }
        $thName = "警戒"
        $thVal = $WarnPct
        if ($d.Level -ge 3) { $thName = "高風險"; $thVal = $HighPct }
        Add-RiskReason -Result $r -Level $d.Level -Reason ("Step/Layer {0} 參數 {1}: {2} 平均 {3} vs {4} 平均 {5}, 差異 {6}% 超過{7}門檻 {8}%。" -f `
            $d.StepKey, $d.Param, $TestName, $d.TestMean.ToString("0.####", $inv), $BaseName, $d.BaseMean.ToString("0.####", $inv), `
            $d.DeltaPct.ToString("+0.##;-0.##", $inv), $thName, $thVal)
        $listed++
    }
    if ($r.ExceedCount -gt $listed) {
        $r.Reasons += ("(其餘 {0} 筆超標差異詳見報告差異排行)" -f ($r.ExceedCount - $listed))
    }
    if ($r.Reasons.Count -eq 0) {
        $r.Reasons += ("{0} vs {1}: 共比對 {2} 個儲存格, StepCode 平均差異皆在門檻內。" -f $TestName, $BaseName, $r.ComparedCells)
    }
    if ($excludedCols.Count -gt 0) {
        # 留痕: 排除了哪些參數 (依設定 LogCompareExcludeParams)
        $r.Reasons += ("(依設定排除 {0} 個參數不列入比對: {1})" -f $excludedCols.Count, ($excludedCols -join ";"))
    }
    return $r
}

function Invoke-PreRunLogDiagnosis {
    <#
      Run 前「前 Run Log 分層診斷」:
        N-1 (上一 Run) 秒級 Log vs N-2 (上上 Run) 秒級 Log  (LogDelta 門檻)
        N-1 秒級 Log vs Golden Run 基準線 (Data\Golden)      (GoldenDelta 門檻)
      產出文字報告至 Reports\LogPreCheck_*.txt 並留審計紀錄。
      僅做整理 / 比對 / 提醒, 不判定開 Run 與否。
    #>
    param(
        [string]$PrevLogPath,                 # 必填: N-1 機台秒級 Log
        [string]$PrevPrevLogPath = "",        # 選填: N-2 機台秒級 Log
        [string]$GoldenLogPath = "",          # 選填: Golden Run 基準 Log
        [string]$RunAlias = ""                # 選填: 本 Run 代號 (僅用於報告標題)
    )

    if ([string]::IsNullOrEmpty($PrevLogPath)) { throw "請指定上一 Run (N-1) 的秒級 Log 檔。" }
    $swDiag = [System.Diagnostics.Stopwatch]::StartNew()
    $tPrev = Get-StepCodeMeanTable -LogPath $PrevLogPath

    $warn   = Get-Threshold -Name "LogDeltaWarnPct"    -DefaultValue 3.0
    $high   = Get-Threshold -Name "LogDeltaHighPct"    -DefaultValue 8.0
    $gWarn  = Get-Threshold -Name "GoldenDeltaWarnPct" -DefaultValue 5.0
    $gHigh  = Get-Threshold -Name "GoldenDeltaHighPct" -DefaultValue 10.0
    $minSec  = [int](Get-Threshold -Name "LogCompareMinSec" -DefaultValue 5)
    $topN    = [int](Get-Threshold -Name "LogCompareTopN"   -DefaultValue 10)
    $minBase = Get-Threshold -Name "LogCompareMinBaseAbs" -DefaultValue 0.0
    $excPats = @(Get-LogCompareExcludePatterns)

    $cmpPrevPrev = $null
    if (-not [string]::IsNullOrEmpty($PrevPrevLogPath)) {
        $tPP = Get-StepCodeMeanTable -LogPath $PrevPrevLogPath
        $cmpPrevPrev = Compare-StepCodeMeanTables -BaseTable $tPP -TestTable $tPrev `
            -BaseName ("N-2 " + $tPP.SourceName) -TestName ("N-1 " + $tPrev.SourceName) `
            -WarnPct $warn -HighPct $high -MinSec $minSec -TopN $topN -MinBaseAbs $minBase -ExcludeParams $excPats
    }
    $cmpGolden = $null
    if (-not [string]::IsNullOrEmpty($GoldenLogPath)) {
        $tG = Get-StepCodeMeanTable -LogPath $GoldenLogPath
        $cmpGolden = Compare-StepCodeMeanTables -BaseTable $tG -TestTable $tPrev `
            -BaseName ("Golden " + $tG.SourceName) -TestName ("N-1 " + $tPrev.SourceName) `
            -WarnPct $gWarn -HighPct $gHigh -MinSec $minSec -TopN $topN -MinBaseAbs $minBase -ExcludeParams $excPats
    }

    $levels = @(0)
    if ($null -ne $cmpPrevPrev) { $levels += $cmpPrevPrev.Level }
    if ($null -ne $cmpGolden)   { $levels += $cmpGolden.Level }
    $overall = Measure-MaxRisk $levels
    $swDiag.Stop()
    $elapsedSec = [math]::Round($swDiag.Elapsed.TotalSeconds, 1)

    # ---- 報告 ----
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("前 Run Log 分層診斷報告 (Run 前)")
    [void]$sb.AppendLine("產出時間:" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    [void]$sb.AppendLine("AI 版本:" + (Get-AIVersion) + " (規則引擎 + 模板摘要, 未使用外部 LLM)")
    if (-not [string]::IsNullOrEmpty($RunAlias)) { [void]$sb.AppendLine("本 Run:" + $RunAlias) }
    [void]$sb.AppendLine("N-1 Log:" + $tPrev.SourceName + " (分組欄: " + $tPrev.GroupColumn + ", " + $tPrev.Order.Count + " Step / " + $tPrev.RowCount + " 秒)")
    [void]$sb.AppendLine("比對耗時:" + $elapsedSec.ToString($inv) + " 秒")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("零、資料時效說明")
    [void]$sb.AppendLine("Run 前僅能即時取得上一 Run (N-1) 的機台秒級 Log;")
    [void]$sb.AppendLine("PL / XRD / Thickness / Rs / AOI 量測資料最快僅到上上 Run (N-2)。")
    [void]$sb.AppendLine("因此本報告以 N-1 Log 對 N-2 Log / Golden 基準線比對, 提前預警機台行為變化;")
    [void]$sb.AppendLine("N-1 的量測結果確認, 請於量測產出後再行覆核。")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("一、總體風險")
    [void]$sb.AppendLine("風險等級:" + (Get-RiskName $overall))
    [void]$sb.AppendLine("建議狀態:需製程 / 設備工程師覆判後決定後續處置")
    [void]$sb.AppendLine("")

    $section = {
        param([hashtable]$Cmp, [string]$Title, [double]$W, [double]$H)
        [void]$sb.AppendLine($Title)
        if ($null -eq $Cmp) {
            [void]$sb.AppendLine("(未提供基準檔, 未執行此項比對)")
            [void]$sb.AppendLine("")
            return
        }
        [void]$sb.AppendLine("等級:" + (Get-RiskName $Cmp.Level) + ("  (警戒 {0}% / 高風險 {1}%)" -f $W, $H))
        $i = 1
        foreach ($x in $Cmp.Reasons) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
        if (@($Cmp.SkippedSteps).Count -gt 0) {
            [void]$sb.AppendLine(("略過短 Step (秒數 < 門檻): " + (@($Cmp.SkippedSteps) -join ";")))
        }
        if (@($Cmp.TopDeltas).Count -gt 0) {
            [void]$sb.AppendLine("差異排行 (|差異%| 由大到小):")
            foreach ($d in $Cmp.TopDeltas) {
                [void]$sb.AppendLine(("  Step/Layer {0} | {1} | 基準 {2} | N-1 {3} | 差異 {4}% | 等級 {5}" -f `
                    $d.StepKey, $d.Param, $d.BaseMean.ToString("0.####", $inv), $d.TestMean.ToString("0.####", $inv), `
                    $d.DeltaPct.ToString("+0.##;-0.##", $inv), (Get-RiskName $d.Level)))
            }
        }
        [void]$sb.AppendLine("")
    }.GetNewClosure()

    & $section $cmpPrevPrev "二、N-1 vs N-2 秒級 Log 比對" $warn $high
    & $section $cmpGolden   "三、N-1 vs Golden 基準線比對" $gWarn $gHigh

    [void]$sb.AppendLine("四、AI 不可判定事項")
    [void]$sb.AppendLine("AI 不判定是否可開 Run。")
    [void]$sb.AppendLine("AI 不建議修改正式 Recipe。")
    [void]$sb.AppendLine("AI 不判定產品放行或報廢。")
    [void]$sb.AppendLine("AI 不取代 SPC、MES、QMS 或工程師覆判。")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("五、資料來源摘要 (可追溯性)")
    [void]$sb.AppendLine("- N-1 Log: " + $tPrev.SourceName)
    if ($null -ne $cmpPrevPrev) { [void]$sb.AppendLine("- N-2 Log: " + $cmpPrevPrev.BaseName) }
    if ($null -ne $cmpGolden)   { [void]$sb.AppendLine("- Golden : " + $cmpGolden.BaseName) }
    [void]$sb.AppendLine("本報告僅使用檔案內數值之相對比較, 請確認檔名已去識別化再對外提供。")

    $text = $sb.ToString()
    $fname = "LogPreCheck_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    $rpath = Join-Path (Get-ReportRootPath) $fname
    Write-AllTextUtf8 -Path $rpath -Text $text
    Write-AuditLog -Action "LOGPRECHECK" -Detail ("{0} overall={1}" -f $tPrev.SourceName, (Get-RiskName $overall))

    return @{
        Level = $overall; Text = $text; ReportPath = $rpath
        CompareToPrevPrev = $cmpPrevPrev; CompareToGolden = $cmpGolden
        GeneratedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        ElapsedSeconds = $elapsedSec
    }
}

# ============================================================
# 資料存取層 (CSV 為主, SQL 為預留介面)
# ============================================================
function Get-DataRows {
    <#
      統一資料入口。DataSource=CSV 時讀 Data\Import\<TableName>.csv
      DataSource=SQL 且已啟用時, 以 config.Sql.Queries.<TableName> 查詢 (唯讀)。
    #>
    param([string]$TableName)
    $cfg = $script:AppState.Config
    $mode = [string](Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "DataSource" -DefaultValue "CSV")
    if ($mode.ToUpper() -eq "SQL") {
        $sqlCfg = Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "Sql" -DefaultValue $null
        $enabled = Test-TruthyFlag (Get-ObjectPropertyValueSafe -Object $sqlCfg -PropertyName "Enabled" -DefaultValue $false)
        if (-not $enabled) {
            throw "SQL 資料來源尚未啟用 (config.json: Sql.Enabled=false)。請改用 CSV 或完成資安審核後啟用。"
        }
        return Get-SqlDataRows -TableName $TableName -SqlConfig $sqlCfg
    }
    $path = Join-Path $script:ImportRoot ($TableName + ".csv")
    return Import-CsvSafe -Path $path
}

function Get-SqlDataRows {
    # SQL 預留介面: 僅唯讀查詢; 連線字串建議 Integrated Security=SSPI, 禁止明碼帳密
    param([string]$TableName, [object]$SqlConfig)
    $connStr = [string](Get-ObjectPropertyValueSafe -Object $SqlConfig -PropertyName "ConnectionString" -DefaultValue "")
    if ([string]::IsNullOrEmpty($connStr)) { throw "SQL 連線字串未設定。" }
    if ($connStr -match "(?i)password\s*=") {
        Write-AppLog "[SECURITY] 偵測到連線字串包含明碼密碼, 建議改用 Integrated Security。"
    }
    $queries = Get-ObjectPropertyValueSafe -Object $SqlConfig -PropertyName "Queries" -DefaultValue $null
    $query = [string](Get-ObjectPropertyValueSafe -Object $queries -PropertyName $TableName -DefaultValue "")
    if ([string]::IsNullOrEmpty($query)) { throw ("未設定資料表 {0} 的查詢語法。" -f $TableName) }
    if ($query -notmatch "(?i)^\s*select\s") { throw "僅允許 SELECT 唯讀查詢。" }

    Add-Type -AssemblyName System.Data
    $table = New-Object System.Data.DataTable
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $query
        $cmd.CommandTimeout = 30
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        [void]$adapter.Fill($table)
    } finally {
        $conn.Dispose()
    }
    $rows = @()
    foreach ($dr in $table.Rows) {
        $h = New-Object PSObject
        foreach ($col in $table.Columns) {
            $h | Add-Member -MemberType NoteProperty -Name $col.ColumnName -Value ([string]$dr[$col])
        }
        $rows += $h
    }
    return $rows
}

function Find-RowByRun {
    <#
      以 RunID_Alias 查表列。
      v1.0.11: 傳入 ToolAlias 時容錯 — 工程師常以「機台+Run 碼」填寫
      (如 MAT06261176), 與系統的「Run 碼」(261176) 同視為相符;
      接受格式: Run碼 / 機台Run碼 / 機台_Run碼 / 機台-Run碼 (不分大小寫)。
    #>
    param([object[]]$Rows, [string]$RunAlias, [string]$ToolAlias = "")
    if ([string]::IsNullOrEmpty($RunAlias)) { return $null }
    $ra = $RunAlias.Trim().ToUpper()
    $keys = New-Object System.Collections.Generic.List[string]
    [void]$keys.Add($ra)
    if (-not [string]::IsNullOrEmpty($ToolAlias)) {
        $ta = $ToolAlias.Trim().ToUpper()
        [void]$keys.Add($ta + $ra)
        [void]$keys.Add($ta + "_" + $ra)
        [void]$keys.Add($ta + "-" + $ra)
    }
    foreach ($r in $Rows) {
        $rid = (Get-FieldString -Object $r -PropertyName "RunID_Alias").ToUpper()
        if ($rid.Length -gt 0 -and $keys.Contains($rid)) { return $r }
    }
    return $null
}

# ============================================================
# 風險等級工具
# ============================================================
# 等級: 0=低, 1=中, 2=中高, 3=高
$script:RiskNames = @("低", "中", "中高", "高")

function Get-RiskName {
    param([int]$Level)
    if ($Level -lt 0) { $Level = 0 }
    if ($Level -gt 3) { $Level = 3 }
    return $script:RiskNames[$Level]
}

function Measure-MaxRisk {
    param([int[]]$Levels)
    $max = 0
    foreach ($l in $Levels) { if ($l -gt $max) { $max = $l } }
    return $max
}

function Add-RiskReason {
    # Result: hashtable @{ Level; Reasons=@() } ; 依規則升級並記錄原因
    param([hashtable]$Result, [int]$Level, [string]$Reason)
    if ($Level -gt $Result.Level) { $Result.Level = $Level }
    $Result.Reasons += $Reason
}

# ============================================================
# Layer 1 規則引擎 (SOP 8.2 / 10.1 / 10.2 / 10.3)
# ============================================================
function Get-StabilityRisk {
    # 前一 Run 機台穩定度診斷 (SOP 10.1)
    param([object]$StabilityRow, [object]$PrevRunRow, [hashtable]$Missing)
    $r = @{ Level = 0; Reasons = @() }
    if ($null -eq $StabilityRow) {
        $Missing.Items += "缺少前一 Run 機台穩定度資料 (ToolStability)"
        Add-RiskReason -Result $r -Level 1 -Reason "無前一 Run 穩定度資料, 無法確認機台狀態, 保守列為中風險。"
        return $r
    }
    $scoreMed  = Get-Threshold -Name "StabilityScoreMedium" -DefaultValue 85
    $scoreHigh = Get-Threshold -Name "StabilityScoreHigh" -DefaultValue 70
    $alarmMed  = [int](Get-Threshold -Name "AlarmCountMedium" -DefaultValue 5)
    $vacWarn   = Get-Threshold -Name "VacuumRecoveryWarnMin" -DefaultValue 20

    $scoreFields = @(
        @{ Name = "TempStabilityScore";     Label = "溫度穩定分數" },
        @{ Name = "PressureStabilityScore"; Label = "壓力穩定分數" },
        @{ Name = "MFCStabilityScore";      Label = "MFC 穩定分數" },
        @{ Name = "RotationStabilityScore"; Label = "轉速穩定分數" }
    )
    foreach ($f in $scoreFields) {
        if (Test-FieldMissing -Object $StabilityRow -PropertyName $f.Name) {
            $Missing.Items += ("缺少欄位 " + $f.Name)
            continue
        }
        $v = Get-FieldDouble -Object $StabilityRow -PropertyName $f.Name -DefaultValue 100
        if ($v -lt $scoreHigh) {
            Add-RiskReason -Result $r -Level 3 -Reason ("{0} {1} 低於高風險門檻 {2}。" -f $f.Label, $v, $scoreHigh)
        } elseif ($v -lt $scoreMed) {
            Add-RiskReason -Result $r -Level 1 -Reason ("{0} {1} 低於警戒門檻 {2}。" -f $f.Label, $v, $scoreMed)
        }
    }

    $critical = Get-FieldInt -Object $StabilityRow -PropertyName "CriticalAlarmCount" -DefaultValue 0
    $alarms   = Get-FieldInt -Object $StabilityRow -PropertyName "AlarmCountLastRun" -DefaultValue 0
    $interlock = Get-FieldInt -Object $StabilityRow -PropertyName "InterlockEventCount" -DefaultValue 0
    if ($critical -ge 2 -or $interlock -ge 1) {
        Add-RiskReason -Result $r -Level 3 -Reason ("前一 Run critical alarm {0} 次 / interlock {1} 次, 機台穩定度高風險。" -f $critical, $interlock)
    } elseif ($critical -ge 1) {
        # SOP 8.2: 前一 Run 有 critical alarm -> 至少 Medium
        Add-RiskReason -Result $r -Level 1 -Reason ("前一 Run 出現 critical alarm {0} 次, 依規則至少列為中風險。" -f $critical)
    }
    if ($alarms -ge $alarmMed) {
        Add-RiskReason -Result $r -Level 1 -Reason ("前一 Run alarm 總數 {0} 次, 高於警戒值 {1}。" -f $alarms, $alarmMed)
    }
    if (-not (Test-FieldMissing -Object $StabilityRow -PropertyName "VacuumRecoveryTimeMin")) {
        $vac = Get-FieldDouble -Object $StabilityRow -PropertyName "VacuumRecoveryTimeMin" -DefaultValue 0
        if ($vac -gt $vacWarn) {
            Add-RiskReason -Result $r -Level 1 -Reason ("真空恢復時間 {0} 分, 高於警戒值 {1} 分。" -f $vac, $vacWarn)
        }
    }

    if ($null -ne $PrevRunRow) {
        $prevResult = (Get-FieldString -Object $PrevRunRow -PropertyName "RunResult" -DefaultValue "").ToUpper()
        if ($prevResult -eq "ABNORMAL" -or $prevResult -eq "ABORT") {
            Add-RiskReason -Result $r -Level 3 -Reason ("前一 Run 結果為 {0}, 屬異常中止, 機台穩定度高風險。" -f $prevResult)
        } elseif ($prevResult -eq "HOLD") {
            Add-RiskReason -Result $r -Level 2 -Reason "前一 Run 結果為 Hold, 需確認 Hold 原因是否已排除。"
        }
    } else {
        $Missing.Items += "缺少前一 Run 基本資料 (RunSummary)"
    }

    if ($r.Reasons.Count -eq 0) {
        $r.Reasons += "前一 Run 無 critical alarm, 穩定度分數皆在門檻內, trend 未見異常。"
    }
    return $r
}

function Get-MaintenanceRisk {
    # 維修機率診斷 (SOP 10.2)
    param([object]$MaintRow, [object]$StabilityRow, [hashtable]$Missing)
    $r = @{ Level = 0; Reasons = @() }
    if ($null -eq $MaintRow) {
        $Missing.Items += "缺少維修 / PM 特徵資料 (Maintenance)"
        Add-RiskReason -Result $r -Level 1 -Reason "無 PM / 維修資料, 無法評估維修風險, 保守列為中風險。"
        return $r
    }
    $pmTarget  = Get-Threshold -Name "PMTargetDays" -DefaultValue 21
    $pmRatio   = Get-Threshold -Name "PMWarnRatio" -DefaultValue 0.8
    $sameAlarmHigh = [int](Get-Threshold -Name "SameAlarmHighCount" -DefaultValue 3)
    $watchRuns = [int](Get-Threshold -Name "PostPmWatchRuns" -DefaultValue 3)

    $daysAfterPm = Get-FieldDouble -Object $MaintRow -PropertyName "DaysAfterPM" -DefaultValue 0
    if (Test-FieldMissing -Object $MaintRow -PropertyName "DaysAfterPM") {
        $Missing.Items += "缺少欄位 DaysAfterPM"
    } elseif ($daysAfterPm -gt $pmTarget) {
        Add-RiskReason -Result $r -Level 3 -Reason ("距上次 PM {0} 天, 已超過 PM 目標週期 {1} 天。" -f $daysAfterPm, $pmTarget)
    } elseif ($daysAfterPm -gt ($pmTarget * $pmRatio)) {
        # SOP 10.2: DaysAfterPM > 目標週期 80% -> Medium
        Add-RiskReason -Result $r -Level 1 -Reason ("距上次 PM {0} 天, 超過目標週期 {1} 天的 {2:P0}。" -f $daysAfterPm, $pmTarget, $pmRatio)
    }

    $sameAlarm = Get-FieldInt -Object $MaintRow -PropertyName "RecentSameAlarmCount7d" -DefaultValue 0
    if ($sameAlarm -ge $sameAlarmHigh) {
        # SOP 8.2: 同一 alarm 7 天內 >= 3 次 -> High
        Add-RiskReason -Result $r -Level 3 -Reason ("同一 alarm 7 天內出現 {0} 次 (門檻 {1}), 維修風險高。" -f $sameAlarm, $sameAlarmHigh)
    } elseif ($sameAlarm -ge 2) {
        Add-RiskReason -Result $r -Level 1 -Reason ("同一 alarm 7 天內出現 {0} 次, 需留意重複性。" -f $sameAlarm)
    }

    if ($null -ne $StabilityRow) {
        $crit24h = Test-TruthyFlag (Get-ObjectPropertyValueSafe -Object $StabilityRow -PropertyName "CriticalAlarmWithin24h" -DefaultValue "0")
        if ($crit24h) {
            # SOP 10.2: critical alarm 24 小時內發生 -> High
            Add-RiskReason -Result $r -Level 3 -Reason "24 小時內曾發生 critical alarm, 維修風險高。"
        }
    }

    $runsAfterPm = Get-FieldInt -Object $MaintRow -PropertyName "RunsAfterPM" -DefaultValue 999
    $alarms = 0
    if ($null -ne $StabilityRow) { $alarms = Get-FieldInt -Object $StabilityRow -PropertyName "AlarmCountLastRun" -DefaultValue 0 }
    if ($runsAfterPm -le $watchRuns -and $alarms -gt 0) {
        # SOP 8.2: PM 後 3 Run 內且 alarm 增加 -> Medium
        Add-RiskReason -Result $r -Level 1 -Reason ("PM 後第 {0} Run (觀察期 {1} Run 內) 且前一 Run 有 alarm {2} 次。" -f $runsAfterPm, $watchRuns, $alarms)
    }

    $mtbf = (Get-FieldString -Object $MaintRow -PropertyName "MTBFTrend" -DefaultValue "").ToUpper()
    if ($mtbf -eq "DOWN") {
        # SOP 10.2: MTBF 連續下降 -> Medium / High
        Add-RiskReason -Result $r -Level 2 -Reason "MTBF 趨勢下降, 故障間隔縮短中。"
    }

    $recentCm = Test-TruthyFlag (Get-ObjectPropertyValueSafe -Object $MaintRow -PropertyName "RecentCorrectiveMaintenance" -DefaultValue "No")
    $golden = (Get-FieldString -Object $MaintRow -PropertyName "GoldenRunVerified" -DefaultValue "NA").ToUpper()
    if ($recentCm -and $golden -eq "NO") {
        # SOP 10.2: 維修後未完成 golden run 驗證 -> High
        Add-RiskReason -Result $r -Level 3 -Reason "近期有 corrective maintenance 且尚未完成 golden run 驗證。"
    } elseif ($recentCm) {
        Add-RiskReason -Result $r -Level 1 -Reason "近期曾執行 corrective maintenance, 建議確認維修後首批結果。"
    }

    if ($r.Reasons.Count -eq 0) {
        $r.Reasons += "PM 週期內、無重複 alarm、MTBF 趨勢正常, 維修風險低。"
    }
    return $r
}

function Get-DriftRisk {
    # 產品特性飄移診斷 (SOP 10.3)
    param([object]$DriftRow, [hashtable]$Missing)
    $r = @{ Level = 0; Reasons = @() }
    if ($null -eq $DriftRow) {
        $Missing.Items += "缺少產品特性飄移資料 (ProductDrift)"
        Add-RiskReason -Result $r -Level 1 -Reason "無 Run 後量測趨勢資料, 無法評估產品漂移, 保守列為中風險。"
        return $r
    }
    $plRuns  = [int](Get-Threshold -Name "PLConsecutiveRuns" -DefaultValue 3)
    $thkWarn = Get-Threshold -Name "ThicknessWarnPct" -DefaultValue 2.0
    $thkHigh = Get-Threshold -Name "ThicknessHighPct" -DefaultValue 3.0
    $rsWarn  = Get-Threshold -Name "RsWarnPct" -DefaultValue 2.0

    $overControl = Test-TruthyFlag (Get-ObjectPropertyValueSafe -Object $DriftRow -PropertyName "OverSpcControl" -DefaultValue "0")
    $overWarning = Test-TruthyFlag (Get-ObjectPropertyValueSafe -Object $DriftRow -PropertyName "OverSpcWarning" -DefaultValue "0")
    if ($overControl) {
        # SOP 10.3: 超過 control limit -> 高風險
        Add-RiskReason -Result $r -Level 3 -Reason "量測值已超過 SPC control limit。"
    } elseif ($overWarning) {
        # SOP 10.3: 超過 SPC warning limit -> 中高風險
        Add-RiskReason -Result $r -Level 2 -Reason "量測值已超過 SPC warning limit。"
    }

    $plConsec = Get-FieldInt -Object $DriftRow -PropertyName "ConsecutivePLShiftRuns" -DefaultValue 0
    $plShift  = Get-FieldDouble -Object $DriftRow -PropertyName "PLPeakShiftNm" -DefaultValue 0
    if ($plConsec -ge $plRuns) {
        # SOP 8.2: PL peak 連續 3 Run 同方向偏移 -> Medium
        Add-RiskReason -Result $r -Level 1 -Reason ("PL peak 連續 {0} Run 同方向偏移 (最近偏移 {1} nm)。" -f $plConsec, $plShift)
    }

    if (Test-FieldMissing -Object $DriftRow -PropertyName "ThicknessDeltaPct") {
        $Missing.Items += "缺少欄位 ThicknessDeltaPct"
    } else {
        $thk = Get-FieldDouble -Object $DriftRow -PropertyName "ThicknessDeltaPct" -DefaultValue 0
        if ([math]::Abs($thk) -gt $thkHigh) {
            Add-RiskReason -Result $r -Level 3 -Reason ("厚度相對偏差 {0}%, 超過高風險門檻 {1}%。" -f $thk, $thkHigh)
        } elseif ([math]::Abs($thk) -gt $thkWarn) {
            Add-RiskReason -Result $r -Level 1 -Reason ("厚度相對偏差 {0}%, 超過警戒門檻 {1}%。" -f $thk, $thkWarn)
        }
    }

    $rs = Get-FieldDouble -Object $DriftRow -PropertyName "RsDeltaPct" -DefaultValue 0
    if ([math]::Abs($rs) -gt $rsWarn) {
        Add-RiskReason -Result $r -Level 1 -Reason ("Rs 相對偏差 {0}%, 超過警戒門檻 {1}%。" -f $rs, $rsWarn)
    }

    $uni = (Get-FieldString -Object $DriftRow -PropertyName "UniformityTrend" -DefaultValue "").ToUpper()
    if ($uni -eq "WORSE") {
        Add-RiskReason -Result $r -Level 1 -Reason "均勻性趨勢惡化, 建議確認 showerhead / rotation / thermal field。"
    }
    $aoi = (Get-FieldString -Object $DriftRow -PropertyName "AOIDefectTrend" -DefaultValue "").ToUpper()
    if ($aoi -eq "UP") {
        Add-RiskReason -Result $r -Level 1 -Reason "AOI defect 趨勢上升, 建議確認 particle / chamber memory。"
    }
    $plInt = (Get-FieldString -Object $DriftRow -PropertyName "PLIntensityTrend" -DefaultValue "").ToUpper()
    if ($plInt -eq "DOWN") {
        Add-RiskReason -Result $r -Level 1 -Reason "PL 強度趨勢下降, 可能與材料品質或 interface 有關。"
    }

    if ($r.Reasons.Count -eq 0) {
        $r.Reasons += "PL / XRD / Thickness / Rs / AOI 趨勢皆在門檻內, 未見明顯漂移。"
    }
    return $r
}

function Get-ContextChecks {
    # 情境補查項: 產品族切換 / Recipe 版本變更 (SOP 8.2)
    param([object]$RunRow, [hashtable]$Result)
    if ($null -eq $RunRow) { return }
    $pf   = Get-FieldString -Object $RunRow -PropertyName "ProductFamily"
    $prev = Get-FieldString -Object $RunRow -PropertyName "PreviousProductFamily"
    if (-not [string]::IsNullOrEmpty($prev) -and $prev -ne $pf) {
        Add-RiskReason -Result $Result -Level 1 -Reason ("本 Run 產品族 {0} 與前一 Run {1} 不同, chamber memory effect 需人工確認。" -f $pf, $prev)
    }
    $ecn = Get-FieldString -Object $RunRow -PropertyName "RecipeChangeNote"
    $verChanged = Test-TruthyFlag (Get-ObjectPropertyValueSafe -Object $RunRow -PropertyName "RecipeVersionChanged" -DefaultValue "0")
    if ($verChanged -and [string]::IsNullOrEmpty($ecn)) {
        # SOP 8.2: Recipe 版本有相對變更但無 ECN 備註 -> 補查項
        Add-RiskReason -Result $Result -Level 1 -Reason "Recipe 版本群組有變更但無 ECN 備註, 列為補查項。"
    }
}

# ============================================================
# 相似案例搜尋 (Layer 2 簡化版: 特徵比對)
# ============================================================
function Find-SimilarCases {
    param([object[]]$Cases, [string]$ToolAlias, [string]$ProductFamily, [string[]]$RiskCategories, [string[]]$Keywords)
    $maxN = [int](Get-Threshold -Name "SimilarCaseMax" -DefaultValue 5)
    $scored = @()
    foreach ($c in $Cases) {
        $score = 0
        if ((Get-FieldString -Object $c -PropertyName "ToolAlias") -eq $ToolAlias) { $score += 2 }
        if ((Get-FieldString -Object $c -PropertyName "ProductFamily") -eq $ProductFamily) { $score += 2 }
        $cat = Get-FieldString -Object $c -PropertyName "RiskCategory"
        foreach ($rc in $RiskCategories) {
            if ($cat -eq $rc) { $score += 3 }
        }
        $kwRaw = Get-FieldString -Object $c -PropertyName "Keywords"
        if (-not [string]::IsNullOrEmpty($kwRaw)) {
            $caseKws = $kwRaw.Split(";")
            foreach ($k in $Keywords) {
                foreach ($ck in $caseKws) {
                    if ($ck.Trim().ToUpper() -eq $k.Trim().ToUpper() -and $k.Trim() -ne "") { $score += 2 }
                }
            }
        }
        if ($score -gt 0) {
            $scored += ,@{ Score = $score; Case = $c }
        }
    }
    $sorted = @($scored | Sort-Object -Property @{ Expression = { $_.Score }; Descending = $true })
    $top = @()
    $i = 0
    foreach ($s in $sorted) {
        if ($i -ge $maxN) { break }
        $top += ,$s
        $i++
    }
    return $top
}

# ============================================================
# 診斷主流程 (SOP 9 流程圖)
# ============================================================
function Invoke-RunDiagnosis {
    <#
      對指定 Run 執行 Run 前診斷, 回傳診斷結果 hashtable。
      僅做: 整理 / 比對 / 提醒 / 建議。不判定開 Run 與否。
    #>
    param([string]$RunAlias)

    $runRows   = @(Get-DataRows -TableName "RunSummary")
    $stabRows  = @(Get-DataRows -TableName "ToolStability")
    $maintRows = @(Get-DataRows -TableName "Maintenance")
    $driftRows = @(Get-DataRows -TableName "ProductDrift")
    $caseRows  = @(Get-DataRows -TableName "HistoryCases")

    $runRow = Find-RowByRun -Rows $runRows -RunAlias $RunAlias

    # v1.0.10: Log 檔名目錄一律查找 (RunSummary 已建檔的 Run 也要保有
    # 秒級 Log 比對與量測總檔整合); 僅 RunSummary 查無時才以檔名建立基本資料
    $logCatalog = @(Get-RunLogCatalog)
    $logEntry = $null
    foreach ($e in $logCatalog) {
        if ([string]$e.RunId -eq $RunAlias) { $logEntry = $e; break }
    }
    $isSyntheticRun = $false
    if ($null -eq $runRow) {
        if ($null -eq $logEntry) { throw ("找不到 Run 基本資料: " + $RunAlias) }
        $isSyntheticRun = $true
        $runRow = @{
            RunID_Alias   = $RunAlias
            ToolAlias     = [string]$logEntry.Tool
            ProductFamily = [string]$logEntry.Product
            ChamberAlias  = ""
            RecipeFamily  = ""
            RunStartTime  = ""
        }
    }

    $toolAlias = Get-FieldString -Object $runRow -PropertyName "ToolAlias"
    $pf        = Get-FieldString -Object $runRow -PropertyName "ProductFamily"

    # 前一 Run: 同機台、RunStartTime 早於本 Run 的最近一筆
    $prevRunRow = $null
    $thisStart = Get-FieldString -Object $runRow -PropertyName "RunStartTime"
    foreach ($r in $runRows) {
        if ((Get-FieldString -Object $r -PropertyName "ToolAlias") -ne $toolAlias) { continue }
        $st = Get-FieldString -Object $r -PropertyName "RunStartTime"
        if ([string]::Compare($st, $thisStart) -ge 0) { continue }
        if ($null -eq $prevRunRow) { $prevRunRow = $r }
        else {
            $pst = Get-FieldString -Object $prevRunRow -PropertyName "RunStartTime"
            if ([string]::Compare($st, $pst) -gt 0) { $prevRunRow = $r }
        }
    }

    $prevRunAlias = ""
    if ($null -ne $prevRunRow) { $prevRunAlias = Get-FieldString -Object $prevRunRow -PropertyName "RunID_Alias" }

    # 上上 Run (N-2): 資料時效限制 — Run 前僅能取得 N-1 機台 Log,
    # PL / XRD / Thickness / Rs / AOI 量測資料最快只到 N-2
    $prevPrevRunRow = $null
    if ($null -ne $prevRunRow) {
        $prevStart = Get-FieldString -Object $prevRunRow -PropertyName "RunStartTime"
        foreach ($r in $runRows) {
            if ((Get-FieldString -Object $r -PropertyName "ToolAlias") -ne $toolAlias) { continue }
            $st = Get-FieldString -Object $r -PropertyName "RunStartTime"
            if ([string]::Compare($st, $prevStart) -ge 0) { continue }
            if ($null -eq $prevPrevRunRow) { $prevPrevRunRow = $r }
            else {
                $ppst = Get-FieldString -Object $prevPrevRunRow -PropertyName "RunStartTime"
                if ([string]::Compare($st, $ppst) -gt 0) { $prevPrevRunRow = $r }
            }
        }
    }
    $prevPrevRunAlias = ""
    if ($null -ne $prevPrevRunRow) { $prevPrevRunAlias = Get-FieldString -Object $prevPrevRunRow -PropertyName "RunID_Alias" }

    # v1.0.7/1.0.10: 由 Log 檔名目錄以同機台 Run 碼數字排序推得前一 / 上上 Run
    # (供秒級 Log 比對用; RunSummary 已推得的前一 Run 不覆蓋, 僅補空)
    $prevLogEntry = $null
    if ($null -ne $logEntry) {
        $thisId = [long]$RunAlias
        $earlier = @()
        foreach ($e in $logCatalog) {
            if ([string]$e.Tool -ne $toolAlias) { continue }
            if ([long]$e.RunId -lt $thisId) { $earlier += $e }
        }
        $earlier = @($earlier | Sort-Object -Property @{ Expression = { [long]$_.RunId }; Descending = $true })
        if ($earlier.Count -ge 1) {
            $prevLogEntry = $earlier[0]
            if ([string]::IsNullOrEmpty($prevRunAlias)) { $prevRunAlias = [string]$earlier[0].RunId }
        }
        if ($earlier.Count -ge 2 -and [string]::IsNullOrEmpty($prevPrevRunAlias)) {
            $prevPrevRunAlias = [string]$earlier[1].RunId
        }
    }

    # 特徵資料:
    #   機台面 (ToolStability / Maintenance): 以「前一 Run (N-1)」為主 — Log 可即時取得
    #   量測面 (ProductDrift):               以「上上 Run (N-2)」為主 — 量測資料時效限制
    $stabRow  = $null
    $maintRow = $null
    $driftRow = $null
    $driftSourceAlias = ""
    if (-not [string]::IsNullOrEmpty($prevRunAlias)) {
        $stabRow  = Find-RowByRun -Rows $stabRows  -RunAlias $prevRunAlias -ToolAlias $toolAlias
        $maintRow = Find-RowByRun -Rows $maintRows -RunAlias $prevRunAlias -ToolAlias $toolAlias
    }
    if (-not [string]::IsNullOrEmpty($prevPrevRunAlias)) {
        $driftRow = Find-RowByRun -Rows $driftRows -RunAlias $prevPrevRunAlias -ToolAlias $toolAlias
        if ($null -ne $driftRow) { $driftSourceAlias = $prevPrevRunAlias }
    }
    if ($null -eq $driftRow -and -not [string]::IsNullOrEmpty($prevRunAlias)) {
        # 若 N-1 量測例外地已產出 (或無 N-2 資料), 退回 N-1
        $driftRow = Find-RowByRun -Rows $driftRows -RunAlias $prevRunAlias -ToolAlias $toolAlias
        if ($null -ne $driftRow) { $driftSourceAlias = $prevRunAlias }
    }
    if ($null -eq $stabRow)  { $stabRow  = Find-RowByRun -Rows $stabRows  -RunAlias $RunAlias -ToolAlias $toolAlias }
    if ($null -eq $maintRow) { $maintRow = Find-RowByRun -Rows $maintRows -RunAlias $RunAlias -ToolAlias $toolAlias }
    if ($null -eq $driftRow) {
        $driftRow = Find-RowByRun -Rows $driftRows -RunAlias $RunAlias -ToolAlias $toolAlias
        if ($null -ne $driftRow) { $driftSourceAlias = $RunAlias }
    }

    # v1.0.8: Log 檔名目錄模式 — ProductDrift 查無資料時, 改以量測總檔
    # (STRUCTURE,REACTOR,RUN_NO,POS_NO,PF,...) 為產品特性飄移來源;
    # 取 RunId <= 本 Run 的最近量測 Run (符合量測資料時效, 通常為 N-2)
    $measInfo = $null
    if ($null -ne $logEntry -and $null -eq $driftRow) {
        $measInfo = Get-MeasurementDriftInfo -ToolAlias $toolAlias -UpToRunId ([long]$RunAlias)
        if ($null -ne $measInfo) {
            $driftRow = $measInfo.DriftRow
            $driftSourceAlias = [string]$measInfo.Current.RunId
        }
    }

    $missing = @{ Items = @() }
    $stab  = Get-StabilityRisk   -StabilityRow $stabRow  -PrevRunRow $prevRunRow -Missing $missing
    $maint = Get-MaintenanceRisk -MaintRow $maintRow -StabilityRow $stabRow -Missing $missing
    $drift = Get-DriftRisk       -DriftRow $driftRow -Missing $missing

    # 資料時效註記: 量測取自 N-2 時提醒工程師 N-1 量測尚無法取得
    if (-not [string]::IsNullOrEmpty($prevPrevRunAlias) -and $driftSourceAlias -eq $prevPrevRunAlias) {
        $drift.Reasons += ("(資料時效) 量測資料取自上上 Run {0};上一 Run {1} 之 PL/XRD/Thickness/Rs/AOI 尚未產出, 請搭配「前 Run Log 診斷」確認 N-1 機台行為。" -f $prevPrevRunAlias, $prevRunAlias)
    }
    if (-not [string]::IsNullOrEmpty($prevPrevRunAlias) -and $null -eq (Find-RowByRun -Rows $driftRows -RunAlias $prevPrevRunAlias -ToolAlias $toolAlias)) {
        $missing.Items += ("缺少上上 Run (N-2) 量測資料 (ProductDrift: " + $prevPrevRunAlias + ")")
    }

    # v1.0.8: 量測總檔摘要與 PF 最終判定 (僅轉述量測系統結果, 不放行/不判 Fail 原因)
    if ($null -ne $measInfo) {
        $mCur = $measInfo.Current
        $prevTxt = "無更早量測 Run 可比較"
        if ($null -ne $measInfo.Previous) { $prevTxt = ("對比前一量測 Run " + [string]$measInfo.Previous.RunId) }
        $drift.Reasons += ("(量測總檔) 來源 Run {0} ({1}, {2} 片: Pass {3} / Fail {4}; {5}; 檔案: {6})" -f `
            [string]$mCur.RunId, [string]$mCur.Structure, $mCur.WaferCount, $mCur.PassCount, $mCur.FailCount, $prevTxt, [string]$mCur.SourceFile)
        if (-not [string]::IsNullOrEmpty([string]$measInfo.ValueSummary)) {
            $drift.Reasons += ("(量測總檔) " + $measInfo.ValueSummary)
        }
        if ($mCur.FailCount -gt 0) {
            $codes = (@($mCur.FailCodes) -join ";")
            if ([string]::IsNullOrEmpty($codes)) { $codes = "未填 FAIL_NO" }
            Add-RiskReason -Result $drift -Level 3 -Reason ("量測最終判定 (PF) 出現 Fail {0} 片 (FAIL_NO: {1}), 請確認 Fail 原因是否與本機台相關。" -f $mCur.FailCount, $codes)
        }
    } elseif ($null -ne $logEntry) {
        $missing.Items += "Import/RawImport 未找到此機台的量測總檔 (STRUCTURE,REACTOR,RUN_NO,POS_NO,PF,... 格式)"
    }

    # v1.0.7: Log 檔名目錄模式 — 自動以本 Run Log vs 前一 Run Log 做
    # StepCode 平均差異比對, 結果併入機台穩定度 (機台行為面提前預警)
    $logCmpResult = $null
    if ($null -ne $logEntry) {
        if ($null -ne $prevLogEntry) {
            try {
                $tSel  = Get-StepCodeMeanTable -LogPath ([string]$logEntry.Path)
                $tPrev = Get-StepCodeMeanTable -LogPath ([string]$prevLogEntry.Path)
                $lw = Get-Threshold -Name "LogDeltaWarnPct" -DefaultValue 3.0
                $lh = Get-Threshold -Name "LogDeltaHighPct" -DefaultValue 8.0
                $lm = [int](Get-Threshold -Name "LogCompareMinSec" -DefaultValue 5)
                $lt = [int](Get-Threshold -Name "LogCompareTopN"   -DefaultValue 10)
                $lb = Get-Threshold -Name "LogCompareMinBaseAbs" -DefaultValue 0.0
                $logCmp = Compare-StepCodeMeanTables -BaseTable $tPrev -TestTable $tSel `
                    -BaseName ("前一 Run " + [string]$prevLogEntry.RunId) -TestName ("Run " + $RunAlias) `
                    -WarnPct $lw -HighPct $lh -MinSec $lm -TopN $lt -MinBaseAbs $lb `
                    -ExcludeParams (@(Get-LogCompareExcludePatterns))
                $logCmpResult = $logCmp
                foreach ($x in $logCmp.Reasons) { $stab.Reasons += ("(Log 比對) " + $x) }
                if ($logCmp.Level -gt $stab.Level) { $stab.Level = $logCmp.Level }
            } catch {
                $missing.Items += ("Log 檔比對失敗: " + $_.Exception.Message)
            }
        } else {
            $missing.Items += "同機台無更早的 Log 檔, 無法進行本 Run vs 前一 Run 秒級 Log 比對"
        }
        if ($isSyntheticRun) {
            $missing.Items += "此 Run 由 Log 檔名建立, RunSummary / ToolStability / Maintenance / ProductDrift 尚無對應資料"
        }
    }

    $overall = @{ Level = (Measure-MaxRisk @($stab.Level, $maint.Level, $drift.Level)); Reasons = @() }
    Get-ContextChecks -RunRow $runRow -Result $overall
    if ($isSyntheticRun) {
        $overall.Reasons += ("(資料來源) 本 Run 資訊由 Log 檔名解析建立 (檔案: {0}); 表格資料備齊前, 缺漏面向以保守等級與 Log 比對呈現。" -f [string]$logEntry.FileName)
    }

    # 相似案例
    $riskCats = @()
    if ($stab.Level  -ge 1) { $riskCats += "Stability" }
    if ($maint.Level -ge 1) { $riskCats += "Maintenance" }
    if ($drift.Level -ge 1) { $riskCats += "Drift" }
    $keywords = @()
    if ($null -ne $stabRow) {
        $codes = Get-FieldString -Object $stabRow -PropertyName "RecentAlarmCodes"
        if (-not [string]::IsNullOrEmpty($codes)) { $keywords += $codes.Split(";") }
    }
    $similar = @(Find-SimilarCases -Cases $caseRows -ToolAlias $toolAlias -ProductFamily $pf -RiskCategories $riskCats -Keywords $keywords)

    # 建議人工檢查項 (依風險組合, 最多 8 項; SOP 12.1)
    $checks = @()
    if ($stab.Level -ge 1) {
        $checks += "查前一 Run process trend (溫度 / 壓力 / MFC / 轉速) 原始曲線。"
        $checks += "查前一 Run alarm / event log 與處置紀錄。"
    }
    if ($maint.Level -ge 1) {
        $checks += "查 PM / 換件 / corrective maintenance 紀錄與 golden run 驗證結果。"
        $checks += "確認清腔紀錄與 seasoning run 結果。"
    }
    if ($drift.Level -ge 1) {
        $checks += "比對同產品族最近 5 Run 的 PL / XRD / Thickness / Rs trend。"
        $checks += "確認本 Run 是否需要增加 Run 後量測項目。"
    }
    foreach ($reason in $overall.Reasons) {
        if ($reason -match "chamber memory") { $checks += "確認產品族切換後之 chamber 狀態 (清腔 / seasoning / dummy run)。" }
        if ($reason -match "ECN") { $checks += "向製程主管確認 Recipe 版本變更之 ECN 與變更原因。" }
    }
    if (-not [string]::IsNullOrEmpty($prevRunAlias)) {
        $checks += "以「前 Run Log 診斷」比對上一 Run (N-1) 與上上 Run (N-2) 之秒級機台 Log 差異。"
    }
    $checks += "由製程與設備工程師共同覆判本 Run 之後續處置。"
    if ($checks.Count -gt 8) { $checks = @($checks[0..7]) }

    return @{
        RunAlias      = $RunAlias
        PrevRunAlias  = $prevRunAlias
        PrevPrevRunAlias = $prevPrevRunAlias
        DriftSourceRun   = $driftSourceAlias
        ToolAlias     = $toolAlias
        ChamberAlias  = Get-FieldString -Object $runRow -PropertyName "ChamberAlias"
        ProductFamily = $pf
        RecipeFamily  = Get-FieldString -Object $runRow -PropertyName "RecipeFamily"
        Stability     = $stab
        Maintenance   = $maint
        Drift         = $drift
        Overall       = $overall
        Checks        = $checks
        MissingData   = $missing.Items
        SimilarCases  = $similar
        LogCompare    = $logCmpResult
        DataSources   = $(if ($null -ne $logEntry) {
                            @("RunLog 檔名目錄 (" + [string]$logEntry.FileName + ")", "RunSummary", "ToolStability", "Maintenance", "ProductDrift", "HistoryCases")
                          } else {
                            @("RunSummary", "ToolStability", "Maintenance", "ProductDrift", "HistoryCases")
                          })
        GeneratedAt   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        AIVersion     = $script:AIVersion
    }
}

# ============================================================
# Layer 3: 模板式診斷報告 (SOP 11 格式; 不使用外部 LLM)
# ============================================================
function New-DiagnosisReportText {
    param([hashtable]$D)
    $nl = [Environment]::NewLine
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Run 前 AI 診斷報告")
    [void]$sb.AppendLine("產出時間:" + $D.GeneratedAt)
    [void]$sb.AppendLine("AI 版本:" + $D.AIVersion + " (規則引擎 + 模板摘要, 未使用外部 LLM)")
    $ppText = ""
    if ($D.ContainsKey("PrevPrevRunAlias") -and -not [string]::IsNullOrEmpty([string]$D.PrevPrevRunAlias)) {
        $ppText = "  上上 Run:" + $D.PrevPrevRunAlias
    }
    [void]$sb.AppendLine("Run:" + $D.RunAlias + "  (前一 Run:" + $D.PrevRunAlias + $ppText + ")")
    [void]$sb.AppendLine("Tool:" + $D.ToolAlias + "  Chamber:" + $D.ChamberAlias)
    [void]$sb.AppendLine("ProductFamily:" + $D.ProductFamily + "  RecipeGroup:" + $D.RecipeFamily)
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("一、總體風險")
    [void]$sb.AppendLine("風險等級:" + (Get-RiskName $D.Overall.Level))
    [void]$sb.AppendLine("建議狀態:需製程 / 設備工程師覆判後決定後續處置")
    if ($D.Overall.Reasons.Count -gt 0) {
        [void]$sb.AppendLine("情境提醒:")
        $i = 1
        foreach ($x in $D.Overall.Reasons) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
    }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("二、機台穩定度")
    [void]$sb.AppendLine("等級:" + (Get-RiskName $D.Stability.Level))
    [void]$sb.AppendLine("原因:")
    $i = 1
    foreach ($x in $D.Stability.Reasons) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
    # v1.0.13: Log 比對差異排行 (主報告原僅有「詳見差異排行」而無排行內容)
    $logCmpRpt = $null
    if ($D.ContainsKey("LogCompare")) { $logCmpRpt = $D.LogCompare }
    if ($null -ne $logCmpRpt -and @($logCmpRpt.TopDeltas).Count -gt 0) {
        $invLc = [System.Globalization.CultureInfo]::InvariantCulture
        [void]$sb.AppendLine("Log 比對差異排行 (|差異%| 由大到小):")
        foreach ($dd in $logCmpRpt.TopDeltas) {
            [void]$sb.AppendLine(("  Step/Layer {0} | {1} | 前一 Run {2} | 本 Run {3} | 差異 {4}% | 等級 {5}" -f `
                $dd.StepKey, $dd.Param, $dd.BaseMean.ToString("0.####", $invLc), $dd.TestMean.ToString("0.####", $invLc), `
                $dd.DeltaPct.ToString("+0.##;-0.##", $invLc), (Get-RiskName $dd.Level)))
        }
        if (@($logCmpRpt.SkippedSteps).Count -gt 0) {
            [void]$sb.AppendLine(("  略過短 Step (秒數 < 門檻): " + (@($logCmpRpt.SkippedSteps) -join ";")))
        }
    }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("三、維修風險")
    [void]$sb.AppendLine("等級:" + (Get-RiskName $D.Maintenance.Level))
    [void]$sb.AppendLine("原因:")
    $i = 1
    foreach ($x in $D.Maintenance.Reasons) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("四、產品特性飄移")
    [void]$sb.AppendLine("等級:" + (Get-RiskName $D.Drift.Level))
    if ($D.ContainsKey("DriftSourceRun") -and -not [string]::IsNullOrEmpty([string]$D.DriftSourceRun)) {
        [void]$sb.AppendLine("量測資料來源 Run:" + $D.DriftSourceRun + " (資料時效: PL/XRD/Thickness/Rs/AOI 最快僅到上上 Run)")
    }
    [void]$sb.AppendLine("原因:")
    $i = 1
    foreach ($x in $D.Drift.Reasons) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("五、建議人工檢查")
    $i = 1
    foreach ($x in $D.Checks) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("六、資料不足處")
    if ($D.MissingData.Count -eq 0) {
        [void]$sb.AppendLine("本次診斷所需欄位皆有資料。")
    } else {
        $i = 1
        foreach ($x in $D.MissingData) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
    }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("七、相似歷史案例")
    if ($D.SimilarCases.Count -eq 0) {
        [void]$sb.AppendLine("未找到相似案例。")
    } else {
        foreach ($s in $D.SimilarCases) {
            $c = $s.Case
            [void]$sb.AppendLine(("- {0} | {1} | {2} | 處置:{3} | 結果:{4}" -f `
                (Get-FieldString -Object $c -PropertyName "CaseID"),
                (Get-FieldString -Object $c -PropertyName "RiskCategory"),
                (Get-FieldString -Object $c -PropertyName "Summary"),
                (Get-FieldString -Object $c -PropertyName "Action"),
                (Get-FieldString -Object $c -PropertyName "Outcome")))
        }
    }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("八、AI 不可判定事項")
    [void]$sb.AppendLine("AI 不判定是否可開 Run。")
    [void]$sb.AppendLine("AI 不建議修改正式 Recipe。")
    [void]$sb.AppendLine("AI 不判定產品放行或報廢。")
    [void]$sb.AppendLine("AI 不取代 SPC、MES、QMS 或工程師覆判。")
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("九、資料來源摘要 (可追溯性)")
    foreach ($ds in $D.DataSources) { [void]$sb.AppendLine("- " + $ds) }
    [void]$sb.AppendLine("本報告僅使用去識別化資料, 未含真實機密參數。")

    return $sb.ToString()
}

function Save-DiagnosisReport {
    param([hashtable]$D)
    $fname = "RunPreCheck_{0}_{1}.txt" -f (ConvertTo-SafeFileName $D.RunAlias), (Get-Date -Format "yyyyMMdd_HHmmss")
    $path = Join-Path $script:ReportRoot $fname
    Write-AllTextUtf8 -Path $path -Text (New-DiagnosisReportText -D $D)
    Write-AuditLog -Action "REPORT_SAVED" -Detail ("{0} overall={1}" -f $D.RunAlias, (Get-RiskName $D.Overall.Level))
    return $path
}

# ============================================================
# 覆判留痕 (SOP 6.5 HumanReview)
# ============================================================
function Initialize-HumanReviewCsv {
    if (-not (Test-Path -LiteralPath $script:HumanReviewPath)) {
        $header = "ReviewTime,RunID_Alias,ToolAlias,AI_OverallRisk,AI_StabilityRisk,AI_MaintRisk,AI_DriftRisk,AI_Version,EngineerReview,FinalDecision,FalsePositive,FalseNegative,ReviewerRole,Comment,ReportFile"
        Write-FileSafe -Path $script:HumanReviewPath -Text $header
    }
}

function Save-HumanReview {
    param([hashtable]$D, [string]$EngineerReview, [string]$FinalDecision,
          [string]$FalsePositive, [string]$FalseNegative,
          [string]$ReviewerRole, [string]$Comment, [string]$ReportFile)
    Initialize-HumanReviewCsv
    $reportName = ""
    if (-not [string]::IsNullOrEmpty($ReportFile)) { $reportName = Split-Path -Leaf $ReportFile }
    $line = Join-CsvLine @(
        (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
        $D.RunAlias, $D.ToolAlias,
        (Get-RiskName $D.Overall.Level), (Get-RiskName $D.Stability.Level),
        (Get-RiskName $D.Maintenance.Level), (Get-RiskName $D.Drift.Level),
        $D.AIVersion, $EngineerReview, $FinalDecision,
        $FalsePositive, $FalseNegative, $ReviewerRole, $Comment,
        $reportName
    )
    Add-LinesSafe -Path $script:HumanReviewPath -Lines @($line)
    Write-AuditLog -Action "REVIEW_SAVED" -Detail ("{0} decision={1}" -f $D.RunAlias, $FinalDecision)
}

# ============================================================
# 自我測試 (-SelfTest): 不開 GUI, 於暫存資料夾驗證核心邏輯
# ============================================================
function New-TestRow {
    param([hashtable]$Fields)
    $o = New-Object PSObject
    foreach ($k in $Fields.Keys) {
        $o | Add-Member -MemberType NoteProperty -Name $k -Value $Fields[$k]
    }
    return $o
}

function Invoke-SelfTest {
    $t = @{ Pass = 0; Fail = 0 }
    $assert = {
        param([bool]$Cond, [string]$Name)
        if ($Cond) { $t.Pass++; Write-Host ("[PASS] " + $Name) }
        else       { $t.Fail++; Write-Host ("[FAIL] " + $Name) }
    }.GetNewClosure()

    # --- 測試環境: 暫存資料夾 ---
    $tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("RunPreCheckAI_SelfTest_" + [System.Guid]::NewGuid().ToString("N").Substring(0, 8))
    Set-AppRootPaths -Root $tmpRoot
    Initialize-AppFolders
    Initialize-Config
    Write-Host ("SelfTest 暫存資料夾: " + $tmpRoot)

    # --- 1. CSV helpers ---
    & $assert ((ConvertTo-CsvValue 'a,b') -eq '"a,b"') "CSV 逗號跳脫"
    & $assert ((ConvertTo-CsvValue 'say "hi"') -eq '"say ""hi"""') "CSV 引號跳脫"
    & $assert ((Join-CsvLine @("a", "b,c", "")) -eq 'a,"b,c",') "Join-CsvLine 組合"

    # --- 2. 安全欄位存取 (StrictMode) ---
    $row = New-TestRow @{ A = "1"; B = "" }
    & $assert ((Get-FieldInt -Object $row -PropertyName "A") -eq 1) "Get-FieldInt 轉型"
    & $assert ((Get-FieldDouble -Object $row -PropertyName "NotExist" -DefaultValue 9.9) -eq 9.9) "缺欄位回預設值"
    & $assert (Test-FieldMissing -Object $row -PropertyName "B") "空字串視為缺漏"
    & $assert ((Get-RiskName 3) -eq "高") "風險等級名稱"
    & $assert ((Measure-MaxRisk @(0, 2, 1)) -eq 2) "最大風險計算"

    # --- 3. 規則引擎: 機台穩定度 ---
    $m1 = @{ Items = @() }
    $stabHigh = New-TestRow @{
        TempStabilityScore = "65"; PressureStabilityScore = "90"
        MFCStabilityScore = "92"; RotationStabilityScore = "95"
        VacuumRecoveryTimeMin = "10"; AlarmCountLastRun = "2"
        CriticalAlarmCount = "0"; InterlockEventCount = "0"
        CriticalAlarmWithin24h = "0"; RecentAlarmCodes = "TEMP_DRIFT"
    }
    $r1 = Get-StabilityRisk -StabilityRow $stabHigh -PrevRunRow $null -Missing $m1
    & $assert ($r1.Level -eq 3) "穩定度: 分數低於高風險門檻 -> 高"

    $m2 = @{ Items = @() }
    $stabLow = New-TestRow @{
        TempStabilityScore = "95"; PressureStabilityScore = "93"
        MFCStabilityScore = "96"; RotationStabilityScore = "97"
        VacuumRecoveryTimeMin = "8"; AlarmCountLastRun = "0"
        CriticalAlarmCount = "0"; InterlockEventCount = "0"
        CriticalAlarmWithin24h = "0"; RecentAlarmCodes = ""
    }
    $prevOk = New-TestRow @{ RunID_Alias = "RUN_X"; RunResult = "Normal"; RunStartTime = "2026-07-09 08:00:00" }
    $r2 = Get-StabilityRisk -StabilityRow $stabLow -PrevRunRow $prevOk -Missing $m2
    & $assert ($r2.Level -eq 0) "穩定度: 全部正常 -> 低"

    $m2b = @{ Items = @() }
    $stabCrit = New-TestRow @{
        TempStabilityScore = "90"; PressureStabilityScore = "90"
        MFCStabilityScore = "90"; RotationStabilityScore = "90"
        VacuumRecoveryTimeMin = "5"; AlarmCountLastRun = "1"
        CriticalAlarmCount = "1"; InterlockEventCount = "0"
        CriticalAlarmWithin24h = "0"; RecentAlarmCodes = "PRESSURE_SPIKE"
    }
    $r2b = Get-StabilityRisk -StabilityRow $stabCrit -PrevRunRow $prevOk -Missing $m2b
    & $assert ($r2b.Level -ge 1) "穩定度: critical alarm -> 至少中"

    # --- 4. 規則引擎: 維修風險 ---
    $m3 = @{ Items = @() }
    $maintHigh = New-TestRow @{
        DaysAfterPM = "18"; RunsAfterPM = "42"; DaysAfterPartChange = "7"
        RecentSameAlarmCount7d = "5"; MTBFTrend = "Down"
        RecentCorrectiveMaintenance = "No"; GoldenRunVerified = "NA"
    }
    $r3 = Get-MaintenanceRisk -MaintRow $maintHigh -StabilityRow $stabLow -Missing $m3
    & $assert ($r3.Level -eq 3) "維修: 同 alarm 7天內 5 次 -> 高"

    $m4 = @{ Items = @() }
    $maintLow = New-TestRow @{
        DaysAfterPM = "5"; RunsAfterPM = "10"; DaysAfterPartChange = "30"
        RecentSameAlarmCount7d = "0"; MTBFTrend = "Flat"
        RecentCorrectiveMaintenance = "No"; GoldenRunVerified = "NA"
    }
    $r4 = Get-MaintenanceRisk -MaintRow $maintLow -StabilityRow $stabLow -Missing $m4
    & $assert ($r4.Level -eq 0) "維修: 正常 -> 低"

    $m4b = @{ Items = @() }
    $maintCm = New-TestRow @{
        DaysAfterPM = "2"; RunsAfterPM = "1"; DaysAfterPartChange = "2"
        RecentSameAlarmCount7d = "0"; MTBFTrend = "Flat"
        RecentCorrectiveMaintenance = "Yes"; GoldenRunVerified = "No"
    }
    $r4b = Get-MaintenanceRisk -MaintRow $maintCm -StabilityRow $stabLow -Missing $m4b
    & $assert ($r4b.Level -eq 3) "維修: 維修後未 golden run 驗證 -> 高"

    # --- 5. 規則引擎: 產品漂移 ---
    $m5 = @{ Items = @() }
    $driftHigh = New-TestRow @{
        PLPeakShiftNm = "1.8"; ConsecutivePLShiftRuns = "3"; PLIntensityTrend = "Flat"
        XRDPeakShiftDeg = "0.006"; ThicknessDeltaPct = "1.5"; UniformityTrend = "Flat"
        RsDeltaPct = "0.5"; AOIDefectTrend = "Flat"; OverSpcWarning = "0"; OverSpcControl = "1"
    }
    $r5 = Get-DriftRisk -DriftRow $driftHigh -Missing $m5
    & $assert ($r5.Level -eq 3) "漂移: 超過 control limit -> 高"

    $m6 = @{ Items = @() }
    $driftMed = New-TestRow @{
        PLPeakShiftNm = "0.9"; ConsecutivePLShiftRuns = "3"; PLIntensityTrend = "Flat"
        XRDPeakShiftDeg = "0.001"; ThicknessDeltaPct = "0.5"; UniformityTrend = "Flat"
        RsDeltaPct = "0.3"; AOIDefectTrend = "Flat"; OverSpcWarning = "0"; OverSpcControl = "0"
    }
    $r6 = Get-DriftRisk -DriftRow $driftMed -Missing $m6
    & $assert ($r6.Level -eq 1) "漂移: PL 連續 3 Run 同向 -> 中"

    # --- 6. 缺資料保守處理 ---
    $m7 = @{ Items = @() }
    $r7 = Get-DriftRisk -DriftRow $null -Missing $m7
    & $assert ($r7.Level -eq 1 -and $m7.Items.Count -eq 1) "漂移: 無資料 -> 中 + 列入資料不足"

    # --- 7. 原子寫檔 + 檔案鎖 ---
    $p = Join-Path $script:DataRoot "locktest.csv"
    Write-FileSafe -Path $p -Text "Col1,Col2"
    Add-LinesSafe -Path $p -Lines @("a,b", "c,d")
    $content = Read-AllTextUtf8 -Path $p
    & $assert ($content -match "Col1,Col2\r?\na,b\r?\nc,d") "Write-FileSafe + Add-LinesSafe (含換行邊界)"

    # --- 8. 去識別化轉換 ---
    $rawCsv = Join-Path $script:RawImportRoot "RunSummary.csv"
    $rawLines = @(
        "RunID,ToolID,ChamberID,ProductID,RecipeName,PrevProductID,RunStartTime,RunEndTime,RunResult,OperatorShift",
        "R26-0710-001,MOCVD-A7,CH1,VCSEL-940-X2,RCP-940-V13,HBT-2G-B1,2026-07-10 08:00:00,2026-07-10 11:30:00,Normal,Day",
        "R26-0710-002,MOCVD-A7,CH1,VCSEL-940-X2,RCP-940-V13,VCSEL-940-X2,2026-07-10 12:00:00,2026-07-10 15:30:00,Normal,Day"
    )
    Write-AllTextUtf8 -Path $rawCsv -Text ($rawLines -join [Environment]::NewLine)
    $outCsv = Join-Path $script:ImportRoot "RunSummary_deident_test.csv"
    $n = Convert-RawImportFile -RawPath $rawCsv -OutPath $outCsv -MapPath $script:DeidentMapPath
    $outText = Read-AllTextUtf8 -Path $outCsv
    & $assert ($n -eq 2) "去識別化: 轉出 2 筆"
    & $assert (-not ($outText -match "MOCVD-A7")) "去識別化: 真實機台代號未外洩"
    & $assert (-not ($outText -match "VCSEL-940-X2")) "去識別化: 真實料號未外洩"
    & $assert ($outText -match "RunID_Alias") "去識別化: 欄位改名為 Alias"
    & $assert ($outText -match "TOOL_") "去識別化: Tool alias 產生"
    $mapText = Read-AllTextUtf8 -Path $script:DeidentMapPath
    & $assert ($mapText -match "MOCVD-A7") "對照表: 保留原始值供受控追溯"
    # 同值應對應同 alias
    $map2 = Import-DeidentMap -MapPath $script:DeidentMapPath
    $st = @{ NewRows = @() }
    $a1 = Get-DeidentAlias -Map $map2 -Category "TOOL" -RawValue "MOCVD-A7" -Prefix "TOOL" -State $st
    $a2 = Get-DeidentAlias -Map $map2 -Category "TOOL" -RawValue "MOCVD-A7" -Prefix "TOOL" -State $st
    & $assert ($a1 -eq $a2 -and $st.NewRows.Count -eq 0) "對照表: 同值同 alias, 不重複建檔"

    # --- 8b. 去識別化效能修正 (v1.0.6) 行為驗證 ---
    # 計數快取的編號需與原版全表掃描結果相同: 新值接續既有編號
    $st2 = @{ NewRows = @() }
    $a3 = Get-DeidentAlias -Map $map2 -Category "TOOL" -RawValue "MOCVD-B9" -Prefix "TOOL" -State $st2
    & $assert ($a3 -eq "TOOL_002" -and $st2.NewRows.Count -eq 1) "對照表: 新值編號接續既有 (TOOL_002)"
    # 重複轉換: 同值同 alias, 對照表不增列, 輸出內容完全一致
    $mapRowsBefore = @(Import-CsvSafe -Path $script:DeidentMapPath).Count
    $n2 = Convert-RawImportFile -RawPath $rawCsv -OutPath $outCsv -MapPath $script:DeidentMapPath
    $outText2 = Read-AllTextUtf8 -Path $outCsv
    $mapRowsAfter = @(Import-CsvSafe -Path $script:DeidentMapPath).Count
    & $assert ($n2 -eq 2 -and $mapRowsAfter -eq $mapRowsBefore) "去識別化: 重複轉換不重複建檔"
    & $assert ($outText2 -eq $outText) "去識別化: 重複轉換輸出一致"
    Remove-Item -LiteralPath $outCsv -Force -ErrorAction SilentlyContinue

    # --- 9. 端到端: 建立 Import 範例 -> Invoke-RunDiagnosis -> 報告 ---
    $importFiles = @{
        "RunSummary.csv" = @(
            "RunID_Alias,ToolAlias,ChamberAlias,ProductFamily,RecipeFamily,RecipeVersionGroup,PreviousProductFamily,RunStartTime,RunEndTime,RunResult,OperatorShift,RecipeVersionChanged,RecipeChangeNote",
            "RUN_001,TOOL_01,CH_A,PF_VCSEL_A,RF_001,VG_12,PF_HBT_B,2026-07-09 08:00:00,2026-07-09 11:30:00,Normal,Day,0,",
            "RUN_002,TOOL_01,CH_A,PF_VCSEL_A,RF_001,VG_12,PF_VCSEL_A,2026-07-10 08:00:00,,Planned,Day,1,"
        )
        "ToolStability.csv" = @(
            "RunID_Alias,ToolAlias,TempStabilityScore,PressureStabilityScore,MFCStabilityScore,RotationStabilityScore,VacuumRecoveryTimeMin,AlarmCountLastRun,CriticalAlarmCount,InterlockEventCount,CriticalAlarmWithin24h,RecentAlarmCodes",
            "RUN_001,TOOL_01,92,78,90,95,12,3,1,0,1,TEMP_DRIFT;PRESSURE_SPIKE"
        )
        "Maintenance.csv" = @(
            "RunID_Alias,ToolAlias,DaysAfterPM,RunsAfterPM,DaysAfterPartChange,RecentSameAlarmCount7d,MTBFTrend,RecentCorrectiveMaintenance,GoldenRunVerified",
            "RUN_001,TOOL_01,18,42,7,2,Down,No,NA"
        )
        "ProductDrift.csv" = @(
            "RunID_Alias,ToolAlias,ProductFamily,PLPeakShiftNm,ConsecutivePLShiftRuns,PLIntensityTrend,XRDPeakShiftDeg,ThicknessDeltaPct,UniformityTrend,RsDeltaPct,AOIDefectTrend,OverSpcWarning,OverSpcControl",
            "RUN_001,TOOL_01,PF_VCSEL_A,1.2,3,Down,0.004,2.2,Flat,1.0,Flat,1,0"
        )
        "HistoryCases.csv" = @(
            "CaseID,ToolAlias,ProductFamily,RiskCategory,Keywords,Summary,Action,Outcome,CaseDate",
            "CASE_102,TOOL_01,PF_VCSEL_A,Stability,TEMP_DRIFT;PRESSURE_SPIKE,前 Run 溫壓不穩後本 Run 厚度偏移,清腔+seasoning 後恢復,恢復正常,2026-05-12",
            "CASE_155,TOOL_02,PF_HBT_B,Drift,PL_SHIFT,PL 連續偏移與 MFC 漂移相關,校正 MFC,恢復正常,2026-04-03"
        )
    }
    foreach ($fn in $importFiles.Keys) {
        Write-AllTextUtf8 -Path (Join-Path $script:ImportRoot $fn) -Text ($importFiles[$fn] -join [Environment]::NewLine)
    }
    $diag = Invoke-RunDiagnosis -RunAlias "RUN_002"
    & $assert ($diag.PrevRunAlias -eq "RUN_001") "診斷: 正確找到前一 Run"
    & $assert ($diag.Overall.Level -ge 2) "診斷: 綜合風險為中高以上"
    & $assert ($diag.Maintenance.Level -ge 1) "診斷: 維修風險至少中 (24h critical alarm)"
    & $assert ($diag.Drift.Level -ge 2) "診斷: 漂移風險中高 (SPC warning + 厚度)"
    & $assert ($diag.SimilarCases.Count -ge 1) "診斷: 找到相似案例"
    & $assert ($diag.Checks.Count -le 8 -and $diag.Checks.Count -ge 3) "診斷: 檢查項 3~8 項"

    $report = New-DiagnosisReportText -D $diag
    foreach ($section in @("一、總體風險", "二、機台穩定度", "三、維修風險", "四、產品特性飄移",
                           "五、建議人工檢查", "六、資料不足處", "七、相似歷史案例",
                           "八、AI 不可判定事項", "九、資料來源摘要")) {
        & $assert ($report.Contains($section)) ("報告含章節: " + $section)
    }
    & $assert ($report.Contains("AI 不判定是否可開 Run")) "報告含 AI 權限邊界聲明"
    $rpath = Save-DiagnosisReport -D $diag
    & $assert (Test-Path -LiteralPath $rpath) "報告存檔成功"

    # --- 10. 覆判留痕 ---
    Save-HumanReview -D $diag -EngineerReview "同意" -FinalDecision "補查資料" `
        -FalsePositive "" -FalseNegative "" -ReviewerRole "ENG_01" `
        -Comment "測試覆判" -ReportFile $rpath
    $reviews = @(Import-CsvSafe -Path $script:HumanReviewPath)
    & $assert ($reviews.Count -eq 1) "覆判: 寫入 1 筆"
    & $assert ((Get-FieldString -Object $reviews[0] -PropertyName "FinalDecision") -eq "補查資料") "覆判: FinalDecision 正確"
    & $assert ((Get-FieldString -Object $reviews[0] -PropertyName "ReviewerRole") -eq "ENG_01") "覆判: 僅記 RoleCode"

    # --- 11. 使用者管理 ---
    Initialize-UsersCsv
    $newRc = Add-UserAccount -UserId "eng03" -DisplayName "測試工程師" -Role "Engineer"
    & $assert ($newRc -eq "ENG_02") "使用者: RoleCode 自動編號 (ENG_02)"
    & $assert ($null -ne (Find-UserById -UserId "eng03")) "使用者: 新增後可登入查得"
    $dupBlocked = $false
    try { Add-UserAccount -UserId "ENG03" -DisplayName "重複" -Role "Viewer" | Out-Null } catch { $dupBlocked = $true }
    & $assert $dupBlocked "使用者: 重複工號正確拒絕 (不分大小寫)"
    $badRoleBlocked = $false
    try { Add-UserAccount -UserId "x01" -DisplayName "x" -Role "Boss" | Out-Null } catch { $badRoleBlocked = $true }
    & $assert $badRoleBlocked "使用者: 無效角色正確拒絕"
    Set-UserActive -UserId "eng03" -Active $false
    & $assert ($null -eq (Find-UserById -UserId "eng03")) "使用者: 停用後無法登入"
    Set-UserActive -UserId "eng03" -Active $true
    & $assert ($null -ne (Find-UserById -UserId "eng03")) "使用者: 重新啟用後恢復"
    $lastAdminBlocked = $false
    try { Set-UserActive -UserId "admin" -Active $false } catch { $lastAdminBlocked = $true }
    & $assert $lastAdminBlocked "使用者: 最後一位 Admin 不可停用"

    # --- 12. Config 儲存 / 回讀 / 缺欄位回填 ---
    $script:AppState.Config.Thresholds.PMTargetDays = 30
    Save-AppConfig
    $script:AppState.Config = $null
    Initialize-Config
    & $assert ((Get-Threshold -Name "PMTargetDays") -eq 30) "Config: 門檻修改後存檔並回讀"
    Write-AllTextUtf8 -Path $script:ConfigPath -Text '{ "DataSource": "CSV" }'
    Initialize-Config
    & $assert ((Get-Threshold -Name "SameAlarmHighCount") -eq 3) "Config: 舊檔缺欄位以預設值回填"
    & $assert (([string]$script:AppState.Config.DataSource) -eq "CSV") "Config: 保留既有設定值"

    # --- 13. SQL 介面守門 ---
    $script:AppState.Config.DataSource = "SQL"
    $sqlBlocked = $false
    try { Get-DataRows -TableName "RunSummary" | Out-Null } catch { $sqlBlocked = $true }
    & $assert $sqlBlocked "SQL 未啟用時正確拒絕"
    $script:AppState.Config.DataSource = "CSV"

    # --- 14. StepCode 平均 (以秒計 Log -> 各參數平均) ---
    $secLog = Join-Path $script:RawImportRoot "SecLog_test.csv"
    $secLines = @(
        "Timestamp,Step,Stepcode_s,Temp,Flow,Note",
        "11:00:00,1,10,100,200,x",
        "11:00:01,1,10,110,,x",
        "11:00:02,2,,50,80,x",
        "11:00:03,2,,70,120,x"
    )
    Write-AllTextUtf8 -Path $secLog -Text ($secLines -join [Environment]::NewLine)
    $sm = Export-StepCodeMeanReport -LogPath $secLog
    & $assert (Test-Path -LiteralPath $sm.OutPath) "StepMean: 輸出檔存在"
    & $assert ($sm.GroupCount -eq 2 -and $sm.GroupColumn -eq "Step") "StepMean: 依 Step 分 2 組"
    $smRows = @(Import-CsvSafe -Path $sm.OutPath)
    & $assert ((Get-FieldDouble -Object $smRows[0] -PropertyName "Temp") -eq 105) "StepMean: Step1 Temp 平均 105"
    & $assert ((Get-FieldDouble -Object $smRows[0] -PropertyName "Flow") -eq 200) "StepMean: 空白儲存格不列入平均"
    & $assert ((Get-FieldString -Object $smRows[1] -PropertyName "Stepcode_s") -eq "(blank)") "StepMean: 空白 Stepcode 標記 (blank)"
    & $assert ((Get-FieldInt -Object $smRows[1] -PropertyName "N_sec") -eq 2) "StepMean: 秒數統計正確"
    & $assert (Test-FieldMissing -Object $smRows[0] -PropertyName "Note") "StepMean: 非數值欄位輸出空白"

    # --- 15. 前 Run Log 分層比對 (N-1 vs N-2 / Golden; 資料時效) ---
    $mkLog = {
        param([string]$Name, [double]$T1, [double]$T2)
        $lines = @("Timestamp,Step,Stepcode_s,Temp,Flow")
        for ($i = 0; $i -lt 6; $i++) { $lines += ("10:00:0{0},1,11,{1},50" -f $i, $T1.ToString([System.Globalization.CultureInfo]::InvariantCulture)) }
        for ($i = 0; $i -lt 6; $i++) { $lines += ("10:01:0{0},2,22,{1},80" -f $i, $T2.ToString([System.Globalization.CultureInfo]::InvariantCulture)) }
        $lines += "10:02:00,3,33,300,10"
        $lines += "10:02:01,3,33,300,10"
        $p = Join-Path $script:RawImportRoot $Name
        Write-AllTextUtf8 -Path $p -Text ($lines -join [Environment]::NewLine)
        return $p
    }
    $logN2 = & $mkLog "SecLog_N2.csv" 100.0 200.0   # 基準 (上上 Run)
    $logN1 = & $mkLog "SecLog_N1.csv" 104.0 220.0   # 上一 Run: Step1 +4%, Step2 +10%

    $tN2 = Get-StepCodeMeanTable -LogPath $logN2
    $tN1 = Get-StepCodeMeanTable -LogPath $logN1
    & $assert ($tN2.Order.Count -eq 3) "LogCmp: 平均表分 3 Step"
    & $assert ([math]::Abs([double]$tN2.Groups["1"].Mean["Temp"] - 100) -lt 1e-9) "LogCmp: 基準 Step1 Temp 平均 100"

    $cmp = Compare-StepCodeMeanTables -BaseTable $tN2 -TestTable $tN1 -BaseName "N-2" -TestName "N-1" `
        -WarnPct 3.0 -HighPct 8.0 -MinSec 5 -TopN 10
    & $assert ($cmp.Level -eq 3) "LogCmp: Step2 +10% 超過高風險門檻 -> 高"
    & $assert ($cmp.ExceedCount -eq 2) "LogCmp: 超標數 2 (Step1 Temp / Step2 Temp)"
    & $assert (@($cmp.SkippedSteps) -contains "3") "LogCmp: 短 Step (秒數<5) 略過"
    & $assert ($cmp.ComparedCells -eq 4) "LogCmp: 比對儲存格數 4 (2 Step x 2 參數)"
    & $assert ([math]::Abs([double]$cmp.TopDeltas[0].DeltaPct - 10.0) -lt 0.001) "LogCmp: 最大差異 +10% 排行第一"

    $cmpG = Compare-StepCodeMeanTables -BaseTable $tN2 -TestTable $tN1 -BaseName "Golden" -TestName "N-1" `
        -WarnPct 5.0 -HighPct 10.0 -MinSec 5 -TopN 10
    & $assert ($cmpG.Level -eq 1) "LogCmp: Golden 門檻 (5/10%) 下僅中風險"
    & $assert ($cmpG.ExceedCount -eq 1) "LogCmp: Golden 門檻下超標數 1"

    $logShort = Join-Path $script:RawImportRoot "SecLog_struct.csv"
    $shortLines = @("Timestamp,Step,Stepcode_s,Temp,Flow")
    for ($i = 0; $i -lt 6; $i++) { $shortLines += ("10:00:0{0},1,11,100,50" -f $i) }
    Write-AllTextUtf8 -Path $logShort -Text ($shortLines -join [Environment]::NewLine)
    $tShort = Get-StepCodeMeanTable -LogPath $logShort
    $cmpS = Compare-StepCodeMeanTables -BaseTable $tN2 -TestTable $tShort -BaseName "N-2" -TestName "N-1" `
        -WarnPct 3.0 -HighPct 8.0 -MinSec 5 -TopN 10
    & $assert (@($cmpS.OnlyInBase).Count -eq 2) "LogCmp: Step 結構差異偵測 (基準獨有 2 Step)"
    & $assert ($cmpS.Level -ge 1) "LogCmp: Step 結構不一致 -> 至少中風險"

    $pre = Invoke-PreRunLogDiagnosis -PrevLogPath $logN1 -PrevPrevLogPath $logN2 -GoldenLogPath $logN2 -RunAlias "RUN_TEST"
    & $assert ($pre.Level -eq 3) "LogPreCheck: 總體風險取最大 (高)"
    & $assert (Test-Path -LiteralPath $pre.ReportPath) "LogPreCheck: 報告存檔成功"
    & $assert ($pre.Text.Contains("資料時效")) "LogPreCheck: 報告含資料時效說明"
    & $assert ($pre.Text.Contains("AI 不判定是否可開 Run")) "LogPreCheck: 報告含 AI 權限邊界聲明"

    # --- 15b. StepCode 快速彙總引擎 (v1.0.5 效能修正) ---
    & $assert ($pre.Text.Contains("比對耗時")) "LogPreCheck: 報告含比對耗時"

    # 快速引擎 (C#) 與退回模式 (純 PS 串流) 結果必須完全一致
    $tFb = Get-StepCodeMeanTableFallback -LogPath $logN1
    $consistent = ($tFb.Order.Count -eq $tN1.Order.Count) -and ($tFb.RowCount -eq $tN1.RowCount)
    if ($consistent) {
        foreach ($k in $tN1.Order) {
            if (-not $tFb.Groups.ContainsKey($k)) { $consistent = $false; break }
            if ($tFb.Groups[$k].N -ne $tN1.Groups[$k].N) { $consistent = $false; break }
            foreach ($c in $tN1.NumCols) {
                $a = $tN1.Groups[$k].Mean[$c]; $b = $tFb.Groups[$k].Mean[$c]
                if (($null -eq $a) -ne ($null -eq $b)) { $consistent = $false; break }
                if ($null -ne $a -and [math]::Abs([double]$a - [double]$b) -gt 1e-9) { $consistent = $false; break }
            }
            if (-not $consistent) { break }
        }
    }
    & $assert $consistent "StepMean: 快速引擎與退回模式結果一致"

    $logQuoted = Join-Path $script:RawImportRoot "SecLog_quoted.csv"
    $qLines = @(
        'Timestamp,Step,Stepcode_s,Temp,Note',
        '10:00:00,1,11,100,"hello, world"',
        '',
        '10:00:01,1,11,110,"say ""hi"""'
    )
    Write-AllTextUtf8 -Path $logQuoted -Text ($qLines -join [Environment]::NewLine)
    $tQ = Get-StepCodeMeanTable -LogPath $logQuoted
    & $assert ($tQ.RowCount -eq 2) "StepMean: 空白行不列入秒數"
    & $assert ([math]::Abs([double]$tQ.Groups["1"].Mean["Temp"] - 105) -lt 1e-9) "StepMean: 引號欄位 (含逗號) 正確解析"

    $logEmpty = Join-Path $script:RawImportRoot "SecLog_empty.csv"
    Write-AllTextUtf8 -Path $logEmpty -Text "Timestamp,Step,Stepcode_s,Temp"
    $emptyBlocked = $false
    try { Get-StepCodeMeanTable -LogPath $logEmpty | Out-Null } catch { $emptyBlocked = $true }
    & $assert $emptyBlocked "StepMean: 僅標頭無資料正確拒絕"

    # --- 16. 診斷主流程: 量測資料以上上 Run (N-2) 為主 (資料時效) ---
    $importFiles2 = @{
        "RunSummary.csv" = @(
            "RunID_Alias,ToolAlias,ChamberAlias,ProductFamily,RecipeFamily,RecipeVersionGroup,PreviousProductFamily,RunStartTime,RunEndTime,RunResult,OperatorShift,RecipeVersionChanged,RecipeChangeNote",
            "RUN_A,TOOL_09,CH_A,PF_VCSEL_A,RF_001,VG_12,PF_VCSEL_A,2026-07-07 08:00:00,2026-07-07 11:30:00,Normal,Day,0,",
            "RUN_B,TOOL_09,CH_A,PF_VCSEL_A,RF_001,VG_12,PF_VCSEL_A,2026-07-08 08:00:00,2026-07-08 11:30:00,Normal,Day,0,",
            "RUN_C,TOOL_09,CH_A,PF_VCSEL_A,RF_001,VG_12,PF_VCSEL_A,2026-07-09 08:00:00,,Planned,Day,0,"
        )
        "ToolStability.csv" = @(
            "RunID_Alias,ToolAlias,TempStabilityScore,PressureStabilityScore,MFCStabilityScore,RotationStabilityScore,VacuumRecoveryTimeMin,AlarmCountLastRun,CriticalAlarmCount,InterlockEventCount,CriticalAlarmWithin24h,RecentAlarmCodes",
            "RUN_B,TOOL_09,95,94,96,97,8,0,0,0,0,"
        )
        "Maintenance.csv" = @(
            "RunID_Alias,ToolAlias,DaysAfterPM,RunsAfterPM,DaysAfterPartChange,RecentSameAlarmCount7d,MTBFTrend,RecentCorrectiveMaintenance,GoldenRunVerified",
            "RUN_B,TOOL_09,5,10,30,0,Flat,No,NA"
        )
        "ProductDrift.csv" = @(
            "RunID_Alias,ToolAlias,ProductFamily,PLPeakShiftNm,ConsecutivePLShiftRuns,PLIntensityTrend,XRDPeakShiftDeg,ThicknessDeltaPct,UniformityTrend,RsDeltaPct,AOIDefectTrend,OverSpcWarning,OverSpcControl",
            "RUN_A,TOOL_09,PF_VCSEL_A,0.3,0,Flat,0.001,2.5,Flat,0.5,Flat,0,0"
        )
        "HistoryCases.csv" = @(
            "CaseID,ToolAlias,ProductFamily,RiskCategory,Keywords,Summary,Action,Outcome,CaseDate"
        )
    }
    foreach ($fn in $importFiles2.Keys) {
        Write-AllTextUtf8 -Path (Join-Path $script:ImportRoot $fn) -Text ($importFiles2[$fn] -join [Environment]::NewLine)
    }
    $diag2 = Invoke-RunDiagnosis -RunAlias "RUN_C"
    & $assert ($diag2.PrevRunAlias -eq "RUN_B") "N-2 診斷: 前一 Run = RUN_B"
    & $assert ($diag2.PrevPrevRunAlias -eq "RUN_A") "N-2 診斷: 上上 Run = RUN_A"
    & $assert ($diag2.DriftSourceRun -eq "RUN_A") "N-2 診斷: 量測資料取自上上 Run"
    & $assert (($diag2.Drift.Reasons -join "|").Contains("資料時效")) "N-2 診斷: 漂移原因含資料時效註記"
    & $assert (($diag2.Checks -join "|").Contains("前 Run Log 診斷")) "N-2 診斷: 檢查項含前 Run Log 診斷"
    $report2 = New-DiagnosisReportText -D $diag2
    & $assert ($report2.Contains("量測資料來源 Run:RUN_A")) "N-2 診斷: 報告註明量測來源 Run"
    & $assert ($report2.Contains("上上 Run:RUN_A")) "N-2 診斷: 報告標頭含上上 Run"

    # --- 17. 機台 Log 檔名目錄 (v1.0.7): 檔名帶入 機台/Run/產品 與端到端診斷 ---
    $fi = Get-RunLogFileInfo -FileName "H01B9N.PRODUCT.B9N.P02069_M06_P(NGR)_MAT06261175_17155.csv"
    & $assert ($null -ne $fi -and $fi.Tool -eq "MAT06" -and $fi.RunId -eq "261175" -and $fi.Product -eq "H01B9N") "LogCatalog: 檔名解析 機台/Run/產品"
    & $assert ($null -eq (Get-RunLogFileInfo -FileName "RunSummary.csv")) "LogCatalog: 資料表檔不誤判為 Log"

    # 建兩個同機台 Log 檔 (本 Run Step2 Temp +10%) 供目錄掃描與診斷
    $mkMatLog = {
        param([string]$Name, [double]$T2)
        $lines = @("Timestamp,Step,Stepcode_s,Temp,Flow")
        for ($i = 0; $i -lt 6; $i++) { $lines += ("10:00:0{0},1,11,100,50" -f $i) }
        for ($i = 0; $i -lt 6; $i++) { $lines += ("10:01:0{0},2,22,{1},80" -f $i, $T2.ToString([System.Globalization.CultureInfo]::InvariantCulture)) }
        $p = Join-Path $script:ImportRoot $Name
        Write-AllTextUtf8 -Path $p -Text ($lines -join [Environment]::NewLine)
    }
    & $mkMatLog "H01B9N.PRODUCT.B9N.P02069_M06_P(NGR)_MAT06261174_17154.csv" 200.0
    & $mkMatLog "H01B9N.PRODUCT.B9N.P02069_M06_P(NGR)_MAT06261175_17155.csv" 220.0
    $cat = @(Get-RunLogCatalog)
    $mat06 = @()
    foreach ($e in $cat) { if ($e.Tool -eq "MAT06") { $mat06 += $e } }
    & $assert ($mat06.Count -eq 2) "LogCatalog: 目錄掃描找到 2 個 MAT06 Run"

    $dLog = Invoke-RunDiagnosis -RunAlias "261175"
    & $assert ($dLog.ToolAlias -eq "MAT06" -and $dLog.ProductFamily -eq "H01B9N") "LogCatalog: 診斷帶入機台與產品"
    & $assert ($dLog.PrevRunAlias -eq "261174") "LogCatalog: 前一 Run 由檔名數字排序推得"
    $hasLogCmpReason = $false
    foreach ($x in $dLog.Stability.Reasons) { if ($x.StartsWith("(Log 比對)")) { $hasLogCmpReason = $true; break } }
    & $assert ($hasLogCmpReason -and $dLog.Stability.Level -ge 3) "LogCatalog: Log 比對併入穩定度 (Temp +10% -> 高)"

    # --- 18. 量測總檔 (v1.0.8): 前言跳過 / Run 彙總 / 診斷整合 ---
    $measPath = Join-Path $script:ImportRoot "Measurement_test.txt"
    $measLines = @(
        "#量測總檔表頭說明",
        "",
        "STRUCTURE,REACTOR,RUN_NO,POS_NO,",
        "-STRUCTURE=產品名,",
        "-REACTOR=EQPID,",
        "",
        "#以下為去識別後的Data",
        "",
        "STRUCTURE,REACTOR,RUN_NO,POS_NO,PF,FAIL_NO,COMMENT,R2R_W2W,ODD1,ODD2,Area_Cnt,Area_total,Haze avg,LEHI_RS,LEHI_UNIF_STD,AGA_WL_AVG,IGA_WL_AVG,LPD3+4",
        "PRO-001,Tool06,261172,A,Pass,,,,0.195,17,2,0.0542,0.052,6.50,0.40,,,30",
        "PRO-001,Tool06,261172,B,Pass,,,,0.227,31,2,0.0716,0.050,6.51,0.41,,,32",
        "PRO-001,Tool06,261172,C,Pass,,,,0.208,17,2,0.0538,0.051,6.49,0.39,,,28",
        "PRO-001,Tool06,261173,A,Pass,,,,0.403,20,10,0.272,0.055,6.70,0.42,,,45",
        "PRO-001,Tool06,261173,B,Fail,F123,,,0.377,15,5,0.193,0.056,6.71,0.41,,,44",
        "PRO-001,Tool06,261173,C,Pass,,,,0.305,17,6,0.134,0.054,6.69,0.42,,,46"
    )
    Write-AllTextUtf8 -Path $measPath -Text ($measLines -join [Environment]::NewLine)

    $measRows = @(Import-MeasurementMasterRows -Path $measPath)
    & $assert ($measRows.Count -eq 6 -and (Get-FieldString -Object $measRows[0] -PropertyName "LEHI_RS") -eq "6.50") "量測總檔: 前言跳過並解析資料列"
    $measCat = @(Get-MeasurementRunCatalog)
    $m06 = @()
    foreach ($m in $measCat) { if ([int]$m.ToolNum -eq 6) { $m06 += $m } }
    $m173 = $null
    foreach ($m in $m06) { if ($m.RunId -eq "261173") { $m173 = $m } }
    & $assert ($m06.Count -eq 2 -and $null -ne $m173 -and $m173.WaferCount -eq 3 -and $m173.FailCount -eq 1) "量測總檔: 依機台+Run 彙總 (含 Fail 統計)"

    $dMeas = Invoke-RunDiagnosis -RunAlias "261175"
    & $assert ($dMeas.DriftSourceRun -eq "261173") "量測總檔: 量測來源取最近可用 Run (資料時效)"
    $driftText = ($dMeas.Drift.Reasons -join " | ")
    & $assert ($dMeas.Drift.Level -ge 3 -and $driftText.Contains("量測最終判定")) "量測總檔: PF Fail -> 高風險"
    & $assert ($driftText.Contains("Rs 相對偏差") -and $driftText.Contains("(量測總檔)")) "量測總檔: Rs 偏差併入既有規則與摘要"

    # --- 19. 現場回饋修正 (v1.0.9) ---
    # 資料表範本: 既有檔案不覆蓋, 缺少者補建, 欄位說明檔產出
    $rsBefore = Read-AllTextUtf8 -Path (Join-Path $script:ImportRoot "RunSummary.csv")
    Initialize-ImportTableTemplates | Out-Null
    $rsAfter = Read-AllTextUtf8 -Path (Join-Path $script:ImportRoot "RunSummary.csv")
    & $assert ((Test-Path -LiteralPath (Join-Path $script:ImportRoot "MeasurementTrend.csv")) -and `
               (Test-Path -LiteralPath (Join-Path $script:DataRoot "資料表欄位說明.txt"))) "範本: 缺少的資料表與欄位說明已建立"
    & $assert ($rsAfter -eq $rsBefore) "範本: 既有資料表不覆蓋"

    # 單列 CSV 去識別化 (PS 5.1 StrictMode .Count 回歸)
    $rawOne = Join-Path $script:RawImportRoot "OneRow.csv"
    Write-AllTextUtf8 -Path $rawOne -Text ("RunID,ToolID" + [Environment]::NewLine + "R26-000001,MOCVD-Z9")
    $nOne = Convert-RawImportFile -RawPath $rawOne -OutPath (Join-Path $script:ImportRoot "OneRow.csv") -MapPath $script:DeidentMapPath
    & $assert ($nOne -eq 1) "去識別化: 單列檔案正常 (Count 回歸)"

    # 機台 Log 檔與量測總檔於批次轉換時自動略過
    Copy-Item -LiteralPath (Join-Path $script:ImportRoot "H01B9N.PRODUCT.B9N.P02069_M06_P(NGR)_MAT06261174_17154.csv") `
        -Destination (Join-Path $script:RawImportRoot "H01B9N.PRODUCT.B9N.P02069_M06_P(NGR)_MAT06261174_17154.csv") -Force
    Copy-Item -LiteralPath $measPath -Destination (Join-Path $script:RawImportRoot "Measurement_raw.csv") -Force
    $convRes = @(Convert-AllRawImports)
    $skipLog = $false; $skipMeas = $false
    foreach ($x in $convRes) {
        if ($x.Contains("機台 Log 檔, 免轉換")) { $skipLog = $true }
        if ($x.Contains("量測總檔")) { $skipMeas = $true }
    }
    & $assert ($skipLog -and $skipMeas) "去識別化: 機台 Log 與量測總檔自動略過"

    # Log 比對近零基準門檻 (LogCompareMinBaseAbs)
    $tinyBase = @{ Order = @("1"); Groups = @{ "1" = @{ Key = "1"; Stepcode = "11"; N = 10; Mean = @{ P = 0.02 } } }
                   NumCols = @("P"); RowCount = 10; SourceName = "b" }
    $tinyTest = @{ Order = @("1"); Groups = @{ "1" = @{ Key = "1"; Stepcode = "11"; N = 10; Mean = @{ P = -1.5 } } }
                   NumCols = @("P"); RowCount = 10; SourceName = "t" }
    $cmpTiny = Compare-StepCodeMeanTables -BaseTable $tinyBase -TestTable $tinyTest -BaseName "b" -TestName "t" `
        -WarnPct 3.0 -HighPct 8.0 -MinSec 5 -TopN 5 -MinBaseAbs 0.05
    & $assert ($cmpTiny.ComparedCells -eq 0 -and $cmpTiny.Level -eq 0) "LogCmp: 近零基準參數依門檻略過"

    # --- 20. 資料表自動補列 (v1.0.10) ---
    $upd1 = Update-DerivedTables
    $rsText = Read-AllTextUtf8 -Path (Join-Path $script:ImportRoot "RunSummary.csv")
    & $assert ($upd1.RunSummaryAdded -ge 2 -and $rsText.Contains("261175,MAT06,,H01B9N")) "自動補列: RunSummary 由 Log 檔名產生"
    $mtText = Read-AllTextUtf8 -Path (Join-Path $script:ImportRoot "MeasurementTrend.csv")
    & $assert ($upd1.MeasurementTrendAdded -ge 2 -and $mtText.Contains("261173,MAT06")) "自動補列: MeasurementTrend 由量測總檔產生"
    $upd2 = Update-DerivedTables
    & $assert ($upd2.RunSummaryAdded -eq 0 -and $upd2.MeasurementTrendAdded -eq 0) "自動補列: 重複執行不重複增列"

    # Run 已建入 RunSummary 後, 診斷仍保有 Log 比對與量測總檔整合 (解耦回歸)
    $dTbl = Invoke-RunDiagnosis -RunAlias "261175"
    $tblStabText = ($dTbl.Stability.Reasons -join " | ")
    $tblDriftText = ($dTbl.Drift.Reasons -join " | ")
    $tblOverallText = ($dTbl.Overall.Reasons -join " | ")
    & $assert ($tblStabText.Contains("(Log 比對)") -and $tblDriftText.Contains("(量測總檔)") -and `
               -not $tblOverallText.Contains("由 Log 檔名解析建立")) "自動補列: 建表後診斷仍含 Log 比對與量測整合"

    # --- 21. RunID_Alias 鍵格式容錯 (v1.0.11): 機台+Run 碼填法 ---
    $rowFlex = New-TestRow @{ RunID_Alias = "MAT06261176"; ToolAlias = "MAT06"; TempStabilityScore = "90" }
    & $assert ($null -ne (Find-RowByRun -Rows @($rowFlex) -RunAlias "261176" -ToolAlias "MAT06")) "鍵容錯: 機台+Run 碼可查得"
    & $assert ($null -eq (Find-RowByRun -Rows @($rowFlex) -RunAlias "261176")) "鍵容錯: 未指定機台時維持嚴格比對"

    # e2e: 重現現場填法 — ToolStability / Maintenance 以 MAT06261174 為鍵
    Write-AllTextUtf8 -Path (Join-Path $script:ImportRoot "ToolStability.csv") -Text (@(
        "RunID_Alias,ToolAlias,TempStabilityScore,PressureStabilityScore,MFCStabilityScore,RotationStabilityScore,VacuumRecoveryTimeMin,AlarmCountLastRun,CriticalAlarmCount,InterlockEventCount,CriticalAlarmWithin24h,RecentAlarmCodes",
        "MAT06261174,MAT06,90,90,90,80,120,0,0,0,0,"
    ) -join [Environment]::NewLine)
    Write-AllTextUtf8 -Path (Join-Path $script:ImportRoot "Maintenance.csv") -Text (@(
        "RunID_Alias,ToolAlias,DaysAfterPM,RunsAfterPM,DaysAfterPartChange,RecentSameAlarmCount7d,MTBFTrend,RecentCorrectiveMaintenance,GoldenRunVerified",
        "MAT06261174,MAT06,3,10,3,0,Flat,No,NA"
    ) -join [Environment]::NewLine)
    $dFlex = Invoke-RunDiagnosis -RunAlias "261175"
    $flexStab = ($dFlex.Stability.Reasons -join " | ")
    $flexMaint = ($dFlex.Maintenance.Reasons -join " | ")
    & $assert ((-not $flexStab.Contains("無前一 Run 穩定度資料")) -and $flexStab.Contains("轉速穩定分數")) "鍵容錯: 穩定度資料以機台+Run 碼查得"
    & $assert (-not $flexMaint.Contains("無 PM / 維修資料")) "鍵容錯: 維修資料以機台+Run 碼查得"

    # --- 22. 趨勢圖資料: 無 PL 資料時改用 Rs (v1.0.12) ---
    $trend = Get-TrendChartData -ToolAlias "MAT06"
    & $assert ((-not $trend.UsePl) -and $trend.UseRs) "趨勢圖: 無 PL 資料時改用 Rs"
    $tItems = @($trend.Items)
    $lastIt = $tItems[$tItems.Count - 1]
    & $assert ($tItems.Count -ge 2 -and [string]$lastIt.RunId -eq "261173" -and `
               $null -ne $lastIt.Rs -and [math]::Abs([double]$lastIt.Rs - 6.7) -lt 0.01) "趨勢圖: Rs 數列由量測總檔合併 (舊->新排序)"

    # --- 23. 現場回饋 3 項 (v1.0.13) ---
    # (a) 參數排除: Dop1.* 不列入比對且留痕
    & $assert (@(Get-LogCompareExcludePatterns) -contains "Dop1.*") "排除參數: 預設含 Dop1.*"
    $excBase = @{ Order = @("1"); Groups = @{ "1" = @{ Key = "1"; Stepcode = "11"; N = 10; Mean = @{ "Dop1.dp_SP" = 0.01; Temp = 100.0 } } }
                  NumCols = @("Dop1.dp_SP", "Temp"); RowCount = 10; SourceName = "b" }
    $excTest = @{ Order = @("1"); Groups = @{ "1" = @{ Key = "1"; Stepcode = "11"; N = 10; Mean = @{ "Dop1.dp_SP" = -1.5; Temp = 110.0 } } }
                  NumCols = @("Dop1.dp_SP", "Temp"); RowCount = 10; SourceName = "t" }
    $cmpExc = Compare-StepCodeMeanTables -BaseTable $excBase -TestTable $excTest -BaseName "b" -TestName "t" `
        -WarnPct 3.0 -HighPct 8.0 -MinSec 5 -TopN 5 -ExcludeParams @("Dop1.*")
    $excText = ($cmpExc.Reasons -join " | ")
    & $assert ($cmpExc.ComparedCells -eq 1 -and $cmpExc.ExceedCount -eq 1 -and `
               $excText.Contains("依設定排除 1 個參數")) "排除參數: Dop1 不列入且報告留痕"

    # (b) 主診斷報告含 Log 比對差異排行
    $dRank = Invoke-RunDiagnosis -RunAlias "261175"
    & $assert ($dRank.ContainsKey("LogCompare") -and $null -ne $dRank.LogCompare) "報告排行: 診斷結果含 Log 比對明細"
    $rankReport = New-DiagnosisReportText -D $dRank
    & $assert ($rankReport.Contains("Log 比對差異排行") -and $rankReport.Contains("| Temp |")) "報告排行: 主報告含差異排行內容"

    # (c) 趨勢圖聯集: 量測總檔新 Run 未入 MeasurementTrend 表時圖仍涵蓋
    $measLines2 = @(
        "STRUCTURE,REACTOR,RUN_NO,POS_NO,PF,FAIL_NO,LEHI_RS",
        "PRO-001,Tool06,261176,A,Pass,,6.55",
        "PRO-001,Tool06,261176,B,Pass,,6.57"
    )
    Write-AllTextUtf8 -Path (Join-Path $script:ImportRoot "Measurement_new.csv") -Text ($measLines2 -join [Environment]::NewLine)
    $trend2 = Get-TrendChartData -ToolAlias "MAT06"
    $t2Items = @($trend2.Items)
    $t2Last = $t2Items[$t2Items.Count - 1]
    & $assert ([string]$t2Last.RunId -eq "261176" -and $null -ne $t2Last.Rs) "趨勢圖: 表未更新時仍聯集量測總檔新 Run"

    # --- 24. StepLabel 分組 (v1.0.14): StepLabel 對應真實 Layer ---
    $logLabel = Join-Path $script:RawImportRoot "SecLog_label.csv"
    $lblLines = @("Timestamp,Step,StepLabel,Stepcode_s,Temp,Flow")
    for ($i = 0; $i -lt 6; $i++) { $lblLines += ("10:00:0{0},1,Buffer,11,100,50" -f $i) }
    for ($i = 0; $i -lt 6; $i++) { $lblLines += ("10:01:0{0},2,Buffer,12,110,50" -f $i) }
    for ($i = 0; $i -lt 6; $i++) { $lblLines += ("10:02:0{0},3,QW,21,200,80" -f $i) }
    Write-AllTextUtf8 -Path $logLabel -Text ($lblLines -join [Environment]::NewLine)
    $tLbl = Get-StepCodeMeanTable -LogPath $logLabel
    & $assert ($tLbl.GroupColumn -eq "StepLabel" -and $tLbl.Order.Count -eq 2 -and $tLbl.Groups["Buffer"].N -eq 12) "StepLabel: 依 StepLabel 分組 (跨 Step 合併同 Layer)"
    & $assert ([math]::Abs([double]$tLbl.Groups["Buffer"].Mean["Temp"] - 105) -lt 1e-9 -and `
               (@($tLbl.NumCols) -notcontains "StepLabel") -and (@($tLbl.NumCols) -notcontains "Step")) "StepLabel: Layer 平均正確且標籤欄不列入數值比對"
    $tLblFb = Get-StepCodeMeanTableFallback -LogPath $logLabel
    & $assert ($tLblFb.GroupColumn -eq "StepLabel" -and [math]::Abs([double]$tLblFb.Groups["QW"].Mean["Temp"] - 200) -lt 1e-9) "StepLabel: 退回模式同樣依 StepLabel 分組"

    # --- 25. 資料表管理 (v1.0.15): 標頭定義集中共用 ---
    $hdrs = Get-ImportTableHeaders
    & $assert ($hdrs.Count -eq 6 -and $hdrs.ContainsKey("RunSummary.csv") -and `
               ([string]$hdrs["HistoryCases.csv"]).StartsWith("CaseID,")) "資料表管理: 六張表標頭定義集中"
    $mtHdrLine = ((Read-AllTextUtf8 -Path (Join-Path $script:ImportRoot "MeasurementTrend.csv")) -split "`r?`n")[0].Trim()
    & $assert ($mtHdrLine -eq [string]$hdrs["MeasurementTrend.csv"]) "資料表管理: 範本檔標頭與定義一致"

    # --- 收尾 ---
    Write-Host ""
    Write-Host ("SelfTest 結果: PASS={0} FAIL={1}" -f $t.Pass, $t.Fail)
    try { Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    if ($t.Fail -gt 0) { return 1 }
    return 0
}

# ============================================================
# UI 工廠
# ============================================================
function New-Form {
    param([string]$Text, [int]$Width, [int]$Height)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Text; $f.Width = $Width; $f.Height = $Height
    $f.StartPosition = "CenterScreen"
    $f.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 9)
    return $f
}
function New-UiLabel {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 100, [int]$Height = 24)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Left = $X; $l.Top = $Y
    $l.Width = $Width; $l.Height = $Height; $l.TextAlign = "MiddleLeft"
    return $l
}
function New-UiTextBox {
    param([int]$X, [int]$Y, [int]$Width = 160, [int]$Height = 24, [string]$Text = "")
    $t = New-Object System.Windows.Forms.TextBox
    $t.Left = $X; $t.Top = $Y; $t.Width = $Width; $t.Height = $Height; $t.Text = $Text
    return $t
}
function New-UiButton {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 110, [int]$Height = 30)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Left = $X; $b.Top = $Y; $b.Width = $Width; $b.Height = $Height
    return $b
}
function New-UiComboBox {
    param([int]$X, [int]$Y, [int]$Width = 180)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.Left = $X; $c.Top = $Y; $c.Width = $Width
    $c.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    return $c
}

function Get-RiskColor {
    param([int]$Level)
    switch ($Level) {
        0 { return [System.Drawing.Color]::FromArgb(46, 139, 87) }    # 低: 綠
        1 { return [System.Drawing.Color]::FromArgb(218, 165, 32) }   # 中: 金黃
        2 { return [System.Drawing.Color]::FromArgb(230, 126, 34) }   # 中高: 橘
        default { return [System.Drawing.Color]::FromArgb(192, 57, 43) }  # 高: 紅
    }
}

function New-RiskPanel {
    # 燈號面板: 回傳 hashtable @{ Panel; Title; Value; SetLevel(scriptblock) }
    param([string]$Title, [int]$X, [int]$Y, [int]$Width = 200, [int]$Height = 72)
    $p = New-Object System.Windows.Forms.Panel
    $p.Left = $X; $p.Top = $Y; $p.Width = $Width; $p.Height = $Height
    $p.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $p.BackColor = [System.Drawing.Color]::Gainsboro

    $lt = New-UiLabel -Text $Title -X 8 -Y 6 -Width ($Width - 16) -Height 22
    $lt.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $lt.ForeColor = [System.Drawing.Color]::White
    $lt.BackColor = [System.Drawing.Color]::Transparent
    $p.Controls.Add($lt)

    $lv = New-UiLabel -Text "--" -X 8 -Y 30 -Width ($Width - 16) -Height 34
    $lv.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 16, [System.Drawing.FontStyle]::Bold)
    $lv.ForeColor = [System.Drawing.Color]::White
    $lv.BackColor = [System.Drawing.Color]::Transparent
    $p.Controls.Add($lv)

    return @{ Panel = $p; TitleLabel = $lt; ValueLabel = $lv }
}

function Set-RiskPanelLevel {
    param([hashtable]$RiskPanel, [object]$Level)
    if ($null -eq $Level) {
        $RiskPanel.Panel.BackColor = [System.Drawing.Color]::Gainsboro
        $RiskPanel.ValueLabel.Text = "--"
        return
    }
    $lv = [int]$Level
    $RiskPanel.Panel.BackColor = (Get-RiskColor -Level $lv)
    $RiskPanel.ValueLabel.Text = (Get-RiskName $lv)
}

# ============================================================
# 趨勢圖 (MeasurementTrend.csv; 無圖表元件時降級為文字)
# ============================================================
function New-TrendChart {
    param([string]$ToolAlias, [int]$X, [int]$Y, [int]$Width, [int]$Height)
    if (-not $script:ChartingAvailable) { return $null }

    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.Left = $X; $chart.Top = $Y; $chart.Width = $Width; $chart.Height = $Height
    $chart.BackColor = [System.Drawing.Color]::White
    $chart.Anchor = "Top,Bottom,Left,Right"

    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea("MainArea")
    $area.AxisX.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
    $area.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
    $area.AxisX.LabelStyle.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 8)
    $area.AxisX.LabelStyle.Angle = -45
    $area.AxisY.Title = "相對偏移 (%/nm)"
    $area.AxisY2.Title = "Alarm 數"
    $area.AxisY2.Enabled = [System.Windows.Forms.DataVisualization.Charting.AxisEnabled]::True
    [void]$chart.ChartAreas.Add($area)

    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $legend.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 8.5)
    [void]$chart.Legends.Add($legend)
    return $chart
}

function Get-TrendChartData {
    <#
      趨勢圖資料準備 (v1.0.12; 與 UI 分離, 供 SelfTest 驗證):
      - 以 MeasurementTrend 表為主 (Alarm / PL 偏移 / 厚度偏差)
      - 由量測總檔目錄合併每 Run 的 Rs 平均 (LEHI_RS)
      - 該機台完全無 PL 資料時 UseRs=true — 主線改畫 Rs
        (現場需求: 部分產品不量 PL, 原 PL 線恆為空)
      - 表無資料時直接以量測總檔的 Run 清單作圖
      回傳 @{ Items (舊->新; RunId/RunDate/Alarm/Pl/Thk/Rs, 缺值為 $null);
              UsePl; UseRs }
    #>
    param([string]$ToolAlias)
    $maxN = [int](Get-Threshold -Name "TrendRunCount" -DefaultValue 10)

    # 量測總檔: Rs 平均與 Run 清單 (機台以 MAT 尾碼對應)
    $rsByRun = @{}
    $catRuns = @()
    try {
        foreach ($m in @(Get-MeasurementRunCatalog)) {
            if (("MAT{0:D2}" -f [int]$m.ToolNum) -ne $ToolAlias) { continue }
            $catRuns += $m
            if ($null -ne $m.Mean["LEHI_RS"]) { $rsByRun[[string]$m.RunId] = [double]$m.Mean["LEHI_RS"] }
        }
    } catch { }

    $rows = @()
    try { $rows = @(Get-DataRows -TableName "MeasurementTrend") } catch { $rows = @() }
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) {
        if ((Get-FieldString -Object $r -PropertyName "ToolAlias") -ne $ToolAlias) { continue }
        $rid = Get-FieldString -Object $r -PropertyName "RunID_Alias"
        if ([string]::IsNullOrEmpty($rid)) { continue }
        $alarm = $null; $pl = $null; $thk = $null
        if (-not (Test-FieldMissing -Object $r -PropertyName "AlarmCount"))        { $alarm = Get-FieldInt    -Object $r -PropertyName "AlarmCount" }
        if (-not (Test-FieldMissing -Object $r -PropertyName "PLPeakShiftNm"))     { $pl    = Get-FieldDouble -Object $r -PropertyName "PLPeakShiftNm" }
        if (-not (Test-FieldMissing -Object $r -PropertyName "ThicknessDeltaPct")) { $thk   = Get-FieldDouble -Object $r -PropertyName "ThicknessDeltaPct" }
        $rs = $null
        if ($rsByRun.ContainsKey($rid)) { $rs = [double]$rsByRun[$rid] }
        [void]$items.Add(@{ RunId = $rid; RunDate = (Get-FieldString -Object $r -PropertyName "RunDate")
                            Alarm = $alarm; Pl = $pl; Thk = $thk; Rs = $rs })
    }
    # v1.0.13: 與量測總檔目錄「聯集」— MeasurementTrend 表未更新
    # (如檔案被 Excel 開啟鎖住致補列失敗) 時, 圖仍即時涵蓋全部量測 Run
    $haveIds = @{}
    foreach ($it in $items) { $haveIds[[string]$it.RunId] = $true }
    foreach ($m in $catRuns) {
        if ($haveIds.ContainsKey([string]$m.RunId)) { continue }
        $rs = $null
        if ($rsByRun.ContainsKey([string]$m.RunId)) { $rs = [double]$rsByRun[[string]$m.RunId] }
        [void]$items.Add(@{ RunId = [string]$m.RunId; RunDate = [string]$m.GDate
                            Alarm = $null; Pl = $null; Thk = $null; Rs = $rs })
        $haveIds[[string]$m.RunId] = $true
    }

    # 舊 -> 新排序: Run 碼皆為數字時以數字為主 (不受日期格式差異影響);
    # 否則以 RunDate 為主、Run 碼為輔
    $allNumeric = ($items.Count -gt 0)
    foreach ($it in $items) {
        $v = 0L
        if (-not [long]::TryParse([string]$it.RunId, [ref]$v)) { $allNumeric = $false; break }
    }
    if ($allNumeric) {
        $sorted = @($items | Sort-Object -Property @{ Expression = { [long]$_.RunId } })
    } else {
        $sorted = @($items | Sort-Object -Property `
            @{ Expression = { [string]$_.RunDate } }, `
            @{ Expression = { $v = 0L; [void][long]::TryParse([string]$_.RunId, [ref]$v); $v } })
    }
    if ($sorted.Count -gt $maxN) { $sorted = @($sorted[($sorted.Count - $maxN)..($sorted.Count - 1)]) }

    $usePl = $false
    $anyRs = $false
    foreach ($it in $sorted) {
        if ($null -ne $it.Pl) { $usePl = $true }
        if ($null -ne $it.Rs) { $anyRs = $true }
    }
    return @{ Items = $sorted; UsePl = $usePl; UseRs = ((-not $usePl) -and $anyRs) }
}

function Update-TrendChart {
    param([object]$Chart, [string]$ToolAlias)
    if ($null -eq $Chart) { return }
    $Chart.Series.Clear()
    $data = Get-TrendChartData -ToolAlias $ToolAlias
    $items = @($data.Items)
    if ($items.Count -eq 0) { return }

    $sAlarm = New-Object System.Windows.Forms.DataVisualization.Charting.Series("Alarm 數")
    $sAlarm.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
    $sAlarm.YAxisType = [System.Windows.Forms.DataVisualization.Charting.AxisType]::Secondary
    $sAlarm.Color = [System.Drawing.Color]::FromArgb(120, 149, 165, 166)

    # v1.0.12: 無 PL 資料時主線改畫 Rs (由量測總檔), Y 軸標題同步切換
    $mainName = "PL peak 偏移 (nm)"
    $axisTitle = "相對偏移 (%/nm)"
    if ($data.UseRs) {
        $mainName = "Rs 平均 (LEHI_RS)"
        $axisTitle = "Rs (LEHI_RS 平均)"
    }
    try { $Chart.ChartAreas[0].AxisY.Title = $axisTitle } catch { }
    $sMain = New-Object System.Windows.Forms.DataVisualization.Charting.Series($mainName)
    $sMain.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $sMain.BorderWidth = 2
    $sMain.MarkerStyle = "Circle"; $sMain.MarkerSize = 6
    $sMain.Color = [System.Drawing.Color]::FromArgb(41, 128, 185)

    $sThk = New-Object System.Windows.Forms.DataVisualization.Charting.Series("厚度偏差 (%)")
    $sThk.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $sThk.BorderWidth = 2
    $sThk.MarkerStyle = "Square"; $sThk.MarkerSize = 6
    $sThk.Color = [System.Drawing.Color]::FromArgb(192, 57, 43)

    $anyAlarm = $false
    $anyThk = $false
    foreach ($it in $items) {
        $x = [string]$it.RunId
        $i1 = $sAlarm.Points.AddXY($x, $(if ($null -ne $it.Alarm) { [int]$it.Alarm } else { 0 }))
        if ($null -eq $it.Alarm) { $sAlarm.Points[$i1].IsEmpty = $true } else { $anyAlarm = $true }
        $mv = $null
        if ($data.UseRs) { $mv = $it.Rs } else { $mv = $it.Pl }
        $i2 = $sMain.Points.AddXY($x, $(if ($null -ne $mv) { [double]$mv } else { 0 }))
        if ($null -eq $mv) { $sMain.Points[$i2].IsEmpty = $true }
        $i3 = $sThk.Points.AddXY($x, $(if ($null -ne $it.Thk) { [double]$it.Thk } else { 0 }))
        if ($null -eq $it.Thk) { $sThk.Points[$i3].IsEmpty = $true } else { $anyThk = $true }
    }
    if ($anyAlarm) { [void]$Chart.Series.Add($sAlarm) }
    [void]$Chart.Series.Add($sMain)
    if ($anyThk) { [void]$Chart.Series.Add($sThk) }
}

# ============================================================
# Tab 1: Run 前診斷儀表板
# ============================================================
function Show-LogCompareDialog {
    <#
      前 Run Log 分層診斷對話框 (v1.0.4):
      N-1 秒級 Log (必選) vs N-2 秒級 Log / Golden 基準 (選填)。
      結果顯示於視窗並自動存檔至 Reports\LogPreCheck_*.txt。
    #>
    param([hashtable]$St)

    $form = New-Form -Text "前 Run Log 診斷 (N-1 vs N-2 / Golden) — 資料時效: Run 前僅有 N-1 機台 Log" -Width 800 -Height 620
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    $rows = @(
        @{ Key = "Prev";   Label = "N-1 (上一 Run) 秒級 Log [必選]:";                     Y = 12;  UseGolden = $false },
        @{ Key = "PP";     Label = "N-2 (上上 Run) 秒級 Log [選填, 同 recipe 比對基準]:"; Y = 70;  UseGolden = $false },
        @{ Key = "Golden"; Label = "Golden Run 基準 Log [選填, 建議存放 Data\Golden]:";    Y = 128; UseGolden = $true }
    )
    $boxes = @{}
    foreach ($rowDef in $rows) {
        $form.Controls.Add((New-UiLabel -Text $rowDef.Label -X 10 -Y $rowDef.Y -Width 620))
        $tb = New-UiTextBox -X 10 -Y ($rowDef.Y + 24) -Width 650
        $form.Controls.Add($tb)
        $btn = New-UiButton -Text "瀏覽..." -X 670 -Y ($rowDef.Y + 22) -Width 100
        $form.Controls.Add($btn)
        $boxes[$rowDef.Key] = $tb
        $useGolden = [bool]$rowDef.UseGolden
        $btn.Add_Click({
            try {
                # Pitfall 1: 閉包內不讀 $script: 變數, 路徑一律經函式取得
                $dlg = New-Object System.Windows.Forms.OpenFileDialog
                $dlg.Title = "選擇秒級 Run Log CSV"
                $dlg.Filter = "CSV 檔 (*.csv)|*.csv"
                if ($useGolden) { $dlg.InitialDirectory = (Get-GoldenRootPath) }
                else { $dlg.InitialDirectory = (Get-RawImportRootPath) }
                if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $tb.Text = $dlg.FileName
                }
            } catch {
                Write-ErrorLog ("LogCompare browse: " + $_.Exception.Message)
                Show-ErrorMessage $_.Exception.Message
            }
        }.GetNewClosure())
    }

    $btnRun = New-UiButton -Text "執行比對" -X 10 -Y 186 -Width 120
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($btnRun)
    $btnOpenRpt = New-UiButton -Text "開啟報告資料夾" -X 140 -Y 186 -Width 140
    $form.Controls.Add($btnOpenRpt)
    $lblStatus = New-UiLabel -Text "" -X 290 -Y 190 -Width 480 -Height 22
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblStatus)

    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Left = 10; $rtb.Top = 220; $rtb.Width = 760; $rtb.Height = 350
    $rtb.ReadOnly = $true
    $rtb.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 9)
    $rtb.Text = "資料時效說明: Run 前僅能即時取得上一 Run (N-1) 的機台秒級 Log;" + [Environment]::NewLine + `
        "PL / XRD / Thickness / Rs / AOI 量測最快僅到上上 Run (N-2)。" + [Environment]::NewLine + `
        "本工具以 N-1 Log 對 N-2 Log / Golden 基準做 StepCode 平均差異比對, 提前預警機台行為變化;" + [Environment]::NewLine + `
        "僅做整理 / 比對 / 提醒, 不判定是否可開 Run。"
    $form.Controls.Add($rtb)

    $btnRun.Add_Click({
        try {
            $p1 = $boxes["Prev"].Text.Trim()
            if ([string]::IsNullOrEmpty($p1)) { Show-Warn "請先選擇 N-1 (上一 Run) 的秒級 Log 檔。"; return }
            $p2 = $boxes["PP"].Text.Trim()
            $pg = $boxes["Golden"].Text.Trim()
            if ([string]::IsNullOrEmpty($p2) -and [string]::IsNullOrEmpty($pg)) {
                Show-Warn "請至少提供 N-2 Log 或 Golden 基準 Log 其中一項作為比對基準。"; return
            }
            # v1.0.5: 執行中顯示等待游標與狀態 (大檔比對數秒內完成)
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnRun.Enabled = $false
            $lblStatus.Text = "比對中, 請稍候..."
            $lblStatus.Refresh()
            try {
                $res = Invoke-PreRunLogDiagnosis -PrevLogPath $p1 -PrevPrevLogPath $p2 -GoldenLogPath $pg
            } finally {
                $btnRun.Enabled = $true
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
            $rtb.Text = $res.Text
            $lblStatus.Text = ("風險等級: {0} / 耗時 {1} 秒 / 報告: {2}" -f `
                (Get-RiskName $res.Level), $res.ElapsedSeconds, (Split-Path -Leaf $res.ReportPath))
        } catch {
            Write-ErrorLog ("btnLogCompareRun: " + $_.Exception.Message)
            Show-ErrorMessage ("前 Run Log 診斷失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())

    $btnOpenRpt.Add_Click({
        try { Invoke-Item -LiteralPath (Get-ReportRootPath) } catch { Show-ErrorMessage $_.Exception.Message }
    }.GetNewClosure())

    [void]$form.ShowDialog()
}

function Build-DiagnosisTab {
    param([System.Windows.Forms.TabPage]$Tab, [hashtable]$St)

    $Tab.Controls.Add((New-UiLabel -Text "機台:" -X 10 -Y 14 -Width 45))
    $cbTool = New-UiComboBox -X 58 -Y 12 -Width 130
    $Tab.Controls.Add($cbTool)
    $Tab.Controls.Add((New-UiLabel -Text "Run:" -X 200 -Y 14 -Width 40))
    $cbRun = New-UiComboBox -X 242 -Y 12 -Width 190
    $Tab.Controls.Add($cbRun)
    $btnDiag = New-UiButton -Text "執行診斷" -X 450 -Y 10 -Width 100
    $btnDiag.BackColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
    $btnDiag.ForeColor = [System.Drawing.Color]::White
    $Tab.Controls.Add($btnDiag)
    $btnReport = New-UiButton -Text "產出報告" -X 560 -Y 10 -Width 100
    $btnReport.Enabled = $false
    $Tab.Controls.Add($btnReport)
    $btnReload = New-UiButton -Text "重載資料" -X 670 -Y 10 -Width 100
    $Tab.Controls.Add($btnReload)
    $btnLogCmp = New-UiButton -Text "前 Run Log 診斷 (N-1/N-2)" -X 780 -Y 10 -Width 200
    $Tab.Controls.Add($btnLogCmp)

    $lblGen = New-UiLabel -Text "" -X 220 -Y 392 -Width 500 -Height 22
    $lblGen.ForeColor = [System.Drawing.Color]::DimGray
    $Tab.Controls.Add($lblGen)

    # 燈號列
    $panels = @{
        Overall     = New-RiskPanel -Title "總體風險"     -X 10  -Y 48 -Width 250
        Stability   = New-RiskPanel -Title "機台穩定度"   -X 270 -Y 48 -Width 250
        Maintenance = New-RiskPanel -Title "維修風險"     -X 530 -Y 48 -Width 250
        Drift       = New-RiskPanel -Title "產品特性飄移" -X 790 -Y 48 -Width 250
    }
    foreach ($k in $panels.Keys) { $Tab.Controls.Add($panels[$k].Panel) }

    # 左: 診斷內容; 右: 趨勢圖
    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Left = 10; $rtb.Top = 130; $rtb.Width = 520; $rtb.Height = 250
    $rtb.ReadOnly = $true
    $rtb.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 9.5)
    $rtb.Text = "請選擇機台與 Run 後按「執行診斷」。" + [Environment]::NewLine + [Environment]::NewLine + `
        "本系統僅做資料彙整與風險提醒, 不判定是否可開 Run, 最終決策由製程 / 設備工程師覆判。"
    $rtb.Anchor = "Top,Bottom,Left"
    $Tab.Controls.Add($rtb)

    $chart = New-TrendChart -ToolAlias "" -X 540 -Y 130 -Width 500 -Height 250
    if ($null -ne $chart) {
        $chart.Anchor = "Top,Bottom,Left,Right"
        $Tab.Controls.Add($chart)
    } else {
        $lblNoChart = New-UiLabel -Text "(此環境無 DataVisualization 圖表元件, 趨勢圖停用)" -X 540 -Y 130 -Width 480 -Height 24
        $lblNoChart.ForeColor = [System.Drawing.Color]::Gray
        $Tab.Controls.Add($lblNoChart)
    }

    # 相似案例
    $Tab.Controls.Add((New-UiLabel -Text "相似歷史案例:" -X 10 -Y 388 -Width 200))
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 10; $grid.Top = 412; $grid.Width = 1030; $grid.Height = 130
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.Anchor = "Bottom,Left,Right"
    foreach ($col in @("案例", "類別", "摘要", "處置", "結果", "日期")) {
        [void]$grid.Columns.Add($col, $col)
    }
    $grid.Columns[2].FillWeight = 220
    $Tab.Controls.Add($grid)

    $fillRunCombo = {
        try {
            $cbRun.Items.Clear()
            $tool = [string]$cbTool.SelectedItem
            if ([string]::IsNullOrEmpty($tool)) { return }
            $rows = @(Get-DataRows -TableName "RunSummary")
            $toolRuns = @()
            foreach ($r in $rows) {
                if ((Get-FieldString -Object $r -PropertyName "ToolAlias") -eq $tool) { $toolRuns += $r }
            }
            $sorted = @($toolRuns | Sort-Object -Property @{ Expression = { Get-FieldString -Object $_ -PropertyName "RunStartTime" }; Descending = $true })
            # v1.0.10: 顯示 "Run碼 | 產品"; 表格與 Log 檔名目錄同 Run 不重複列出
            $seenRuns = @{}
            foreach ($r in $sorted) {
                $rid = Get-FieldString -Object $r -PropertyName "RunID_Alias"
                if ([string]::IsNullOrEmpty($rid) -or $seenRuns.ContainsKey($rid)) { continue }
                $seenRuns[$rid] = $true
                $txt = $rid
                $pfv = Get-FieldString -Object $r -PropertyName "ProductFamily"
                if (-not [string]::IsNullOrEmpty($pfv)) { $txt = $rid + " | " + $pfv }
                [void]$cbRun.Items.Add($txt)
            }
            # v1.0.7: 併入 Log 檔名目錄的 Run (新到舊)
            $cat = @()
            if ($null -ne $St.LogCatalog) { $cat = @($St.LogCatalog) }
            $logRuns = @()
            foreach ($e in $cat) { if ([string]$e.Tool -eq $tool) { $logRuns += $e } }
            $logRuns = @($logRuns | Sort-Object -Property @{ Expression = { [long]$_.RunId }; Descending = $true })
            foreach ($e in $logRuns) {
                $rid = [string]$e.RunId
                if ($seenRuns.ContainsKey($rid)) { continue }
                $seenRuns[$rid] = $true
                $txt = $rid
                if (-not [string]::IsNullOrEmpty([string]$e.Product)) { $txt = $rid + " | " + [string]$e.Product }
                [void]$cbRun.Items.Add($txt)
            }
            if ($cbRun.Items.Count -gt 0) { $cbRun.SelectedIndex = 0 }
        } catch {
            Write-ErrorLog ("fillRunCombo failed: " + $_.Exception.Message)
        }
    }.GetNewClosure()

    $reloadTools = {
        try {
            # v1.0.10: 由 Log 檔名目錄 / 量測總檔自動補列 RunSummary / MeasurementTrend
            try { Update-DerivedTables | Out-Null } catch { Write-ErrorLog ("Update-DerivedTables: " + $_.Exception.Message) }
            $cbTool.Items.Clear()
            $rows = @(Get-DataRows -TableName "RunSummary")
            $seen = @{}
            foreach ($r in $rows) {
                $tl = Get-FieldString -Object $r -PropertyName "ToolAlias"
                if (-not [string]::IsNullOrEmpty($tl) -and -not $seen.ContainsKey($tl)) {
                    $seen[$tl] = $true
                    [void]$cbTool.Items.Add($tl)
                }
            }
            # v1.0.7: 併入 Log 檔名目錄的機台 (Input 僅放 Log 檔時選單不再空白)
            $St.LogCatalog = @(Get-RunLogCatalog)
            foreach ($e in $St.LogCatalog) {
                $tl = [string]$e.Tool
                if (-not [string]::IsNullOrEmpty($tl) -and -not $seen.ContainsKey($tl)) {
                    $seen[$tl] = $true
                    [void]$cbTool.Items.Add($tl)
                }
            }
            if ($cbTool.Items.Count -gt 0) { $cbTool.SelectedIndex = 0 }
            & $fillRunCombo
        } catch {
            Write-ErrorLog ("reloadTools failed: " + $_.Exception.Message)
            Show-ErrorMessage ("重載資料失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure()

    $showDiagnosis = {
        param([hashtable]$D)
        Set-RiskPanelLevel -RiskPanel $panels.Overall     -Level $D.Overall.Level
        Set-RiskPanelLevel -RiskPanel $panels.Stability   -Level $D.Stability.Level
        Set-RiskPanelLevel -RiskPanel $panels.Maintenance -Level $D.Maintenance.Level
        Set-RiskPanelLevel -RiskPanel $panels.Drift       -Level $D.Drift.Level
        $lblGen.Text = ("產出: {0} / AI v{1}" -f $D.GeneratedAt, $D.AIVersion)

        $nl = [Environment]::NewLine
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine(("Run {0} (前一 Run: {1})  Tool {2} / {3}  產品族 {4}" -f $D.RunAlias, $D.PrevRunAlias, $D.ToolAlias, $D.ChamberAlias, $D.ProductFamily))
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("【主要原因】")
        $i = 1
        foreach ($x in ($D.Stability.Reasons + $D.Maintenance.Reasons + $D.Drift.Reasons + $D.Overall.Reasons)) {
            [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++
        }
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("【建議人工檢查】")
        $i = 1
        foreach ($x in $D.Checks) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("【資料不足處】")
        if ($D.MissingData.Count -eq 0) { [void]$sb.AppendLine("無") }
        else {
            $i = 1
            foreach ($x in $D.MissingData) { [void]$sb.AppendLine(("{0}. {1}" -f $i, $x)); $i++ }
        }
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("【AI 不可判定事項】AI 不判定是否可開 Run、不建議修改 Recipe、不取代工程師覆判。")
        $rtb.Text = $sb.ToString()

        $grid.Rows.Clear()
        foreach ($s in $D.SimilarCases) {
            $c = $s.Case
            [void]$grid.Rows.Add(
                (Get-FieldString -Object $c -PropertyName "CaseID"),
                (Get-FieldString -Object $c -PropertyName "RiskCategory"),
                (Get-FieldString -Object $c -PropertyName "Summary"),
                (Get-FieldString -Object $c -PropertyName "Action"),
                (Get-FieldString -Object $c -PropertyName "Outcome"),
                (Get-FieldString -Object $c -PropertyName "CaseDate"))
        }
        if ($null -ne $chart) { Update-TrendChart -Chart $chart -ToolAlias $D.ToolAlias }
    }.GetNewClosure()

    $cbTool.Add_SelectedIndexChanged({
        try { & $fillRunCombo } catch { Write-ErrorLog ("cbTool changed: " + $_.Exception.Message) }
    }.GetNewClosure())

    $btnReload.Add_Click({
        try { & $reloadTools; Show-Info "資料已重新載入。" } catch {
            Write-ErrorLog ("btnReload: " + $_.Exception.Message)
            Show-ErrorMessage $_.Exception.Message
        }
    }.GetNewClosure())

    $btnDiag.Add_Click({
        try {
            $runAlias = [string]$cbRun.SelectedItem
            if ([string]::IsNullOrEmpty($runAlias)) { Show-Warn "請先選擇 Run。"; return }
            # v1.0.7: Log 檔名目錄項目顯示為 "Run碼 | 產品", 取 Run 碼部分
            $runAlias = $runAlias.Split('|')[0].Trim()
            $d = Invoke-RunDiagnosis -RunAlias $runAlias
            $St.LastDiagnosis = $d
            $St.LastReportPath = ""
            & $showDiagnosis $d
            $btnReport.Enabled = $true
            Write-AuditLog -Action "DIAGNOSIS_RUN" -Detail ("{0} overall={1}" -f $runAlias, (Get-RiskName $d.Overall.Level))
        } catch {
            Write-ErrorLog ("btnDiag: " + $_.Exception.Message)
            Show-ErrorMessage ("診斷失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())

    $btnReport.Add_Click({
        try {
            if ($null -eq $St.LastDiagnosis) { Show-Warn "請先執行診斷。"; return }
            $path = Save-DiagnosisReport -D $St.LastDiagnosis
            $St.LastReportPath = $path
            Show-Info ("報告已產出:" + [Environment]::NewLine + $path)
            try { Invoke-Item -LiteralPath $path } catch { }
        } catch {
            Write-ErrorLog ("btnReport: " + $_.Exception.Message)
            Show-ErrorMessage ("報告產出失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())

    $btnLogCmp.Add_Click({
        try {
            if ($null -eq $St.CurrentUser) { Show-Warn "請先登入。"; return }
            Show-LogCompareDialog -St $St
        } catch {
            Write-ErrorLog ("btnLogCmp: " + $_.Exception.Message)
            Show-ErrorMessage ("前 Run Log 診斷開啟失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())

    # 啟動時載入
    try { & $reloadTools } catch { Write-ErrorLog ("Diagnosis tab init: " + $_.Exception.Message) }
}

# ============================================================
# Tab 2: 工程師覆判
# ============================================================
function Build-ReviewTab {
    param([System.Windows.Forms.TabPage]$Tab, [hashtable]$St)

    $lblDiag = New-UiLabel -Text "(尚未執行診斷)" -X 10 -Y 12 -Width 1020 -Height 24
    $lblDiag.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $Tab.Controls.Add($lblDiag)

    $Tab.Controls.Add((New-UiLabel -Text "覆判結果:" -X 10 -Y 52 -Width 80))
    $cbReview = New-UiComboBox -X 95 -Y 50 -Width 150
    foreach ($x in @("同意 AI 提醒", "部分同意", "不同意")) { [void]$cbReview.Items.Add($x) }
    $Tab.Controls.Add($cbReview)

    $Tab.Controls.Add((New-UiLabel -Text "最終處置:" -X 270 -Y 52 -Width 80))
    $cbDecision = New-UiComboBox -X 355 -Y 50 -Width 170
    foreach ($x in @("開 Run", "補查資料", "通知設備", "Hold", "Escalation")) { [void]$cbDecision.Items.Add($x) }
    $Tab.Controls.Add($cbDecision)

    $Tab.Controls.Add((New-UiLabel -Text "AI 誤報:" -X 550 -Y 52 -Width 70))
    $cbFp = New-UiComboBox -X 622 -Y 50 -Width 100
    foreach ($x in @("待確認", "是", "否")) { [void]$cbFp.Items.Add($x) }
    $cbFp.SelectedIndex = 0
    $Tab.Controls.Add($cbFp)

    $Tab.Controls.Add((New-UiLabel -Text "AI 漏報:" -X 740 -Y 52 -Width 70))
    $cbFn = New-UiComboBox -X 812 -Y 50 -Width 100
    foreach ($x in @("待確認", "是", "否")) { [void]$cbFn.Items.Add($x) }
    $cbFn.SelectedIndex = 0
    $Tab.Controls.Add($cbFn)

    $Tab.Controls.Add((New-UiLabel -Text "備註 (異常說明 / 補查內容 / Run 後實際結果):" -X 10 -Y 90 -Width 400))
    $txtComment = New-Object System.Windows.Forms.TextBox
    $txtComment.Left = 10; $txtComment.Top = 116; $txtComment.Width = 1020; $txtComment.Height = 110
    $txtComment.Multiline = $true
    $txtComment.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $Tab.Controls.Add($txtComment)

    $btnSave = New-UiButton -Text "儲存覆判 (留痕)" -X 10 -Y 240 -Width 150
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $Tab.Controls.Add($btnSave)

    $lblHint = New-UiLabel -Text "說明: 覆判紀錄僅保存 RoleCode (去識別化), 寫入 Data\records\HumanReview.csv, 作為後續模型驗證與 FP/FN 統計之依據。" -X 175 -Y 244 -Width 860 -Height 24
    $lblHint.ForeColor = [System.Drawing.Color]::DimGray
    $Tab.Controls.Add($lblHint)

    $refreshHeader = {
        try {
            $d = $St.LastDiagnosis
            if ($null -eq $d) { $lblDiag.Text = "(尚未執行診斷 - 請先於「Run 前診斷」頁執行)"; return }
            $lblDiag.Text = ("待覆判: Run {0} / Tool {1} / 總體風險 {2} (穩定度 {3} / 維修 {4} / 漂移 {5}) / AI v{6}" -f `
                $d.RunAlias, $d.ToolAlias, (Get-RiskName $d.Overall.Level), (Get-RiskName $d.Stability.Level),
                (Get-RiskName $d.Maintenance.Level), (Get-RiskName $d.Drift.Level), $d.AIVersion)
        } catch { Write-ErrorLog ("review refreshHeader: " + $_.Exception.Message) }
    }.GetNewClosure()
    $Tab.Tag = $refreshHeader   # 供主表單切頁時呼叫

    $btnSave.Add_Click({
        try {
            $d = $St.LastDiagnosis
            if ($null -eq $d) { Show-Warn "請先於「Run 前診斷」頁執行診斷。"; return }
            $u = $St.CurrentUser
            if ($null -eq $u) { Show-Warn "請先登入後再儲存覆判。"; return }
            $role = Get-FieldString -Object $u -PropertyName "Role"
            if ($role -ne "Engineer" -and $role -ne "Admin") {
                Show-Warn "僅工程師 (Engineer) 或管理員可儲存覆判。"; return
            }
            if ($null -eq $cbReview.SelectedItem) { Show-Warn "請選擇覆判結果。"; return }
            if ($null -eq $cbDecision.SelectedItem) { Show-Warn "請選擇最終處置。"; return }
            $roleCode = Get-FieldString -Object $u -PropertyName "RoleCode" -DefaultValue "UNKNOWN"
            Save-HumanReview -D $d `
                -EngineerReview ([string]$cbReview.SelectedItem) `
                -FinalDecision ([string]$cbDecision.SelectedItem) `
                -FalsePositive ([string]$cbFp.SelectedItem) `
                -FalseNegative ([string]$cbFn.SelectedItem) `
                -ReviewerRole $roleCode `
                -Comment $txtComment.Text.Trim() `
                -ReportFile $St.LastReportPath
            Show-Info "覆判已儲存並留痕。"
            $txtComment.Clear()
        } catch {
            Write-ErrorLog ("btnSaveReview: " + $_.Exception.Message)
            Show-ErrorMessage ("儲存覆判失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())
}

# ============================================================
# Tab 3: 覆判紀錄 / 留痕查詢
# ============================================================
function Build-HistoryTab {
    param([System.Windows.Forms.TabPage]$Tab, [hashtable]$St)

    $btnRefresh = New-UiButton -Text "重新整理" -X 10 -Y 10 -Width 100
    $Tab.Controls.Add($btnRefresh)
    $btnOpenFolder = New-UiButton -Text "開啟報告資料夾" -X 120 -Y 10 -Width 140
    $Tab.Controls.Add($btnOpenFolder)
    $lblCount = New-UiLabel -Text "" -X 280 -Y 14 -Width 500
    $lblCount.ForeColor = [System.Drawing.Color]::DimGray
    $Tab.Controls.Add($lblCount)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 10; $grid.Top = 48; $grid.Width = 1030; $grid.Height = 490
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.Anchor = "Top,Bottom,Left,Right"
    $cols = @("覆判時間", "Run", "Tool", "AI總體", "AI穩定度", "AI維修", "AI漂移", "覆判", "最終處置", "誤報", "漏報", "覆判角色", "備註")
    foreach ($c in $cols) { [void]$grid.Columns.Add($c, $c) }
    $grid.Columns[12].FillWeight = 180
    $Tab.Controls.Add($grid)

    $loadReviews = {
        try {
            $grid.Rows.Clear()
            # Pitfall 1: 閉包內不可讀 $script: 變數, 改用函式取路徑
            $rows = @(Import-CsvSafe -Path (Get-HumanReviewCsvPath))
            $sorted = @($rows | Sort-Object -Property @{ Expression = { Get-FieldString -Object $_ -PropertyName "ReviewTime" }; Descending = $true })
            foreach ($r in $sorted) {
                [void]$grid.Rows.Add(
                    (Get-FieldString -Object $r -PropertyName "ReviewTime"),
                    (Get-FieldString -Object $r -PropertyName "RunID_Alias"),
                    (Get-FieldString -Object $r -PropertyName "ToolAlias"),
                    (Get-FieldString -Object $r -PropertyName "AI_OverallRisk"),
                    (Get-FieldString -Object $r -PropertyName "AI_StabilityRisk"),
                    (Get-FieldString -Object $r -PropertyName "AI_MaintRisk"),
                    (Get-FieldString -Object $r -PropertyName "AI_DriftRisk"),
                    (Get-FieldString -Object $r -PropertyName "EngineerReview"),
                    (Get-FieldString -Object $r -PropertyName "FinalDecision"),
                    (Get-FieldString -Object $r -PropertyName "FalsePositive"),
                    (Get-FieldString -Object $r -PropertyName "FalseNegative"),
                    (Get-FieldString -Object $r -PropertyName "ReviewerRole"),
                    (Get-FieldString -Object $r -PropertyName "Comment"))
            }
            $lblCount.Text = ("共 {0} 筆覆判紀錄 (FP/FN 統計為每月治理會議輸入)" -f $sorted.Count)
        } catch {
            Write-ErrorLog ("loadReviews: " + $_.Exception.Message)
            Show-ErrorMessage $_.Exception.Message
        }
    }.GetNewClosure()

    $btnRefresh.Add_Click({ try { & $loadReviews } catch { Write-ErrorLog $_.Exception.Message } }.GetNewClosure())
    $btnOpenFolder.Add_Click({
        try { Invoke-Item -LiteralPath (Get-ReportRootPath) } catch {
            Write-ErrorLog ("open reports: " + $_.Exception.Message)
            Show-ErrorMessage $_.Exception.Message
        }
    }.GetNewClosure())

    try { & $loadReviews } catch { Write-ErrorLog ("History tab init: " + $_.Exception.Message) }
}

# ============================================================
# 設定編輯對話框 (config.json; 僅 Admin 由設定頁開啟)
# ============================================================
function Show-TableManagerDialog {
    <#
      資料表管理 (v1.0.15): 六張 Data\Import 資料表的檢視 / 編輯介面。
      - 表格直接編輯; 底部空白列輸入即新增; 點選列頭後按 Delete 刪除列
      - 儲存: 原子寫檔 + 寫檔前自動備份 (Backup\BeforeWrite) + 稽核留痕
      - 標頭以檔案現況為準 (保留使用者自訂欄位); 檔案不存在用範本標頭
      - 未儲存的變更在切換資料表 / 關閉視窗時會提醒
    #>
    param([hashtable]$St)

    $form = New-Form -Text "資料表管理 (Data\Import) — 儲存時自動備份與稽核留痕" -Width 1010 -Height 660
    $form.MinimizeBox = $false

    $form.Controls.Add((New-UiLabel -Text "資料表:" -X 10 -Y 14 -Width 60))
    $cbTable = New-UiComboBox -X 74 -Y 12 -Width 190
    foreach ($t in @("RunSummary", "ToolStability", "Maintenance", "ProductDrift", "MeasurementTrend", "HistoryCases")) {
        [void]$cbTable.Items.Add($t)
    }
    $form.Controls.Add($cbTable)

    $btnSave = New-UiButton -Text "儲存 (留痕)" -X 280 -Y 10 -Width 110
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($btnSave)
    $btnReloadT = New-UiButton -Text "重新載入" -X 400 -Y 10 -Width 100
    $form.Controls.Add($btnReloadT)
    $btnFolder = New-UiButton -Text "開啟資料夾" -X 510 -Y 10 -Width 100
    $form.Controls.Add($btnFolder)
    $btnColHelp = New-UiButton -Text "欄位說明" -X 620 -Y 10 -Width 90
    $form.Controls.Add($btnColHelp)

    $lblInfo = New-UiLabel -Text "" -X 10 -Y 44 -Width 980 -Height 22
    $lblInfo.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblInfo)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 10; $grid.Top = 70; $grid.Width = 974; $grid.Height = 490
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToDeleteRows = $true
    $grid.RowHeadersVisible = $true
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
    $grid.Anchor = "Top,Bottom,Left,Right"
    $form.Controls.Add($grid)

    $lblHint = New-UiLabel -Text "提示: 直接點儲存格修改; 最下方空白列輸入即新增; 點選列頭後按 Delete 刪除列。RunID_Alias 接受 Run 碼 (261176) 或 機台+Run 碼 (MAT06261176)。儲存後請於診斷頁按「重載資料」。" -X 10 -Y 570 -Width 980 -Height 40
    $lblHint.ForeColor = [System.Drawing.Color]::DimGray
    $lblHint.Anchor = "Bottom,Left,Right"
    $form.Controls.Add($lblHint)

    $descMap = @{
        RunSummary       = "Run 基本資料 (可由 Log 檔名自動補列; 人工列不會被覆蓋)"
        ToolStability    = "前一 Run 機台穩定度 (alarm / 分數; 需 MES 匯出或人工維護)"
        Maintenance      = "維修 / PM 特徵 (需 MES 匯出或人工維護)"
        ProductDrift     = "產品特性飄移 (量測總檔存在時可不填, 系統自動彙總)"
        MeasurementTrend = "趨勢圖資料 (可由量測總檔自動補列)"
        HistoryCases     = "歷史案例知識庫 (人工維護, 供相似案例推薦)"
    }
    $ui = @{ Table = ""; Header = @(); Dirty = $false; Loading = $false; Switching = $false }

    $loadTable = {
        try {
            $name = [string]$cbTable.SelectedItem
            if ([string]::IsNullOrEmpty($name)) { return }
            $ui.Loading = $true
            $grid.Rows.Clear()
            $grid.Columns.Clear()
            $path = Join-Path (Get-ImportRootPath) ($name + ".csv")
            $header = @()
            if (Test-Path -LiteralPath $path) {
                $sr = New-Object System.IO.StreamReader($path, [System.Text.Encoding]::UTF8, $true)
                try {
                    $hl = $sr.ReadLine()
                    while ($null -ne $hl -and $hl.Trim().Length -eq 0) { $hl = $sr.ReadLine() }
                    if (-not [string]::IsNullOrEmpty($hl)) { $header = @(Split-CsvLineFields -Line $hl) }
                } finally { $sr.Dispose() }
            }
            $hdrMap = Get-ImportTableHeaders
            if ($header.Count -eq 0 -and $hdrMap.ContainsKey($name + ".csv")) {
                $header = ([string]$hdrMap[$name + ".csv"]).Split(',')
            }
            foreach ($c in $header) { [void]$grid.Columns.Add([string]$c, [string]$c) }
            $rows = @(Import-CsvSafe -Path $path)
            foreach ($r in $rows) {
                $vals = @()
                foreach ($c in $header) { $vals += (Get-FieldString -Object $r -PropertyName ([string]$c)) }
                [void]$grid.Rows.Add([object[]]$vals)
            }
            $ui.Table = $name
            $ui.Header = @($header)
            $ui.Dirty = $false
            $desc = [string]$descMap[$name]
            $lblInfo.Text = ("{0}.csv — {1}  ({2} 筆)  路徑: {3}" -f $name, $desc, $rows.Count, $path)
        } catch {
            Write-ErrorLog ("TableManager load: " + $_.Exception.Message)
            Show-ErrorMessage ("載入資料表失敗: " + $_.Exception.Message)
        } finally { $ui.Loading = $false }
    }.GetNewClosure()

    $saveTable = {
        try {
            $name = [string]$ui.Table
            if ([string]::IsNullOrEmpty($name)) { return }
            $path = Join-Path (Get-ImportRootPath) ($name + ".csv")
            $header = @($ui.Header)
            $lines = New-Object System.Collections.Generic.List[string]
            $lines.Add((Join-CsvLine $header))
            $n = 0
            foreach ($row in $grid.Rows) {
                if ($row.IsNewRow) { continue }
                $vals = @()
                $allEmpty = $true
                for ($i = 0; $i -lt $header.Count; $i++) {
                    $v = ""
                    if ($i -lt $row.Cells.Count -and $null -ne $row.Cells[$i].Value) {
                        $v = ([string]$row.Cells[$i].Value).Trim()
                    }
                    if ($v.Length -gt 0) { $allEmpty = $false }
                    $vals += $v
                }
                if ($allEmpty) { continue }
                $lines.Add((Join-CsvLine $vals))
                $n++
            }
            Write-FileSafe -Path $path -Text (($lines.ToArray()) -join [Environment]::NewLine)
            Write-AuditLog -Action "TABLE_EDITED" -Detail ("{0}.csv rows={1}" -f $name, $n)
            $ui.Dirty = $false
            $lblInfo.Text = ("{0}.csv — 已儲存 ({1} 筆)  路徑: {2}" -f $name, $n, $path)
            Show-Info (("已儲存 {0}.csv ({1} 筆)。" -f $name, $n) + [Environment]::NewLine + `
                "寫檔前版本已自動備份於 Backup\BeforeWrite。" + [Environment]::NewLine + `
                "請於診斷頁按「重載資料」讓變更生效。")
        } catch {
            Write-ErrorLog ("TableManager save: " + $_.Exception.Message)
            Show-ErrorMessage ("儲存失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure()

    $confirmDiscard = {
        if (-not $ui.Dirty) { return $true }
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "目前資料表有未儲存的變更, 繼續將遺失。確定?", "未儲存的變更",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return ($ans -eq [System.Windows.Forms.DialogResult]::Yes)
    }.GetNewClosure()

    $grid.Add_CellValueChanged({ if (-not $ui.Loading) { $ui.Dirty = $true } }.GetNewClosure())
    $grid.Add_UserDeletedRow({ if (-not $ui.Loading) { $ui.Dirty = $true } }.GetNewClosure())

    $cbTable.Add_SelectedIndexChanged({
        try {
            if ($ui.Switching) { return }
            if (-not (& $confirmDiscard)) {
                $ui.Switching = $true
                $cbTable.SelectedItem = $ui.Table
                $ui.Switching = $false
                return
            }
            & $loadTable
        } catch { Write-ErrorLog ("TableManager switch: " + $_.Exception.Message) }
    }.GetNewClosure())

    $btnSave.Add_Click({ & $saveTable }.GetNewClosure())
    $btnReloadT.Add_Click({
        try { if (& $confirmDiscard) { & $loadTable } } catch { Show-ErrorMessage $_.Exception.Message }
    }.GetNewClosure())
    $btnFolder.Add_Click({
        try { Invoke-Item -LiteralPath (Get-ImportRootPath) } catch { Show-ErrorMessage $_.Exception.Message }
    }.GetNewClosure())
    $btnColHelp.Add_Click({
        try {
            Initialize-ImportTableTemplates | Out-Null
            Invoke-Item -LiteralPath (Join-Path (Split-Path -Parent (Get-DeidentMapPath)) "資料表欄位說明.txt")
        } catch { Show-ErrorMessage $_.Exception.Message }
    }.GetNewClosure())
    $form.Add_FormClosing({
        param($s, $e)
        if (-not (& $confirmDiscard)) { $e.Cancel = $true }
    }.GetNewClosure())

    $cbTable.SelectedIndex = 0
    [void]$form.ShowDialog()
}

function Show-ConfigEditorDialog {
    # 回傳 $true = 已儲存 (Pitfall 4: 以 $form.Tag 傳回結果)
    param([hashtable]$St)
    $cfg = $St.Config
    if ($null -eq $cfg) { throw "設定尚未載入。" }
    $sqlCfg = Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "Sql" -DefaultValue $null
    if ($null -eq $sqlCfg -or -not ($sqlCfg -is [hashtable])) {
        $sqlCfg = @{ Enabled = $false; ConnectionString = ""; Queries = @{} }
        $cfg["Sql"] = $sqlCfg
    }

    $form = New-Form -Text "編輯設定 (Data\config.json) - 存檔後立即生效" -Width 680 -Height 680
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.Tag = $null

    $form.Controls.Add((New-UiLabel -Text "資料來源模式:" -X 15 -Y 16 -Width 100))
    $cbMode = New-UiComboBox -X 118 -Y 14 -Width 90
    foreach ($x in @("CSV", "SQL")) { [void]$cbMode.Items.Add($x) }
    $curMode = ([string](Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "DataSource" -DefaultValue "CSV")).ToUpper()
    if ($curMode -eq "SQL") { $cbMode.SelectedIndex = 1 } else { $cbMode.SelectedIndex = 0 }
    $form.Controls.Add($cbMode)

    $chkSql = New-Object System.Windows.Forms.CheckBox
    $chkSql.Text = "啟用 SQL 介面 (僅 SELECT 唯讀; 啟用前須資安審核)"
    $chkSql.Left = 230; $chkSql.Top = 14; $chkSql.Width = 420
    $chkSql.Checked = (Test-TruthyFlag (Get-ObjectPropertyValueSafe -Object $sqlCfg -PropertyName "Enabled" -DefaultValue $false))
    $form.Controls.Add($chkSql)

    $form.Controls.Add((New-UiLabel -Text "SQL 連線字串:" -X 15 -Y 50 -Width 100))
    $txtConn = New-UiTextBox -X 118 -Y 48 -Width 530
    $txtConn.Text = [string](Get-ObjectPropertyValueSafe -Object $sqlCfg -PropertyName "ConnectionString" -DefaultValue "")
    $form.Controls.Add($txtConn)

    $lblNote = New-UiLabel -Text "資安要求: 建議 Integrated Security=SSPI, 禁止明碼帳密。各表 SELECT 查詢請直接編輯 config.json 的 Sql.Queries。" -X 15 -Y 76 -Width 640 -Height 22
    $lblNote.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblNote)

    # v1.0.13: Log 比對排除參數 (不是所有 Device 參數都需比對)
    $form.Controls.Add((New-UiLabel -Text "Log 比對排除參數 (萬用字元, 分號分隔; 如 Dop1.*;Hyd1.*;*.dp_SP):" -X 15 -Y 104 -Width 400))
    $txtExc = New-UiTextBox -X 420 -Y 102 -Width 230
    $txtExc.Text = [string](Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "LogCompareExcludeParams" -DefaultValue "")
    $form.Controls.Add($txtExc)

    $form.Controls.Add((New-UiLabel -Text "規則門檻 (數值; 對應 SOP 8.2 / 10.x 規則, 建議由治理小組會議決議後修改):" -X 15 -Y 132 -Width 640))
    $gridT = New-Object System.Windows.Forms.DataGridView
    $gridT.Left = 15; $gridT.Top = 158; $gridT.Width = 635; $gridT.Height = 394
    $gridT.AllowUserToAddRows = $false
    $gridT.AllowUserToDeleteRows = $false
    $gridT.RowHeadersVisible = $false
    $gridT.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    [void]$gridT.Columns.Add("Key", "門檻參數")
    [void]$gridT.Columns.Add("Value", "數值 (可編輯)")
    $gridT.Columns[0].ReadOnly = $true
    $th = Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "Thresholds" -DefaultValue $null
    if ($null -ne $th) {
        foreach ($k in ($th.Keys | Sort-Object)) {
            [void]$gridT.Rows.Add([string]$k, [string]$th[$k])
        }
    }
    $form.Controls.Add($gridT)

    $btnSave = New-UiButton -Text "儲存設定" -X 15 -Y 566 -Width 120
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($btnSave)
    $btnCancel = New-UiButton -Text "取消" -X 145 -Y 566 -Width 90
    $form.Controls.Add($btnCancel)

    $btnSave.Add_Click({
        try {
            $newTh = @{}
            foreach ($row in $gridT.Rows) {
                $k = [string]$row.Cells[0].Value
                $vRaw = ([string]$row.Cells[1].Value).Trim()
                $out = 0.0
                $style = [System.Globalization.NumberStyles]::Float
                $culture = [System.Globalization.CultureInfo]::InvariantCulture
                if (-not [double]::TryParse($vRaw, $style, $culture, [ref]$out)) {
                    Show-Warn ("門檻「{0}」的值「{1}」不是有效數字, 未儲存。" -f $k, $vRaw)
                    return
                }
                $newTh[$k] = $out
            }
            $mode = [string]$cbMode.SelectedItem
            $connStr = $txtConn.Text.Trim()
            if ($mode -eq "SQL" -and -not $chkSql.Checked) {
                Show-Warn "資料來源選 SQL 但未勾選啟用, 系統將拒絕讀取。請確認後再儲存。"
                return
            }
            if ($mode -eq "SQL" -and [string]::IsNullOrEmpty($connStr)) {
                Show-Warn "資料來源選 SQL 但連線字串為空, 未儲存。"
                return
            }
            if ($connStr -match "(?i)password\s*=") {
                Show-Warn "連線字串包含明碼密碼, 依資安要求建議改用 Integrated Security=SSPI。"
            }
            $cfg["Thresholds"] = $newTh
            $cfg["DataSource"] = $mode
            $cfg["LogCompareExcludeParams"] = $txtExc.Text.Trim()
            $sqlCfg["Enabled"] = [bool]$chkSql.Checked
            $sqlCfg["ConnectionString"] = $connStr
            Save-AppConfig
            $form.Tag = $true
            $form.Close()
        } catch {
            Write-ErrorLog ("ConfigEditor save: " + $_.Exception.Message)
            Show-ErrorMessage ("儲存設定失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())

    $btnCancel.Add_Click({ $form.Tag = $null; $form.Close() }.GetNewClosure())

    [void]$form.ShowDialog()
    return ($form.Tag -eq $true)
}

# ============================================================
# Tab 4: 設定 / 資料來源 / 去識別化
# ============================================================
function Build-SettingsTab {
    param([System.Windows.Forms.TabPage]$Tab, [hashtable]$St)

    $y = 12
    $info = New-Object System.Windows.Forms.RichTextBox
    $info.Left = 10; $info.Top = $y; $info.Width = 1030; $info.Height = 210
    $info.ReadOnly = $true
    $info.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 9.5)
    $Tab.Controls.Add($info)
    $y = 234

    # Pitfall 1: 閉包內不讀 $script: 變數, 路徑與版本一律經函式取得
    $refreshInfo = {
        try {
            $cfg = $St.Config
            $mode = [string](Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "DataSource" -DefaultValue "CSV")
            $sqlCfg = Get-ObjectPropertyValueSafe -Object $cfg -PropertyName "Sql" -DefaultValue $null
            $sqlEnabled = Test-TruthyFlag (Get-ObjectPropertyValueSafe -Object $sqlCfg -PropertyName "Enabled" -DefaultValue $false)
            $sqlText = "未啟用"
            if ($sqlEnabled) { $sqlText = "已啟用" }
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine(("系統版本: {0}    診斷方式: Layer1 規則引擎 + 模板摘要 (離線, 不外送資料)" -f (Get-AIVersion)))
            [void]$sb.AppendLine(("資料來源模式: {0}" -f $mode))
            [void]$sb.AppendLine(("SQL 介面: {0} (可由「編輯設定」調整; 啟用需通過資安審核, 僅允許 SELECT 唯讀)" -f $sqlText))
            [void]$sb.AppendLine(("LLM 介面: 預留, 預設停用 (啟用前須資安審核, 僅允許去識別化相對值)"))
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("資料夾:")
            [void]$sb.AppendLine(("  匯入區 (去識別化後): " + (Get-ImportRootPath)))
            [void]$sb.AppendLine(("  原始匯入區 (轉換前): " + (Get-RawImportRootPath)))
            [void]$sb.AppendLine(("  覆判留痕: " + (Get-HumanReviewCsvPath)))
            [void]$sb.AppendLine(("  診斷報告: " + (Get-ReportRootPath)))
            [void]$sb.AppendLine(("  去識別化對照表 (受控, 勿外流): " + (Get-DeidentMapPath)))
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("AI 權限邊界 (SOP Phase 0): 只做整理/比對/提醒/建議/留痕; 不開 Run、不改 Recipe、不放行產品、不取代 SPC/MES/QMS。")
            $info.Text = $sb.ToString()
        } catch { Write-ErrorLog ("refreshInfo: " + $_.Exception.Message) }
    }.GetNewClosure()

    $btnDeident = New-UiButton -Text "去識別化轉換 (RawImport->Import)" -X 10 -Y $y -Width 250
    $Tab.Controls.Add($btnDeident)
    $btnTables = New-UiButton -Text "資料表管理 (編輯六張表)" -X 268 -Y $y -Width 180
    $btnTables.BackColor = [System.Drawing.Color]::FromArgb(142, 68, 173)
    $btnTables.ForeColor = [System.Drawing.Color]::White
    $Tab.Controls.Add($btnTables)
    $btnOpenImport = New-UiButton -Text "開啟匯入資料夾" -X 456 -Y $y -Width 130
    $Tab.Controls.Add($btnOpenImport)
    $btnOpenLogs = New-UiButton -Text "開啟 Logs" -X 594 -Y $y -Width 90
    $Tab.Controls.Add($btnOpenLogs)
    $btnEditCfg = New-UiButton -Text "編輯設定 (Admin)" -X 692 -Y $y -Width 140
    $btnEditCfg.BackColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
    $btnEditCfg.ForeColor = [System.Drawing.Color]::White
    $Tab.Controls.Add($btnEditCfg)
    $btnStepMean = New-UiButton -Text "Log 秒檔 StepCode 平均" -X 840 -Y $y -Width 200
    $Tab.Controls.Add($btnStepMean)
    $y += 44

    $Tab.Controls.Add((New-UiLabel -Text "規則門檻 (按「編輯設定」修改, 建議治理小組每季檢視; SOP 14.1):" -X 10 -Y $y -Width 600))
    $y += 26
    $gridTh = New-Object System.Windows.Forms.DataGridView
    $gridTh.Left = 10; $gridTh.Top = $y; $gridTh.Width = 600; $gridTh.Height = 240
    $gridTh.ReadOnly = $true
    $gridTh.AllowUserToAddRows = $false
    $gridTh.RowHeadersVisible = $false
    $gridTh.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    [void]$gridTh.Columns.Add("Key", "門檻參數")
    [void]$gridTh.Columns.Add("Value", "現值")
    $Tab.Controls.Add($gridTh)

    $loadThresholds = {
        try {
            $gridTh.Rows.Clear()
            $th = Get-ObjectPropertyValueSafe -Object $St.Config -PropertyName "Thresholds" -DefaultValue $null
            if ($null -ne $th) {
                foreach ($k in ($th.Keys | Sort-Object)) {
                    [void]$gridTh.Rows.Add([string]$k, [string]$th[$k])
                }
            }
        } catch { Write-ErrorLog ("loadThresholds: " + $_.Exception.Message) }
    }.GetNewClosure()

    & $refreshInfo
    & $loadThresholds

    # ---- 使用者管理 (僅 Admin 可異動) ----
    $Tab.Controls.Add((New-UiLabel -Text "使用者管理 (僅 Admin 可異動; 覆判留痕僅記 RoleCode):" -X 630 -Y ($y - 26) -Width 410))
    $gridUsers = New-Object System.Windows.Forms.DataGridView
    $gridUsers.Left = 630; $gridUsers.Top = $y; $gridUsers.Width = 410; $gridUsers.Height = 140
    $gridUsers.ReadOnly = $true
    $gridUsers.AllowUserToAddRows = $false
    $gridUsers.RowHeadersVisible = $false
    $gridUsers.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridUsers.MultiSelect = $false
    $gridUsers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    foreach ($c in @("工號", "名稱", "RoleCode", "角色", "啟用")) { [void]$gridUsers.Columns.Add($c, $c) }
    $Tab.Controls.Add($gridUsers)

    $uy = $y + 150
    $Tab.Controls.Add((New-UiLabel -Text "工號:" -X 630 -Y ($uy + 2) -Width 45))
    $txtNewUid = New-UiTextBox -X 678 -Y $uy -Width 90
    $Tab.Controls.Add($txtNewUid)
    $Tab.Controls.Add((New-UiLabel -Text "名稱:" -X 776 -Y ($uy + 2) -Width 45))
    $txtNewName = New-UiTextBox -X 824 -Y $uy -Width 100
    $Tab.Controls.Add($txtNewName)
    $cbNewRole = New-UiComboBox -X 932 -Y $uy -Width 105
    foreach ($x in @("Engineer", "Viewer", "Admin")) { [void]$cbNewRole.Items.Add($x) }
    $cbNewRole.SelectedIndex = 0
    $Tab.Controls.Add($cbNewRole)

    $uy = $uy + 34
    $btnAddUser = New-UiButton -Text "新增使用者" -X 630 -Y $uy -Width 120
    $btnAddUser.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
    $btnAddUser.ForeColor = [System.Drawing.Color]::White
    $Tab.Controls.Add($btnAddUser)
    $btnToggleUser = New-UiButton -Text "停用 / 啟用選取帳號" -X 760 -Y $uy -Width 170
    $Tab.Controls.Add($btnToggleUser)
    $btnReloadUsers = New-UiButton -Text "重新整理" -X 940 -Y $uy -Width 97
    $Tab.Controls.Add($btnReloadUsers)

    $loadUsers = {
        try {
            $gridUsers.Rows.Clear()
            foreach ($u in (Get-AllUsers)) {
                $act = "停用"
                if ((Get-FieldString -Object $u -PropertyName "IsActive" -DefaultValue "0") -eq "1") { $act = "啟用" }
                [void]$gridUsers.Rows.Add(
                    (Get-FieldString -Object $u -PropertyName "UserId"),
                    (Get-FieldString -Object $u -PropertyName "DisplayName"),
                    (Get-FieldString -Object $u -PropertyName "RoleCode"),
                    (Get-FieldString -Object $u -PropertyName "Role"),
                    $act)
            }
        } catch {
            Write-ErrorLog ("loadUsers: " + $_.Exception.Message)
            Show-ErrorMessage $_.Exception.Message
        }
    }.GetNewClosure()

    $requireAdmin = {
        $u = $St.CurrentUser
        if ($null -eq $u) { Show-Warn "請先登入。"; return $false }
        if ((Get-FieldString -Object $u -PropertyName "Role") -ne "Admin") {
            Show-Warn "僅管理員 (Admin) 可異動使用者。"; return $false
        }
        return $true
    }.GetNewClosure()

    $btnAddUser.Add_Click({
        try {
            if (-not (& $requireAdmin)) { return }
            $roleCode = Add-UserAccount -UserId $txtNewUid.Text -DisplayName $txtNewName.Text -Role ([string]$cbNewRole.SelectedItem)
            Show-Info ("使用者已新增, RoleCode: " + $roleCode)
            $txtNewUid.Clear(); $txtNewName.Clear()
            & $loadUsers
        } catch {
            Write-ErrorLog ("btnAddUser: " + $_.Exception.Message)
            Show-ErrorMessage ("新增使用者失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())

    $btnToggleUser.Add_Click({
        try {
            if (-not (& $requireAdmin)) { return }
            if ($null -eq $gridUsers.CurrentRow) { Show-Warn "請先在清單中選取帳號。"; return }
            $uid = [string]$gridUsers.CurrentRow.Cells[0].Value
            $nowActive = ([string]$gridUsers.CurrentRow.Cells[4].Value -eq "啟用")
            Set-UserActive -UserId $uid -Active (-not $nowActive)
            & $loadUsers
        } catch {
            Write-ErrorLog ("btnToggleUser: " + $_.Exception.Message)
            Show-ErrorMessage ("變更帳號狀態失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())

    $btnReloadUsers.Add_Click({
        try { & $loadUsers } catch { Write-ErrorLog ("btnReloadUsers: " + $_.Exception.Message) }
    }.GetNewClosure())

    try { & $loadUsers } catch { Write-ErrorLog ("Settings tab users init: " + $_.Exception.Message) }

    $btnDeident.Add_Click({
        try {
            $u = $St.CurrentUser
            if ($null -eq $u -or (Get-FieldString -Object $u -PropertyName "Role") -ne "Admin") {
                Show-Warn "僅管理員 (Admin) 可執行去識別化轉換。"; return
            }
            # v1.0.6: 執行中顯示等待游標, 完成訊息含耗時
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
            try {
                $results = @(Convert-AllRawImports)
            } finally {
                [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
            }
            $sw.Stop()
            if ($results.Count -eq 0) { Show-Info ("RawImport 資料夾內無 CSV 檔:" + [Environment]::NewLine + (Get-RawImportRootPath)); return }
            Show-Info (("轉換完成 (耗時 {0} 秒):" -f [math]::Round($sw.Elapsed.TotalSeconds, 1)) + [Environment]::NewLine + ($results -join [Environment]::NewLine))
        } catch {
            Write-ErrorLog ("btnDeident: " + $_.Exception.Message)
            Show-ErrorMessage ("去識別化轉換失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())
    $btnOpenImport.Add_Click({
        try { Invoke-Item -LiteralPath (Get-ImportRootPath) } catch { Show-ErrorMessage $_.Exception.Message }
    }.GetNewClosure())
    $btnOpenLogs.Add_Click({
        try { Invoke-Item -LiteralPath (Get-LogRootPath) } catch { Show-ErrorMessage $_.Exception.Message }
    }.GetNewClosure())
    $btnStepMean.Add_Click({
        try {
            if ($null -eq $St.CurrentUser) { Show-Warn "請先登入。"; return }
            # Pitfall 1: 閉包內不讀 $script: 變數, 路徑一律經函式取得
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = "選擇以秒計的 Run Log CSV"
            $dlg.Filter = "CSV 檔 (*.csv)|*.csv"
            $dlg.InitialDirectory = (Get-RawImportRootPath)
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            # v1.0.5: 執行中顯示等待游標
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
            try {
                $res = Export-StepCodeMeanReport -LogPath $dlg.FileName
            } finally {
                [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
            }
            Show-Info (("StepCode 平均已輸出 (分組欄: {0}, 共 {1} 個 Step, {2} 秒資料):" -f `
                $res.GroupColumn, $res.GroupCount, $res.RowCount) + [Environment]::NewLine + $res.OutPath)
        } catch {
            Write-ErrorLog ("btnStepMean: " + $_.Exception.Message)
            Show-ErrorMessage ("StepCode 平均計算失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())
    $btnTables.Add_Click({
        try {
            $u = $St.CurrentUser
            if ($null -eq $u) { Show-Warn "請先登入。"; return }
            $role = Get-FieldString -Object $u -PropertyName "Role"
            if ($role -ne "Admin" -and $role -ne "Engineer") {
                Show-Warn "僅 Admin / Engineer 可編輯資料表 (Viewer 僅供瀏覽)。"; return
            }
            Show-TableManagerDialog -St $St
        } catch {
            Write-ErrorLog ("btnTables: " + $_.Exception.Message)
            Show-ErrorMessage ("資料表管理開啟失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())
    $btnEditCfg.Add_Click({
        try {
            $u = $St.CurrentUser
            if ($null -eq $u -or (Get-FieldString -Object $u -PropertyName "Role") -ne "Admin") {
                Show-Warn "僅管理員 (Admin) 可編輯設定。"; return
            }
            $saved = Show-ConfigEditorDialog -St $St
            if ($saved) {
                & $loadThresholds
                & $refreshInfo
                Show-Info "設定已儲存 (Data\config.json), 新門檻立即生效。"
            }
        } catch {
            Write-ErrorLog ("btnEditCfg: " + $_.Exception.Message)
            Show-ErrorMessage ("編輯設定失敗: " + $_.Exception.Message)
        }
    }.GetNewClosure())
}

# ============================================================
# 主表單
# ============================================================
function Show-MainForm {
    $form = New-Form -Text ($script:AppName + "  v" + $script:AIVersion + "  [輔助診斷, 不取代工程師覆判]") -Width 1100 -Height 700

    # 登入列
    $form.Controls.Add((New-UiLabel -Text "工號:" -X 15 -Y 14 -Width 45))
    $txtUid = New-UiTextBox -X 62 -Y 12 -Width 120
    $form.Controls.Add($txtUid)
    $btnLogin  = New-UiButton -Text "登入" -X 192 -Y 10 -Width 70
    $btnLogout = New-UiButton -Text "登出" -X 268 -Y 10 -Width 70
    $btnLogout.Enabled = $false
    $form.Controls.Add($btnLogin); $form.Controls.Add($btnLogout)

    $lblUser = New-UiLabel -Text "狀態: 尚未登入 (可瀏覽; 覆判與去識別化需登入)" -X 350 -Y 14 -Width 720 -Height 22
    $lblUser.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $form.Controls.Add($lblUser)

    $st = $script:AppState

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Left = 10; $tabs.Top = 44; $tabs.Width = 1064; $tabs.Height = 600
    $tabs.Anchor = "Top,Bottom,Left,Right"
    $form.Controls.Add($tabs)

    $tabDiag    = New-Object System.Windows.Forms.TabPage; $tabDiag.Text    = "Run 前診斷"
    $tabReview  = New-Object System.Windows.Forms.TabPage; $tabReview.Text  = "工程師覆判"
    $tabHistory = New-Object System.Windows.Forms.TabPage; $tabHistory.Text = "覆判紀錄 / 留痕"
    $tabSetting = New-Object System.Windows.Forms.TabPage; $tabSetting.Text = "設定 / 資料來源"
    foreach ($tp in @($tabDiag, $tabReview, $tabHistory, $tabSetting)) {
        [void]$tabs.TabPages.Add($tp)
    }

    # Pitfall 7.7: 非作用中 TabPage 尚無實際尺寸, 先強制設計尺寸再建子控件
    $designSize = New-Object System.Drawing.Size(1056, 572)
    foreach ($tp in @($tabDiag, $tabReview, $tabHistory, $tabSetting)) { $tp.Size = $designSize }

    Build-DiagnosisTab -Tab $tabDiag    -St $st
    Build-ReviewTab    -Tab $tabReview  -St $st
    Build-HistoryTab   -Tab $tabHistory -St $st
    Build-SettingsTab  -Tab $tabSetting -St $st

    $tabs.Add_SelectedIndexChanged({
        try {
            if ($tabs.SelectedTab -eq $tabReview -and $null -ne $tabReview.Tag) {
                & $tabReview.Tag
            }
        } catch { Write-ErrorLog ("tab changed: " + $_.Exception.Message) }
    }.GetNewClosure())

    $doLogin = {
        try {
            $uid = $txtUid.Text.Trim()
            if ([string]::IsNullOrEmpty($uid)) { Show-Warn "請輸入工號。"; return }
            $u = Find-UserById -UserId $uid
            if ($null -eq $u) { Show-Warn "找不到使用者或帳號已停用。"; return }
            $st.CurrentUser = $u
            $st.IsAdmin = ((Get-FieldString -Object $u -PropertyName "Role") -eq "Admin")
            $lblUser.Text = ("狀態: 已登入 {0} ({1}) / 角色 {2}" -f `
                (Get-FieldString -Object $u -PropertyName "RoleCode"),
                (Get-FieldString -Object $u -PropertyName "DisplayName"),
                (Get-FieldString -Object $u -PropertyName "Role"))
            $lblUser.ForeColor = [System.Drawing.Color]::DarkGreen
            $btnLogin.Enabled = $false; $btnLogout.Enabled = $true; $txtUid.Enabled = $false
            Write-AuditLog -Action "LOGIN" -Detail (Get-FieldString -Object $u -PropertyName "RoleCode")
        } catch {
            Write-ErrorLog ("Login failed: " + $_.Exception.Message)
            Show-ErrorMessage $_.Exception.Message
        }
    }.GetNewClosure()

    $btnLogin.Add_Click({ & $doLogin }.GetNewClosure())
    $btnLogout.Add_Click({
        try {
            if ($null -ne $st.CurrentUser) {
                Write-AuditLog -Action "LOGOUT" -Detail (Get-FieldString -Object $st.CurrentUser -PropertyName "RoleCode")
            }
            $st.CurrentUser = $null
            $st.IsAdmin = $false
            $lblUser.Text = "狀態: 尚未登入 (可瀏覽; 覆判與去識別化需登入)"
            $lblUser.ForeColor = [System.Drawing.Color]::DarkSlateGray
            $btnLogin.Enabled = $true; $btnLogout.Enabled = $false
            $txtUid.Enabled = $true; $txtUid.Clear()
        } catch { Write-ErrorLog ("Logout failed: " + $_.Exception.Message) }
    }.GetNewClosure())
    $txtUid.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true; & $doLogin
        }
    }.GetNewClosure())

    [void]$form.ShowDialog()
}

# ============================================================
# Startup
# ============================================================
try {
    if ($SelfTest) {
        Initialize-AppFolders   # 先確保正式資料夾存在 (log 用)
        $code = Invoke-SelfTest
        exit $code
    }
    Initialize-AppFolders
    Initialize-Config
    Initialize-UsersCsv
    Initialize-HumanReviewCsv
    Initialize-ImportTableTemplates | Out-Null   # v1.0.9: 資料表空白範本 (既有檔案不動)
    Write-AppLog ($script:AppName + " v" + $script:AIVersion + " started.")
    Show-MainForm
    Write-AppLog ($script:AppName + " closed.")
} catch {
    Write-ErrorLog ("Fatal: " + $_.Exception.Message + " | " + $_.ScriptStackTrace)
    Show-ErrorMessage ("系統發生錯誤: " + $_.Exception.Message)
    if ($SelfTest) { exit 1 }
}
# EOF RunPreCheckAI.ps1 v1.0.15



