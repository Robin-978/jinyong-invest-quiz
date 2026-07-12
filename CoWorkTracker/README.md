# CoWorkTracker 協同工作紀錄系統

以 PowerShell 5.1 + .NET WinForms 打造、JSON 為資料庫、可在內網多人共用的協同工作紀錄系統。

## 功能特色

- 主頁專案列表：新使用者註冊、新專案申請、申請參與專案
- 管理頁面：人員管理、專案管理、Log 查詢/匯出、備份/還原
- 專案：名稱、編碼、說明、機密等級、成員權限、起迄期限
- 階段管理：KPI、負責人、協同人員、進度條、進度說明、附件上傳（圖檔/檔案）
- 進度視覺化：整體進度條與各階段橫條圖
- 權限分級：Admin、Owner（發起人/主持人）、Editor（協作可編輯）、Viewer（唯讀）
- 新使用者註冊；申請加入專案需發起人/主持人核可
- 密碼採 PBKDF2-SHA256 雜湊（隨機 salt、可調 iterations）
- 全程 Log 採 AES-256-CBC 加密，Admin/發起人可查詢與匯出 CSV
- 每個專案具備獨立編碼（自動產生）
- 啟動時自動備份 Data + Logs + Uploads 為 ZIP，可隨時手動備份或從備份還原
- 多人內網共用：所有 JSON 寫入採檔案鎖 + 重試 + 原子化置換，避免讀寫衝突
- 完整防空值：所有資料模型欄位皆做 null 防呆，UI 操作前後皆檢查清單範圍

## 系統需求

- Windows 7 SP1 / Windows 10 / Windows 11 / Windows Server 2012R2 以上
- Windows PowerShell 5.1（Windows 內建）
- .NET Framework 4.5 以上（含 System.IO.Compression.FileSystem）

均為 Windows 原生元件，不需額外安裝。

## 目錄結構

```
CoWorkTracker/
├── Start-CoWorkTracker.ps1   ← 啟動程式
├── Test-Syntax.ps1           ← 語法自我驗證
├── config.json               ← 設定檔（含 DataRoot 可指向網路共用路徑）
├── Modules/
│   ├── Core.psm1
│   ├── Security.psm1
│   ├── Logging.psm1
│   ├── Backup.psm1
│   ├── Users.psm1
│   └── Projects.psm1
├── UI/
│   ├── LoginForm.ps1
│   ├── RegisterForm.ps1
│   ├── MainForm.ps1
│   ├── ProjectForm.ps1
│   └── AdminForm.ps1
├── Data/                     ← users.json / projects.json / memberships.json / join_requests.json / master.key
├── Logs/                     ← log_YYYYMMDD.enc（加密）
├── Backups/                  ← backup_*.zip
└── Uploads/                  ← <專案編碼>/<階段ID>/<檔案>
```

## 快速開始

1. 將整個 `CoWorkTracker` 資料夾解壓到要使用的位置（單機或網路共用路徑）。
2. 若要多人共用，建議將整個資料夾放在「網路芳鄰」共用資料夾，所有人從相同 UNC 路徑執行。
3. 開啟 PowerShell：
   ```powershell
   Set-Location <資料夾路徑>
   powershell -ExecutionPolicy Bypass -File .\Start-CoWorkTracker.ps1
   ```
   若需指定資料根目錄（可放在共用磁碟 / NAS）：
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Start-CoWorkTracker.ps1 -DataRoot '\\server\share\CoWorkTracker'
   ```
4. 預設管理員：`admin / admin@123`，**首次登入請立刻使用「修改密碼」變更**。

## 多人內網共用注意事項

- 將整個資料夾放在一台機器分享出來，或將 `DataRoot` 指向網路共用路徑。
- 所有使用者需要對該共用資料夾擁有「修改」NTFS 權限。
- `Data/master.key` 為加密 Log 的金鑰，請以 NTFS 權限限制只允許具備 Admin/發起人 角色的 Windows 帳號讀取。
- JSON 寫入採取暫存檔 + 原子化置換 + 鎖檔重試，避免多人同時寫入造成損毀；同時每次寫入會保留 `.bak`。

## 權限模型

| 角色 | 主頁能看見 | 專案內權限 |
| ---- | ---- | ---- |
| Admin（系統管理員） | 全部專案 | 全部專案皆可管理 |
| Owner（發起人/主持人） | 含自己參與的專案 | 完整管理、核可加入、設定成員權限 |
| Editor（協作人員） | 自己參與的專案 | 可新增/編輯階段、回報進度、上傳附件 |
| Viewer（參與者） | 自己參與的專案 | 唯讀 |

## 安全性

- 密碼：PBKDF2-HMAC-SHA256，10 萬次迭代、32 byte salt、32 byte 雜湊（皆 Base64 儲存）
- Log：AES-256-CBC，每筆獨立 IV，金鑰來自 `Data/master.key`
- 附件：保留來源檔名 + 時間戳，並計算 SHA-256 供完整性驗證
- Log 匯出僅 Admin 進入管理頁面後可執行；專案發起人可於專案頁進入後利用管理員權限轉移後匯出（依需求可進一步限縮）

## 備份 / 還原

- 啟動時自動建立啟動備份（可在 `config.json` 關閉 `AutoBackupOnStart`）
- 管理頁面 → Log 查詢分頁，可手動備份與還原
- 還原前會自動再做一次「還原前備份」，避免操作失誤丟失資料
- 保留份數可在 `config.json` 透過 `BackupKeepCount` 設定（預設 30 份）

## 自我驗證

執行下列指令可驗證所有 PowerShell 檔案語法正確：

```powershell
powershell -ExecutionPolicy Bypass -File .\Test-Syntax.ps1
```

通過時會顯示 `全部通過 (N 檔案)`。

## Log 動作清單（節錄）

| Action | 意義 |
| ---- | ---- |
| login.success / login.fail | 登入結果 |
| user.create / user.update / user.changePassword / user.resetPassword | 帳號管理 |
| project.create / project.update / project.delete / project.open | 專案操作 |
| project.join.request / project.join.approved / project.join.rejected | 加入申請 |
| membership.set / membership.remove | 成員權限 |
| phase.add / phase.update / phase.remove / phase.attach | 階段操作 |
| log.export | Log 匯出 |
| backup.manual / backup.restore | 備份/還原 |
| app.start / app.exit / app.error | 程式生命週期 |

## 變更/擴充

- 所有資料模型位於 `Modules/Projects.psm1` / `Modules/Users.psm1`，UI 與資料分離，便於擴充。
- 若需要其他統計圖表，可在 `UI/ProjectForm.ps1` 的 `Draw-Chart` 區塊以 `System.Drawing.Graphics` 直接繪製。
