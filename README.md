# jinyong-invest-quiz

## 🍉 體感接水果（fruit-catcher/）

用手機**前鏡頭**偵測手部動作的體感網頁遊戲：

- 🖐️ 揮動手掌接住從左右飛來的水果，**水果越大、分數越高**
- 💣 炸彈 −40 分、🟢 榴槤 −30 分，千萬別碰
- 🔥 連續接到有連擊加分，一回合 60 秒
- 手部偵測使用 MediaPipe HandLandmarker；若模型載入失敗會自動退回「移動偵測」模式
- 相機畫面只在裝置上即時處理，不會上傳

**遊玩方式**：以 HTTPS 開啟 `fruit-catcher/index.html`（例如 GitHub Pages），允許相機權限即可開始。
