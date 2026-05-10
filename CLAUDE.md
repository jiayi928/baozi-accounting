# 包子記帳 (Baozi Accounting) — 專案說明

## 專案概覽
一個 PWA 記帳 App，資料存在 Google Sheets，支援 LINE Bot 推播通知。
- **線上網址**: https://jiayi928.github.io/baozi-accounting/
- **GitHub**: https://github.com/jiayi928/baozi-accounting
- **部署方式**: GitHub Pages（推送 master 分支自動更新）

## 技術架構
- **前端**: 純 HTML/CSS/JS，無框架，PWA（含 Service Worker）
- **資料儲存**: Google Sheets API v4（OAuth 2.0）
- **LINE 通知**: Google Cloud Run 代理 → LINE Messaging API
- **LIFF**: LINE Front-end Framework，讓用戶一鍵綁定 LINE 帳號

## 重要設定值（js/config.js）
```js
CLIENT_ID: '50636553434-f7pqcgi9ovbd8f6a4hi0tcmensg0nbjb.apps.googleusercontent.com'
SPREADSHEET_TITLE: '包子記帳資料'
GAS_PROXY_URL: 'https://line-bot-proxy-50636553434.asia-east1.run.app'
LINE_AUTH_KEY: 'baozi2025'
LIFF_ID: '2009972460-c1ivDfyC'
```

## 檔案結構
```
/
├── index.html          # 主 App（所有頁面用 display:none 切換）
├── sw.js               # Service Worker（快取版本：目前 baozi-v9）
├── manifest.json       # PWA manifest
├── liff-callback.html  # LINE 帳號綁定回調頁
├── guide.html          # 使用說明頁
├── privacy.html        # 隱私權政策
├── js/
│   ├── config.js       # 全域設定
│   ├── auth.js         # Google OAuth 2.0
│   ├── sheets.js       # Google Sheets 所有讀寫邏輯
│   └── app.js          # 主要 UI 邏輯
└── CLAUDE.md           # 本檔案
```

## Google Sheets 試算表結構
試算表名稱：「包子記帳資料」，共 5 個工作表：

### 紀錄頁（RECORDS）
| A 日期時間 | B 收支類型 | C 項目名稱 | D 金額 | E 會計科目 | F 轉出帳戶(扣款) | G 轉入帳戶(存入) | H 備註 |

### 帳戶與信用卡（ACCOUNTS）
銀行：A 帳戶名稱 | B 初始餘額 | C 目前餘額（SUMIFS公式）| D 備註
信用卡：F 信用卡名稱 | G 總信用額度 | H 本月應付（公式）| I 剩餘額度（公式）| J 結帳日 | K 繳款日 | L 自動扣款帳戶

### AI 專用分析與預算（BUDGET）
A 會計科目 | B 每月預算 | C 本月實際花費（SUMIFS公式）| D 剩餘預算差異（公式）

### 固定支出與轉帳（FIXED）
A 項目名稱 | B 收支類型 | C 每月金額 | D 每月扣款日 | E 總期數 | F 已繳期數 | G 轉出帳戶 | H 轉入帳戶 | I 會計科目 | J 未繳總餘額（公式）

### 個人設定（SETTINGS）
key-value 格式，對應 CONFIG.DEFAULT_SETTINGS

## 重要邏輯說明

### Service Worker 快取
每次修改前端檔案後，必須更新 `sw.js` 的 `CACHE` 版本號（baozi-v1, v2... 遞增）否則用戶看不到更新。

### 帳戶餘額計算
App 在 JS 裡自己算（不依賴試算表公式）：
- `_calcBankBalance(name, initial, allRecords)` → 初始餘額 + 轉入 - 轉出
- `_calcCardDebt(name, allRecords)` → 累計刷卡 - 累計還款
- `_calcCurrentBill(name, totalDebt, billDate, allRecords)` → 本期帳單

### 類型選擇器
記帳 Modal 有兩區：
1. Segmented Control（支出/收入/轉帳）— id="mainTypeScroll"
2. 小型 Chip 列（繳卡費/固定支出/訂閱/點數折抵/手續費/借出/收回）— class="type-chips"

`selectMainType(v, b)` 只清除 `#recordForm .type-btn` 的 active，不影響其他頁面的按鈕。

### LINE 整合流程
1. 用戶點「連結 LINE 帳號」→ 跳到 LIFF (`liff.line.me/2009972460-c1ivDfyC`)
2. LIFF 取得 LINE userId → 用 URL 參數帶回 App
3. App 讀取 URL 參數 → 存入 Google Sheets 設定頁
4. 記帳時呼叫 Cloud Run Proxy → 推播 LINE 訊息

## 外部服務
| 服務 | 用途 | 備註 |
|------|------|------|
| Google Cloud Console | OAuth Client ID、API 啟用 | 專案：baozi-accounting |
| Google Cloud Run | LINE Bot Proxy | asia-east1，服務名：line-bot-proxy |
| LINE Developers | Messaging API + LIFF | Bot ID: @570lljaw |
| GitHub Pages | 靜態網站部署 | 推 master 自動上線 |

## 常見操作
- **推送更新**: `git add . && git commit -m "..." && git push origin master`
- **重置試算表**: App 內 設定 → 進階管理 → 重置
- **強制用戶更新**: 修改 sw.js 的 CACHE 版本號
