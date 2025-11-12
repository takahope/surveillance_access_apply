# 監視器調閱權限申請系統 / Surveillance Camera Access Request System

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://script.google.com)
[![Bootstrap](https://img.shields.io/badge/Bootstrap-4.5.2-563D7C?logo=bootstrap&logoColor=white)](https://getbootstrap.com)
[![License](https://img.shields.io/badge/License-Proprietary-red.svg)](LICENSE)

一個基於 Google Apps Script 開發的企業級監視器調閱權限申請與審核系統，提供完整的多階段審核流程、角色權限管理及自動化通知功能。

An enterprise-grade surveillance camera access request and approval system built on Google Apps Script, featuring multi-stage approval workflows, role-based access control, and automated notifications.

---

## 📋 目錄 / Table of Contents

- [功能特色 / Features](#功能特色--features)
- [系統架構 / Architecture](#系統架構--architecture)
- [技術堆疊 / Tech Stack](#技術堆疊--tech-stack)
- [專案結構 / Project Structure](#專案結構--project-structure)
- [資料庫架構 / Database Schema](#資料庫架構--database-schema)
- [安裝部署 / Installation](#安裝部署--installation)
- [使用說明 / Usage](#使用說明--usage)
- [權限管理 / Permission Management](#權限管理--permission-management)
- [開發指南 / Development](#開發指南--development)
- [版本歷史 / Version History](#版本歷史--version-history)

---

## 🚀 功能特色 / Features

### 核心功能 / Core Features

- **三階段審核工作流** / **Three-Stage Approval Workflow**
  - 第一階段審核 (審核1)
  - 第二階段審核 (審核2)
  - 系統開通階段 (開通)
  - 即時狀態追蹤：審核1中 → 審核2中 → 開通中 → 已開通

- **角色權限控制 (RBAC)** / **Role-Based Access Control**
  - 三種管理員角色：審核1人員、審核2人員、開通人員
  - 前後端雙重權限驗證
  - Google Workspace 帳號認證整合

- **申請管理** / **Application Management**
  - 線上提交攝影機調閱申請
  - 可選擇申請人與攝影機地點
  - 可設定開通天數 (1-5天)
  - 詳細事由說明欄位

- **批次操作** / **Batch Operations**
  - 批次核准多筆申請
  - 批次更新申請狀態
  - 全選/取消全選功能

- **郵件通知** / **Email Notifications**
  - 每階段自動發送通知信
  - 自訂郵件範本與操作連結
  - HTML 格式化郵件內容

- **效能最佳化** / **Performance Optimization**
  - 管理員清單快取 (10分鐘)
  - 使用者名稱快取 (6小時)
  - 批次 API 操作減少呼叫次數

- **安全防護** / **Security Features**
  - 公式注入防護
  - XSS 跨站腳本攻擊防護
  - 敏感操作權限檢查
  - 網域限制存取

---

## 🏗️ 系統架構 / Architecture

```
┌─────────────────────────────────────────────────────────┐
│                   Google Workspace                       │
│  ┌──────────────────────────────────────────────────┐  │
│  │         Google Apps Script Web App               │  │
│  │  ┌────────────────────────────────────────────┐  │  │
│  │  │  Frontend (HTML/CSS/JavaScript)            │  │  │
│  │  │  - index.html (Portal)                     │  │  │
│  │  │  - 表單.html (Application Form)            │  │  │
│  │  │  - myapply.html (User Applications)        │  │  │
│  │  │  - review.html (Admin Panel)               │  │  │
│  │  │  - unauthorized.html (Access Denied)       │  │  │
│  │  └────────────────────────────────────────────┘  │  │
│  │               ↕️ google.script.run                │  │
│  │  ┌────────────────────────────────────────────┐  │  │
│  │  │  Backend (Apps Script)                     │  │  │
│  │  │  - 程式碼.js (Main Logic - 732 lines)      │  │  │
│  │  │    ├── Form Processing                     │  │  │
│  │  │    ├── Approval Workflow                   │  │  │
│  │  │    ├── Permission Management               │  │  │
│  │  │    ├── Email Notifications                 │  │  │
│  │  │    └── Data Access Layer                   │  │  │
│  │  └────────────────────────────────────────────┘  │  │
│  └────────────────┬────────────────────────────────┘  │
│                   ↕️                                    │
│  ┌────────────────────────────────────────────────┐  │
│  │  Google Sheets (Database)                      │  │
│  │  - 申請紀錄 (Application Records)              │  │
│  │  - 申請人與攝影機資料 (Requester-Camera Map)   │  │
│  │  - 審核開通 (Admin Configuration)              │  │
│  │  - 申請人/信箱 (User Data)                     │  │
│  └────────────────────────────────────────────────┘  │
│                   ↕️                                    │
│  ┌────────────────────────────────────────────────┐  │
│  │  Gmail (Email Service)                         │  │
│  │  - Automated Notifications                     │  │
│  └────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────┘
```

**架構類型** / **Architecture Type**: Serverless Web Application
**前端** / **Frontend**: HTML5 + CSS3 + JavaScript (ES6+)
**後端** / **Backend**: Google Apps Script (V8 Runtime)
**資料庫** / **Database**: Google Sheets (NoSQL-like)
**託管** / **Hosting**: Google Apps Script Web App Service

---

## 🛠️ 技術堆疊 / Tech Stack

### 前端技術 / Frontend Technologies
- **HTML5** - 網頁結構
- **CSS3** - 樣式設計與動畫效果
- **JavaScript (ES6+)** - 前端邏輯處理
- **Bootstrap 4.5.2** - UI 框架與響應式設計
- **Google Apps Script Client API** - 前後端通訊

### 後端技術 / Backend Technologies
- **Google Apps Script** - 伺服器端運行環境
- **SpreadsheetApp** - Google 試算表 API
- **MailApp** - Gmail API
- **ScriptApp** - Apps Script 服務
- **CacheService** - 快取服務
- **Session** - 使用者會話管理

### 開發工具 / Development Tools
- **Git** - 版本控制系統
- **Google Apps Script Editor** - 線上 IDE
- **clasp** (可選) - 命令列部署工具

---

## 📁 專案結構 / Project Structure

```
surveillance_access_apply/
├── .git/                      # Git 版本控制
├── .gitignore                 # Git 忽略設定
├── appsscript.json            # Apps Script 專案配置
├── 程式碼.js                   # 主要後端邏輯 (732 lines)
│   ├── 全域配置
│   ├── 欄位索引設定
│   ├── 資料格式化函數
│   ├── 權限管理模組
│   ├── 審核工作流引擎
│   ├── 表單處理邏輯
│   ├── 郵件通知系統
│   └── 自訂選單功能
├── index.html                 # 入口/首頁 (98 lines)
├── 表單.html                   # 申請表單頁面 (170 lines)
├── myapply.html               # 個人申請紀錄 (115 lines)
├── review.html                # 管理員審核面板 (150 lines)
└── unauthorized.html          # 存取拒絕頁面 (36 lines)

總計: 1,301 行程式碼
```

### 檔案說明 / File Descriptions

| 檔案 / File | 用途 / Purpose | 存取權限 / Access |
|------------|---------------|------------------|
| `程式碼.js` | 核心後端邏輯、工作流引擎、API 端點 | 內部 |
| `index.html` | 系統入口與導航頁面 | 公開 |
| `表單.html` | 監視器調閱申請表單 | 已認證使用者 |
| `myapply.html` | 個人申請紀錄查詢 | 已認證使用者 |
| `review.html` | 管理員審核控制台 | 僅管理員 |
| `unauthorized.html` | 權限不足提示頁面 | 公開 |
| `appsscript.json` | Apps Script 專案設定 | 配置 |

---

## 🗄️ 資料庫架構 / Database Schema

系統使用 Google 試算表作為資料庫，包含 4 個工作表：

### 1. 申請紀錄 (Application Records)

主要資料表，記錄所有申請案件與審核狀態。

| 欄位 | 欄位名稱 | 資料類型 | 說明 |
|-----|---------|---------|------|
| A | 填表帳號 | String | 提交表單的使用者信箱 |
| B | 申請時間 | DateTime | 申請提交時間戳記 |
| C | 申請人姓名 | String | 實際申請人姓名 |
| D | 攝影機地點 | String | 欲調閱的攝影機位置 |
| E | 調閱事由 | String | 申請原因說明 |
| F | 審核狀態 | Enum | 審核1中/審核2中/開通中/已開通 |
| G | 審核1帳號 | String | 第一階段審核人員信箱 |
| H | 審核1時間 | DateTime | 第一階段審核時間 |
| I | 審核2帳號 | String | 第二階段審核人員信箱 |
| J | 審核2時間 | DateTime | 第二階段審核時間 |
| K | 開通人帳號 | String | 系統開通人員信箱 |
| L | 開通時間 | DateTime | 系統開通時間 |
| M | 開通天數 | Number | 權限有效天數 (1-5天) |
| N | 系統帶入的申請人姓名 | String | 系統自動解析的申請人名稱 |

### 2. 申請人與攝影機資料 (Requester-Camera Mapping)

定義申請人與可申請攝影機的對應關係。

| 欄位 | 欄位名稱 | 說明 |
|-----|---------|------|
| 1 | 攝影機地點 | 攝影機位置名稱 |
| 3 | 申請人姓名 | 可申請該攝影機的人員 |

### 3. 審核開通 (Admin Configuration)

設定三階段審核人員名單。

| 欄位 | 欄位名稱 | 說明 |
|-----|---------|------|
| B | 審核1人員 | 第一階段審核人員信箱清單 |
| D | 審核2人員 | 第二階段審核人員信箱清單 |
| F | 開通人員 | 系統開通人員信箱清單 |

### 4. 申請人/信箱 (User Data)

使用者信箱與顯示名稱對應表。

| 欄位 | 欄位名稱 | 說明 |
|-----|---------|------|
| A | 使用者名稱 | 顯示用的中文姓名 |
| B | 登入者帳號 | Google Workspace 信箱 |

---

## 📦 安裝部署 / Installation

### 前置需求 / Prerequisites

1. **Google Workspace 帳號** (G Suite)
2. **Google Drive 存取權限**
3. **Google Apps Script 專案建立權限**

### 部署步驟 / Deployment Steps

#### 方法 1: 透過 Google Apps Script Editor (推薦)

1. **建立 Google 試算表**
   ```
   1. 前往 Google Drive
   2. 建立新的 Google 試算表
   3. 建立 4 個工作表：申請紀錄、申請人與攝影機資料、審核開通、申請人/信箱
   ```

2. **開啟 Apps Script 編輯器**
   ```
   1. 在試算表中點選「擴充功能」→「Apps Script」
   2. 將所有 .js 和 .html 檔案內容複製到編輯器
   3. 確保檔案名稱與專案結構一致
   ```

3. **設定專案屬性**
   ```
   1. 點選「專案設定」
   2. 勾選「在編輯器中顯示 appsscript.json」
   3. 將 appsscript.json 內容複製到編輯器
   ```

4. **部署為網頁應用程式**
   ```
   1. 點選「部署」→「新增部署」
   2. 選擇類型：「網頁應用程式」
   3. 設定：
      - 執行身分：我
      - 具有存取權的使用者：貴機構中的任何人
   4. 點選「部署」並授權必要權限
   5. 複製網頁應用程式 URL
   ```

5. **初始化資料**
   ```
   1. 在試算表中執行自訂選單「系統管理」→「初始化資料」
   2. 設定管理員名單於「審核開通」工作表
   3. 建立申請人與攝影機對應關係
   ```

#### 方法 2: 透過 clasp (命令列工具)

```bash
# 安裝 clasp
npm install -g @google/clasp

# 登入 Google 帳號
clasp login

# Clone 此專案
git clone <repository-url>
cd surveillance_access_apply

# 建立 Apps Script 專案
clasp create --type standalone --title "監視器調閱權限申請系統"

# 推送程式碼
clasp push

# 開啟專案進行部署
clasp open
```

### 環境設定 / Configuration

在 `程式碼.js` 中確認以下設定：

```javascript
// 工作表名稱設定
const DATA_SHEET_NAME = "申請人與攝影機資料";
const LOG_SHEET_NAME = "申請紀錄";
const ADMIN_SHEET_NAME = "審核開通";
const USER_DATA_SHEET_NAME = "申請人/信箱";

// 時區設定 (appsscript.json)
"timeZone": "Asia/Taipei"
```

---

## 📖 使用說明 / Usage

### 一般使用者操作流程

#### 1. 提交申請

1. 開啟系統首頁
2. 點選「申請表單」
3. 填寫申請資訊：
   - 選擇申請人
   - 選擇攝影機地點
   - 選擇開通天數 (1-5天)
   - 填寫調閱事由
4. 點選「送出申請」

#### 2. 查詢申請紀錄

1. 點選「我的申請」
2. 檢視所有個人申請紀錄與審核狀態

### 管理員操作流程

#### 1. 登入審核面板

1. 以管理員帳號登入
2. 點選「審核面板」
3. 系統自動顯示待審核案件

#### 2. 單筆審核

1. 檢視申請詳情
2. 點選「核准」按鈕
3. 系統自動推進至下一階段

#### 3. 批次審核

1. 勾選多個待審核案件
2. 點選「批次核准」
3. 確認操作並送出

### 審核流程說明

```
使用者提交申請
      ↓
   [審核1中]
      ↓
審核1人員核准 → 發送通知給審核2人員
      ↓
   [審核2中]
      ↓
審核2人員核准 → 發送通知給開通人員
      ↓
   [開通中]
      ↓
開通人員執行 → 發送通知給申請人
      ↓
   [已開通]
```

---

## 🔐 權限管理 / Permission Management

### 角色定義 / Role Definitions

| 角色 | 權限 | 操作範圍 |
|-----|------|---------|
| **一般使用者** | 提交申請、查詢個人紀錄 | index.html, 表單.html, myapply.html |
| **審核1人員** | 第一階段審核 | review.html (僅顯示審核1中案件) |
| **審核2人員** | 第二階段審核 | review.html (僅顯示審核2中案件) |
| **開通人員** | 系統開通作業 | review.html (僅顯示開通中案件) |

### 權限檢查機制

#### 前端權限檢查
- 頁面載入時驗證使用者角色
- 動態顯示/隱藏功能按鈕
- 未授權自動導向 unauthorized.html

#### 後端權限檢查
- 所有敏感操作前進行權限驗證
- `checkStagePermission()` 函數驗證審核權限
- 防止跨階段操作與越權存取

#### 管理員設定

在「審核開通」工作表中設定各角色人員：

```
欄位 B: 審核1人員
user1@company.com
user2@company.com

欄位 D: 審核2人員
manager1@company.com
manager2@company.com

欄位 F: 開通人員
admin1@company.com
admin2@company.com
```

---

## 👨‍💻 開發指南 / Development

### 本地開發環境設定

1. **安裝 clasp**
   ```bash
   npm install -g @google/clasp
   ```

2. **Clone 專案並連結**
   ```bash
   git clone <repository-url>
   cd surveillance_access_apply
   clasp login
   # 編輯 .clasp.json 設定 scriptId
   ```

3. **開發流程**
   ```bash
   # 修改本地檔案
   nano 程式碼.js

   # 推送到 Apps Script
   clasp push

   # 開啟線上編輯器
   clasp open
   ```

### 程式碼規範

#### JavaScript 風格
- 使用 `const` 和 `let`，避免 `var`
- 函數命名採用駝峰式 (camelCase)
- 常數命名採用大寫加底線 (UPPER_SNAKE_CASE)
- 保持函數單一職責

#### 註解規範
```javascript
/**
 * 函數功能說明
 * @param {type} paramName - 參數說明
 * @return {type} 回傳值說明
 */
function functionName(paramName) {
  // 實作邏輯
}
```

#### 安全性最佳實踐

1. **輸入驗證**
   ```javascript
   // 使用 sanitizeForSheet 防止公式注入
   const safeName = sanitizeForSheet(userName);
   ```

2. **XSS 防護**
   ```javascript
   // 使用 escapeHtml 防止跨站腳本攻擊
   const safeHtml = escapeHtml(userInput);
   ```

3. **權限檢查**
   ```javascript
   // 每個敏感操作前驗證權限
   if (!checkStagePermission(currentStage)) {
     throw new Error('權限不足');
   }
   ```

### API 函數參考

#### 公開 API (可從前端呼叫)

```javascript
// 取得網頁應用程式 URL
getWebAppUrl()

// 取得申請人與攝影機對應資料
getRequesterData()

// 處理表單提交
processForm(formData)

// 取得使用者申請紀錄
getMyApplications()

// 取得當前使用者待處理任務
getTasksForCurrentUser()

// 處理單筆審核
processApproval(rowNum)

// 處理批次審核
processBatchApproval(rowNumbers)

// 取得使用者顯示名稱
getUserNameByEmail(email)
```

#### 內部函數

```javascript
// 檢查權限
checkStagePermission(stage)

// 取得管理員清單
getAdminEmails()

// 發送通知郵件
sendNotificationEmail(toEmail, stage, data)

// 資料清理
sanitizeForSheet(value)
escapeHtml(value)
```

### 除錯技巧

1. **使用 Logger**
   ```javascript
   Logger.log('Debug info: ' + JSON.stringify(data));
   ```

2. **Apps Script 偵錯工具**
   - 在編輯器中設定中斷點
   - 使用「執行」→「偵錯」功能
   - 查看執行記錄：「檢視」→「執行記錄」

3. **Cloud Logging**
   - 在 Apps Script 專案設定中啟用 Stackdriver
   - 查看詳細執行日誌與錯誤追蹤

---

## 📊 效能最佳化 / Performance Optimization

### 快取策略

```javascript
// 管理員清單快取 (10分鐘)
const CACHE_KEY_ADMINS = 'adminEmails';
const CACHE_DURATION = 600; // 秒

// 使用者名稱快取 (6小時)
const CACHE_DURATION_USER = 21600; // 秒
```

### 批次操作

- 使用 `getValues()` 和 `setValues()` 進行批次讀寫
- 減少 API 呼叫次數
- 避免在迴圈中重複存取試算表

### 資料庫查詢最佳化

```javascript
// 好的做法：一次讀取所有資料
const allData = sheet.getDataRange().getValues();
const filtered = allData.filter(row => row[0] === userEmail);

// 避免：多次單筆查詢
// for (let i = 1; i <= lastRow; i++) {
//   const value = sheet.getRange(i, 1).getValue();
// }
```

---

## 🔒 安全性考量 / Security Considerations

### 實作的安全機制

1. **存取控制**
   - 網域限制：僅限組織內部使用者
   - 角色權限驗證
   - 前後端雙重檢查

2. **資料驗證**
   - 公式注入防護 (`sanitizeForSheet`)
   - XSS 防護 (`escapeHtml`)
   - 輸入資料類型驗證

3. **稽核日誌**
   - 記錄所有審核操作
   - 時間戳記與操作人員
   - 完整狀態變更歷程

4. **資料隱私**
   - 使用者僅能查看自己的申請
   - 管理員僅能看到職責範圍內案件
   - 敏感資料加密傳輸 (HTTPS)

### 建議安全措施

- 定期檢視管理員名單
- 啟用兩步驟驗證 (2FA)
- 定期備份試算表資料
- 監控異常存取行為
- 限制試算表編輯權限

---

## 🔄 版本歷史 / Version History

### v14.0 (Latest)
- ✨ 新增未授權存取頁面 (unauthorized.html)
- ✨ 新增開通天數功能 (1-5天可選)
- 🔧 強化審核工作流程
- 📝 改進使用者體驗

### v13.0
- 🚀 完整三階段審核系統
- 📧 自動化郵件通知機制
- 👥 批次審核功能
- 🔐 角色權限管理
- 💾 快取效能最佳化

### 早期版本
- 基礎申請與審核功能
- Google Sheets 整合
- 使用者介面開發

---

## 📝 授權聲明 / License

此專案為企業內部專有系統，所有權利保留。

This project is proprietary software for internal enterprise use. All rights reserved.

---

## 🙏 致謝 / Acknowledgments

感謝以下技術與平台：

- Google Workspace 團隊
- Google Apps Script 社群
- Bootstrap 開發團隊

---

**最後更新** / **Last Updated**: 2025-11-12
**專案狀態** / **Project Status**: 🟢 Active Development
**總程式碼行數** / **Total Lines of Code**: 1,301
