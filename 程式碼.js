/**
 * @fileoverview 後端邏輯，支援多層級審核工作流程及批量審核。
 * @version 13.0 - Added Activation Days feature
 */

// --- 全域設定 ---
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const DATA_SHEET_NAME = "申請人與攝影機資料";
const LOG_SHEET_NAME = "申請紀錄";
const ADMIN_SHEET_NAME = "審核開通";
const USER_DATA_SHEET_NAME = "申請人/信箱"; // ✨ 新增：使用者資料工作表名稱

// --- ✨ 更新後的欄位索引設定 ---
const DS_CAMERA_COLUMN_INDEX = 1, DS_REQUESTER_NAME_COLUMN_INDEX = 3;
const LOG_USER_EMAIL_COLUMN_INDEX = 1;     // A欄: 填表帳號
const LOG_TIMESTAMP_COLUMN_INDEX = 2;      // B欄: 申請時間
const LOG_REQUESTER_NAME_COLUMN_INDEX = 3; // C欄: 申請人姓名
const LOG_CAMERA_COLUMN_INDEX = 4;         // D欄: 攝影機地點
const LOG_REASON_COLUMN_INDEX = 5;         // E欄: 調閱事由
const LOG_STATUS_COLUMN_INDEX = 6;         // F欄: 審核狀態
const LOG_APPROVER1_EMAIL_COLUMN_INDEX = 7; // G欄: 審核1帳號
const LOG_APPROVER1_TIME_COLUMN_INDEX = 8;  // H欄: 審核1時間
const LOG_APPROVER2_EMAIL_COLUMN_INDEX = 9; // I欄: 審核2帳號
const LOG_APPROVER2_TIME_COLUMN_INDEX = 10; // J欄: 審核2時間
const LOG_ACTIVATOR_EMAIL_COLUMN_INDEX = 11;// K欄: 開通人帳號
const LOG_ACTIVATOR_TIME_COLUMN_INDEX = 12; // L欄: 開通時間
const LOG_ACTIVATION_DAYS_COLUMN_INDEX = 13;// M欄: ✨ 新增：開通天數
const LOG_SYSTEM_REQUESTER_NAME_COLUMN_INDEX = 14; // ✨ 新增：系統帶入的申請人姓名 (N欄)


// ================================================================
// --- ✨ 重新加入遺失的 formatDataForFrontend 函式 ✨ ---
// ================================================================
/**
 * 通用的資料格式化輔助函式
 * @param {Array} dataRows - 從 Sheet 讀取出的二維陣列資料 (不含表頭)。
 * @returns {Array} 格式化處理過的二維陣列。
 */
function formatDataForFrontend(dataRows) {
  return dataRows.map(row => {
    return row.map((cell, index) => { // 'index' 是從 0 開始的欄位索引
      if (cell instanceof Date) {
        // 根據欄位索引決定格式
        switch (index + 1) { // 轉換為 1-based 索引來比對
          case LOG_TIMESTAMP_COLUMN_INDEX:
          case LOG_APPROVER1_TIME_COLUMN_INDEX:
          case LOG_APPROVER2_TIME_COLUMN_INDEX:
          case LOG_ACTIVATOR_TIME_COLUMN_INDEX:
            return Utilities.formatDate(cell, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
          default:
            // 對於其他未知的日期欄位，給予一個預設的日期格式
            return Utilities.formatDate(cell, "Asia/Taipei", "yyyy/MM/dd");
        }
      }
      return cell;
    });
  });
}


/// ================================================================
// --- ✨ 權限管理重構 (分層) ---
// ================================================================
function getApprover1Emails() { return getAdminListFromSheet('B'); }
function getApprover2Emails() { return getAdminListFromSheet('D'); }
function getActivatorEmails() { return getAdminListFromSheet('F'); }

function getAdminListFromSheet(column) {
  const cacheKey = `admin_list_${column}`;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) { return JSON.parse(cached); }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    const range = sheet.getRange(`${column}2:${column}${sheet.getLastRow()}`);
    const emails = range.getValues().map(row => row[0]).filter(String);
    cache.put(cacheKey, JSON.stringify(emails), 600); // 快取 10 分鐘
    return emails;
  } catch(e) { return []; }
}


/**
 * ✨ 新增：根據 Email 高效率查詢對應的使用者名稱。
 * @param {string} email - 要查詢的使用者 Email。
 * @returns {string} 對應的使用者名稱，若找不到則回傳原始 Email。
 */
function getUserNameByEmail(email) {
  if (!email) return "未知帳號";

  const cacheKey = 'user_name_map';
  const cache = CacheService.getScriptCache();
  const cachedMap = cache.get(cacheKey);

  let userNameMap;
  if (cachedMap) {
    userNameMap = JSON.parse(cachedMap);
  } else {
    // 如果快取中沒有，則從試算表讀取並建立 Map
    userNameMap = {};
    try {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USER_DATA_SHEET_NAME);
      // 從第二列開始讀取，避免讀到標題
      if (sheet.getLastRow() > 1) {
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
        data.forEach(row => {
          const name = row[0]; // A欄: 使用者名稱
          const userEmail = row[1]; // B欄: 登入者帳號
          if (userEmail) {
            userNameMap[userEmail.toLowerCase().trim()] = name;
          }
        });
      }
      // 將建立好的 Map 存入快取，有效期為 6 小時
      cache.put(cacheKey, JSON.stringify(userNameMap), 21600);
    } catch (e) {
      Logger.log(`讀取使用者資料表 (${USER_DATA_SHEET_NAME}) 失敗: ${e.toString()}`);
      // 即使讀取失敗，也回傳原始 email，避免流程中斷
      return email;
    }
  }

  // 從 Map 中查詢並回傳結果
  const foundName = userNameMap[email.toLowerCase().trim()];
  return foundName || email; // 如果找不到對應名稱，則回傳原始 email 作為備用
}

// ================================================================
// --- ✨ 核心工作流程引擎 ✨ ---
// ================================================================

/**
 * [前端可呼叫] 根據當前使用者身分，獲取其待辦事項。
 */
function getTasksForCurrentUser() {
  // ✨ 修正：在函式開頭加入權限檢查，確保只有審核者能呼叫
  const userEmail = Session.getActiveUser().getEmail();
  if (!isUserAnAdmin(userEmail)) {
      Logger.log(`[權限警告] 非管理員 ${userEmail} 嘗試呼叫 getTasksForCurrentUser()`);
      return { header: [], records: [] }; // 回傳空資料
  }
  const allData = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME).getDataRange().getValues();
  const originalHeader = allData[0];
  let tasks = [];

   // --- ✨ 定義要顯示的欄位索引 (0-based) ---
  // B, C, D, E, F, H, K, M
  const desiredColumnIndexes = [
    LOG_TIMESTAMP_COLUMN_INDEX - 1,      // B欄: 申請時間
    LOG_REQUESTER_NAME_COLUMN_INDEX - 1, // C欄: 申請人姓名
    LOG_CAMERA_COLUMN_INDEX - 1,         // D欄: 攝影機地點
    LOG_REASON_COLUMN_INDEX - 1,         // E欄: 調閱事由
    LOG_STATUS_COLUMN_INDEX - 1,         // F欄: 審核狀態
    LOG_APPROVER1_TIME_COLUMN_INDEX - 1, // H欄: 審核1時間
    LOG_APPROVER2_TIME_COLUMN_INDEX - 1, // I欄: 審核2時間
    LOG_ACTIVATION_DAYS_COLUMN_INDEX - 1 // M欄: 開通天數
  ];

  // --- ✨ 根據索引篩選表頭 ---
  const filteredHeader = desiredColumnIndexes.map(index => originalHeader[index]);

  const isApprover1 = getApprover1Emails().includes(userEmail);
  const isApprover2 = getApprover2Emails().includes(userEmail);
  const isActivator = getActivatorEmails().includes(userEmail);

  const dataRows = allData.slice(1);
  if (isApprover1) {
    tasks = tasks.concat(filterTasksByStatus(dataRows, "審核1中"));
  }
  if (isApprover2) {
    tasks = tasks.concat(filterTasksByStatus(dataRows, "審核2中"));
  }
  if (isActivator) {
    tasks = tasks.concat(filterTasksByStatus(dataRows, "開通中"));
  }
  
  const formattedTasks = tasks.map(task => {
    const formattedData = formatDataForFrontend([task.data])[0];
    // --- ✨ 根據索引篩選資料列 ---
    const filteredData = desiredColumnIndexes.map(index => formattedData[index]);
    return { rowNum: task.rowNum, data: filteredData };
});

return { header: filteredHeader, records: formattedTasks };
}

function filterTasksByStatus(dataRows, status) {
  const tasks = [];
  dataRows.forEach((row, index) => {
    if (row[LOG_STATUS_COLUMN_INDEX - 1] === status) {
      tasks.push({ rowNum: index + 2, data: row });
    }
  });
  return tasks;
}

/**
 * [前端可呼叫] 處理審核動作的核心函式。
 * @param {number} rowNum - 要處理的紀錄在 Sheet 中的列號。
 */
function processApproval(rowNum) {
  const userEmail = Session.getActiveUser().getEmail();
  const logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
  const targetRow = logSheet.getRange(rowNum, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const currentStatus = targetRow[LOG_STATUS_COLUMN_INDEX - 1];
  const requesterEmail = targetRow[LOG_USER_EMAIL_COLUMN_INDEX - 1];
  const requesterName = targetRow[LOG_REQUESTER_NAME_COLUMN_INDEX - 1];
  const camera = targetRow[LOG_CAMERA_COLUMN_INDEX - 1];
  const activationDays = targetRow[LOG_ACTIVATION_DAYS_COLUMN_INDEX - 1]; // ✨ 新增：讀取開通天數

  let notifyEmails = [];
  let subject = "";
  let body = "";

  switch(currentStatus) {
    case "審核1中":
      if (!getApprover1Emails().includes(userEmail)) return "權限不足";
      logSheet.getRange(rowNum, LOG_APPROVER1_EMAIL_COLUMN_INDEX).setValue(userEmail);
      logSheet.getRange(rowNum, LOG_APPROVER1_TIME_COLUMN_INDEX).setValue(new Date());
      logSheet.getRange(rowNum, LOG_STATUS_COLUMN_INDEX).setValue("審核2中");
      
      notifyEmails = getApprover2Emails();
      subject = `[待審核] 影像調閱申請需要您進行第二階段審核`;
      body = `申請人 ${requesterName} 的影像調閱申請 (地點: ${camera}) 已通過第一階段審核，需要您進行第二階段審核。`;
      break;
    
    case "審核2中":
      if (!getApprover2Emails().includes(userEmail)) return "權限不足";
      logSheet.getRange(rowNum, LOG_APPROVER2_EMAIL_COLUMN_INDEX).setValue(userEmail);
      logSheet.getRange(rowNum, LOG_APPROVER2_TIME_COLUMN_INDEX).setValue(new Date());
      logSheet.getRange(rowNum, LOG_STATUS_COLUMN_INDEX).setValue("開通中");

      notifyEmails = getActivatorEmails();
      subject = `[待開通] 影像調閱申請已審核通過`;
      body = `申請人 ${requesterName} 的影像調閱申請 (地點: ${camera}) 已通過兩階段審核，申請開通天數為【${activationDays}】天，請您進行影像開通作業。`; // ✨ 更新：在信件內文中加入天數
      break;

      case "開通中":
        if (!getActivatorEmails().includes(userEmail)) return "權限不足";
        logSheet.getRange(rowNum, LOG_ACTIVATOR_EMAIL_COLUMN_INDEX).setValue(userEmail);
        logSheet.getRange(rowNum, LOG_ACTIVATOR_TIME_COLUMN_INDEX).setValue(new Date());
        logSheet.getRange(rowNum, LOG_STATUS_COLUMN_INDEX).setValue("已開通");
  
        notifyEmails = [requesterEmail]; // 只通知申請人
        subject = `[通知] 您的影像調閱申請已開通`;
        body = `您好，您申請的影像調閱 (地點: ${camera}) 已處理完畢並開通，開通天數為【${activationDays}】天。`;
        
        // ✨ 執行郵件通知 (帶上自訂連結)
        if (notifyEmails.length > 0) {
          sendNotificationEmail(notifyEmails.join(','), subject, body, { 
            linkPage: 'myapply', 
            linkText: '點此查看我的申請紀錄' 
          });
        }
        return "操作成功！"; // ✨ 提早結束，避免重複發信
  
      default:
        return "此案件狀態不正確或已處理。";
    }
  
    // ✨ 其他狀態的通用郵件通知 (使用預設連結)
    if (notifyEmails.length > 0) {
      sendNotificationEmail(notifyEmails.join(','), subject, body);
    }
    
    return "操作成功！";
  }

/**
 * [前端可呼叫] ✨ 新增：處理批量審核動作。
 * @param {Array<number>} rowNumbers - 要處理的紀錄在 Sheet 中的列號陣列。
 * @returns {String} 批次處理的結果摘要。
 */
function processBatchApproval(rowNumbers) {
  // 權限檢查：確保執行者至少是一個審核者
  const userEmail = Session.getActiveUser().getEmail();
  if (!isUserAnAdmin(userEmail)) {
    return "權限不足，操作失敗。";
  }

  if (!rowNumbers || !Array.isArray(rowNumbers) || rowNumbers.length === 0) {
    return "沒有選擇任何項目。";
  }

  let successCount = 0;
  let failCount = 0;
  let errorMessages = [];

  // 遍歷所有傳入的列號
  rowNumbers.forEach(rowNum => {
    // ✨ 核心邏輯：重複利用現有的單筆審核函式
    const result = processApproval(rowNum);
    if (result === "操作成功！") {
      successCount++;
    } else {
      failCount++;
      // 記錄下失敗的原因（例如：權限不足、狀態不符等）
      errorMessages.push(`第 ${rowNum} 列: ${result}`);
    }
  });

  let summary = `批次處理完成。\n成功：${successCount} 筆。\n失敗/略過：${failCount} 筆。`;
  if (failCount > 0) {
    summary += `\n\n失敗原因摘要：\n${errorMessages.join('\n')}`;
  }
  
  return summary;
}


/**
 * ✨ 新增：中央權限檢查函式
 * 檢查使用者是否為任何層級的審核者。
 * @param {string} userEmail - 要檢查的使用者 Email。
 * @returns {boolean} 如果是審核者則回傳 true，否則回傳 false。
 */
function isUserAnAdmin(userEmail) {
  if (!userEmail) return false;
  const list1 = getApprover1Emails();
  const list2 = getApprover2Emails();
  const list3 = getActivatorEmails();
  // 只要 email 出現在任何一個列表中，就認定為管理員
  return list1.includes(userEmail) || list2.includes(userEmail) || list3.includes(userEmail);
}



/**
 * 根據 URL 參數決定顯示哪個 HTML 頁面。
 */
function doGet(e) {
  const page = e.parameter.page;
  
  if (page === 'apply') {
    return HtmlService.createTemplateFromFile('表單').evaluate()
        .setTitle('駐站攝影機影像調閱申請');
  }
  if (page === 'myapply') {
    return HtmlService.createTemplateFromFile('myapply').evaluate()
        .setTitle('我的申請紀錄');
  }
  if (page === 'review') {
    const userEmail = Session.getActiveUser().getEmail();
    // ✨ 使用新的中央權限檢查函式
    if (isUserAnAdmin(userEmail)) {
      return HtmlService.createTemplateFromFile('review').evaluate().setTitle('審核面板');
    } else {
      return HtmlService.createTemplateFromFile('unauthorized').evaluate().setTitle('權限不足');
    }
  }
  
  // ✨ 如果沒有任何 page 參數，預設顯示入口網站 index.html
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('駐站攝影機影像調閱申請');
}

/**
 * [前端可呼叫] ✨ 新增：獲取當前 Web App 的部署 URL。
 * @returns {String} Web App 的 URL。
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ================================================================
// --- ✨ 修正後的 getRequesterData 函式 ✨ ---
// ================================================================

/**
 * [前端可呼叫] 獲取申請人與其對應的攝影機資料。
 * @returns {Object} 一個物件，key 是申請人名稱，value 是該申請人擁有的攝影機陣列。
 */
function getRequesterData() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log(`警告：工作表 "${DATA_SHEET_NAME}" 不存在或沒有資料。`);
      return {};
    }
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const data = dataRange.getValues();
    
    const requesterMap = {};
    
    data.forEach(row => {
      const camera = row[DS_CAMERA_COLUMN_INDEX - 1];
      const requester = row[DS_REQUESTER_NAME_COLUMN_INDEX - 1];
      
      if (camera && requester) {
        if (!requesterMap[requester]) {
          requesterMap[requester] = [];
        }
        requesterMap[requester].push(camera);
      }
    });
    
    // ✨ --- 關鍵修正：回傳處理過的 requesterMap，而不是原始的 requesters 陣列 --- ✨
    return requesterMap;

  } catch (e) {
    Logger.log("讀取申請人資料失敗: " + e.toString() + e.stack);
    return { "錯誤": ["讀取申請人資料失敗，請檢查後台日誌"] };
  }
}

/**
 * 輔助函式：清理使用者輸入，防止公式注入。
 * @param {*} input - 使用者輸入的資料。
 * @returns {*} 清理後的資料。
 */
function sanitizeForSheet(input) {
  if (typeof input !== 'string') {
    return input;
  }
  // 如果字串以 =, +, -, @ 開頭，則在前面加上一個單引號
  if (/^(=|\+|-|@)/.test(input)) {
    return "'" + input;
  }
  return input;
}

/**
 * [前端可呼叫] ✨ 更新後的表單處理函式
 */
function processForm(formData) {
  try {
    const logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) { throw new Error("無法獲取使用者帳號"); }

    // ✨ 新增：呼叫查詢函式，獲取使用者真實姓名
    const systemUserName = getUserNameByEmail(userEmail);

    const newRow = [];
    newRow[LOG_USER_EMAIL_COLUMN_INDEX - 1] = userEmail;
    newRow[LOG_TIMESTAMP_COLUMN_INDEX - 1] = new Date();
    newRow[LOG_REQUESTER_NAME_COLUMN_INDEX - 1] = formData.requester;
    newRow[LOG_CAMERA_COLUMN_INDEX - 1] = formData.camera;
    newRow[LOG_REASON_COLUMN_INDEX - 1] = formData.reason;
    newRow[LOG_STATUS_COLUMN_INDEX - 1] = "審核1中";
    newRow[LOG_ACTIVATION_DAYS_COLUMN_INDEX - 1] = formData.activationDays; // ✨ 新增：寫入開通天數
    // ✨ 新增：將系統查詢到的姓名寫入 N 欄
    newRow[LOG_SYSTEM_REQUESTER_NAME_COLUMN_INDEX - 1] = systemUserName;

    logSheet.appendRow(newRow);

    // 觸發給第一審核人的通知
    const notifyEmails = getApprover1Emails();
    if(notifyEmails.length > 0) {
      const subject = "[新申請] 您有一筆新的影像調閱申請待審核";
      const body = `申請人 ${formData.requester} 提出了一筆新的影像調閱申請 (地點: ${formData.camera})，需要您進行第一階段審核。`;
      sendNotificationEmail(notifyEmails.join(','), subject, body);
    }
    return "申請提交成功！";
  } catch (e) { return "錯誤：" + e.message; }
}



/**
 * [前端可呼叫] 獲取所有待處理的申請紀錄給管理員。(v7.1 更新)
 */
/**
 * [前端可呼叫] 獲取所有待處理的申請紀錄給管理員。
 */
function getPendingApplications() {
  // ✨ 雙重保險：同樣使用新的函式檢查權限
  const userEmail = Session.getActiveUser().getEmail();
  if (!getAdminEmails().includes(userEmail)) {
    Logger.log(`[權限警告] 非管理員 ${userEmail} 嘗試呼叫 getPendingApplications()`);
    return null;
  }
  try {
    const logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
    if (logSheet.getLastRow() < 2) { return { header: [], records: [] }; }

    const allData = logSheet.getDataRange().getValues();
    const header = allData[0];
    const pendingApps = [];

    for (let i = 1; i < allData.length; i++) {
      const rowData = allData[i];
      if (rowData[LOG_STATUS_COLUMN_INDEX - 1] === "申請中") {
        pendingApps.push({
          rowNum: i + 1,
          data: rowData 
        });
      }
    }
    
    // ✨ 在打包前，先對所有待處理的紀錄進行格式化
    const formattedPendingApps = pendingApps.map(app => {
      // formatDataForFrontend 需要二維陣列，所以我們傳入 [app.data] 並取回第一個結果
      const formattedData = formatDataForFrontend([app.data])[0];
      return {
        rowNum: app.rowNum,
        data: formattedData
      };
    });

    return { header: header, records: formattedPendingApps };
  } catch (e) {
    Logger.log("getPendingApplications 失敗: " + e.toString() + e.stack);
    return null;
  }
}



/**
 * [前端可呼叫] 更新指定申請紀錄的狀態。
 * @param {Array<number>} rowNumbers 要更新狀態的列號陣列。
 * @returns {String} 執行結果訊息。
 */
function updateApplicationStatus(rowNumbers) {
  // 1. 安全性檢查：確認執行者是否為管理員
  const userEmail = Session.getActiveUser().getEmail();
  if (!isUserAnAdmin(userEmail)) {
    Logger.log(`[權限警告] 非管理員 ${userEmail} 嘗試呼叫舊的 updateApplicationStatus()`);
    return "權限不足，操作失敗。";
  }

  // 2. 輸入驗證：檢查前端傳來的參數是否有效
  if (!rowNumbers || !Array.isArray(rowNumbers) || rowNumbers.length === 0) {
    return "沒有選擇任何項目，或傳入的資料格式不正確。";
  }

  try {
    // 3. 執行更新操作
    const logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
    
    rowNumbers.forEach(rowNum => {
      // 將 rowNum 轉換為整數以防萬一
      const row = parseInt(rowNum, 10);
      if (row > 1) { // 避免意外修改到表頭 (row 1)
        // 使用我們定義好的常數，精準定位要更新的欄位 (F欄)
        logSheet.getRange(row, LOG_STATUS_COLUMN_INDEX).setValue("已調閱");
      }
    });

    SpreadsheetApp.flush(); // 強制伺服器立即寫入所有變更

    // 4. 回傳成功訊息
    return "成功更新 " + rowNumbers.length + " 筆紀錄的狀態為「已調閱」！";

  } catch (e) {
    // 5. 錯誤處理
    Logger.log("更新狀態時發生錯誤: " + e.toString() + e.stack);
    return "更新失敗，請檢查後台日誌。";
  }
}


/**
 * [前端可呼叫] 獲取當前登入使用者的所有申請紀錄。(v7.1 更新)
 */
function getMyApplications() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) { return [[]]; } // 回傳空陣列以避免前端錯誤

    // ✨ 新增：根據 Email 獲取使用者在「使用者資料」表中登記的姓名
    const userName = getUserNameByEmail(userEmail);


    const logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
    const allData = logSheet.getDataRange().getValues();
    const originalHeader = allData[0];

    // --- ✨ 定義要顯示的欄位索引 (0-based) ---
    // B=2, C=3, D=4, E=5, F=6, M=13
    const desiredColumnIndexes = [
      LOG_SYSTEM_REQUESTER_NAME_COLUMN_INDEX - 1, //A欄: 申請人姓名
      LOG_TIMESTAMP_COLUMN_INDEX - 1,      // B欄: 申請時間
      LOG_REQUESTER_NAME_COLUMN_INDEX - 1, // C欄: 申請人姓名
      LOG_CAMERA_COLUMN_INDEX - 1,         // D欄: 攝影機地點
      LOG_REASON_COLUMN_INDEX - 1,         // E欄: 調閱事由
      LOG_STATUS_COLUMN_INDEX - 1,         // F欄: 審核狀態
      LOG_ACTIVATOR_TIME_COLUMN_INDEX - 1, // L欄: 開通時間
      LOG_ACTIVATION_DAYS_COLUMN_INDEX - 1 // M欄: 開通天數
    ];

    // --- ✨ 根據索引篩選表頭 ---
    const filteredHeader = desiredColumnIndexes.map(index => originalHeader[index]);

    // 如果工作表只有表頭，直接回傳篩選過的表頭
    if (logSheet.getLastRow() < 2) {
      return [filteredHeader];
    }

    // --- ✨ 核心修改：擴充 filter 的篩選條件 ---
    const userRecords = allData.slice(1)
      .filter(row => {
        const rowSubmitterEmail = row[LOG_USER_EMAIL_COLUMN_INDEX - 1]; // A 欄: 實際申請操作帳號
        const rowRequesterName = row[LOG_REQUESTER_NAME_COLUMN_INDEX - 1];  // C 欄: 申請人姓名

        // 條件：只要「操作帳號」是本人，或者「申請人姓名」是本人，就納入結果
        return rowSubmitterEmail === userEmail || rowRequesterName === userName;
      })
      .map(row => {
        // 維持您原有的欄位提取和日期格式化邏輯
        return desiredColumnIndexes.map(index => {
          const cell = row[index];
          if ((index === LOG_TIMESTAMP_COLUMN_INDEX - 1 || index === LOG_ACTIVATOR_TIME_COLUMN_INDEX - 1) && cell instanceof Date && !isNaN(cell)) {
            return Utilities.formatDate(cell, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
          }
          return cell;
        });
      });

    // --- ✨ 回傳篩選並格式化後的資料 ---
    return [filteredHeader, ...userRecords];

  } catch (e) {
    Logger.log("getMyApplications 發生錯誤: " + e.toString() + e.stack);
    // 回傳 null 可能導致前端 JS 錯誤，回傳帶有錯誤訊息的結構更安全
    return [['錯誤'], ['讀取申請紀錄失敗，請聯繫管理員。']];
  }
}


/**
* ✨ 通用郵件通知函式 (v14.0 更新，支援自訂連結)
 * @param {string} recipients - 收件者 Email，多個請用逗號分隔。
 * @param {string} subject - 郵件主旨。
 * @param {string} body - 郵件內文 (HTML 格式)。
 * @param {object} [options] - 可選參數物件。
 * @param {string} [options.linkPage='review'] - 連結指向的頁面 (例如 'myapply')。
 * @param {string} [options.linkText='點此前往審核儀表板'] - 連結的顯示文字。
 */
function sendNotificationEmail(recipients, subject, body, options = {}) {
  const { linkPage = 'review', linkText = '點此前往審核儀表板' } = options;
  
  const webAppUrl = ScriptApp.getService().getUrl();
  const linkUrl = `${webAppUrl}?page=${linkPage}`;

  const htmlBody = `
    <p>${body}</p>
    <p>您可以點擊以下連結進入系統查看：</p>
    <p><a href="${linkUrl}">${linkText}</a></p>
    <br>
    <p><i>此為系統自動發送的郵件。</i></p>
  `;
  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    htmlBody: htmlBody
  });
}


// =================================================================
// ✨ 新增功能：Google Sheet 自訂選單 ✨
// =================================================================

/**
 * 當試算表被開啟時，自動執行的函式，用於建立自訂選單。
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🖥️ 監視器調閱系統') // 主選單名稱
      .addItem('🔗 開啟系統入口網站', 'openPortalPage')
      .addSeparator()
      .addItem('➤ 新增調閱申請', 'openApplyPage')
      .addItem('➤ 查詢我的申請', 'openMyApplyPage')
      .addSeparator()
      .addItem('🛡️ 管理員審核面板', 'openReviewPage')
      .addToUi();
}

/**
 * 輔助函式：在對話方塊中開啟入口網站頁面 (index.html)
 */
function openPortalPage() {
  const html = HtmlService.createTemplateFromFile('index').evaluate()
      .setWidth(800)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '駐站攝影機調閱申請');
}

/**
 * 輔助函式：在對話方塊中開啟新增申請頁面 (表單.html)
 */
function openApplyPage() {
  const html = HtmlService.createTemplateFromFile('表單').evaluate()
      .setWidth(650)
      .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '新增調閱申請');
}

/**
 * 輔助函式：在對話方塊中開啟我的申請紀錄頁面 (myapply.html)
 */
function openMyApplyPage() {
  const html = HtmlService.createTemplateFromFile('myapply').evaluate()
      .setWidth(950)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '我的申請紀錄');
}

/**
 * 輔助函式：在對話方塊中開啟管理員審核頁面 (review.html)
 */
function openReviewPage() {
  const html = HtmlService.createTemplateFromFile('review').evaluate()
      .setWidth(1250)
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '審核面板');
}

/**
 * 通用的開啟 Web App 頁面函式。
 * @param {string | null} pageName - 要在 URL 中指定的 page 參數，若為首頁則為 null。
 */
function openWebAppPage(pageName) {
  let url = ScriptApp.getService().getUrl();
  if (pageName) {
    url += '?page=' + pageName;
  }

  // 產生一個包含 JavaScript 的迷你 HTML，用來在新分頁開啟網址並自動關閉彈窗
  const html = `
    <script>
      window.open("${url}", "_blank");
      google.script.host.close();
    </script>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(100)
      .setHeight(50);
      
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '開啟中...');
}