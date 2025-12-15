/**
 * ===================================================================
 * 園遊會商品訂購系統 - Google Apps Script 後端 (v4.2 - 安全性修正完整版)
 * ===================================================================
 * * 修改說明：
 * 1. 改為使用 Script Properties 讀取 SPREADSHEET_ID 與 ADMIN_EMAILS。
 * 2. 移除舊的 Hardcode 設定，避免 "Identifier has already been declared" 錯誤。
 */

// -------------------
// 【重要】全域設定 (改為從 Script Properties 讀取)
// -------------------

// 1. 取得 Script Properties 物件
const scriptProperties = PropertiesService.getScriptProperties();

// 2. 讀取試算表 ID
// 如果讀取失敗 (例如忘記設定)，這行會變成 null，建議加上錯誤檢查或預設值
const SPREADSHEET_ID = scriptProperties.getProperty('SPREADSHEET_ID');

if (!SPREADSHEET_ID) {
  throw new Error("嚴重錯誤：未設定 SPREADSHEET_ID。請至專案設定 > 指令碼屬性中設定。");
}

// 3. 讀取管理員 Email
// 因為屬性存的是純字串 (例如 "a@g.com,b@g.com")，所以取出來後要用 split(',') 轉回陣列
const ADMIN_EMAILS_RAW = scriptProperties.getProperty('ADMIN_EMAILS');
const ADMIN_EMAILS = ADMIN_EMAILS_RAW ? ADMIN_EMAILS_RAW.split(',') : [];

// -------------------
// 其他固定設定
// -------------------
const ORDER_SHEET_NAME = "訂單紀錄";
const SUMMARY_SHEET_NAME = "商品彙總";
const INVENTORY_SHEET_NAME = "商品庫存"; 

// -------------------
// 網頁服務 (GET 請求)
// -------------------
function doGet(e) {
  try {
    const params = e.parameter;
    const page = params.page;
    const userEmail = getActiveUserEmail();

    if (page === 'admin') {
      // -------------------
      // 載入管理後台 (admin.html)
      // -------------------
      if (isAdmin(userEmail)) {
        return HtmlService.createTemplateFromFile('admin')
          .evaluate()
          .setTitle('訂單管理後台')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
      } else {
        return HtmlService
          .createHtmlOutput(`
            <div style="font-family: Arial, sans-serif; padding: 40px; text-align: center;">
              <h1 style="color: #d93025;">存取遭拒 (403)</h1>
              <p style="font-size: 1.2em;">您沒有權限存取此頁面。</p>
              <p style="color: #5f6368;">請確認您已登入正確的管理員 Google 帳號。</p>
              <hr style="margin: 20px auto; width: 100px; border: 1px solid #ddd;">
              <p style="color: #5f6368;">
                <b>偵測到的目前帳號：</b><br>
                <span style="font-weight: bold; font-size: 1.1em; color: #333;">${userEmail || '未偵測到 (尚未登入?)'}</span>
              </p>
            </div>
          `)
          .setTitle('存取遭拒');
      }
    } else {
      // -------------------
      // 載入訂購頁面 (index.html)
      // -------------------
      const template = HtmlService.createTemplateFromFile('index');
      template.userEmail = userEmail; // 將 Email 傳遞到 HTML 模板
      
      // 【v2 修改】載入商品資料並傳遞到前端
      const productsData = getProducts(); 
      template.productsData = productsData; // 將商品 JSON 傳遞給 HTML
      
      return template.evaluate()
        .setTitle('園遊會商品訂購')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
    }
  } catch (e) {
    Logger.log("doGet 發生嚴重錯誤: " + String(e));
    return HtmlService.createHtmlOutput("伺服器載入頁面時發生錯誤: " + String(e));
  }
}

/**
 * 檢查目前使用者是否為管理員
 */
function isAdmin(email) {
  if (!email) {
    return false;
  }
  return ADMIN_EMAILS.map(adminEmail => adminEmail.toLowerCase()).includes(email.toLowerCase());
}

/**
 * 獲取目前登入者的 Email
 */
function getActiveUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return ""; // 發生錯誤 (例如在匿名模式下)
  }
}


// -------------------
// 前端呼叫的後端函式 (Web API)
// -------------------

/**
 * 【v3 修改】獲取商品列表 (由 index.html 呼叫)
 * - 讀取 "商品庫存" 工作表，只回傳庫存 > 0 的商品
 * - 新增 F 欄 (分類)
 * @returns {String} - JSON 字串，[{ id, name, price, stock, imageUrl, category }, ...]
 */
function getProducts() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(INVENTORY_SHEET_NAME);
    if (!sheet) {
      throw new Error(`伺服器設定錯誤：找不到 "${INVENTORY_SHEET_NAME}" 工作表。`);
    }
    
    // 假設欄位: A=ID, B=名稱, C=價格, D=庫存, E=圖片URL, F=分類
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return JSON.stringify([]); // 沒有商品
    }
    
    // 【v3 修改】讀取 6 欄 (A:F)
    const range = sheet.getRange(2, 1, lastRow - 1, 6); 
    const values = range.getValues();
    
    const products = [];
    values.forEach(row => {
      const id = row[0];
      const name = row[1];
      const price = parseFloat(row[2]);
      const stock = parseInt(row[3], 10);
      const imageUrl = row[4];
      const category = row[5]; // 【v3 新增】
      
      // 確保資料有效，且庫存 > 0
      if (id && name && !isNaN(price) && !isNaN(stock) && stock > 0) {
        products.push({
          id: id,
          name: name,
          price: price,
          stock: stock,
          imageUrl: imageUrl || 'https://placehold.co/400x400/e0e0e0/7f7f7f?text=IMG', // 預設圖片
          category: category || '未分類' // 【v3 新增】
        });
      }
    });
    
    return JSON.stringify(products);
    
  } catch (e) {
    Logger.log("getProducts 錯誤: " + String(e));
    // 即使出錯，也回傳空陣列，避免前端頁面完全崩潰
    return JSON.stringify([]); 
  }
}

/**
 * 【v19 修改】驗證學生身份 (三重檢核 - 錯誤修正)
 * @param {Object} userInfo - 包含 { grade, class, seatNumber, name, studentId, gender }
 * @returns {String} - JSON 字串，{ isValid: true } 或 { isValid: false, message: "..." }
 */
function validateStudent(userInfo) {
  try {
    const { grade, class: classNum, seatNumber, name, studentId, gender } = userInfo;
    
    // 老師/教職員 直接通過
    if (grade === '老師') {
      return JSON.stringify({ isValid: true });
    }

    // 準備比對用的資料
    const targetClass = classNum; // 來自前端，應為 "01", "02" ...
    const targetSeat = parseInt(seatNumber, 10);
    const targetSeatStr = targetSeat.toString();
    const targetName = name.trim();
    const targetStudentId = studentId.trim();
    const targetGender = gender.trim(); // 來自前端 "男" 或 "女"

    // -------------------
    // 【檢核 1：學號 vs 表單】(v19 修正)
    // -------------------
    
    // 1a. 檢查學號長度
    if (targetStudentId.length !== 7) {
      return JSON.stringify({ isValid: false, message: "學號格式驗證失敗 (長度應為7碼)。" });
    }

    // 1b. 檢查年級前綴 (14/13/12)
    const idPrefix = targetStudentId.substring(0, 2);
    let expectedPrefix;
    if (grade === "一") expectedPrefix = "14";
    else if (grade === "二") expectedPrefix = "13";
    else if (grade === "三") expectedPrefix = "12";
    else {
      // 雖然前端會擋，但後端還是要驗證
      throw new Error("無效的年級：" + grade);
    }

    if (idPrefix !== expectedPrefix) {
      return JSON.stringify({ isValid: false, message: "學號與年級驗證失敗。" });
    }
    
    // 1c. 檢查學號中的「性別」
    const idGenderDigit = targetStudentId.substring(2, 3); // 第3碼 (1或2)
    const expectedGenderDigit = (targetGender === "男") ? "1" : "2";
    
    if (idGenderDigit !== expectedGenderDigit) {
      // 【v17 模糊錯誤】
      return JSON.stringify({ isValid: false, message: "學號與性別驗證失败。" });
    }

    // 1d. 檢查學號中的「班級」
    const idClass = targetStudentId.substring(3, 5); // 取得學號中的班級 (e.g., "01")
    if (targetClass !== idClass) {
      // 【v17 模糊錯誤】
      return JSON.stringify({ isValid: false, message: "學號與班級驗證失败。" });
    }
    
    // 1e. 檢查學號中的「座號」
    const idSeat = parseInt(targetStudentId.substring(5, 7), 10).toString(); // 取得學號中的座號 (e.g., "1")
    if (targetSeatStr !== idSeat) {
      // 【v17 模糊錯誤】
      return JSON.stringify({ isValid: false, message: "學號與座號驗證失败。" });
    }

    // --- 檢核 1 (學號 vs 表單) 通過 ---

    // -------------------
    // 【檢核 2：名冊 vs 表單】
    // -------------------
    const sheetName = '國' + grade + '名冊';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      // 這個 throw 會被下面的 catch 捕獲，並回傳 v17 的模糊錯誤
      throw new Error(`伺服器設定錯誤：找不到 "${sheetName}" 工作表。`);
    }

    // 【v16 修改】讀取 A:D 欄 (班級, 座號, 姓名, 性別)
    const data = sheet.getRange("A:D").getValues();
    
    // 遍歷名冊 (跳過 A1, B1, C1, D1 標頭)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // 【v18 健壯性修正】
      if (!row[0] || !row[1] || !row[2] || !row[3]) {
        continue; 
      }

      // 【v18 健壯性修正】安全地讀取資料
      const rowClass = (row[0] || '').toString().trim().padStart(2, '0'); // A欄: 班級
      const rowSeat = parseInt(row[1], 10);                               // B欄: 座號
      const rowName = (row[2] || '').toString().trim();                   // C欄: 姓名
      const rowGender = (row[3] || '').toString().trim();                 // D欄: 性別

      if (isNaN(rowSeat)) {
        continue;
      }
      
      // 比對 班級 和 座號
      if (rowClass === targetClass && rowSeat === targetSeat) {
        
        // 2a. 檢查「姓名」
        if (rowName !== targetName) {
          return JSON.stringify({
            isValid: false,
            message: "姓名驗證失敗。"
          });
        }
        
        // 2b. 檢查「性別」
        if (rowGender !== targetGender) {
           return JSON.stringify({
            isValid: false,
            message: "性別驗證失敗。"
          });
        }
        
        // 所有檢核都通過
        return JSON.stringify({ isValid: true });
      }
    }
    
    // 如果迴圈跑完都沒找到
    return JSON.stringify({
      isValid: false,
      message: "查無此學生資料。"
    });

  } catch (e) {
    Logger.log("validateStudent 錯誤: " + String(e));
    return JSON.stringify({
      isValid: false,
      message: "驗證時發生伺服器錯誤，請稍後再試。"
    });
  }
}


/**
 * 【v2 重大修改】處理訂單 (由 index.html 呼叫)
 * - 增加庫存檢查與扣除
 * @param {Object} formData - 包含訂購人資訊和商品陣列
 * @returns {String} - 回傳 JSON 字串 (成功或失敗)
 */
function processOrder(formData) {
  
  // 【v2 新增】使用鎖定服務，防止多人同時下單造成庫存計算錯誤
  const lock = LockService.getScriptLock();
  try {
    // 嘗試獲取鎖，最多等待 10 秒
    lock.waitLock(10000); 
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const inventorySheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
    const orderSheet = ss.getSheetByName(ORDER_SHEET_NAME);

    if (!inventorySheet) {
      throw new Error(`試算表錯誤：找不到名稱為 "${INVENTORY_SHEET_NAME}" 的工作表。`);
    }
    if (!orderSheet) {
      throw new Error(`試算表錯誤：找不到名稱為 "${ORDER_SHEET_NAME}" 的工作表。`);
    }

    // ---------------------------------
    // 1. 讀取並驗證庫存 (在鎖定狀態下)
    // ---------------------------------
    const inventoryLastRow = inventorySheet.getLastRow();
    if (inventoryLastRow <= 1) {
      throw new Error("商品庫存表為空。");
    }
    // 讀取 A:D (ID, Name, Price, Stock)
    const inventoryRange = inventorySheet.getRange(2, 1, inventoryLastRow - 1, 4);
    const inventoryValues = inventoryRange.getValues();
    
    // 建立一個即時庫存地圖 (Map)
    // 'P01' -> { row: 2, name: '商品A', stock: 10, price: 50 }
    const inventoryMap = new Map();
    inventoryValues.forEach((row, index) => {
      const id = row[0];
      if (id) {
        inventoryMap.set(id, {
          rowNum: index + 2, // 實際列號 (從 2 開始)
          name: row[1],
          price: parseFloat(row[2]),
          stock: parseInt(row[3], 10)
        });
      }
    });

    // 遍歷訂單商品，檢查庫存
    const items = formData.items;
    let stockErrorMessage = "";
    
    for (const item of items) {
      const { id, name, quantity } = item;
      if (!inventoryMap.has(id)) {
        // 商品 ID 不在庫存表中 (可能已下架)
        stockErrorMessage = `商品 "${name}" 已下架或不存在。`;
        break;
      }
      
      const inventoryItem = inventoryMap.get(id);
      if (inventoryItem.stock < quantity) {
        // 庫存不足
        stockErrorMessage = `商品 "${inventoryItem.name}" 庫存不足 (剩餘 ${inventoryItem.stock} 份)。`;
        break;
      }
      
      // 【v2 修改】驗證價格，防止前端竄改
      // 允許一定的誤差 (例如浮點數問題)
      if (Math.abs(inventoryItem.price - item.price) > 0.01) {
        Logger.log(`價格綁定錯誤：ID ${id} 前端價格 ${item.price} vs 後端價格 ${inventoryItem.price}`);
        stockErrorMessage = `商品 "${inventoryItem.name}" 的價格驗證失敗。`;
        break;
      }
    }

    // 如果庫存檢查失敗，立即回傳錯誤
    if (stockErrorMessage) {
      return JSON.stringify({
        status: "error",
        message: stockErrorMessage
      });
    }

    // ---------------------------------
    // 2. 庫存充足，開始處理訂單
    // ---------------------------------
    const timestamp = new Date();
    const orderId = timestamp.getTime(); // 使用時間戳作為唯一的訂單編號
    const totalItems = items.length;
    
    // 準備寫入「訂單紀錄」的資料
    const rowsData = [];
    items.forEach((item, index) => {
      const rowData = [
        orderId,                                          // A: 訂單編號
        timestamp,                                        // B: 訂購時間
        formData.identifier,                              // C: 識別碼 (Email或學號)
        formData.grade,                                   // D: 年級
        formData.class,                                   // E: 班級
        formData.seatNumber,                              // F: 座號
        formData.name,                                    // G: 姓名
        item.name,                                        // H: 商品明細
        item.price,                                       // I: 單價
        item.quantity,                                    // J: 數量
        (index === totalItems - 1) ? formData.totalPrice : "", // K: 總金額
        "未取貨"                                          // L: 取貨狀態
      ];
      rowsData.push(rowData);
    });
    
    // 一次性寫入訂單紀錄 (效能較佳)
    const startRow = orderSheet.getLastRow() + 1; // 取得即將寫入的第一列
    orderSheet.getRange(startRow, 1, rowsData.length, rowsData[0].length).setValues(rowsData);

    // 【v4 修改】強制設定 B 欄 (第 2 欄) 的時間格式
    const timeFormatRange = orderSheet.getRange(startRow, 2, rowsData.length, 1); // (起始列, B欄, 寫入的列數, 1欄)
    timeFormatRange.setNumberFormat("yyyy-mm-dd hh:mm:ss"); // 設定為 年-月-日 時:分:秒

    // ---------------------------------
    // 3. 更新彙總表 (同原邏輯)
    // ---------------------------------
    updateSummary(items, ss); // 傳入 ss 物件避免重複開啟

    // ---------------------------------
    // 4. 扣除庫存 (重要)
    // ---------------------------------
    items.forEach(item => {
      const { id, quantity } = item;
      const inventoryItem = inventoryMap.get(id);
      const newStock = inventoryItem.stock - quantity;
      
      // 更新庫存表 (D 欄)
      inventorySheet.getRange(inventoryItem.rowNum, 4).setValue(newStock);
      
      // 更新 Map (以防同一筆訂單有多個相同商品，雖然前端邏輯應已合併)
      inventoryItem.stock = newStock;
    });

    // ---------------------------------
    // 5. 成功回傳
    // ---------------------------------
    return JSON.stringify({
      status: "success",
      orderId: orderId,
      message: "訂單已成功提交"
    });

  } catch (e) {
    Logger.log("processOrder 錯誤 (可能來自鎖定或庫存檢查): " + String(e));
    return JSON.stringify({
      status: "error",
      message: `伺服器錯誤：${e.message}`
    });
  } finally {
    // 【v2 新增】務必釋放鎖
    lock.releaseLock();
  }
}


/**
 * 【v2 修改】更新彙總表 (由 processOrder 呼叫)
 * @param {Array} items - 訂單商品陣列
 * @param {Spreadsheet} ss - (可選) 傳入的 Spreadsheet 物件
 */
function updateSummary(items, ss) {
  try {
    // 如果沒有傳入 ss 物件，則重新開啟
    const spreadsheet = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    let summarySheet = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME);
    
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet(SUMMARY_SHEET_NAME);
      summarySheet.getRange("A1:D1").setValues([["商品 ID", "商品名稱", "總數量", "總金額"]]);
      summarySheet.getRange("A:D").setHorizontalAlignment("left");
      summarySheet.getRange("C:D").setHorizontalAlignment("right");
      summarySheet.getRange("A1:D1").setFontWeight("bold");
    }
    
    const lastRow = summarySheet.getLastRow();
    const numRows = lastRow - 1;
    let summaryData = []; 
    
    if (numRows > 0) {
      const dataRange = summarySheet.getRange(2, 1, numRows, 4);
      summaryData = dataRange.getValues();
    }
    
    const summaryMap = new Map();
    summaryData.forEach((row, index) => {
      if (row[0]) {
        summaryMap.set(row[0], {
          rowNum: index + 2, 
          name: row[1],
          totalQty: row[2],
          totalAmount: row[3]
        });
      }
    });

    items.forEach(item => {
      const { id, name, price, quantity } = item;
      const amount = price * quantity;
      
      if (summaryMap.has(id)) {
        const existing = summaryMap.get(id);
        const newTotalQty = existing.totalQty + quantity;
        const newTotalAmount = existing.totalAmount + amount;
        
        existing.totalQty = newTotalQty;
        existing.totalAmount = newTotalAmount;
        
        summarySheet.getRange(existing.rowNum, 3, 1, 2).setValues([[newTotalQty, newTotalAmount]]);
        
      } else {
        const newRow = [id, name, quantity, amount];
        summarySheet.appendRow(newRow);
        
        summaryMap.set(id, {
          rowNum: summarySheet.getLastRow(),
          name: name,
          totalQty: quantity,
          totalAmount: amount
        });
      }
    });
    
  } catch (e) {
    Logger.log("updateSummary 錯誤: " + String(e));
  }
}


// -------------------
// 管理後台 API (由 admin.html 呼叫)
// -------------------

/**
 * 檢查呼叫者是否為管理員
 */
function checkAdminAccess() {
  const email = getActiveUserEmail();
  if (!isAdmin(email)) {
    throw new Error("存取遭拒。您必須是管理員。");
  }
  return email;
}


/**
 * 獲取所有訂單資料 (由 admin.html 呼叫)
 * @returns {String} - JSON 字串，包含所有訂單項目
 */
function getOrdersForAdmin() {
  try {
    checkAdminAccess(); 
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDER_SHEET_NAME);
    if (!sheet) {
      return JSON.stringify({ status: "error", message: "找不到訂單紀錄工作表" });
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return JSON.stringify([]); 
    }
    
    const range = sheet.getRange(2, 1, lastRow - 1, 12); 
    const values = range.getValues();

    const orders = values.map((row, index) => {
      return {
        row: index + 2,
        orderId: row[0],
        timestamp: row[1],
        identifier: row[2],
        grade: row[3],
        class: row[4],
        seat: row[5],
        name: row[6],
        productName: row[7],
        price: row[8],
        quantity: row[9],
        totalAmount: row[10],
        status: row[11]
      };
    }).filter(order => order.orderId);

    orders.sort((a, b) => b.orderId - a.orderId);

    return JSON.stringify(orders);
    
  } catch (e) {
    Logger.log("getOrdersForAdmin 錯誤: " + String(e));
    return JSON.stringify({ status: "error", message: e.message });
  }
}


/**
 * 更新取貨狀態 (由 admin.html 呼叫)
 * @param {Number} startRow - 該筆訂單在試算表上的起始列
 * @param {Number} numItems - 該筆訂單包含的商品數量 (列數)
 * @param {String} newStatus - "已取貨" 或 "未取貨"
 * @returns {String} - JSON 字串 (成功或失敗)
 */
function updatePickupStatus(startRow, numItems, newStatus) {
  try {
    checkAdminAccess(); 
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDER_SHEET_NAME);
    if (!sheet) {
      throw new Error("找不到訂單紀錄工作表");
    }
    
    const statusRange = sheet.getRange(startRow, 12, numItems, 1);
    
    const statuses = Array(numItems).fill([newStatus]);
    statusRange.setValues(statuses);
    
    return JSON.stringify({ status: "success", newStatus: newStatus });
    
  } catch (e) {
    Logger.log("updatePickupStatus 錯誤: " + String(e));
    return JSON.stringify({ status: "error", message: e.message });
  }
}


/**
 * 【v2 新增】獲取庫存資料 (由 admin.html 呼叫)
 * @returns {String} - JSON 字串，包含所有商品庫存 (包含 0)
 */
function getInventoryForAdmin() {
  try {
    checkAdminAccess(); 
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(INVENTORY_SHEET_NAME);
    if (!sheet) {
      throw new Error(`伺服器設定錯誤：找不到 "${INVENTORY_SHEET_NAME}" 工作表。`);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return JSON.stringify([]); // 沒有商品
    }
    
    // 讀取 A:D (ID, Name, Price, Stock)
    const range = sheet.getRange(2, 1, lastRow - 1, 4);
    const values = range.getValues();
    
    const inventory = values.map((row, index) => {
      const id = row[0];
      const name = row[1];
      const price = parseFloat(row[2]);
      const stock = parseInt(row[3], 10);
      
      if (id && name && !isNaN(price) && !isNaN(stock)) {
        return {
          rowNum: index + 2, // 實際列號
          id: id,
          name: name,
          price: price,
          stock: stock
        };
      }
      return null;
    }).filter(item => item !== null);
    
    return JSON.stringify({ status: "success", data: inventory });
    
  } catch (e) {
    Logger.log("getInventoryForAdmin 錯誤: " + String(e));
    return JSON.stringify({ status: "error", message: e.message });
  }
}


/**
 * 【v2 新增】更新庫存 (由 admin.html 呼叫)
 * @param {String} productId - 要更新的商品 ID
 * @param {Number} newStock - 新的庫存數量
 * @returns {String} - JSON 字串 (成功或失敗)
 */
function updateInventoryForAdmin(productId, newStock) {
  try {
    checkAdminAccess();
    
    const stock = parseInt(newStock, 10);
    if (isNaN(stock) || stock < 0) {
      throw new Error("無效的庫存數量。必須是大於等於 0 的數字。");
    }
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(INVENTORY_SHEET_NAME);
    if (!sheet) {
      throw new Error(`伺服器設定錯誤：找不到 "${INVENTORY_SHEET_NAME}" 工作表。`);
    }

    // 尋找該 Product ID 所在的列
    const idColumn = sheet.getRange("A:A").getValues();
    let targetRow = -1;
    for (let i = 1; i < idColumn.length; i++) { // 從第 2 列 (index 1) 開始
      if (idColumn[i][0] == productId) {
        targetRow = i + 1; // 陣列 index + 1 = 列號
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`找不到 ID 為 "${productId}" 的商品。`);
    }
    
    // 更新 D 欄 (庫存)
    sheet.getRange(targetRow, 4).setValue(stock);
    
    return JSON.stringify({ status: "success", productId: productId, newStock: stock });
    
  } catch (e) {
    Logger.log("updateInventoryForAdmin 錯誤: " + String(e));
    return JSON.stringify({ status: "error", message: e.message });
  }
}
