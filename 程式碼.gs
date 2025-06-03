// 🔧 詮栩 CNC 雲端系統 - Google Apps Script 完整整合版
// 功能涵蓋：刀具入庫、查詢、加工登記與查詢，支援 RESTful API 結構與 JSON 資料格式

/**
 * 🌐 主函式：處理 GET 請求
 * 說明：根據 URL 參數中的 page 決定執行哪一個功能模組，並回傳對應資料。
 * 支援功能：刀具入庫、刀具查詢、加工登記頁面下拉資料、加工數量查詢
 */
function doGet(e) {
  const page = e.parameter.page;
  const keyword = e.parameter.keyword;

  if (page === 'inventory') return getDropdownData();
  if (page === 'check') return handleToolCheck(keyword, e.parameter);
  if (page === 'production_entry') {
    if (e.parameter.type === 'last_submitted') return getLastSubmittedProduction();
    return handleProductionEntryLoad(e.parameter);
  }
  if (page === 'production_query') {
    const type = e.parameter.type;
    if (e.parameter.mode === 'latest' && e.parameter.pd && e.parameter.od) {
      return getLatestProductionByProduct(e.parameter.pd, e.parameter.od);
    }
    if (type === 'export_data') return exportProductionData(e.parameter.pd, e.parameter.od);
    if (type === 'dropdown') return getDropdownValues();
    return getProductionQuery(e.parameter);
  }
  return ContentService.createTextOutput(JSON.stringify({ error: "無效的 page 參數" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 📬 主函式：處理 POST 請求
 * 說明：根據前端送來的資料內容，自動分派功能給「刀具入庫」、「加工登記」、「刪除最後一筆加工資料」
 * 使用 JSON 格式傳遞資料，前端用 fetch 傳送 body 即可
 */
function doPost(e) {
  const data = JSON.parse(e.postData.contents); // 將前端送來的 JSON 字串解析成物件
  // ✅ 判斷為「刀具入庫」功能（需有 category 與 spec 欄位）
  if (data.category && data.spec) {
    return handleToolInventoryPost(data);  // 執行刀具入庫處理邏輯（寫入入庫表、更新庫存）
  }
  // ✅ 判斷為「加工數量登記」功能（需有機台與產品編號）
  else if (data.mc && data.pd) {
    return addProductionData(data); // 寫入加工數據，含檢查是否重複、加上時間戳
  }
  // ✅ 判斷為「刪除最後一筆加工紀錄」功能（需明確指定 action 欄位）
  else if (data.action === 'deleteLastProduction') {
    return deleteLastProductionRow(); // 找出最後一筆資料並刪除
  }
  // ⚠️ 如果傳入格式不符合任何一種，回傳失敗（前端可接收 { success: false } 處理錯誤提示）
  return ContentService.createTextOutput(JSON.stringify({ success: false }))
    .setMimeType(ContentService.MimeType.JSON); // 回傳格式為 JSON
}

/**
 * 📥 getDropdownData：取得刀具入庫頁面所需下拉選單資料
 * 說明：從「data」工作表取得 類別、廠牌、廠商，
 *      並從「刀具庫存表」中取得所有品名規格（自動去重與過濾空白）
 * 回傳格式：{ categories:[], brands:[], vendors:[], specs:[] }（JSON）
 */
function getDropdownData() {
  // 取得「data」工作表（內含類別、廠牌、廠商資料）
  const sheet = SpreadsheetApp.getActive().getSheetByName("data");
  // 取得整張表的所有資料，並略過第1列（標題列）
  const values = sheet.getDataRange().getValues().slice(1);
  // 使用 Set 結構儲存不重複的選項（自動去重）
  const categories = new Set(), brands = new Set(), vendors = new Set();
  // 遍歷每一列資料，依照欄位加入 Set 中
  values.forEach(r => {
    if (r[0]) categories.add(r[0]); // A欄：類別
    if (r[1]) brands.add(r[1]);     // B欄：廠牌
    if (r[2]) vendors.add(r[2]);    // C欄：廠商
  });
  // 再從「刀具庫存表」抓取 C 欄（第 3 欄）的品名規格資料
  const specSheet = SpreadsheetApp.getActive().getSheetByName("刀具庫存表");
  let specs = [];
  if (specSheet) {
    // 取得第2列起所有品名規格欄的值（第3欄），忽略空白與重複
    const specValues = specSheet.getRange(2, 3, specSheet.getLastRow() - 1 || 1).getValues();
    specs = [...new Set(specValues.flat().filter(v => v))]; // 展平、過濾空白、去重
  }
  // 最後將所有下拉選單資料組成 JSON 格式並回傳給前端
  return ContentService.createTextOutput(JSON.stringify({
    categories: [...categories], // 類別選單
    brands: [...brands],         // 廠牌選單
    vendors: [...vendors],       // 廠商選單
    specs                        // 品名規格選單
  })).setMimeType(ContentService.MimeType.JSON); // 指定回傳格式為 JSON
}

/**
 * 🔍 handleToolCheck：處理刀具查詢請求
 * 說明：依照傳入的查詢參數，自動分派至：
 *       1. 入庫明細查詢（searchInDetailSheet）
 *       2. 庫存模糊查詢（searchStock）
 *       3. 類別查詢（getStockByCategory）
 *       4. 顯示可選的類別（getDropdownData2）
 * 傳入參數：
 *   - keyword：關鍵字（模糊比對用）
 *   - params：其他 GET 參數（target, category...）
 * 回傳：查詢結果 JSON（透過 ContentService 傳回）
 */
function handleToolCheck(keyword, params) {
  // ✅ 若有 keyword 且查詢目標為 detail，代表是查「入庫明細表」
  if (keyword && params.target === 'detail') {
    return ContentService.createTextOutput(JSON.stringify(searchInDetailSheet(keyword))) // 回傳明細查詢結果
      .setMimeType(ContentService.MimeType.JSON); // 指定格式為 JSON
  }
  // ✅ 若只有 keyword（但沒有 detail），代表查的是「刀具庫存表」模糊比對
  else if (keyword) {
    return ContentService.createTextOutput(JSON.stringify(searchStock(keyword))) // 回傳庫存查詢結果
      .setMimeType(ContentService.MimeType.JSON); // 指定格式為 JSON
  }
  // ✅ 若傳入的是 category（例如「銑刀」），代表做分類查詢
  else if (params.category) {
    return getStockByCategory(params.category); // 依類別回傳庫存資料
  }
  // ✅ 若沒傳入 keyword 也沒傳入 category，就只回傳所有類別下拉選項
  else {
    return getDropdownData2(); // 顯示分類選單用（給前端下拉選）
  }
}

/**
 * 📂 getDropdownData2：取得刀具查詢頁面用的「類別下拉選單」資料
 * 說明：從「data」工作表中擷取第 1 欄（類別），去除重複後組成選單資料
 * 回傳格式：{ categories: [...] }（JSON）
 */
function getDropdownData2() {
  // 取得目前 Spreadsheet 中名為「data」的工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
  // 抓取整張表的資料，並略過標題列（第一列）
  const data = sheet.getDataRange().getValues().slice(1);
  // 建立一個 Set 結構來儲存不重複的類別項目
  const categories = new Set();
  // 遍歷每一列資料，將 A 欄（類別欄）不為空的值加入 Set
  data.forEach(row => row[0] && categories.add(row[0]));
  // 將 Set 展開為陣列，轉成 JSON 格式回傳給前端
  return ContentService.createTextOutput(JSON.stringify({ categories: [...categories] }))
    .setMimeType(ContentService.MimeType.JSON); // 指定回傳格式為 JSON
}

/**
 * 🗂️ getStockByCategory：根據刀具類別回傳庫存資料
 * 說明：從「刀具庫存表」中，找出指定類別的所有品項，並將結果格式化為 JSON 回傳。
 * 傳入參數：
 *   - category（字串）：查詢的刀具類別（如：車刀、鑽頭等）
 * 回傳格式：[{ 編號, 類別, 品名規格, 廠牌, 庫存數量 }, ...]
 */
function getStockByCategory(category) {
  // 取得名為「刀具庫存表」的工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("刀具庫存表");
  // 取得整張表的所有資料（含標題列）
  const data = sheet.getDataRange().getValues();
  // 建立空陣列，準備存放符合條件的資料
  const result = [];
  // 遍歷所有資料列（包含標題，實務上建議可跳過第 1 列）
  data.forEach(row => {
    // 若第 2 欄（類別欄）與傳入參數一致，則加入結果陣列
    if (row[1] === category) {
      result.push({
        編號: row[0],        // A欄：編號
        類別: row[1],        // B欄：類別
        品名規格: row[2],    // C欄：品名規格
        廠牌: row[3],        // D欄：廠牌
        庫存數量: row[4]     // E欄：庫存數量
      });
    }
  });
  // 將查詢結果轉換成 JSON 字串並回傳，指定為 JSON 格式供前端 fetch 處理
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🛒 handleToolInventoryPost：處理刀具入庫資料
 * 說明：寫入「刀具入庫表」，同時更新「刀具庫存表」，若品項不存在則新增
 * 傳入參數：
 *   - data：來自前端 POST 傳入的物件，包含：category, spec, brand, vendor, unit_price, quantity, total_price, date, note
 * 回傳格式：
 *   - { success: true, stock: {...} }（更新後的庫存資訊）
 */
function handleToolInventoryPost(data) {
  // 取得「刀具入庫表」與「刀具庫存表」兩個工作表
  const sheet_in = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("刀具入庫表");
  const sheet_stock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("刀具庫存表");
  // 從 data 中取出各欄位資料（解構賦值）
  const { category, spec, brand, vendor, unit_price, quantity, total_price, date, note } = data;
  // 新入庫資料的編號（用目前最後一行當作流水號）
  const newId = sheet_in.getLastRow();
  // ✅ 寫入「刀具入庫表」一筆紀錄
  sheet_in.appendRow([
    newId, category, spec, brand, vendor, unit_price, quantity, total_price, date, note
  ]);
  // 取得庫存資料（含標題列）
  const stock_data = sheet_stock.getDataRange().getValues();
  // 初始化變數：是否找到舊資料、更新後的庫存資訊、如新增用的新編號
  let found = false, updatedStockInfo = null, newStockId = '';
  // 🔁 遍歷「刀具庫存表」資料，從第 2 列開始
  for (let i = 1; i < stock_data.length; i++) {
    const row = stock_data[i];
    // ✅ 若規格 + 廠牌完全一致，視為已有此刀具，直接累加數量
    if (row[2] === spec && row[3] === brand) {
      const newQty = Number(row[4]) + Number(quantity); // 加總新舊數量
      sheet_stock.getRange(i + 1, 5).setValue(newQty);   // 更新庫存數值（E 欄）
      found = true;
      // 建立更新後要回傳的庫存資訊
      updatedStockInfo = {
        編號: row[0], 類別: row[1], 品名規格: spec, 廠牌: brand, 數量: newQty
      };
      break;
    }
  }

  // 🆕 若未找到該品項，則新增一筆至「刀具庫存表」
  if (!found) {
    const prefix = category.trim().charAt(0); // 用類別第一字元當作編號前綴
    const categoryRows = stock_data.filter(row => row[1]?.charAt(0) === prefix); // 同一類別的行
    // 依照同類項目數量產生新編號（如：C0002）
    newStockId = prefix + (
      categoryRows.length > 0
        ? String(parseInt(categoryRows[categoryRows.length - 1][0].slice(1)) + 1).padStart(4, '0')
        : '0001'
    );
    // 寫入新項目至庫存表
    sheet_stock.appendRow([newStockId, category, spec, brand, Number(quantity)]);
    sortStockSheet(); // 排序庫存表，確保編號有順序
    // 建立回傳的庫存資料
    updatedStockInfo = {
      編號: newStockId, 類別: category, 品名規格: spec, 廠牌: brand, 數量: Number(quantity)
    };
  }
  // ✏️ 若廠牌/廠商是新的，則寫入「data」表並排序（可輸入新選項）
  updateDataSheetIfNew(brand, vendor);
  // ✅ 將處理結果與更新後的庫存資訊一起回傳（提供前端即時顯示）
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    stock: updatedStockInfo
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🔢 sortStockSheet：將「刀具庫存表」依照刀具編號排序
 * 說明：當新增新刀具項目後，為了保持資料一致性，會依據 A 欄「編號」欄位做升冪排序
 */
function sortStockSheet() {
  // 取得「刀具庫存表」工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("刀具庫存表");
  // 取得整張表格的所有資料（含標題列）
  const data = sheet.getDataRange().getValues();
  // 將標題列（第1列）單獨取出
  const header = data[0];
  // 將資料列（第2列以後）依照編號欄（A欄）使用 localeCompare 排序（字串排序）
  const rows = data.slice(1).sort((a, b) => a[0].localeCompare(b[0]));
  // 將排序後的資料覆蓋回原表（從第2列開始，不動標題列）
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

/**
 * 🆕 updateDataSheetIfNew：自動新增廠牌與廠商（如為新資料）
 * 說明：當使用者在入庫時輸入新廠牌或新廠商，會自動將這些新選項加入 data 表，
 *       並對廠牌（B欄）與廠商（C欄）進行依字數排序（短的優先）
 */
function updateDataSheetIfNew(brand, vendor) {
  // 取得「data」工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
  // 從第2列開始取得 A~C 欄的資料（含：類別、廠牌、廠商）
  const data = sheet.getRange(2, 1, sheet.getLastRow(), 3).getValues();
  // 提取目前已有的廠牌（B欄）與廠商（C欄）內容，並過濾空白
  const existingBrands = data.map(r => r[1]).filter(v => v);
  const existingVendors = data.map(r => r[2]).filter(v => v);
  // 準備要新增的資料列（格式為三欄：類別、廠牌、廠商）
  const newEntries = [];
  // 若輸入的廠牌不在既有資料中，則加入一筆新的廠牌（B欄填值，其它為 null）
  if (brand && !existingBrands.includes(brand)) newEntries.push([null, brand, null]);
  // 若輸入的廠商不在既有資料中，則加入一筆新的廠商（C欄填值，其它為 null）
  if (vendor && !existingVendors.includes(vendor)) newEntries.push([null, null, vendor]);
  // 寫入新資料到 data 表（如果有新項目）
  if (newEntries.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newEntries.length, 3).setValues(newEntries);
  }
  // 重新取得所有資料（從第2列起，略過標題列）
  const updated = sheet.getDataRange().getValues().slice(1);
  // 重新整理出所有廠牌與廠商清單（去空值）
  const brandList = updated.map(r => r[1]).filter(v => v);
  const vendorList = updated.map(r => r[2]).filter(v => v);
  // 利用自訂排序函式（依字數長短排序）處理
  const sortedBrands = countAndSortByLength(brandList);
  const sortedVendors = countAndSortByLength(vendorList);
  // 將排序後的廠牌依序寫回 B 欄
  for (let i = 0; i < sortedBrands.length; i++) {
    sheet.getRange(i + 2, 2).setValue(sortedBrands[i]);
  }
  // 將排序後的廠商依序寫回 C 欄
  for (let i = 0; i < sortedVendors.length; i++) {
    sheet.getRange(i + 2, 3).setValue(sortedVendors[i]);
  }
}

/**
 * ✏️ countAndSortByLength：移除重複後依字數排序（短字優先）
 * 說明：接收一個字串陣列，去除重複項目後，依照每個字串的長度從短到長排序。
 * 用途：常用於廠牌、廠商排序（名稱越短越先顯示於下拉選單）
 * 
 * @param {Array} list - 要排序的字串陣列（例如所有廠牌或廠商名稱）
 * @return {Array} 排序後的陣列（已去除重複，並依長度由小到大排序）
 */
function countAndSortByLength(list) {
  return [...new Set(list)]              // 將陣列轉成 Set 再展開成新陣列（移除重複）
    .sort((a, b) => a.length - b.length); // 依字串長度升冪排序（短字優先）
}

/**
 * 🔍 searchStock：於「刀具庫存表」中進行模糊搜尋
 * 說明：根據使用者輸入的關鍵字（可支援多關鍵字以 + 分隔），從 C 欄「品名規格」欄位中查找符合的列
 * 傳入參數：
 *   - keyword（字串）：查詢字串，可包含 "+" 代表多個關鍵詞（例如 WNMG+0804）
 * 回傳格式：
 *   - { header: [...], results: [...] }（包含標題列與符合條件的資料列）
 */
function searchStock(keyword) {
  // 取得名為「刀具庫存表」的工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("刀具庫存表");
  // 取得整張表的所有資料（包含標題列）
  const data = sheet.getDataRange().getValues();
  // 取出標題列（第1列），將來一併回傳
  const header = data[0];
  // 建立結果儲存用的陣列
  const results = [];
  // 處理關鍵字字串：
  // 1. 去除前後空白 2. 全轉大寫（比對用）3. 以「+」切割 4. 過濾空字串
  const keywords = keyword.trim().toUpperCase().split("+").map(k => k.trim()).filter(k => k);
  // 從第2列開始逐行掃描資料
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // 將 C 欄「品名規格」轉成大寫字串做比對
    const name = row[2]?.toString().toUpperCase();
    // 若 name 存在，且所有關鍵字都包含在品名中，則加入結果
    if (name && keywords.every(kw => name.includes(kw))) {
      results.push(row); // 將整列加入查詢結果
    }
  }
  // 回傳：包含標題列與查詢結果
  return { header, results };
}

/**
 * 📋 searchInDetailSheet：搜尋「刀具入庫表」的品名紀錄（模糊查詢）
 * 說明：與 searchStock 類似，但此處查的是入庫明細紀錄（包含廠商、單價、備註等）
 * 支援多關鍵字（用 + 隔開）進行「品名規格」欄位模糊查詢
 * 
 * 傳入參數：
 *   - keyword：查詢用字串（可支援 "+" 分隔的多關鍵詞）
 * 回傳格式：
 *   - { header: [...], results: [...] }（包含表頭與符合條件的列）
 */
function searchInDetailSheet(keyword) {
  // 取得名為「刀具入庫表」的工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("刀具入庫表");
  // 取得整張表格的所有內容（包含標題列）
  const data = sheet.getDataRange().getValues();
  // 將第1列（標題）取出，準備回傳給前端顯示表頭用
  const header = data[0];
  // 建立搜尋結果陣列
  const results = [];
  // 將關鍵字做以下處理：
  // 1. 去除首尾空白 2. 全部轉為大寫 3. 用「+」拆開成陣列 4. 去除空字串
  const keywords = keyword.trim().toUpperCase().split("+").map(k => k.trim()).filter(k => k);
  // 從第2列開始遍歷每一筆資料
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // 取出第 3 欄（C欄，品名規格）轉為大寫，做比對
    const name = row[2]?.toString().toUpperCase();
    // 若品名存在，且所有關鍵詞都包含於品名中 → 加入結果
    if (name && keywords.every(kw => name.includes(kw))) {
      results.push(row); // 將符合的整列資料加入結果
    }
  }
  // 回傳結果：包含表頭與查到的資料列
  return { header, results };
}

/**
 * 🏭 handleProductionEntryLoad：載入加工數量登記頁所需的下拉選單資料
 * 說明：根據傳入的 type 參數，回傳對應的下拉選單資料（機台、產品、訂單）
 * 使用方式：前端以 ?page=production_entry&type=machine / product / order 呼叫
 * 
 * @param {Object} param - URL 傳入的參數物件
 *   - type: 指定要取得哪一種下拉資料（machine / product / order）
 *   - pd: 若 type=order 時需提供產品編號（pd）
 * 
 * 回傳：ContentService.createTextOutput(JSON)，回傳對應下拉清單陣列
 */
function handleProductionEntryLoad(param) {
  const type = param.type; // 取得傳入的 type 參數（用來決定要抓哪一類資料）
  // ✅ 機台下拉選單：從 data 表 D 欄抓取
  if (type === 'machine') {
    return ContentService.createTextOutput(JSON.stringify(getMachineList()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ✅ 產品編號下拉選單：從 data 表 E 欄抓取
  if (type === 'product') {
    return ContentService.createTextOutput(JSON.stringify(getProductList()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ✅ 訂單編號下拉選單：需提供對應產品編號（pd），從訂單編號表中查找
  if (type === 'order' && param.pd) {
    return ContentService.createTextOutput(JSON.stringify(getOrderListByProduct(param.pd)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ⚠️ 若 type 參數不符合上述任一情況，則回傳錯誤訊息
  return ContentService.createTextOutput(JSON.stringify({ error: "缺少必要參數" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🏭 getMachineList：取得機台清單
 * 說明：從「data」工作表 D 欄（第4欄）擷取機台名稱清單，忽略空值。
 * 回傳格式：一維陣列，如 ["機台1", "機台2", "NP20CM", ...]
 */
function getMachineList() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("data")               // 取得 data 表
    .getRange("D2:D")                     // 從第2列開始抓 D 欄（機台）
    .getValues()                          // 取得值，為二維陣列
    .flat()                               // 展平成一維陣列
    .filter(String);                      // 過濾空字串，只保留有值的機台名稱
}

/**
 * 🧾 getProductList：取得產品編號清單
 * 說明：從「data」工作表 E 欄（第5欄）擷取產品代號清單，忽略空值。
 * 回傳格式：一維陣列，如 ["Z1234", "X4567", "AZ789", ...]
 */
function getProductList() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("data")               // 取得 data 表
    .getRange("E2:E")                     // 從第2列開始抓 E 欄（產品編號）
    .getValues()                          // 取得值，為二維陣列
    .flat()                               // 展平成一維陣列
    .filter(String);                      // 過濾空字串，只保留有值的產品編號
}

/**
 * 📦 getOrderListByProduct：根據產品編號取得所有對應的訂單編號
 * 說明：從「訂單編號表」中，找出第1欄符合指定產品代號的列，並擷取第2欄（訂單編號）
 * 
 * 傳入參數：
 *   - pd：產品代號（例如 "Z1234"）
 * 回傳格式：一維陣列，如 ["ORD-001", "ORD-002"]
 */
function getOrderListByProduct(pd) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("訂單編號表");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues(); // 去除標題列
  let orderNumbers = [];

  data.forEach(row => {
    if (row[0] === pd) {
      orderNumbers = row.slice(1); // 從第2欄開始取到最後一欄
    }
  });

  return orderNumbers.filter(String); // 移除空白或 undefined
}

/**
 * 📊 sortProductionSheet：排序「加工數量表」資料
 * 說明：依據 A 欄（日期）與 B 欄（機台）做升冪排序，
 *       讓資料按照時間與機台順序排列，便於管理與查詢。
 */
function sortProductionSheet() {
  // 取得「加工數量表」工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  // 取得目前最後一列的列號，用來界定要排序的範圍
  const lastRow = sheet.getLastRow();
  // 取得從第2列（資料列）開始，到最後一列的 A~Z 欄範圍（保留完整欄位）
  const range = sheet.getRange("A2:Z" + lastRow);
  // 執行排序：
  // 1. 先依 A 欄（日期）升冪排序
  // 2. 若日期相同，依 B 欄（機台）升冪排序
  range.sort([
    { column: 1, ascending: true },  // A欄：日期
    { column: 2, ascending: true }   // B欄：機台
  ]);
}

/**
 * 🕒 getLastSubmittedProduction：查詢「加工數量表」中最新一筆資料
 * 說明：根據第 9 欄（I 欄，時間戳記）找出最後一筆加工資料，
 *       回傳完整欄位內容與時間戳，供前端提示用。
 * 
 * 回傳格式：
 *   {
 *     dt: 日期,
 *     mc: 機台,
 *     pd: 產品,
 *     or: 訂單,
 *     hr: 小時,
 *     min: 分鐘,
 *     num: 數量,
 *     com: 備註,
 *     timestamp: 時間戳字串,
 *     rowIndex: Google Sheet 中實際列號
 *   }
 */
function getLastSubmittedProduction() {
  // 取得「加工數量表」工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  // 取得所有資料（含標題列）
  const data = sheet.getDataRange().getValues();
  // 若資料只有一列（只有標題），直接回傳 null
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify(null))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 🔍 用 reduce 找出第 9 欄（時間戳）最新的資料索引（排除標題列）
  const lastIndex = data.slice(1).reduce((latestIdx, row, i, arr) => {
    const timeA = new Date(arr[latestIdx][8]); // 當前最大時間
    const timeB = new Date(row[8]);           // 比對時間
    return timeB > timeA ? i : latestIdx;     // 若 timeB 更新，就換掉索引
  }, 0) + 1; // 加回標題列偏移量
  // 取出最後一筆資料
  const lastRow = data[lastIndex];
  // 將原始 timestamp 格式化為台灣時區 yyyy-MM-dd HH:mm:ss
  const rawTimestamp = new Date(lastRow[8]);
  const tz = Utilities.formatDate(rawTimestamp, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  // 整理結果內容（欄位順序依表格欄對應）
  const result = {
    dt: lastRow[0],     // A欄：日期
    mc: lastRow[1],     // B欄：機台
    pd: lastRow[2],     // C欄：產品編號
    or: lastRow[3],     // D欄：訂單編號
    hr: lastRow[4],     // E欄：小時
    min: lastRow[5],    // F欄：分鐘
    num: lastRow[6],    // G欄：加工數量
    com: lastRow[7],    // H欄：備註
    timestamp: tz,      // I欄（格式化）
    rowIndex: lastIndex + 1 // 真實在表格中的列數（給刪除用）
  };
  // 回傳 JSON 格式的結果（前端 fetch 使用）
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ➕ addProductionData：新增一筆加工數據
 * 說明：寫入一筆加工登記資料（包含防重複、加上時間戳、排序），回傳成功與否結果。
 * 
 * 傳入參數（rowData）應包含：
 *   - dt：加工日期（YYYY-MM-DD）
 *   - mc：機台名稱
 *   - pd：產品編號
 *   - or：訂單編號
 *   - hr：小時
 *   - min：分鐘
 *   - num：加工數量
 *   - comments：備註
 */
function addProductionData(rowData) {
  // 取得「加工數量表」工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  // 建立目前時間作為時間戳（第9欄）
  const currentDate = new Date();
  // ✅ 若目前表格只有標題列（沒有資料）
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([
      rowData.dt,         // A欄：日期
      rowData.mc,         // B欄：機台
      rowData.pd,         // C欄：產品
      rowData.or,         // D欄：訂單
      rowData.hr,         // E欄：小時
      rowData.min,        // F欄：分鐘
      rowData.num,        // G欄：數量
      rowData.comments,   // H欄：備註
      currentDate         // I欄：時間戳
    ]);
    // 👉 加入置中對齊
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 9).setHorizontalAlignment("center");
    sortProductionSheet(); // 執行排序（依日期與機台）
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 📄 若已有資料，先取出所有現有資料（不含標題列）
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  // 將輸入的日期轉換格式，方便後續與既有資料比對（YYYY/MM/DD）
  const inputDate = new Date(rowData.dt);
  const formattedInputDate = inputDate.getFullYear() + '/' +
    (inputDate.getMonth() + 1) + '/' +
    inputDate.getDate();
  // 🔁 檢查是否已有相同資料（同一天 + 同機台 + 同產品 + 同訂單）
  for (let i = 0; i < dataRange.length; i++) {
    const existingDate = new Date(dataRange[i][0]); // A欄
    const formattedExistingDate = existingDate.getFullYear() + '/' +
      (existingDate.getMonth() + 1) + '/' +
      existingDate.getDate();
    if (
      formattedExistingDate === formattedInputDate &&
      dataRange[i][1] === rowData.mc &&   // B欄：機台
      dataRange[i][2] === rowData.pd &&   // C欄：產品
      dataRange[i][3] === rowData.or      // D欄：訂單
    ) {
      // ⚠️ 發現重複資料，不允許寫入，回傳失敗
      return ContentService.createTextOutput(JSON.stringify({ success: false }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  // ✅ 若確認無重複，寫入新資料
  sheet.appendRow([
    formattedInputDate,  // A欄：格式化後的日期
    rowData.mc,          // B欄：機台
    rowData.pd,          // C欄：產品
    rowData.or,          // D欄：訂單
    rowData.hr,          // E欄：小時
    rowData.min,         // F欄：分鐘
    rowData.num,         // G欄：數量
    rowData.comments,    // H欄：備註
    currentDate          // I欄：時間戳
  ]);
  // 👉 加入置中對齊
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 9).setHorizontalAlignment("center");
  // 自動重新排序（依日期 + 機台）
  sortProductionSheet();
  // 回傳成功訊息
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ❌ deleteLastProductionRow：刪除「加工數量表」中最後一筆資料
 * 說明：根據時間戳（I欄）找出時間最新的一筆資料並刪除，常搭配「誤登記後立即撤回」的功能。
 * 
 * 回傳格式：
 *   - { success: true } 若刪除成功
 *   - { success: false } 若表中無資料可刪
 */
function deleteLastProductionRow() {
  // 取得「加工數量表」工作表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  // 取得整張表的所有資料（含標題列）
  const data = sheet.getDataRange().getValues();
  // 若表格只有標題列，代表沒有可刪的資料，回傳失敗
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ success: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 🕒 找出第 9 欄（時間戳）欄位中，最新的一筆資料所在索引
  const lastIndex = data.slice(1).reduce((latestIdx, row, i, arr) => {
    const timeA = new Date(arr[latestIdx][8]); // 目前最新的時間
    const timeB = new Date(row[8]);           // 當前比對列的時間
    return timeB > timeA ? i : latestIdx;     // 若 timeB 較新，更新索引
  }, 0) + 2; // 加回標題列偏移量（slice(1) + index 基於 0 → 所以是 +2）
  // 刪除該列（Google Sheets 索引從 1 開始）
  sheet.deleteRow(lastIndex);
  // 回傳成功訊息
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🔍 getLatestProductionByProduct：查詢指定產品.訂單的最後一次加工紀錄
 * 說明：先找出該產品最後加工的日期，再擷取當天的所有加工紀錄（可能多筆），供前端顯示使用。
 * 
 * 傳入參數：
 *   - pd：產品編號（如 Z1123）
 * 
 * 回傳格式：
 *   - [[日期, 組合文字, 數量], ...]
 *     → 組合文字格式為「產品編號 . 機台名稱」
 */
function getLatestProductionByProduct(pd, od) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const data = ws.getDataRange().getValues();
  let filteredData = [];
  let latestDate = null;

  // 第一步：找出指定「產品+訂單」的最新加工日期
  for (let i = 1; i < data.length; i++) {
    const recordDate = new Date(data[i][0]); // A欄：日期
    const productCode = data[i][2];          // C欄：產品編號
    const orderCode = data[i][3];            // D欄：訂單編號

    if (productCode === pd && orderCode === od) {
      if (!latestDate || recordDate > latestDate) {
        latestDate = recordDate;
      }
    }
  }

  // 第二步：篩出該「產品+訂單」在最新日期的資料
  if (latestDate) {
    const formattedDate = `${latestDate.getFullYear()}/${latestDate.getMonth() + 1}/${latestDate.getDate()}`;
    for (let j = 1; j < data.length; j++) {
      const recordDate = new Date(data[j][0]);
      const productCode = data[j][2];
      const orderCode = data[j][3];
      const machineCode = data[j][1];
      const quantity = data[j][6];

      const checkDate = `${recordDate.getFullYear()}/${recordDate.getMonth() + 1}/${recordDate.getDate()}`;

      if (productCode === pd && orderCode === od && checkDate === formattedDate) {
        filteredData.push([checkDate, `${productCode}  .  ${machineCode}  .  ${orderCode}`, quantity]);
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify(filteredData))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 📅 getProductionQuery：查詢某產品在特定日期區間內的所有加工紀錄
 * 說明：篩選出符合產品編號、加工日期在指定範圍內的所有資料，格式化後回傳。
 *
 * 傳入參數 param 包含：
 *   - pd：產品編號（例如 Z1234）
 *   - dt1：起始日期（例如 2024-04-01）
 *   - dt2：結束日期（例如 2024-04-11）
 * 
 * 回傳格式：
 *   - [[日期, "產品 . 機台", 數量], ...]（JSON 格式）
 */
function getProductionQuery(param) {
  const { pd, od, dt1, dt2 } = param; // 多解構一個 od（訂單編號）
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const data = ws.getDataRange().getValues();
  const filteredData = [];
  const startDate = new Date(dt1);
  const endDate = new Date(dt2);

  for (let i = 1; i < data.length; i++) { // 從第1列開始，跳過標題列
    const recordDate = new Date(data[i][0]); // A欄：日期
    const machineCode = data[i][1];          // B欄：機台
    const productCode = data[i][2];          // C欄：產品編號
    const orderCode = data[i][3];            // D欄：訂單編號
    const quantity = data[i][6];             // G欄：加工數量

    if (
      productCode === pd &&
      orderCode === od &&
      recordDate >= startDate &&
      recordDate <= endDate
    ) {
      const formattedDate = `${recordDate.getFullYear()}/${recordDate.getMonth() + 1}/${recordDate.getDate()}`;
      filteredData.push([
        formattedDate,
        `${productCode} . ${machineCode} . ${orderCode}`,
        quantity
      ]);
    }
  }

  return ContentService.createTextOutput(JSON.stringify(filteredData))
    .setMimeType(ContentService.MimeType.JSON);
}


///////////////////////////////////////////////////////////////////////
/**
 * 🔽 匯出資料查詢：取得指定產品與訂單編號的所有資料
 */
function exportProductionData(productCode, orderCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(`A1:H${lastRow}`).getDisplayValues();

  const filtered = data.filter((row, index) => {
    if (index === 0) return false;
    return row[2] === productCode && row[3] === orderCode;
  });

  return ContentService.createTextOutput(JSON.stringify(filtered))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🔽 匯出用下拉選單選項（產品與訂單編號）
 */
function getDropdownValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const data = sheet.getDataRange().getValues();

  const productSet = new Set();
  const orderSet = new Set();

  data.forEach((row, index) => {
    if (index === 0) return;
    if (row[2]) productSet.add(row[2]);
    if (row[3]) orderSet.add(row[3]);
  });

  return ContentService.createTextOutput(JSON.stringify({
    products: [...productSet],
    orders: [...orderSet]
  })).setMimeType(ContentService.MimeType.JSON);
}