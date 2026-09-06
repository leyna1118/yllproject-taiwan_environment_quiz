// 綠色旅遊測驗．匿名使用記錄後端
//
// 部署方式：
//   1. 開一個新的 Google Sheet，選單 擴充功能 → Apps Script
//   2. 把這份檔案整份內容貼進去，取代預設的 Code.gs，存檔
//   3. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
//        執行身分：我
//        誰可以存取：所有人
//   4. 部署後複製網址，貼到 index.html 最上方的 TRACK_ENDPOINT
//
// 之後如果改了這份程式碼，要「管理部署作業」→ 編輯現有部署 → 換一個新版本，
// 存檔不會讓網址上的舊版本生效。

var SHEET_NAME = "記錄";
var HEADER = ["時間戳", "事件", "匿名ID", "第幾次作答", "地區類型", "主線任務", "知識↔實踐", "深度↔廣度", "來源"];

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    getSheet_().appendRow([
      new Date(),
      payload.event || "",
      payload.uid || "",
      payload.attempt != null ? payload.attempt : "",
      payload.area || "",
      payload.activity || "",
      payload.know != null ? payload.know : "",
      payload.depth != null ? payload.depth : "",
      payload.source || ""
    ]);
  } catch (err) {
    // 前端用 no-cors 送出，不會讀這裡的回應，記錄失敗就靜靜跳過即可
  }
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  return ContentService.createTextOutput("綠色旅遊測驗記錄端點運作中");
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
  }
  return sheet;
}
