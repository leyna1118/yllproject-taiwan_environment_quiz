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
//
// 這個端點的網址會出現在公開的原始碼裡，任何人都可以直接 POST 資料進來，
// 所以 doPost 會先用下面的白名單驗證，格式不符就安靜丟棄、不寫入試算表。

var SHEET_NAME = "記錄";
var HEADER = ["時間戳", "事件", "匿名ID", "第幾次作答", "地區類型", "主線任務", "知識↔實踐", "深度↔廣度", "來源"];

var FEEDBACK_SHEET_NAME = "回饋";
var FEEDBACK_HEADER = ["時間戳", "事件", "匿名ID", "地區類型", "主線任務", "評分", "內容", "來源"];

var MAX_BODY_LENGTH = 2000;
var VALID_EVENTS = ["載入", "開始", "完成", "評分", "回饋"];
var FEEDBACK_EVENTS = ["評分", "回饋"];
var VALID_SOURCES = ["正式", "測試"];
var VALID_AREAS = ["海邊", "城市", "鄉間", "山林"];

// 16 個主線任務行政區，從 data/quiz-data.js 的 activities 底下每個地區類型
// 各 4 筆的 region 欄位抓出來的固定清單。
// 之後如果在 data.xlsx 改了主線任務（新增/改名/刪除），這份清單要跟著手動更新，
// 不然新的主線任務送出的「完成」事件會被白名單擋掉、不會被記錄。
var VALID_ACTIVITIES = [
  "小琉球", "蘭嶼", "金門金城（建功嶼）", "花蓮縣新城鄉",
  "台北‧大稻埕", "台南‧山上花園水道博物館", "高雄‧哈瑪星", "台中‧富興工廠1962文創聚落",
  "苗栗淺山（通霄／苑裡等）", "花蓮富里羅山村", "池上鄉萬安社區（台東）", "屏東恆春社頂",
  "新竹北埔鹿寮坑", "貓空（台北文山區）", "南投集集", "阿里山（嘉義）"
];

function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents;
    if (typeof raw !== "string" || raw.length === 0 || raw.length > MAX_BODY_LENGTH) {
      return ContentService.createTextOutput("ok");
    }

    var payload = JSON.parse(raw);
    if (!isValidPayload_(payload)) {
      return ContentService.createTextOutput("ok");
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return ContentService.createTextOutput("ok"); // 10 秒內拿不到鎖就放棄這一筆，不要卡住整個請求
    }
    try {
      if (FEEDBACK_EVENTS.indexOf(payload.event) !== -1) {
        getFeedbackSheet_().appendRow([
          new Date(),
          payload.event,
          payload.uid,
          payload.area,
          payload.activity,
          payload.rating,
          payload.text,
          payload.source
        ]);
      } else {
        getSheet_().appendRow([
          new Date(),
          payload.event,
          payload.uid,
          payload.attempt,
          payload.area,
          payload.activity,
          payload.know,
          payload.depth,
          payload.source
        ]);
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    // 前端用 no-cors 送出，不會讀這裡的回應，任何失敗（含格式錯誤的 JSON）都靜靜跳過即可
  }
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  return ContentService.createTextOutput("綠色旅遊測驗記錄端點運作中");
}

// 白名單驗證：欄位型別／範圍都要符合，且「完成」事件必須四個結果欄位都有值，
// 「載入」「開始」則必須四個結果欄位都是空字串，「評分」「回饋」則走各自的驗證，
// 否則視為格式不符。
function isValidPayload_(p) {
  if (!p || typeof p !== "object") return false;
  if (VALID_EVENTS.indexOf(p.event) === -1) return false;
  if (VALID_SOURCES.indexOf(p.source) === -1) return false;
  if (typeof p.uid !== "string" || p.uid.length > 64) return false;

  if (FEEDBACK_EVENTS.indexOf(p.event) !== -1) {
    return isValidFeedbackPayload_(p);
  }

  if (!isIntInRange_(p.attempt, 0, 10000)) return false;
  if (p.area !== "" && VALID_AREAS.indexOf(p.area) === -1) return false;
  if (p.activity !== "" && VALID_ACTIVITIES.indexOf(p.activity) === -1) return false;
  if (p.know !== "" && !isIntInRange_(p.know, 0, 100)) return false;
  if (p.depth !== "" && !isIntInRange_(p.depth, 0, 100)) return false;

  var filledCount = [p.area, p.activity, p.know, p.depth].filter(function (v) {
    return v !== "";
  }).length;
  if (p.event === "完成" && filledCount !== 4) return false;
  if (p.event !== "完成" && filledCount !== 0) return false;

  return true;
}

// 「評分」必須 rating 為 1/2/3 且不得帶 text；「回饋」必須有非空、不超過 500 字的
// text 且不得帶 rating；兩者的 area、activity 都必須是既有白名單值（不能是空字串）。
function isValidFeedbackPayload_(p) {
  if (VALID_AREAS.indexOf(p.area) === -1) return false;
  if (VALID_ACTIVITIES.indexOf(p.activity) === -1) return false;

  if (p.event === "評分") {
    if (p.rating !== 1 && p.rating !== 2 && p.rating !== 3) return false;
    if (p.text !== "") return false;
    return true;
  }

  // 回饋
  if (typeof p.text !== "string" || p.text.length === 0 || p.text.length > 500) return false;
  if (p.rating !== "") return false;
  return true;
}

function isIntInRange_(n, min, max) {
  return typeof n === "number" && isFinite(n) && Math.floor(n) === n && n >= min && n <= max;
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

function getFeedbackSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FEEDBACK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FEEDBACK_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(FEEDBACK_HEADER);
  }
  return sheet;
}
