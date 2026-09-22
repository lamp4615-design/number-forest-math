/**
 * 數字森林闖關 - 後台紀錄系統
 * 部署方式：
 * 1. 開一個新的 Google 表單試算表（Google Sheets）。
 * 2. 上方選單「擴充功能」→「Apps Script」。
 * 3. 把這個檔案的全部內容貼進去，取代原本的內容。
 * 4. 上方工具列的函式下拉選單選「setupSheets」，按執行（▶），
 *    第一次執行會跳出授權視窗，允許即可。這會自動建立三個工作表：
 *    「名單」「作答紀錄」「學生總覽」。
 * 5. 執行完後，去「名單」工作表，把全校/全班的班級、座號、姓名填進去
 *    （第一列是標題列，不要刪，從第二列開始填資料）。
 * 6. 上方選單「部署」→「新增部署作業」：
 *      類型選「網頁應用程式」
 *      「執行身份」選「我」
 *      「可存取權限」一定要選「任何人」（不要選「知道 Google 帳戶的任何人」，
 *       不然孩子在遊戲裡送資料會失敗）
 * 7. 按部署，會給你一個網址，長得像：
 *      https://script.google.com/macros/s/AKfycb.../exec
 *    把這個網址複製起來，貼到「數字森林闖關」網頁右上角的齒輪圖示設定裡。
 * 8. 之後如果你修改了這個程式碼，要「部署」→「管理部署作業」→
 *    點編輯（鉛筆）→版本選「新版本」→部署，網址才會套用新的程式碼。
 */

var SHEET_ROSTER = '名單';
var SHEET_LOG = '作答紀錄';
var SHEET_SUMMARY = '學生總覽';

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

/**
 * 執行一次即可：建立三個工作表與標題列/公式。
 */
function setupSheets() {
  var roster = getSheet_(SHEET_ROSTER);
  if (roster.getLastRow() === 0) {
    roster.appendRow(['班級', '座號', '姓名']);
    roster.setFrozenRows(1);
  }

  var log = getSheet_(SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.appendRow(['時間', '班級', '姓名', '關卡ID', '關卡名稱', '分類', '答對題數', '總題數', '星數']);
    log.setFrozenRows(1);
  }

  var summary = getSheet_(SHEET_SUMMARY);
  summary.clear();
  summary.appendRow(['班級', '姓名', '分類', '平均星數', '作答次數', '答對率(%)']);
  summary.setFrozenRows(1);
  summary.getRange('A2').setFormula(
    '=IFERROR(QUERY(' + SHEET_LOG + '!A:I,' +
    '"select B, C, F, avg(I), count(I), avg(G)/avg(H)*100 ' +
    'where B is not null and B <> \'\' ' +
    'group by B, C, F ' +
    'order by B, C, F ' +
    'label avg(I) \'\', count(I) \'\', avg(G)/avg(H)*100 \'\'", 0),"")'
  );

  SpreadsheetApp.getActiveSpreadsheet().toast('設定完成！請到「名單」工作表填班級與姓名。');
}

/**
 * 給遊戲讀取班級名單用的（GET ?action=roster）
 */
function doGet(e) {
  var action = e.parameter.action;
  if (action === 'roster') {
    var sheet = getSheet_(SHEET_ROSTER);
    var lastRow = sheet.getLastRow();
    var roster = {};
    if (lastRow > 1) {
      var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      for (var i = 0; i < data.length; i++) {
        var cls = String(data[i][0] || '').trim();
        var name = String(data[i][2] || '').trim();
        if (!cls || !name) continue;
        if (!roster[cls]) roster[cls] = [];
        roster[cls].push(name);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, roster: roster }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: '數字森林闖關 API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 給遊戲送出成績用的（POST，body 是 JSON）
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheet = getSheet_(SHEET_LOG);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['時間', '班級', '姓名', '關卡ID', '關卡名稱', '分類', '答對題數', '總題數', '星數']);
    }
    sheet.appendRow([
      new Date(),
      body.className || '',
      body.studentName || '',
      body.stageId || '',
      body.stageTitle || '',
      body.category || '',
      body.correct || 0,
      body.total || 0,
      body.stars || 0
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
