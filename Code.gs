/**
 * 數字森林闖關 - 後台紀錄系統
 * 部署方式：
 * 1. 開一個新的 Google 表單試算表（Google Sheets）。
 * 2. 上方選單「擴充功能」→「Apps Script」。
 * 3. 把這個檔案的全部內容貼進去，取代原本的內容。
 * 4. 上方工具列的函式下拉選單選「setupSheets」，按執行（▶），
 *    第一次執行會跳出授權視窗，允許即可。這會自動建立四個工作表：
 *    「名單」「作答紀錄」「學生總覽」「數感特訓紀錄」。
 *    這一份後台程式同時服務「數字森林闖關」和「數感快手」兩個遊戲，
 *    只要部署一次、一個網址，兩邊遊戲都能用。
 * 5. 執行完後，去「名單」工作表，把全校/全班的班級、座號、姓名填進去
 *    （第一列是標題列，不要刪，從第二列開始填資料）。
 * 6. 上方選單「部署」→「新增部署作業」：
 *      類型選「網頁應用程式」
 *      「執行身份」選「我」
 *      「可存取權限」一定要選「任何人」（不要選「知道 Google 帳戶的任何人」，
 *       不然孩子在遊戲裡送資料會失敗）
 * 7. 按部署，會給你一個網址，長得像：
 *      https://script.google.com/macros/s/AKfycb.../exec
 *    把這個網址複製起來，貼到任一個遊戲網頁右上角的齒輪圖示設定裡即可（兩個遊戲共用同一個網址）。
 * 8. 之後如果你修改了這個程式碼，要「部署」→「管理部署作業」→
 *    點編輯（鉛筆）→版本選「新版本」→部署，網址才會套用新的程式碼。
 */

var SHEET_ROSTER = '名單';
var SHEET_LOG = '作答紀錄';
var SHEET_SUMMARY = '學生總覽';
var SHEET_NS = '數感特訓紀錄';

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

/**
 * 執行一次即可：建立四個工作表與標題列/公式。
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

  var ns = getSheet_(SHEET_NS);
  if (ns.getLastRow() === 0) {
    ns.appendRow(['時間', '班級', '姓名', '模式', '答對題數', '總題數', 'XP', '平均反應時間(ms)', '等級']);
    ns.setFrozenRows(1);
  }

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
  if (action === 'nsLeaderboard') {
    var nsSheet = getSheet_(SHEET_NS);
    var nsLastRow = nsSheet.getLastRow();
    var totals = {};
    if (nsLastRow > 1) {
      var nsData = nsSheet.getRange(2, 1, nsLastRow - 1, 9).getValues();
      for (var i = 0; i < nsData.length; i++) {
        var cls = String(nsData[i][1] || '').trim();
        var name = String(nsData[i][2] || '').trim();
        var xp = Number(nsData[i][6]) || 0;
        var lvl = Number(nsData[i][8]) || 1;
        if (!cls || !name) continue;
        var key = cls + '\u0001' + name;
        if (!totals[key]) totals[key] = { className: cls, studentName: name, totalXp: 0, maxLevel: 1 };
        totals[key].totalXp += xp;
        if (lvl > totals[key].maxLevel) totals[key].maxLevel = lvl;
      }
    }
    var list = Object.keys(totals).map(function (k) { return totals[k]; });
    list.sort(function (a, b) { return b.totalXp - a.totalXp; });
    return ContentService.createTextOutput(JSON.stringify({ ok: true, leaderboard: list.slice(0, 50) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: '數字森林闖關 API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 給遊戲送出成績用的（POST，body 是 JSON）
 * body.type === 'numberSense' 時記錄到「數感特訓紀錄」，否則記錄到「作答紀錄」（數字森林闖關）
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.type === 'numberSense') {
      var nsSheet = getSheet_(SHEET_NS);
      if (nsSheet.getLastRow() === 0) {
        nsSheet.appendRow(['時間', '班級', '姓名', '模式', '答對題數', '總題數', 'XP', '平均反應時間(ms)', '等級']);
      }
      nsSheet.appendRow([
        new Date(),
        body.className || '',
        body.studentName || '',
        body.mode || '',
        body.correct || 0,
        body.total || 0,
        body.xp || 0,
        body.avgReactionMs || 0,
        body.level || 1
      ]);
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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
