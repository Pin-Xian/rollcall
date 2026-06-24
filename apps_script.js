// ================================================
// Google Apps Script 後端
// 貼到 https://script.google.com 新專案
// ================================================

// ★ 請填入你的 Google Sheet ID（從網址複製）
const SHEET_ID = 'YOUR_SHEET_ID_HERE';

// Sheet 名稱常數
const SHEET_ROSTER   = '名冊';
const SHEET_SEATS    = '座位表';
const SHEET_ATTEND   = '點名紀錄';

// ──────────────────────────────────────────
// doPost：學生簽到時呼叫
// ──────────────────────────────────────────
function doPost(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);

  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'checkin') {
      const result = saveCheckin(data);
      res.setContent(JSON.stringify({ok: true, ...result}));
    } else {
      res.setContent(JSON.stringify({ok: false, msg: '未知動作'}));
    }
  } catch(err) {
    res.setContent(JSON.stringify({ok: false, msg: err.message}));
  }
  return res;
}

// ──────────────────────────────────────────
// doGet：老師端讀取資料時呼叫
// ──────────────────────────────────────────
function doGet(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);

  // CORS header 處理（Apps Script 不支援自訂 header，用 callback 方式）
  try {
    const action = e.parameter.action;

    if (action === 'getRoster') {
      res.setContent(JSON.stringify({ok: true, data: getRoster(e.parameter.cls)}));
    } else if (action === 'getSeats') {
      res.setContent(JSON.stringify({ok: true, data: getSeats(e.parameter.cls)}));
    } else if (action === 'getAttendance') {
      res.setContent(JSON.stringify({ok: true, data: getAttendance(e.parameter.cls, e.parameter.date)}));
    } else if (action === 'getClasses') {
      res.setContent(JSON.stringify({ok: true, data: getClasses()}));
    } else if (action === 'getWeekly') {
      res.setContent(JSON.stringify({ok: true, data: getWeeklyReport(e.parameter.cls, e.parameter.from, e.parameter.to)}));
    } else {
      res.setContent(JSON.stringify({ok: false, msg: '未知動作'}));
    }
  } catch(err) {
    res.setContent(JSON.stringify({ok: false, msg: err.message}));
  }
  return res;
}

// ──────────────────────────────────────────
// 學生簽到：寫入「點名紀錄」Sheet
// ──────────────────────────────────────────
function saveCheckin(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_ATTEND);

  // 第一次使用時自動建立 sheet
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ATTEND);
    sheet.appendRow(['日期','課程','班級','座位','座號','姓名','學號','簽到時間','狀態','Token']);
    sheet.setFrozenRows(1);
    // 格式化標題列
    sheet.getRange(1,1,1,10).setBackground('#1A1A18').setFontColor('#FFFFFF').setFontWeight('bold');
  }

  // 防重複：同一 token + 座位 只能簽到一次
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (existing[i][9] === data.token && existing[i][3] === data.seat) {
      return {msg: '已簽到', duplicate: true};
    }
  }

  // 判斷遲到
  const now      = new Date();
  const timeStr  = Utilities.formatDate(now, 'Asia/Taipei', 'HH:mm:ss');
  const [dh, dm] = (data.deadline || '23:59').split(':').map(Number);
  const nowMin   = now.getHours() * 60 + now.getMinutes();
  const isLate   = nowMin > dh * 60 + dm;
  const status   = isLate ? '遲到' : '準時';

  sheet.appendRow([
    data.date,
    data.course,
    data.cls || '',
    data.seat,
    data.sno  || '',
    data.name,
    data.sid  || '',
    timeStr,
    status,
    data.token
  ]);

  // 自動依時間排序（讓新資料在最下面即可，不做排序節省時間）
  return {status, timeStr};
}

// ──────────────────────────────────────────
// 讀取名冊
// ──────────────────────────────────────────
function getRoster(cls) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ROSTER);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];

  const headers = rows[0];
  const result  = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0] && !row[1] && !row[2]) continue; // 空行跳過
    const obj = {};
    headers.forEach((h, j) => obj[String(h).trim()] = String(row[j] || '').trim());
    if (cls && obj['班級'] !== cls) continue;
    result.push(obj);
  }
  return result;
}

// ──────────────────────────────────────────
// 讀取座位表
// ──────────────────────────────────────────
function getSeats(cls) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SEATS);
  if (!sheet) return {};

  const rows   = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < rows.length; i++) {
    const [c, seat, sno, name, sid] = rows[i].map(v => String(v||'').trim());
    if (!seat || !name) continue;
    if (cls && c !== cls) continue;
    result[seat] = {cls:c, sno, name, sid};
  }
  return result;
}

// ──────────────────────────────────────────
// 讀取點名紀錄（指定日期）
// ──────────────────────────────────────────
function getAttendance(cls, date) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ATTEND);
  if (!sheet) return {};

  const rows   = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < rows.length; i++) {
    const [d, course, c, seat, sno, name, sid, timeStr, status, token] = rows[i].map(v=>String(v||'').trim());
    if (date && d !== date) continue;
    if (cls  && c !== cls)  continue;
    result[seat] = {date:d, course, cls:c, seat, sno, name, sid, timeStr, status, token, late: status==='遲到'};
  }
  return result;
}

// ──────────────────────────────────────────
// 取得所有班級
// ──────────────────────────────────────────
function getClasses() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ROSTER);
  if (!sheet) return [];

  const rows    = sheet.getDataRange().getValues();
  const classes = new Set();
  for (let i = 1; i < rows.length; i++) {
    const cls = String(rows[i][0]||'').trim();
    if (cls) classes.add(cls);
  }
  return [...classes];
}

// ──────────────────────────────────────────
// 週報表：指定班級、日期範圍
// ──────────────────────────────────────────
function getWeeklyReport(cls, from, to) {
  const ss     = SpreadsheetApp.openById(SHEET_ID);
  const attend = ss.getSheetByName(SHEET_ATTEND);
  if (!attend) return {dates:[], students:[], rows:{}};

  const rows   = attend.getDataRange().getValues();
  const dates  = new Set();
  const byName = {}; // name -> {date -> status}

  for (let i = 1; i < rows.length; i++) {
    const [d, course, c, seat, sno, name, sid, timeStr, status] = rows[i].map(v=>String(v||'').trim());
    if (!d || !name) continue;
    if (cls && c !== cls) continue;
    if (from && d < from) continue;
    if (to   && d > to)   continue;
    dates.add(d);
    if (!byName[name]) byName[name] = {sno, sid};
    byName[name][d] = status;
  }

  // 補上缺席的學生（從名冊）
  const roster = getRoster(cls);
  roster.forEach(s => {
    if (!byName[s['姓名']]) byName[s['姓名']] = {sno: s['座號']||s['sno']||'', sid: s['學號']||s['sid']||''};
  });

  const sortedDates = [...dates].sort();
  const students    = Object.entries(byName).map(([name, info]) => ({
    name,
    sno: info.sno,
    sid: info.sid,
    records: sortedDates.map(d => info[d] || '缺席')
  })).sort((a,b) => parseInt(a.sno||999) - parseInt(b.sno||999));

  return {dates: sortedDates, students};
}

// ──────────────────────────────────────────
// 初始化 Sheet（第一次使用時執行）
// ──────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 名冊 Sheet
  if (!ss.getSheetByName(SHEET_ROSTER)) {
    const s = ss.insertSheet(SHEET_ROSTER);
    s.appendRow(['班級','座號','姓名','學號']);
    s.setFrozenRows(1);
    s.getRange(1,1,1,4).setBackground('#2D6A4F').setFontColor('#FFFFFF').setFontWeight('bold');
    // 範例資料
    s.appendRow(['資訊科技概論','1','王小明','S001']);
    s.appendRow(['資訊科技概論','2','李雅婷','S002']);
  }

  // 座位表 Sheet
  if (!ss.getSheetByName(SHEET_SEATS)) {
    const s = ss.insertSheet(SHEET_SEATS);
    s.appendRow(['班級','座位','座號','姓名','學號']);
    s.setFrozenRows(1);
    s.getRange(1,1,1,5).setBackground('#2D6A4F').setFontColor('#FFFFFF').setFontWeight('bold');
    s.appendRow(['資訊科技概論','1A','1','王小明','S001']);
    s.appendRow(['資訊科技概論','1B','2','李雅婷','S002']);
  }

  // 點名紀錄 Sheet
  if (!ss.getSheetByName(SHEET_ATTEND)) {
    const s = ss.insertSheet(SHEET_ATTEND);
    s.appendRow(['日期','課程','班級','座位','座號','姓名','學號','簽到時間','狀態','Token']);
    s.setFrozenRows(1);
    s.getRange(1,1,1,10).setBackground('#1A1A18').setFontColor('#FFFFFF').setFontWeight('bold');
  }

  SpreadsheetApp.flush();
  Logger.log('初始化完成！');
  return '✅ 初始化完成';
}
