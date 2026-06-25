// ================================================
// 點名系統 Apps Script
// 部署為網頁應用程式：
//   執行身分 → 我
//   存取權   → 所有人
// ================================================

const SHEET_CHECKIN = '簽到紀錄';
const SHEET_SEATS   = '座位表';

// ── POST：儲存座位表 或 學生簽到 ──
function doPost(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'saveSeats') {
      saveSeats_(data);
      res.setContent(JSON.stringify({ ok: true }));
    } else if (data.action === 'checkin') {
      const result = saveCheckin_(data);
      res.setContent(JSON.stringify({ ok: true, ...result }));
    } else {
      res.setContent(JSON.stringify({ ok: false, msg: '未知 action' }));
    }
  } catch(err) {
    res.setContent(JSON.stringify({ ok: false, msg: err.message }));
  }
  return res;
}

// ── GET：讀取座位表 或 簽到紀錄 ──
function doGet(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);
  try {
    const action = e.parameter.action || 'getCheckins';

    if (action === 'getSeats') {
      const cls = e.parameter.cls;
      res.setContent(JSON.stringify({ ok: true, data: getSeats_(cls) }));
    } else if (action === 'getCheckins') {
      const token = e.parameter.token;
      const cls   = e.parameter.cls;
      res.setContent(JSON.stringify({ ok: true, data: getCheckins_(token, cls) }));
    } else {
      res.setContent(JSON.stringify({ ok: false, msg: '未知 action' }));
    }
  } catch(err) {
    res.setContent(JSON.stringify({ ok: false, msg: err.message }));
  }
  return res;
}

// ── 儲存座位表 ──
function saveSeats_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_SEATS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SEATS);
    sheet.appendRow(['班級','座位','座號','姓名','學號']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,5).setBackground('#18181A').setFontColor('#fff').setFontWeight('bold');
  }

  const cls   = data.cls;
  const seats = data.seats; // {seat: {sno,name,sid}}

  // 刪除此班舊資料
  const all = sheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = all.length - 1; i >= 1; i--) {
    if (all[i][0] === cls) toDelete.push(i + 1);
  }
  toDelete.forEach(row => sheet.deleteRow(row));

  // 寫入新資料
  Object.entries(seats).forEach(([seat, info]) => {
    sheet.appendRow([cls, seat, info.sno || '', info.name || '', info.sid || '']);
  });
  SpreadsheetApp.flush();
}

// ── 讀取座位表 ──
function getSeats_(cls) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SEATS);
  if (!sheet) return {};

  const all    = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < all.length; i++) {
    const [rowCls, seat, sno, name, sid] = all[i];
    if (!cls || rowCls === cls) {
      result[seat] = { sno: String(sno), name: String(name), sid: String(sid || '') };
    }
  }
  return result;
}

// ── 儲存簽到 ──
function saveCheckin_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_CHECKIN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CHECKIN);
    sheet.appendRow(['日期','課程','班級','token','座位','座號','姓名','學號','時間','狀態']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,10).setBackground('#18181A').setFontColor('#fff').setFontWeight('bold');
  }

  // 防重複
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (all[i][3] === data.token && all[i][4] === data.seat) {
      return { duplicate: true };
    }
  }

  sheet.appendRow([
    data.date    || '',
    data.course  || '',
    data.cls     || '',
    data.token   || '',
    data.seat    || '',
    data.sno     || '',
    data.name    || '',
    data.sid     || '',
    data.timeStr || '',
    data.late ? '遲到' : '準時'
  ]);
  SpreadsheetApp.flush();
  return { status: data.late ? '遲到' : '準時' };
}

// ── 讀取簽到紀錄 ──
function getCheckins_(token, cls) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CHECKIN);
  if (!sheet) return {};

  const all    = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < all.length; i++) {
    if (all[i][3] === token && all[i][2] === cls) {
      const seat = all[i][4];
      result[seat] = {
        date: all[i][0], course: all[i][1], cls: all[i][2],
        token: all[i][3], seat, sno: String(all[i][5]),
        name: String(all[i][6]), sid: String(all[i][7] || ''),
        timeStr: all[i][8], late: all[i][9] === '遲到'
      };
    }
  }
  return result;
}
