// ================================================
// 點名系統 Apps Script
// 貼到 script.google.com 新專案，部署為網頁應用程式
// ================================================

const SHEET_NAME = '簽到紀錄';

function doPost(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // 第一次自動建立工作表
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['日期','課程','班級','token','座位','座號','姓名','學號','時間','狀態']);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,10)
           .setBackground('#18181A')
           .setFontColor('#ffffff')
           .setFontWeight('bold');
    }

    // 防重複：同 token + 座位只能一筆
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][3] === data.token && all[i][4] === data.seat) {
        res.setContent(JSON.stringify({ ok: true, duplicate: true }));
        return res;
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
    res.setContent(JSON.stringify({ ok: true }));
  } catch(err) {
    res.setContent(JSON.stringify({ ok: false, msg: err.message }));
  }
  return res;
}

function doGet(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);
  try {
    const token = e.parameter.token;
    const cls   = e.parameter.cls;
    if (!token || !cls) {
      res.setContent(JSON.stringify({ ok: false, msg: '缺少參數' }));
      return res;
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      res.setContent(JSON.stringify({ ok: true, data: {} }));
      return res;
    }

    const all  = sheet.getDataRange().getValues();
    const data = {};
    for (let i = 1; i < all.length; i++) {
      const row = all[i];
      if (row[3] === token && row[2] === cls) {
        const seat = row[4];
        data[seat] = {
          date:    row[0],
          course:  row[1],
          cls:     row[2],
          token:   row[3],
          seat:    seat,
          sno:     row[5],
          name:    row[6],
          sid:     row[7],
          timeStr: row[8],
          late:    row[9] === '遲到'
        };
      }
    }
    res.setContent(JSON.stringify({ ok: true, data }));
  } catch(err) {
    res.setContent(JSON.stringify({ ok: false, msg: err.message }));
  }
  return res;
}
