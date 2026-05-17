// ================================================================
// 家族信用卡福利管理 - Google Apps Script 後端
//
// 使用步驟：
// 1. 打開 https://script.google.com → 新增專案
// 2. 把這整份程式碼貼入
// 3. 修改下方 SHEET_ID 為你的試算表 ID
// 4. 點「部署」→「新增部署」→「網頁應用程式」
//    執行身份：「我」｜存取者：「所有人」
// 5. 複製部署 URL，貼到 index.html 的設定頁面
// ================================================================

const SHEET_ID = '1PhCvipzoZi-zDI4YXy-jGo16fLvQ9yPWDKfwL9Aat_8';
const COMPLETIONS_TAB = 'BenefitCompletions';
const CARD_STATUS_TAB = 'CardStatus';

// ----------------------------------------------------------------
// 取得或建立分頁
// ----------------------------------------------------------------
function getSheet(tabName, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function getCompletionsSheet() {
  return getSheet(COMPLETIONS_TAB, ['key', 'completedAt', 'completedBy', 'deleted']);
}

function getCardStatusSheet() {
  return getSheet(CARD_STATUS_TAB, ['person', 'card', 'openDate', 'closeDate', 'notes', 'updatedAt']);
}

// ----------------------------------------------------------------
// GET：讀取資料
// ----------------------------------------------------------------
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  // 讀取福利完成紀錄
  if (action === 'getAll') {
    const sheet = getCompletionsSheet();
    const rows = sheet.getDataRange().getValues();
    const completions = {};
    for (let i = 1; i < rows.length; i++) {
      const [key, at, by, deleted] = rows[i];
      if (key && !deleted) {
        completions[key] = { at, by };
      }
    }
    return jsonResponse({ completions, ts: new Date().toISOString() });
  }

  // 讀取卡片開關卡狀態
  if (action === 'getCardStatus') {
    const sheet = getCardStatusSheet();
    const rows = sheet.getDataRange().getValues();
    const cards = {};
    for (let i = 1; i < rows.length; i++) {
      const [person, card, openDate, closeDate, notes] = rows[i];
      if (person && card) {
        const key = `${person}::${card}`;
        cards[key] = {
          openDate: openDate ? fmtDate(openDate) : null,
          closeDate: closeDate ? fmtDate(closeDate) : null,
          notes: notes || ''
        };
      }
    }
    return jsonResponse({ cards, ts: new Date().toISOString() });
  }

  // 透過 GET 參數寫入卡片開關卡（批次匯入用）
  if (action === 'setCardStatusGet') {
    const key = e.parameter.key || '';
    const openDate = e.parameter.openDate || '';
    const closeDate = e.parameter.closeDate || '';
    const notes = e.parameter.notes || '';
    if (key) {
      const sheet = getCardStatusSheet();
      const rows = sheet.getDataRange().getValues();
      const [person, card] = key.split('::');
      let found = false;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === person && rows[i][1] === card) {
          sheet.getRange(i + 1, 1, 1, 6).setValues([[person, card, openDate, closeDate, notes, new Date().toISOString()]]);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([person, card, openDate, closeDate, notes, new Date().toISOString()]);
    }
    return jsonResponse({ ok: true, key, openDate });
  }

  // 測試 fmtDate 函式
  if (action === 'testFmtDate') {
    const sheet = getCardStatusSheet();
    const rows = sheet.getDataRange().getValues();
    const d = rows[1][2];
    return jsonResponse({
      type: typeof d,
      isDate: d instanceof Date,
      fmtResult: fmtDate(d),
      utilsResult: d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'not a date',
      tz: Session.getScriptTimeZone(),
      raw: String(d)
    });
  }

  // debug：回傳 CardStatus 原始資料
  if (action === 'debugCardStatus') {
    const sheet = getCardStatusSheet();
    const rows = sheet.getDataRange().getValues();
    return jsonResponse({ rows: rows.slice(1, 5).map(r => ({
      person: r[0], card: r[1],
      openDate_raw: String(r[2]), openDate_type: typeof r[2],
      notes: r[4]
    }))});
  }

  return ContentService.createTextOutput('OK');
}

// ----------------------------------------------------------------
// POST：寫入資料
// ----------------------------------------------------------------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, key, val } = body;

    // 更新福利完成紀錄
    if (action === 'set' && key) {
      const sheet = getCompletionsSheet();
      const rows = sheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === key) {
          if (val) {
            sheet.getRange(i + 1, 1, 1, 4).setValues([[key, val.at, val.by, '']]);
          } else {
            sheet.getRange(i + 1, 4).setValue('deleted');
          }
          found = true;
          break;
        }
      }
      if (!found && val) {
        sheet.appendRow([key, val.at, val.by, '']);
      }
    }

    // 更新卡片開關卡狀態
    if (action === 'setCardStatus' && key) {
      const sheet = getCardStatusSheet();
      const rows = sheet.getDataRange().getValues();
      const [person, card] = key.split('::');
      let found = false;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === person && rows[i][1] === card) {
          if (val) {
            sheet.getRange(i + 1, 1, 1, 6).setValues([[
              person, card,
              val.openDate || '', val.closeDate || '',
              val.notes || '', new Date().toISOString()
            ]]);
          } else {
            sheet.deleteRow(i + 1);
          }
          found = true;
          break;
        }
      }
      if (!found && val) {
        sheet.appendRow([person, card, val.openDate || '', val.closeDate || '', val.notes || '', new Date().toISOString()]);
      }
    }

    return jsonResponse({ ok: true });

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ----------------------------------------------------------------
// 工具函式
// ----------------------------------------------------------------
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function fmtDate(d) {
  if (!d) return null;
  if (typeof d === 'string') return d.slice(0, 10);
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) {
    return null;
  }
}
