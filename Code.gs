const SHEET_RECORDS = 'Records';
const SHEET_USERS = 'Users';

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const action = (e && e.parameter && e.parameter.action) || 'pull_all';
    const recSheet = getOrCreateSheet(SHEET_RECORDS, ['id', 'payload_json']);
    const usrSheet = getOrCreateSheet(SHEET_USERS, ['id', 'username', 'password', 'payload_json']);

    const records = readRecords(recSheet);
    const users = readUsers(usrSheet);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, action, records, users }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const body = JSON.parse(raw);
    const action = body.action || '';

    const recSheet = getOrCreateSheet(SHEET_RECORDS, ['id', 'payload_json']);
    const usrSheet = getOrCreateSheet(SHEET_USERS, ['id', 'username', 'password', 'payload_json']);

    if (action === 'append_record' && body.row) {
      upsertRecord(recSheet, body.row);
    } else if (action === 'delete_record' && body.id) {
      deleteRecordById(recSheet, body.id);
    } else if (action === 'clear_records') {
      clearSheetData(recSheet);
    } else if (action === 'bulk_sync') {
      if (Array.isArray(body.rows)) {
        replaceAllRecords(recSheet, body.rows);
      }
      if (Array.isArray(body.users)) {
        replaceAllUsers(usrSheet, body.users);
      }
    } else if (action === 'sync_users' && Array.isArray(body.users)) {
      replaceAllUsers(usrSheet, body.users);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, action }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function readRecords(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const list = [];
  for (const [id, jsonStr] of values) {
    if (!jsonStr) continue;
    try {
      const obj = JSON.parse(jsonStr);
      if (!obj.id && id) obj.id = id;
      list.push(obj);
    } catch (_) {}
  }
  return list;
}

function upsertRecord(sheet, row) {
  const id = String(row.id || ('REC-' + Date.now()));
  row.id = id;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const idx = ids.findIndex(v => String(v) === id);
    if (idx !== -1) {
      sheet.getRange(idx + 2, 2).setValue(JSON.stringify(row));
      return;
    }
  }
  sheet.appendRow([id, JSON.stringify(row)]);
}

function deleteRecordById(sheet, targetId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i]) === String(targetId)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function clearSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

function replaceAllRecords(sheet, rows) {
  clearSheetData(sheet);
  if (!rows.length) return;
  const out = rows.map(r => {
    const id = String(r.id || ('REC-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6)));
    return [id, JSON.stringify({ ...r, id })];
  });
  sheet.getRange(2, 1, out.length, 2).setValues(out);
}

function readUsers(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const list = [];
  for (const [id, username, password, jsonStr] of values) {
    if (jsonStr) {
      try {
        list.push(JSON.parse(jsonStr));
        continue;
      } catch (_) {}
    }
    if (username) {
      list.push({ id: id || ('USR-' + Date.now()), username: String(username), password: String(password || '') });
    }
  }
  return list;
}

function replaceAllUsers(sheet, users) {
  clearSheetData(sheet);
  if (!users.length) return;
  const out = users.map(u => {
    const id = String(u.id || ('USR-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6)));
    const payload = { ...u, id };
    return [id, String(u.username || ''), String(u.password || ''), JSON.stringify(payload)];
  });
  sheet.getRange(2, 1, out.length, 4).setValues(out);
}
