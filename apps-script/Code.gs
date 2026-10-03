/**
 * 습관 트래커 동기화 서버 (Google Apps Script + Google Sheets)
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드를 붙여넣기
 *       → 아래 TOKEN을 본인만 아는 긴 문자열로 변경 → 웹 앱으로 배포
 */
const TOKEN = 'CHANGE_ME_TO_A_LONG_RANDOM_STRING';

const HABITS_SHEET = 'Habits'; // id, name, deleted, order, updatedAt
const LOGS_SHEET = 'Logs';     // key("habitId|YYYY-MM-DD"), done, updatedAt

function doGet() {
  return json_({ ok: true, message: 'habit-tracker sync server' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad request' });
  }
  if (body.token !== TOKEN || TOKEN === 'CHANGE_ME_TO_A_LONG_RANDOM_STRING') {
    return json_({ ok: false, error: 'unauthorized' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const habits = readHabits_();
    const logs = readLogs_();

    // 클라이언트가 보낸 레코드를 병합 (updatedAt이 더 큰 쪽이 승리)
    (body.habits || []).forEach(h => {
      const cur = habits[h.id];
      if (!cur || h.t > cur.t) habits[h.id] = { id: String(h.id), name: String(h.name), deleted: !!h.deleted, o: Number(h.o) || 0, t: Number(h.t) };
    });
    (body.logs || []).forEach(l => {
      const cur = logs[l.k];
      if (!cur || l.t > cur.t) logs[l.k] = { k: String(l.k), done: !!l.done, t: Number(l.t) };
    });

    writeHabits_(habits);
    writeLogs_(logs);
    return json_({ ok: true, habits: Object.values(habits), logs: Object.values(logs) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function readHabits_() {
  const sh = sheet_(HABITS_SHEET, ['id', 'name', 'deleted', 'order', 'updatedAt']);
  const out = {};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    out[String(r[0])] = { id: String(r[0]), name: String(r[1]), deleted: r[2] === true || r[2] === 'TRUE', o: Number(r[3]) || 0, t: Number(r[4]) || 0 };
  });
  return out;
}

function readLogs_() {
  const sh = sheet_(LOGS_SHEET, ['key', 'done', 'updatedAt']);
  const out = {};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    out[String(r[0])] = { k: String(r[0]), done: r[1] === true || r[1] === 'TRUE', t: Number(r[2]) || 0 };
  });
  return out;
}

function writeHabits_(habits) {
  const sh = sheet_(HABITS_SHEET, ['id', 'name', 'deleted', 'order', 'updatedAt']);
  const rows = Object.values(habits).map(h => [h.id, h.name, h.deleted, h.o, h.t]);
  sh.clearContents();
  sh.getRange(1, 1, 1, 5).setValues([['id', 'name', 'deleted', 'order', 'updatedAt']]);
  if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
}

function writeLogs_(logs) {
  const sh = sheet_(LOGS_SHEET, ['key', 'done', 'updatedAt']);
  const rows = Object.values(logs).map(l => [l.k, l.done, l.t]);
  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([['key', 'done', 'updatedAt']]);
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
