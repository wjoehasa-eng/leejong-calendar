/**
 * 이종대왕 달력 — 구글 시트 백업/동기화 서버 (Google Apps Script)
 * ============================================================
 * 쓰는 법 (5분, 한 번만)
 * 1. 구글 드라이브에서 새 스프레드시트를 하나 만듭니다. (예: "이종대왕 달력 백업")
 * 2. 그 시트에서 확장 프로그램 → Apps Script 를 엽니다.
 * 3. 기본으로 열린 Code.gs 내용을 모두 지우고, 이 파일 내용을 그대로 붙여 넣습니다.
 * 4. 저장(💾) → 배포 → 새 배포 → 유형 선택(⚙) = 웹 앱
 *    · 설명: 이종대왕 달력 백업
 *    · 다음 사용자로 실행: 나
 *    · 액세스 권한: 모든 사용자        ← 이게 "모든 사용자"여야 폰 앱에서 접속됩니다
 *    → 배포 → 권한 검토/허용(내 계정) → 웹 앱 URL 복사
 * 5. 복사한 URL(.../exec)을 달력 앱 설정 → "구글 시트 클라우드 동기화" 칸에 붙여 넣고
 *    "시트로 저장 (백업)" 을 누릅니다.
 *
 * 코드를 고친 뒤에는 반드시 배포 → 배포 관리 → 수정(연필) → 버전: 새 버전 → 배포.
 *
 * 시트 구성 (자동으로 만들어집니다)
 *  · 이벤트 : 기록 1개 = 1행. 엑셀처럼 직접 보고 고칠 수 있습니다.
 *             (id, date, end, title, done, star, tags, desc, links, raw)
 *  · 백업   : 통째로 저장한 원본 JSON 을 45,000자씩 잘라 담습니다(셀 5만자 한도 회피).
 *             '이벤트' 시트가 지워져도 여기서 완전히 복구됩니다.
 *  · 메타   : 마지막 저장 시각, 기록 수 등.
 *  · 사용법 : 처음 한 번 만들어 두는 짧은 안내.
 *
 * 개인정보: 이 스크립트는 사용자 본인의 스프레드시트에만 쓰고, 다른 곳으로 보내지 않습니다.
 */

/** 기록 1개를 1행으로 저장하는 시트 이름 */
const SHEET_EVENTS = '이벤트';
/** 원본 JSON 통째 백업(청크) 시트 이름 (구글 시트 셀 한도 5만자 회피) */
const SHEET_BACKUP = '백업';
/** 마지막 저장 시각·개수 시트 이름 */
const SHEET_META = '메타';
/** 처음 한 번만 만드는 안내 시트 이름 */
const SHEET_HELP = '사용법';
/** 한 셀에 넣을 최대 글자 수 (구글 시트 한도 50,000자보다 여유 있게) */
const CHUNK_SIZE = 45000;
/** 이벤트 시트 열 이름 */
const EVENT_HEADERS = ['id', 'date', 'end', 'title', 'done', 'star', 'tags', 'desc', 'links', 'raw'];

/* ------------------------------------------------------------------ 라우팅 */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'status';
  try {
    if (action === 'load') return json_(loadAll_());
    if (action === 'ping') return json_({ ok: true, now: new Date().toISOString() });
    return json_(status_());
  } catch (err) {
    return json_({ ok: false, error: errText_(err) });
  }
}

function doPost(e) {
  var req = {};
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '요청 본문(JSON)을 읽지 못했습니다.' });
  }
  try {
    var action = req.action || 'ping';
    if (action === 'save') return json_(saveAll_(req));
    if (action === 'load') return json_(loadAll_());
    return json_(status_());
  } catch (err) {
    return json_({ ok: false, error: errText_(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errText_(err) {
  return String((err && err.message) || err || '알 수 없는 오류');
}

/* ------------------------------------------------------------------ 저장 */

function saveAll_(req) {
  var events = req.events;
  if (!(events instanceof Array)) throw new Error('events 배열이 없습니다.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = ss_();
    var now = new Date();
    var savedAt = req.savedAt || now.toISOString();

    // 1) 사람이 읽는 표: 기록 1개 = 1행
    writeEventRows_(ss, events);

    // 2) 손실 없는 원본: 통째 JSON 을 청크로
    var payload = {
      app: req.app || '이종대왕 달력',
      version: req.version || 1,
      calendar: req.calendar || '이종대왕 달력',
      savedAt: savedAt,
      count: events.length,
      events: events
    };
    writeChunks_(ss, JSON.stringify(payload));

    // 3) 사용법(없을 때만) + 메타
    ensureHelp_(ss);
    writeMeta_(ss, {
      savedAt: savedAt,
      updatedAt: now.toISOString(),
      count: String(events.length),
      app: payload.app
    });

    return {
      ok: true,
      count: events.length,
      savedAt: savedAt,
      updatedAt: now.toISOString(),
      sheetUrl: ss.getUrl(),
      spreadsheetId: ss.getId()
    };
  } finally {
    lock.releaseLock();
  }
}

function writeEventRows_(ss, events) {
  var sh = ensureSheet_(ss, SHEET_EVENTS, EVENT_HEADERS);
  var rows = events.map(function (ev) { return eventToRow_(ev); });
  var width = EVENT_HEADERS.length;
  if (rows.length) {
    var startRow = 2;
    // 기존 데이터 지우기(헤더 유지)
    var last = sh.getLastRow();
    if (last >= startRow) sh.getRange(startRow, 1, last - startRow + 1, width).clearContent();
    // 날짜 열은 글자로 고정해 두어야 Date 로 바뀌지 않습니다
    sh.getRange(startRow, 1, rows.length, width).setNumberFormat('@');
    sh.getRange(startRow, 1, rows.length, width).setValues(rows);
    sh.getRange(startRow, 1, rows.length, width).setNumberFormat('@');
  } else {
    var last2 = sh.getLastRow();
    if (last2 >= 2) sh.getRange(2, 1, last2 - 1, width).clearContent();
  }
  sh.setFrozenRows(1);
}

function eventToRow_(ev) {
  ev = ev || {};
  return [
    str_(ev.id),
    str_(ev.date),
    str_(ev.end),
    str_(ev.title),
    ev.done ? 'TRUE' : 'FALSE',
    ev.star ? 'TRUE' : 'FALSE',
    (ev.tags || []).join(' '),
    str_(ev.desc),
    (ev.links && ev.links.length) ? JSON.stringify(ev.links) : '',
    JSON.stringify(ev)
  ];
}

function writeChunks_(ss, text) {
  var sh = ensureSheet_(ss, SHEET_BACKUP, ['chunk']);
  sh.clearContents();
  sh.getRange(1, 1, 1, 2).setValues([['chunk', 'savedChars']]);
  var chunks = [];
  for (var i = 0; i < text.length; i += CHUNK_SIZE) chunks.push([text.slice(i, i + CHUNK_SIZE)]);
  if (chunks.length) {
    sh.getRange(1, 1, 1, 1).setValue('chunk');            // A1 머리글 유지
    sh.getRange(1, 2).setValue(String(text.length));
    sh.getRange(2, 1, chunks.length, 1).setNumberFormat('@').setValues(chunks);
  }
  sh.setFrozenRows(1);
}

function ensureHelp_(ss) {
  var sh = ss.getSheetByName(SHEET_HELP);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_HELP);
  var lines = [
    ['이종대왕 달력 — 구글 시트 백업', ''],
    ['', ''],
    ['이 스프레드시트는 달력 앱의 클라우드 백업입니다.', ''],
    ['앱: 설정(⚙) → 구글 시트 클라우드 동기화 → 시트로 저장 / 시트에서 불러오기', ''],
    ['', ''],
    ['시트 이름', '설명'],
    [SHEET_EVENTS, '기록 1개 = 1행. 여기서 제목·날짜·완료 여부를 직접 고쳐도 앱에서 불러올 때 반영됩니다.'],
    [SHEET_BACKUP, '통째로 저장한 원본 JSON(45,000자씩 나눠 담음). 이벤트 시트가 지워져도 여기서 완전히 복구됩니다.'],
    [SHEET_META, '마지막 저장 시각, 기록 수'],
    ['', ''],
    ['주의', '이벤트 시트에서 행을 지우면 그 기록은 다음 불러오기에서 사라집니다. 지우기 전에 앱에서 타임머신 스냅샷을 한 번 저장해 두세요.'],
    ['스크립트 수정', '코드를 고친 뒤에는 배포 → 배포 관리 → 수정 → 버전: 새 버전 → 배포 를 해야 앱에 반영됩니다.']
  ];
  sh.getRange(1, 1, lines.length, 2).setValues(lines);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.getRange(6, 1, 1, 2).setFontWeight('bold');
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 420);
  return sh;
}

function writeMeta_(ss, meta) {
  var sh = ensureSheet_(ss, SHEET_META, ['key', 'value']);
  sh.clearContents();
  var rows = [['key', 'value']];
  Object.keys(meta).forEach(function (k) { rows.push([k, meta[k]]); });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.setFrozenRows(1);
}

/* ------------------------------------------------------------------ 불러오기 */

function loadAll_() {
  var ss = ss_();
  var meta = readMeta_(ss);
  var events = readEventRows_(ss);
  var source = SHEET_EVENTS;

  // '이벤트' 시트가 비었으면 청크 백업에서 복구
  if (!events.length) {
    var fromChunks = readChunks_(ss);
    if (fromChunks && fromChunks.events instanceof Array) {
      events = fromChunks.events;
      source = SHEET_BACKUP;
      if (fromChunks.savedAt) meta.savedAt = fromChunks.savedAt;
    }
  }

  return {
    ok: true,
    app: meta.app || '이종대왕 달력',
    count: events.length,
    savedAt: meta.savedAt || '',
    updatedAt: meta.updatedAt || '',
    source: source,
    sheetUrl: ss.getUrl(),
    events: events
  };
}

function readEventRows_(ss) {
  var sh = ss.getSheetByName(SHEET_EVENTS);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var width = EVENT_HEADERS.length;
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var ev = rowToEvent_(values[i]);
    if (ev) out.push(ev);
  }
  return out;
}

function rowToEvent_(row) {
  var base = {};
  try { base = JSON.parse(row[9] || '{}') || {}; } catch (e) { base = {}; }

  var ev = {};
  Object.keys(base).forEach(function (k) { ev[k] = base[k]; });

  var id = str_(row[0]) || str_(base.id);
  var date = ymd_(row[1]) || str_(base.date);
  if (!date) return null;                       // 날짜가 없으면 기록이 아님

  ev.id = id || ('e' + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 7));
  ev.date = date;

  var end = ymd_(row[2]);
  if (end) ev.end = end; else delete ev.end;

  var title = str_(row[3]);
  ev.title = title !== '' ? title : str_(base.title);

  ev.done = truthy_(row[4]);
  ev.star = truthy_(row[5]) || ev.title.indexOf('★') >= 0;

  var tags = str_(row[6]).trim();
  if (tags) ev.tags = tags.split(/[\s,]+/).filter(function (t) { return !!t; });
  else delete ev.tags;

  var desc = str_(row[7]);
  if (desc) ev.desc = desc; else delete ev.desc;

  var links = str_(row[8]).trim();
  if (links) {
    try {
      var parsed = JSON.parse(links);
      if (parsed instanceof Array) ev.links = parsed; else delete ev.links;
    } catch (e) { /* 표에서 직접 고치다 깨졌으면 raw 의 값을 그대로 둔다 */ }
  } else {
    delete ev.links;
  }
  return ev;
}

function readChunks_(ss) {
  var sh = ss.getSheetByName(SHEET_BACKUP);
  if (!sh) return null;
  var last = sh.getLastRow();
  if (last < 2) return null;
  var values = sh.getRange(2, 1, last - 1, 1).getValues();
  var text = values.map(function (r) { return str_(r[0]); }).join('');
  if (!text) return null;
  try { return JSON.parse(text); } catch (e) { return null; }
}

function readMeta_(ss) {
  var sh = ss.getSheetByName(SHEET_META);
  var out = {};
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var values = sh.getRange(2, 1, last - 1, 2).getValues();
  values.forEach(function (r) {
    var k = str_(r[0]);
    if (k) out[k] = str_(r[1]);
  });
  return out;
}

/* ------------------------------------------------------------------ 상태 */

function status_() {
  var ss = ss_();
  var meta = readMeta_(ss);
  var count = 0;
  var sh = ss.getSheetByName(SHEET_EVENTS);
  if (sh) count = Math.max(0, sh.getLastRow() - 1);
  return {
    ok: true,
    app: meta.app || '이종대왕 달력',
    count: count,
    savedAt: meta.savedAt || '',
    updatedAt: meta.updatedAt || '',
    sheetUrl: ss.getUrl(),
    message: '이종대왕 달력 백업 통로입니다. 앱의 설정 → 구글 시트 클라우드 동기화에서 사용하세요.'
  };
}

/* ------------------------------------------------------------------ 도우미 */

/** 시트에 묶여 있으면 그 시트를, 아니면(독립 스크립트) 전용 스프레드시트를 찾거나 만듭니다. */
function ss_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 지워졌으면 새로 만든다 */ }
  }
  var created = SpreadsheetApp.create('이종대왕 달력 백업');
  props.setProperty('SPREADSHEET_ID', created.getId());
  return created;
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && headers.length) {
    var first = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    var same = headers.every(function (h, i) { return str_(first[i]) === h; });
    if (!same) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sh;
}

function str_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  return String(v);
}

/** Date 로 바뀐 셀도 yyyy-MM-dd 로 되살립니다. */
function ymd_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  var s = String(v).trim();
  var m = s.match(/(\d{4})[.\-\/]\s*(\d{1,2})[.\-\/]\s*(\d{1,2})/);
  if (!m) return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function truthy_(v) {
  if (v === true) return true;
  var s = String(v === null || v === undefined ? '' : v).trim().toUpperCase();
  return s === 'TRUE' || s === 'Y' || s === 'YES' || s === 'O' || s === '✓' || s === '1';
}
