/**
 * 과제 제출 트래커 — Google Apps Script 웹 앱
 *
 * - doGet / doPost : GitHub Pages 프런트엔드가 호출하는 JSON API
 * - setup()          : 스프레드시트·시트·스크립트 속성 초기화 (최초 1회 수동 실행)
 * - installTriggers(): 매일 09시 sendDailyDigest 트리거 등록 (최초 1회 수동 실행)
 * - sendDailyDigest(): 운영자 요약 메일 + 미제출자 리마인더 메일 발송
 */

var SHEET_ROSTER = 'Roster';
var SHEET_SUBMISSIONS = 'Submissions';
var ROSTER_HEADER = ['이름', '이메일'];
var SUBMISSION_HEADER = ['제출시각', '회차', '이름', '이메일', '제목', '링크', '메모'];
var TZ = 'Asia/Seoul';

// ───────────────────────── 설정 ─────────────────────────

function props_() {
  return PropertiesService.getScriptProperties();
}

function config_() {
  var p = props_().getProperties();
  return {
    spreadsheetId: p.SPREADSHEET_ID || '',
    currentRound: Number(p.CURRENT_ROUND || 6),
    deadline: p.DEADLINE || '',            // 예: 2026-10-19T23:59:00+09:00
    ownerEmail: p.OWNER_EMAIL || Session.getEffectiveUser().getEmail(),
    remindStudents: p.REMIND_STUDENTS !== 'false'
  };
}

function db_() {
  var id = config_().spreadsheetId;
  if (!id) throw new Error('SPREADSHEET_ID 가 없습니다. setup() 을 먼저 실행하세요.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  var sh = db_().getSheetByName(name);
  if (!sh) throw new Error('시트가 없습니다: ' + name);
  return sh;
}

// ───────────────────────── 초기화 ─────────────────────────

function setup() {
  var p = props_();
  var ss;
  var id = p.getProperty('SPREADSHEET_ID');
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('과제 제출 트래커 DB');
    p.setProperty('SPREADSHEET_ID', ss.getId());
  }
  ensureSheet_(ss, SHEET_ROSTER, ROSTER_HEADER);
  ensureSheet_(ss, SHEET_SUBMISSIONS, SUBMISSION_HEADER);

  var defaults = { CURRENT_ROUND: '6', REMIND_STUDENTS: 'true' };
  Object.keys(defaults).forEach(function (k) {
    if (!p.getProperty(k)) p.setProperty(k, defaults[k]);
  });

  var first = ss.getSheets()[0];
  if (first.getName() === 'Sheet1' || first.getName() === '시트1') ss.deleteSheet(first);

  Logger.log('DB: %s', ss.getUrl());
  return ss.getUrl();
}

function ensureSheet_(ss, name, header) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyDigest').timeBased().everyDays(1).atHour(9).inTimezone(TZ).create();
}

// ───────────────────────── HTTP 엔드포인트 ─────────────────────────

function doGet(e) {
  var q = (e && e.parameter) || {};
  try {
    switch (q.action || 'summary') {
      case 'summary':
        return json_({ ok: true, data: summary_(q.round ? Number(q.round) : config_().currentRound) });
      case 'recent':
        return json_({ ok: true, data: recent_(Math.min(Number(q.limit) || 20, 100)) });
      case 'ping':
        return json_({ ok: true, data: { now: new Date().toISOString() } });
      default:
        return json_({ ok: false, error: '알 수 없는 action: ' + q.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// 프런트엔드는 CORS preflight 를 피하려고 Content-Type: text/plain 으로 JSON 을 보낸다.
function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action !== 'submit') return json_({ ok: false, error: '알 수 없는 action' });
    return json_({ ok: true, data: submit_(body) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── 도메인 로직 ─────────────────────────

function submit_(b) {
  var round = Number(b.round);
  var name = String(b.name || '').trim();
  var email = String(b.email || '').trim().toLowerCase();
  var title = String(b.title || '').trim();
  var link = String(b.link || '').trim();
  var note = String(b.note || '').trim();

  if (!(round >= 1 && round <= 30)) throw new Error('회차가 올바르지 않습니다.');
  if (!name || name.length > 40) throw new Error('이름을 확인하세요.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('이메일 형식을 확인하세요.');
  if (!title || title.length > 120) throw new Error('제목을 확인하세요.');
  if (!/^https?:\/\/\S+$/i.test(link) || link.length > 500) throw new Error('링크는 http(s) 주소여야 합니다.');
  if (note.length > 500) throw new Error('메모는 500자 이내로 입력하세요.');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var now = new Date();
    // 시트 수식 주입 방지: = + - @ 로 시작하는 값은 문자열로 고정
    sheet_(SHEET_SUBMISSIONS).appendRow([now, round, safe_(name), email, safe_(title), link, safe_(note)]);
    upsertRoster_(name, email);
    return { round: round, submittedAt: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function upsertRoster_(name, email) {
  var sh = sheet_(SHEET_ROSTER);
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues() : [];
  var exists = rows.some(function (r) { return String(r[0]).toLowerCase() === email; });
  if (!exists) sh.appendRow([safe_(name), email]);
}

function readRows_(name) {
  var sh = sheet_(name);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

function summary_(round) {
  var cfg = config_();
  var roster = readRows_(SHEET_ROSTER).map(function (r) {
    return { name: String(r[0]), email: String(r[1]).toLowerCase() };
  });

  var latest = {};
  readRows_(SHEET_SUBMISSIONS).forEach(function (r) {
    if (Number(r[1]) !== round) return;
    var email = String(r[3]).toLowerCase();
    var at = r[0] instanceof Date ? r[0] : new Date(r[0]);
    if (!latest[email] || latest[email].at < at) latest[email] = { at: at, title: String(r[4]), link: String(r[5]) };
  });

  var people = roster.map(function (p) {
    var s = latest[p.email];
    return {
      name: p.name,
      email: maskEmail_(p.email),
      submitted: !!s,
      submittedAt: s ? s.at.toISOString() : null,
      title: s ? s.title : null,
      link: s ? s.link : null
    };
  });

  return {
    round: round,
    currentRound: cfg.currentRound,
    deadline: cfg.deadline || null,
    total: people.length,
    submitted: people.filter(function (p) { return p.submitted; }).length,
    people: people
  };
}

function recent_(limit) {
  return readRows_(SHEET_SUBMISSIONS)
    .map(function (r) {
      return {
        at: (r[0] instanceof Date ? r[0] : new Date(r[0])).toISOString(),
        round: Number(r[1]),
        name: String(r[2]),
        title: String(r[4]),
        link: String(r[5])
      };
    })
    .sort(function (a, b) { return a.at < b.at ? 1 : -1; })
    .slice(0, limit);
}

// 공개 API 에서는 이메일 전체를 노출하지 않는다.
function maskEmail_(email) {
  var parts = email.split('@');
  if (parts.length !== 2) return '';
  var id = parts[0];
  return id.slice(0, 2) + '***@' + parts[1];
}

// ───────────────────────── 자동화: 일일 알림 ─────────────────────────

function sendDailyDigest() {
  var cfg = config_();
  var s = summary_(cfg.currentRound);
  var fullRoster = readRows_(SHEET_ROSTER).map(function (r) {
    return { name: String(r[0]), email: String(r[1]).toLowerCase() };
  });
  var pending = fullRoster.filter(function (p, i) { return !s.people[i].submitted; });
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var deadlineText = cfg.deadline
    ? Utilities.formatDate(new Date(cfg.deadline), TZ, 'M월 d일 HH:mm')
    : '미정';

  var body = [
    '[' + today + '] ' + s.round + '회차 제출 현황',
    '',
    '제출 ' + s.submitted + ' / ' + s.total + '명 · 마감 ' + deadlineText,
    '',
    '미제출: ' + (pending.length ? pending.map(function (p) { return p.name; }).join(', ') : '없음'),
    '',
    'DB: ' + db_().getUrl()
  ].join('\n');
  MailApp.sendEmail(cfg.ownerEmail, '[과제 트래커] ' + s.round + '회차 제출 ' + s.submitted + '/' + s.total, body);

  if (!cfg.remindStudents || !cfg.deadline) return;
  var msLeft = new Date(cfg.deadline).getTime() - Date.now();
  if (msLeft <= 0 || msLeft > 3 * 24 * 3600 * 1000) return; // 마감 3일 전부터만 리마인드

  pending.forEach(function (p) {
    if (MailApp.getRemainingDailyQuota() < 1) return;
    MailApp.sendEmail(
      p.email,
      '[리마인더] ' + s.round + '회차 과제 마감 ' + deadlineText,
      p.name + '님, ' + s.round + '회차 과제가 아직 제출되지 않았습니다.\n마감: ' + deadlineText
    );
  });
}
