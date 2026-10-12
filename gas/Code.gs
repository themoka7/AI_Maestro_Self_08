/**
 * AX·정보화혁신본부 주간보고 — Apps Script 웹 앱
 *
 * 최초 1회 수동 실행: setup() → installTriggers()
 * 조직 데이터는 저장소에 없다. setup() 이 DB 시트 / ORG_SHEET_ID / ORG_JSON / INITIAL_ADMIN 순으로 찾는다.
 * 배포: 웹 앱 / 실행 사용자: 나 / 액세스: pusan.ac.kr 도메인 사용자
 *
 * 화면(Index.html)은 google.script.run.api(method, args) 하나로 서버를 호출하고,
 * 실제 로직·권한 검사는 Core.js 가 담당한다.
 */

var DOMAIN = 'pusan.ac.kr';
var TZ = 'Asia/Seoul';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AX·정보화혁신본부 주간보고')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function api(method, args) {
  var me = currentUser_();
  var lock = null;
  if (Core.isWrite(method)) {
    lock = LockService.getScriptLock();
    lock.waitLock(15000);
  }
  try {
    var result = Core.call(new SheetStore_(db_()), me, method, args);
    if (result.events.length) notify_(result.events);
    return JSON.parse(JSON.stringify(result.data));
  } finally {
    if (lock) lock.releaseLock();
  }
}

// 도메인 내부 배포에서는 접속자 이메일을 신뢰할 수 있다.
function currentUser_() {
  var email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  var at = email.lastIndexOf('@');
  if (at < 0) throw new Error('로그인 정보를 확인할 수 없습니다. ' + DOMAIN + ' 계정으로 접속하세요.');
  if (email.slice(at + 1) !== DOMAIN) throw new Error(DOMAIN + ' 계정만 사용할 수 있습니다.');
  return email.slice(0, at);
}

// ───────────────────────── 저장소 (Google Sheets) ─────────────────────────

function db_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_ID 가 없습니다. setup() 을 먼저 실행하세요.');
  return SpreadsheetApp.openById(id);
}

function SheetStore_(ss) {
  this.ss = ss;
  this.cache = {};
}

SheetStore_.prototype.sheet_ = function (t) {
  var sh = this.ss.getSheetByName(t);
  if (!sh) throw new Error('시트가 없습니다: ' + t + ' — setup() 을 다시 실행하세요.');
  return sh;
};

SheetStore_.prototype.load_ = function (t) {
  if (this.cache[t]) return this.cache[t];
  var sh = this.sheet_(t);
  var width = Object.keys(Core.SCHEMA[t]).length;
  var n = sh.getLastRow() - 1;
  var rows = n > 0 ? sh.getRange(2, 1, n, width).getDisplayValues() : [];
  var list = [];
  var rowOf = {};
  rows.forEach(function (r, i) {
    if (!r[0]) return;
    var o = Core.decode(t, r);
    list.push(o);
    rowOf[o.id] = i + 2;
  });
  this.cache[t] = { list: list, rowOf: rowOf };
  return this.cache[t];
};

SheetStore_.prototype.all = function (t) {
  return this.load_(t).list.map(function (o) { return Object.assign({}, o); });
};

SheetStore_.prototype.get = function (t, id) {
  var o = this.load_(t).list.filter(function (x) { return x.id === id; })[0];
  return o ? Object.assign({}, o) : null;
};

SheetStore_.prototype.save = function (t, obj) {
  var sh = this.sheet_(t);
  var row = Core.encode(t, obj);
  var at = this.load_(t).rowOf[obj.id] || sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  delete this.cache[t];
};

SheetStore_.prototype.remove = function (t, id) {
  var at = this.load_(t).rowOf[id];
  if (at) this.sheet_(t).deleteRow(at);
  delete this.cache[t];
};

// ───────────────────────── 초기화 ─────────────────────────

function setup() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('SPREADSHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('AX·정보화혁신본부 주간보고 DB');
  if (!id) p.setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(Core.SCHEMA).forEach(function (t) {
    var sh = ss.getSheetByName(t) || ss.insertSheet(t);
    var header = Object.keys(Core.SCHEMA[t]);
    // 날짜·전화번호 자동 변환을 막기 위해 전체를 일반 텍스트로 둔다.
    sh.getRange(1, 1, sh.getMaxRows(), header.length).setNumberFormat('@');
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  ['Sheet1', '시트1'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  var store = new SheetStore_(ss);
  if (store.all('people').length) {
    Logger.log('DB: %s (people 시트에 기존 조직 유지)', ss.getUrl());
    return ss.getUrl();
  }
  var org = loadOrg_();
  if (!org) {
    Logger.log('DB: %s\n조직 데이터가 없습니다. 다음 중 하나를 한 뒤 setup() 을 다시 실행하세요.\n' +
      ' - DB 의 depts/teams/people 시트에 직접 붙여넣기 (이 경우 setup 재실행 불필요)\n' +
      ' - 스크립트 속성 ORG_SHEET_ID: 조직 원본 스프레드시트 ID\n' +
      ' - 스크립트 속성 ORG_JSON: {"depts":[],"teams":[],"people":[]} (9KB 이하)\n' +
      ' - 스크립트 속성 INITIAL_ADMIN: 관리자 이메일 아이디 (나머지는 관리 화면에서 등록)', ss.getUrl());
    return ss.getUrl();
  }
  Core.importOrg(store, org.data);
  Logger.log('DB: %s (조직 %s명 등록, 출처: %s)', ss.getUrl(), org.data.people.length, org.source);
  return ss.getUrl();
}

// 실제 조직 데이터는 저장소(공개)에 두지 않는다. 운영자가 아래 중 하나로 넣는다.
function loadOrg_() {
  var p = PropertiesService.getScriptProperties();

  // 1) 별도 비공개 스프레드시트 (시트 이름 depts / teams / people, 1행은 필드명)
  var sourceId = p.getProperty('ORG_SHEET_ID');
  if (sourceId) {
    var src = SpreadsheetApp.openById(sourceId);
    var data = {};
    ['depts', 'teams', 'people'].forEach(function (t) {
      var sh = src.getSheetByName(t);
      data[t] = sh && sh.getLastRow() > 1 ? Core.fromTable(sh.getDataRange().getDisplayValues()) : [];
    });
    return { source: 'ORG_SHEET_ID', data: data };
  }

  // 2) 스크립트 속성 JSON (속성 값 한도 9KB — 작은 조직용)
  var json = p.getProperty('ORG_JSON');
  if (json) return { source: 'ORG_JSON', data: JSON.parse(json) };

  // 3) 관리자 1명만 등록하고 나머지는 관리 화면에서 추가
  var admin = String(p.getProperty('INITIAL_ADMIN') || '').trim().toLowerCase().replace(/@.*$/, '');
  if (admin) {
    return { source: 'INITIAL_ADMIN', data: { depts: [], teams: [], people: [{ id: admin, position: '관리자', admin: true, active: true }] } };
  }
  return null;
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'remindDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('remindDaily').timeBased().everyDays(1).atHour(9).inTimezone(TZ).create();
}

// ───────────────────────── 메일 ─────────────────────────

function remindDaily() {
  sendMails_(Core.reminders(new SheetStore_(db_()), new Date()));
}

function notify_(events) {
  var store = new SheetStore_(db_());
  var notify = store.get('settings', 'notify');
  if (notify && notify.value === 'false') return;
  try {
    sendMails_(events);
  } catch (err) {
    // 메일 실패가 보고 저장을 되돌리지 않도록 기록만 남긴다.
    console.error('알림 메일 실패: ' + err);
  }
}

function sendMails_(events) {
  var url = ScriptApp.getService().getUrl();
  events.forEach(function (e) {
    if (MailApp.getRemainingDailyQuota() < e.to.length) return;
    MailApp.sendEmail({
      to: e.to.map(function (id) { return id + '@' + DOMAIN; }).join(','),
      subject: e.subject,
      body: e.body + (url ? '\n\n주간보고 바로가기: ' + url : '')
    });
  });
}
