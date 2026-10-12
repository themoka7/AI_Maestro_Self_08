/**
 * 주간보고 도메인 로직 — Apps Script(서버)와 브라우저(데모)에서 같이 쓴다.
 *
 * 저장소(store)는 아래 4개 메서드만 있으면 된다.
 *   all(table) → object[]      get(table, id) → object|null
 *   save(table, obj)           remove(table, id)
 * GAS 에서는 Sheets(SheetStore_), 데모에서는 localStorage(Core.MemoryStore).
 *
 * 보고 흐름: 개인(reports) → 팀장 취합(teamReports) → 과장 확인/반려
 */
var Core = (function () {
  'use strict';

  var SCHEMA = {
    depts: { id: 's', name: 's', chiefIds: 's', order: 'n' },
    teams: { id: 's', deptId: 's', name: 's', leaderIds: 's', order: 'n' },
    people: { id: 's', name: 's', position: 's', teamId: 's', duty: 's', phone: 's', reporter: 'b', admin: 'b', viewer: 'b', active: 'b' },
    reports: { id: 's', week: 's', personId: 's', teamId: 's', done: 's', plan: 's', issue: 's', status: 's', updatedAt: 's', submittedAt: 's', returnNote: 's', returnedBy: 's' },
    teamReports: { id: 's', week: 's', teamId: 's', authorId: 's', done: 's', plan: 's', issue: 's', status: 's', updatedAt: 's', submittedAt: 's', reviewNote: 's', reviewerId: 's', reviewedAt: 's' },
    settings: { id: 's', value: 's' }
  };

  var DEFAULT_SETTINGS = { personalDueDay: '4', teamDueDay: '5', notify: 'true', remind: 'true' };
  var MAX_TEXT = 5000;
  var STATUS = { draft: '작성 중', submitted: '제출', returned: '반려', approved: '확인' };

  // ───────────── 직렬화 (Sheets 행 ↔ 객체) ─────────────
  // 시트 수식 주입 방지: = + - @ 로 시작하면 앞에 zero-width space 를 붙여 저장한다.
  var ZW = '​';

  function encode(table, obj) {
    return Object.keys(SCHEMA[table]).map(function (k) {
      var v = obj[k];
      if (SCHEMA[table][k] === 'b') return v ? 'TRUE' : 'FALSE';
      v = v == null ? '' : String(v);
      return /^[=+\-@]/.test(v) ? ZW + v : v;
    });
  }

  function decode(table, row) {
    var o = {};
    Object.keys(SCHEMA[table]).forEach(function (k, i) {
      var v = row[i] == null ? '' : String(row[i]);
      if (v.charAt(0) === ZW) v = v.slice(1);
      var t = SCHEMA[table][k];
      o[k] = t === 'b' ? v.toUpperCase() === 'TRUE' : t === 'n' ? Number(v) || 0 : v;
    });
    return o;
  }

  // ───────────── 주차 (월요일 기준, KST) ─────────────
  var KST = 9 * 3600 * 1000;

  function weekOf(date) {
    var d = new Date((date ? new Date(date).getTime() : Date.now()) + KST);
    d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7));
    return d.toISOString().slice(0, 10);
  }

  function isWeek(w) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(w || '')) return false;
    var d = new Date(w + 'T00:00:00Z');
    return !isNaN(d) && d.getUTCDay() === 1;
  }

  function kstDay(date) {
    return new Date((date ? new Date(date).getTime() : Date.now()) + KST).getUTCDay();
  }

  // ───────────── 공통 ─────────────
  function ids(s) {
    return String(s || '').split(',').map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean);
  }
  function now() { return new Date().toISOString(); }
  function fail(msg) { throw new Error(msg); }
  function byOrder(a, b) { return (a.order - b.order) || a.name.localeCompare(b.name); }
  function text(v, label, required) {
    v = String(v == null ? '' : v).replace(/\r\n/g, '\n').trim();
    if (required && !v) fail(label + '을(를) 입력하세요.');
    if (v.length > MAX_TEXT) fail(label + '은(는) ' + MAX_TEXT + '자 이내로 입력하세요.');
    return v;
  }
  function needWeek(w) { if (!isWeek(w)) fail('주차 형식이 올바르지 않습니다.'); return w; }
  function displayName(p) { return p ? (p.name || p.id) : ''; }

  function settingsOf(store) {
    var s = {};
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) { s[k] = DEFAULT_SETTINGS[k]; });
    store.all('settings').forEach(function (r) { s[r.id] = r.value; });
    return s;
  }

  // ───────────── 요청 컨텍스트 / 권한 ─────────────
  function context(store, meId) {
    meId = String(meId || '').toLowerCase();
    var people = store.all('people');
    var depts = store.all('depts').sort(byOrder);
    var deptRank = {};
    depts.forEach(function (d, i) { deptRank[d.id] = i; });
    var teams = store.all('teams').sort(function (a, b) {
      return ((deptRank[a.deptId] || 0) - (deptRank[b.deptId] || 0)) || byOrder(a, b);
    });
    var me = people.filter(function (p) { return p.id === meId; })[0];
    if (!me || !me.active) fail('등록되지 않은 사용자입니다 (' + meId + '). 관리자에게 등록을 요청하세요.');
    return {
      store: store, me: me, people: people, teams: teams, depts: depts, events: [],
      leaderOf: teams.filter(function (t) { return ids(t.leaderIds).indexOf(me.id) >= 0; }).map(function (t) { return t.id; }),
      chiefOf: depts.filter(function (d) { return ids(d.chiefIds).indexOf(me.id) >= 0; }).map(function (d) { return d.id; }),
      admin: !!me.admin,
      viewer: !!me.viewer || !!me.admin
    };
  }

  function team(c, id) { return c.teams.filter(function (t) { return t.id === id; })[0] || fail('팀을 찾을 수 없습니다.'); }
  function dept(c, id) { return c.depts.filter(function (d) { return d.id === id; })[0] || fail('과를 찾을 수 없습니다.'); }
  function person(c, id) { return c.people.filter(function (p) { return p.id === id; })[0]; }

  function leads(c, teamId) { return c.leaderOf.indexOf(teamId) >= 0; }
  function chiefs(c, deptId) { return c.chiefOf.indexOf(deptId) >= 0; }
  function canViewTeam(c, teamId) { return leads(c, teamId) || chiefs(c, team(c, teamId).deptId) || c.viewer; }
  function canViewDept(c, deptId) { return chiefs(c, deptId) || c.viewer; }

  // 이름 입력 권한: 본인, 소속 팀장, 소속 과장, 관리자
  function canName(c, p) {
    if (p.id === c.me.id || c.admin) return true;
    if (!p.teamId) return false;
    var t = c.teams.filter(function (x) { return x.id === p.teamId; })[0];
    return !!t && (leads(c, t.id) || chiefs(c, t.deptId));
  }

  function members(c, teamId) {
    return c.people.filter(function (p) { return p.active && p.reporter && p.teamId === teamId; })
      .sort(function (a, b) { return a.id.localeCompare(b.id); });
  }

  function emit(c, to, subject, body) {
    to = to.filter(function (id) { return id && id.charAt(0) !== '_'; });
    if (to.length) c.events.push({ to: to, subject: subject, body: body });
  }

  function weekText(w) {
    var thu = new Date(w + 'T00:00:00Z');
    thu.setUTCDate(thu.getUTCDate() + 3);
    return (thu.getUTCMonth() + 1) + '월 ' + Math.ceil(thu.getUTCDate() / 7) + '주차';
  }

  // ───────────── API ─────────────
  var API = {};
  var WRITES = {};

  API.bootstrap = function (c) {
    var me = Object.assign({}, c.me, { leaderOf: c.leaderOf, chiefOf: c.chiefOf, isAdmin: c.admin, isViewer: c.viewer });
    return { me: me, depts: c.depts, teams: c.teams, people: c.people, settings: settingsOf(c.store), thisWeek: weekOf() };
  };

  // 이름 기억: 한 번 입력하면 people 시트에 저장된다.
  WRITES.setName = true;
  API.setName = function (c, a) {
    var p = person(c, String(a.personId || c.me.id).toLowerCase()) || fail('사람을 찾을 수 없습니다.');
    if (!canName(c, p)) fail('이 사람의 이름을 입력할 권한이 없습니다.');
    var name = String(a.name || '').trim();
    if (!name || name.length > 20) fail('이름은 1~20자로 입력하세요.');
    p.name = name;
    c.store.save('people', p);
    return { id: p.id, name: name };
  };

  // ── 개인 ──
  function teamLocked(c, week, teamId) {
    var tr = c.store.get('teamReports', week + '_' + teamId);
    return !!tr && (tr.status === 'submitted' || tr.status === 'approved');
  }

  API.myReport = function (c, a) {
    var week = needWeek(a.week);
    var r = c.store.get('reports', week + '_' + c.me.id);
    var locked = !!c.me.teamId && teamLocked(c, week, c.me.teamId) && !(r && r.status === 'returned');
    return { report: r, locked: locked };
  };

  API.myHistory = function (c) {
    return c.store.all('reports')
      .filter(function (r) { return r.personId === c.me.id; })
      .sort(function (a, b) { return a.week < b.week ? 1 : -1; })
      .slice(0, 12)
      .map(function (r) { return { week: r.week, status: r.status, submittedAt: r.submittedAt, returnNote: r.returnNote }; });
  };

  WRITES.saveMyReport = true;
  API.saveMyReport = function (c, a) {
    var week = needWeek(a.week);
    if (!c.me.reporter) fail('주간보고 작성 대상이 아닙니다.');
    if (!c.me.teamId) fail('소속 팀이 지정되지 않았습니다. 관리자에게 문의하세요.');
    var id = week + '_' + c.me.id;
    var prev = c.store.get('reports', id);
    if (teamLocked(c, week, c.me.teamId) && !(prev && prev.status === 'returned')) {
      fail('팀 보고가 이미 과장님께 제출되어 수정할 수 없습니다. 팀장에게 문의하세요.');
    }
    var submit = !!a.submit;
    var r = {
      id: id, week: week, personId: c.me.id, teamId: c.me.teamId,
      done: text(a.done, '금주 실적', submit), plan: text(a.plan, '차주 계획', submit), issue: text(a.issue, '건의·특이사항'),
      status: submit ? 'submitted' : (prev && prev.status === 'returned' ? 'returned' : 'draft'),
      updatedAt: now(),
      submittedAt: submit ? now() : (prev ? prev.submittedAt : ''),
      returnNote: submit ? '' : (prev ? prev.returnNote : ''),
      returnedBy: submit ? '' : (prev ? prev.returnedBy : '')
    };
    c.store.save('reports', r);

    if (submit) {
      var t = team(c, c.me.teamId);
      var mem = members(c, t.id);
      var done = c.store.all('reports').filter(function (x) { return x.week === week && x.teamId === t.id && x.status === 'submitted'; }).length;
      if (done >= mem.length) {
        emit(c, ids(t.leaderIds), '[주간보고] ' + t.name + ' 팀원 보고 제출 완료 (' + weekText(week) + ')',
          t.name + ' 팀원 ' + mem.length + '명이 모두 ' + weekText(week) + ' 주간보고를 제출했습니다.\n팀 보고를 작성해 과장님께 제출해 주세요.');
      }
    }
    return r;
  };

  // ── 팀장 ──
  API.teamWeek = function (c, a) {
    var week = needWeek(a.week);
    var t = team(c, a.teamId);
    if (!canViewTeam(c, t.id)) fail('이 팀의 보고를 볼 권한이 없습니다.');
    var reports = {};
    c.store.all('reports').forEach(function (r) { if (r.week === week && r.teamId === t.id) reports[r.personId] = r; });
    return {
      team: t,
      members: members(c, t.id).map(function (p) { return { person: p, report: reports[p.id] || null, canName: canName(c, p) }; }),
      teamReport: c.store.get('teamReports', week + '_' + t.id),
      canEdit: leads(c, t.id)
    };
  };

  WRITES.returnReport = true;
  API.returnReport = function (c, a) {
    var r = c.store.get('reports', a.id) || fail('보고를 찾을 수 없습니다.');
    if (!leads(c, r.teamId)) fail('팀장만 반려할 수 있습니다.');
    if (r.status !== 'submitted') fail('제출된 보고만 반려할 수 있습니다.');
    r.status = 'returned';
    r.returnNote = text(a.note, '반려 사유', true);
    r.returnedBy = c.me.id;
    r.updatedAt = now();
    c.store.save('reports', r);
    emit(c, [r.personId], '[주간보고] ' + weekText(r.week) + ' 보고가 반려되었습니다',
      displayName(c.me) + ' 팀장님이 ' + weekText(r.week) + ' 주간보고를 반려했습니다.\n\n사유: ' + r.returnNote + '\n\n수정 후 다시 제출해 주세요.');
    return r;
  };

  WRITES.saveTeamReport = true;
  API.saveTeamReport = function (c, a) {
    var week = needWeek(a.week);
    var t = team(c, a.teamId);
    if (!leads(c, t.id)) fail('팀장만 팀 보고를 작성할 수 있습니다.');
    var id = week + '_' + t.id;
    var prev = c.store.get('teamReports', id);
    if (prev && prev.status === 'approved') fail('과장님이 이미 확인한 보고입니다.');
    var submit = !!a.submit;
    var r = {
      id: id, week: week, teamId: t.id, authorId: c.me.id,
      done: text(a.done, '금주 실적', submit), plan: text(a.plan, '차주 계획', submit), issue: text(a.issue, '건의·특이사항'),
      status: submit ? 'submitted' : (prev && prev.status === 'returned' ? 'returned' : 'draft'),
      updatedAt: now(),
      submittedAt: submit ? now() : (prev ? prev.submittedAt : ''),
      reviewNote: submit ? '' : (prev ? prev.reviewNote : ''),
      reviewerId: prev ? prev.reviewerId : '',
      reviewedAt: prev ? prev.reviewedAt : ''
    };
    c.store.save('teamReports', r);
    if (submit) {
      var d = dept(c, t.deptId);
      emit(c, ids(d.chiefIds), '[주간보고] ' + t.name + ' 팀 보고 제출 (' + weekText(week) + ')',
        t.name + ' ' + weekText(week) + ' 팀 주간보고가 제출되었습니다.');
    }
    return r;
  };

  // ── 과장 ──
  function teamRow(c, t, week, reportsByTeam) {
    var mem = members(c, t.id);
    var subs = (reportsByTeam[t.id] || []).filter(function (r) { return r.status === 'submitted'; });
    return {
      team: t,
      leaders: ids(t.leaderIds).map(function (id) { return displayName(person(c, id)) || id; }),
      total: mem.length,
      submitted: subs.length,
      teamReport: c.store.get('teamReports', week + '_' + t.id)
    };
  }

  function groupReports(c, week) {
    var g = {};
    c.store.all('reports').forEach(function (r) {
      if (r.week !== week) return;
      (g[r.teamId] = g[r.teamId] || []).push(r);
    });
    return g;
  }

  API.deptWeek = function (c, a) {
    var week = needWeek(a.week);
    var d = dept(c, a.deptId);
    if (!canViewDept(c, d.id)) fail('이 과의 보고를 볼 권한이 없습니다.');
    var g = groupReports(c, week);
    return {
      dept: d,
      teams: c.teams.filter(function (t) { return t.deptId === d.id; }).map(function (t) { return teamRow(c, t, week, g); }),
      canReview: chiefs(c, d.id)
    };
  };

  WRITES.reviewTeamReport = true;
  API.reviewTeamReport = function (c, a) {
    var r = c.store.get('teamReports', a.id) || fail('팀 보고를 찾을 수 없습니다.');
    var t = team(c, r.teamId);
    if (!chiefs(c, t.deptId)) fail('과장만 확인/반려할 수 있습니다.');
    if (r.status !== 'submitted' && r.status !== 'approved') fail('제출된 보고만 처리할 수 있습니다.');
    var approve = a.action === 'approve';
    r.status = approve ? 'approved' : 'returned';
    r.reviewNote = text(a.note, '의견', !approve);
    r.reviewerId = c.me.id;
    r.reviewedAt = now();
    c.store.save('teamReports', r);
    emit(c, ids(t.leaderIds), '[주간보고] ' + t.name + ' 팀 보고 ' + (approve ? '확인' : '반려') + ' (' + weekText(r.week) + ')',
      displayName(c.me) + ' 과장님이 ' + t.name + ' ' + weekText(r.week) + ' 팀 보고를 ' + (approve ? '확인' : '반려') + '했습니다.' +
      (r.reviewNote ? '\n\n의견: ' + r.reviewNote : ''));
    return r;
  };

  // 취합본: deptId 가 '*' 이면 볼 수 있는 모든 과
  API.compiled = function (c, a) {
    var week = needWeek(a.week);
    var targets = a.deptId === '*' ? c.depts.filter(function (d) { return canViewDept(c, d.id); }) : [dept(c, a.deptId)];
    if (!targets.length || !targets.every(function (d) { return canViewDept(c, d.id); })) fail('취합본을 볼 권한이 없습니다.');
    var g = groupReports(c, week);
    return targets.map(function (d) {
      return {
        dept: d,
        teams: c.teams.filter(function (t) { return t.deptId === d.id; }).map(function (t) {
          var row = teamRow(c, t, week, g);
          if (a.withMembers) {
            row.memberReports = (g[t.id] || []).filter(function (r) { return r.status === 'submitted'; })
              .map(function (r) { return Object.assign({ name: displayName(person(c, r.personId)) || r.personId }, r); });
          }
          return row;
        })
      };
    });
  };

  // ── 관리 ──
  function needAdmin(c) { if (!c.admin) fail('관리자만 사용할 수 있습니다.'); }
  function cleanId(v, label) {
    v = String(v || '').trim().toLowerCase();
    if (!/^[a-z0-9_][a-z0-9._-]{0,63}$/.test(v)) fail(label + ' 형식이 올바르지 않습니다 (영문 소문자·숫자·._-).');
    return v;
  }
  function cleanIds(v) {
    return ids(v).map(function (x) { return cleanId(x, '아이디'); }).join(',');
  }

  WRITES.savePerson = true;
  API.savePerson = function (c, a) {
    needAdmin(c);
    var p = a.person || {};
    var id = cleanId(p.id, '이메일 아이디');
    var prev = person(c, id);
    if (a.isNew && prev) fail('이미 등록된 아이디입니다: ' + id);
    if (p.teamId) team(c, p.teamId);
    var next = {
      id: id, name: String(p.name || '').trim().slice(0, 20), position: String(p.position || '').trim().slice(0, 30),
      teamId: p.teamId || '', duty: text(p.duty, '담당업무'), phone: String(p.phone || '').trim().slice(0, 40),
      reporter: !!p.reporter, admin: !!p.admin, viewer: !!p.viewer, active: p.active !== false
    };
    var admins = c.people.filter(function (x) { return x.id !== id && x.admin && x.active; }).length;
    if (!(next.admin && next.active) && admins === 0) fail('관리자가 최소 1명은 있어야 합니다.');
    c.store.save('people', next);
    return next;
  };

  WRITES.removePerson = true;
  API.removePerson = function (c, a) {
    needAdmin(c);
    if (a.id === c.me.id) fail('자기 자신은 삭제할 수 없습니다.');
    c.store.remove('people', a.id);
    return true;
  };

  WRITES.saveTeam = true;
  API.saveTeam = function (c, a) {
    needAdmin(c);
    var t = a.team || {};
    var id = cleanId(t.id, '팀 코드');
    if (a.isNew && c.teams.some(function (x) { return x.id === id; })) fail('이미 있는 팀 코드입니다.');
    dept(c, t.deptId);
    var next = { id: id, deptId: t.deptId, name: text(t.name, '팀 이름', true).slice(0, 40), leaderIds: cleanIds(t.leaderIds), order: Number(t.order) || 0 };
    c.store.save('teams', next);
    return next;
  };

  WRITES.removeTeam = true;
  API.removeTeam = function (c, a) {
    needAdmin(c);
    if (c.people.some(function (p) { return p.teamId === a.id; })) fail('소속 구성원이 있는 팀은 삭제할 수 없습니다.');
    c.store.remove('teams', a.id);
    return true;
  };

  WRITES.saveDept = true;
  API.saveDept = function (c, a) {
    needAdmin(c);
    var d = a.dept || {};
    var id = cleanId(d.id, '과 코드');
    if (a.isNew && c.depts.some(function (x) { return x.id === id; })) fail('이미 있는 과 코드입니다.');
    var next = { id: id, name: text(d.name, '과 이름', true).slice(0, 40), chiefIds: cleanIds(d.chiefIds), order: Number(d.order) || 0 };
    c.store.save('depts', next);
    return next;
  };

  WRITES.removeDept = true;
  API.removeDept = function (c, a) {
    needAdmin(c);
    if (c.teams.some(function (t) { return t.deptId === a.id; })) fail('소속 팀이 있는 과는 삭제할 수 없습니다.');
    c.store.remove('depts', a.id);
    return true;
  };

  WRITES.saveSettings = true;
  API.saveSettings = function (c, a) {
    needAdmin(c);
    var s = a.settings || {};
    ['personalDueDay', 'teamDueDay'].forEach(function (k) {
      var v = Number(s[k]);
      if (!(v >= 1 && v <= 5)) fail('마감 요일은 월~금 중에서 선택하세요.');
      c.store.save('settings', { id: k, value: String(v) });
    });
    ['notify', 'remind'].forEach(function (k) { c.store.save('settings', { id: k, value: s[k] ? 'true' : 'false' }); });
    return settingsOf(c.store);
  };

  function call(store, meId, method, args) {
    if (!Object.prototype.hasOwnProperty.call(API, method)) fail('알 수 없는 요청: ' + method);
    var c = context(store, meId);
    var data = API[method](c, args || {});
    return { data: data, events: c.events };
  }

  // ───────────── 리마인더 (GAS 트리거용) ─────────────
  function reminders(store, date) {
    var s = settingsOf(store);
    if (s.remind !== 'true') return [];
    var day = kstDay(date), week = weekOf(date), out = [];
    var people = store.all('people'), teams = store.all('teams');
    var reports = store.all('reports').filter(function (r) { return r.week === week; });

    if (day === Number(s.personalDueDay)) {
      people.filter(function (p) { return p.active && p.reporter && p.teamId; }).forEach(function (p) {
        var r = reports.filter(function (x) { return x.personId === p.id; })[0];
        if (!r || r.status !== 'submitted') {
          out.push({ to: [p.id], subject: '[주간보고] 오늘은 ' + weekText(week) + ' 개인 보고 마감일입니다',
            body: (p.name || p.id) + '님, ' + weekText(week) + ' 주간보고를 아직 제출하지 않았습니다.' });
        }
      });
    }
    if (day === Number(s.teamDueDay)) {
      teams.forEach(function (t) {
        var tr = store.get('teamReports', week + '_' + t.id);
        if (!tr || (tr.status !== 'submitted' && tr.status !== 'approved')) {
          out.push({ to: ids(t.leaderIds), subject: '[주간보고] 오늘은 ' + t.name + ' 팀 보고 마감일입니다',
            body: t.name + ' ' + weekText(week) + ' 팀 보고를 아직 제출하지 않았습니다.' });
        }
      });
    }
    return out.map(function (e) { e.to = e.to.filter(function (id) { return id.charAt(0) !== '_'; }); return e; })
      .filter(function (e) { return e.to.length; });
  }

  // ───────────── 조직 데이터 ─────────────
  // 실제 조직 데이터는 저장소에 넣지 않는다 (공개 저장소).
  // 운영: Code.gs 의 setup() 이 스프레드시트·스크립트 속성에서 읽어 importOrg() 로 넣는다.
  // 데모: 아래 가상 조직(팀A~팀E, user01~)을 쓴다.
  var DEMO_ORG = (function () {
    var depts = [
      { id: 'dept-a', name: 'A과', chiefIds: 'user02', order: 1 },
      { id: 'dept-b', name: 'B과', chiefIds: 'user02', order: 2 }
    ];
    // 한 사람이 여러 팀의 팀장(user03, user04)이고, 한 과장이 두 과(user02)를 맡는 경우를 포함한다.
    var teams = [
      { id: 'team-a', deptId: 'dept-a', name: '팀A', leaderIds: 'user03', order: 1 },
      { id: 'team-b', deptId: 'dept-a', name: '팀B', leaderIds: 'user03', order: 2 },
      { id: 'team-c', deptId: 'dept-a', name: '팀C', leaderIds: 'user04', order: 3 },
      { id: 'team-d', deptId: 'dept-b', name: '팀D', leaderIds: 'user04', order: 1 },
      { id: 'team-e', deptId: 'dept-b', name: '팀E', leaderIds: 'user05', order: 2 }
    ];
    // [아이디, 직위, 팀, 보고대상, 열람권한, 관리자]
    var rows = [
      ['user01', '본부장', '', 0, 1, 0],
      ['user02', '과장', '', 0, 0, 0],
      ['user03', '팀장', 'team-a', 0, 0, 0],
      ['user04', '팀장', 'team-c', 0, 0, 0],
      ['user05', '팀장', 'team-e', 0, 0, 0],
      ['user06', '주무관', 'team-a', 1, 0, 1],
      ['user07', '주무관', 'team-a', 1, 0, 0],
      ['user08', '주무관', 'team-a', 1, 0, 0],
      ['user09', '주무관', 'team-b', 1, 0, 0],
      ['user10', '주무관', 'team-b', 1, 0, 0],
      ['user11', '주무관', 'team-c', 1, 0, 0],
      ['user12', '주무관', 'team-c', 1, 0, 0],
      ['user13', '주무관', 'team-c', 1, 0, 0],
      ['user14', '주무관', 'team-d', 1, 0, 0],
      ['user15', '주무관', 'team-d', 1, 0, 0],
      ['user16', '주무관', 'team-e', 1, 0, 0],
      ['user17', '주무관', 'team-e', 1, 0, 0],
      ['user18', '주무관', 'team-e', 1, 0, 0]
    ];
    var people = rows.map(function (r, i) {
      var team = teams.filter(function (t) { return t.id === r[2]; })[0];
      return {
        id: r[0], name: '', position: r[1], teamId: r[2],
        duty: team ? team.name + ' 업무 ' + (i + 1) : '총괄',
        phone: '000-' + String(1000 + i), reporter: !!r[3], viewer: !!r[4], admin: !!r[5], active: true
      };
    });
    return { depts: depts, teams: teams, people: people };
  })();

  // org = { depts: [], teams: [], people: [] } (각 객체는 SCHEMA 필드). 비어 있는 저장소에만 넣는다.
  function importOrg(store, org) {
    if (store.all('people').length) return false;
    if (!org || !org.people || !org.people.length) fail('조직 데이터에 people 이 없습니다.');
    if (!org.people.some(function (p) { return p.admin; })) fail('조직 데이터에 관리자(admin=TRUE)가 최소 1명 있어야 합니다.');
    ['depts', 'teams', 'people'].forEach(function (t) {
      (org[t] || []).forEach(function (o) {
        var row = {};
        Object.keys(SCHEMA[t]).forEach(function (k) {
          var type = SCHEMA[t][k], v = o[k];
          if (type === 'b') row[k] = v === true || String(v).toUpperCase() === 'TRUE' || (k === 'active' && v == null);
          else if (type === 'n') row[k] = Number(v) || 0;
          else row[k] = v == null ? '' : String(v).trim();
        });
        row.id = row.id.toLowerCase();
        if (!row.id) fail(t + ' 에 id 가 빈 행이 있습니다.');
        store.save(t, row);
      });
    });
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      if (!store.get('settings', k)) store.save('settings', { id: k, value: DEFAULT_SETTINGS[k] });
    });
    return true;
  }

  // 헤더 행 기준으로 표(2차원 배열)를 객체 배열로 바꾼다. 열 순서가 달라도 된다.
  function fromTable(values) {
    var header = (values[0] || []).map(function (h) { return String(h).trim(); });
    var idCol = header.indexOf('id');
    if (idCol < 0) fail('1행에 id 열이 없습니다.');
    return values.slice(1).filter(function (r) { return String(r[idCol] || '').trim(); }).map(function (r) {
      var o = {};
      header.forEach(function (h, i) { if (h) o[h] = r[i]; });
      return o;
    });
  }

  function seedDemo(store) { return importOrg(store, DEMO_ORG); }

  // ───────────── 데모용 메모리 저장소 ─────────────
  function MemoryStore(data, persist) {
    data = data || {};
    Object.keys(SCHEMA).forEach(function (t) { data[t] = data[t] || []; });
    function clone(o) { return JSON.parse(JSON.stringify(o)); }
    return {
      data: data,
      all: function (t) { return clone(data[t]); },
      get: function (t, id) { var o = data[t].filter(function (x) { return x.id === id; })[0]; return o ? clone(o) : null; },
      save: function (t, o) {
        var i = data[t].map(function (x) { return x.id; }).indexOf(o.id);
        if (i >= 0) data[t][i] = clone(o); else data[t].push(clone(o));
        if (persist) persist(data);
      },
      remove: function (t, id) {
        data[t] = data[t].filter(function (x) { return x.id !== id; });
        if (persist) persist(data);
      }
    };
  }

  return {
    SCHEMA: SCHEMA, STATUS: STATUS, WRITES: WRITES,
    encode: encode, decode: decode,
    weekOf: weekOf, weekText: weekText, isWeek: isWeek,
    call: call, isWrite: function (m) { return !!WRITES[m]; },
    reminders: reminders, importOrg: importOrg, fromTable: fromTable, seedDemo: seedDemo,
    DEMO_ORG: DEMO_ORG, MemoryStore: MemoryStore
  };
})();

if (typeof module !== 'undefined') module.exports = Core;
