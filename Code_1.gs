/**
 * 香山高中 足球無人機社｜點名計分簿 — Google 試算表後端
 *
 * 1. 把下面的 PASSWORD 改成你自己的密碼（網頁連線時要輸入）
 * 2. 上方函式選單選 setup → 執行（第一次會要求授權）
 * 3. 部署 → 新增部署作業 → 網頁應用程式
 *    執行身分：我　／　誰可以存取：所有人
 *
 * 試算表分頁：
 *   名單         可以直接在試算表改姓名、班級、座號（網頁按「重新載入」就會更新）
 *   總表         自動產生，每次存檔都會重算，請不要手動改
 *   加減分紀錄   自動產生，請不要手動改
 *   _資料、_組別 隱藏分頁，網頁的原始紀錄，請不要動（分組請在網頁上拖曳調整）
 */
const PASSWORD = '請改成你的密碼';

const SESSIONS = ['2026-10-01', '2026-10-15', '2026-11-05', '2026-11-12', '2026-12-10', '2027-01-07'];
const TZ = 'Asia/Taipei';
const SH = { roster: '名單', data: '_資料', groups: '_組別', summary: '總表', log: '加減分紀錄' };
const ROSTER_HEAD = ['key', '班級', '座號', '學號', '姓名', '狀態', '組別代碼'];
const NC = ROSTER_HEAD.length;
const STAT = { P: '出席', T: '遲到', L: '請假', A: '曠課' };
const KEY_RE = /^[A-Za-z0-9_-]{1,60}$/;

/* ---------- 網頁入口 ---------- */
function doGet() {
  return json_({ ok: true, msg: '點名計分簿後端運作中，請從網頁使用。' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (!PASSWORD || PASSWORD === '請改成你的密碼') return json_({ ok: false, error: 'change_password' });
  if (req.token !== PASSWORD) { Utilities.sleep(800); return json_({ ok: false, error: 'wrong_password' }); }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return json_({ ok: false, error: 'busy' });
  try {
    ensureSheets_();
    switch (req.action) {
      case 'load': return json_(Object.assign({ ok: true }, loadAll_()));
      case 'saveDay': saveDay_(req.date, req.day); rebuild_(); return json_({ ok: true });
      case 'saveStudents': saveStudents_(req.students || []); rebuild_(); return json_({ ok: true });
      case 'deleteStudent': deleteStudent_(req.key); rebuild_(); return json_({ ok: true });
      case 'saveGroups': saveGroups_(req.groups || []); rebuild_(); return json_({ ok: true });
      default: return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: 'server', detail: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 第一次使用時手動執行一次：建立分頁 */
function setup() {
  ensureSheets_();
  const ss = ss_();
  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s);
  });
  rebuild_();
}

/* ---------- 工具 ---------- */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { return ss_().getSheetByName(name); }
function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 40); }
function txt_(v, max) { v = clean_(v, max); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function sortStu_(a, b) {
  return String(a.cls).localeCompare(String(b.cls)) || (Number(a.seat) || 0) - (Number(b.seat) || 0) || String(a.name).localeCompare(String(b.name));
}

function ensureSheets_() {
  const ss = ss_();
  let r = ss.getSheetByName(SH.roster);
  if (!r) {
    r = ss.insertSheet(SH.roster, 0);
    r.getRange('A:G').setNumberFormat('@');
    r.getRange(1, 1, 1, NC).setValues([ROSTER_HEAD]).setFontWeight('bold').setBackground('#dff2fa');
    r.setFrozenRows(1);
    r.hideColumns(1);
    r.hideColumns(7);
  }
  let g = ss.getSheetByName(SH.groups);
  if (!g) {
    g = ss.insertSheet(SH.groups);
    g.getRange('A:B').setNumberFormat('@');
    g.getRange(1, 1, 1, 2).setValues([['組別代碼', '組名']]);
    g.hideSheet();
  }
  let d = ss.getSheetByName(SH.data);
  if (!d) {
    d = ss.insertSheet(SH.data);
    d.getRange('A:B').setNumberFormat('@');
    d.getRange(1, 1, 1, 2).setValues([['日期', '資料(JSON)']]);
    d.hideSheet();
  }
  if (!ss.getSheetByName(SH.summary)) ss.insertSheet(SH.summary, 1);
  if (!ss.getSheetByName(SH.log)) ss.insertSheet(SH.log, 2);
}

/* ---------- 讀取 ---------- */
function readRoster_() {
  const sh = sheet_(SH.roster);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const vals = sh.getRange(2, 1, n - 1, NC).getDisplayValues();
  const used = new Set(vals.map(v => v[0]).filter(Boolean));
  const out = [];
  vals.forEach((v, i) => {
    if (!v[3] && !v[4]) return; // 空白列
    let key = v[0];
    if (!key) { // 老師直接在試算表新增的列：補一個 key
      key = 's' + (v[3].replace(/[^A-Za-z0-9]/g, '') || Utilities.getUuid().slice(0, 8));
      while (used.has(key)) key += 'x';
      used.add(key);
      sh.getRange(i + 2, 1).setNumberFormat('@').setValue(key);
    }
    out.push({ key: key, cls: v[1], seat: v[2], sid: v[3], name: v[4], active: v[5] !== '退社', group: v[6] || '', row: i + 2 });
  });
  return out;
}

function readDays_() {
  const sh = sheet_(SH.data);
  const n = sh.getLastRow();
  const o = {};
  if (n < 2) return o;
  sh.getRange(2, 1, n - 1, 2).getDisplayValues().forEach(r => {
    if (!r[0]) return;
    try { o[r[0]] = JSON.parse(r[1]); } catch (e) { /* 壞掉的列略過 */ }
  });
  return o;
}

function readGroups_() {
  const sh = sheet_(SH.groups);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 2).getDisplayValues().filter(r => r[0]).map(r => ({ id: r[0], name: r[1] }));
}

function loadAll_() {
  return {
    students: readRoster_().map(s => ({ key: s.key, cls: s.cls, seat: s.seat, sid: s.sid, name: s.name, active: s.active, group: s.group })),
    groups: readGroups_(),
    days: readDays_(),
  };
}

/* ---------- 寫入 ---------- */
function saveStudents_(list) {
  const sh = sheet_(SH.roster);
  const cur = readRoster_();
  list.forEach(s => {
    const key = clean_(s.key, 60);
    if (!KEY_RE.test(key)) return;
    const group = KEY_RE.test(s.group || '') ? s.group : '';
    const row = [key, txt_(s.cls), txt_(s.seat), txt_(s.sid), txt_(s.name), s.active === false ? '退社' : '在籍', group];
    const hit = cur.find(c => c.key === key);
    if (hit) {
      sh.getRange(hit.row, 1, 1, NC).setNumberFormat('@').setValues([row]);
    } else {
      const r = sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, NC).setNumberFormat('@').setValues([row]);
      cur.push({ key: key, row: r });
    }
  });
}

function saveGroups_(list) {
  const rows = [];
  list.slice(0, 30).forEach(g => {
    const id = clean_(g && g.id, 60);
    if (KEY_RE.test(id) && !rows.some(r => r[0] === id)) rows.push([id, txt_(g.name, 20) || '未命名']);
  });
  const sh = sheet_(SH.groups);
  const n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, 2).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
}

function deleteStudent_(key) {
  const hit = readRoster_().find(s => s.key === key);
  if (hit) sheet_(SH.roster).deleteRow(hit.row);
}

function saveDay_(date, day) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('bad date');
  day = day || {};
  const att = {}, pts = {}, gpts = {}, log = [];
  Object.keys(day.att || {}).forEach(k => {
    const v = day.att[k];
    if (KEY_RE.test(k) && /^[TLA]$/.test(v)) att[k] = v;
  });
  Object.keys(day.pts || {}).forEach(k => {
    const v = Math.round(Number(day.pts[k]));
    if (KEY_RE.test(k) && isFinite(v) && v) pts[k] = Math.max(-99, Math.min(99, v));
  });
  Object.keys(day.gpts || {}).forEach(k => {
    const v = Math.round(Number(day.gpts[k]));
    if (KEY_RE.test(k) && isFinite(v) && v) gpts[k] = Math.max(-999, Math.min(999, v));
  });
  (Array.isArray(day.log) ? day.log : []).slice(-800).forEach(e => {
    if (!e || (e.d !== 1 && e.d !== -1)) return;
    const base = { d: e.d, r: clean_(e.r, 30), t: clean_(e.t, 25) };
    if (KEY_RE.test(e.k || '')) log.push(Object.assign({ k: e.k }, base));
    else if (KEY_RE.test(e.g || '')) log.push(Object.assign({ g: e.g }, base));
  });
  const json = JSON.stringify({ date: date, att: att, pts: pts, gpts: gpts, log: log });
  const sh = sheet_(SH.data);
  const n = sh.getLastRow();
  const dates = n > 1 ? sh.getRange(2, 1, n - 1, 1).getDisplayValues().map(r => r[0]) : [];
  const i = dates.indexOf(date);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(json);
  else sh.getRange(n + 1, 1, 1, 2).setNumberFormat('@').setValues([[date, json]]);
}

/* ---------- 重算「總表」與「加減分紀錄」 ---------- */
function rebuild_() {
  const stus = readRoster_().sort(sortStu_);
  const days = readDays_();
  const groups = readGroups_();
  const gname = {};
  groups.forEach(g => { gname[g.id] = g.name; });
  const today = today_();
  const has = d => { const x = days[d]; return !!x && (Object.keys(x.att || {}).length + Object.keys(x.pts || {}).length + Object.keys(x.gpts || {}).length + (x.log || []).length > 0); };
  const rec = Object.keys(days).filter(has);
  const cols = Array.from(new Set(SESSIONS.concat(rec))).sort();
  const counted = new Set(rec.concat(SESSIONS.filter(d => d <= today)));
  const label = d => { const i = SESSIONS.indexOf(d); return (i >= 0 ? '第' + (i + 1) + '次 ' : '補課 ') + Number(d.slice(5, 7)) + '/' + Number(d.slice(8)); };

  const head = ['班級', '座號', '學號', '姓名', '組別', '狀態'].concat(cols.map(label), ['出席', '遲到', '請假', '曠課', '總分']);
  const rows = stus.map(s => {
    let P = 0, T = 0, L = 0, A = 0, pts = 0;
    const cells = cols.map(d => {
      if (!counted.has(d)) return '';
      const x = days[d] || {};
      const st = (x.att || {})[s.key] || 'P';
      const p = (x.pts || {})[s.key] || 0;
      pts += p;
      if (st === 'P') P++; else if (st === 'T') T++; else if (st === 'L') L++; else A++;
      return (st === 'P' ? '' : STAT[st]) + (p ? (p > 0 ? '+' + p : String(p)) : '');
    });
    s.pts = pts;
    return [s.cls, s.seat, s.sid, s.name, gname[s.group] || '', s.active ? '在籍' : '退社'].concat(cells, [P + T, T, L, A, pts]);
  });

  const sh = sheet_(SH.summary);
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#dff2fa');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 6 + cols.length).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  }
  let next = rows.length + 3;
  if (groups.length) {
    const ghead = ['組別', '人數', '組別分數'].concat(cols.map(label), ['成員個人分數合計']);
    const grows = groups.map(g => {
      const mem = stus.filter(s => s.group === g.id && s.active);
      let tot = 0;
      const cells = cols.map(d => {
        if (!counted.has(d)) return '';
        const p = ((days[d] || {}).gpts || {})[g.id] || 0;
        tot += p;
        return p ? (p > 0 ? '+' + p : String(p)) : '';
      });
      return [g.name, mem.length, tot].concat(cells, [mem.reduce((a, s) => a + (s.pts || 0), 0)]);
    });
    sh.getRange(next, 1, 1, ghead.length).setValues([ghead]).setFontWeight('bold').setBackground('#fff3d1');
    sh.getRange(next + 1, 4, grows.length, cols.length).setNumberFormat('@');
    sh.getRange(next + 1, 1, grows.length, ghead.length).setValues(grows);
    next += grows.length + 2;
  }
  sh.getRange(next, 1).setValue('最後更新：' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') + '（此分頁自動產生，請勿手動修改）');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);

  const name = {};
  stus.forEach(s => { name[s.key] = s.name; });
  const lr = [];
  Object.keys(days).sort().forEach(d => (days[d].log || []).forEach(e => {
    const who = e.g ? '【' + (gname[e.g] || '已刪除的組') + '】' : (name[e.k] || '（已刪除）');
    lr.push([d, String(e.t || '').slice(11, 16), who, e.d > 0 ? '+1' : '-1', e.r || '']);
  }));
  const lg = sheet_(SH.log);
  lg.clear();
  lg.getRange(1, 1, 1, 5).setValues([['日期', '時間', '姓名／組別', '加減分', '原因']]).setFontWeight('bold').setBackground('#dff2fa');
  if (lr.length) lg.getRange(2, 1, lr.length, 5).setNumberFormat('@').setValues(lr);
  lg.setFrozenRows(1);
}
