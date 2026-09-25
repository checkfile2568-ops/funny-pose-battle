/**
 * ============================================================
 *  ฮาท่าเทพ (Funny Pose Battle) — Backend บน Google Apps Script
 *  ------------------------------------------------------------
 *  หน้าที่:
 *   1) เก็บคะแนนทุกเกม (ชีต Scores) + สรุปคะแนนดีที่สุดรายวัน (ชีต Daily)
 *   2) คำนวณตารางอันดับ ซีซั่น / สัปดาห์นี้ / วันนี้
 *   3) ห้องแข่ง (Room) แบบเครื่องใครเครื่องมัน — เก็บสถานะใน CacheService
 *
 *  ติดตั้ง:
 *   - สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วางไฟล์นี้
 *   - รันฟังก์ชัน setup() หนึ่งครั้ง (อนุญาตสิทธิ์)
 *   - ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป
 *       ดำเนินการในฐานะ: ฉัน  |  ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 *   - นำ URL (/exec) ไปใส่ใน index.html ที่ค่า API_URL
 *
 *  จุดต่อยอด (ค้นหาคำว่า "ต่อยอด:")
 * ============================================================
 */

const APP_VERSION = '1.0.0';
const TZ = 'Asia/Bangkok';

const SHEET_CONFIG = 'Config';
const SHEET_SCORES = 'Scores';
const SHEET_DAILY = 'Daily';
const SHEET_STANDINGS = 'Standings';

const SCORES_HEADERS = ['เวลา', 'วันที่', 'playerKey', 'ชื่อ', 'แผนก', 'โหมด', 'รหัสห้อง',
  'คะแนน', 'ผ่าน(ด่าน)', 'จำนวนด่าน', 'ท่าเด่น', 'ฉายา', 'นับคะแนนซีซั่น', 'เวอร์ชัน'];
const DAILY_HEADERS = ['วันที่', 'playerKey', 'ชื่อ', 'แผนก', 'คะแนนดีที่สุด', 'จำนวนเกม', 'อัปเดตล่าสุด'];

const MAX_ROUND_SCORE = 400;           // เพดานคะแนนต่อด่าน (กันค่าผิดปกติ)
const MAX_ROUNDS = 10;
const MODES_COUNTED = ['daily', 'room', 'duo'];   // ต่อยอด: เพิ่มโหมดที่นับคะแนนซีซั่น
const ROOM_TTL_SEC = 21600;            // CacheService เก็บได้สูงสุด 6 ชม.
const ROOM_MAX_PLAYERS = 8;
const ROOM_IDLE_KICK_MS = 45000;       // ไม่ติดต่อมาเกินนี้ในล็อบบี้ = ออกจากห้อง
const BOARD_CACHE_SEC = 20;

const DEFAULT_CONFIG = [
  ['seasonName', 'ซีซั่น 1 : ฮาท่าเทพ', 'ชื่อซีซั่น (แสดงบนหัวเกม)'],
  ['seasonStart', null, 'วันเริ่มซีซั่น (ใส่เป็นวันที่ เช่น 25/9/2569)'],
  ['seasonEnd', null, 'วันจบซีซั่น (รวมวันนี้ด้วย)'],
  ['finalDayMultiplier', 2, 'ตัวคูณคะแนนวันสุดท้าย'],
  ['hideLastDays', 0, 'ซ่อนตารางอันดับกี่วันสุดท้าย (0 = ไม่ซ่อน)'],
  ['orgName', 'ออฟฟิศฮาไม่หยุด', 'ชื่อหน่วยงานบนบัตรพนักงานเกียรติยศ'],
  ['difficulty', 'normal', 'ความยากการตรวจท่า: easy / normal / hard'],
  ['roundsPerGame', 5, 'จำนวนท่าต่อเกม (ด่านประจำวัน/ห้องแข่ง)'],
  ['prizeNote', 'ประกาศรางวัลประจำสัปดาห์ทุกวันศุกร์', 'ข้อความรางวัล (แสดงในหน้าตารางอันดับ)']
];

/* ============================================================
 *  ติดตั้ง / เมนู
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎮 ฮาท่าเทพ')
    .addItem('ตั้งค่าเริ่มต้น (setup)', 'setup')
    .addItem('สรุปอันดับลงชีต Standings', 'exportStandings')
    .addSeparator()
    .addItem('เริ่มซีซั่นใหม่ (เก็บข้อมูลเก่า)', 'archiveSeason')
    .addToUi();
}

/** รันครั้งแรกครั้งเดียว: สร้างชีตและค่าตั้งต้น */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const cfg = ensureSheet_(ss, SHEET_CONFIG, ['key', 'value', 'คำอธิบาย']);
  if (cfg.getLastRow() < 2) {
    const today = new Date();
    const end = new Date(today.getTime() + 13 * 86400000);
    const rows = DEFAULT_CONFIG.map(r => {
      if (r[0] === 'seasonStart') return [r[0], today, r[2]];
      if (r[0] === 'seasonEnd') return [r[0], end, r[2]];
      return r.slice();
    });
    cfg.getRange(2, 1, rows.length, 3).setValues(rows);
    cfg.getRange(2, 2, rows.length, 1).setHorizontalAlignment('left');
    cfg.getRange('B3:B4').setNumberFormat('dd/mm/yyyy');
    cfg.setColumnWidth(1, 170); cfg.setColumnWidth(2, 260); cfg.setColumnWidth(3, 340);
  } else {
    // เติม key ที่ยังไม่มี (กรณีอัปเกรดเวอร์ชัน)
    const have = cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
    DEFAULT_CONFIG.forEach(r => { if (have.indexOf(r[0]) < 0 && r[1] !== null) cfg.appendRow(r); });
  }
  ensureSheet_(ss, SHEET_SCORES, SCORES_HEADERS);
  ensureSheet_(ss, SHEET_DAILY, DAILY_HEADERS);
  bumpBoardVersion_();
  CacheService.getScriptCache().remove('cfg');
  return 'ok';
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    // คอลัมน์วันที่เก็บเป็นข้อความ yyyy-MM-dd (กันชีตแปลงเป็นวันที่ตามโซนเวลาของไฟล์)
    const dateCol = name === SHEET_SCORES ? 'B:B' : (name === SHEET_DAILY ? 'A:A' : '');
    if (dateCol) sh.getRange(dateCol).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#0b1f4d').setFontColor('#f5c542');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ============================================================
 *  HTTP entry
 * ============================================================ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action) return json_(route_(p));
  return json_({ ok: true, app: 'funny-pose-battle', version: APP_VERSION, serverNow: Date.now() });
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  return json_(route_(body));
}

function route_(b) {
  try {
    const a = String(b.action || '');
    let out;
    switch (a) {
      case 'config': out = apiConfig_(); break;
      case 'submit': out = apiSubmit_(b); break;
      case 'leaderboard': out = apiLeaderboard_(b); break;
      case 'roomCreate': out = apiRoomCreate_(b); break;
      case 'roomJoin': out = apiRoomJoin_(b); break;
      case 'roomSync': out = apiRoomSync_(b); break;
      case 'roomStart': out = apiRoomStart_(b); break;
      case 'roomReset': out = apiRoomReset_(b); break;
      // ต่อยอด: เพิ่ม action ใหม่ตรงนี้
      default: out = { ok: false, error: 'unknown_action' };
    }
    out.serverNow = Date.now();
    return out;
  } catch (err) {
    return { ok: false, error: String(err && err.message || err), serverNow: Date.now() };
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 *  Config
 * ============================================================ */

function readConfig_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('cfg');
  if (hit) return JSON.parse(hit);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('ยังไม่ได้รัน setup()');
  const raw = {};
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => { if (r[0] !== '') raw[String(r[0]).trim()] = r[1]; });
  const today = todayKey_();
  const cfg = {
    seasonName: String(raw.seasonName || 'ซีซั่น 1'),
    seasonStart: toDateKey_(raw.seasonStart) || today,
    seasonEnd: toDateKey_(raw.seasonEnd) || addDays_(today, 13),
    finalDayMultiplier: num_(raw.finalDayMultiplier, 2, 1, 10),
    hideLastDays: Math.round(num_(raw.hideLastDays, 0, 0, 30)),
    orgName: String(raw.orgName || 'ออฟฟิศฮาไม่หยุด'),
    difficulty: ['easy', 'normal', 'hard'].indexOf(String(raw.difficulty)) >= 0 ? String(raw.difficulty) : 'normal',
    roundsPerGame: Math.round(num_(raw.roundsPerGame, 5, 3, MAX_ROUNDS)),
    prizeNote: String(raw.prizeNote || '')
  };
  if (cfg.seasonEnd < cfg.seasonStart) cfg.seasonEnd = cfg.seasonStart;
  cache.put('cfg', JSON.stringify(cfg), 60);
  return cfg;
}

function apiConfig_() {
  const cfg = readConfig_();
  const today = todayKey_();
  return {
    ok: true, version: APP_VERSION, today: today,
    status: seasonStatus_(cfg, today),
    config: cfg
  };
}

function seasonStatus_(cfg, today) {
  if (today < cfg.seasonStart) return 'upcoming';
  if (today > cfg.seasonEnd) return 'ended';
  return 'live';
}

/* ============================================================
 *  ส่งคะแนน
 * ============================================================ */

function apiSubmit_(b) {
  const cfg = readConfig_();
  const mode = String(b.mode || 'daily');
  const rounds = Math.round(num_(b.rounds, cfg.roundsPerGame, 1, cfg.roundsPerGame)); // ต่อยอด: ถ้าเพิ่มโหมดที่มีด่านมากกว่านี้ ให้ปรับเพดาน
  const score = Math.round(num_(b.score, 0, 0, rounds * MAX_ROUND_SCORE));
  const cleared = Math.round(num_(b.cleared, 0, 0, rounds));
  const players = (mode === 'duo' && Array.isArray(b.players)) ? b.players.slice(0, 2) : [{ name: b.name, dept: b.dept }];
  const clean = players.map(p => ({ name: cleanText_(p && p.name, 40), dept: cleanText_(p && p.dept, 40) }))
    .filter(p => p.name);
  if (!clean.length) return { ok: false, error: 'ต้องระบุชื่อผู้เล่น' };

  const today = todayKey_();
  const counted = MODES_COUNTED.indexOf(mode) >= 0 && seasonStatus_(cfg, today) === 'live';
  const now = new Date();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'ระบบไม่ว่าง ลองใหม่อีกครั้ง' };
  try {
    const ss = SpreadsheetApp.getActive();
    const sc = ss.getSheetByName(SHEET_SCORES);
    const rows = clean.map(p => [now, today, playerKey_(p.name), p.name, p.dept, mode,
      cleanText_(b.roomCode, 8), score, cleared, rounds, cleanText_(b.bestPose, 60),
      cleanText_(b.title, 120), counted ? 'Y' : 'N', cleanText_(b.clientVersion, 20)]);
    const at = sc.getLastRow() + 1;
    sc.getRange(at, 2, rows.length, 1).setNumberFormat('@');
    sc.getRange(at, 1, rows.length, rows[0].length).setValues(rows);

    if (counted) clean.forEach(p => upsertDaily_(ss, today, p, score, now));
    bumpBoardVersion_();
  } finally {
    lock.releaseLock();
  }

  const me = clean[0];
  const board = computeBoard_('season', cfg);
  const mine = findMe_(board, me.name);
  const todayBoard = computeBoard_('today', cfg);
  const mineToday = findMe_(todayBoard, me.name);
  return {
    ok: true, counted: counted, today: today,
    seasonTotal: mine.total, seasonRank: board.hidden ? null : mine.rank,
    todayBest: mineToday.best, todayRank: todayBoard.hidden ? null : mineToday.rank,
    hidden: board.hidden
  };
}

function upsertDaily_(ss, dateKey, p, score, now) {
  const sh = ss.getSheetByName(SHEET_DAILY);
  const key = playerKey_(p.name);
  const last = sh.getLastRow();
  if (last >= 2) {
    const ab = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = ab.length - 1; i >= 0; i--) {
      if (toDateKey_(ab[i][0]) === dateKey && String(ab[i][1]) === key) {
        const r = i + 2;
        const cur = sh.getRange(r, 1, 1, 7).getValues()[0];
        const best = Math.max(Number(cur[4]) || 0, score);
        sh.getRange(r, 3, 1, 5).setValues([[p.name, p.dept || cur[3], best, (Number(cur[5]) || 0) + 1, now]]);
        return;
      }
    }
  }
  sh.getRange(last + 1, 1).setNumberFormat('@');
  sh.getRange(last + 1, 1, 1, 7).setValues([[dateKey, key, p.name, p.dept, score, 1, now]]);
}

/* ============================================================
 *  ตารางอันดับ
 * ============================================================ */

function apiLeaderboard_(b) {
  const cfg = readConfig_();
  const scope = ['season', 'week', 'today'].indexOf(String(b.scope)) >= 0 ? String(b.scope) : 'season';
  const board = computeBoard_(scope, cfg);
  const out = {
    ok: true, scope: scope, today: board.today, from: board.from, to: board.to,
    hidden: board.hidden, cutoff: board.cutoff, status: seasonStatus_(cfg, board.today),
    config: cfg, rows: board.rows.slice(0, 100)
  };
  if (b.name) {
    const m = findMe_(board, b.name);
    // ช่วงซ่อนอันดับ: แสดงเฉพาะคะแนนของตัวเอง (คำนวณเต็มช่วง) แต่ไม่บอกอันดับ
    if (board.hidden) {
      const full = computeBoard_(scope, cfg, true);
      const mf = findMe_(full, b.name);
      out.me = { total: mf.total, best: mf.best, rank: null };
    } else {
      out.me = { total: m.total, best: m.best, rank: m.rank };
    }
  }
  return out;
}

function findMe_(board, name) {
  const key = playerKey_(cleanText_(name, 40));
  for (let i = 0; i < board.rows.length; i++) {
    if (board.rows[i].key === key) return { total: board.rows[i].total, best: board.rows[i].best, rank: i + 1 };
  }
  return { total: 0, best: 0, rank: null };
}

/**
 * คำนวณตารางอันดับ
 * - คะแนนซีซั่น = ผลรวม "คะแนนดีที่สุดของแต่ละวัน" (วันสุดท้าย x finalDayMultiplier)
 * - ignoreHidden = true ใช้คำนวณคะแนนของตัวเองในช่วงซ่อนอันดับ
 */
function computeBoard_(scope, cfg, ignoreHidden) {
  const today = todayKey_();
  const cache = CacheService.getScriptCache();
  const ck = 'board_' + scope + '_' + (ignoreHidden ? 'f' : 'h') + '_' + getBoardVersion_() + '_' + today;
  const hit = cache.get(ck);
  if (hit) return JSON.parse(hit);

  const start = cfg.seasonStart, end = cfg.seasonEnd;
  let from = start, to = end, hidden = false, cutoff = null;
  if (seasonStatus_(cfg, today) === 'live' && cfg.hideLastDays > 0 && !ignoreHidden) {
    cutoff = addDays_(end, -cfg.hideLastDays);
    if (today > cutoff) hidden = true;
  }
  if (scope === 'today') {
    from = today; to = today;
  } else if (scope === 'week') {
    const ws = weekStart_(today);
    from = ws > start ? ws : start;
    const we = addDays_(ws, 6);
    to = we < end ? we : end;
  }
  let rows = [];
  if (!(hidden && scope === 'today')) {
    if (hidden) to = to < cutoff ? to : cutoff;
    const map = {};
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DAILY);
    const last = sh ? sh.getLastRow() : 0;
    if (last >= 2) {
      const vals = sh.getRange(2, 1, last - 1, 6).getValues();
      const disp = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (let i = 0; i < vals.length; i++) {
        const d = toDateKey_(vals[i][0]) || disp[i][0];
        if (!d || d < from || d > to || d < start || d > end) continue;
        const key = String(vals[i][1]);
        const best = Number(vals[i][4]) || 0;
        const mult = d === end ? cfg.finalDayMultiplier : 1;
        const p = map[key] || (map[key] = { key: key, name: '', dept: '', total: 0, days: 0, best: 0, games: 0, lastDate: '' });
        p.total += Math.round(best * mult);
        p.days += 1;
        p.games += Number(vals[i][5]) || 0;
        if (best > p.best) p.best = best;
        if (d >= p.lastDate) { p.lastDate = d; p.name = String(vals[i][2]); p.dept = String(vals[i][3] || ''); }
      }
    }
    rows = Object.keys(map).map(k => map[k]);
    rows.sort((a, b) => (b.total - a.total) || (b.best - a.best) || (a.games - b.games));
  }
  const board = { scope: scope, today: today, from: from, to: to, hidden: hidden, cutoff: cutoff, rows: rows };
  try { cache.put(ck, JSON.stringify(board), BOARD_CACHE_SEC); } catch (e) { /* ใหญ่เกิน cache ก็ไม่เป็นไร */ }
  return board;
}

function getBoardVersion_() {
  return CacheService.getScriptCache().get('boardVer') || '0';
}
function bumpBoardVersion_() {
  CacheService.getScriptCache().put('boardVer', Date.now() + '_' + Math.floor(Math.random() * 1e6), ROOM_TTL_SEC);
}

/** เมนู: สรุปอันดับซีซั่น + สัปดาห์ ลงชีต Standings (ไว้พิมพ์/ประกาศ) */
function exportStandings() {
  const cfg = readConfig_();
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_STANDINGS);
  if (!sh) sh = ss.insertSheet(SHEET_STANDINGS);
  sh.clear();
  const season = computeBoard_('season', cfg, true);
  const week = computeBoard_('week', cfg, true);
  const out = [];
  out.push([cfg.seasonName + ' — สรุป ณ ' + thaiDate_(todayKey_()), '', '', '', '', '']);
  out.push(['อันดับซีซั่น (' + thaiDate_(cfg.seasonStart) + ' – ' + thaiDate_(cfg.seasonEnd) + ')', '', '', '', '', '']);
  out.push(['อันดับ', 'ชื่อ', 'แผนก', 'คะแนนรวม', 'วันที่เล่น', 'เกมดีที่สุด']);
  season.rows.forEach((r, i) => out.push([i + 1, r.name, r.dept, r.total, r.days, r.best]));
  out.push(['', '', '', '', '', '']);
  out.push(['อันดับสัปดาห์นี้ (' + thaiDate_(week.from) + ' – ' + thaiDate_(week.to) + ')', '', '', '', '', '']);
  out.push(['อันดับ', 'ชื่อ', 'แผนก', 'คะแนนรวม', 'วันที่เล่น', 'เกมดีที่สุด']);
  week.rows.forEach((r, i) => out.push([i + 1, r.name, r.dept, r.total, r.days, r.best]));
  sh.getRange(1, 1, out.length, 6).setValues(out);
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sh.autoResizeColumns(1, 6);
  return 'ok';
}

/** เมนู: ปิดซีซั่นเดิม — เปลี่ยนชื่อชีต Scores/Daily เก็บไว้ แล้วสร้างชุดใหม่ */
function archiveSeason() {
  const ss = SpreadsheetApp.getActive();
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm');
  [SHEET_SCORES, SHEET_DAILY].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh) sh.setName(n + '_' + stamp);
  });
  ensureSheet_(ss, SHEET_SCORES, SCORES_HEADERS);
  ensureSheet_(ss, SHEET_DAILY, DAILY_HEADERS);
  bumpBoardVersion_();
  CacheService.getScriptCache().remove('cfg');
  try {
    SpreadsheetApp.getUi().alert('เก็บข้อมูลซีซั่นเดิมแล้ว — อย่าลืมแก้ seasonName / seasonStart / seasonEnd ในชีต Config');
  } catch (e) { /* รันจาก editor */ }
}

/* ============================================================
 *  ห้องแข่ง (Room) — สถานะเก็บใน CacheService, เขียนภายใต้ LockService
 * ============================================================ */

const ROOM_CHARS = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';

function roomGet_(code) {
  const raw = CacheService.getScriptCache().get('room_' + code);
  return raw ? JSON.parse(raw) : null;
}
function roomPut_(room) {
  CacheService.getScriptCache().put('room_' + room.code, JSON.stringify(room), ROOM_TTL_SEC);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}
function cleanCode_(c) {
  return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 6);
}
function cleanPid_(p) {
  return String(p || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 24);
}

function newPlayer_(b, now) {
  return {
    name: cleanText_(b.name, 40) || 'ผู้เล่น', dept: cleanText_(b.dept, 40),
    ready: false, joinedAt: now, lastSeen: now,
    score: 0, prog: null, clears: {}, rs: {}, done: false
  };
}

function sanitizePoseIds_(ids) {
  if (!Array.isArray(ids)) return [];
  return ids.slice(0, MAX_ROUNDS).map(x => String(x).replace(/[^a-z0-9_-]/gi, '').slice(0, 24)).filter(Boolean);
}

function apiRoomCreate_(b) {
  const pid = cleanPid_(b.pid);
  if (!pid) return { ok: false, error: 'no_pid' };
  const poseIds = sanitizePoseIds_(b.poseIds);
  if (poseIds.length < 1) return { ok: false, error: 'no_poses' };
  return withLock_(() => {
    let code = '';
    for (let tries = 0; tries < 20; tries++) {
      code = '';
      for (let i = 0; i < 4; i++) code += ROOM_CHARS.charAt(Math.floor(Math.random() * ROOM_CHARS.length));
      if (!roomGet_(code)) break;
    }
    const now = Date.now();
    const room = {
      code: code, hostId: pid, status: 'lobby', createdAt: now,
      seed: Math.floor(Math.random() * 1e9), poseIds: poseIds,
      roundMs: Math.round(num_(b.roundMs, 18000, 8000, 60000)),
      startAt: 0, game: 1, players: {}
    };
    room.players[pid] = newPlayer_(b, now);
    roomPut_(room);
    return { ok: true, room: publicRoom_(room, now) };
  });
}

function apiRoomJoin_(b) {
  const code = cleanCode_(b.code), pid = cleanPid_(b.pid);
  if (!code || !pid) return { ok: false, error: 'bad_request' };
  return withLock_(() => {
    const room = roomGet_(code);
    if (!room) return { ok: false, error: 'ไม่พบห้องนี้ (อาจหมดอายุแล้ว)' };
    const now = Date.now();
    refreshStatus_(room, now);
    if (!room.players[pid]) {
      if (room.status !== 'lobby') return { ok: false, error: 'ห้องนี้เริ่มเล่นไปแล้ว รอรอบถัดไปนะ' };
      if (Object.keys(room.players).length >= ROOM_MAX_PLAYERS) return { ok: false, error: 'ห้องเต็มแล้ว (สูงสุด ' + ROOM_MAX_PLAYERS + ' คน)' };
      room.players[pid] = newPlayer_(b, now);
    } else {
      room.players[pid].lastSeen = now;
      if (b.name) room.players[pid].name = cleanText_(b.name, 40);
    }
    roomPut_(room);
    return { ok: true, room: publicRoom_(room, now) };
  });
}

/**
 * ซิงก์สถานะ: ส่งอัปเดตของตัวเอง (ถ้ามี) แล้วรับสถานะห้องทั้งหมดกลับไป
 *  b.ready    : true/false
 *  b.prog     : { r: ด่าน, p: 0..1, n: จำนวนครั้ง }
 *  b.clear    : { r: ด่าน, ms: เวลาที่ใช้ (ms), pts: คะแนนด่านนั้น }
 *  b.fail     : { r: ด่าน, pts: คะแนนบางส่วน }
 *  b.done     : true เมื่อจบเกม
 *  b.leave    : true ออกจากห้อง
 */
function apiRoomSync_(b) {
  const code = cleanCode_(b.code), pid = cleanPid_(b.pid);
  if (!code || !pid) return { ok: false, error: 'bad_request' };
  const hasWrite = ('ready' in b) || b.prog || b.clear || b.fail || b.done || b.leave;
  const now = Date.now();
  if (!hasWrite) {
    // อ่านอย่างเดียว: ไม่ต้องล็อก (เร็วกว่า) — แต่อัปเดต lastSeen ทุก ๆ ~10 วินาที
    const room0 = roomGet_(code);
    if (!room0) return { ok: false, error: 'room_gone' };
    const me0 = room0.players[pid];
    if (me0 && now - (me0.lastSeen || 0) < 10000) {
      refreshStatus_(room0, now);
      return { ok: true, room: publicRoom_(room0, now) };
    }
  }
  return withLock_(() => {
    const room = roomGet_(code);
    if (!room) return { ok: false, error: 'room_gone' };
    refreshStatus_(room, now);
    const me = room.players[pid];
    if (!me) return { ok: false, error: 'not_in_room', room: publicRoom_(room, now) };
    me.lastSeen = now;
    if (b.leave) {
      delete room.players[pid];
      if (room.hostId === pid) {
        const rest = Object.keys(room.players);
        room.hostId = rest.length ? rest[0] : '';
      }
      roomPut_(room);
      return { ok: true, left: true };
    }
    if ('ready' in b && room.status === 'lobby') me.ready = !!b.ready;
    if (b.prog && typeof b.prog === 'object') {
      me.prog = { r: Math.round(num_(b.prog.r, 0, 0, MAX_ROUNDS)), p: num_(b.prog.p, 0, 0, 1), n: Math.round(num_(b.prog.n, 0, 0, 99)), g: room.game };
    }
    if (b.clear && typeof b.clear === 'object') {
      const r = String(Math.round(num_(b.clear.r, 0, 0, MAX_ROUNDS)));
      if (!(r in me.clears)) {
        me.clears[r] = Math.round(num_(b.clear.ms, 0, 0, 60000));
        me.rs[r] = Math.round(num_(b.clear.pts, 0, 0, MAX_ROUND_SCORE));
      }
    }
    if (b.fail && typeof b.fail === 'object') {
      const r = String(Math.round(num_(b.fail.r, 0, 0, MAX_ROUNDS)));
      if (!(r in me.rs)) me.rs[r] = Math.round(num_(b.fail.pts, 0, 0, MAX_ROUND_SCORE));
    }
    if (b.done) me.done = true;
    me.score = Object.keys(me.rs).reduce((s, k) => s + (me.rs[k] || 0), 0);
    // เตะคนที่หลุดนานในล็อบบี้
    if (room.status === 'lobby') {
      Object.keys(room.players).forEach(k => {
        if (k !== pid && now - room.players[k].lastSeen > ROOM_IDLE_KICK_MS) delete room.players[k];
      });
      if (!room.players[room.hostId]) room.hostId = pid;
    }
    roomPut_(room);
    return { ok: true, room: publicRoom_(room, now) };
  });
}

function apiRoomStart_(b) {
  const code = cleanCode_(b.code), pid = cleanPid_(b.pid);
  return withLock_(() => {
    const room = roomGet_(code);
    if (!room) return { ok: false, error: 'room_gone' };
    const now = Date.now();
    refreshStatus_(room, now);
    if (room.hostId !== pid) return { ok: false, error: 'เฉพาะเจ้าของห้องเท่านั้นที่กดเริ่มได้' };
    if (room.status !== 'lobby') return { ok: true, room: publicRoom_(room, now) };
    // ไม่รอคนที่ยังไม่กด READY — ตัดออกจากเกมนี้ (ยังอยู่ในห้องรอบหน้า)
    const ready = Object.keys(room.players).filter(k => room.players[k].ready);
    if (ready.length < 1) return { ok: false, error: 'ยังไม่มีใครกด READY' };
    room.status = 'playing';
    room.startAt = now + 5000;
    room.playing = ready;
    roomPut_(room);
    return { ok: true, room: publicRoom_(room, now) };
  });
}

/** เจ้าของห้องกด "เล่นอีกรอบ" — ใช้ห้องเดิม สุ่มท่าใหม่ */
function apiRoomReset_(b) {
  const code = cleanCode_(b.code), pid = cleanPid_(b.pid);
  const poseIds = sanitizePoseIds_(b.poseIds);
  return withLock_(() => {
    const room = roomGet_(code);
    if (!room) return { ok: false, error: 'room_gone' };
    const now = Date.now();
    if (room.hostId !== pid) return { ok: false, error: 'เฉพาะเจ้าของห้องเท่านั้น' };
    room.status = 'lobby';
    room.startAt = 0;
    room.game = (room.game || 1) + 1;
    room.seed = Math.floor(Math.random() * 1e9);
    if (poseIds.length) room.poseIds = poseIds;
    room.playing = [];
    Object.keys(room.players).forEach(k => {
      const p = room.players[k];
      p.ready = false; p.score = 0; p.prog = null; p.clears = {}; p.rs = {}; p.done = false;
    });
    roomPut_(room);
    return { ok: true, room: publicRoom_(room, now) };
  });
}

function refreshStatus_(room, now) {
  if (room.status === 'playing' && room.startAt) {
    const endAt = room.startAt + room.poseIds.length * room.roundMs + 8000;
    if (now > endAt) room.status = 'done';
  }
}

function publicRoom_(room, now) {
  const players = Object.keys(room.players).map(k => {
    const p = room.players[k];
    return {
      pid: k, name: p.name, dept: p.dept, ready: p.ready, host: k === room.hostId,
      playing: !room.playing || !room.playing.length || room.playing.indexOf(k) >= 0,
      score: p.score, prog: p.prog, clears: p.clears, rs: p.rs, done: p.done,
      online: now - p.lastSeen < 20000
    };
  });
  return {
    code: room.code, status: room.status, hostId: room.hostId, seed: room.seed,
    poseIds: room.poseIds, roundMs: room.roundMs, startAt: room.startAt, game: room.game || 1,
    players: players
  };
}

/* ============================================================
 *  Utils
 * ============================================================ */

function todayKey_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

/** รับ Date หรือข้อความ (2026-09-25 / 25/9/2569 / 25/09/2026) → 'yyyy-MM-dd' */
function toDateKey_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return fmtKey_(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (m) return fmtKey_(+m[3], +m[2], +m[1]);
  return '';
}
function fmtKey_(y, mo, d) {
  if (y > 2400) y -= 543; // พ.ศ. → ค.ศ.
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}
function keyToUtc_(k) {
  const p = k.split('-');
  return Date.UTC(+p[0], +p[1] - 1, +p[2]);
}
function addDays_(k, n) {
  const d = new Date(keyToUtc_(k) + n * 86400000);
  return fmtKey_(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}
function weekStart_(k) {
  const dow = new Date(keyToUtc_(k)).getUTCDay(); // 0 = อาทิตย์
  return addDays_(k, -((dow + 6) % 7));           // จันทร์
}
function thaiDate_(k) {
  if (!k) return '';
  const p = k.split('-');
  const mons = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  return (+p[2]) + ' ' + mons[+p[1] - 1] + ' ' + (+p[0] + 543);
}
function num_(v, def, min, max) {
  const n = Number(v);
  if (!isFinite(n)) return def;
  return Math.min(max, Math.max(min, n));
}
function cleanText_(v, max) {
  return String(v === undefined || v === null ? '' : v)
    .replace(/[​-‍﻿]/g, '')
    .replace(/[<>]/g, '')
    .replace(/^[=+\-@]+/, '')     // กันสูตรในชีต
    .replace(/\s+/g, ' ').trim().slice(0, max);
}
function playerKey_(name) {
  return cleanText_(name, 40).toLowerCase();
}
