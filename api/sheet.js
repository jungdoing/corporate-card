/**
 * 법인카드 관리 - Vercel 서버리스 함수 (구글 시트 직접 접근)
 * 서비스 계정으로 Sheets REST API 호출. 외부 npm 의존성 없음(Node 내장만 사용).
 *
 * 필요한 환경변수(Vercel Project Settings → Environment Variables):
 *  - SHEET_ID             : 스프레드시트 ID (URL 의 /d/<여기>/edit)
 *  - GOOGLE_CLIENT_EMAIL  : 서비스 계정 이메일 (...@...iam.gserviceaccount.com)
 *  - GOOGLE_PRIVATE_KEY   : 서비스 계정 개인키 (BEGIN PRIVATE KEY ... 전체, 줄바꿈은 \n 으로)
 *  - API_TOKEN            : 프론트(index.html)의 TOKEN 과 동일한 값
 */

const crypto = require('crypto');

const SHEET_ID = process.env.SHEET_ID;
const CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL;
const PRIVATE_KEY = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
const TOKEN = process.env.API_TOKEN;

const ENTRIES = '내역';
const SETTINGS = '설정';
const HEADERS = ['id', '날짜', '시간', '금액', '가맹점/메모', '카드사', '구분', '한도반영', '입력방식', '등록일시', '메모'];

/* ===== 인증: 서비스 계정 JWT → access token ===== */
let _tok = null, _exp = 0;
function b64url(buf) { return Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
async function accessToken() {
  const now = Math.floor(Date.now() / 1000);
  if (_tok && now < _exp - 60) return _tok;
  const header = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim = b64url(JSON.stringify({
    iss: CLIENT_EMAIL, scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token', exp: now + 3600, iat: now
  }));
  const sig = b64url(crypto.createSign('RSA-SHA256').update(header + '.' + claim).sign(PRIVATE_KEY));
  const jwt = header + '.' + claim + '.' + sig;
  const r = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: 'grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=' + jwt
  });
  const j = await r.json();
  if (!j.access_token) throw new Error('auth 실패: ' + JSON.stringify(j));
  _tok = j.access_token; _exp = now + (j.expires_in || 3600);
  return _tok;
}

/* ===== Sheets REST 헬퍼 ===== */
async function rest(path, method, body) {
  const tok = await accessToken();
  const r = await fetch('https://sheets.googleapis.com/v4/spreadsheets/' + SHEET_ID + path, {
    method: method || 'GET',
    headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined
  });
  const j = await r.json();
  if (j.error) throw new Error(j.error.message || JSON.stringify(j.error));
  return j;
}
async function getValues(range) {
  const j = await rest('/values/' + encodeURIComponent(range) + '?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER');
  return j.values || [];
}
async function updateValues(range, values) {
  await rest('/values/' + encodeURIComponent(range) + '?valueInputOption=RAW', 'PUT', { values });
}
async function appendValues(range, row) {
  await rest('/values/' + encodeURIComponent(range) + ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS', 'POST', { values: [row] });
}
async function batch(requests) { await rest(':batchUpdate', 'POST', { requests }); }

let _ids = null;
async function sheetIds() {
  if (_ids) return _ids;
  const j = await rest('?fields=sheets(properties(sheetId,title))');
  _ids = {};
  (j.sheets || []).forEach(s => { _ids[s.properties.title] = s.properties.sheetId; });
  return _ids;
}

/* ===== 시트 보장 ===== */
async function ensureEntries() {
  const ids = await sheetIds();
  if (ids[ENTRIES] == null) { await batch([{ addSheet: { properties: { title: ENTRIES } } }]); _ids = null; await updateValues(ENTRIES + '!A1', [HEADERS]); }
}
async function ensureSettings() {
  const ids = await sheetIds();
  if (ids[SETTINGS] == null) { await batch([{ addSheet: { properties: { title: SETTINGS } } }]); _ids = null; await updateValues(SETTINGS + '!A1', [['월', '한도']]); }
}

async function entriesTable() {
  await ensureEntries();
  let rows = await getValues(ENTRIES + '!A1:Z100000');
  if (rows.length === 0) { await updateValues(ENTRIES + '!A1', [HEADERS]); rows = [HEADERS.slice()]; }
  let header = rows[0].map(String);
  const missing = HEADERS.filter(h => header.indexOf(h) < 0);
  if (missing.length) {
    await updateValues(ENTRIES + '!' + colLetter(header.length + 1) + '1', [missing]);
    header = header.concat(missing);
  }
  const map = {}; header.forEach((h, i) => map[h] = i);
  return { header, map, data: rows.slice(1) };
}

/* ===== 비즈니스 로직 ===== */
async function getMonth(ym) {
  ym = ym || ymNow();
  const t = await entriesTable(), m = t.map;
  const entries = [];
  for (let i = 0; i < t.data.length; i++) {
    const r = t.data[i];
    const id = r[m['id']]; if (id === '' || id == null) continue;
    const date = normDate(r[m['날짜']]);
    if (String(date).indexOf(ym) !== 0) continue;
    entries.push({
      id: String(id), date, time: normTime(r[m['시간']]),
      amount: Number(r[m['금액']]) || 0, memo: r[m['가맹점/메모']] || '', card: r[m['카드사']] || '',
      type: r[m['구분']] || '승인', inLimit: (r[m['한도반영']] === '제외') ? '제외' : '적용',
      note: r[m['메모']] || '', source: r[m['입력방식']] || ''
    });
  }
  entries.sort((a, b) => a.date !== b.date ? (a.date < b.date ? 1 : -1) : (a.time < b.time ? 1 : (a.time > b.time ? -1 : 0)));
  let used = 0, extra = 0;
  entries.forEach(x => { const v = (x.type === '취소' ? -1 : 1) * x.amount; if (x.inLimit === '제외') extra += v; else used += v; });
  let sv = []; try { sv = await getValues(SETTINGS + '!A2:B100000'); } catch (e) {}
  let limit = 0;
  for (const row of sv) { if (normYm(row[0]) === ym) { limit = Number(row[1]) || 0; break; } }
  return { ym, limit, used, extra, remaining: limit - used, entries };
}

async function setLimit(ym, limit) {
  await ensureSettings();
  limit = Number(limit) || 0;
  let sv = []; try { sv = await getValues(SETTINGS + '!A2:B100000'); } catch (e) {}
  for (let i = 0; i < sv.length; i++) {
    if (normYm(sv[i][0]) === ym) { await updateValues(SETTINGS + '!A' + (i + 2), [[ym, limit]]); return getMonth(ym); }
  }
  await appendValues(SETTINGS + '!A1', [ym, limit]);
  return getMonth(ym);
}

async function addEntry(e) {
  const t = await entriesTable(), m = t.map;
  const date = normDate(e.date) || ymdNow();
  const row = new Array(t.header.length).fill('');
  set(row, m, 'id', uuid());
  set(row, m, '날짜', date);
  set(row, m, '시간', e.time || '');
  set(row, m, '금액', Number(e.amount) || 0);
  set(row, m, '가맹점/메모', e.memo || '');
  set(row, m, '카드사', e.card || '');
  set(row, m, '구분', e.type || '승인');
  set(row, m, '한도반영', e.inLimit === '제외' ? '제외' : '적용');
  set(row, m, '입력방식', e.source || '수동');
  set(row, m, '등록일시', nowStr());
  set(row, m, '메모', e.note || '');
  await appendValues(ENTRIES + '!A1', row);
  return getMonth(date.substring(0, 7));
}

async function updateEntry(id, e) {
  const t = await entriesTable(), m = t.map;
  for (let i = 0; i < t.data.length; i++) {
    if (String(t.data[i][m['id']]) === String(id)) {
      const row = t.data[i].slice();
      while (row.length < t.header.length) row.push('');
      const date = normDate(e.date) || normDate(row[m['날짜']]);
      set(row, m, '날짜', date);
      set(row, m, '시간', e.time || '');
      set(row, m, '금액', Number(e.amount) || 0);
      set(row, m, '가맹점/메모', e.memo || '');
      set(row, m, '카드사', e.card || '');
      set(row, m, '구분', e.type || '승인');
      set(row, m, '한도반영', e.inLimit === '제외' ? '제외' : '적용');
      set(row, m, '메모', e.note || '');
      const rr = i + 2;
      await updateValues(ENTRIES + '!A' + rr + ':' + colLetter(t.header.length) + rr, [row]);
      return getMonth(date.substring(0, 7));
    }
  }
  return getMonth(normDate(e.date).substring(0, 7));
}

async function deleteEntry(id, ym) {
  const t = await entriesTable(), m = t.map;
  const ids = await sheetIds();
  for (let i = 0; i < t.data.length; i++) {
    if (String(t.data[i][m['id']]) === String(id)) {
      const idx = i + 1; // 0-based, 헤더 포함 → 데이터행 i 는 시트 인덱스 i+1
      await batch([{ deleteDimension: { range: { sheetId: ids[ENTRIES], dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 } } }]);
      break;
    }
  }
  return getMonth(ym);
}

/* ===== 유틸 ===== */
function set(row, m, k, v) { if (m[k] != null) row[m[k]] = v; }
function uuid() { return crypto.randomUUID(); }
function p2(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
function colLetter(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function kst() { return new Date(Date.now() + 9 * 3600 * 1000); }
function ymNow() { const d = kst(); return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1); }
function ymdNow() { const d = kst(); return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate()); }
function nowStr() { const d = kst(); return ymdNow() + ' ' + p2(d.getUTCHours()) + ':' + p2(d.getUTCMinutes()) + ':' + p2(d.getUTCSeconds()); }
function serialToDate(n) { return new Date(Math.round((n - 25569) * 86400 * 1000)); }
function normDate(v) {
  if (typeof v === 'number') { const d = serialToDate(v); return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate()); }
  const s = String(v == null ? '' : v).trim();
  const m = s.match(/^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
  return m ? m[1] + '-' + p2(m[2]) + '-' + p2(m[3]) : s;
}
function normTime(v) {
  if (typeof v === 'number') { const d = serialToDate(v); return p2(d.getUTCHours()) + ':' + p2(d.getUTCMinutes()); }
  const s = String(v == null ? '' : v).trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  return m ? p2(m[1]) + ':' + m[2] : s;
}
function normYm(v) {
  if (typeof v === 'number') { const d = serialToDate(v); return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1); }
  return String(v == null ? '' : v).trim();
}

/* ===== 핸들러 ===== */
module.exports = async (req, res) => {
  if (req.method !== 'POST') { res.status(200).json({ ok: true, service: '법인카드 관리 API' }); return; }
  try {
    if (!SHEET_ID || !CLIENT_EMAIL || !PRIVATE_KEY) { res.status(200).json({ error: '서버 환경변수 미설정 (SHEET_ID/GOOGLE_CLIENT_EMAIL/GOOGLE_PRIVATE_KEY)' }); return; }
    let body = req.body;
    if (typeof body === 'string') body = JSON.parse(body || '{}');
    if (!body || typeof body !== 'object') body = {};
    if (body.token !== TOKEN) { res.status(200).json({ error: 'unauthorized' }); return; }
    let data;
    switch (body.action) {
      case 'getMonth': data = await getMonth(body.ym); break;
      case 'setLimit': data = await setLimit(body.ym, body.limit); break;
      case 'add': data = await addEntry(body.entry); break;
      case 'update': data = await updateEntry(body.id, body.entry); break;
      case 'delete': data = await deleteEntry(body.id, body.ym); break;
      default: res.status(200).json({ error: 'unknown action: ' + body.action }); return;
    }
    res.status(200).json({ ok: true, data });
  } catch (err) {
    res.status(200).json({ error: String(err && err.message ? err.message : err) });
  }
};
