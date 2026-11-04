const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const crypto = require('crypto');

// ============================================================================
// ⚙️ การตั้งค่าระบบสมาชิก (Member Portal Settings)
// ============================================================================
const PORT = process.env.PORT || 8088;

// กำหนดปีงบประมาณปัจจุบัน (Current Fiscal Year)
// หากต้องการระบุเจาะจงสามารถใส่ได้ เช่น '2026' หรือส่งผ่าน Environment Variable: CURRENT_YEAR
// หากปล่อยว่าง ('') ระบบจะเลือกปีล่าสุดที่มีข้อมูลในระบบให้อัตโนมัติ (เช่น 2026 / 2569)
const ACTIVE_YEAR = process.env.CURRENT_YEAR || '';

const BASE_DIR = __dirname;
const DS_DIR = path.join(BASE_DIR, 'ds');
const CONFIG_FILE = path.join(BASE_DIR, 'config.json');

const yearData = {};
let availableYears = [];

// Session store: token -> { memo, year, createdAt }
const sessions = new Map();
const SESSION_MAX_AGE_MS = 24 * 60 * 60 * 1000; // 24 hours

// Clean expired sessions periodically
setInterval(() => {
  const now = Date.now();
  for (const [token, sess] of sessions.entries()) {
    if (now - sess.createdAt > SESSION_MAX_AGE_MS) {
      sessions.delete(token);
    }
  }
}, 60 * 60 * 1000);

// Load data from ds/
function loadYearData(yearKey, jsonFileName) {
  let filePath = path.join(DS_DIR, jsonFileName);
  if (!fs.existsSync(filePath)) {
    filePath = path.join(BASE_DIR, jsonFileName);
  }

  if (fs.existsSync(filePath)) {
    try {
      const raw = fs.readFileSync(filePath, 'utf-8');
      const list = JSON.parse(raw);
      yearData[yearKey] = list;
      console.log(`[Member Server] Loaded Year ${yearKey}: ${list.length} records.`);
    } catch (e) {
      console.error(`[Member Server] Error loading ${filePath}:`, e);
    }
  }
}

function initData() {
  Object.keys(yearData).forEach(k => delete yearData[k]);
  availableYears = [];

  const yearSet = new Set();
  const yearRegex = /^data(\d{4})\.json$/i;

  if (fs.existsSync(DS_DIR)) {
    try {
      const dsFiles = fs.readdirSync(DS_DIR);
      dsFiles.forEach(f => {
        const m = f.match(yearRegex);
        if (m) yearSet.add(m[1]);
      });
    } catch (err) {
      console.error('Error scanning DS_DIR:', err);
    }
  }

  try {
    const baseFiles = fs.readdirSync(BASE_DIR);
    baseFiles.forEach(f => {
      const m = f.match(yearRegex);
      if (m) yearSet.add(m[1]);
    });
  } catch (err) {}

  const sortedYears = Array.from(yearSet).sort((a, b) => parseInt(b, 10) - parseInt(a, 10));

  sortedYears.forEach(yKey => {
    loadYearData(yKey, `data${yKey}.json`);
    if (yearData[yKey]) {
      const beYear = parseInt(yKey, 10) + 543;
      availableYears.push({
        id: yKey,
        beYear: beYear,
        label: `ปีงบประมาณ ${beYear} (${yKey})`,
        count: yearData[yKey].length
      });
    }
  });

  console.log(`[Member Server] Available years: ${availableYears.map(y => `${y.id}(${y.count})`).join(', ')}`);
}

initData();

// Helper: Get active current year
function getCurrentYear() {
  if (ACTIVE_YEAR && yearData[ACTIVE_YEAR]) {
    return ACTIVE_YEAR;
  }
  return availableYears[0] ? availableYears[0].id : '2026';
}

function readFullConfig() {
  try {
    if (fs.existsSync(CONFIG_FILE)) {
      return JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf-8'));
    }
  } catch (err) {
    console.error('Error reading config file:', err);
  }
  return {};
}

function getMimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const mimeTypes = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
    '.pdf': 'application/pdf'
  };
  return mimeTypes[ext] || 'application/octet-stream';
}

function parseCookies(req) {
  const list = {};
  const cookieHeader = req.headers.cookie;
  if (!cookieHeader) return list;

  cookieHeader.split(';').forEach(cookie => {
    let [name, ...rest] = cookie.split('=');
    name = name?.trim();
    if (!name) return;
    const value = rest.join('=').trim();
    list[name] = decodeURIComponent(value);
  });
  return list;
}

function getSession(req) {
  const cookies = parseCookies(req);
  const token = cookies.member_sid;
  if (token && sessions.has(token)) {
    const sess = sessions.get(token);
    if (Date.now() - sess.createdAt <= SESSION_MAX_AGE_MS) {
      return sess;
    } else {
      sessions.delete(token);
    }
  }
  return null;
}

// Mask sensitive ID card for display: e.g. 3100501754770 -> 3-1005-XXXXX-77-0
function maskCitizenId(idp) {
  if (!idp || typeof idp !== 'string' && typeof idp !== 'number') return '-';
  const str = String(idp).trim();
  if (str.length === 13) {
    return `${str[0]}-${str.substring(1, 5)}-xxxxx-${str.substring(9, 11)}-${str[12]}`;
  }
  return str.length > 4 ? `xxxxxxx${str.slice(-4)}` : str;
}

// Mask bank account number: e.g. 0392256020 -> 039-x-xx602-0
function maskBankAccount(acc) {
  if (!acc) return '-';
  const str = String(acc).trim();
  if (str.length === 10) {
    return `${str.substring(0, 3)}-x-xx${str.substring(6, 9)}-${str[9]}`;
  }
  return str.length > 4 ? `xxx-xx${str.slice(-4)}` : str;
}

const server = http.createServer((req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;
  const method = req.method;

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // 1. API: Current Fiscal Year (ปีปัจจุบันอย่างเดียว)
  if (pathname === '/api/current-year' && method === 'GET') {
    const curYear = getCurrentYear();
    const beYear = parseInt(curYear, 10) + 543;
    const allCfg = readFullConfig();
    const yearCfg = allCfg[curYear] || allCfg['2026'] || allCfg['2025'] || {};
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      year: curYear,
      beYear: beYear,
      label: `ปีงบประมาณ ${beYear}`,
      config: yearCfg
    }));
    return;
  }

  // API: Available Years (backward compatibility)
  if (pathname === '/api/years' && method === 'GET') {
    const curYear = getCurrentYear();
    const beYear = parseInt(curYear, 10) + 543;
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify([{
      id: curYear,
      beYear: beYear,
      label: `ปีงบประมาณ ${beYear}`,
      count: yearData[curYear] ? yearData[curYear].length : 0
    }]));
    return;
  }

  // 2. API: Config (reused by print.html and member dashboard)
  if (pathname === '/api/config' && method === 'GET') {
    const qYear = parsedUrl.query.year || getCurrentYear();
    const allCfg = readFullConfig();
    const yearCfg = allCfg[qYear] || allCfg['2026'] || allCfg['2025'] || {};
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(yearCfg));
    return;
  }

  // 3. API: Member Login (ดึงปีปัจจุบันอย่างเดียว)
  if (pathname === '/api/member/login' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { memo, id4 } = JSON.parse(body || '{}');

        if (!memo || !id4) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ success: false, message: 'กรุณากรอกรหัสสมาชิกและเลขบัตรประชาชน 4 หลักท้าย' }));
          return;
        }

        // ล็อกปีปัจจุบันอย่างเดียว
        const targetYear = getCurrentYear();
        const beYear = parseInt(targetYear, 10) + 543;
        const list = yearData[targetYear];

        if (!list || list.length === 0) {
          res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ success: false, message: `ไม่พบชุดข้อมูลสำหรับปีงบประมาณ ${beYear}` }));
          return;
        }

        // Normalize memo: pad to 5 digits (e.g. 22 -> 00022)
        const cleanMemoNum = parseInt(memo, 10);
        const paddedMemo = !isNaN(cleanMemoNum) ? String(cleanMemoNum).padStart(5, '0') : String(memo).trim();
        const inputId4 = String(id4).trim();

        // Find member
        const member = list.find(m => {
          const mMemo = String(m.memo || '').padStart(5, '0');
          return mMemo === paddedMemo;
        });

        if (!member) {
          res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ success: false, message: `ไม่พบรหัสสมาชิก "${paddedMemo}" ในข้อมูลปี ${beYear}` }));
          return;
        }

        // Verify last 4 digits of citizen ID (idp)
        const idpStr = String(member.idp || '').trim();
        const idpLast4 = idpStr.slice(-4);

        if (idpLast4 !== inputId4) {
          res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ success: false, message: 'เลขบัตรประชาชน 4 หลักท้ายไม่ถูกต้อง' }));
          return;
        }

        // Generate session token
        const token = crypto.randomBytes(32).toString('hex');
        sessions.set(token, {
          memo: paddedMemo,
          year: targetYear,
          createdAt: Date.now()
        });

        // Set HttpOnly cookie
        res.setHeader('Set-Cookie', `member_sid=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=86400`);
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });

        // Safe sanitized member data for client view
        const safeMember = { ...member };
        safeMember.idp_masked = maskCitizenId(member.idp);
        safeMember.tmb_masked = maskBankAccount(member.tmb);
        safeMember.idp_last4 = idpLast4;

        res.end(JSON.stringify({
          success: true,
          message: 'เข้าสู่ระบบสำเร็จ',
          year: targetYear,
          beYear: beYear,
          member: safeMember
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ success: false, message: 'เกิดข้อผิดพลาดในการประมวลผล: ' + err.message }));
      }
    });
    return;
  }

  // 4. API: Check current session (ปีปัจจุบัน)
  if (pathname === '/api/member/me' && method === 'GET') {
    const sess = getSession(req);
    if (!sess) {
      res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ authenticated: false }));
      return;
    }

    const curYear = getCurrentYear();
    const beYear = parseInt(curYear, 10) + 543;
    const list = yearData[curYear] || [];
    const member = list.find(m => String(m.memo || '').padStart(5, '0') === sess.memo);

    if (!member) {
      res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({
        authenticated: true,
        foundInYear: false,
        memo: sess.memo,
        year: curYear,
        beYear: beYear,
        message: `ไม่พบข้อมูลสมาชิกในงวดปี ${beYear}`
      }));
      return;
    }

    const safeMember = { ...member };
    safeMember.idp_masked = maskCitizenId(member.idp);
    safeMember.tmb_masked = maskBankAccount(member.tmb);
    safeMember.idp_last4 = String(member.idp || '').trim().slice(-4);

    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      authenticated: true,
      foundInYear: true,
      year: curYear,
      beYear: beYear,
      member: safeMember
    }));
    return;
  }

  // 5. API: Member Logout
  if (pathname === '/api/member/logout' && method === 'POST') {
    const cookies = parseCookies(req);
    const token = cookies.member_sid;
    if (token) {
      sessions.delete(token);
    }
    res.setHeader('Set-Cookie', 'member_sid=; Path=/; HttpOnly; Max-Age=0');
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ success: true, message: 'ออกจากระบบเรียบร้อยแล้ว' }));
    return;
  }

  // 6. API: Members endpoint for print.html compatibility!
  // Reuses the exact print.html without modifying it, while protecting other members' privacy.
  if (pathname === '/api/members' && method === 'GET') {
    const q = parsedUrl.query;
    const curYear = getCurrentYear();
    const sess = getSession(req);

    // If user is not authenticated, block access to prevent member data leaks
    if (!sess) {
      res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ error: 'Unauthorized: กรุณาเข้าสู่ระบบก่อนพิมพ์รายงาน', data: [] }));
      return;
    }

    const list = yearData[curYear] || [];
    // Only return the authenticated member's record!
    const member = list.find(m => String(m.memo || '').padStart(5, '0') === sess.memo);

    if (member) {
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({
        year: curYear,
        total: 1,
        page: 1,
        limit: 1,
        totalPages: 1,
        data: [member]
      }));
    } else {
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({
        year: curYear,
        total: 0,
        page: 1,
        limit: 1,
        totalPages: 1,
        data: []
      }));
    }
    return;
  }

  // 7. Static File Serving
  let reqPath = pathname === '/' ? '/member.html' : pathname;
  if (reqPath === '/index.html') {
    reqPath = '/member.html';
  }

  const safePath = path.normalize(reqPath).replace(/^(\.\.[\/\\])+/, '');
  let filePath = path.join(BASE_DIR, safePath);

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    const mime = getMimeType(filePath);
    res.writeHead(200, { 'Content-Type': mime });
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('404 Not Found: ' + pathname);
  }
});

const curYear = getCurrentYear();
const beYear = parseInt(curYear, 10) + 543;

server.listen(PORT, () => {
  console.log(`================================================================`);
  console.log(`🚀 ระบบตรวจสอบเงินปันผลเฉลี่ยคืน สำหรับสมาชิก (Member Portal)`);
  console.log(`สหกรณ์ออมทรัพย์กองบัญชาการกองทัพไทย จำกัด`);
  console.log(`📌 ปีงบประมาณปัจจุบัน: ${beYear} (${curYear})`);
  console.log(`Member Portal กำลังทำงานที่: http://localhost:${PORT}`);
  console.log(`================================================================`);
});
