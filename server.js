const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 8080;
const BASE_DIR = __dirname;
const DS_DIR = path.join(BASE_DIR, 'ds');
const CONFIG_FILE = path.join(BASE_DIR, 'config.json');

// Ensure ds directory exists
if (!fs.existsSync(DS_DIR)) {
  try { fs.mkdirSync(DS_DIR, { recursive: true }); } catch (e) {}
}

const yearData = {};
const yearDepts = {};
let availableYears = [];

// Helper to load data for a year (checks ds/ first, then BASE_DIR)
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

      const deptMap = {};
      list.forEach(m => {
        const d = m.dept ?? 'N/A';
        const dName = m.deptname || 'ไม่ระบุ';
        const key = `${d}`;
        if (!deptMap[key]) {
          deptMap[key] = { dept: d, deptname: dName, count: 0 };
        }
        deptMap[key].count++;
      });

      yearDepts[yearKey] = Object.values(deptMap).sort((a, b) => {
        const na = parseInt(a.dept, 10);
        const nb = parseInt(b.dept, 10);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(a.dept).localeCompare(String(b.dept));
      });

      console.log(`Loaded Year ${yearKey} from ${path.relative(BASE_DIR, filePath)}: ${list.length} records, ${yearDepts[yearKey].length} departments.`);
    } catch (e) {
      console.error(`Error loading ${filePath}:`, e);
    }
  }
}

function initData() {
  // Clear old cache before reloading
  Object.keys(yearData).forEach(k => delete yearData[k]);
  Object.keys(yearDepts).forEach(k => delete yearDepts[k]);
  availableYears = [];

  const yearSet = new Set();
  const yearRegex = /^data(\d{4})\.json$/i;

  // Scan files in ds directory
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

  // Also scan BASE_DIR for backward compatibility
  try {
    const baseFiles = fs.readdirSync(BASE_DIR);
    baseFiles.forEach(f => {
      const m = f.match(yearRegex);
      if (m) yearSet.add(m[1]);
    });
  } catch (err) {
    console.error('Error scanning BASE_DIR:', err);
  }

  // Sort years descending (newest first, e.g. 2028, 2027, 2026, 2025)
  const sortedYears = Array.from(yearSet).sort((a, b) => parseInt(b, 10) - parseInt(a, 10));

  sortedYears.forEach(yKey => {
    loadYearData(yKey, `data${yKey}.json`);
    if (yearData[yKey]) {
      const beYear = parseInt(yKey, 10) + 543;
      availableYears.push({
        id: yKey,
        label: `ข้อมูลปี ${beYear} (${yKey})`,
        count: yearData[yKey].length
      });
    }
  });

  console.log(`[initData] Completed loading. Available years: ${availableYears.map(y => `${y.id}(${y.count})`).join(', ')}`);
}

initData();

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
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.pdf': 'application/pdf'
  };
  return mimeTypes[ext] || 'application/octet-stream';
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

const server = http.createServer((req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;
  const method = req.method;

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // API: Reload data from JSON files
  if (pathname === '/api/reload') {
    initData();
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ success: true, message: 'โหลดข้อมูลใหม่เรียบร้อยแล้ว', years: availableYears }));
    return;
  }

  // API: Get available years
  if (pathname === '/api/years' && method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(availableYears));
    return;
  }

  // API: Config
  if (pathname === '/api/config') {
    const qYear = parsedUrl.query.year || '2026';
    const allCfg = readFullConfig();

    if (method === 'GET') {
      const yearCfg = allCfg[qYear] || allCfg['2026'] || allCfg['2025'] || allCfg;
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify(yearCfg));
      return;
    } else if (method === 'POST') {
      let body = '';
      req.on('data', chunk => body += chunk);
      req.on('end', () => {
        try {
          const newYearCfg = JSON.parse(body);
          allCfg[qYear] = newYearCfg;
          fs.writeFileSync(CONFIG_FILE, JSON.stringify(allCfg, null, 2), 'utf-8');
          res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ success: true, message: `บันทึกการตั้งค่าปี ${qYear} เรียบร้อยแล้ว` }));
        } catch (err) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: err.message }));
        }
      });
      return;
    }
  }

  // API: Departments list by year
  if (pathname === '/api/depts' && method === 'GET') {
    const qYear = parsedUrl.query.year || (availableYears[0] ? availableYears[0].id : '2026');
    const depts = yearDepts[qYear] || [];
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(depts));
    return;
  }

  // API: Members list by year
  if (pathname === '/api/members' && method === 'GET') {
    const q = parsedUrl.query;
    const qYear = q.year || (availableYears[0] ? availableYears[0].id : '2026');
    let filtered = yearData[qYear] || [];

    // Filter by specific memos
    if (q.memos) {
      const memoList = q.memos.split(',').map(s => s.trim().padStart(5, '0'));
      const memoSet = new Set(memoList);
      filtered = filtered.filter(m => memoSet.has(String(m.memo).padStart(5, '0')));
    } else {
      // Status filter
      if (q.status_filter === 'exclude_inactive' || !q.status_filter) {
        filtered = filtered.filter(m => m.status !== 'เสียชีวิต' && m.status !== 'ให้ออก' && m.status !== 'พ้นสภาพ' && m.status !== 'พ้นสมาชิก');
      } else if (q.status_filter === 'ศาลอายัด') {
        filtered = filtered.filter(m => m.status === 'ศาลอายัด');
      } else if (q.status_filter === 'เงินสด') {
        filtered = filtered.filter(m => m.status === 'เงินสด');
      } else if (q.status_filter === 'โอนธนาคาร') {
        filtered = filtered.filter(m => m.status !== 'ศาลอายัด' && m.status !== 'เงินสด' && m.status !== 'เสียชีวิต' && m.status !== 'ให้ออก' && m.status !== 'พ้นสภาพ' && m.status !== 'พ้นสมาชิก');
      }

      // Dept filter
      if (q.dept && q.dept !== 'all') {
        filtered = filtered.filter(m => String(m.dept) === String(q.dept));
      }

      // Memo range
      if (q.memo_from) {
        const fromNum = parseInt(q.memo_from, 10);
        filtered = filtered.filter(m => {
          const mNum = parseInt(m.memo, 10);
          return !isNaN(mNum) && mNum >= fromNum;
        });
      }
      if (q.memo_to) {
        const toNum = parseInt(q.memo_to, 10);
        filtered = filtered.filter(m => {
          const mNum = parseInt(m.memo, 10);
          return !isNaN(mNum) && mNum <= toNum;
        });
      }

      // General search
      if (q.search) {
        const term = q.search.trim().toLowerCase();
        filtered = filtered.filter(m => {
          const memoStr = String(m.memo || '').toLowerCase();
          const idpStr = String(m.idp || '').toLowerCase();
          const nameStr = String(m.name || '').toLowerCase();
          const deptNameStr = String(m.deptname || '').toLowerCase();
          return memoStr.includes(term) || idpStr.includes(term) || nameStr.includes(term) || deptNameStr.includes(term);
        });
      }
    }

    const total = filtered.length;
    const page = parseInt(q.page, 10) || 1;
    const limit = parseInt(q.limit, 10) || (q.all === '1' ? total : 50);

    const startIndex = (page - 1) * limit;
    const paginated = limit > 0 ? filtered.slice(startIndex, startIndex + limit) : filtered;

    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      year: qYear,
      total,
      page,
      limit,
      totalPages: limit > 0 ? Math.ceil(total / limit) : 1,
      data: paginated
    }));
    return;
  }

  // API: Download Excel by Year
  if (pathname === '/download-excel') {
    const qYear = parsedUrl.query.year || (availableYears[0] ? availableYears[0].id : '2026');
    let excelFile = path.join(DS_DIR, `data${qYear}.xlsx`);
    if (!fs.existsSync(excelFile)) {
      excelFile = path.join(BASE_DIR, `data${qYear}.xlsx`);
    }

    if (fs.existsSync(excelFile)) {
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="data${qYear}.xlsx"`
      });
      fs.createReadStream(excelFile).pipe(res);
      return;
    } else {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end(`ไม่พบไฟล์ Excel สำหรับปี ${qYear}`);
      return;
    }
  }

  // Static File Serving
  let reqPath = pathname === '/' ? '/index.html' : pathname;
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

server.listen(PORT, () => {
  console.log(`=====================================================`);
  console.log(`ระบบรายงานเงินปันผลเฉลี่ยคืน สหกรณ์ออมทรัพย์ บก.ทท.`);
  console.log(`รองรับระบบหลายปี (Multi-Year Support): ${availableYears.map(y => y.id).join(', ')}`);
  console.log(`Server is running at: http://localhost:${PORT}`);
  console.log(`=====================================================`);
});
