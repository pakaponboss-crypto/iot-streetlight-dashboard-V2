'use strict';
const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');
const fs      = require('fs');
const drive   = require('./drive');

const app  = express();
const PORT = 8501;

// ── Local cache dir for raw pole files (speed up re-reads) ───────────────────
const CACHE_DIR = path.join(__dirname, 'cache');
fs.mkdirSync(CACHE_DIR, { recursive: true });

// ── DB cache (refresh every 5s to keep multi-user in sync) ───────────────────
let _dbCache = null;
let _dbTime  = 0;
const DB_TTL = 5000;

async function readDB() {
  const now = Date.now();
  if (_dbCache && (now - _dbTime) < DB_TTL) return _dbCache;
  const data = await drive.readJSON('db.json');
  _dbCache = data || { uploads: [], rows: [] };
  _dbTime  = now;
  return _dbCache;
}
async function writeDB(db) {
  _dbCache = db;
  _dbTime  = Date.now();
  await drive.writeJSON('db.json', db);
}

// ── Raw pole data: local cache → Drive fallback ──────────────────────────────
function _cachePath(uploadId) { return path.join(CACHE_DIR, `raw_${uploadId}.json`); }

async function readRaw(uploadId) {
  const local = _cachePath(uploadId);
  if (fs.existsSync(local)) {
    try { return JSON.parse(fs.readFileSync(local, 'utf8')); } catch {}
  }
  const data = await drive.readJSON(`raw_${uploadId}.json`);
  if (data) { fs.writeFileSync(local, JSON.stringify(data)); return data; }
  return [];
}
async function writeRaw(uploadId, rows) {
  fs.writeFileSync(_cachePath(uploadId), JSON.stringify(rows)); // local cache
  await drive.writeJSON(`raw_${uploadId}.json`, rows);          // Drive
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function nextId(arr)   { return arr.length ? Math.max(...arr.map(x => x.id)) + 1 : 1; }
function safeSheet(n)  { return String(n).replace(/[:\\/?*[\]]/g, '-').slice(0, 31); }

// ── Column Detection ─────────────────────────────────────────────────────────
// Format A: pre-aggregated  — ผู้รับเหมา, เลขที่สัญญา, max โคม, จำนวนจริง, IOT ได้, API เชื่อม
// Format B: pole-level      — รหัสเสาไฟ, เลขที่สัญญา, ผู้รับเหมา, ควบคุมiot, จำนวนโคม, สถานะข้อมูลจากAPI

function cleanH(h) { return String(h).trim().replace(/^["']+|["']+$/g, '').trim(); }

function detectFormat(rawHeaders) {
  const cleaned = rawHeaders.map(h => ({ orig: h, clean: cleanH(h).toLowerCase() }));
  const isPoleLevel = cleaned.some(c =>
    c.clean.includes('ควบคุมiot') || c.clean.includes('รหัสเสาไฟ') || c.clean.includes('สถานะข้อมูลจากapi')
  );

  if (isPoleLevel) {
    const m = {};
    for (const { orig, clean } of cleaned) {
      if (!m.contract_no && (clean.includes('เลขที่สัญญา') || clean.includes('สัญญา')))      m.contract_no  = orig;
      if (!m.contractor  && (clean.includes('ผู้รับเหมา')  || clean.includes('ผู้รับจ้าง'))) m.contractor   = orig;
      if (!m.max_poles   && (clean.includes('จำนวนโคม')    || clean.includes('maxโคม')))      m.max_poles    = orig;
      if (!m.iot_control && (clean.includes('ควบคุมiot')   || clean.includes('iot')))         m.iot_control  = orig;
      if (!m.api_status  && (clean.includes('สถานะข้อมูลจากapi') || clean.includes('api')))   m.api_status   = orig;
    }
    return { format: 'B', mapping: m, missing: ['contract_no','contractor','iot_control'].filter(f => !m[f]) };
  }

  const hints = {
    contractor:    ['ผู้รับเหมา','contractor','บริษัท'],
    contract_no:   ['เลขที่สัญญา','สัญญา','contract'],
    max_poles:     ['max โคม','maxโคม','max'],
    actual_poles:  ['จำนวนจริง','จำนวนเสา','actual','จำนวน'],
    iot_available: ['iot ได้','iotได้','iot'],
    api_connected: ['api เชื่อม','apiเชื่อม','api'],
  };
  const m = {};
  for (const { orig, clean } of cleaned) {
    for (const [field, hs] of Object.entries(hints)) {
      if (!m[field] && hs.some(h => clean.includes(h.toLowerCase()))) m[field] = orig;
    }
  }
  return { format: 'A', mapping: m, missing: Object.keys(hints).filter(f => !m[f]) };
}

function aggregatePoleData(rows, mapping) {
  const by = {};
  for (const row of rows) {
    const cn = String(row[mapping.contract_no] || '').replace(/^["']+|["']+$/g,'').trim();
    if (!cn) continue;
    if (!by[cn]) by[cn] = {
      contractor:    String(row[mapping.contractor] || '').replace(/^["']+|["']+$/g,'').trim(),
      contract_no:   cn,
      max_poles:     mapping.max_poles ? (parseInt(row[mapping.max_poles]) || 0) : 0,
      actual_poles:  0, iot_available: 0, api_connected: 0,
    };
    by[cn].actual_poles++;
    if (String(row[mapping.iot_control]||'').replace(/^["']+|["']+$/g,'').trim() === 'ได้') by[cn].iot_available++;
    const av = row[mapping.api_status];
    if (av === 1 || av === '1' || av === true) by[cn].api_connected++;
  }
  return Object.values(by);
}

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// ── GET /api/uploads ─────────────────────────────────────────────────────────
app.get('/api/uploads', async (req, res) => {
  try {
    const db = await readDB();
    res.json([...db.uploads].reverse());
  } catch (e) { res.status(500).json({ error: String(e) }); }
});

// ── POST /api/upload ─────────────────────────────────────────────────────────
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    const wb   = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) return res.status(400).json({ error: 'ไฟล์ไม่มีข้อมูล' });

    const { format, mapping, missing } = detectFormat(Object.keys(rows[0]));
    if (missing.length) return res.status(400).json({
      error: `ไม่พบคอลัมน์ที่จำเป็น: ${missing.join(', ')}\nคอลัมน์ในไฟล์: ${Object.keys(rows[0]).join(', ')}`,
    });

    let summaryRows;
    if (format === 'B') {
      summaryRows = aggregatePoleData(rows, mapping);
    } else {
      summaryRows = rows.map(row => ({
        contractor:    String(row[mapping.contractor]    || ''),
        contract_no:   String(row[mapping.contract_no]  || '').trim(),
        max_poles:     parseInt(row[mapping.max_poles])     || 0,
        actual_poles:  parseInt(row[mapping.actual_poles])  || 0,
        iot_available: parseInt(row[mapping.iot_available]) || 0,
        api_connected: parseInt(row[mapping.api_connected]) || 0,
      })).filter(r => r.contract_no);
    }

    const now      = new Date();
    const db       = await readDB();
    const uploadId = nextId(db.uploads);

    db.uploads.push({
      id: uploadId,
      upload_date: now.toISOString().split('T')[0],
      upload_time: now.toTimeString().split(' ')[0],
      filename: req.file.originalname,
      format,
      raw_rows: rows.length,
    });

    for (const r of summaryRows) {
      db.rows.push({
        id: nextId(db.rows), upload_id: uploadId,
        contractor: r.contractor, contract_no: r.contract_no,
        max_poles: r.max_poles, actual_poles: r.actual_poles,
        iot_available: r.iot_available, api_connected: r.api_connected,
      });
    }

    // Save raw pole data (Format B) for pole-level export
    if (format === 'B') {
      const clean = v => String(v || '').replace(/^["']+|["']+$/g, '').trim();
      await writeRaw(uploadId, rows.map(r => ({
        pole_id:     clean(r[Object.keys(r)[0]]),
        contract_no: clean(r[mapping.contract_no]),
        contractor:  clean(r[mapping.contractor]),
        lamp_code:   String(r['lamp_code'] || r[' lamp_code'] || ''),
        node_id:     String(r['node_id']   || r[' node_id']   || ''),
        iot:         clean(r[mapping.iot_control]),
        max_poles:   mapping.max_poles ? (parseInt(r[mapping.max_poles]) || 0) : 0,
        api:         (r[mapping.api_status] === 1 || r[mapping.api_status] === '1') ? 1 : 0,
      })));
    }

    await writeDB(db);
    res.json({ success: true, uploadId, rowCount: summaryRows.length, rawRows: rows.length, format });
  } catch (e) { res.status(500).json({ error: String(e) }); }
});

// ── GET /api/data/:uploadId ──────────────────────────────────────────────────
app.get('/api/data/:uploadId', async (req, res) => {
  try {
    const db = await readDB();
    res.json(db.rows.filter(r => r.upload_id == req.params.uploadId));
  } catch (e) { res.status(500).json({ error: String(e) }); }
});

// ── DELETE /api/uploads/:uploadId ────────────────────────────────────────────
app.delete('/api/uploads/:uploadId', async (req, res) => {
  try {
    const id = parseInt(req.params.uploadId);
    const db = await readDB();
    db.uploads = db.uploads.filter(u => u.id !== id);
    db.rows    = db.rows.filter(r => r.upload_id !== id);
    await writeDB(db);
    // Delete raw files from Drive + local cache
    await drive.deleteFile(`raw_${id}.json`);
    try { fs.unlinkSync(_cachePath(id)); } catch {}
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: String(e) }); }
});

// ── GET /api/contractors/:uploadId ───────────────────────────────────────────
app.get('/api/contractors/:uploadId', async (req, res) => {
  try {
    const raw = await readRaw(req.params.uploadId);
    if (!raw.length) return res.json([]);
    const tree = {};
    for (const r of raw) {
      if (!tree[r.contractor]) tree[r.contractor] = new Set();
      tree[r.contractor].add(r.contract_no);
    }
    res.json(Object.entries(tree).map(([c, s]) => ({ contractor: c, contracts: [...s].sort() })));
  } catch (e) { res.status(500).json({ error: String(e) }); }
});

// ── GET /api/export/:uploadId ─────────────────────────────────────────────────
app.get('/api/export/:uploadId', async (req, res) => {
  try {
    const { groupBy } = req.query;
    const db   = await readDB();
    const data = db.rows.filter(r => r.upload_id == req.params.uploadId);

    const toRow = r => ({
      'ผู้รับเหมา':    r.contractor,  'เลขที่สัญญา': r.contract_no,
      'max โคม':      r.max_poles,   'จำนวนจริง':   r.actual_poles,
      'ส่วนต่าง':     r.actual_poles - r.max_poles,
      'IOT ได้':      r.iot_available,'API เชื่อม':  r.api_connected,
      'ยังไม่เชื่อม': r.iot_available - r.api_connected,
    });

    const wb = XLSX.utils.book_new();
    if (groupBy === 'contractor') {
      const grouped = {};
      for (const r of data) {
        const k = r.contractor || 'ไม่ระบุ';
        if (!grouped[k]) grouped[k] = [];
        grouped[k].push(toRow(r));
      }
      for (const [name, rows] of Object.entries(grouped))
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), safeSheet(name));
    } else {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.map(toRow)), 'รายเลขที่สัญญา');
    }

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', `attachment; filename="export_${Date.now()}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (e) { res.status(500).json({ error: String(e) }); }
});

// ── GET /api/export/:uploadId/poles ──────────────────────────────────────────
app.get('/api/export/:uploadId/poles', async (req, res) => {
  try {
    const { contractor, contract } = req.query;
    let raw = await readRaw(req.params.uploadId);
    if (!raw.length) return res.status(404).json({ error: 'ไม่พบข้อมูลรายโคม (กรุณาอัพโหลดไฟล์ใหม่)' });

    if (contractor) raw = raw.filter(r => r.contractor === contractor);
    if (contract)   raw = raw.filter(r => r.contract_no === contract);

    const toRow = r => ({
      'รหัสเสาไฟ':           r.pole_id,      'เลขที่สัญญา': r.contract_no,
      'ผู้รับเหมา':           r.contractor,   'lamp_code':   r.lamp_code,
      'node_id':              r.node_id,      'ควบคุม IOT':  r.iot,
      'สถานะ API':            r.api === 1 ? 'เชื่อมแล้ว' : 'ยังไม่เชื่อม',
      'จำนวนโคม (max สัญญา)': r.max_poles,
    });

    const wb = XLSX.utils.book_new();
    if (contract) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(raw.map(toRow)), safeSheet(contract));
    } else {
      const grouped = {};
      for (const r of raw) {
        const k = r.contract_no || (contractor ? r.contract_no : r.contractor) || 'ไม่ระบุ';
        if (!grouped[k]) grouped[k] = [];
        grouped[k].push(toRow(r));
      }
      for (const [name, rows] of Object.entries(grouped))
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), safeSheet(name));
    }

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', `attachment; filename="poles_${Date.now()}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (e) { res.status(500).json({ error: String(e) }); }
});

// ── GET /api/template ────────────────────────────────────────────────────────
app.get('/api/template', (req, res) => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
    { 'ผู้รับเหมา': 'บริษัท ตัวอย่าง จำกัด', 'เลขที่สัญญา': 'สกบ.1/2568', 'max โคม': 100, 'จำนวนจริง': 100, 'IOT ได้': 80, 'API เชื่อม': 75 },
    { 'ผู้รับเหมา': 'บริษัท ABC จำกัด',       'เลขที่สัญญา': 'สกบ.2/2568', 'max โคม': 50,  'จำนวนจริง': 48,  'IOT ได้': 0,  'API เชื่อม': 0  },
    { 'ผู้รับเหมา': 'บริษัท XYZ จำกัด',       'เลขที่สัญญา': 'สนย.3/2568', 'max โคม': 200, 'จำนวนจริง': 215, 'IOT ได้': 150,'API เชื่อม': 0  },
  ]), 'Template');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename="template_streetlight.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`✅ Server running → http://localhost:${PORT}`);
  console.log(`📁 Drive folder  → https://drive.google.com/drive/folders/${  '1mmDo3SHRZGcClBSq5HNf_0MLM19-3Ywe'}`);
});
