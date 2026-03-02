/**
 * ============================================
 *  DataMorph — Universal PDF to Excel Converter
 *  Server v6 — Works with ANY PDF
 * ============================================
 *  Pipeline:
 *   1. Try text extraction (pdfjs-dist / pdf-parse)
 *   2. If empty → OCR (render → Tesseract.js)
 *   3. Smart table detection + parsing
 *   4. Generate styled Excel workbook
 * ============================================
 */

const express = require('express');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const { createCanvas } = require('canvas');
const Tesseract = require('tesseract.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// ── Directories ────────────────────────────────────────────────
const uploadsDir = path.join(__dirname, 'uploads');
const outputsDir = path.join(__dirname, 'outputs');
[uploadsDir, outputsDir].forEach((d) => {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

// ── Multer config ──────────────────────────────────────────────
const storage = multer.diskStorage({
  destination: (_, __, cb) => cb(null, uploadsDir),
  filename: (_, file, cb) => cb(null, `${Date.now()}-${file.originalname}`),
});
const upload = multer({
  storage,
  fileFilter: (_, file, cb) => {
    const ok = file.mimetype === 'application/pdf';
    cb(ok ? null : new Error('Only PDF files are allowed'), ok);
  },
  limits: { fileSize: 25 * 1024 * 1024 },
});

app.use(express.static(path.join(__dirname, 'public')));

function safeDelete(fp) {
  try { if (fp && fs.existsSync(fp)) fs.unlinkSync(fp); } catch {}
}

// ═══════════════════════════════════════════════════════════════
//  LAYER 1: Position-aware text extraction (pdfjs-dist)
// ═══════════════════════════════════════════════════════════════

async function extractTextWithPositions(pdfBuffer) {
  const data = new Uint8Array(pdfBuffer);
  const doc = await pdfjsLib.getDocument({ data }).promise;
  const allPageLines = [];

  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const tc = await page.getTextContent();
    if (!tc.items || tc.items.length === 0) continue;

    // Group items by Y coordinate to reconstruct rows
    const lineMap = new Map();
    for (const item of tc.items) {
      if (!item.str || !item.str.trim()) continue;
      // Round Y to group items on the same line (within 3px)
      const y = Math.round(item.transform[5] / 3) * 3;
      if (!lineMap.has(y)) lineMap.set(y, []);
      lineMap.get(y).push({
        x: item.transform[4],
        text: item.str,
        width: item.width || item.str.length * 5,
      });
    }

    // Sort by Y descending (PDF y-axis is bottom-up)
    const sortedYs = [...lineMap.keys()].sort((a, b) => b - a);
    for (const y of sortedYs) {
      const items = lineMap.get(y).sort((a, b) => a.x - b.x);
      // Reconstruct line: insert tabs for large gaps
      let line = '';
      let lastEnd = 0;
      for (const item of items) {
        const gap = item.x - lastEnd;
        if (line && gap > 20) line += '\t';
        else if (line && gap > 6) line += ' ';
        line += item.text;
        lastEnd = item.x + item.width;
      }
      if (line.trim()) allPageLines.push(line.trim());
    }
  }

  return allPageLines.join('\n');
}

// ═══════════════════════════════════════════════════════════════
//  LAYER 1B: pdf-parse fallback
// ═══════════════════════════════════════════════════════════════

async function extractTextFallback(pdfBuffer) {
  const result = await pdfParse(pdfBuffer);
  return result.text || '';
}

// ═══════════════════════════════════════════════════════════════
//  LAYER 2: OCR Pipeline (render to PNG → Tesseract)
// ═══════════════════════════════════════════════════════════════

class NodeCanvasFactory {
  create(w, h) {
    const canvas = createCanvas(w, h);
    return { canvas, context: canvas.getContext('2d') };
  }
  reset(cd, w, h) { cd.canvas.width = w; cd.canvas.height = h; }
  destroy() {}
}

async function ocrExtract(pdfBuffer) {
  const data = new Uint8Array(pdfBuffer);
  const doc = await pdfjsLib.getDocument({ data }).promise;
  const numPages = doc.numPages;
  console.log(`  [OCR] Processing ${numPages} pages...`);

  // Use single worker to save memory on free-tier containers
  const worker = await Tesseract.createWorker('eng');

  const allText = [];
  // Scale 2.0 is a good balance of accuracy vs memory
  const scale = 2.0;

  for (let p = 1; p <= numPages; p++) {
    try {
      const page = await doc.getPage(p);
      const vp = page.getViewport({ scale });
      const factory = new NodeCanvasFactory();
      const { canvas, context } = factory.create(vp.width, vp.height);

      // White background is crucial for OCR accuracy
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, vp.width, vp.height);

      await page.render({
        canvasContext: context,
        viewport: vp,
        canvasFactory: factory,
      }).promise;

      const pngBuffer = canvas.toBuffer('image/png');
      console.log(`  [OCR] Page ${p}/${numPages}: rendered (${(pngBuffer.length / 1024).toFixed(0)} KB)`);

      const { data: ocrData } = await worker.recognize(pngBuffer);
      allText.push(ocrData.text);
      console.log(`  [OCR] Page ${p}/${numPages}: done (conf: ${ocrData.confidence}%)`);
    } catch (e) {
      console.log(`  [OCR] Page ${p} failed: ${e.message}`);
    }
  }

  await worker.terminate();
  return allText.join('\n');
}

// ═══════════════════════════════════════════════════════════════
//  MASTER TEXT EXTRACTOR
// ═══════════════════════════════════════════════════════════════

async function extractAllText(pdfBuffer) {
  // Method 1: pdfjs-dist position-aware
  try {
    const text = await extractTextWithPositions(pdfBuffer);
    if (text.trim().length > 30) {
      console.log(`  [Extract] pdfjs: ${text.length} chars`);
      return { text, method: 'pdfjs' };
    }
  } catch (e) {
    console.log(`  [Extract] pdfjs error: ${e.message}`);
  }

  // Method 2: pdf-parse fallback
  try {
    const text = await extractTextFallback(pdfBuffer);
    if (text.trim().length > 30) {
      console.log(`  [Extract] pdf-parse: ${text.length} chars`);
      return { text, method: 'pdf-parse' };
    }
  } catch (e) {
    console.log(`  [Extract] pdf-parse error: ${e.message}`);
  }

  // Method 3: OCR
  console.log('  [Extract] No text found — starting OCR (this may take 1-3 minutes)...');
  try {
    const text = await ocrExtract(pdfBuffer);
    if (text.trim().length > 10) {
      console.log(`  [Extract] OCR: ${text.length} chars`);
      return { text, method: 'ocr' };
    }
  } catch (e) {
    console.log(`  [Extract] OCR error: ${e.message}`);
  }

  return { text: '', method: 'none' };
}

// ═══════════════════════════════════════════════════════════════
//  UNIVERSAL TABLE PARSER
//  Tries multiple strategies, returns first one that works
// ═══════════════════════════════════════════════════════════════

function parseToTable(rawText) {
  const strategies = [
    { name: 'Tab-delimited', fn: parseTabSeparated },
    { name: 'Reg number anchored', fn: parseByRegNumbers },
    { name: 'Header keyword', fn: parseByHeaders },
    { name: 'Structured lines', fn: parseStructuredLines },
    { name: 'Number-heavy lines', fn: parseNumberLines },
  ];

  for (const { name, fn } of strategies) {
    try {
      const result = fn(rawText);
      if (result && result.rows.length >= 1) {
        console.log(`  [Parse] "${name}" → ${result.rows.length} rows x ${result.headers.length} cols`);
        return result;
      }
    } catch (e) {
      console.log(`  [Parse] "${name}" error: ${e.message}`);
    }
  }

  // Absolute fallback: each line as a row
  const lines = rawText.split(/\r?\n/).map((l) => l.trim()).filter((l) => l.length > 1);
  return {
    headers: ['Line #', 'Content'],
    rows: lines.map((l, i) => [String(i + 1), l]),
  };
}

// ── Strategy 1: Tab-separated (from position-aware extraction) ──

function parseTabSeparated(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.includes('\t') && l.trim());
  if (lines.length < 2) return null;

  const parsed = lines.map((l) => l.split('\t').map((s) => s.trim()).filter(Boolean));
  const colCounts = parsed.map((r) => r.length);
  const mc = mode(colCounts);
  const goodRows = parsed.filter((r) => Math.abs(r.length - mc) <= 2 && r.length >= 2);

  if (goodRows.length < 2) return null;

  // Check if first row is a header
  const first = goodRows[0];
  const isHeader = first.every((cell) => !/^\d+(\.\d+)?$/.test(cell));

  let headers, rows;
  if (isHeader) {
    headers = first;
    rows = goodRows.slice(1);
  } else {
    headers = Array.from({ length: mc }, (_, i) => `Column ${i + 1}`);
    rows = goodRows;
  }

  normalizeTable(headers, rows);
  return { headers, rows };
}

// ── Strategy 2: Registration-number anchored (SRM / university) ──

function normalizeRegNo(raw) {
  let s = raw.toUpperCase().replace(/[\[\](){}|,\s]/g, '');
  const m = s.match(/003010(\d{2,4})/);
  if (m) return 'RA2411003010' + m[1].padStart(3, '0');
  return s;
}

function parseByRegNumbers(text) {
  // SRM-style: anything with 003010 core
  const srmPattern = /[A-Za-z0-9]{1,10}003010\d{2,4}/g;
  // Generic university reg numbers: 2-4 letters + 8-15 digits
  const genericPattern = /\b([A-Z]{2,4}\d{8,15})\b/gi;

  const seen = new Set();
  const matches = [];

  for (const m of text.matchAll(srmPattern)) {
    const n = normalizeRegNo(m[0]);
    if (!seen.has(n)) { seen.add(n); matches.push({ regNo: n, raw: m[0], idx: m.index }); }
  }
  for (const m of text.matchAll(genericPattern)) {
    const n = m[1].toUpperCase();
    if (!seen.has(n) && n.length >= 10) {
      seen.add(n);
      matches.push({ regNo: n, raw: m[0], idx: m.index });
    }
  }

  matches.sort((a, b) => a.idx - b.idx);
  if (matches.length < 2) return null;

  // Detect subject/course codes (21MAB204T(A), 21CSC204J(B), etc.)
  const coursePattern = /\b(2[01][A-Z]{2,5}\d{2,4}[A-Z]*(?:\([A-Z]\))?)\b/gi;
  const courses = [];
  const seenCourses = new Set();
  for (const m of text.matchAll(coursePattern)) {
    const c = m[1].toUpperCase();
    if (!seenCourses.has(c) && !/^RA\d/.test(c)) {
      seenCourses.add(c);
      courses.push(c);
    }
  }

  const rows = [];
  for (let i = 0; i < matches.length; i++) {
    const { regNo, raw, idx } = matches[i];
    const end = i + 1 < matches.length ? matches[i + 1].idx : text.length;
    const chunk = text.substring(idx + raw.length, end);

    // Extract name: text before first number/course code
    const numStart = chunk.search(/\b\d{2,3}\.\d/);
    const codeStart = chunk.search(/\b2[01][A-Z]{2,5}\d/i);
    let nameEnd = Math.min(200, chunk.length);
    if (numStart > 0) nameEnd = Math.min(nameEnd, numStart);
    if (codeStart > 0) nameEnd = Math.min(nameEnd, codeStart);

    let name = chunk.substring(0, nameEnd)
      .replace(/[^a-zA-Z\s.'\-]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    name = name.split(/\s+/).filter((w) => w.length > 1).join(' ');
    if (!name || name.length < 2) name = 'Unknown';
    if (name.length > 60) name = name.substring(0, 60).trim();

    // Extract decimal scores
    const scores = [...chunk.matchAll(/\b(\d{1,3}\.\d{1,2})\b/g)].map((m) => m[1]);
    rows.push([String(rows.length + 1), regNo, name, ...scores]);
  }

  if (rows.length === 0) return null;

  const headers = ['S.No', 'Reg. No', 'Name'];
  courses.slice(0, 15).forEach((c) => headers.push(c));
  normalizeTable(headers, rows);
  return { headers, rows };
}

// ── Strategy 3: Header keyword detection ─────────────────────

function parseByHeaders(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const headerKW = /\b(s\.?no|sr\.?no|roll|reg\.?no|name|student|attendance|marks?|score|grade|subject|total|percent|date|class|section|department|semester|course)\b/i;

  let headerIdx = -1;
  for (let i = 0; i < Math.min(lines.length, 40); i++) {
    const kwMatches = lines[i].match(new RegExp(headerKW.source, 'gi'));
    if (kwMatches && kwMatches.length >= 2) {
      headerIdx = i;
      break;
    }
  }

  if (headerIdx === -1) return null;

  const headers = tokenize(lines[headerIdx]);
  if (headers.length < 2) return null;

  const rows = [];
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const tokens = tokenize(lines[i]);
    if (tokens.length >= 2 && tokens.some((t) => /\d/.test(t))) {
      rows.push(tokens);
    }
  }

  if (rows.length === 0) return null;
  normalizeTable(headers, rows);
  return { headers, rows };
}

// ── Strategy 4: Lines with consistent structure ──────────────

function parseStructuredLines(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter((l) => l.length > 3);
  const tokenized = lines.map((l) => tokenize(l)).filter((t) => t.length >= 2);

  if (tokenized.length < 2) return null;

  const colCounts = tokenized.map((r) => r.length);
  const mc = mode(colCounts);
  const goodRows = tokenized.filter((r) => Math.abs(r.length - mc) <= 2 && r.length >= 2);

  if (goodRows.length < 2) return null;

  const first = goodRows[0];
  const isHeader = first.every((cell) => !/^\d+(\.\d+)?$/.test(cell));

  let headers, rows;
  if (isHeader && goodRows.length > 1) {
    headers = first;
    rows = goodRows.slice(1);
  } else {
    headers = Array.from({ length: mc }, (_, i) => `Column ${i + 1}`);
    rows = goodRows;
  }

  normalizeTable(headers, rows);
  return { headers, rows };
}

// ── Strategy 5: Lines with multiple numbers ──────────────────

function parseNumberLines(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const rows = [];
  for (const line of lines) {
    const nums = line.match(/\d+\.?\d*/g);
    if (nums && nums.length >= 2) {
      const tokens = tokenize(line);
      if (tokens.length >= 2) rows.push(tokens);
    }
  }
  if (rows.length === 0) return null;

  const mc = Math.max(...rows.map((r) => r.length));
  const first = rows[0];
  const isHeader = first.every((cell) => !/^\d+(\.\d+)?$/.test(cell));

  let headers, dataRows;
  if (isHeader && rows.length > 1) {
    headers = first;
    dataRows = rows.slice(1);
  } else {
    headers = Array.from({ length: mc }, (_, i) => `Column ${i + 1}`);
    dataRows = rows;
  }

  normalizeTable(headers, dataRows);
  return { headers, rows: dataRows };
}

// ═══════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════

function tokenize(line) {
  // Try tab-separated first
  if (line.includes('\t')) {
    const t = line.split('\t').map((s) => s.trim()).filter(Boolean);
    if (t.length >= 2) return t;
  }
  // Try pipe-separated
  if (line.includes('|')) {
    const t = line.split('|').map((s) => s.trim()).filter(Boolean);
    if (t.length >= 2) return t;
  }
  // Try multiple spaces
  const sp = line.split(/\s{2,}/).map((s) => s.trim()).filter(Boolean);
  if (sp.length >= 2) return sp;
  // Try comma-separated
  if (line.includes(',')) {
    const t = line.split(',').map((s) => s.trim()).filter(Boolean);
    if (t.length >= 2) return t;
  }
  // Fall back to single-space split
  return line.split(/\s+/).map((s) => s.trim()).filter(Boolean);
}

function normalizeTable(headers, rows) {
  const maxCols = Math.max(headers.length, ...rows.map((r) => r.length));
  while (headers.length < maxCols) headers.push(`Col ${headers.length + 1}`);
  rows.forEach((r) => {
    while (r.length < maxCols) r.push('');
    r.length = maxCols;
  });
}

function mode(arr) {
  const freq = {};
  arr.forEach((v) => (freq[v] = (freq[v] || 0) + 1));
  return Number(Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0]);
}

// ═══════════════════════════════════════════════════════════════
//  EXCEL GENERATION
// ═══════════════════════════════════════════════════════════════

function buildExcel(headers, rows) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'DataMorph';
  wb.created = new Date();
  const ws = wb.addWorksheet('Data');

  // Dynamic column widths based on content
  ws.columns = headers.map((h, i) => {
    const maxContentLen = Math.max(
      (h || '').length,
      ...rows.slice(0, 20).map((r) => (r[i] || '').toString().length)
    );
    return {
      header: h,
      key: `c${i}`,
      width: Math.max(10, Math.min(40, maxContentLen + 3)),
    };
  });

  // Header row styling
  const hr = ws.getRow(1);
  hr.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
  hr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
  hr.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  hr.height = 30;

  // Data rows
  rows.forEach((rowData, idx) => {
    const obj = {};
    headers.forEach((_, i) => {
      obj[`c${i}`] = rowData[i] !== undefined ? rowData[i] : '';
    });
    const row = ws.addRow(obj);
    row.alignment = { vertical: 'middle', wrapText: true };
    if (idx % 2 === 0) {
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0FF' } };
    }
    // Auto-convert numeric strings to numbers
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = String(cell.value).trim();
      if (/^\d+(\.\d+)?$/.test(v)) {
        cell.value = parseFloat(v);
        cell.numFmt = v.includes('.') ? '0.00' : '0';
      }
    });
  });

  // Freeze header and add auto-filter
  const colLetter = getColLetter(headers.length);
  ws.autoFilter = `A1:${colLetter}1`;
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  return wb;
}

function getColLetter(n) {
  let s = '', num = n;
  while (num > 0) { num--; s = String.fromCharCode(65 + (num % 26)) + s; num = Math.floor(num / 26); }
  return s;
}

// ═══════════════════════════════════════════════════════════════
//  UPLOAD ROUTE
// ═══════════════════════════════════════════════════════════════

app.post('/upload', upload.single('pdfFile'), async (req, res) => {
  let uploadedPath = null;
  let outputPath = null;

  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No PDF file uploaded.' });
    }

    uploadedPath = req.file.path;
    const pdfBuffer = fs.readFileSync(uploadedPath);

    console.log(`\n${'='.repeat(55)}`);
    console.log(`  File: ${req.file.originalname} (${(pdfBuffer.length / 1024).toFixed(1)} KB)`);
    console.log(`${'='.repeat(55)}`);

    // ── Extract text ──────────────────────────────────────────
    const { text, method } = await extractAllText(pdfBuffer);

    if (!text || text.trim().length < 5) {
      return res.status(422).json({
        error: 'Could not extract any text from this PDF. The file may be corrupted or contain only images.',
      });
    }

    console.log(`  Extraction method: ${method} | Text: ${text.length} chars`);

    // ── Parse into table ──────────────────────────────────────
    console.log('  Parsing table structure...');
    const { headers, rows } = parseToTable(text);

    console.log(`  Result: ${rows.length} rows x ${headers.length} columns`);
    if (rows[0]) console.log(`  First row: ${rows[0].slice(0, 5).join(' | ')}`);

    // ── Generate Excel ────────────────────────────────────────
    const wb = buildExcel(headers, rows);
    const outName = `data-${Date.now()}.xlsx`;
    outputPath = path.join(outputsDir, outName);
    await wb.xlsx.writeFile(outputPath);

    // Download filename = original PDF name but .xlsx
    const dlName = req.file.originalname.replace(/\.pdf$/i, '') + '.xlsx';

    res.download(outputPath, dlName, (err) => {
      safeDelete(uploadedPath);
      safeDelete(outputPath);
      if (err && !res.headersSent) {
        console.error('  Download error:', err.message);
      }
    });

  } catch (err) {
    console.error('  CONVERSION ERROR:', err);
    safeDelete(uploadedPath);
    safeDelete(outputPath);
    if (!res.headersSent) {
      res.status(500).json({
        error: `Conversion failed: ${err.message}`,
      });
    }
  }
});

// ── Error handler ──────────────────────────────────────────────
app.use((err, _req, res, _next) => {
  if (err instanceof multer.MulterError) {
    return res.status(err.code === 'LIMIT_FILE_SIZE' ? 413 : 400).json({
      error: err.message,
    });
  }
  if (err) {
    return res.status(400).json({ error: err.message || 'Something went wrong.' });
  }
});

// ── Start server ───────────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n  ========================================`);
  console.log(`    DataMorph — PDF to Excel Converter`);
  console.log(`    Running on http://0.0.0.0:${PORT}`);
  console.log(`    Ready for any PDF!`);
  console.log(`  ========================================\n`);
});
