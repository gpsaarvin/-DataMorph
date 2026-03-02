/**
 * ============================================
 *  DataMorphp — Smart PDF to Excel Converter
 *  Server v5 — OCR-Powered Extraction
 * ============================================
 *  Handles Zoho Creator / SRM Type3-font PDFs
 *  that no text-extraction library can decode.
 *
 *  Pipeline:
 *   1. Try text extraction (pdf-parse / pdfjs-dist)
 *   2. If empty → render pages to images → Tesseract OCR
 *   3. Parse OCR text with multi-strategy parser
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
const debugDir = path.join(__dirname, 'debug');
[uploadsDir, outputsDir, debugDir].forEach((dir) => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// ── Multer config ──────────────────────────────────────────────
const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, uploadsDir),
  filename: (_req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`),
});
const fileFilter = (_req, file, cb) => {
  const ok = file.mimetype === 'application/pdf';
  cb(ok ? null : new Error('Only PDF files are allowed'), ok);
};
const upload = multer({ storage, fileFilter, limits: { fileSize: 25 * 1024 * 1024 } });

// ── Static files ───────────────────────────────────────────────
app.use(express.static(path.join(__dirname, 'public')));

function safeDelete(fp) {
  try { if (fs.existsSync(fp)) fs.unlinkSync(fp); } catch {}
}

// ═══════════════════════════════════════════════════════════════
//  EXTRACTION LAYER 1: Standard text extraction
// ═══════════════════════════════════════════════════════════════

async function tryTextExtraction(pdfBuffer) {
  let text = '';

  // Method A — pdfjs-dist getTextContent
  try {
    const data = new Uint8Array(pdfBuffer);
    const doc = await pdfjsLib.getDocument({ data }).promise;
    const parts = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const tc = await page.getTextContent();
      const pageText = tc.items
        .filter((it) => it.str && it.str.trim())
        .map((it) => it.str)
        .join(' ');
      if (pageText.trim()) parts.push(pageText);
    }
    text = parts.join('\n');
    if (text.trim().length > 50) {
      console.log(`   pdfjs-dist extracted ${text.length} chars`);
      return text;
    }
  } catch (e) {
    console.log(`   pdfjs-dist error: ${e.message}`);
  }

  // Method B — pdf-parse
  try {
    const result = await pdfParse(pdfBuffer);
    text = result.text || '';
    if (text.trim().length > 50) {
      console.log(`   pdf-parse extracted ${text.length} chars`);
      return text;
    }
  } catch (e) {
    console.log(`   pdf-parse error: ${e.message}`);
  }

  console.log(`   Text extraction: no usable text found`);
  return '';
}

// ═══════════════════════════════════════════════════════════════
//  EXTRACTION LAYER 2: OCR Pipeline
//  Render PDF pages → PNG images → Tesseract OCR
// ═══════════════════════════════════════════════════════════════

class NodeCanvasFactory {
  create(w, h) {
    const canvas = createCanvas(w, h);
    return { canvas, context: canvas.getContext('2d') };
  }
  reset(canvasAndContext, w, h) {
    canvasAndContext.canvas.width = w;
    canvasAndContext.canvas.height = h;
  }
  destroy() {}
}

async function ocrExtractText(pdfBuffer, progressCb) {
  const data = new Uint8Array(pdfBuffer);
  const doc = await pdfjsLib.getDocument({ data }).promise;
  const numPages = doc.numPages;
  console.log(`   OCR: Rendering ${numPages} pages...`);

  // Create a Tesseract scheduler for parallel OCR
  const scheduler = Tesseract.createScheduler();
  const workerCount = Math.min(numPages, 3); // Use up to 3 workers

  for (let i = 0; i < workerCount; i++) {
    const worker = await Tesseract.createWorker('eng');
    scheduler.addWorker(worker);
  }

  const pageTexts = [];
  const renderScale = 3.0; // Higher = better OCR accuracy

  for (let p = 1; p <= numPages; p++) {
    const pct = Math.round((p / numPages) * 80);
    if (progressCb) progressCb(pct);

    const page = await doc.getPage(p);
    const vp = page.getViewport({ scale: renderScale });
    const factory = new NodeCanvasFactory();
    const canvasCtx = factory.create(vp.width, vp.height);

    // White background (important for OCR)
    canvasCtx.context.fillStyle = '#ffffff';
    canvasCtx.context.fillRect(0, 0, vp.width, vp.height);

    await page.render({
      canvasContext: canvasCtx.context,
      viewport: vp,
      canvasFactory: factory,
    }).promise;

    const pngBuffer = canvasCtx.canvas.toBuffer('image/png');
    console.log(`   Page ${p}/${numPages}: rendered (${(pngBuffer.length / 1024).toFixed(0)} KB)`);

    // Save first page for debug
    if (p === 1) {
      try {
        fs.writeFileSync(path.join(debugDir, 'page1-ocr.png'), pngBuffer);
      } catch {}
    }

    const { data: ocrData } = await scheduler.addJob('recognize', pngBuffer);
    pageTexts.push(ocrData.text);
    console.log(`   Page ${p}/${numPages}: OCR done (confidence: ${ocrData.confidence}%)`);
  }

  await scheduler.terminate();

  const fullText = pageTexts.join('\n');
  console.log(`   OCR total: ${fullText.length} chars across ${numPages} pages`);

  // Save OCR text for debugging
  try {
    fs.writeFileSync(path.join(debugDir, `ocr-text-${Date.now()}.txt`), fullText, 'utf8');
  } catch {}

  return fullText;
}

// ═══════════════════════════════════════════════════════════════
//  TEXT PARSING — Multi-strategy parser for student data
// ═══════════════════════════════════════════════════════════════

function parseStudentData(rawText) {
  const lines = rawText.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  let result;

  // Strategy 1: Registration number anchored (SRM-style)
  result = parseByRegNumbers(rawText);
  if (result.rows.length >= 2) {
    console.log(`   [RegNo Strategy] ${result.rows.length} rows`);
    return result;
  }

  // Strategy 2: OCR table lines — look for lines with RA numbers + scores
  result = parseOcrTableLines(lines);
  if (result.rows.length >= 2) {
    console.log(`   [OCR Table Strategy] ${result.rows.length} rows`);
    return result;
  }

  // Strategy 3: Header-keyword detection
  result = parseByHeaderDetection(lines);
  if (result.rows.length >= 2) {
    console.log(`   [Header Strategy] ${result.rows.length} rows`);
    return result;
  }

  // Strategy 4: Lines with multiple numbers
  result = parseNumberLines(lines);
  if (result.rows.length >= 2) {
    console.log(`   [Number Lines Strategy] ${result.rows.length} rows`);
    return result;
  }

  // Strategy 5: Brute force — any multi-token lines
  result = parseBruteForce(lines);
  if (result.rows.length >= 1) {
    console.log(`   [Brute Force Strategy] ${result.rows.length} rows`);
    return result;
  }

  // Last resort: dump all lines
  return {
    headers: ['Extracted Text'],
    rows: lines.filter((l) => l.length > 2).map((l) => [l]),
  };
}

// ── Strategy 1: Registration Number Anchored ─────────────────

/** Normalize an OCR'd registration number — fix common misreads */
function normalizeRegNo(raw) {
  let s = raw.toUpperCase()
    .replace(/[\[\](){}|,]/g, '')  // strip OCR brackets
    .replace(/\s/g, '');           // strip spaces
  // Common OCR substitutions: T↔1, O↔0, I↔1, S↔5, B↔8
  // Try to reconstruct RA2411003010XXX pattern for SRM numbers
  if (/003010\d{2,4}/.test(s)) {
    // Found the SRM core — reconstruct
    const coreMatch = s.match(/003010(\d{2,4})/);
    if (coreMatch) {
      const studentId = coreMatch[1].padStart(3, '0'); // ensure 3 digits
      // The full SRM reg number is always RA2411003010XXX
      return 'RA2411003010' + studentId;
    }
  }
  // For non-SRM: just clean up
  return s;
}

function parseByRegNumbers(text) {
  // Very broad pattern to catch OCR-garbled registration numbers
  // Catches: RA2411003010324, Ra2411003010340, RA24T1003010338,
  //          n2a11003010334, RAZATIO030I0S39, a2a1003010365, etc.
  // Strategy: look for anything containing the SRM core pattern 003010
  // plus a broader alpha-digit pattern for other universities
  const srmPattern = /[A-Za-z0-9]{1,10}003010\d{2,4}/g;
  const genericPattern = /\b([A-Z]{2,4}\d{8,15})\b/gi;

  // Collect all SRM-style matches
  const srmMatches = [...text.matchAll(srmPattern)].map(m => ({
    raw: m[0], index: m.index, regNo: normalizeRegNo(m[0])
  }));

  // Collect generic matches
  const genericMatches = [...text.matchAll(genericPattern)].map(m => ({
    raw: m[0], index: m.index, regNo: m[1].toUpperCase()
  }));

  // Merge, preferring SRM matches, dedup by normalized regNo
  const allCandidates = [...srmMatches, ...genericMatches];
  const seen = new Set();
  const regMatches = [];
  for (const c of allCandidates) {
    if (!seen.has(c.regNo)) {
      seen.add(c.regNo);
      regMatches.push(c);
    }
  }

  // Sort by position in text
  regMatches.sort((a, b) => a.index - b.index);

  console.log(`   Found ${regMatches.length} unique reg numbers (SRM: ${srmMatches.length}, generic: ${genericMatches.length})`);
  if (regMatches.length < 2) return { headers: [], rows: [] };

  // Detect subject/course code headers from text
  // Pattern: 21MAB204T(A), 21CSC204J(B), etc.
  const coursePattern = /\b(2[01][A-Z]{2,5}\d{2,4}[A-Z]*(?:\([A-Z]\))?)\b/gi;
  const seenCourses = new Set();
  const orderedCourses = [];
  for (const m of text.matchAll(coursePattern)) {
    const code = m[1].toUpperCase();
    // Skip if looks like a reg number
    if (/^RA\d{7,}/.test(code)) continue;
    if (!seenCourses.has(code)) {
      seenCourses.add(code);
      orderedCourses.push(code);
    }
  }

  // Limit to reasonable number of subject columns (likely <= 15)
  const subjectHeaders = orderedCourses.slice(0, 15);
  console.log(`   Detected ${subjectHeaders.length} subject codes`);

  // Extract per-student data
  const rows = [];
  for (let i = 0; i < regMatches.length; i++) {
    const { regNo, raw, index: start } = regMatches[i];
    const end = i + 1 < regMatches.length ? regMatches[i + 1].index : text.length;
    const chunk = text.substring(start, end);

    // Name: text after the raw match, before scores/course codes
    const afterReg = chunk.substring(raw.length).trim();
    // Find where numbers start (scores like 100.00, 84.62)
    const numStart = afterReg.search(/\b\d{2,3}\.\d/);
    // Find where course codes appear (21MAB204T, etc)
    const codeStart = afterReg.search(/\b2[01][A-Z]{2,5}\d{2,4}/i);
    let nameEnd = Math.min(200, afterReg.length); // cap at 200 chars
    if (numStart > 0) nameEnd = Math.min(nameEnd, numStart);
    if (codeStart > 0) nameEnd = Math.min(nameEnd, codeStart);

    let name = afterReg.substring(0, nameEnd)
      .replace(/[\[\]|{}()\d]/g, ' ')     // Remove OCR artifacts
      .replace(/[^a-zA-Z\s.'\-]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    // Remove single-char OCR noise words
    name = name.split(/\s+/).filter(w => w.length > 1 || /^[A-Z]$/.test(w)).join(' ');

    // Clean up
    if (!name || name.length < 2) name = 'Unknown';
    if (name.length > 50) name = name.substring(0, 50).trim();

    // Extract all decimal scores (like 100.00, 77.78, 66.67)
    // But also handle OCR errors like "10000" → should be "100.00"
    const scorePattern = /\b(\d{1,3}\.\d{1,2})\b/g;
    const scores = [...chunk.matchAll(scorePattern)].map((m) => m[1]);

    const row = [String(rows.length + 1), regNo, name, ...scores];
    rows.push(row);
  }

  if (rows.length === 0) return { headers: [], rows: [] };

  // Build headers
  const headers = ['S.No', 'Reg. No', 'Name'];
  if (subjectHeaders.length > 0) {
    subjectHeaders.forEach((c) => headers.push(c));
  }

  // Normalise column count
  const maxCols = Math.max(headers.length, ...rows.map((r) => r.length));
  while (headers.length < maxCols) headers.push(`Score ${headers.length - 2}`);
  rows.forEach((r) => { while (r.length < maxCols) r.push(''); });
  // Trim excess columns from rows
  rows.forEach((r) => { if (r.length > maxCols) r.length = maxCols; });

  return { headers, rows };
}

// ── Strategy 2: OCR Table Lines ──────────────────────────────
// OCR output often has each student as one messy line
function parseOcrTableLines(lines) {
  // Find lines containing both a reg number and decimal scores
  const regPattern = /[A-Za-z0-9]{1,10}003010\d{2,4}|(?:RA|AP|SRM)\d{7,15}/i;
  const dataLines = [];

  for (const line of lines) {
    if (regPattern.test(line)) {
      const scores = line.match(/\b\d{1,3}\.\d{1,2}\b/g);
      if (scores && scores.length >= 1) {
        dataLines.push(line);
      }
    }
  }

  if (dataLines.length < 2) return { headers: [], rows: [] };

  // Parse each data line
  const rows = [];
  for (const line of dataLines) {
    const regMatch = line.match(/[A-Za-z0-9]{1,10}003010\d{2,4}|(?:RA|AP|SRM)\d{7,15}/i);
    if (!regMatch) continue;
    const regNo = normalizeRegNo(regMatch[0]);

    // Text after reg number, before first score
    const afterReg = line.substring(line.indexOf(regMatch[0]) + regMatch[0].length);
    const firstScoreIdx = afterReg.search(/\b\d{1,3}\.\d/);
    let name = firstScoreIdx > 0
      ? afterReg.substring(0, firstScoreIdx).replace(/[\[\]|{}\d]/g, ' ').replace(/[^a-zA-Z\s.'\-]/g, ' ').replace(/\s+/g, ' ').trim()
      : 'Unknown';
    if (name.length < 2) name = 'Unknown';

    const scores = [...afterReg.matchAll(/\b(\d{1,3}\.\d{1,2})\b/g)].map((m) => m[1]);
    rows.push([String(rows.length + 1), regNo, name, ...scores]);
  }

  if (rows.length === 0) return { headers: [], rows: [] };

  const maxCols = Math.max(...rows.map((r) => r.length));
  const headers = ['S.No', 'Reg. No', 'Name'];
  while (headers.length < maxCols) headers.push(`Subject ${headers.length - 2}`);
  rows.forEach((r) => { while (r.length < maxCols) r.push(''); });

  return { headers, rows };
}

// ── Strategy 3: Header detection ─────────────────────────────
function parseByHeaderDetection(lines) {
  const headerKW = /\b(s\.?no|sr\.?no|roll|reg|name|student|attendance|marks?|score|grade|subject|total|percentage)\b/i;
  let headerIdx = -1, headerTokens = [];

  for (let i = 0; i < Math.min(lines.length, 25); i++) {
    if (headerKW.test(lines[i])) {
      const matches = lines[i].match(new RegExp(headerKW.source, 'gi'));
      if (matches && matches.length >= 2) {
        headerIdx = i;
        headerTokens = tokeniseLine(lines[i]);
        break;
      }
    }
  }

  if (headerIdx === -1 || headerTokens.length < 2) return { headers: [], rows: [] };

  const headers = headerTokens;
  const rows = [];
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const tokens = tokeniseLine(lines[i]);
    if (tokens.length >= 2 && tokens.some((t) => /\d/.test(t))) {
      rows.push(tokens);
    }
  }

  if (rows.length === 0) return { headers: [], rows: [] };
  normalise(headers, rows);
  return { headers, rows };
}

// ── Strategy 4: Number-heavy lines ───────────────────────────
function parseNumberLines(lines) {
  const rows = [];
  for (const line of lines) {
    const nums = line.match(/\d+\.\d+/g);
    if (nums && nums.length >= 2) {
      const tokens = tokeniseLine(line);
      if (tokens.length >= 2) rows.push(tokens);
    }
  }
  if (rows.length === 0) return { headers: [], rows: [] };
  const mc = Math.max(...rows.map((r) => r.length));
  const headers = Array.from({ length: mc }, (_, i) => `Column ${i + 1}`);
  normalise(headers, rows);
  return { headers, rows };
}

// ── Strategy 5: Brute force ──────────────────────────────────
function parseBruteForce(lines) {
  const rows = [];
  for (const line of lines) {
    if (line.length < 3) continue;
    const tokens = tokeniseLine(line);
    if (tokens.length >= 2) {
      const hasAlpha = tokens.some((t) => /[a-zA-Z]{2,}/.test(t));
      const hasNum = tokens.some((t) => /\d/.test(t));
      if (hasAlpha || hasNum) rows.push(tokens);
    }
  }
  if (rows.length === 0) return { headers: [], rows: [] };

  const first = rows[0];
  const isHeader = first.every((t) => !/^\d+(\.\d+)?$/.test(t.trim()));
  let headers, dataRows;
  if (isHeader && rows.length > 1) { headers = first; dataRows = rows.slice(1); }
  else {
    const mc = Math.max(...rows.map((r) => r.length));
    headers = Array.from({ length: mc }, (_, i) => `Column ${i + 1}`);
    dataRows = rows;
  }
  normalise(headers, dataRows);
  return { headers, rows: dataRows };
}

// ── Helpers ──────────────────────────────────────────────────
function tokeniseLine(line) {
  if (line.includes('\t')) { const t = line.split('\t').map((s) => s.trim()).filter(Boolean); if (t.length >= 2) return t; }
  if (line.includes('|')) { const t = line.split('|').map((s) => s.trim()).filter(Boolean); if (t.length >= 2) return t; }
  const sp = line.split(/\s{2,}/).map((s) => s.trim()).filter(Boolean);
  if (sp.length >= 2) return sp;
  if (line.includes(',')) { const t = line.split(',').map((s) => s.trim()).filter(Boolean); if (t.length >= 2) return t; }
  return line.split(/\s+/).map((s) => s.trim()).filter(Boolean);
}

function normalise(headers, rows) {
  const maxCols = Math.max(headers.length, ...rows.map((r) => r.length));
  while (headers.length < maxCols) headers.push(`Col ${headers.length + 1}`);
  rows.forEach((r) => { while (r.length < maxCols) r.push(''); });
}

// ═══════════════════════════════════════════════════════════════
//  EXCEL GENERATION
// ═══════════════════════════════════════════════════════════════

function buildExcel(headers, rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'DataMorphp';
  workbook.created = new Date();
  const sheet = workbook.addWorksheet('Student Data');

  sheet.columns = headers.map((h, i) => ({
    header: h,
    key: `col_${i}`,
    width: Math.max(10, Math.min(30, (h || '').length + 4)),
  }));

  // Style header row
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  headerRow.height = 30;

  // Add data
  rows.forEach((rowData, idx) => {
    const obj = {};
    headers.forEach((_, i) => { obj[`col_${i}`] = rowData[i] !== undefined ? rowData[i] : ''; });
    const row = sheet.addRow(obj);
    row.alignment = { vertical: 'middle', wrapText: true };
    if (idx % 2 === 0) {
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0FF' } };
    }
    // Convert numeric-looking strings to numbers
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = String(cell.value).trim();
      if (/^\d+(\.\d+)?$/.test(v)) {
        cell.value = parseFloat(v);
        cell.numFmt = v.includes('.') ? '0.00' : '0';
      }
    });
  });

  // Auto filter + freeze header
  const colLetter = getColLetter(headers.length);
  sheet.autoFilter = `A1:${colLetter}1`;
  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  return workbook;
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
  let uploadedPath = null, outputPath = null;

  try {
    if (!req.file) return res.status(400).json({ error: 'No PDF file uploaded.' });
    uploadedPath = req.file.path;
    const pdfBuffer = fs.readFileSync(uploadedPath);

    console.log(`\n${'='.repeat(60)}`);
    console.log(`  File: ${req.file.originalname} (${(pdfBuffer.length / 1024).toFixed(1)} KB)`);
    console.log(`${'='.repeat(60)}`);

    let rawText = '';

    // ── LAYER 1: Standard text extraction ─────────────────────
    console.log('\n  Layer 1: Standard text extraction...');
    rawText = await tryTextExtraction(pdfBuffer);

    // ── LAYER 2: OCR if text extraction failed ────────────────
    if (rawText.trim().length < 50) {
      console.log('\n  Layer 2: OCR Pipeline (rendering + Tesseract)...');
      console.log('  This may take 30-90 seconds for multi-page PDFs...');
      rawText = await ocrExtractText(pdfBuffer);
    }

    if (!rawText || rawText.trim().length < 10) {
      return res.status(422).json({
        error: 'Could not extract any data from this PDF. The file may be corrupted or contain only images without text.',
      });
    }

    // Save combined text for debugging
    try {
      fs.writeFileSync(path.join(debugDir, `final-text-${Date.now()}.txt`), rawText, 'utf8');
    } catch {}

    // ── PARSE structured data ─────────────────────────────────
    console.log('\n  Parsing structured data...');
    const { headers, rows } = parseStudentData(rawText);

    console.log(`\n  FINAL RESULT: ${rows.length} data rows x ${headers.length} columns`);
    console.log(`  Headers: ${headers.slice(0, 8).join(' | ')}${headers.length > 8 ? ' ...' : ''}`);
    if (rows[0]) console.log(`  Row 1: ${rows[0].slice(0, 6).join(' | ')}${rows[0].length > 6 ? ' ...' : ''}`);
    if (rows[rows.length - 1]) console.log(`  Last : ${rows[rows.length - 1].slice(0, 6).join(' | ')}`);

    // ── Build Excel ───────────────────────────────────────────
    const workbook = buildExcel(headers, rows);
    const outputName = `students-${Date.now()}.xlsx`;
    outputPath = path.join(outputsDir, outputName);
    await workbook.xlsx.writeFile(outputPath);

    res.download(outputPath, 'StudentData.xlsx', (err) => {
      safeDelete(uploadedPath);
      safeDelete(outputPath);
      if (err && !res.headersSent) console.error('Download error:', err.message);
    });

  } catch (err) {
    console.error('  Conversion error:', err);
    if (uploadedPath) safeDelete(uploadedPath);
    if (outputPath) safeDelete(outputPath);
    if (!res.headersSent) {
      res.status(500).json({ error: 'An error occurred during conversion. Please try again.' });
    }
  }
});

// ── Error handler ──────────────────────────────────────────────
app.use((err, _req, res, _next) => {
  if (err instanceof multer.MulterError) {
    return res.status(err.code === 'LIMIT_FILE_SIZE' ? 413 : 400).json({ error: err.message });
  }
  if (err) return res.status(400).json({ error: err.message || 'Something went wrong.' });
});

// ── Start ──────────────────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`
  ========================================
    DataMorph - PDF to Excel Converter
    Server running: http://0.0.0.0:${PORT}
    OCR-Powered for Zoho/SRM PDFs
  ========================================
  `);
});
