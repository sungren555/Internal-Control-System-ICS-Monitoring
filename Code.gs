/**
 * ICS Monitoring - Google Apps Script backend
 *
 * SETUP
 * 1. Open the Google Sheet "Internal Control System (ICS) Monitoring"
 * 2. Extensions > Apps Script, paste this whole file, Save
 * 3. Deploy > New deployment > type: Web app
 *      Execute as: Me      Who has access: Anyone
 * 4. Copy the Web app URL (ends with /exec) into index.html (API_URL)
 *
 * After changing this code, use Deploy > Manage deployments > Edit > New version.
 */

const SS_ID = '';              // leave empty when the script is opened from the Sheet itself
const TZ = 'Asia/Phnom_Penh';
const REGISTER_SHEET = 'Register farmers';
const REGISTER_FIRST_ROW = 4;  // data starts on row 4 (row 3 is the header)

const COMMON = ['No', 'Inspector', 'Date', 'IRPGs', 'Village', 'Farmily_Code',
                'Hus_Name(Eng)', 'Wif_Name(Eng)', 'year Join', 'Farmer Status', 'Compliance'];

const TARGETS = {
  'ICS-I':     { sheet: 'ICS-I',     header: COMMON.concat(['Planting']) },
  'NC':        { sheet: 'NC/Resign', header: COMMON.concat(['Note']) },
  'ICS-II':    { sheet: 'ICS-II',    header: COMMON.concat(['Threshing']) },
  'ICS-III':   { sheet: 'ICS-III',   header: COMMON.concat(['Burn Check']) }
};

function getSS_() {
  return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- READ ---------- */

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'sync';
    if (action === 'ping') return json_({ ok: true });
    const ss = getSS_();
    return json_({
      ok: true,
      time: new Date().toISOString(),
      farmers: readFarmers_(ss),
      ics1: readRecords_(ss, TARGETS['ICS-I'].sheet),
      nc:   readRecords_(ss, TARGETS['NC'].sheet),
      ics2: readRecords_(ss, TARGETS['ICS-II'].sheet),
      ics3: readRecords_(ss, TARGETS['ICS-III'].sheet)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function clean_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v).trim();
}

// Register farmers: A No | B IRPG | C Village | D Family code | E Hus (KH) | F Hus (EN)
//                   G Wife (KH) | H Wife (EN) | I Year join | J Status | K Level
function readFarmers_(ss) {
  const sh = ss.getSheetByName(REGISTER_SHEET);
  if (!sh) throw new Error('Sheet not found: ' + REGISTER_SHEET);
  const last = sh.getLastRow();
  if (last < REGISTER_FIRST_ROW) return [];
  const rows = sh.getRange(REGISTER_FIRST_ROW, 1, last - REGISTER_FIRST_ROW + 1, 11).getValues();
  const out = [];
  rows.forEach(function (r) {
    const code = clean_(r[3]);
    if (!code) return;
    out.push({
      irpg: clean_(r[1]), village: clean_(r[2]), code: code,
      husKh: clean_(r[4]), husEn: clean_(r[5]),
      wifeKh: clean_(r[6]), wifeEn: clean_(r[7]),
      yearJoin: clean_(r[8]), status: clean_(r[9]), level: clean_(r[10])
    });
  });
  return out;
}

// Output sheets: A No | B Inspector | C Date | D IRPG | E Village | F Code | G Hus | H Wife
//                I Year | J Status | K Compliance | L extra (Planting / Note / Threshing / Burn Check)
function readRecords_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, 12).getValues();
  const out = [];
  rows.forEach(function (r) {
    const code = clean_(r[5]);
    if (!code) return;
    out.push({
      no: clean_(r[0]), inspector: clean_(r[1]), date: clean_(r[2]),
      irpg: clean_(r[3]), village: clean_(r[4]), code: code,
      hus: clean_(r[6]), wife: clean_(r[7]), yearJoin: clean_(r[8]),
      status: clean_(r[9]), compliance: clean_(r[10]), extra: clean_(r[11])
    });
  });
  return out;
}

/* ---------- WRITE ---------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const body = JSON.parse(e.postData.contents);
    const records = body.records || [];
    const ss = getSS_();
    const idSheet = ensureIdSheet_(ss);
    const known = loadIds_(idSheet);

    const saved = [], duplicates = [];
    const rowsBySheet = {};   // target key -> array of row arrays
    const newIds = [];

    records.forEach(function (rec) {
      if (!rec.id) return;
      if (known[rec.id]) { duplicates.push(rec.id); return; }
      const key = targetKey_(rec);
      if (!key) return;
      (rowsBySheet[key] = rowsBySheet[key] || []).push(toRow_(rec, key));
      known[rec.id] = true;
      newIds.push([rec.id, new Date()]);
      saved.push(rec.id);
    });

    Object.keys(rowsBySheet).forEach(function (key) {
      appendRows_(ss, TARGETS[key], rowsBySheet[key]);
    });

    if (newIds.length) {
      idSheet.getRange(idSheet.getLastRow() + 1, 1, newIds.length, 2).setValues(newIds);
    }
    return json_({ ok: true, saved: saved, duplicates: duplicates });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// ICS-I: Compliance Yes -> "ICS-I"; No / Resign -> "NC/Resign"
function targetKey_(rec) {
  if (rec.section === 'ICS-I') return rec.compliance === 'Yes' ? 'ICS-I' : 'NC';
  if (rec.section === 'ICS-II') return 'ICS-II';
  if (rec.section === 'ICS-III') return 'ICS-III';
  return null;
}

function toRow_(rec, key) {
  const extra = key === 'NC' ? (rec.note || '') : (rec.extra || '');
  return [
    '',  // No - filled in appendRows_
    rec.inspector || '',
    rec.date || '',
    rec.irpg || '',
    rec.village || '',
    rec.code || '',
    rec.hus || '',
    rec.wife || '',
    rec.yearJoin || '',
    rec.status || '',
    rec.compliance || '',
    extra
  ];
}

function appendRows_(ss, target, rows) {
  let sh = ss.getSheetByName(target.sheet);
  if (!sh) sh = ss.insertSheet(target.sheet);
  if (!sh.getRange(1, 1).getValue()) {
    sh.getRange(1, 1, 1, target.header.length).setValues([target.header]).setFontWeight('bold');
  }
  const start = sh.getLastRow() + 1;
  rows.forEach(function (r, i) {
    r[0] = start + i - 1;                                  // running number
    const p = String(r[2]).split('-');                     // yyyy-mm-dd -> real date
    if (p.length === 3) r[2] = new Date(+p[0], +p[1] - 1, +p[2]);
  });
  const range = sh.getRange(start, 1, rows.length, 12);
  range.setNumberFormat('@');                              // keep codes and years as text
  sh.getRange(start, 3, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  range.setValues(rows);
}

function ensureIdSheet_(ss) {
  let sh = ss.getSheetByName('_ids');
  if (!sh) {
    sh = ss.insertSheet('_ids');
    sh.getRange(1, 1, 1, 2).setValues([['id', 'received']]);
    sh.hideSheet();
  }
  return sh;
}

function loadIds_(sh) {
  const map = {};
  const last = sh.getLastRow();
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { if (r[0]) map[r[0]] = true; });
  return map;
}
