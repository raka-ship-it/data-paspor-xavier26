const SHEET_NAME = 'Data';
const FIRST_DATA_ROW = 2;
const MAX_NOTE_LENGTH = 1000;
const REVISION_COUNT_COLUMN = 17;
const CONFIRM_COUNT_COLUMN = 18;
const REVISION_LIMIT = 2;
const CONFIRM_LIMIT = 3;

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const sheet = getSheet();
    const rowCount = Math.max(sheet.getLastRow() - FIRST_DATA_ROW + 1, 0);
    const rows = rowCount ? sheet.getRange(FIRST_DATA_ROW, 4, rowCount, 15).getDisplayValues() : [];
    if (action === 'names') {
      const names = [...new Set(rows.map(row => String(row[0] || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
      return json({ ok: true, names });
    }
    if (action !== 'search') return json({ ok: false, message: 'Action tidak valid.' }, 400);
    const name = normalize(e.parameter.name);
    const prefix = normalize(e.parameter.passportPrefix).slice(0, 3);
    if (!name || prefix.length !== 3) return json({ ok: false, message: 'Nama dan 3 awalan paspor wajib diisi.' }, 400);
    const results = rows.map((r, i) => ({ r, rowId: FIRST_DATA_ROW + i }))
      .filter(x => normalize(x.r[0]) === name && normalize(x.r[4]).slice(0, 3) === prefix)
      .map(x => mapRow(x.r, x.rowId));
    return json({ ok: true, results });
  } catch (err) { return json({ ok: false, message: err.message }, 500); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const body = JSON.parse(e.postData.contents || '{}');
    const rowId = Number(body.rowId);
    if (!Number.isInteger(rowId) || rowId < FIRST_DATA_ROW) return json({ ok: false, message: 'Baris tidak valid.' }, 400);
    const sheet = getSheet();
    if (rowId > sheet.getLastRow()) return json({ ok: false, message: 'Data tidak ditemukan.' }, 404);
    const revisionCount = readCount(sheet, rowId, REVISION_COUNT_COLUMN);
    const confirmCount = readCount(sheet, rowId, CONFIRM_COUNT_COLUMN);
    if (body.action === 'confirm') {
      if (confirmCount >= CONFIRM_LIMIT) return json({ ok: false, message: 'Batas konfirmasi sudah tercapai.' }, 429);
      sheet.getRange(rowId, 14).setValue('Terkonfirmasi');
      sheet.getRange(rowId, CONFIRM_COUNT_COLUMN).setValue(confirmCount + 1);
    } else if (body.action === 'revision') {
      if (revisionCount >= REVISION_LIMIT) return json({ ok: false, message: 'Batas revisi sudah tercapai. Silakan konfirmasi data benar.' }, 429);
      const note = String(body.note || '').trim();
      if (!note || note.length > MAX_NOTE_LENGTH) return json({ ok: false, message: 'Catatan revisi wajib diisi dan maksimal 1.000 karakter.' }, 400);
      sheet.getRange(rowId, 16).setValue(note);
      sheet.getRange(rowId, REVISION_COUNT_COLUMN).setValue(revisionCount + 1);
    } else return json({ ok: false, message: 'Action tidak valid.' }, 400);
    sheet.getRange(rowId, 15).setValue(new Date());
    SpreadsheetApp.flush();
    return json({ ok: true, rowId, timestamp: new Date().toISOString(), revisionCount: body.action === 'revision' ? revisionCount + 1 : revisionCount, confirmCount: body.action === 'confirm' ? confirmCount + 1 : confirmCount, revisionLimit: REVISION_LIMIT, confirmLimit: CONFIRM_LIMIT });
  } catch (err) { return json({ ok: false, message: err.message }, 500); } finally { try { lock.releaseLock(); } catch (_) {} }
}

function getSheet() { const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME); if (!sheet) throw new Error('Tab Data tidak ditemukan.'); return sheet; }
function readCount(sheet, rowId, column) { const value = Number(sheet.getRange(rowId, column).getValue()); return Number.isFinite(value) && value >= 0 ? value : 0; }
function normalize(value) { return String(value || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function mapRow(r, rowId) { return { rowId, name: r[0], birthPlace: r[1], birthDate: r[2], gender: r[3], passportNumber: r[4], issueDate: r[5], expiryDate: r[6], jacketSize: r[9], confirmed: String(r[10] || '').trim().toLowerCase() === 'terkonfirmasi', revisionCount: Number(r[13]) || 0, confirmCount: Number(r[14]) || 0, revisionLimit: REVISION_LIMIT, confirmLimit: CONFIRM_LIMIT }; }
function json(payload) { return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON); }
