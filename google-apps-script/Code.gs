const SHEET_NAME = 'Data';
const FIRST_DATA_ROW = 2;
const MAX_NOTE_LENGTH = 1000;

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const sheet = getSheet();
    const rowCount = Math.max(sheet.getLastRow() - FIRST_DATA_ROW + 1, 0);
    const rows = rowCount ? sheet.getRange(FIRST_DATA_ROW, 4, rowCount, 13).getDisplayValues() : [];
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
    if (body.action === 'confirm') {
      sheet.getRange(rowId, 14).setValue('Terkonfirmasi');
    } else if (body.action === 'revision') {
      const note = String(body.note || '').trim();
      if (!note || note.length > MAX_NOTE_LENGTH) return json({ ok: false, message: 'Catatan revisi wajib diisi dan maksimal 1.000 karakter.' }, 400);
      sheet.getRange(rowId, 16).setValue(note);
    } else return json({ ok: false, message: 'Action tidak valid.' }, 400);
    sheet.getRange(rowId, 15).setValue(new Date());
    SpreadsheetApp.flush();
    return json({ ok: true, rowId, timestamp: new Date().toISOString() });
  } catch (err) { return json({ ok: false, message: err.message }, 500); } finally { try { lock.releaseLock(); } catch (_) {} }
}

function getSheet() { const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME); if (!sheet) throw new Error('Tab Data tidak ditemukan.'); return sheet; }
function normalize(value) { return String(value || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function mapRow(r, rowId) { return { rowId, name: r[0], birthPlace: r[1], birthDate: r[2], gender: r[3], passportNumber: r[4], issueDate: r[5], expiryDate: r[6], jacketSize: r[9] }; }
function json(payload) { return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON); }
