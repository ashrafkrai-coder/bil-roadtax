/**
 * Bil & Roadtax - Google Apps Script (Web App)
 *
 * Cara pasang:
 * 1. Buka Google Sheet baharu > Extensions > Apps Script.
 * 2. Padam kod sedia ada, tampal kod ini.
 * 3. Pastikan API_KEY sama dengan kunci dalam app (Tetapan > Kunci API).
 * 4. Deploy > New deployment > jenis: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Salin URL (berakhir /exec) dan tampal dalam app > Tetapan > URL Web App.
 *
 * Selepas mengubah kod: Deploy > Manage deployments > pensel > New version.
 */
const API_KEY = 'xCSFHhQev6US0P-ZfRw6fv0D';
const SHEET_NAME = 'Rekod';
const HEADERS = ['id', 'kategori', 'tarikh', 'tarikhLuput', 'plat', 'jumlah', 'nota', 'updatedAt', 'deleted'];
const TZ = 'Asia/Kuala_Lumpur';

function doGet() {
  return json_({ ok: true, info: 'Bil & Roadtax API aktif' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.key !== API_KEY) return json_({ ok: false, error: 'Kunci API salah' });

    const sh = getSheet_();
    const last = sh.getLastRow();
    const rowOf = {};
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 1).getValues().forEach((r, i) => {
        if (r[0] !== '') rowOf[String(r[0])] = i + 2;
      });
    }

    (body.records || []).forEach(rec => {
      const row = [
        String(rec.id), rec.kategori || '', rec.tarikh || '', rec.tarikhLuput || '',
        rec.plat || '', Number(rec.jumlah) || 0, rec.nota || '',
        Number(rec.updatedAt) || 0, rec.deleted === true
      ];
      const r = rowOf[String(rec.id)];
      if (r) {
        // Hanya tulis ganti jika versi dari app lebih baharu
        const current = Number(sh.getRange(r, 8).getValue()) || 0;
        if (row[7] >= current) sh.getRange(r, 1, 1, HEADERS.length).setValues([row]);
      } else {
        sh.appendRow(row);
        rowOf[String(rec.id)] = sh.getLastRow();
      }
    });

    return json_({ ok: true, records: readAll_(sh) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({
      id: String(r[0]),
      kategori: r[1],
      tarikh: asDate_(r[2]),
      tarikhLuput: asDate_(r[3]),
      plat: r[4],
      jumlah: Number(r[5]) || 0,
      nota: r[6],
      updatedAt: Number(r[7]) || 0,
      deleted: r[8] === true || String(r[8]).toUpperCase() === 'TRUE'
    }));
}

function asDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v ? String(v) : '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
