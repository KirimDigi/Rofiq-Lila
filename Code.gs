// =====================================================
// RSVP & Ucapan - Rofiq & Lila -> Google Sheets
// Cara pakai (tinggal copas):
// 1. Buka spreadsheet > menu Extensions > Apps Script
// 2. Hapus isi editor, tempel SELURUH file ini, Save
// 3. Deploy > New deployment > tipe "Web app"
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Deploy > salin URL yang berakhiran /exec
// 5. Kirim URL itu untuk dipasang ke index.html
// Sheet otomatis dibuat dengan header:
// Timestamp | Nama | Kehadiran | JumlahTamu | Ucapan
// =====================================================

const SHEET_ID = '';    // opsional: isi ID spreadsheet jika script ini berdiri sendiri
const SHEET_NAME = 'RSVP';

function getSheet_() {
  const ss = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Nama', 'Kehadiran', 'JumlahTamu', 'Ucapan']);
  } else if (sh.getLastColumn() < 5 && sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].join('|').indexOf('Ucapan') >= 0) {
    // migrasi sheet lama 4 kolom -> tambah kolom JumlahTamu di D
    sh.insertColumnAfter(3);
    sh.getRange(1, 4).setValue('JumlahTamu');
  }
  return sh;
}

function mapHadir_(v) {
  const map = { present: 'Hadir', notpresent: 'Tidak Hadir', notsure: 'Masih Ragu' };
  return map[v] || v || '';
}

// Menerima kiriman dari form undangan (POST: nama, kehadiran, tamu/guest, ucapan)
function doPost(e) {
  try {
    const d = (e && e.parameter) || {};
    const tamu = d.tamu || d.guest || d.jumlahTamu || '';
    getSheet_().appendRow([new Date(), d.nama || '', mapHadir_(d.kehadiran), String(tamu), d.ucapan || '']);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Mengambil daftar ucapan untuk ditampilkan di undangan (maks 100, terbaru dulu)
// Mendukung JSONP (?callback=...) agar bisa dimuat tanpa hambatan CORS
function doGet(e) {
  try {
    const sh = getSheet_();
    const out = [];
    const n = sh.getLastRow();
    if (n > 1) {
      const cols = sh.getLastColumn();
      const vals = sh.getRange(2, 1, n - 1, cols).getValues();
      const head = sh.getRange(1, 1, 1, cols).getValues()[0].map(String);
      const is5 = cols >= 5 && head[3] === 'JumlahTamu';
      for (let i = vals.length - 1; i >= 0 && out.length < 100; i--) {
        out.push({
          waktu: vals[i][0] instanceof Date ? vals[i][0].toISOString() : String(vals[i][0]),
          nama: String(vals[i][1]),
          kehadiran: String(vals[i][2]),
          tamu: String(is5 ? (vals[i][3] || '') : ''),
          ucapan: String(is5 ? vals[i][4] : vals[i][3])
        });
      }
    }
    const json = JSON.stringify(out);
    const cb = e && e.parameter && e.parameter.callback;
    if (cb) {
      return ContentService
        .createTextOutput(cb + '(' + json + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
