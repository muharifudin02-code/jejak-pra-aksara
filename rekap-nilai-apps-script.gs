/**
 * Rekap Nilai "Jejak Pra-Aksara Nusantara"
 * Tempel seluruh kode ini di Google Sheets → Ekstensi → Apps Script.
 * GANTI sandi di bawah ini. Sandi hanya tersimpan di sini (akun Google guru), tidak di game.
 */
const ADMIN_PASSWORD = 'GANTI_SANDI_INI';

const LOG_H = ['Waktu', 'Kelas', 'Absen', 'Nama', 'Mode', 'Bab', 'Kuis', 'Skor', 'Total', 'Nilai', 'Lulus'];
const PRO_H = ['Kunci', 'Kelas', 'Absen', 'Nama', 'Mode', 'Tahap selesai', 'Jumlah tahap', 'Bintang', 'Rincian per bab', 'Diperbarui'];

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); s.getRange(1, 1, 1, head.length).setFontWeight('bold'); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Game mengirim hasil kuis dan tahapan ke sini. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const list = Array.isArray(data) ? data : [data];
    const log = sheet_('Log', LOG_H), pro = sheet_('Progres', PRO_H);
    list.forEach(function (x) {
      const t = new Date(x.waktu || Date.now());
      if (x.t === 'kuis') {
        log.appendRow([t, String(x.kelas), x.absen, x.nama, x.mode, x.bab, x.kuis, x.skor, x.total, x.persen, x.lulus ? 'Ya' : 'Belum']);
      } else if (x.t === 'progres') {
        const key = x.kelas + '|' + x.absen + '|' + x.mode;
        const row = [key, String(x.kelas), x.absen, x.nama, x.mode, x.tahap, x.tahapTotal, x.bintang, x.detail, t];
        const n = pro.getLastRow();
        const keys = n > 1 ? pro.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
        const i = keys.indexOf(key);
        if (i >= 0) pro.getRange(i + 2, 1, 1, row.length).setValues([row]); else pro.appendRow(row);
      }
    });
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/** Halaman admin di game membaca rekap dari sini (wajib sandi). */
function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== ADMIN_PASSWORD) return out_({ ok: false, error: 'sandi' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = function (n) {
    const s = ss.getSheetByName(n);
    return s && s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues() : [];
  };
  const ms = function (d) { return d instanceof Date ? d.getTime() : (Number(new Date(d)) || 0); };
  return out_({
    ok: true,
    log: rows('Log').map(function (r) {
      return { waktu: ms(r[0]), kelas: String(r[1]), absen: Number(r[2]), nama: r[3], mode: r[4], bab: r[5], kuis: r[6], skor: r[7], total: r[8], persen: Number(r[9]), lulus: r[10] === 'Ya' };
    }),
    progres: rows('Progres').map(function (r) {
      return { kelas: String(r[1]), absen: Number(r[2]), nama: r[3], mode: r[4], tahap: Number(r[5]), tahapTotal: Number(r[6]), bintang: Number(r[7]), detail: r[8], waktu: ms(r[9]) };
    })
  });
}
