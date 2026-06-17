// ============================================================
//  AIDA BIMBEL – Google Apps Script Backend
//  Programs: Tutor · Swim · Art · Mandarin
//
//  SETUP:
//  1. Buat Google Spreadsheet baru → Extensions → Apps Script
//  2. Paste seluruh file ini, Save
//  3. Deploy → New deployment → Web App → Execute as Me,
//     Who has access: Anyone → Deploy
//  4. Copy URL /exec → paste ke api.js (SCRIPT_URL)
//  5. Buka admin.html → console browser → ketik: API.setup()
//     lalu (opsional, sekali saja): API.seed()
//     → mengisi data guru & murid awal
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();
const FOLDER_NAME = 'AidaBimbel_BuktiBayar';

// ── Helpers ─────────────────────────────────────────────────
function uid() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}
function getSheet(name) {
  return SS.getSheetByName(name) || SS.insertSheet(name);
}
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const WA_COLS   = ['wa_ortu', 'wa_laporan', 'wa'];
  const DATE_COLS = ['tanggal', 'tgl_masuk', 'tgl_berhenti', 'tgl_kirim'];
  const tz = Session.getScriptTimeZone();
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
      if (WA_COLS.includes(h)) val = String(val || '');
      // Sheets suka auto-convert "2026-06" jadi tanggal — normalisasi balik ke YYYY-MM
      if (h === 'bulan' && val) val = String(val).slice(0, 7);
      obj[h] = val;
    });
    return obj;
  });
}
function ok(data) { return ContentService.createTextOutput(JSON.stringify({ status: 'ok', data })).setMimeType(ContentService.MimeType.JSON); }
function err(msg) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: msg })).setMimeType(ContentService.MimeType.JSON); }

// Tulis baris berdasarkan NAMA kolom (aman walau urutan/jumlah kolom beda)
function appendByHeaders(sheet, obj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sheet.appendRow(row);
  return headers;
}

function deleteRowById(sheetName, id) {
  const sheet = getSheet(sheetName);
  const data  = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idIdx] === id) { sheet.deleteRow(i + 1); return { deleted: id }; }
  }
  throw new Error(sheetName + ' tidak ditemukan: ' + id);
}

function updateRowById(sheetName, id, fields) {
  const sheet   = getSheet(sheetName);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx   = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      for (const [key, val] of Object.entries(fields)) {
        const col = headers.indexOf(key);
        if (col >= 0 && val !== undefined) {
          const cell = sheet.getRange(i + 1, col + 1);
          if (['wa_ortu', 'wa_laporan', 'wa'].includes(key)) cell.setNumberFormat('@');
          cell.setValue(val);
        }
      }
      return { updated: id };
    }
  }
  throw new Error(sheetName + ' tidak ditemukan: ' + id);
}

// ── Entry Points ─────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter || {};
    switch (p.action) {
      case 'getMuridByLink': return ok(getMuridByLink(p.link_id));
      default:               return err('Unknown GET action: ' + p.action);
    }
  } catch (ex) { return err(ex.message); }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    switch (p.action) {
      // SETUP
      case 'setup':                return ok(setupSheets());
      case 'seed':                 return ok(seedAida());
      case 'addColumn':            return ok(addColumn(p.sheet, p.kolom, p.default_val));
      case 'migrate':              return ok(migrateSheets());
      // LOGIN
      case 'loginOrtu':            return ok(loginOrtu(p.phone));
      case 'loginGuru':            return ok(loginGuru(p.phone));
      case 'loginAdmin':           return ok(loginAdmin(p.phone, p.password));
      case 'generatePin':          return ok(generatePin(p.murid_id));
      // MURID
      case 'getMurid':             return ok(getMurid(p));
      case 'getMuridByLink':       return ok(getMuridByLink(p.link_id));
      case 'addMurid':             return ok(addMurid(p));
      case 'updateMurid':          return ok(updateMurid(p));
      case 'deleteMurid':          return ok(deleteMurid(p.id));
      // KELAS (enrollment murid per program)
      case 'getKelas':             return ok(getKelas(p));
      case 'addKelas':             return ok(addKelas(p));
      case 'updateKelas':          return ok(updateKelas(p));
      case 'deleteKelas':          return ok(deleteKelas(p.id));
      case 'getSesiAllKelas':      return ok(getSesiAllKelas());
      // GURU
      case 'getGuru':              return ok(getGuru());
      case 'addGuru':              return ok(addGuru(p));
      case 'updateGuru':           return ok(updateGuru(p));
      case 'deleteGuru':           return ok(deleteGuru(p.id));
      case 'getFeeGuru':           return ok(getFeeGuru(p.bulan));
      // LAPORAN
      case 'getLaporan':           return ok(getLaporan(p));
      case 'addLaporan':           return ok(addLaporan(p));
      case 'updateLaporan':        return ok(updateLaporan(p));
      case 'deleteLaporan':        return ok(deleteLaporan(p.id));
      // ABSENSI
      case 'getAbsensi':           return ok(getAbsensi(p));
      case 'saveAbsensi':          return ok(saveAbsensi(p.tanggal, p.data));
      case 'upsertAbsensi':        return ok(upsertAbsensi(p));
      // INVOICE
      case 'getInvoice':           return ok(getInvoice(p));
      case 'generateInvoice':      return ok(generateInvoice(p));
      case 'updateInvoiceStatus':  return ok(updateInvoiceStatus(p.id, p.status, p.tgl_kirim));
      case 'updateInvoiceNominal': return ok(updateInvoiceNominal(p.id, p.sesi, p.rate, p.nominal));
      case 'deleteInvoice':        return ok(deleteInvoice(p.id));
      case 'uploadBukti':          return ok(uploadBukti(p));
      // SETTINGS
      case 'getSettings':          return ok(getSettings());
      case 'saveSettings':         return ok(saveSettings(p));
      default:                     return err('Unknown action: ' + p.action);
    }
  } catch (ex) { return err(ex.message); }
}

// ══════════════════════════════════════════════════════════════
//  SETUP — buat semua sheet dengan header yang benar
// ══════════════════════════════════════════════════════════════
const SCHEMAS = {
  'Murid':   ['id', 'nama', 'grade', 'kelas_sekolah', 'wa_ortu', 'wa_laporan', 'rate_invoice', 'pin', 'tgl_masuk', 'tgl_berhenti', 'aktif', 'link_id'],
  'Kelas':   ['id', 'murid_id', 'nama_murid', 'guru_id', 'nama_guru', 'program', 'tipe', 'jadwal', 'sesi_kuota', 'fee_guru', 'tgl_mulai_term', 'tgl_akhir_term', 'aktif'],
  'Guru':    ['id', 'nama', 'wa', 'default_fee', 'aktif'],
  'Laporan': ['id', 'kelas_id', 'murid_id', 'nama_murid', 'guru_id', 'nama_guru', 'tanggal', 'program', 'tipe', 'materi_json', 'catatan', 'kehadiran', 'extra', 'foto', 'ttd', 'status', 'timestamp'],
  'Absensi': ['id', 'kelas_id', 'murid_id', 'nama', 'tanggal', 'status', 'catatan'],
  'Invoice': ['id', 'murid_id', 'nama_murid', 'bulan', 'program_label', 'sesi', 'rate', 'nominal', 'status', 'tgl_kirim', 'bukti_url', 'catatan'],
  'Setting': ['key', 'value'],
};

// Tambah kolom yang kurang ke sheet lama (aman utk data existing) — jalankan: API.migrate()
function migrateSheets() {
  const results = [];
  for (const [name, headers] of Object.entries(SCHEMAS)) {
    const sheet = getSheet(name);
    const existing = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    if (existing.length === 0 || existing[0] === '') {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#F7B6C2').setFontColor('#5B3A47').setFontWeight('bold');
      sheet.setFrozenRows(1);
      results.push(name + ': created');
      continue;
    }
    headers.forEach(h => {
      if (!existing.includes(h)) {
        addColumn(name, h, h === 'status' && name === 'Laporan' ? 'approved' : undefined);
        results.push(name + ': +' + h);
      }
    });
  }
  // Default kredensial admin (ganti di sheet Setting!)
  const st = getSheet('Setting');
  if (st.getLastRow() < 2) {
    if (st.getLastRow() === 0 || st.getRange(1, 1).getValue() === '') {
      st.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
      st.getRange(1, 1, 1, 2).setBackground('#F7B6C2').setFontColor('#5B3A47').setFontWeight('bold');
    }
    st.appendRow(['admin_phone', '085133136377']);
    st.appendRow(['admin_password', 'aidabimbel123']);
    results.push('Setting: login admin default dibuat → HP 085133136377, password aidabimbel123 (GANTI di sheet Setting!)');
  }
  return results.length ? results : ['Semua sheet sudah terbaru'];
}

function setupSheets() {
  const schemas = SCHEMAS;
  const results = [];
  for (const [name, headers] of Object.entries(schemas)) {
    const sheet = getSheet(name);
    const existing = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    if (existing.length === 0 || existing[0] === '') {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#F7B6C2').setFontColor('#5B3A47').setFontWeight('bold');
      sheet.setFrozenRows(1);
      ['wa_ortu', 'wa_laporan', 'wa'].forEach(key => {
        const col = headers.indexOf(key);
        if (col >= 0) sheet.getRange(2, col + 1, 1000, 1).setNumberFormat('@');
      });
      results.push(name + ': created');
    } else {
      results.push(name + ': already exists');
    }
  }
  return results;
}

// ── Tambah kolom baru ke sheet yang sudah ada (aman, tanpa hapus data)
function addColumn(sheetName, kolom, defaultVal) {
  if (!sheetName || !kolom) throw new Error('Wajib: sheet & kolom. Contoh: API.addColumn("Murid","catatan")');
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.includes(kolom)) return { skipped: 'Kolom "' + kolom + '" sudah ada di ' + sheetName };
  const col = headers.length + 1;
  sheet.getRange(1, col).setValue(kolom)
    .setBackground('#F7B6C2').setFontColor('#5B3A47').setFontWeight('bold');
  const lastRow = sheet.getLastRow();
  if (defaultVal !== undefined && defaultVal !== '' && lastRow > 1) {
    sheet.getRange(2, col, lastRow - 1, 1).setValue(defaultVal);
  }
  return { added: kolom, sheet: sheetName, default_val: defaultVal || null };
}

// ══════════════════════════════════════════════════════════════
//  MURID
// ══════════════════════════════════════════════════════════════
function getMurid(p) {
  const rows = sheetToObjects(getSheet('Murid'));
  if (p.id)    return rows.find(r => r.id === p.id) || null;
  if (p.aktif) return rows.filter(r => r.aktif === p.aktif);
  return rows;
}

function getMuridByLink(link_id) {
  const rows  = sheetToObjects(getSheet('Murid'));
  const murid = rows.find(r => r.link_id === link_id);
  if (!murid) throw new Error('Murid tidak ditemukan untuk link: ' + link_id);
  murid.kelas = sheetToObjects(getSheet('Kelas')).filter(k => k.murid_id === murid.id && k.aktif === 'aktif');
  return murid;
}

function addMurid(p) {
  const sheet   = getSheet('Murid');
  const id      = uid();
  const link_id = String(p.nama || 'murid').toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '') + '-' + id.slice(0, 6);
  const hdrs = appendByHeaders(sheet, {
    id,
    nama:          p.nama || '',
    grade:         p.grade || '',
    kelas_sekolah: p.kelas_sekolah || '',
    wa_ortu:       p.wa_ortu || '',
    wa_laporan:    p.wa_laporan || '',
    rate_invoice:  Number(p.rate_invoice) || 0,
    pin:           p.pin || String(Math.floor(100000 + Math.random() * 900000)),
    tgl_masuk:     p.tgl_masuk || '',
    tgl_berhenti:  p.tgl_berhenti || '',
    aktif:         p.aktif || 'aktif',
    link_id,
  });
  const lastRow = sheet.getLastRow();
  ['wa_ortu', 'wa_laporan'].forEach(key => {
    const col = hdrs.indexOf(key);
    if (col >= 0) sheet.getRange(lastRow, col + 1).setNumberFormat('@');
  });
  return { id, link_id };
}

function updateMurid(p) {
  const r = updateRowById('Murid', p.id, {
    nama: p.nama, grade: p.grade, kelas_sekolah: p.kelas_sekolah,
    wa_ortu: p.wa_ortu, wa_laporan: p.wa_laporan,
    rate_invoice: p.rate_invoice !== undefined ? Number(p.rate_invoice) || 0 : undefined,
    pin: p.pin, tgl_masuk: p.tgl_masuk, tgl_berhenti: p.tgl_berhenti, aktif: p.aktif,
  });
  // Sinkronkan nama murid di sheet Kelas
  if (p.nama) {
    const kelas = getSheet('Kelas');
    const data  = kelas.getDataRange().getValues();
    if (data.length > 1) {
      const h = data[0], mi = h.indexOf('murid_id'), ni = h.indexOf('nama_murid');
      for (let i = 1; i < data.length; i++) {
        if (data[i][mi] === p.id) kelas.getRange(i + 1, ni + 1).setValue(p.nama);
      }
    }
  }
  return r;
}

function deleteMurid(id) { return deleteRowById('Murid', id); }

// ══════════════════════════════════════════════════════════════
//  KELAS — enrollment murid per program (kuota & fee guru per kelas)
// ══════════════════════════════════════════════════════════════
function getKelas(p) {
  let rows = sheetToObjects(getSheet('Kelas'));
  if (p.id)       return rows.find(r => r.id === p.id) || null;
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.guru_id)  rows = rows.filter(r => r.guru_id === p.guru_id);
  if (p.aktif)    rows = rows.filter(r => r.aktif === p.aktif);
  return rows;
}

function addKelas(p) {
  const sheet = getSheet('Kelas');
  const id = uid();
  const murid = getMurid({ id: p.murid_id });
  if (!murid) throw new Error('Murid tidak ditemukan di list murid — refresh halaman lalu coba lagi');
  const guru  = sheetToObjects(getSheet('Guru')).find(g => g.id === p.guru_id);
  appendByHeaders(sheet, {
    id,
    murid_id:   p.murid_id || '',
    nama_murid: murid ? murid.nama : (p.nama_murid || ''),
    guru_id:    p.guru_id || '',
    nama_guru:  guru ? guru.nama : (p.nama_guru || ''),
    program:    p.program || 'Tutor',
    tipe:       p.tipe || 'Private',
    jadwal:     p.jadwal || '',
    sesi_kuota: Number(p.sesi_kuota) || 8,
    fee_guru:   Number(p.fee_guru) || 0,
    tgl_mulai_term: p.tgl_mulai_term || '',
    tgl_akhir_term: p.tgl_akhir_term || '',
    aktif: p.aktif || 'aktif',
  });
  return { id };
}

function updateKelas(p) {
  return updateRowById('Kelas', p.id, {
    guru_id: p.guru_id, nama_guru: p.nama_guru,
    program: p.program, tipe: p.tipe, jadwal: p.jadwal,
    sesi_kuota: p.sesi_kuota !== undefined ? Number(p.sesi_kuota) || 0 : undefined,
    fee_guru:   p.fee_guru   !== undefined ? Number(p.fee_guru)   || 0 : undefined,
    tgl_mulai_term: p.tgl_mulai_term, tgl_akhir_term: p.tgl_akhir_term,
    aktif: p.aktif,
  });
}

function deleteKelas(id) { return deleteRowById('Kelas', id); }

// ══════════════════════════════════════════════════════════════
//  GURU
// ══════════════════════════════════════════════════════════════
function getGuru() { return sheetToObjects(getSheet('Guru')); }

function addGuru(p) {
  const sheet = getSheet('Guru');
  const id = uid();
  const hdrs = appendByHeaders(sheet, { id, nama: p.nama || '', wa: p.wa || '', default_fee: Number(p.default_fee) || 0, aktif: p.aktif || 'aktif' });
  const waCol = hdrs.indexOf('wa');
  if (waCol >= 0) sheet.getRange(sheet.getLastRow(), waCol + 1).setNumberFormat('@');
  return { id };
}

function updateGuru(p) {
  const r = updateRowById('Guru', p.id, {
    nama: p.nama, wa: p.wa,
    default_fee: p.default_fee !== undefined ? Number(p.default_fee) || 0 : undefined,
    aktif: p.aktif,
  });
  if (p.nama) {
    const kelas = getSheet('Kelas');
    const data  = kelas.getDataRange().getValues();
    if (data.length > 1) {
      const h = data[0], gi = h.indexOf('guru_id'), ni = h.indexOf('nama_guru');
      for (let i = 1; i < data.length; i++) {
        if (data[i][gi] === p.id) kelas.getRange(i + 1, ni + 1).setValue(p.nama);
      }
    }
  }
  return r;
}

function deleteGuru(id) { return deleteRowById('Guru', id); }

// ── FEE GURU — dihitung dari laporan harian yang diisi guru ──
function getFeeGuru(bulan) {
  if (!bulan) throw new Error('Parameter bulan (YYYY-MM) wajib diisi');
  const laporan = sheetToObjects(getSheet('Laporan')).filter(r => String(r.tanggal).startsWith(bulan));
  const kelas   = sheetToObjects(getSheet('Kelas'));
  const guru    = sheetToObjects(getSheet('Guru'));
  const kelasMap = {};
  kelas.forEach(k => { kelasMap[k.id] = k; });

  const perGuru = {};
  laporan.forEach(l => {
    const gid = l.guru_id || 'unknown';
    if (!perGuru[gid]) perGuru[gid] = {};
    const kid = l.kelas_id || 'tanpa-kelas';
    if (!perGuru[gid][kid]) perGuru[gid][kid] = 0;
    perGuru[gid][kid]++;
  });

  return Object.entries(perGuru).map(([gid, kelasCounts]) => {
    const g = guru.find(x => x.id === gid);
    let total_sesi = 0, total_fee = 0;
    const detail = Object.entries(kelasCounts).map(([kid, count]) => {
      const k   = kelasMap[kid];
      const fee = k ? (Number(k.fee_guru) || Number(g && g.default_fee) || 0) : (Number(g && g.default_fee) || 0);
      total_sesi += count;
      total_fee  += fee * count;
      return {
        kelas_id: kid,
        nama_murid: k ? k.nama_murid : '(kelas terhapus)',
        program: k ? k.program : '-',
        tipe: k ? k.tipe : '-',
        sesi: count,
        fee_per_sesi: fee,
        subtotal: fee * count,
      };
    });
    return { guru_id: gid, nama_guru: g ? g.nama : '(tidak dikenal)', wa: g ? g.wa : '', total_sesi, total_fee, detail };
  }).sort((a, b) => a.nama_guru > b.nama_guru ? 1 : -1);
}

// ══════════════════════════════════════════════════════════════
//  LAPORAN — materi_json: [{"m":"materi","s":3}, ...]  s = bintang 1-4
// ══════════════════════════════════════════════════════════════
function getLaporan(p) {
  let rows = sheetToObjects(getSheet('Laporan'));
  rows.forEach(r => { if (!r.status) r.status = 'approved'; }); // baris lama dianggap approved
  if (p.status)   rows = rows.filter(r => r.status === p.status);
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.guru_id)  rows = rows.filter(r => r.guru_id === p.guru_id);
  if (p.kelas_id) rows = rows.filter(r => r.kelas_id === p.kelas_id);
  if (p.bulan)    rows = rows.filter(r => String(r.tanggal).startsWith(p.bulan));
  if (p.dari)     rows = rows.filter(r => String(r.tanggal) >= p.dari);
  if (p.sampai)   rows = rows.filter(r => String(r.tanggal) <= p.sampai);
  return rows.sort((a, b) => b.tanggal > a.tanggal ? 1 : -1);
}

function addLaporan(p) {
  const sheet = getSheet('Laporan');
  const id = uid();
  const kelas = p.kelas_id ? getKelas({ id: p.kelas_id }) : null;
  const murid_id   = p.murid_id || (kelas ? kelas.murid_id : '');
  const guru_id    = p.guru_id  || (kelas ? kelas.guru_id  : '');
  const muridRows  = sheetToObjects(getSheet('Murid'));
  const guruRows   = sheetToObjects(getSheet('Guru'));
  const murid = muridRows.find(r => r.id === murid_id);
  const guru  = guruRows.find(r => r.id === guru_id);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  const set = (key, val) => { const i = headers.indexOf(key); if (i >= 0) row[i] = val; };
  set('id', id);
  set('kelas_id',   p.kelas_id || '');
  set('murid_id',   murid_id);
  set('nama_murid', murid ? murid.nama : '');
  set('guru_id',    guru_id);
  set('nama_guru',  guru ? guru.nama : '');
  set('tanggal',    p.tanggal || '');
  set('program',    p.program || (kelas ? kelas.program : ''));
  set('tipe',       p.tipe || (kelas ? kelas.tipe : ''));
  set('materi_json', typeof p.materi === 'string' ? p.materi : JSON.stringify(p.materi || []));
  set('catatan',    p.catatan || '');
  set('kehadiran',  p.kehadiran || 'hadir');
  set('extra',      typeof p.extra === 'string' ? p.extra : (p.extra ? JSON.stringify(p.extra) : ''));
  set('foto',       p.foto || '');
  set('ttd',        p.ttd  || '');
  set('status',     p.status || 'pending'); // menunggu approval admin sebelum dikirim ke ortu
  set('timestamp',  new Date().toISOString());
  sheet.appendRow(row);

  // Auto-absen dengan status kehadiran dari guru (hadir atau alpha)
  if (p.auto_absen && p.kelas_id && p.tanggal) {
    const absenStatus = ['hadir', 'alpha'].includes(p.kehadiran) ? p.kehadiran : 'hadir';
    upsertAbsensi({ kelas_id: p.kelas_id, murid_id, nama: murid ? murid.nama : '', tanggal: p.tanggal, status: absenStatus, catatan: 'auto: laporan guru' });
  }
  return { id };
}

// Edit laporan (admin: koreksi typo/bahasa) & ubah status approval
function updateLaporan(p) {
  return updateRowById('Laporan', p.id, {
    tanggal:   p.tanggal,
    materi_json: p.materi !== undefined ? (typeof p.materi === 'string' ? p.materi : JSON.stringify(p.materi)) : undefined,
    catatan:   p.catatan,
    kehadiran: p.kehadiran,
    extra:     p.extra !== undefined ? (typeof p.extra === 'string' ? p.extra : JSON.stringify(p.extra)) : undefined,
    foto:      p.foto  !== undefined ? p.foto  : undefined,
    ttd:       p.ttd   !== undefined ? p.ttd   : undefined,
    status:    p.status,
  });
}

function deleteLaporan(id) { return deleteRowById('Laporan', id); }

// ══════════════════════════════════════════════════════════════
//  LOGIN ORTU — nomor HP + PIN (PIN di-generate admin)
// ══════════════════════════════════════════════════════════════
function normPhone(s) {
  let d = String(s || '').replace(/\D/g, '');
  if (d.startsWith('62')) d = '0' + d.slice(2);
  return d;
}

function loginOrtu(phone) {
  const ph = normPhone(phone);
  if (!ph) throw new Error('Isi nomor HP dulu ya');
  const murid = sheetToObjects(getSheet('Murid'));
  const samePhone = murid.filter(m => normPhone(m.wa_ortu) === ph || normPhone(m.wa_laporan) === ph);
  if (!samePhone.length) throw new Error('Nomor HP tidak terdaftar. Hubungi admin AIDA BIMBEL ya 🙏');
  const anak = samePhone.filter(m => m.aktif === 'aktif');
  if (!anak.length) throw new Error('Status murid nonaktif — akses portal dimatikan. Hubungi admin AIDA BIMBEL 🙏');
  const kelas = sheetToObjects(getSheet('Kelas'));
  return anak.map(m => {
    const k = kelas.filter(x => x.murid_id === m.id && x.aktif === 'aktif');
    return { id: m.id, nama: m.nama, grade: m.grade, kelas_sekolah: m.kelas_sekolah, kelas: k };
  });
}

function loginGuru(phone) {
  const ph = normPhone(phone);
  if (!ph) throw new Error('Isi nomor HP dulu ya');
  const guru = sheetToObjects(getSheet('Guru')).find(g => normPhone(g.wa) === ph);
  if (!guru) throw new Error('Nomor HP tidak terdaftar sebagai guru. Hubungi admin 🙏');
  if (guru.aktif === 'nonaktif') throw new Error('Status guru nonaktif — akses dimatikan. Hubungi admin 🙏');
  return { id: guru.id, nama: guru.nama };
}

function getSetting(key) {
  const rows = sheetToObjects(getSheet('Setting'));
  const r = rows.find(x => String(x.key) === key);
  return r ? String(r.value || '') : '';
}

function loginAdmin(phone, password) {
  const adminPhone = normPhone(getSetting('admin_phone'));
  const adminPass  = getSetting('admin_password');
  if (!adminPass) throw new Error('Login admin belum diset — jalankan API.migrate() lalu cek sheet Setting');
  if (adminPhone && normPhone(phone) !== adminPhone) throw new Error('Nomor HP admin salah');
  if (String(password || '') !== adminPass) throw new Error('Password salah');
  return { ok: true };
}

function generatePin(murid_id) {
  const pin = String(Math.floor(100000 + Math.random() * 900000));
  updateRowById('Murid', murid_id, { pin });
  return { murid_id, pin };
}

// ══════════════════════════════════════════════════════════════
//  ABSENSI
// ══════════════════════════════════════════════════════════════
function getAbsensi(p) {
  let rows = sheetToObjects(getSheet('Absensi'));
  if (p.tanggal)  rows = rows.filter(r => r.tanggal === p.tanggal);
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.kelas_id) rows = rows.filter(r => r.kelas_id === p.kelas_id);
  if (p.bulan)    rows = rows.filter(r => String(r.tanggal).startsWith(p.bulan));
  if (p.dari)     rows = rows.filter(r => String(r.tanggal) >= p.dari);
  if (p.sampai)   rows = rows.filter(r => String(r.tanggal) <= p.sampai);
  return rows;
}

// Simpan absensi per tanggal (replace semua baris tanggal tsb)
function saveAbsensi(tanggal, data) {
  const sheet    = getSheet('Absensi');
  const existing = sheet.getDataRange().getValues();
  const headers  = existing[0];
  const tglIdx   = headers.indexOf('tanggal');
  const tz = Session.getScriptTimeZone();
  for (let i = existing.length - 1; i >= 1; i--) {
    let v = existing[i][tglIdx];
    if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    if (v === tanggal) sheet.deleteRow(i + 1);
  }
  data.forEach(d => {
    appendByHeaders(sheet, { id: uid(), kelas_id: d.kelas_id || '', murid_id: d.murid_id, nama: d.nama || '', tanggal, status: d.status, catatan: d.catatan || '' });
  });
  return { saved: data.length };
}

// Upsert satu baris absensi (dipakai auto-absen dari laporan guru)
function upsertAbsensi(p) {
  const sheet   = getSheet('Absensi');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const tglIdx  = headers.indexOf('tanggal');
  const kidIdx  = headers.indexOf('kelas_id');
  const stIdx   = headers.indexOf('status');
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < data.length; i++) {
    let v = data[i][tglIdx];
    if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    if (v === p.tanggal && data[i][kidIdx] === p.kelas_id) {
      sheet.getRange(i + 1, stIdx + 1).setValue(p.status);
      return { updated: true };
    }
  }
  appendByHeaders(sheet, { id: uid(), kelas_id: p.kelas_id || '', murid_id: p.murid_id || '', nama: p.nama || '', tanggal: p.tanggal, status: p.status || 'hadir', catatan: p.catatan || '' });
  return { inserted: true };
}

// ── SESI BERJALAN per kelas — hadir sejak invoice lunas terakhir
function getSesiAllKelas() {
  const invoices = sheetToObjects(getSheet('Invoice'));
  const absensi  = sheetToObjects(getSheet('Absensi'));
  const kelas    = sheetToObjects(getSheet('Kelas')).filter(k => k.aktif === 'aktif');
  const tz = Session.getScriptTimeZone();

  const sinceMap = {}; // murid_id → tanggal lunas terakhir
  invoices.filter(i => i.status === 'lunas').forEach(i => {
    let d = i.tgl_kirim || '';
    if (!d && i.bulan) {
      const parts = String(i.bulan).split('-').map(Number);
      d = Utilities.formatDate(new Date(parts[0], parts[1], 0), tz, 'yyyy-MM-dd');
    }
    if (d && (!sinceMap[i.murid_id] || d > sinceMap[i.murid_id])) sinceMap[i.murid_id] = d;
  });

  return kelas.map(k => {
    const since = sinceMap[k.murid_id] || null;
    // Hadir dan Alpha sama-sama menghitung sesi (slot sudah dialokasikan)
    const sesi = absensi.filter(a => {
      if (a.kelas_id !== k.id || !['hadir', 'alpha'].includes(a.status)) return false;
      if (since) return a.tanggal > since;
      return true;
    }).length;
    return {
      kelas_id: k.id, murid_id: k.murid_id, nama: k.nama_murid,
      guru_id: k.guru_id, nama_guru: k.nama_guru,
      program: k.program, tipe: k.tipe, jadwal: k.jadwal,
      sesi, kuota: Number(k.sesi_kuota) || 8, since,
      tgl_mulai_term: k.tgl_mulai_term || '', tgl_akhir_term: k.tgl_akhir_term || '',
    };
  });
}

// ══════════════════════════════════════════════════════════════
//  INVOICE
// ══════════════════════════════════════════════════════════════
function getInvoice(p) {
  let rows = sheetToObjects(getSheet('Invoice'));
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.bulan)    rows = rows.filter(r => String(r.bulan).slice(0, 7) === String(p.bulan).slice(0, 7));
  return rows.sort((a, b) => b.bulan > a.bulan ? 1 : -1);
}

function generateInvoice(p) {
  const murid = getMurid({ id: p.murid_id });
  if (!murid) throw new Error('Murid tidak ditemukan');
  const sesi = Number(p.sesi) || 0;
  if (sesi <= 0) throw new Error('Jumlah sesi harus lebih dari 0');
  const rate    = Number(p.rate) || Number(murid.rate_invoice) || 0;
  const nominal = p.nominal !== undefined && p.nominal !== '' ? Number(p.nominal) : rate * sesi;

  const existing = getInvoice({ murid_id: p.murid_id, bulan: p.bulan });
  if (existing.length > 0) throw new Error('Invoice ' + murid.nama + ' untuk ' + p.bulan + ' sudah ada');

  // Label program otomatis dari kelas aktif murid bila tidak diisi
  let label = p.program_label || '';
  if (!label) {
    const progs = [...new Set(getKelas({ murid_id: p.murid_id, aktif: 'aktif' }).map(k => k.program))];
    label = progs.join(' & ');
  }

  const sheet = getSheet('Invoice');
  const id = uid();
  appendByHeaders(sheet, {
    id, murid_id: p.murid_id, nama_murid: murid.nama, bulan: p.bulan, program_label: label,
    sesi, rate, nominal, status: 'belum_bayar', tgl_kirim: '', bukti_url: '', catatan: p.catatan || '',
  });
  return { id, sesi, rate, nominal, program_label: label };
}

function updateInvoiceStatus(id, status, tgl_kirim) {
  return updateRowById('Invoice', id, { status, tgl_kirim: tgl_kirim || undefined });
}

function updateInvoiceNominal(id, sesi, rate, nominal) {
  return updateRowById('Invoice', id, {
    sesi:    sesi    !== undefined ? Number(sesi)    : undefined,
    rate:    rate    !== undefined ? Number(rate)    : undefined,
    nominal: nominal !== undefined ? Number(nominal) : undefined,
  });
}

function deleteInvoice(id) { return deleteRowById('Invoice', id); }

function uploadBukti(p) {
  let folder;
  try {
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  } catch (e) { folder = DriveApp.getRootFolder(); }

  const base64 = p.base64.split(',')[1];
  const mime   = p.base64.split(';')[0].split(':')[1];
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), mime, p.filename || 'bukti.jpg');
  const file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/uc?id=' + file.getId();

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  updateRowById('Invoice', p.invoice_id, { status: 'lunas', bukti_url: url, tgl_kirim: today });
  return { url };
}

// ══════════════════════════════════════════════════════════════
//  SEED — isi data awal (jalankan sekali via API.seed())
// ══════════════════════════════════════════════════════════════
function seedAida() {
  if (sheetToObjects(getSheet('Guru')).length > 0) throw new Error('Data sudah ada — seed dibatalkan agar tidak dobel.');

  const guruNames = ['Coach Rere', 'Miss Devi', 'Miss Sherly', 'Laoshi Lia', 'Miss Keren', 'Miss Lala', 'Miss Siska', 'Mr Hendra', 'Miss Cey'];
  const guruIds = {};
  guruNames.forEach(n => { guruIds[n] = addGuru({ nama: n }).id; });

  // [nama murid, guru, program, tipe, jadwal, kuota]
  const seed = [
    ['Karsa', 'Coach Rere', 'Swim', 'Private', 'Mon 15.30', 4],
    ['Ezio', 'Coach Rere', 'Tutor', 'Private', '15.00', 5],
    ['Shinnaz', 'Coach Rere', 'Swim', 'Private', 'Tue-Thu 17.00', 8],
    ['Russel', 'Coach Rere', 'Swim', 'Private', 'Sat 15.15 / Wed 17.15', 8],
    ['Millie Karen Keiko', 'Coach Rere', 'Swim', 'Semi Private', 'Wed 15.00', 10],
    ['Clara', 'Coach Rere', 'Swim', 'Private', 'Sat 14.30', 4],
    ['Faith', 'Miss Devi', 'Tutor', 'Semi Private', 'Tue-Thu 16.00', 8],
    ['Winston', 'Miss Devi', 'Tutor', 'Private', 'Mon-Thu 15.00', 8],
    ['Clinton', 'Miss Devi', 'Tutor', 'Private', 'Mon-Thu 14.00', 8],
    ['Aaron', 'Miss Devi', 'Tutor', 'Private', 'Tue, Wed-Thu 12.00', 12],
    ['Joshua', 'Miss Devi', 'Tutor', 'Private', 'Mon, Wed/Fri 13.00', 12],
    ['Valerie Axelle', 'Miss Devi', 'Tutor', 'Private', 'Tue/Fri 17.30', 8],
    ['Naomi', 'Miss Sherly', 'Tutor', 'Semi Private', 'Mon, Wed 15.45', 8],
    ['Diego', 'Miss Sherly', 'Tutor', 'Semi Private', 'Mon/Tue/Thu/Fri 15.45', 8],
    ['Ayko', 'Miss Sherly', 'Tutor', 'Semi Private', 'Tue/Thu 15.45', 8],
    ['Gabby', 'Miss Sherly', 'Tutor', 'Semi Private', 'Tue/Wed/Fri 15.45', 8],
    ['Grace', 'Miss Sherly', 'Tutor', 'Home Service', '17.30', 8],
    ['Grace', 'Laoshi Lia', 'Mandarin', 'Home Service', 'Wed/Fri 15.30', 8],
    ['Alicia', 'Miss Keren', 'Tutor', 'Home Service', 'Mon/Tue/Thu 17.30', 11],
    ['Wilson', 'Miss Lala', 'Tutor', 'Private', 'Tue/Thu 13.00', 8],
    ['Caellen', 'Miss Lala', 'Tutor', 'Private', 'Mon/Wed/Fri 13.00', 8],
    ['Jason Luisxander', 'Miss Lala', 'Tutor', 'Private', 'Mon/Thu 14.00', 8],
    ['Jason Luisxander', 'Mr Hendra', 'Tutor', 'Private', 'Mon/Wed/Thu', 8],
    ['Alvaro', 'Miss Lala', 'Tutor', 'Semi Private', 'Mon/Thu/Fri 15.30', 8],
    ['Hollie', 'Miss Siska', 'Tutor', 'Semi Private', 'Mon/Thu 17.15', 8],
    ['Agatha', 'Miss Siska', 'Tutor', 'Private', 'Wed 18.15', 8],
    ['Kevin', 'Miss Siska', 'Tutor', 'Semi Private', 'Mon/Wed/Thu 17.15', 8],
    ['Brian', 'Miss Cey', 'Tutor', 'Home Service', 'Mon 15.30 / Thu 13.30', 8],
    ['Rachel', 'Miss Cey', 'Tutor', 'Home Service', 'Mon 16.30 / Thu 14.30', 8],
    ['Aldrich', 'Miss Cey', 'Tutor', 'Home Service', 'Sun 8.00 / Tue 15.30', 8],
    ['Caitlyn', 'Miss Cey', 'Tutor', 'Home Service', 'Wed 15.00 / Sat 10.00', 8],
    ['Cherise', 'Miss Cey', 'Tutor', 'Home Service', 'Wed 16.00 / Sat 11.00', 8],
    ['Celine', 'Miss Cey', 'Tutor', 'Home Service', 'Wed 17.00 / Sat 12.00', 8],
  ];
  // Murid di setup.pdf yang belum punya jadwal kelas
  const extraMurid = ['Salvia', 'Elpandya', 'Thalya Tyanna', 'Kei Oliver Heriyanto', 'Marcellino', 'Clinton Steine Louise', 'Galene'];

  const muridIds = {};
  let nKelas = 0;
  seed.forEach(([nama, guru, program, tipe, jadwal, kuota]) => {
    if (!muridIds[nama]) muridIds[nama] = addMurid({ nama }).id;
    addKelas({ murid_id: muridIds[nama], guru_id: guruIds[guru], program, tipe, jadwal, sesi_kuota: kuota });
    nKelas++;
  });
  extraMurid.forEach(nama => { if (!muridIds[nama]) muridIds[nama] = addMurid({ nama }).id; });

  return { guru: guruNames.length, murid: Object.keys(muridIds).length, kelas: nKelas, note: 'Lengkapi rate invoice murid & fee guru per kelas di tab Admin.' };
}

// ══════════════════════════════════════════════════════════════
//  SETTINGS — key-value pairs di sheet Setting
// ══════════════════════════════════════════════════════════════
function getSettings() {
  const rows = sheetToObjects(getSheet('Setting'));
  const out = {};
  rows.forEach(r => { if (r.key) out[String(r.key)] = String(r.value || ''); });
  return out;
}

function saveSettings(p) {
  const sheet   = getSheet('Setting');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIdx  = headers.indexOf('key');
  const valIdx  = headers.indexOf('value');
  if (keyIdx < 0 || valIdx < 0) throw new Error('Sheet Setting belum disetup — jalankan API.migrate() dahulu');

  const allowed = ['nama', 'motto', 'logo', 'alamat', 'wa', 'ig', 'programs', 'teksInvoice', 'teksLaporan'];
  const results = [];
  for (const key of allowed) {
    if (p[key] === undefined) continue;
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][keyIdx] === key) {
        sheet.getRange(i + 1, valIdx + 1).setValue(p[key]);
        found = true; results.push('updated:' + key); break;
      }
    }
    if (!found) {
      appendByHeaders(sheet, { key, value: p[key] });
      results.push('added:' + key);
    }
  }
  return results;
}
