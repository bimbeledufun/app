// ============================================================
//  AIDA BIMBEL – Client API Helper
//  Ganti SCRIPT_URL dengan URL deployment Apps Script kamu (/exec)
// ============================================================

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx99g9vHa892cH3Cb16tzUaQCbEO6S6AW1f1lkbiEgHJhgANE_TlsmlcpcIYCOn6bsK/exec';

const API = {
  async post(payload) {
    const res  = await fetch(SCRIPT_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify(payload) });
    const json = await res.json();
    if (json.status === 'error') throw new Error(json.message);
    return json.data;
  },
  async get(params) {
    const qs  = new URLSearchParams(params).toString();
    const res = await fetch(SCRIPT_URL + '?' + qs);
    const json = await res.json();
    if (json.status === 'error') throw new Error(json.message);
    return json.data;
  },

  // MURID
  getMurid:       (opts = {}) => API.post({ action: 'getMurid', ...opts }),
  getMuridByLink: (link_id)   => API.post({ action: 'getMuridByLink', link_id }),
  addMurid:       (d)         => API.post({ action: 'addMurid', ...d }),
  updateMurid:    (d)         => API.post({ action: 'updateMurid', ...d }),
  deleteMurid:    (id)        => API.post({ action: 'deleteMurid', id }),

  // KELAS (enrollment murid per program)
  getKelas:        (opts = {}) => API.post({ action: 'getKelas', ...opts }),
  addKelas:        (d)         => API.post({ action: 'addKelas', ...d }),
  updateKelas:     (d)         => API.post({ action: 'updateKelas', ...d }),
  deleteKelas:     (id)        => API.post({ action: 'deleteKelas', id }),
  getSesiAllKelas: ()          => API.post({ action: 'getSesiAllKelas' }),

  // GURU
  getGuru:    ()       => API.post({ action: 'getGuru' }),
  addGuru:    (d)      => API.post({ action: 'addGuru', ...d }),
  updateGuru: (d)      => API.post({ action: 'updateGuru', ...d }),
  deleteGuru: (id)     => API.post({ action: 'deleteGuru', id }),
  getFeeGuru: (bulan)  => API.post({ action: 'getFeeGuru', bulan }),

  // LAPORAN  (materi = [{m:'...', s:1-4}, ...] · status: pending → approved)
  getLaporan:    (opts = {}) => API.post({ action: 'getLaporan', ...opts }),
  addLaporan:    (d)         => API.post({ action: 'addLaporan', ...d }),
  updateLaporan: (d)         => API.post({ action: 'updateLaporan', ...d }),
  deleteLaporan: (id)        => API.post({ action: 'deleteLaporan', id }),

  // LOGIN
  loginOrtu:  (phone)           => API.post({ action: 'loginOrtu', phone }),
  loginGuru:  (phone)           => API.post({ action: 'loginGuru', phone }),
  loginAdmin: (phone, password) => API.post({ action: 'loginAdmin', phone, password }),

  // ABSENSI
  getAbsensi:    (opts = {})      => API.post({ action: 'getAbsensi', ...opts }),
  saveAbsensi:   (tanggal, data)  => API.post({ action: 'saveAbsensi', tanggal, data }),
  upsertAbsensi: (d)              => API.post({ action: 'upsertAbsensi', ...d }),

  // INVOICE
  getInvoice:           (opts = {}) => API.post({ action: 'getInvoice', ...opts }),
  generateInvoice:      (d)         => API.post({ action: 'generateInvoice', ...d }),
  updateInvoiceStatus:  (id, status, tgl_kirim) => API.post({ action: 'updateInvoiceStatus', id, status, tgl_kirim }),
  updateInvoiceNominal: (id, sesi, rate, nominal) => API.post({ action: 'updateInvoiceNominal', id, sesi, rate, nominal }),
  deleteInvoice:        (id)        => API.post({ action: 'deleteInvoice', id }),
  uploadBukti: (invoice_id, murid_id, bulan, file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          resolve(await API.post({ action: 'uploadBukti', invoice_id, murid_id, bulan, base64: e.target.result, filename: file.name }));
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    }),

  // SETUP
  setup: () => API.post({ action: 'setup' }),
  seed:  () => API.post({ action: 'seed' }),
  addColumn: (sheet, kolom, default_val) => API.post({ action: 'addColumn', sheet, kolom, default_val }),
  migrate: () => API.post({ action: 'migrate' }),

  // SETTINGS
  getSettings:  ()  => API.post({ action: 'getSettings' }),
  saveSettings: (d) => API.post({ action: 'saveSettings', ...d }),
};

// ── Util bersama ────────────────────────────────────────────
const LU = {
  PROGRAMS: ['Tutor', 'Swim', 'Art', 'Mandarin'],
  getPrograms() {
    try {
      const s = JSON.parse(localStorage.getItem('aida_settings') || '{}');
      if (s.programs) return s.programs.split(',').map(p => p.trim()).filter(Boolean);
    } catch {}
    return this.PROGRAMS;
  },
  getNamaBimbel() {
    try { return JSON.parse(localStorage.getItem('aida_settings') || '{}').nama || 'Bimbel Edufun'; } catch { return 'Bimbel Edufun'; }
  },
  getBank() {
    try { return JSON.parse(localStorage.getItem('aida_settings') || '{}').bank || this.BANK; } catch { return this.BANK; }
  },
  getTeksInvoice() {
    try { return JSON.parse(localStorage.getItem('aida_settings') || '{}').teksInvoice || ''; } catch { return ''; }
  },
  getTeksLaporan() {
    try { return JSON.parse(localStorage.getItem('aida_settings') || '{}').teksLaporan || ''; } catch { return ''; }
  },
  // Terapkan nama bimbel ke semua elemen yang menampilkan nama bimbel
  applyBrandToPage() {
    const nama = this.getNamaBimbel();
    const el = document.getElementById('siteTitle');
    if (el) el.textContent = nama;
    document.title = document.title.replace(/AIDA BIMBEL/g, nama);
    ['taglineBimbel1','taglineBimbel2','heroSubBimbel','loginHintBimbel','waSender'].forEach(id => {
      const e = document.getElementById(id);
      if (e) e.textContent = nama;
    });
    // Admin login card title
    const loginTitle = document.getElementById('adminLoginTitle');
    if (loginTitle) loginTitle.textContent = `🔐 Login Admin — ${nama}`;
  },
  // Ambil settings dari API, simpan ke localStorage, lalu apply ke halaman
  async loadSettingsFromAPI() {
    try {
      const s = await API.getSettings();
      if (s && typeof s === 'object') {
        const current = JSON.parse(localStorage.getItem('aida_settings') || '{}');
        Object.assign(current, s);
        localStorage.setItem('aida_settings', JSON.stringify(current));
      }
    } catch(e) { /* offline / tidak login — pakai cache localStorage */ }
    this.applyBrandToPage();
  },
  TIPE: ['Private', 'Semi Private', 'Home Service'],
  GRADES: ['Growing', 'Improving', 'Advanced'],
  BANK: 'BCA 2881889996 a.n. Clara E',
  rp: (n) => 'Rp ' + (Number(n) || 0).toLocaleString('id-ID'),
  bulanID: (ym) => {
    if (!ym) return '—';
    const [y, m] = ym.split('-').map(Number);
    return new Date(y, m - 1, 1).toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
  },
  fmtDate: (s) => {
    if (!s) return '–';
    try { return new Date(s).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' }); } catch { return s; }
  },
  fmtDateDDMM: (s) => {
    if (!s) return '–';
    const [y, m, d] = String(s).split('-');
    return `${d}/${m}/${y}`;
  },
  parseMateri: (json) => {
    try { const a = JSON.parse(json || '[]'); return Array.isArray(a) ? a : []; } catch { return []; }
  },
  stars: (n) => '★'.repeat(Number(n) || 0) + '☆'.repeat(Math.max(0, 4 - (Number(n) || 0))),
  // Substitusi variabel {key} dalam template
  applyTemplate(tpl, vars) {
    return tpl.replace(/\{(\w+)\}/g, (_, key) => vars[key] !== undefined ? String(vars[key]) : '{' + key + '}');
  },
  // Template WA invoice — support full custom template atau default
  waInvoice: ({ bulan, nama, program, sesi, rate, total }) => {
    const bimbel   = LU.getNamaBimbel();
    const bank     = LU.getBank();
    const template = LU.getTeksInvoice();
    const vars = {
      bimbel, bulan, nama: nama || '—', program: program || '—',
      sesi, rate: LU.rp(rate), total: LU.rp(total), rekening: bank,
    };
    if (template) return LU.applyTemplate(template, vars);
    // Default template
    return `Dear Parents,
Berikut kami informasikan billing bulan ${bulan}

Nama Anak: ${nama}
Program: ${program}
Jumlah Sesi: ${sesi}
💰 Rate: ${LU.rp(rate)}/ sesi
💰 Total: ${LU.rp(total)}

Pembayaran dapat dilakukan melalui:
${bank}

Mohon konfirmasi setelah melakukan pembayaran ya 🙏
Terima kasih banyak atas kepercayaannya 😊
— ${bimbel}`;
  },
  // Template WA laporan harian — support full custom template atau default
  waLaporan: ({ tanggal, nama, program, tipe, materi, catatan, extra }) => {
    const bimbel = LU.getNamaBimbel();
    const tglFmt = tanggal
      ? new Date(tanggal).toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })
      : '—';
    const materiTxt = (materi && materi.length)
      ? materi.map(x => `• ${x.m} ${LU.stars(x.s)}`).join('\n')
      : '—';
    let extraObj = {};
    try { extraObj = (typeof extra === 'string' ? JSON.parse(extra) : extra) || {}; } catch {}
    const template = LU.getTeksLaporan();
    if (template) {
      return LU.applyTemplate(template, {
        bimbel, tanggal: tglFmt, nama: nama || '—',
        program: program || '—', tipe: tipe || '',
        materi: materiTxt,
        rincian:  extraObj.rincian  || '',
        prestasi: extraObj.prestasi || '',
        kendala:  extraObj.kendala  || '',
        catatan:  catatan || '—',
      });
    }
    // Default template
    const rincianTxt  = extraObj.rincian  ? `\n\n📖 Rincian Materi:\n${extraObj.rincian}`  : '';
    const prestasiTxt = extraObj.prestasi ? `\n\n🏆 Prestasi Siswa:\n${extraObj.prestasi}` : '';
    const kendalaTxt  = extraObj.kendala  ? `\n\n⚡ Kendala & Evaluasi:\n${extraObj.kendala}` : '';
    return `Halo parents! 👋

📋 Laporan Belajar — ${bimbel}
${tglFmt}
Nama: ${nama || '—'}
Program: ${program || '—'}${tipe ? ' (' + tipe + ')' : ''}

Materi hari ini:
${materiTxt}${rincianTxt}${prestasiTxt}${kendalaTxt}

📝 Catatan Tutor:
${catatan || '—'}

Terima kasih atas kepercayaannya 🙏
— ${bimbel}`;
  },
};
