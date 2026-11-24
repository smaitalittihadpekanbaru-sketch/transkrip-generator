/* =============================================================================
 *  app.js — Frontend GitHub Pages (dibangun otomatis dari JavaScript.html)
 *  JANGAN diedit langsung: ubah JavaScript.html lalu jalankan tools/build_github_pages.py
 *  - Unggah Excel (SheetJS) & DOCX (JSZip) → JSON → API Apps Script (apiImport)
 *  - Data siswa & riwayat transkrip dari Google Sheets
 *  - Bahasa (3 Bahasa / ID / EN / AR) + format (GPA / Nilai Rapor)
 *  - Generate, unduh, edit di Google Docs, perbarui PDF/DOCX, edit data & nilai
 * ========================================================================== */
'use strict';

/* ---------------------------------------------------------------------------
 * 1. STATE & UTILITAS
 * ------------------------------------------------------------------------- */
const state = {
  students: [], subjects: [], transcripts: [], modes: [], jurusan: [], kurikulum: [],
  selected: new Set(), parsed: null, tab: 'students',
  preview: { nis: null, mode: 'TRI', format: 'TRI' },
  editor: null, busy: false
};

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const isAr = (s) => /[؀-ۿ]/.test(String(s || ''));
const getMode = () => (document.querySelector('input[name="mode"]:checked') || {}).value || 'TRI';
const getJurusan = () => ($('jurusanSel') ? $('jurusanSel').value : '');
const jurusanName = (code) => { const j = state.jurusan.find((x) => x.Kode === code); return j ? j.Nama : (code || ''); };
const getFormat = (mode) => ((mode || getMode()) === 'TRI' ? 'TRI' : ((document.querySelector('input[name="format"]:checked') || {}).value || 'GPA'));
const FORMAT_LABEL = { TRI: '3 Bahasa', GPA: 'GPA', RAPOR: 'Nilai Rapor' };
const MODE_LABEL = { TRI: '3 Bahasa', ID: 'Indonesia', EN: 'English', AR: 'Arab' };
const comboLabel = (mode, format) => (format === 'TRI' ? '3 Bahasa' : `${MODE_LABEL[mode] || mode} · ${FORMAT_LABEL[format] || format}`);

/* ---- Koneksi ke Google Apps Script (GitHub Pages) -------------------------
 * Pengganti RPC bawaan GAS: POST JSON ke /exec. Body dikirim sebagai
 * text/plain agar browser tidak mengirim preflight CORS (GAS tidak mendukung OPTIONS).
 * ------------------------------------------------------------------------- */
const API = {
  URL_KEY: 'transkrip_gas_url',
  get url() {
    let u = '';
    try { u = localStorage.getItem(this.URL_KEY) || ''; } catch (e) { u = this._url || ''; }
    return String(u || (window.APP_CONFIG && window.APP_CONFIG.GAS_URL) || '').trim();
  },
  setUrl(u) {
    this._url = u;
    try { if (u) localStorage.setItem(this.URL_KEY, u); else localStorage.removeItem(this.URL_KEY); } catch (e) { /* memori saja */ }
  },
  KEY: 'transkrip_kode_akses',
  get token() {
    try { return localStorage.getItem(this.KEY) || sessionStorage.getItem(this.KEY) || ''; } catch (e) { return this._mem || ''; }
  },
  setToken(t, remember) {
    this._mem = t;
    try {
      localStorage.removeItem(this.KEY); sessionStorage.removeItem(this.KEY);
      if (t) (remember ? localStorage : sessionStorage).setItem(this.KEY, t);
    } catch (e) { /* penyimpanan browser tidak tersedia: simpan di memori saja */ }
  }
};

class ApiError extends Error { constructor(msg, code) { super(msg); this.code = code || 'ERROR'; } }

async function rawCall(action, args, token) {
  if (!API.url || !/^https:\/\/script\.google(usercontent)?\.com\//.test(API.url)) {
    throw new ApiError('GAS_URL di config.js belum diisi dengan URL Web App (/exec).', 'CONFIG');
  }
  let res;
  try {
    res = await fetch(API.url, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, args, token })
    });
  } catch (e) {
    throw new ApiError('Tidak dapat terhubung ke server Apps Script. Klik "Tes koneksi" di bawah: jika yang muncul bukan teks {"ok":true,…,"version":"2.2.0"}, berarti URL ini belum menjalankan Code.gs versi baru (Deploy → Manage deployments → Edit → Version: New version), atau aksesnya bukan "Siapa saja" (Anyone).', 'NETWORK');
  }
  const text = await res.text();
  let out;
  try { out = JSON.parse(text); } catch (e) {
    const login = /accounts\.google\.com|ServiceLogin|<html/i.test(text);
    throw new ApiError(login
      ? 'Server meminta login Google. Deploy ulang Web App dengan "Who has access: Anyone" (Siapa saja) lalu pakai URL /exec yang baru.'
      : 'Balasan server tidak valid (HTTP ' + res.status + ').', 'NETWORK');
  }
  if (!out.ok) throw new ApiError(out.error || 'Terjadi kesalahan di server.', out.code);
  return out.data;
}

async function call(fn, ...args) {
  try {
    return await rawCall(fn, args, API.token);
  } catch (e) {
    if (e.code === 'AUTH') { API.setToken(''); showLogin(e.message); }
    throw e;
  }
}
async function withLoading(text, fn, sub) { showLoading(text, sub); try { return await fn(); } finally { hideLoading(); } }
function showLoading(text, sub) { $('loadingText').textContent = text || 'Memproses…'; $('loadingSub').textContent = sub || ''; $('loading').classList.remove('hidden'); }
function hideLoading() { $('loading').classList.add('hidden'); }
function toast(msg, type = 'ok', ms = 4500) {
  const colors = { ok: 'bg-brand-700 text-white', err: 'bg-red-600 text-white', warn: 'bg-amber-500 text-white', info: 'bg-slate-800 text-white' };
  const el = document.createElement('div');
  el.className = 'toast rounded-xl shadow-lg px-4 py-3 text-sm font-medium whitespace-pre-line ' + (colors[type] || colors.info);
  el.textContent = msg;
  $('toasts').appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity .3s'; setTimeout(() => el.remove(), 300); }, ms);
}
const errMsg = (e) => (e && e.message ? e.message : String(e)).replace(/^Exception:\s*/, '');
function openModal(id) { $(id).classList.remove('hidden'); document.body.style.overflow = 'hidden'; }
function closeModal(id) { $(id).classList.add('hidden'); if (!document.querySelector('.modal-bg:not(.hidden):not(#loading)')) document.body.style.overflow = ''; }
function confirmBox(title, text, yesLabel = 'Ya, lanjutkan') {
  return new Promise((resolve) => {
    $('confirmTitle').textContent = title; $('confirmText').textContent = text; $('confirmYes').textContent = yesLabel;
    openModal('modalConfirm');
    const done = (v) => { closeModal('modalConfirm'); $('confirmYes').onclick = null; $('confirmNo').onclick = null; resolve(v); };
    $('confirmYes').onclick = () => done(true);
    $('confirmNo').onclick = () => done(false);
  });
}
function saveBase64(b64, mime, name) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  const a = document.createElement('a');
  a.href = url; a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
}
const fmtTime = (iso) => { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? String(iso) : d.toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' }); };

/* ---------------------------------------------------------------------------
 * 2. PARSER EXCEL & DOCX (fungsi murni)
 * ------------------------------------------------------------------------- */
// === PARSER START ===
const ExcelParser = (function () {
  const ROMAN = { I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6 };
  const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  const keyOf = (s) => norm(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const rowsOf = (ws) => XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const isNum = (v) => v !== '' && v !== null && !isNaN(Number(v));

  /** "Matematika/Mathematics/الرياضيات" → {id, en, ar, raw} */
  function splitMapel(name) {
    const raw = norm(name);
    const parts = raw.split('/').map(norm).filter(Boolean);
    if (parts.length >= 3) return { raw, id: parts[0], en: parts.slice(1, -1).join('/'), ar: parts[parts.length - 1] };
    if (parts.length === 2) {
      const arabic = /[؀-ۿ]/.test(parts[1]);
      return { raw, id: parts[0], en: arabic ? '' : parts[1], ar: arabic ? parts[1] : '' };
    }
    return { raw, id: raw, en: '', ar: '' };
  }

  function toIsoDate(v) {
    if (v === '' || v == null) return '';
    if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
    if (typeof v === 'number' && XLSX.SSF) {
      const d = XLSX.SSF.parse_date_code(v);
      if (d) return d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0');
    }
    return norm(v);
  }

  /** Cari baris judul blok semester: sel "SEMESTER I/II/…" + sel "MATA PELAJARAN". */
  function findSemesterBlocks(rows) {
    const limit = Math.min(rows.length, 25);
    for (let r = 0; r < limit; r++) {
      const row = rows[r] || [];
      const blocks = [];
      let subjCol = -1;
      row.forEach((cell, c) => {
        const t = norm(cell);
        if (/^(mata\s*pelajaran|subject)/i.test(t)) subjCol = c;
        const m = /^SEMESTER\s+(VI|IV|V|I{1,3}|[1-6])\b/i.exec(t);
        if (m) {
          const sem = ROMAN[m[1].toUpperCase()] || Number(m[1]);
          const y = /(\d{4})\s*\/\s*(\d{4})/.exec(t);
          blocks.push({ sem, col: c, tahun: y ? y[1] + '/' + y[2] : '' });
        }
      });
      if (blocks.length && subjCol >= 0) return { headerRow: r, subjCol, blocks };
    }
    return null;
  }

  function findIdentity(rows) {
    const id = {};
    const valueRight = (row, c) => { for (let j = c + 1; j < row.length; j++) if (norm(row[j]) !== '') return row[j]; return ''; };
    rows.forEach((row) => {
      (row || []).forEach((cell, c) => {
        const t = norm(cell).toLowerCase();
        if (!t || t.length > 60) return;
        if (/^nama siswa/.test(t) && !id.Nama_Siswa) id.Nama_Siswa = norm(valueRight(row, c));
        else if (/^nis\s*\/\s*nisn/.test(t)) {
          const parts = norm(valueRight(row, c)).split('/').map(norm);
          id.NIS = parts[0] || ''; id.NISN = parts[1] || '';
        } else if (/^nisn\b/.test(t) && !id.NISN) id.NISN = norm(valueRight(row, c));
        else if (/^nis\b/.test(t) && !id.NIS) id.NIS = norm(valueRight(row, c));
        else if (/^masa studi/.test(t)) id.Masa_Studi = norm(valueRight(row, c));
        else if (/^jenis kelamin/.test(t)) id.Jenis_Kelamin = norm(valueRight(row, c));
        else if (/^nama (orang tua|ortu|wali)/.test(t)) id.Nama_Ortu = norm(valueRight(row, c));
        else if (/^nama.*arab/.test(t)) id.Nama_Arab = norm(valueRight(row, c));
        else if (/^(nomor|no\.?) transkrip/.test(t)) id.Nomor_Transkrip = norm(valueRight(row, c));
        else if (/^tanggal terbit/.test(t)) id.Tanggal_Terbit = toIsoDate(valueRight(row, c));
      });
    });
    return id;
  }

  /** FORMAT A — "Kalkulator Transkrip" (satu file = satu siswa). */
  function parseKalkulator(wb) {
    const warnings = [];
    const semSheets = [];
    wb.SheetNames.forEach((name) => {
      const rows = rowsOf(wb.Sheets[name]);
      const info = findSemesterBlocks(rows);
      if (info) semSheets.push({ name, rows, info });
    });
    // Identitas: utamakan sheet REKAP, lalu semua sheet
    let identity = {};
    const rekapName = wb.SheetNames.find((n) => /rekap/i.test(n));
    if (rekapName) identity = findIdentity(rowsOf(wb.Sheets[rekapName]));
    if (!identity.NIS) wb.SheetNames.forEach((n) => { if (!identity.NIS) identity = Object.assign(findIdentity(rowsOf(wb.Sheets[n])), identity); });
    if (!identity.NIS) throw new Error('NIS tidak ditemukan. Pastikan sheet REKAP berisi baris "NIS / NISN".');
    if (!identity.Nama_Siswa) warnings.push('Nama siswa tidak ditemukan di sheet REKAP.');

    const grades = [];
    semSheets.forEach(({ name, rows, info }) => {
      const sub = rows[info.headerRow + 1] || [];
      info.blocks.forEach((b) => {
        b.nilaiCol = -1; b.sksCol = -1;
        for (let c = b.col; c < b.col + 7 && c < sub.length; c++) {
          const t = norm(sub[c]).toUpperCase();
          if (b.nilaiCol < 0 && /^(NILAI|REPORT|SCORE)\b/.test(t)) b.nilaiCol = c;
          if (b.sksCol < 0 && /^(SKS|CREDITS?)\b/.test(t)) b.sksCol = c;
        }
        if (b.nilaiCol < 0 || b.sksCol < 0) warnings.push(`Sheet ${name}: kolom NILAI/SKS semester ${b.sem} tidak ditemukan.`);
      });
      let urutan = 0, started = false;
      for (let r = info.headerRow + 2; r < rows.length; r++) {
        const subj = norm((rows[r] || [])[info.subjCol]);
        if (!subj) { if (started) break; continue; }
        if (/^(jumlah|indeks|total|pemeriksaan)/i.test(subj)) break;
        started = true; urutan++;
        info.blocks.forEach((b) => {
          if (b.nilaiCol < 0) return;
          const nilai = rows[r][b.nilaiCol];
          const sks = b.sksCol >= 0 ? rows[r][b.sksCol] : '';
          if (nilai === '' || nilai == null) return;
          if (!isNum(nilai)) { warnings.push(`Sheet ${name} baris ${r + 1}: nilai "${nilai}" bukan angka.`); return; }
          if (!isNum(sks)) warnings.push(`Sem ${b.sem} · ${subj.split('/')[0]}: SKS kosong.`);
          grades.push({ NIS: identity.NIS, Kategori: 'GPA', Semester: b.sem, Tahun_Pelajaran: b.tahun, Urutan: urutan, Nilai: Number(nilai), SKS: isNum(sks) ? Number(sks) : '', Mapel: splitMapel(subj) });
        });
      }
    });
    if (!grades.length) throw new Error('Tidak ada nilai yang terbaca. Pastikan sheet semester (SEM 1-2, SEM 3-4, …) terisi.');
    return { source: 'KALKULATOR', students: [identity], grades, warnings };
  }

  /** Ubah sheet berheader menjadi array objek dengan alias kolom. */
  function tableOf(ws, aliases) {
    const rows = rowsOf(ws);
    const hIdx = rows.findIndex((r) => r.filter((c) => norm(c) !== '').length >= 2);
    if (hIdx < 0) return [];
    const map = {};
    rows[hIdx].forEach((h, c) => {
      const k = keyOf(h);
      Object.keys(aliases).forEach((field) => { if (map[field] === undefined && aliases[field].includes(k)) map[field] = c; });
    });
    return rows.slice(hIdx + 1)
      .filter((r) => r.some((c) => norm(c) !== ''))
      .map((r, i) => { const o = { _line: hIdx + i + 2 }; Object.keys(map).forEach((f) => { o[f] = r[map[f]]; }); return o; });
  }

  /** FORMAT B — "Format Massal" (sheet SISWA + sheet NILAI, banyak siswa). */
  function parseMassal(wb) {
    const warnings = [];
    const findSheet = (re) => wb.SheetNames.find((n) => re.test(n.trim()));
    const sSiswa = findSheet(/^(data[_ ]?)?siswa$/i), sNilai = findSheet(/^(data[_ ]?)?nilai$/i);
    const students = tableOf(wb.Sheets[sSiswa], {
      NIS: ['nis'], NISN: ['nisn'], Nama_Siswa: ['namasiswa', 'nama', 'namalengkap'],
      Nama_Arab: ['namaarab', 'namaarabic', 'namadalamhurufarab', 'namaar'],
      Jenis_Kelamin: ['jeniskelamin', 'jk', 'lp', 'gender', 'sex'],
      Nama_Ortu: ['namaortu', 'namaorangtua', 'orangtua', 'namawali'],
      Masa_Studi: ['masastudi'], Nomor_Transkrip: ['nomortranskrip', 'notranskrip', 'nomor'],
      Tanggal_Terbit: ['tanggalterbit', 'tglterbit', 'tanggal'], Tempat_Lahir: ['tempatlahir'],
      Tanggal_Lahir: ['tanggallahir', 'tgllahir'], Nomor_Rapor: ['nomorrapor', 'nomortranskriprapor'], Tanggal_Rapor: ['tanggalrapor', 'tglrapor'],
      Jurusan: ['jurusan', 'peminatan', 'program', 'kodejurusan']
    }).map((s) => {
      const o = {};
      ['NIS', 'NISN', 'Nama_Siswa', 'Nama_Arab', 'Jenis_Kelamin', 'Jurusan', 'Tempat_Lahir', 'Nama_Ortu', 'Masa_Studi', 'Nomor_Transkrip', 'Nomor_Rapor'].forEach((k) => { if (s[k] !== undefined) o[k] = norm(s[k]); });
      ['Tanggal_Terbit', 'Tanggal_Lahir', 'Tanggal_Rapor'].forEach((k) => { if (s[k] !== undefined) o[k] = toIsoDate(s[k]); });
      if (!o.NIS) warnings.push(`Sheet ${sSiswa} baris ${s._line}: NIS kosong.`);
      return o;
    }).filter((s) => s.NIS);

    const grades = [];
    tableOf(wb.Sheets[sNilai], {
      NIS: ['nis'], Semester: ['semester', 'sem'], Tahun_Pelajaran: ['tahunpelajaran', 'tahunajaran', 'tahun', 'tp'],
      Mapel: ['matapelajaran', 'mapel', 'namamapel', 'kodemapel', 'kode'], Mapel_EN: ['mapelen', 'subject', 'namaen'],
      Mapel_AR: ['mapelar', 'namaar', 'namaarab'], Nilai: ['nilai', 'score', 'nilaiakhir'], SKS: ['sks', 'credits', 'credit'],
      Urutan: ['urutan', 'no', 'nourut'], Kategori: ['kategori', 'jenis', 'jenisnilai'], Kelompok: ['kelompok', 'kelompokmapel', 'group']
    }).forEach((g) => {
      const nis = norm(g.NIS), mapel = norm(g.Mapel);
      if (!nis || !mapel) { warnings.push(`Sheet ${sNilai} baris ${g._line}: NIS/mapel kosong.`); return; }
      if (g.Nilai === '' || g.Nilai == null) return;
      if (!isNum(g.Nilai)) { warnings.push(`Sheet ${sNilai} baris ${g._line}: nilai "${g.Nilai}" bukan angka.`); return; }
      let sem = norm(g.Semester).toUpperCase().replace(/^SEMESTER\s*/, '');
      sem = ROMAN[sem] || Number(sem);
      const m = mapel.includes('/') ? splitMapel(mapel)
        : { raw: mapel, id: mapel, en: norm(g.Mapel_EN), ar: norm(g.Mapel_AR), kode: /^[A-Z0-9_]{2,15}$/.test(mapel) ? mapel : '' };
      const kat = norm(g.Kategori).toUpperCase() === 'RAPOR' ? 'RAPOR' : 'GPA';
      m.kelompok = norm(g.Kelompok);
      if (kat === 'GPA' && !isNum(g.SKS)) warnings.push(`NIS ${nis} · Sem ${sem} · ${m.id}: SKS kosong.`);
      grades.push({ NIS: nis, Kategori: kat, Semester: sem, Tahun_Pelajaran: norm(g.Tahun_Pelajaran), Urutan: isNum(g.Urutan) ? Number(g.Urutan) : '', Nilai: Number(g.Nilai), SKS: isNum(g.SKS) ? Number(g.SKS) : '', Mapel: m });
    });
    if (!students.length) throw new Error('Sheet SISWA kosong atau kolom NIS tidak ditemukan.');
    return { source: 'MASSAL', students, grades, warnings };
  }

  /** FORMAT C — "Struktur Kurikulum" (sheet KELAS X / XI / XII). */
  function isKurikulum(wb) {
    return wb.SheetNames.some((n) => rowsOf(wb.Sheets[n]).slice(0, 6).some((r) => r.some((c) => /STRUKTUR\s+KURIKULUM\s+KELAS/i.test(norm(c)))));
  }

  function parseKurikulum(wb, fileName) {
    const entries = [], warnings = [];
    wb.SheetNames.forEach((name) => {
      const rows = rowsOf(wb.Sheets[name]);
      let kelas = null, start = -1;
      for (let r = 0; r < Math.min(rows.length, 6) && !kelas; r++) {
        for (const c of rows[r]) {
          const m = /STRUKTUR\s+KURIKULUM\s+KELAS\s+(XII|XI|X)\b/i.exec(norm(c));
          if (m) { kelas = m[1].toUpperCase(); start = r + 1; break; }
        }
      }
      if (!kelas) return;
      let kelompok = '', pemCols = null, jpCol = -1, nameCol = 1;
      for (let r = start; r < rows.length; r++) {
        const row = rows[r] || [];
        const c0 = norm(row[0]);
        if (/^tambahan/i.test(c0) || /^jumlah\s+total/i.test(c0)) break;
        if (/^kelompok/i.test(c0)) { kelompok = c0; pemCols = null; continue; }
        if (/^jumlah$/i.test(c0)) { pemCols = null; continue; }
        if (/^no\.?$/i.test(c0)) {
          const jp = row.findIndex((x, i) => i > 0 && /^jp$/i.test(norm(x)));
          if (jp > 0) jpCol = jp;
          const names = [];
          for (let i = 1; i < (jp > 0 ? jp : row.length); i++) { const t = norm(row[i]); if (t) names.push({ col: i, name: t }); }
          if (names.length >= 2) pemCols = names; else { pemCols = null; if (names[0]) nameCol = names[0].col; }
          continue;
        }
        if (!/^\d+$/.test(c0)) continue;
        const urutan = Number(c0);
        const jp = jpCol > 0 && isNum(row[jpCol]) ? Number(row[jpCol]) : '';
        if (pemCols) pemCols.forEach((pc) => { const n = norm(row[pc.col]); if (n) entries.push({ kelas, peminatan: pc.name, kelompok, urutan, nama: n, jp }); });
        else { const n = norm(row[nameCol]); if (n) entries.push({ kelas, peminatan: '', kelompok, urutan, nama: n, jp }); }
      }
    });
    if (!entries.length) throw new Error('Struktur kurikulum tidak terbaca (cari sheet "STRUKTUR KURIKULUM KELAS X/XI/XII").');
    ['X', 'XI', 'XII'].forEach((k) => { if (!entries.some((e) => e.kelas === k)) warnings.push('Kelas ' + k + ' tidak ditemukan di file.'); });
    const y = /(\d{4})\s*[-_\/]\s*(\d{4})/.exec(fileName || '');
    return { source: 'KURIKULUM', kurikulum: { nama: y ? y[1] + '/' + y[2] : '', entries }, students: [], grades: [], warnings, summary: [] };
  }

  /* ---- FORMAT D — "Ekspor Transkrip" (.xls berisi tabel HTML) ------------------
   * Baris identitas Field | Value (Student, Class, Stream, Grade 10/11/12, Cumulative GPA …),
   * lalu blok "Semester I — 2024/2025 — …" + header Group | Subject name | Report | Grade |
   * Grade point | Credits | Score, ditutup TOTAL / SEMESTER GPA. NIS tidak ada di file. */
  const SEM_TITLE = /^semester\s+(VI|IV|V|I{1,3}|[1-6])\b(.*)$/i;
  const numOf = (v) => { const t = norm(v).replace(',', '.'); return t !== '' && !isNaN(Number(t)) ? Number(t) : null; };
  const filledOf = (row) => (row || []).map(norm).filter(Boolean);
  const isSubjectKey = (k) => /^(subjectname|subject|subjects|matapelajaran|mapel|namamapel)$/.test(k);
  const isCreditKey = (k) => /^(credits?|sks)$/.test(k);

  function eksporBlocks(rows) {
    const out = [];
    for (let r = 0; r < rows.length - 1; r++) {
      const f = filledOf(rows[r]);
      if (f.length !== 1 || !SEM_TITLE.test(f[0])) continue;
      const head = (rows[r + 1] || []).map(keyOf);
      if (head.some(isSubjectKey) && head.some(isCreditKey)) out.push(r);
    }
    return out;
  }
  const isEkspor = (rows) => eksporBlocks(rows).length > 0;

  function parseEkspor(rows) {
    const warnings = [];
    const info = {};
    const student = { NIS: '' };
    const years = {};
    // ---- identitas (sebelum blok semester pertama)
    const blocks = eksporBlocks(rows);
    for (let r = 0; r < blocks[0]; r++) {
      const f = filledOf(rows[r]);
      if (f.length === 1) {
        const m = /^(academic\s+transcript|transkrip(\s+nilai|\s+akademik)?)\s*[—–-]\s*(.+)$/i.exec(f[0]);
        if (m) info.judul = norm(m[3]);
        continue;
      }
      if (f.length < 2) continue;
      const k = keyOf(f[0]), v = f[1];
      if (/^(field|kolom)$/.test(k)) continue;
      if (/^(student|studentname|name|nama|namasiswa|namalengkap)$/.test(k)) student.Nama_Siswa = v;
      else if (/^(nisnisn|nisn?nis)$/.test(k)) { const pr = v.split('/').map(norm); student.NIS = pr[0] || ''; student.NISN = pr[1] || ''; }
      else if (/^(nis|studentid|studentnumber|nomorinduk|noinduk)$/.test(k)) student.NIS = v;
      else if (/^nisn$/.test(k)) student.NISN = v;
      else if (/^(gender|sex|jeniskelamin|jk)$/.test(k)) student.Jenis_Kelamin = v;
      else if (/^(placeofbirth|tempatlahir)$/.test(k)) student.Tempat_Lahir = v;
      else if (/^(dateofbirth|tanggallahir)$/.test(k)) student.Tanggal_Lahir = v;
      else if (/^(parent|parentname|parents|namaortu|namaorangtua|namawali)/.test(k)) student.Nama_Ortu = v;
      else if (/^(class|kelas)$/.test(k)) info.kelas = v;
      else if (/^(stream|jurusan|peminatan|program)/.test(k)) info.stream = v;
      else if (/^(grade|kelas)(10|11|12|x|xi|xii)$/.test(k)) {
        const g = { x: 10, xi: 11, xii: 12 }[k.replace(/^(grade|kelas)/, '')] || Number(k.replace(/^(grade|kelas)/, ''));
        const y = /(\d{4})\s*\/\s*(\d{4})/.exec(v); if (y) years[g] = y[1] + '/' + y[2];
      }
      else if (/^(cumulativegpa|gpa|ipk)$/.test(k)) info.ipk = numOf(v);
      else if (/^(totalcredits|totalsks|jumlahsks)$/.test(k)) info.sks = numOf(v);
      else if (/^(qualitypoints|totalscore|jumlahmutu|totalmutu)$/.test(k)) info.mutu = numOf(v);
      else if (/^(academicstanding|predikat|standing)$/.test(k)) info.predikat = v;
    }
    if (!student.Nama_Siswa && info.judul) student.Nama_Siswa = info.judul;
    if (!student.Nama_Siswa) throw new Error('Nama siswa tidak ditemukan (baris "Student").');
    student.Nama_Siswa = norm(student.Nama_Siswa).toUpperCase();

    // ---- nilai per semester
    const grades = [];
    const semTahun = {};
    blocks.forEach((r0) => {
      const title = filledOf(rows[r0])[0];
      const m = SEM_TITLE.exec(title);
      const sem = ROMAN[m[1].toUpperCase()] || Number(m[1]);
      const y = /(\d{4})\s*\/\s*(\d{4})/.exec(title);
      const tahun = y ? y[1] + '/' + y[2] : (years[sem <= 2 ? 10 : sem <= 4 ? 11 : 12] || '');
      semTahun[sem] = tahun;
      const head = (rows[r0 + 1] || []).map(keyOf);
      const col = (re) => head.findIndex((k) => re.test(k));
      const cGroup = col(/^(group|kelompok|kelompokmapel)$/), cSubj = head.findIndex(isSubjectKey);
      const cNilai = col(/^(report|nilai|mark|marks|reportscore)/), cSks = head.findIndex(isCreditKey);
      const cScore = col(/^(score|mutu|qualitypoints?)$/);
      if (cNilai < 0) { warnings.push(`Semester ${sem}: kolom nilai (Report) tidak ditemukan.`); return; }
      let urutan = 0, sksFile = null, sksSum = 0, mutuSum = 0;
      for (let r = r0 + 2; r < rows.length; r++) {
        const row = rows[r] || [];
        const f = filledOf(row);
        if (!f.length || (f.length === 1 && (SEM_TITLE.test(f[0]) || /^grade\s*scale|^skala/i.test(f[0])))) break;
        const subj = norm(row[cSubj]);
        if (/^(total|jumlah)\b/i.test(subj)) { sksFile = numOf(row[cSks]); continue; }
        if (/^(semester\s*gpa|ips|indeks)/i.test(subj)) break;
        if (!subj) continue;
        urutan++;
        const nilai = numOf(row[cNilai]);
        if (nilai === null) continue; // belum ada nilai (semester berjalan / mapel tanpa nilai)
        const sks = numOf(row[cSks]);
        const group = cGroup >= 0 ? norm(row[cGroup]) : '';
        const mapel = { raw: subj, id: '', en: subj, ar: '', kelompok: group };
        // "Informatics (Coding/AI)" di luar kelompok umum = TIK (Koding/KA); di kelompok umum = Informatika
        if (/^informati(cs|ka)\b/i.test(subj) && group && !/general|umum|wajib/i.test(group)) mapel.kode = 'TIK';
        if (sks === null || sks <= 0) warnings.push(`Sem ${sem} · ${subj}: tanpa SKS (sesuai file, tidak dihitung ke IPS).`);
        else { sksSum += sks; if (cScore >= 0) mutuSum += numOf(row[cScore]) || 0; }
        grades.push({ NIS: student.NIS, Kategori: 'GPA', Semester: sem, Tahun_Pelajaran: tahun, Urutan: urutan, Nilai: nilai, SKS: sks !== null && sks > 0 ? sks : '', Mapel: mapel });
      }
      if (sksFile !== null && sksFile !== sksSum) warnings.push(`Semester ${sem}: total SKS di file ${sksFile}, terbaca ${sksSum}.`);
    });
    if (!grades.length) throw new Error('Tidak ada nilai yang terbaca dari blok semester.');

    const ys = [10, 11, 12].map((g) => years[g]).filter(Boolean);
    const tAll = Object.keys(semTahun).sort().map((k) => semTahun[k]).filter(Boolean);
    const first = ys[0] || tAll[0], last = years[12] || (first ? (Number(first.slice(0, 4)) + 2) + '/' + (Number(first.slice(0, 4)) + 3) : '');
    if (first && last) student.Masa_Studi = first + ' – ' + last;
    const sksTotal = grades.reduce((a, g) => a + (Number(g.SKS) || 0), 0);
    if (info.sks != null && info.sks !== sksTotal) warnings.push(`Total SKS di file ${info.sks}, terbaca ${sksTotal}.`);
    return { source: 'EKSPOR', students: [student], grades, warnings, info };
  }

  /** Tabel HTML (file .xls hasil ekspor web) → array baris; colspan diratakan. */
  function rowsFromHtml(text, DOMParserImpl) {
    const doc = new (DOMParserImpl || DOMParser)().parseFromString(String(text || ''), 'text/html');
    const rows = [];
    Array.from(doc.getElementsByTagName('tr')).forEach((tr) => {
      const row = [];
      Array.from(tr.children).forEach((cell) => {
        if (!/^t[dh]$/i.test(cell.tagName)) return;
        const t = String(cell.textContent || '').replace(/[\s ]+/g, ' ').trim();
        row.push(/^-?\d+(\.\d+)?$/.test(t) ? Number(t) : t);
        const span = Math.min(Number(cell.getAttribute('colspan')) || 1, 50);
        for (let i = 1; i < span; i++) row.push('');
      });
      rows.push(row);
    });
    return rows;
  }

  /** Workbook dari tabel HTML agar bisa diproses parser biasa. */
  function workbookFromHtml(text, DOMParserImpl) {
    const ws = XLSX.utils.aoa_to_sheet(rowsFromHtml(text, DOMParserImpl));
    return { SheetNames: ['Sheet1'], Sheets: { Sheet1: ws } };
  }

  function detect(wb) {
    if (isKurikulum(wb)) return 'KURIKULUM';
    if (wb.SheetNames.some((n) => isEkspor(rowsOf(wb.Sheets[n])))) return 'EKSPOR';
    const names = wb.SheetNames.map((n) => n.trim().toLowerCase());
    if (names.some((n) => /^(data[_ ]?)?siswa$/.test(n)) && names.some((n) => /^(data[_ ]?)?nilai$/.test(n))) return 'MASSAL';
    if (wb.SheetNames.some((n) => findSemesterBlocks(rowsOf(wb.Sheets[n])))) return 'KALKULATOR';
    return null;
  }

  /** Ringkasan per siswa untuk ditampilkan sebelum disimpan. */
  function summarize(result) {
    return result.students.map((s) => {
      const g = result.grades.filter((x) => String(x.NIS) === String(s.NIS));
      const sems = [...new Set(g.map((x) => x.Semester))].sort();
      const rapor = g.length && g.every((x) => x.Kategori === 'RAPOR');
      return { NIS: s.NIS, Nama_Siswa: s.Nama_Siswa || '(tanpa nama)', semesters: sems, count: g.length, rapor, sks: g.reduce((a, x) => a + (Number(x.SKS) || 0), 0) };
    });
  }

  function parse(wb, fileName) {
    const type = detect(wb);
    if (!type) throw new Error('Format file tidak dikenali. Gunakan Kalkulator Transkrip (SEM 1-2 … + REKAP), Format Massal (SISWA + NILAI), Ekspor Transkrip (.xls: Student · Semester I … · Subject name · Credits), atau Struktur Kurikulum (KELAS X/XI/XII).');
    if (type === 'KURIKULUM') { const k = parseKurikulum(wb, fileName); k.fileName = fileName || ''; return k; }
    const res = type === 'MASSAL' ? parseMassal(wb)
      : type === 'EKSPOR' ? parseEkspor(rowsOf(wb.Sheets[wb.SheetNames.find((n) => isEkspor(rowsOf(wb.Sheets[n])))]))
        : parseKalkulator(wb);
    res.fileName = fileName || '';
    res.summary = summarize(res);
    return res;
  }

  return { parse, detect, splitMapel, findSemesterBlocks, findIdentity, summarize, parseKurikulum, parseEkspor, rowsFromHtml, workbookFromHtml };
})();
// === PARSER END ===

// === DOCX PARSER START ===
/**
 * DocxParser — membaca nilai & identitas dari file DOCX transkrip.
 * Format dikenali: (1) Transkrip 3 Bahasa, (2) Academic Transcript GPA (termasuk tabel di
 * dalam text box), (3) Transkrip Nilai Rapor (Kelas X–XII, semester 1–6).
 * parseXml(documentXmlString, DOMParserClass) → {source, students:[...], grades:[...], warnings:[]}
 */
const DocxParser = (function () {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const ROMAN = { I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6 };
  const SEM_RE = /^SEMESTER\s+(VI|IV|V|I{1,3})\b/i;
  const YEAR_RE = /(\d{4})\s*\/\s*(\d{4})/;
  const MONTHS = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember',
    'january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  const norm = (s) => String(s == null ? '' : s).replace(/[\s ]+/g, ' ').trim();
  const isW = (n, name) => n && n.nodeType === 1 && n.namespaceURI === W && n.localName === name;
  const list = (nl) => { const a = []; for (let i = 0; i < nl.length; i++) a.push(nl.item ? nl.item(i) : nl[i]); return a; };
  const kids = (el) => list(el.childNodes).filter((n) => n.nodeType === 1);
  const num = (v) => { const s = norm(v).replace(',', '.'); return s !== '' && !isNaN(Number(s)) ? Number(s) : null; };

  function inFallback(el) {
    for (let p = el.parentNode; p; p = p.parentNode) if (p.localName === 'Fallback') return true;
    return false;
  }

  /** Teks paragraf tanpa isi text box/gambar yang tersisip di dalamnya. */
  function pText(p) {
    let s = '';
    (function walk(n) {
      kids(n).forEach((c) => {
        const ln = c.localName;
        if (ln === 'AlternateContent' || ln === 'drawing' || ln === 'pict' || ln === 'txbxContent' || ln === 'pPr' || ln === 'rPr') return;
        if (isW(c, 't')) s += c.textContent;
        else if (isW(c, 'tab')) s += '\t';
        else if (isW(c, 'br') || isW(c, 'cr')) s += '\n';
        else walk(c);
      });
    })(p);
    return s;
  }

  function cellInfo(tc) {
    const pr = kids(tc).find((c) => isW(c, 'tcPr'));
    const gs = pr ? kids(pr).find((c) => isW(c, 'gridSpan')) : null;
    const span = gs ? parseInt(gs.getAttributeNS(W, 'val') || gs.getAttribute('w:val'), 10) || 1 : 1;
    const text = kids(tc).filter((c) => isW(c, 'p')).map(pText).join('\n');
    return { span, text, ntext: norm(text) };
  }

  function tableRows(tbl) {
    return kids(tbl).filter((c) => isW(c, 'tr')).map((tr) => {
      let col = 0;
      return kids(tr).filter((c) => isW(c, 'tc')).map((tc) => { const ci = cellInfo(tc); ci.col = col; col += ci.span; return ci; });
    });
  }

  const cellAt = (row, col) => row.find((c) => c.col <= col && col < c.col + c.span) || null;
  const textAt = (row, col) => { const c = cellAt(row, col); return c ? c.ntext : ''; };

  function splitMapel(name) {
    const raw = norm(name);
    const parts = raw.split('/').map(norm).filter(Boolean);
    if (parts.length >= 3) return { raw, id: parts[0], en: parts.slice(1, -1).join('/'), ar: parts[parts.length - 1] };
    if (parts.length === 2) { const arabic = /[؀-ۿ]/.test(parts[1]); return { raw, id: parts[0], en: arabic ? '' : parts[1], ar: arabic ? parts[1] : '' }; }
    return { raw, id: raw, en: '', ar: '' };
  }

  function parseDateText(s) {
    s = norm(s);
    let m = /^(\d{1,2})\s+([A-Za-z]+)\.?\s+(\d{4})$/.exec(s);
    let d, mo, y;
    if (m) { d = +m[1]; mo = MONTHS.indexOf(m[2].toLowerCase()); y = +m[3]; }
    else if ((m = /^([A-Za-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})$/.exec(s))) { mo = MONTHS.indexOf(m[1].toLowerCase()); d = +m[2]; y = +m[3]; }
    if (!m || mo < 0) return '';
    mo = mo % 12;
    return y + '-' + String(mo + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  }

  /** Tabel nilai per semester (format GPA & 3 Bahasa). */
  function parseSemesterTable(rows) {
    let hdr = -1;
    const blocks = [];
    for (let r = 0; r < Math.min(rows.length, 6) && hdr < 0; r++) {
      rows[r].forEach((c) => {
        const m = SEM_RE.exec(c.ntext);
        if (m) { const y = YEAR_RE.exec(c.ntext); blocks.push({ sem: ROMAN[m[1].toUpperCase()], col: c.col, span: c.span, tahun: y ? y[1] + '/' + y[2] : '' }); }
      });
      if (blocks.length) hdr = r;
    }
    if (hdr < 0) return null;
    // tahun dari baris di atasnya (format GPA: "YEAR 2023/2024")
    let yearAbove = '';
    for (let r = 0; r < hdr; r++) rows[r].forEach((c) => { const y = YEAR_RE.exec(c.ntext); if (y && !yearAbove) yearAbove = y[1] + '/' + y[2]; });
    blocks.forEach((b) => { if (!b.tahun) b.tahun = yearAbove; });
    let subjCol = -1;
    for (let r = 0; r <= Math.min(hdr + 1, rows.length - 1) && subjCol < 0; r++) {
      const c = rows[r].find((x) => /^(MATA\s*PELAJARAN|SUBJECT)/i.test(x.ntext));
      if (c) subjCol = c.col;
    }
    if (subjCol < 0) subjCol = Math.max(0, blocks[0].col - 1);
    let sub = -1;
    for (let r = hdr; r < Math.min(rows.length, hdr + 3) && sub < 0; r++) if (rows[r].some((c) => /^(NILAI|REPORT)\b/i.test(c.ntext))) sub = r;
    if (sub < 0) return null;
    const english = rows[sub].some((c) => /^REPORT\b/i.test(c.ntext));
    blocks.forEach((b) => {
      rows[sub].forEach((c) => {
        if (c.col < b.col || c.col >= b.col + b.span) return;
        if (b.nilaiCol === undefined && /^(NILAI|REPORT)\b/i.test(c.ntext)) b.nilaiCol = c.col;
        if (b.sksCol === undefined && /^(SKS|CREDITS?)\b/i.test(c.ntext)) b.sksCol = c.col;
      });
    });
    const items = [];
    let urut = 0;
    for (let r = sub + 1; r < rows.length; r++) {
      const subj = textAt(rows[r], subjCol);
      if (!subj) continue;
      if (/^(total|jumlah|grade points?|indeks|ips|rata)/i.test(subj)) break;
      urut++;
      blocks.forEach((b) => {
        if (b.nilaiCol === undefined) return;
        const nilai = num(textAt(rows[r], b.nilaiCol));
        if (nilai === null) return;
        items.push({ sem: b.sem, tahun: b.tahun, name: subj, urutan: urut, nilai, sks: b.sksCol === undefined ? null : num(textAt(rows[r], b.sksCol)) });
      });
    }
    return { items, english };
  }

  /** Tabel rapor: Mata Pelajaran | Kelas X (1,2) | Kelas XI (3,4) | Kelas XII (5,6). */
  function parseRaporTable(rows) {
    const h = rows.findIndex((row) => row.some((c) => /^(kelas|grade)\s+x$/i.test(c.ntext)) && row.some((c) => /^(kelas|grade)\s+xi$/i.test(c.ntext)));
    if (h < 0) return null;
    const subjCell = rows[h].find((c) => /^(mata\s*pelajaran|subjects?)$/i.test(c.ntext)) || rows[h][0];
    const semRow = rows[h + 1] || [];
    const semCol = {};
    semRow.forEach((c) => { const n = parseInt(c.ntext, 10); if (n >= 1 && n <= 6 && String(n) === c.ntext) semCol[n] = c.col; });
    if (Object.keys(semCol).length < 2) return null;
    const items = [];
    let group = '', urut = 0;
    for (let r = h + 2; r < rows.length; r++) {
      const row = rows[r];
      const name = textAt(row, subjCell.col);
      if (!name) continue;
      if (/^(rata|average|jumlah)/i.test(name)) break;
      const vals = Object.keys(semCol).map((k) => textAt(row, semCol[k]));
      if (row.length === 1 || vals.every((v) => v === '')) { group = name; continue; }
      urut++;
      Object.keys(semCol).forEach((k) => {
        const nilai = num(textAt(row, semCol[k]));
        if (nilai === null || nilai <= 0) return; // 0 = tidak diambil
        items.push({ sem: Number(k), name, urutan: urut, nilai, group });
      });
    }
    return { items };
  }

  /** Rapor semester (Laporan Hasil Belajar): No | Mata Pelajaran | Nilai Akhir / Rerata Nilai | Capaian Kompetensi. */
  const GROUP_RE = /^(kelompok\b|muatan\s+lokal|mata\s+pelajaran\s+(umum|wajib|pilihan|muatan)|mapel\s+(umum|wajib|pilihan)|umum$|pilihan$|wajib$)/i;
  function parseLhbTable(rows) {
    const isHead = (row) => row.some((c) => /^no\.?$/i.test(c.ntext)) && row.some((c) => /^mata\s*pelajaran$/i.test(c.ntext)) &&
      row.some((c) => /nilai|rerata/i.test(c.ntext)) && row.some((c) => /capaian|deskripsi/i.test(c.ntext)) &&
      !row.some((c) => /^(kelas|grade)\s+x|sks|credits?|bobot|mutu|predikat/i.test(c.ntext));
    const h = rows.findIndex(isHead);
    if (h < 0) return null;
    const head = rows[h];
    const noCol = head.find((c) => /^no\.?$/i.test(c.ntext)).col;
    const mpCol = head.find((c) => /^mata\s*pelajaran$/i.test(c.ntext)).col;
    const nv = head.find((c) => /nilai|rerata/i.test(c.ntext) && !/capaian|deskripsi/i.test(c.ntext));
    const items = [];
    let group = '';
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r];
      if (isHead(row)) continue;
      const texts = [...new Set(row.map((c) => c.ntext).filter(Boolean))];
      const no = textAt(row, noCol), name = textAt(row, mpCol);
      const nilai = nv ? num(textAt(row, nv.col)) : null;
      if (nilai === null && texts.length === 1 && GROUP_RE.test(texts[0])) { group = texts[0]; continue; }
      if (!name || !/^\d{1,2}\.?$/.test(no) || GROUP_RE.test(name)) continue;
      items.push({ no: parseInt(no, 10), nama: name, nilai, kelompok: group });
    }
    return items.length ? { items } : null;
  }

  function lhbResult(parts, lines) {
    const items = [];
    const seen = {};
    parts.forEach((p) => p.items.forEach((it) => {
      const k = it.no + '|' + it.nama.toLowerCase();
      if (!seen[k]) { seen[k] = 1; items.push(it); }
    }));
    let kelas = '', fase = '', semTxt = '', tahun = '', nama = '', nis = '', nisn = '';
    lines.forEach((raw) => {
      const t = norm(String(raw).replace(/\t/g, '  '));
      let m;
      if (!kelas && (m = /\bKelas\s*:?\s*(XII|XI|X)(?![A-Za-z])/i.exec(t))) kelas = m[1].toUpperCase();
      if (!fase && (m = /\bFase\s*:?\s*([EF])\b/i.exec(t))) fase = m[1].toUpperCase();
      if (!semTxt && (m = /\bSemester\s*:?\s*(\d|ganjil|genap)/i.exec(t))) semTxt = /ganjil/i.test(t) && !/^\d/.test(m[1]) ? 'ganjil' : m[1].toLowerCase();
      if (!tahun && (m = /Tahun\s*Pelajaran\s*:?\s*(\d{4})\s*[-\/]\s*(\d{4})/i.exec(t))) tahun = m[1] + '/' + m[2];
      if (!nama && (m = /\bNama(?:\s+(?:Siswa|Peserta\s*didik))?\s*:\s*(.+?)(?:\s+Kelas\b|\s+NIS\b|$)/i.exec(t))) nama = norm(m[1]);
      if (!nis && (m = /\bNIS\s*\/\s*NISN\s*:?\s*([0-9.]+)\s*\/\s*([0-9]+)/i.exec(t))) { nis = m[1]; nisn = m[2]; }
      if (!nisn && (m = /\bNISN\s*:?\s*([0-9]{6,})/i.exec(t))) nisn = m[1];
      if (!nis && (m = /\bNIS\s*:?\s*([0-9.]{5,})(?!\s*\/)/i.exec(t))) nis = m[1];
    });
    if (!kelas && fase === 'E') kelas = 'X';
    if (!kelas) throw new Error('Rapor terbaca, tetapi kelasnya tidak ditemukan (baris "Kelas : X / XI / XII").');
    const base = { X: 0, XI: 2, XII: 4 }[kelas];
    const d = parseInt(semTxt, 10);
    const semester = d >= 3 ? d : base + (d === 2 || semTxt === 'genap' ? 2 : 1);
    const warnings = [];
    if (!semTxt) warnings.push('Semester tidak terbaca — dianggap semester ganjil.');
    const student = { NIS: nis, NISN: nisn, Nama_Siswa: nama.toUpperCase() };
    const grades = items.filter((it) => it.nilai !== null && it.nilai > 0).map((it) => ({
      NIS: nis, Kategori: 'RAPOR', Semester: semester, Tahun_Pelajaran: tahun, Urutan: it.no, Nilai: it.nilai, SKS: '',
      Mapel: { raw: it.nama, id: it.nama, en: '', ar: '', kelompok: it.kelompok }
    }));
    return { source: 'DOCX-SUSUNAN', susunan: { kelas, fase, semester, tahun, items }, students: nama ? [student] : [], grades: nama ? grades : [], warnings };
  }

  function identityFrom(lines) {
    const id = {};
    const set = (k, v) => { if (v && !id[k]) id[k] = v; };
    lines.forEach((line) => {
      const t = norm(line.replace(/\t/g, ' '));
      let m;
      if ((m = /^(?:nama siswa(?:\s*\/\s*student name)?|nama|name)\s*:\s*(.+)$/i.exec(t))) {
        const parts = m[1].split(' / ').map(norm);
        set('Nama_Siswa', parts.filter((x) => !/[؀-ۿ]/.test(x))[0] || parts[0]);
        const ar = parts.find((x) => /[؀-ۿ]/.test(x));
        if (ar) set('Nama_Arab', ar);
      } else if ((m = /^(?:tempat\s*\/\s*tanggal lahir|place\s*\/\s*date of birth)\s*:\s*(.+?),\s*(.+)$/i.exec(t))) {
        set('Tempat_Lahir', norm(m[1])); set('Tanggal_Lahir', parseDateText(m[2]) || norm(m[2]));
      } else if ((m = /^(?:jenis kelamin(?:\s*\/\s*sex)?|sex)\s*:\s*(.+)$/i.exec(t))) {
        set('Jenis_Kelamin', norm(m[1].split('/')[0]));
      } else if ((m = /^(?:nis\s*\/\s*nisn(?:\s*\/\s*registration number)?|registration number)\s*:\s*([0-9.]+)\s*\/\s*([0-9]+)/i.exec(t))) {
        set('NIS', m[1]); set('NISN', m[2]);
      } else if ((m = /^(?:nama orang tua(?:\s*\/\s*parent'?s? name)?|parent'?s? name)\s*:\s*(.+)$/i.exec(t))) {
        set('Nama_Ortu', norm(m[1]));
      } else if ((m = /^(?:masa studi(?:\s*\/\s*period of study)?|period of study)\s*:\s*(.+)$/i.exec(t))) {
        set('Masa_Studi', norm(m[1]));
      } else if ((m = /^(?:no\.|nomor(?:\s*\/\s*number)?|number)\s*:?\s*(\d[^\s·]*)/i.exec(t))) {
        set('Nomor', m[1]);
      } else if ((m = /^[A-Za-z][A-Za-z .'-]*,\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4}|[A-Za-z]+\s+\d{1,2},\s*\d{4})$/.exec(t))) {
        set('Tanggal', parseDateText(m[1]));
      }
    });
    return id;
  }

  function parseXml(xml, DOMParserImpl) {
    const DP = DOMParserImpl || DOMParser;
    const doc = new DP().parseFromString(xml, 'application/xml');
    const warnings = [];
    const paras = list(doc.getElementsByTagNameNS(W, 'p')).filter((p) => !inFallback(p));
    const tables = list(doc.getElementsByTagNameNS(W, 'tbl')).filter((t) => !inFallback(t)).map(tableRows);

    // baris identitas: paragraf + tabel label|nilai (format 3 Bahasa)
    const lines = paras.map(pText);
    tables.forEach((rows) => rows.forEach((row) => {
      if (row.length >= 2 && row.length % 2 === 0) for (let i = 0; i < row.length; i += 2) {
        const label = row[i].text.split('\n')[0];
        if (label && row[i + 1].ntext && !/^\d/.test(norm(label))) lines.push(norm(label) + ' : ' + row[i + 1].ntext);
      }
    }));
    // Rapor semester (e-Rapor) → susunan mapel
    const lhb = tables.map(parseLhbTable).filter(Boolean);
    if (lhb.length) {
      const all = paras.map(pText).concat(...tables.map((rows) => rows.map((row) => row.map((c) => c.ntext).join('  '))));
      return lhbResult(lhb, all);
    }
    const ident = identityFrom(lines);

    let kind = null;
    const grades = [];
    const seen = {};
    tables.forEach((rows) => {
      const rp = parseRaporTable(rows);
      if (rp) {
        kind = 'RAPOR';
        rp.items.forEach((it) => {
          const k = 'R|' + it.sem + '|' + it.name.toLowerCase();
          if (seen[k]) return; seen[k] = 1;
          grades.push({ Kategori: 'RAPOR', Semester: it.sem, Tahun_Pelajaran: '', Urutan: it.urutan, Nilai: it.nilai, SKS: '',
            Mapel: { raw: it.name, id: it.name, en: '', ar: '', kelompok: it.group } });
        });
        return;
      }
      const sp = parseSemesterTable(rows);
      if (!sp) return;
      if (kind !== 'RAPOR') kind = kind || 'SEM';
      sp.items.forEach((it) => {
        const k = 'G|' + it.sem + '|' + it.name.toLowerCase();
        if (seen[k]) return; seen[k] = 1;
        const mapel = it.name.indexOf('/') >= 0 ? splitMapel(it.name) : (sp.english ? { raw: it.name, id: '', en: it.name, ar: '' } : { raw: it.name, id: it.name, en: '', ar: '' });
        grades.push({ Kategori: 'GPA', Semester: it.sem, Tahun_Pelajaran: it.tahun, Urutan: it.urutan, Nilai: it.nilai, SKS: it.sks === null ? '' : it.sks, Mapel: mapel });
      });
    });

    if (!grades.length) throw new Error('Tidak ada tabel nilai yang dikenali di DOCX ini (format 3 Bahasa, GPA, atau Rapor).');
    if (!ident.NIS) throw new Error('NIS tidak ditemukan di DOCX (cari baris "NIS / NISN" atau "Registration Number").');

    const isRapor = grades.some((g) => g.Kategori === 'RAPOR');
    const tri = grades.some((g) => g.Mapel.ar);
    const student = { NIS: ident.NIS, NISN: ident.NISN || '' };
    ['Nama_Siswa', 'Nama_Arab', 'Jenis_Kelamin', 'Tempat_Lahir', 'Tanggal_Lahir', 'Nama_Ortu', 'Masa_Studi'].forEach((k) => { if (ident[k]) student[k] = ident[k]; });
    if (ident.Nomor) student[isRapor ? 'Nomor_Rapor' : 'Nomor_Transkrip'] = ident.Nomor;
    if (ident.Tanggal) student[isRapor ? 'Tanggal_Rapor' : 'Tanggal_Terbit'] = ident.Tanggal;
    grades.forEach((g) => { g.NIS = ident.NIS; });
    grades.filter((g) => g.Kategori === 'GPA').forEach((g) => { if (g.SKS === '') warnings.push(`Sem ${g.Semester} · ${g.Mapel.raw}: SKS kosong.`); });
    return { source: isRapor ? 'DOCX-RAPOR' : (tri ? 'DOCX-3BAHASA' : 'DOCX-GPA'), students: [student], grades, warnings };
  }

  return { parseXml, parseDateText, splitMapel, parseLhbTable };
})();
// === DOCX PARSER END ===

/* ---------------------------------------------------------------------------
 * 3. BOOTSTRAP
 * ------------------------------------------------------------------------- */
document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  bindLogin();
  renderStudentForm();
  // Dibuka dari menu Google Sheets "Buka website & langsung masuk": #kode=…
  const fromHash = /(?:^#|&)kode=([^&]+)/.exec(location.hash || '');
  if (fromHash) {
    API.setToken(decodeURIComponent(fromHash[1]), true);
    try { history.replaceState(null, '', location.pathname + location.search); } catch (e) { location.hash = ''; }
  }
  if (API.token) loadApp(); else showLogin();
});

function loadApp() {
  return withLoading('Memuat data dari Google Sheets…', () => call('apiBootstrap'))
    .then(applyBootstrap)
    .catch((e) => {
      if (e.code === 'AUTH') return;
      $('studentRows').innerHTML = `<tr><td colspan="6" class="text-center py-10 text-red-600">${esc(errMsg(e))}</td></tr>`;
      toast(errMsg(e), 'err', 9000);
    });
}

/* ---- Login dengan kode akses ------------------------------------------- */
function showLogin(msg) {
  $('loginErr').textContent = msg || '';
  $('loginErr').classList.toggle('hidden', !msg);
  $('loginServer').textContent = API.url ? API.url.replace(/^https:\/\//, '').replace(/\/macros\/s\/(.{6}).*(.{6})\/exec$/, '/macros/s/$1…$2/exec') : '(GAS_URL belum diisi di config.js)';
  $('loginPing').href = API.url ? API.url + '?action=ping' : '#';
  $('loginUrl').value = API.url;
  $('btnLogout').classList.add('hidden');
  openModal('loginModal');
  setTimeout(() => $('loginToken').focus(), 50);
}

function bindLogin() {
  $('loginForm').addEventListener('submit', async (ev) => {
    ev.preventDefault();
    const token = $('loginToken').value.trim();
    if (!token) { $('loginToken').focus(); return; }
    const btn = $('loginBtn');
    btn.disabled = true; btn.textContent = 'Memeriksa…';
    $('loginErr').classList.add('hidden');
    try {
      const d = await rawCall('apiBootstrap', [], token);
      API.setToken(token, $('loginRemember').checked);
      $('loginToken').value = '';
      closeModal('loginModal');
      $('btnLogout').classList.remove('hidden');
      try { applyBootstrap(d); toast('Berhasil masuk.', 'ok'); }
      catch (e) { toast(errMsg(e), 'err', 9000); }
    } catch (e) {
      $('loginErr').textContent = errMsg(e);
      $('loginErr').classList.remove('hidden');
    } finally {
      btn.disabled = false; btn.textContent = 'Masuk';
    }
  });
  $('loginUrlSave').addEventListener('click', () => {
    const u = $('loginUrl').value.trim();
    if (u && !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(u)) {
      $('loginErr').textContent = 'URL harus berbentuk https://script.google.com/macros/s/…/exec (bukan /dev).';
      $('loginErr').classList.remove('hidden'); return;
    }
    const def = String((window.APP_CONFIG && window.APP_CONFIG.GAS_URL) || '').trim();
    API.setUrl(u && u !== def ? u : '');
    $('loginErr').classList.add('hidden');
    showLogin();
    toast(u && u !== def ? 'URL server disimpan di browser ini.' : 'Memakai URL dari config.js.', 'ok');
  });
  $('loginShow').addEventListener('click', () => {
    const i = $('loginToken'); i.type = i.type === 'password' ? 'text' : 'password';
  });
  $('btnLogout').addEventListener('click', () => {
    API.setToken('');
    state.students = []; state.transcripts = []; state.selected.clear();
    $('studentRows').innerHTML = '';
    showLogin();
  });
}

function applyBootstrap(d) {
  if ($('btnLogout')) $('btnLogout').classList.remove('hidden');
  $('schoolName').textContent = d.school || '—';
  $('userEmail').textContent = d.user || 'Pengguna';
  $('appVersion').textContent = `${d.app.NAME} · v${d.app.VERSION}`;
  if (d.spreadsheetUrl) { $('linkSheet').href = d.spreadsheetUrl; $('linkSheet').classList.remove('hidden'); }
  if (d.status.outputFolder.ok) { $('linkFolder').href = d.status.outputFolder.url; $('linkFolder').classList.remove('hidden'); }
  state.modes = d.modes;
  state.subjects = d.subjects || [];
  state.transcripts = d.transcripts || [];
  state.jurusan = d.jurusan || [];
  state.kurikulum = d.kurikulum || [];
  renderJurusanSelect();
  renderStatus(d.status);
  setStudents(d.students || []);
  renderHistory();
}

function renderStatus(s) {
  const chip = (t) => {
    const cls = t.ok ? 'chip-ok' : (t.optional ? 'chip-warn' : 'chip-err');
    const mark = t.ok ? '✓' : (t.optional ? '•' : '✗');
    return `<span class="chip ${cls}" title="${esc(t.message || t.name || '')}">${mark} Template ${esc(t.label)}${t.ok ? '' : ': ' + esc(t.message || '')}</span>`;
  };
  const kur = state.kurikulum.length
    ? `<span class="chip chip-ok">✓ Kurikulum ${esc(state.kurikulum.map((k) => k.nama).join(', '))} · ${state.jurusan.length} jurusan</span>`
    : `<span class="chip chip-warn" title="Unggah file Struktur Kurikulum (.xls) di langkah 1">• Struktur kurikulum belum diunggah</span>`;
  $('statusBar').innerHTML = s.templates.map(chip).join('') +
    `<span class="chip ${s.outputFolder.ok ? 'chip-ok' : 'chip-err'}">${s.outputFolder.ok ? '✓' : '✗'} Folder output</span>` + kur;
  state.lastStatus = s;
}

function renderJurusanSelect() {
  const sel = $('jurusanSel');
  const cur = sel.value;
  sel.innerHTML = `<option value="">Sesuai jurusan masing-masing siswa</option>` +
    state.jurusan.map((j) => `<option value="${esc(j.Kode)}">${esc(j.Nama)}</option>`).join('') +
    `<option value="NONE">Tanpa struktur kurikulum (semua mapel yang ada nilainya)</option>`;
  if ([...sel.options].some((o) => o.value === cur)) sel.value = cur;
}

/* ---------------------------------------------------------------------------
 * 4. EVENT
 * ------------------------------------------------------------------------- */
function bindEvents() {
  const dz = $('dropzone');
  $('fileInput').addEventListener('change', (e) => { if (e.target.files[0]) handleFile(e.target.files[0]); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) handleFile(f); });
  $('btnSaveImport').addEventListener('click', saveImport);
  $('btnCancelImport').addEventListener('click', clearParsed);
  $('btnTemplateMassal').addEventListener('click', downloadMassalTemplate);

  document.querySelectorAll('input[name="mode"]').forEach((r) => r.addEventListener('change', syncFormatBox));
  syncFormatBox();

  $('tabStudents').addEventListener('click', () => switchTab('students'));
  $('tabHistory').addEventListener('click', () => switchTab('history'));
  $('search').addEventListener('input', () => { renderStudents(); renderHistory(); });
  $('chkAll').addEventListener('change', (e) => {
    visibleStudents().forEach((s) => (e.target.checked ? state.selected.add(s.NIS) : state.selected.delete(s.NIS)));
    renderStudents();
  });
  $('studentRows').addEventListener('click', onStudentAction);
  $('studentRows').addEventListener('change', (e) => {
    if (e.target.matches('input[data-nis]')) {
      e.target.checked ? state.selected.add(e.target.dataset.nis) : state.selected.delete(e.target.dataset.nis);
      renderStudents();
    }
  });
  $('historyRows').addEventListener('click', onHistoryAction);
  $('btnRefresh').addEventListener('click', refreshAll);
  $('btnAddStudent').addEventListener('click', () => openStudentForm());
  $('btnSubjects').addEventListener('click', () => { renderSubjects(); openModal('modalSubjects'); });
  $('btnTemplates').addEventListener('click', () => { openModal('modalTemplates'); loadTemplates(); });
  bindKurikulum();
  $('templateRows').addEventListener('click', onTemplateAction);
  $('templateRows').addEventListener('change', onTemplateFile);

  $('btnGenerate').addEventListener('click', () => generateMany([...state.selected], getMode(), getFormat(), getJurusan()));
  $('genResults').addEventListener('click', onDownloadClick);

  document.querySelectorAll('[data-close]').forEach((b) => b.addEventListener('click', () => closeModal(b.dataset.close)));
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') ['modalPreview', 'modalStudent', 'modalSubjects', 'modalEditor'].forEach((id) => { if (!$(id).classList.contains('hidden')) closeModal(id); });
  });
  $('formStudent').addEventListener('submit', saveStudentForm);
  $('btnPreviewExcel').addEventListener('click', () => {
    const p = state.preview;
    if (!p.data) return;
    const st = state.students.find((x) => x.NIS === p.nis) || { Nama_Siswa: p.nis };
    saveExcel(p.data, excelBaseName(p.nis, st.Nama_Siswa, p.mode, p.format));
  });
  $('btnPreviewGenerate').addEventListener('click', () => { const p = state.preview; closeModal('modalPreview'); generateMany([p.nis], p.mode, p.format, getJurusan()); });
  $('subjectRows').addEventListener('click', onSubjectAction);
  $('editorSemTabs').addEventListener('click', (e) => { const b = e.target.closest('button[data-sem]'); if (b) { readEditorGrid(); state.editor.sem = Number(b.dataset.sem); renderEditorGrades(); } });
  $('editorGrades').addEventListener('click', onEditorGridClick);
  $('btnEditorSave').addEventListener('click', () => saveEditor(false));
  $('btnEditorSaveGen').addEventListener('click', () => saveEditor(true));
}

function syncFormatBox() { $('formatBox').classList.toggle('hidden', getMode() === 'TRI'); }

function switchTab(tab) {
  state.tab = tab;
  $('tabStudents').classList.toggle('active', tab === 'students');
  $('tabHistory').classList.toggle('active', tab === 'history');
  $('panelStudents').classList.toggle('hidden', tab !== 'students');
  $('panelHistory').classList.toggle('hidden', tab !== 'history');
  renderStudents(); renderHistory();
}

async function refreshAll() {
  try {
    const [students, transcripts] = await withLoading('Memuat ulang…', () => Promise.all([call('apiListStudents'), call('apiListTranscripts')]));
    state.transcripts = transcripts; setStudents(students); renderHistory();
    toast('Data diperbarui dari Google Sheets.');
  } catch (e) { toast(errMsg(e), 'err'); }
}

/* ---------------------------------------------------------------------------
 * 5. UNGGAH → PARSE → SIMPAN
 * ------------------------------------------------------------------------- */
const SOURCE_LABEL = {
  KALKULATOR: 'Excel · Kalkulator Transkrip', MASSAL: 'Excel · Format Massal',
  KURIKULUM: 'Excel · Struktur Kurikulum', EKSPOR: 'Excel · Ekspor Transkrip', 'DOCX-GPA': 'DOCX · Academic Transcript (GPA)', 'DOCX-RAPOR': 'DOCX · Transkrip Nilai Rapor', 'DOCX-3BAHASA': 'DOCX · Transkrip 3 Bahasa',
  'DOCX-SUSUNAN': 'DOCX · Rapor (susunan mapel)', TEMPLATE: 'DOCX · Template dokumen'
};

/* ---- Template DOCX (berisi token <<…>> / [[…]]) ---------------------------- */
const TEMPLATE_SLOTS = [
  ['TEMPLATE_RAPOR_LTR', 'Rapor (Indonesia/English)'], ['TEMPLATE_RAPOR_RTL', 'Rapor (Arab, RTL)'],
  ['TEMPLATE_GPA_LTR', 'GPA (Indonesia/English)'], ['TEMPLATE_GPA_RTL', 'GPA (Arab, RTL)'],
  ['TEMPLATE_ID_LTR', '3 Bahasa'], ['TEMPLATE_ID_RTL', '3 Bahasa (Arab, RTL)']
];
const slotLabel = (k) => (TEMPLATE_SLOTS.find((x) => x[0] === k) || [k, k])[1];

/** Teks polos dari XML Word (token bisa terpecah ke beberapa run). */
function docxPlainText(xml) {
  return String(xml).replace(/<w:tab\/>/g, '\t').replace(/<\/w:p>/g, '\n').replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

function templateInfoFromDocx(fileName, xml) {
  const plain = docxPlainText(xml);
  const tokens = plain.match(/<<[A-Za-z0-9_]+>>|\[\[[A-Za-z0-9_]+(?:@[A-Za-z0-9_]+)?\.[A-Za-z0-9_]+\]\]/g) || [];
  if (tokens.length < 3) return null;
  const format = /\[\[RAPOR[.@]/.test(plain) ? 'RAPOR' : /\[\[S[1-6]\./.test(plain) ? 'GPA' : /\[\[P[1-3]\./.test(plain) ? 'TRI' : '';
  // RTL: tabel bidiVisual, atau hampir semua paragraf ber-bidi (paragraf Arab di template 3 Bahasa LTR tidak cukup)
  const nP = (xml.match(/<w:p[ >]/g) || []).length, nBidi = (xml.match(/<w:bidi\/>|<w:bidi w:val="(1|true|on)"\/>/g) || []).length;
  const rtl = /<w:bidiVisual\/>/.test(xml) || (nP > 0 && nBidi / nP > 0.6) || /_RTL\b|arab/i.test(fileName);
  let key = (/Template_(Transkrip|ID)_(LTR|RTL)/i.exec(fileName) && 'TEMPLATE_ID_' + /_(LTR|RTL)/i.exec(fileName)[1].toUpperCase()) ||
    (/Template_GPA_(LTR|RTL)/i.exec(fileName) && 'TEMPLATE_GPA_' + /_(LTR|RTL)/i.exec(fileName)[1].toUpperCase()) ||
    (/Template_Rapor_(LTR|RTL)/i.exec(fileName) && 'TEMPLATE_RAPOR_' + /_(LTR|RTL)/i.exec(fileName)[1].toUpperCase()) || '';
  if (!key && format) key = (format === 'TRI' ? 'TEMPLATE_ID_' : 'TEMPLATE_' + format + '_') + (rtl ? 'RTL' : 'LTR');
  return { format, rtl, key, tokens: [...new Set(tokens)] };
}

function bufferToBase64(buf) {
  const bytes = new Uint8Array(buf);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

function handleFile(file) {
  const isDocx = /\.docx$/i.test(file.name);
  if (!isDocx && !/\.(xlsx|xlsm|xls)$/i.test(file.name)) { toast('Pilih file Excel (.xlsx) atau Word (.docx).', 'err'); return; }
  if (isDocx && typeof JSZip === 'undefined') { toast('Library JSZip gagal dimuat. Periksa koneksi internet.', 'err'); return; }
  if (!isDocx && typeof XLSX === 'undefined') { toast('Library SheetJS gagal dimuat. Periksa koneksi internet.', 'err'); return; }
  showLoading(isDocx ? 'Membaca dokumen Word…' : 'Membaca file Excel…', file.name);
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      let res;
      if (isDocx) {
        const zip = await JSZip.loadAsync(e.target.result);
        const part = zip.file('word/document.xml');
        if (!part) throw new Error('File .docx tidak valid (word/document.xml tidak ada).');
        const xml = await part.async('string');
        const tpl = templateInfoFromDocx(file.name, xml);
        if (tpl) {
          res = { source: 'TEMPLATE', template: tpl, base64: bufferToBase64(e.target.result), students: [], grades: [], warnings: [], summary: [] };
          if (!tpl.format) res.warnings.push('Tabel nilai ([[RAPOR.…]], [[S1.…]] atau [[P1.…]]) tidak ditemukan — template akan ditolak saat dipasang.');
        } else {
          res = DocxParser.parseXml(xml);
          res.summary = ExcelParser.summarize(res);
        }
      } else {
        const bytes = new Uint8Array(e.target.result);
        const html = htmlTextOf(bytes);
        const wb = html ? ExcelParser.workbookFromHtml(html) : XLSX.read(bytes, { type: 'array', cellDates: true });
        res = ExcelParser.parse(wb, file.name);
      }
      matchMissingNis(res);
      res.fileName = file.name;
      state.parsed = res;
      renderParsed();
    } catch (err) { clearParsed(); toast(errMsg(err), 'err', 9000); }
    finally { hideLoading(); }
  };
  reader.onerror = () => { hideLoading(); toast('File tidak bisa dibaca.', 'err'); };
  reader.readAsArrayBuffer(file);
}

/** .xls hasil ekspor web sering berupa tabel HTML; kembalikan teksnya (atau '' bila bukan HTML). */
function htmlTextOf(bytes) {
  let i = 0;
  if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) i = 3;
  while (i < bytes.length && i < 64 && /\s/.test(String.fromCharCode(bytes[i]))) i++;
  if (bytes[i] !== 0x3C) return ''; // bukan diawali '<'
  const text = new TextDecoder('utf-8').decode(bytes);
  return /<table[\s>]/i.test(text) ? text : '';
}

const normPersonName = (s) => String(s || '').toUpperCase().replace(/[^A-Z\u00C0-\u024F ]+/g, ' ').replace(/\s+/g, ' ').trim();

/** File tanpa NIS (Ekspor Transkrip): cocokkan dengan siswa yang sudah ada berdasarkan nama. */
function matchMissingNis(res) {
  if (!res || !Array.isArray(res.students)) return;
  res.students.forEach((s) => {
    if (s.NIS) return;
    const nm = normPersonName(s.Nama_Siswa);
    let hits = s.NISN ? state.students.filter((x) => String(x.NISN || '').trim() === String(s.NISN).trim()) : [];
    if (hits.length !== 1) hits = state.students.filter((x) => normPersonName(x.Nama_Siswa) === nm);
    if (hits.length === 1) { assignNis(res, s, hits[0].NIS, hits[0].NISN); res.nisMatched = true; }
  });
}

function assignNis(res, student, nis, nisn) {
  student.NIS = String(nis).trim();
  if (nisn && !student.NISN) student.NISN = String(nisn).trim();
  res.grades.forEach((g) => { if (!g.NIS) g.NIS = student.NIS; });
  res.warnings = (res.warnings || []).filter((w) => !/^NIS tidak ada di file/.test(w));
  res.summary = ExcelParser.summarize(res);
}

function renderParsed() {
  const p = state.parsed;
  $('parseBadge').textContent = SOURCE_LABEL[p.source] || p.source;
  $('parseFile').textContent = p.fileName;
  $('btnSaveImport').textContent = p.source === 'TEMPLATE' ? 'Pasang template' : p.source === 'DOCX-SUSUNAN' ? 'Perbarui susunan mapel' : 'Simpan ke Google Sheets';
  if (p.source === 'KURIKULUM') return renderParsedKurikulum(p);
  if (p.source === 'TEMPLATE') return renderParsedTemplate(p);
  if (p.source === 'DOCX-SUSUNAN') return renderParsedSusunan(p);
  const st = p.students[0] || {};
  const extra = p.source.indexOf('DOCX') === 0
    ? `<div class="text-slate-500 mb-1">${['Nomor_Transkrip', 'Nomor_Rapor', 'Tempat_Lahir', 'Tanggal_Lahir', 'Nama_Ortu'].filter((k) => st[k]).map((k) => `${k.replace(/_/g, ' ')}: <b>${esc(st[k])}</b>`).join(' · ')}</div>` : '';
  const inf = p.info || {};
  const infoLine = p.source === 'EKSPOR'
    ? `<div class="text-slate-500 mb-1">${[inf.kelas && 'Kelas <b>' + esc(inf.kelas) + '</b>', inf.stream && esc(inf.stream), inf.ipk != null && 'IPK di file <b>' + esc(inf.ipk) + '</b>', inf.sks != null && inf.sks + ' SKS'].filter(Boolean).join(' · ')}
       <div class="text-[11px] mt-0.5">Urutan &amp; kelompok mapel mengikuti file. Kolom Grade/Grade point/Score dihitung ulang dari nilai dengan skala di sheet Skala.</div></div>` : '';
  const needNis = p.students.length === 1 && !p.students[0].NIS;
  const nisBox = needNis ? `
    <div class="my-2 p-2.5 rounded-lg bg-amber-50 border border-amber-200">
      <div class="font-semibold text-amber-900 mb-1">NIS tidak ada di file — isi sebelum menyimpan</div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="label" for="importNis">NIS *</label><input id="importNis" class="input input-sm font-mono" list="importNisList" placeholder="mis. 232410017" autocomplete="off"></div>
        <div><label class="label" for="importNisn">NISN</label><input id="importNisn" class="input input-sm font-mono" placeholder="opsional" autocomplete="off"></div>
      </div>
      <datalist id="importNisList">${state.students.map((x) => `<option value="${esc(x.NIS)}">${esc(x.Nama_Siswa)}</option>`).join('')}</datalist>
      <div class="text-[11px] text-amber-800 mt-1">Pilih dari daftar bila siswa sudah ada (data nilai semester yang sama akan diganti), atau ketik NIS baru.</div>
    </div>` : (p.nisMatched ? `<div class="text-[11px] text-emerald-700 mb-1">✓ NIS dicocokkan otomatis dengan siswa yang sudah ada (berdasarkan nama).</div>` : '');
  $('parseSummary').innerHTML = infoLine + nisBox +
    `<div class="text-slate-500 mb-1">${p.summary.length} siswa · ${p.grades.length} nilai terbaca${p.grades.some((g) => g.Kategori === 'RAPOR') ? ' · <b class="text-blue-700">disimpan sebagai nilai rapor</b>' : ''}</div>` + extra +
    p.summary.map((s) => `
      <div class="flex items-center gap-2 py-1 border-b border-slate-100 last:border-0">
        <div class="min-w-0 flex-1"><div class="font-semibold text-slate-800 truncate">${esc(s.Nama_Siswa)}</div>
        <div class="text-slate-500">${s.NIS ? 'NIS ' + esc(s.NIS) : '<span class="text-amber-700">NIS belum diisi</span>'} · ${s.count} nilai${s.rapor ? '' : ' · ' + s.sks + ' SKS'}</div></div>
        <div class="flex gap-0.5">${[1, 2, 3, 4, 5, 6].map((n) => `<span class="sem-dot ${s.semesters.includes(n) ? (s.rapor ? 'sem-rap' : 'sem-on') : 'sem-off'}">${n}</span>`).join('')}</div>
      </div>`).join('');
  const w = p.warnings || [];
  $('parseWarnings').classList.toggle('hidden', !w.length);
  $('parseWarnings').innerHTML = w.length ? `<b>${w.length} catatan:</b><ul class="list-disc pl-4 mt-1">${w.slice(0, 50).map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : '';
  $('parseResult').classList.remove('hidden');
}

function renderParsedKurikulum(p) {
  const e = p.kurikulum.entries;
  const pem = (k) => [...new Set(e.filter((x) => x.kelas === k && x.peminatan).map((x) => x.peminatan))];
  const umum = (k) => e.filter((x) => x.kelas === k && !x.peminatan);
  const isBk = (n) => /^(bk|bimbingan\s*(dan\s*)?konseling)$/i.test(n);
  const kelasHtml = (k) => {
    const peminatan = pem(k);
    return `<div class="py-1.5 border-b border-slate-100 last:border-0">
      <div class="font-bold text-slate-800">Kelas ${k} <span class="font-normal text-slate-500">· ${umum(k).length} mapel umum/kekhasan${peminatan.length ? ' · peminatan: ' + peminatan.map(esc).join(', ') : ''}</span></div>
      <div class="text-slate-600 leading-relaxed">${umum(k).map((x) => isBk(x.nama) ? `<s class="text-red-500" title="Tidak dicetak di transkrip">${esc(x.nama)}</s>` : esc(x.nama)).join(' · ')}</div>
      ${peminatan.map((pm) => `<div class="text-slate-600"><b>${esc(pm)}:</b> ${e.filter((x) => x.kelas === k && x.peminatan === pm).map((x) => esc(x.nama)).join(' · ')}</div>`).join('')}
    </div>`;
  };
  $('parseSummary').innerHTML = `
    <label class="label" for="kurNama">Nama kurikulum</label>
    <input id="kurNama" class="input input-sm mb-2" value="${esc(p.kurikulum.nama)}" placeholder="2026/2027">
    ${['X', 'XI', 'XII'].map(kelasHtml).join('')}
    <div class="text-slate-500 mt-1">Jurusan dibuat otomatis dari peminatan Kelas XI × paket Kelas XII. BK (dicoret) tidak dicetak di transkrip.</div>`;
  const w = p.warnings || [];
  $('parseWarnings').classList.toggle('hidden', !w.length);
  $('parseWarnings').innerHTML = w.map(esc).join('<br>');
  $('parseResult').classList.remove('hidden');
}

function renderParsedTemplate(p) {
  const t = p.template;
  const fmt = { RAPOR: 'Rapor', GPA: 'GPA', TRI: '3 Bahasa' }[t.format] || 'tidak dikenali';
  $('parseSummary').innerHTML = `
    <div class="text-slate-600 mb-2">Terdeteksi <b>template ${esc(fmt)}${t.rtl ? ' (Arab, RTL)' : ''}</b> · ${t.tokens.length} token.</div>
    <label class="label" for="tplKey">Ganti template</label>
    <select id="tplKey" class="input input-sm mb-2">${TEMPLATE_SLOTS.map(([k, l]) => `<option value="${k}" ${k === t.key ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>
    <div class="text-[11px] text-slate-500">Sebelum dipasang, template diuji: token tabel wajib harus ada dan diisi dengan data siswa sungguhan.
      Bila gagal, template lama tetap dipakai. Template lama diarsipkan dan bisa dikembalikan lewat tombol <b>Template</b>.</div>
    <details class="mt-1 text-[11px] text-slate-500"><summary class="cursor-pointer">Token di file ini</summary><div class="font-mono break-all mt-1">${t.tokens.map(esc).join(' ')}</div></details>`;
  const w = p.warnings || [];
  $('parseWarnings').classList.toggle('hidden', !w.length);
  $('parseWarnings').innerHTML = w.map(esc).join('<br>');
  $('parseResult').classList.remove('hidden');
}

async function saveTemplateUpload(p) {
  const key = p.forceKey || $('tplKey').value;
  try {
    const res = await withLoading('Memasang template…', () => call('apiUploadTemplate', { key, fileName: p.fileName, base64: p.base64 }),
      'Konversi ke Google Docs dan uji isi dengan data siswa');
    if (!p.forceKey) clearParsed();
    if (res.status) renderStatus(res.status);
    toast(`Template ${slotLabel(key)} diganti dan langsung dipakai.` + (res.test && !res.test.skipped ? `\nUji isi berhasil dengan data ${res.test.nama}.` : ''), 'ok', 9000);
    (res.warnings || []).forEach((w) => toast(w, 'warn', 12000));
    if (!$('modalTemplates').classList.contains('hidden')) loadTemplates();
  } catch (e) { toast(errMsg(e), 'err', 12000); }
}

function renderParsedSusunan(p) {
  const s = p.susunan;
  const st = p.students[0];
  let last = null;
  const list = s.items.map((it) => {
    const head = it.kelompok && it.kelompok !== last ? `<div class="mt-1 font-bold text-slate-700">${esc(it.kelompok)}</div>` : '';
    last = it.kelompok || last;
    return head + `<div class="flex gap-2"><span class="w-6 text-right text-slate-400">${it.no}</span><span class="flex-1">${esc(it.nama)}</span><span class="text-slate-500">${it.nilai === null ? '' : esc(it.nilai)}</span></div>`;
  }).join('');
  $('parseSummary').innerHTML = `
    <div class="text-slate-600 mb-1">Rapor <b>Kelas ${esc(s.kelas)}</b> · semester ${s.semester}${s.tahun ? ' · ' + esc(s.tahun) : ''}${st ? ' · ' + esc(st.Nama_Siswa) : ''} · <b>${s.items.length} mapel</b></div>
    <div class="text-[11px] text-slate-500 mb-1">Urutan, kelompok, dan nama mapel di rapor ini menjadi susunan transkrip untuk Kelas ${esc(s.kelas)} (sheet Susunan_Mapel).</div>
    ${s.kelas === 'XI' ? '<label class="flex items-center gap-2 my-1"><input id="susXII" type="checkbox" checked> Terapkan juga ke Kelas XII</label>' : ''}
    ${st && p.grades.length ? `<label class="flex items-center gap-2 my-1"><input id="susNilai" type="checkbox"> Simpan juga ${p.grades.length} nilai rapor siswa ini (semester ${s.semester})</label>
      ${!st.NIS ? `<div id="susNisBox" class="hidden"><label class="label" for="importNis">NIS siswa *</label><input id="importNis" class="input input-sm font-mono" list="importNisList" placeholder="NIS" autocomplete="off">
        <datalist id="importNisList">${state.students.map((x) => `<option value="${esc(x.NIS)}">${esc(x.Nama_Siswa)}</option>`).join('')}</datalist></div>` : ''}` : ''}
    <div id="susCheck" class="my-1 text-[11px] text-slate-500">Memeriksa kecocokan mapel…</div>
    <div class="border-t border-slate-100 pt-1">${list}</div>`;
  if ($('susNilai') && $('susNisBox')) $('susNilai').addEventListener('change', (e) => $('susNisBox').classList.toggle('hidden', !e.target.checked));
  const w = p.warnings || [];
  $('parseWarnings').classList.toggle('hidden', !w.length);
  $('parseWarnings').innerHTML = w.map(esc).join('<br>');
  $('parseResult').classList.remove('hidden');
  call('apiSaveSusunan', susunanPayload(p, true)).then((r) => {
    if (state.parsed !== p) return;
    const rep = r.report[s.kelas] || {};
    $('susCheck').innerHTML = `✓ ${rep.umum} mapel umum/muatan lokal${rep.pilihan ? ` · ${rep.pilihan} mapel pilihan → pola <b>${esc(rep.pola)}</b>${rep.polaBaru ? ' (baru)' : ' (diperbarui)'}` : ''}` +
      (rep.extras && rep.extras.length ? ` · tetap disimpan (tidak ada di rapor ini): ${rep.extras.map(esc).join(', ')}` : '') +
      (r.newSubjects.length ? `<div class="text-amber-700">Mapel baru akan dibuat: ${r.newSubjects.map(esc).join(', ')}</div>` : '');
  }).catch((e) => { if ($('susCheck')) $('susCheck').innerHTML = `<span class="text-red-600">${esc(errMsg(e))}</span>`; });
}

function susunanPayload(p, dryRun) {
  const s = p.susunan;
  return { kelas: s.kelas, semester: s.semester, fileName: p.fileName, alsoXII: !!($('susXII') && $('susXII').checked),
    items: s.items.map((it) => ({ no: it.no, nama: it.nama, kelompok: it.kelompok })), dryRun };
}

async function saveSusunan(p) {
  const withNilai = !!($('susNilai') && $('susNilai').checked);
  const st = p.students[0];
  if (withNilai && st && !st.NIS) {
    const nis = $('importNis') ? $('importNis').value.trim() : '';
    if (!/^[0-9A-Za-z.\-\/]{3,30}$/.test(nis)) { toast('Isi NIS siswa untuk menyimpan nilai rapor.', 'err'); return; }
    assignNis(p, st, nis, st.NISN);
  }
  try {
    const res = await withLoading('Memperbarui susunan mapel…', () => call('apiSaveSusunan', susunanPayload(p, false)));
    let msg = `Susunan mapel Kelas ${res.targets.join(' & ')} diperbarui dari rapor.`;
    const rep = res.report[res.kelas] || {};
    if (rep.pilihan) msg += `\nPola pilihan: ${rep.pola}${rep.polaBaru ? ' (baru)' : ''}.`;
    if (withNilai) {
      const imp = await withLoading('Menyimpan nilai rapor…', () => call('apiImport', { source: p.source, fileName: p.fileName, students: p.students, grades: p.grades }));
      setStudents(imp.students);
      msg += `\nNilai rapor tersimpan: ${imp.gradesSaved} nilai.`;
    }
    clearParsed();
    toast(msg, 'ok', 9000);
    if (res.newSubjects.length) toast(`Mapel baru: ${res.newSubjects.join(', ')} — lengkapi nama EN/AR di menu Mata Pelajaran.`, 'warn', 10000);
  } catch (e) { toast(errMsg(e), 'err', 10000); }
}

/* ---- Database Kurikulum (editor) ----------------------------------------- */
const KUR_GROUPS = [['Kelompok Umum', 'U'], ['Kelompok Peminatan', 'P'], ['Kelompok Kekhasan SMAIT', 'K']];
let kurSeq = 0;
const kurGroupOf = (r) => (/peminatan|pilihan/i.test(r.Kelompok) ? 'P' : /kekhasan|muatan|lokal/i.test(r.Kelompok) ? 'K' : 'U');

async function openKurikulum(nama) {
  openModal('modalKurikulum');
  $('kurBody').innerHTML = '<div class="text-center py-10 text-slate-500">Memuat database kurikulum…</div>';
  try { setKur(await call('apiGetKurikulum', nama || '')); }
  catch (e) { $('kurBody').innerHTML = `<div class="text-center py-10 text-red-600">${esc(errMsg(e))}</div>`; }
}

function setKur(d) {
  state.kur = { nama: d.nama, aktif: d.aktif, daftar: d.daftar, bawaan: d.bawaan, subjects: d.subjects || [],
    rows: d.rows.map((r) => Object.assign({ _id: ++kurSeq }, r)), tab: (state.kur && state.kur.tab) || 'X', dirty: !!d.bawaan };
  const names = [...new Set(d.daftar.concat([d.nama]))];
  $('kurPilih').innerHTML = names.map((n) => `<option ${n === d.nama ? 'selected' : ''}>${esc(n)}</option>`).join('');
  $('kurAktif').textContent = d.nama === d.aktif ? '✓ aktif di transkrip' : 'tidak aktif — disimpan = dijadikan aktif';
  $('kurAktif').className = 'chip ' + (d.nama === d.aktif ? 'chip-ok' : 'chip-warn');
  renderKur();
}

function kurSection(kelas, g, pem) {
  return state.kur.rows.filter((r) => r.Kelas === kelas && kurGroupOf(r) === g && (g !== 'P' || r.Peminatan === pem))
    .sort((a, b) => (Number(a.Urutan) || 0) - (Number(b.Urutan) || 0));
}

function renderKur() {
  const k = state.kur, kelas = k.tab;
  document.querySelectorAll('#kurTabs .tab').forEach((b) => b.classList.toggle('active', b.dataset.kelas === kelas));
  const subjOpts = (sel) => `<option value="">(dari nama)</option>` + k.subjects.map((s) => `<option value="${esc(s.kode)}" ${s.kode === sel ? 'selected' : ''}>${esc(s.kode)} — ${esc(s.nama)}</option>`).join('');
  const jpSum = (rows) => rows.filter((r) => r.Tampil !== 'TIDAK').reduce((a, r) => a + (Number(String(r.JP).replace(',', '.')) || 0), 0);
  const table = (rows, g, pem) => `
    <table class="data mb-1">
      <thead><tr><th class="w-10">No</th><th>Mata Pelajaran (struktur)</th><th>Nama di transkrip</th><th class="w-56">Kode</th><th class="w-16">JP</th><th class="w-14">Cetak</th><th class="w-24"></th></tr></thead>
      <tbody>${rows.map((r, i) => `
        <tr data-id="${r._id}" class="${r.Tampil === 'TIDAK' ? 'opacity-50' : ''}">
          <td class="text-slate-400">${i + 1}</td>
          <td><input class="input input-sm" data-f="Nama_Mapel" value="${esc(r.Nama_Mapel)}" placeholder="mis. Sejarah"></td>
          <td><input class="input input-sm" data-f="Nama_Transkrip" value="${esc(r.Nama_Transkrip)}" placeholder="otomatis"></td>
          <td><select class="input input-sm" data-f="Kode_Mapel">${subjOpts(r.Kode_Mapel)}</select></td>
          <td><input class="input input-sm text-right" data-f="JP" value="${esc(r.JP)}" inputmode="decimal"></td>
          <td class="text-center"><input type="checkbox" data-f="Tampil" ${r.Tampil !== 'TIDAK' ? 'checked' : ''}></td>
          <td class="whitespace-nowrap"><button class="btn btn-ghost btn-sm" data-act="up" type="button" title="Naik">↑</button><button class="btn btn-ghost btn-sm" data-act="down" type="button" title="Turun">↓</button><button class="btn btn-ghost btn-sm text-red-600" data-act="del" type="button" title="Hapus">✕</button></td>
        </tr>`).join('')}
      </tbody>
      <tfoot><tr><td colspan="4"><button class="btn btn-ghost btn-sm" data-add="${g}" data-pem="${esc(pem || '')}" type="button">+ Tambah mapel</button></td>
        <td class="text-right font-bold">${jpSum(rows)}</td><td colspan="2" class="text-xs text-slate-500">Jumlah JP</td></tr></tfoot>
    </table>`;
  let html = '';
  const umum = kurSection(kelas, 'U'), khas = kurSection(kelas, 'K');
  html += `<h4 class="font-bold text-brand-800 mt-1 mb-1">Kelompok Umum</h4>` + table(umum, 'U');
  const pems = [...new Set(k.rows.filter((r) => r.Kelas === kelas && kurGroupOf(r) === 'P').map((r) => r.Peminatan))];
  html += `<div class="flex items-center gap-2 mt-4 mb-1"><h4 class="font-bold text-brand-800">Kelompok Peminatan</h4>
    <button class="btn btn-ghost btn-sm" data-addpem="1" type="button">+ Peminatan / Paket</button></div>`;
  if (!pems.length) html += `<div class="text-xs text-slate-500 mb-2">Tidak ada peminatan di kelas ini.</div>`;
  pems.forEach((p) => {
    html += `<div class="mt-2 mb-1 flex items-center gap-2"><span class="text-xs text-slate-500">Peminatan / paket:</span>
      <input class="input input-sm w-64 font-semibold" data-pemname="${esc(p)}" value="${esc(p)}"></div>` + table(kurSection(kelas, 'P', p), 'P', p);
  });
  html += `<h4 class="font-bold text-brand-800 mt-4 mb-1">Kelompok Kekhasan SMAIT</h4>` + table(khas, 'K');
  $('kurBody').innerHTML = html;
  const base = jpSum(umum) + jpSum(khas);
  $('kurTotal').innerHTML = `Jumlah JP Kelas ${kelas}: ` + (pems.length
    ? pems.map((p) => `${esc(p)} <b>${base + jpSum(kurSection(kelas, 'P', p))}</b>`).join(' · ') : `<b>${base}</b>`);
  $('kurDirty').textContent = k.dirty ? (k.bawaan ? 'Database bawaan (struktur kurikulum fix) — klik Simpan untuk memasangnya.' : 'Ada perubahan yang belum disimpan.') : '';
}

function kurRow(el) { const tr = el.closest('tr[data-id]'); return tr ? state.kur.rows.find((r) => r._id === Number(tr.dataset.id)) : null; }
function kurMark() { state.kur.dirty = true; $('kurDirty').textContent = 'Ada perubahan yang belum disimpan.'; }

function bindKurikulum() {
  $('btnKurikulum').addEventListener('click', () => openKurikulum());
  $('kurTabs').addEventListener('click', (e) => { const b = e.target.closest('[data-kelas]'); if (b) { state.kur.tab = b.dataset.kelas; renderKur(); } });
  $('kurPilih').addEventListener('change', (e) => openKurikulum(e.target.value));
  $('kurReload').addEventListener('click', () => openKurikulum(state.kur && state.kur.nama));
  $('kurBody').addEventListener('input', (e) => {
    const r = kurRow(e.target);
    if (r && e.target.dataset.f && e.target.type !== 'checkbox') { r[e.target.dataset.f] = e.target.value; kurMark(); }
  });
  $('kurBody').addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.pemname !== undefined) {
      const old = t.dataset.pemname, nw = t.value.trim();
      if (!nw) { t.value = old; return; }
      state.kur.rows.forEach((r) => { if (r.Kelas === state.kur.tab && kurGroupOf(r) === 'P' && r.Peminatan === old) r.Peminatan = nw; });
      kurMark(); renderKur(); return;
    }
    const r = kurRow(t);
    if (!r) return;
    if (t.dataset.f === 'Tampil') { r.Tampil = t.checked ? 'YA' : 'TIDAK'; kurMark(); renderKur(); }
    else if (t.dataset.f === 'Kode_Mapel') {
      r.Kode_Mapel = t.value;
      const s = state.kur.subjects.find((x) => x.kode === t.value);
      if (s && !r.Nama_Mapel) r.Nama_Mapel = s.nama;
      kurMark(); renderKur();
    } else if (t.dataset.f === 'JP') renderKur();
  });
  $('kurBody').addEventListener('click', (e) => {
    const k = state.kur, kelas = k.tab;
    const add = e.target.closest('[data-add]');
    if (add) {
      const g = add.dataset.add, pem = add.dataset.pem;
      const sec = kurSection(kelas, g, pem);
      k.rows.push({ _id: ++kurSeq, Kelas: kelas, Peminatan: g === 'P' ? pem : '', Kelompok: KUR_GROUPS.find((x) => x[1] === g)[0],
        Urutan: (sec.length ? Number(sec[sec.length - 1].Urutan) || 0 : 0) + 1, Nama_Mapel: '', Nama_Transkrip: '', Kode_Mapel: '', JP: '', Tampil: 'YA' });
      kurMark(); renderKur(); return;
    }
    if (e.target.closest('[data-addpem]')) {
      let n = 1, name = 'Peminatan Baru';
      while (k.rows.some((r) => r.Kelas === kelas && r.Peminatan === name)) name = 'Peminatan Baru ' + (++n);
      k.rows.push({ _id: ++kurSeq, Kelas: kelas, Peminatan: name, Kelompok: 'Kelompok Peminatan', Urutan: 1, Nama_Mapel: '', Nama_Transkrip: '', Kode_Mapel: '', JP: '', Tampil: 'YA' });
      kurMark(); renderKur(); return;
    }
    const btn = e.target.closest('[data-act]');
    const r = btn && kurRow(btn);
    if (!r) return;
    const sec = kurSection(r.Kelas, kurGroupOf(r), r.Peminatan);
    const i = sec.indexOf(r);
    if (btn.dataset.act === 'del') k.rows = k.rows.filter((x) => x !== r);
    else {
      const j = btn.dataset.act === 'up' ? i - 1 : i + 1;
      if (j < 0 || j >= sec.length) return;
      sec.forEach((x, n) => { x.Urutan = n + 1; });
      [sec[i].Urutan, sec[j].Urutan] = [sec[j].Urutan, sec[i].Urutan];
    }
    kurMark(); renderKur();
  });
  $('kurSave').addEventListener('click', saveKurikulumDb);
}

async function saveKurikulumDb() {
  const k = state.kur;
  // urutan bersambung per kelompok agar nomor rapi
  ['X', 'XI', 'XII'].forEach((kelas) => {
    let base = 0;
    kurSection(kelas, 'U').forEach((r) => { r.Urutan = ++base; });
    let maxP = base;
    [...new Set(k.rows.filter((r) => r.Kelas === kelas && kurGroupOf(r) === 'P').map((r) => r.Peminatan))].forEach((p) => {
      let n = base;
      kurSection(kelas, 'P', p).forEach((r) => { r.Urutan = ++n; });
      maxP = Math.max(maxP, n);
    });
    kurSection(kelas, 'K').forEach((r) => { r.Urutan = ++maxP; });
  });
  const rows = k.rows.filter((r) => String(r.Nama_Mapel).trim() || r.Kode_Mapel)
    .map(({ _id, ...r }) => r);
  try {
    const res = await withLoading('Menyimpan database kurikulum…', () => call('apiSaveKurikulum', { nama: k.nama, rows, aktif: true }));
    setKur(res);
    state.kur.dirty = false; renderKur();
    state.jurusan = res.jurusan; state.kurikulum = res.kurikulum;
    renderJurusanSelect(); if (state.lastStatus) renderStatus(state.lastStatus);
    toast(`Kurikulum ${res.nama} disimpan dan aktif. Transkrip berikutnya memakai urutan, nama, dan JP ini.` +
      (res.newSubjects && res.newSubjects.length ? `\nMapel baru: ${res.newSubjects.join(', ')} — lengkapi nama EN/AR di menu Mata Pelajaran.` : ''), 'ok', 9000);
  } catch (e) { toast(errMsg(e), 'err', 15000); }
}

/* ---- Modal Template ------------------------------------------------------ */
async function loadTemplates() {
  $('templateRows').innerHTML = `<tr><td colspan="4" class="text-center py-6 text-slate-500">Memuat…</td></tr>`;
  try {
    const list = await call('apiListTemplates');
    $('templateRows').innerHTML = list.map((t) => `
      <tr>
        <td><div class="font-semibold text-slate-800">${esc(t.label)}</div><div class="text-[11px] text-slate-400 font-mono">${esc(t.fileName)}</div></td>
        <td>${t.ok ? `<span class="chip chip-ok">✓ terpasang</span><div class="text-[11px] text-slate-500 mt-0.5">diubah ${esc(t.updated)}</div>`
          : `<span class="chip ${t.rtl ? 'chip-warn' : 'chip-err'}">${t.rtl ? '•' : '✗'} ${esc(t.message || 'belum')}</span>`}</td>
        <td class="whitespace-nowrap">
          ${t.ok ? `<a class="btn btn-ghost btn-sm" href="${esc(t.url)}" target="_blank" rel="noopener">Buka ↗</a>
          <button class="btn btn-ghost btn-sm" data-tpl-dl="${esc(t.key)}" type="button">⬇ DOCX</button>` : ''}
          <label class="btn btn-primary btn-sm cursor-pointer">⬆ Ganti DOCX<input type="file" accept=".docx" class="hidden" data-tpl-up="${esc(t.key)}"></label>
          ${t.hasPrev ? `<button class="btn btn-ghost btn-sm" data-tpl-restore="${esc(t.key)}" type="button" title="Kembalikan template sebelumnya">↶ Kembalikan</button>` : ''}
        </td>
      </tr>`).join('');
  } catch (e) { $('templateRows').innerHTML = `<tr><td colspan="4" class="text-center py-6 text-red-600">${esc(errMsg(e))}</td></tr>`; }
}

async function onTemplateAction(ev) {
  const dl = ev.target.closest('[data-tpl-dl]'), rs = ev.target.closest('[data-tpl-restore]');
  if (dl) {
    try { const f = await withLoading('Menyiapkan DOCX…', () => call('apiDownloadTemplate', dl.dataset.tplDl)); saveBase64(f.base64, f.mimeType, f.name); }
    catch (e) { toast(errMsg(e), 'err'); }
  } else if (rs) {
    const key = rs.dataset.tplRestore;
    if (!(await confirmBox('Kembalikan template', `Kembalikan template ${slotLabel(key)} ke versi sebelumnya?`))) return;
    try {
      const r = await withLoading('Mengembalikan template…', () => call('apiRestoreTemplate', key));
      if (r.status) renderStatus(r.status);
      toast(`Template ${slotLabel(key)} dikembalikan ke versi sebelumnya.`, 'ok');
      loadTemplates();
    } catch (e) { toast(errMsg(e), 'err'); }
  }
}

function onTemplateFile(ev) {
  const input = ev.target.closest('[data-tpl-up]');
  if (!input || !input.files[0]) return;
  const file = input.files[0], key = input.dataset.tplUp;
  input.value = '';
  if (!/\.docx$/i.test(file.name)) { toast('Pilih file Word (.docx).', 'err'); return; }
  const reader = new FileReader();
  reader.onload = () => saveTemplateUpload({ fileName: file.name, base64: bufferToBase64(reader.result), forceKey: key });
  reader.readAsArrayBuffer(file);
}

function clearParsed() { state.parsed = null; $('parseResult').classList.add('hidden'); $('btnSaveImport').textContent = 'Simpan ke Google Sheets'; }

async function saveImport() {
  const p = state.parsed;
  if (!p) return;
  if (p.source === 'KURIKULUM') return saveKurikulum(p);
  if (p.source === 'TEMPLATE') return saveTemplateUpload(p);
  if (p.source === 'DOCX-SUSUNAN') return saveSusunan(p);
  const missing = p.students.filter((s) => !s.NIS);
  if (missing.length) {
    const nis = $('importNis') ? $('importNis').value.trim() : '';
    if (!/^[0-9A-Za-z.\-\/]{3,30}$/.test(nis)) { toast('Isi NIS siswa terlebih dahulu (3–30 karakter, tanpa spasi).', 'err'); if ($('importNis')) $('importNis').focus(); return; }
    const exist = state.students.find((x) => String(x.NIS) === nis);
    if (exist && normPersonName(exist.Nama_Siswa) !== normPersonName(missing[0].Nama_Siswa) &&
      !(await confirmBox('NIS sudah dipakai', `NIS ${nis} terdaftar atas nama ${exist.Nama_Siswa}.\nData siswa itu akan diperbarui dengan nama ${missing[0].Nama_Siswa} dan nilainya. Lanjutkan?`))) return;
    assignNis(p, missing[0], nis, $('importNisn') ? $('importNisn').value.trim() : '');
  }
  const payload = { source: p.source, fileName: p.fileName, students: p.students, grades: p.grades };
  try {
    const res = await withLoading('Menyimpan ke Google Sheets…', () => call('apiImport', payload), `${p.students.length} siswa · ${p.grades.length} nilai`);
    state.subjects = res.subjects || state.subjects;
    setStudents(res.students);
    p.students.forEach((x) => state.selected.add(String(x.NIS)));
    renderStudents();
    clearParsed();
    let msg = `Tersimpan: ${res.studentsInserted} siswa baru, ${res.studentsUpdated} diperbarui, ${res.gradesSaved} nilai.`;
    if (res.newSubjects.length) msg += `\nMapel baru: ${res.newSubjects.join(', ')} — lengkapi nama EN/AR di menu Mata Pelajaran.`;
    toast(msg, 'ok', 8000);
    if (res.warnings.length) toast(`${res.warnings.length} catatan, contoh: ${res.warnings[0]}`, 'warn', 9000);
  } catch (e) { toast(errMsg(e), 'err', 9000); }
}

async function saveKurikulum(p) {
  const nama = ($('kurNama').value || '').trim();
  if (!nama) { toast('Isi nama kurikulum (mis. 2026/2027).', 'err'); return; }
  try {
    const res = await withLoading('Menyimpan struktur kurikulum…', () => call('apiImportKurikulum', { nama, fileName: p.fileName, entries: p.kurikulum.entries }));
    state.jurusan = res.jurusan; state.kurikulum = res.kurikulum; state.subjects = res.subjects;
    renderJurusanSelect(); if (state.lastStatus) renderStatus(state.lastStatus); setStudents(res.students);
    clearParsed();
    toast(`Kurikulum ${res.nama} tersimpan ke database dan dijadikan aktif (${res.mapel} baris, ${res.jurusan.length} jurusan). Urutan, nama, dan JP transkrip kini mengikuti struktur ini — cek/ubah di menu Kurikulum.` +
      (res.tersembunyi.length ? `\nTidak dicetak di transkrip: ${res.tersembunyi.join(', ')}.` : '') +
      (res.newSubjects.length ? `\nMapel baru: ${res.newSubjects.join(', ')} — lengkapi nama EN/AR di menu Mata Pelajaran.` : ''), 'ok', 9000);
  } catch (e) { toast(errMsg(e), 'err', 9000); }
}

function downloadMassalTemplate() {
  const wb = XLSX.utils.book_new();
  const siswa = [
    ['NIS', 'NISN', 'Nama_Siswa', 'Nama_Arab', 'Jenis_Kelamin', 'Jurusan', 'Tempat_Lahir', 'Tanggal_Lahir', 'Nama_Ortu', 'Masa_Studi', 'Nomor_Transkrip', 'Tanggal_Terbit', 'Nomor_Rapor', 'Tanggal_Rapor'],
    ['232410017', '0085130125', 'ASYFA NAILATUL YUSRA', 'أسيفا نائلة اليسرى', 'P', (state.jurusan[0] || {}).Kode || '', 'Padang', '2008-01-20', 'NUR WISNU SIDHARTA', '2023/2024 – 2025/2026', '09.471/SMAIT/YKPI/V/2026', '2026-05-13', '09.498/SMAIT/YKPI/VI/2026', '2026-05-05']
  ];
  const nilai = [
    ['NIS', 'Kategori', 'Semester', 'Tahun_Pelajaran', 'Urutan', 'Mata_Pelajaran', 'Kelompok', 'Nilai', 'SKS'],
    ['232410017', 'GPA', 1, '2023/2024', 1, 'Pendidikan Agama Islam/Islamic Religious Education/التربية الإسلامية', '', 88, 2],
    ['232410017', 'GPA', 1, '2023/2024', 2, 'MTK', '', 80, 4],
    ['232410017', 'RAPOR', 1, '2023/2024', 1, 'Pendidikan Agama dan Budi Pekerti', 'Wajib', 88, ''],
    ['232410017', 'RAPOR', 1, '2023/2024', 2, 'Ilmu Pengetahuan Alam', 'Wajib', 80, '']
  ];
  const petunjuk = [
    ['PETUNJUK FORMAT MASSAL'],
    ['Sheet SISWA: satu baris per siswa. NIS wajib. Jenis_Kelamin: L / P. Tanggal: YYYY-MM-DD.'],
    ['Jurusan: kode dari sheet Jurusan (mis. ' + (state.jurusan.map((j) => j.Kode).join(', ') || 'SAINS-A-2026') + '). Boleh kosong.'],
    ['Sheet NILAI: satu baris per mapel per semester. Semester 1–6 (atau I–VI).'],
    ['Kategori: GPA (pakai SKS; untuk transkrip 3 Bahasa & GPA) atau RAPOR (tanpa SKS; untuk transkrip Nilai Rapor).'],
    ['Kelompok (khusus RAPOR): Wajib / Pilihan / Muatan Lokal.'],
    ['Mata_Pelajaran boleh: kode (MTK), nama Indonesia, atau "Indonesia/English/Arab".'],
    ['Nilai untuk (NIS, Kategori, Semester) yang ada di file akan MENGGANTI data lama; semester lain tidak berubah.']
  ];
  const wsS = XLSX.utils.aoa_to_sheet(siswa), wsN = XLSX.utils.aoa_to_sheet(nilai), wsP = XLSX.utils.aoa_to_sheet(petunjuk);
  ['A2', 'B2'].forEach((c) => { wsS[c].t = 's'; wsS[c].z = '@'; });
  wsS['!cols'] = [10, 12, 26, 20, 8, 16, 12, 12, 22, 20, 26, 12, 26, 12].map((w) => ({ wch: w }));
  wsN['!cols'] = [12, 9, 9, 14, 8, 58, 14, 7, 6].map((w) => ({ wch: w }));
  wsP['!cols'] = [{ wch: 110 }];
  XLSX.utils.book_append_sheet(wb, wsS, 'SISWA');
  XLSX.utils.book_append_sheet(wb, wsN, 'NILAI');
  XLSX.utils.book_append_sheet(wb, wsP, 'PETUNJUK');
  XLSX.writeFile(wb, 'Format_Massal_Transkrip.xlsx');
}

/* ---------------------------------------------------------------------------
 * 6. TABEL SISWA
 * ------------------------------------------------------------------------- */
function setStudents(list) {
  state.students = list || [];
  const valid = new Set(state.students.map((s) => s.NIS));
  [...state.selected].forEach((n) => { if (!valid.has(n)) state.selected.delete(n); });
  renderStudents();
}

function query() { return $('search').value.trim().toLowerCase(); }

function visibleStudents() {
  const q = query();
  return q ? state.students.filter((s) => [s.Nama_Siswa, s.NIS, s.NISN, s.Nama_Arab].join(' ').toLowerCase().includes(q)) : state.students;
}

function renderStudents() {
  const list = visibleStudents();
  const tb = $('studentRows');
  if (!state.students.length) tb.innerHTML = `<tr><td colspan="6" class="text-center py-12 text-slate-400">Belum ada data. Unggah file Excel/DOCX di langkah 1.</td></tr>`;
  else if (!list.length) tb.innerHTML = `<tr><td colspan="6" class="text-center py-10 text-slate-400">Tidak ada yang cocok dengan pencarian.</td></tr>`;
  else tb.innerHTML = list.map((s) => {
    const sel = state.selected.has(s.NIS);
    const dots = (arr, cls) => [1, 2, 3, 4, 5, 6].map((n) => `<span class="sem-dot ${arr.includes(n) ? cls : 'sem-off'}">${n}</span>`).join(' ');
    const made = (s.transkrip || []).map((c) => { const [m, f] = c.split('-'); return `<span class="chip chip-ok !px-2">${esc(comboLabel(m, f))}</span>`; }).join(' ');
    const warn = !s.lengkap ? `<span class="chip chip-warn !py-0.5 !px-2 ml-1" title="Jenis kelamin / nama orang tua / nomor transkrip belum diisi">identitas belum lengkap</span>` : '';
    return `<tr class="${sel ? 'sel' : ''}">
      <td><input type="checkbox" class="accent-brand-700 w-4 h-4" data-nis="${esc(s.NIS)}" ${sel ? 'checked' : ''} aria-label="Pilih ${esc(s.Nama_Siswa)}"></td>
      <td class="min-w-[220px]"><div class="font-semibold text-slate-900">${esc(s.Nama_Siswa)}${warn}</div>
        <div class="text-xs text-slate-500">NIS ${esc(s.NIS)}${s.NISN ? ' · NISN ' + esc(s.NISN) : ''}${s.Nama_Arab ? ' · <span class="ar text-sm" lang="ar">' + esc(s.Nama_Arab) + '</span>' : ''}</div>
        ${s.Jurusan ? `<div class="text-[11px] text-indigo-700 font-semibold">${esc(s.jurusanNama || s.Jurusan)}</div>` : ''}</td>
      <td class="whitespace-nowrap"><div>${dots(s.semesterTerisi, 'sem-on')}</div>${s.raporTerisi && s.raporTerisi.length ? `<div class="mt-0.5">${dots(s.raporTerisi, 'sem-rap')}</div>` : ''}</td>
      <td class="text-center font-bold tabular-nums text-brand-700">${s.ipk != null ? Number(s.ipk).toFixed(2) : '–'}<div class="text-[10px] font-medium text-slate-500">${esc(s.predikat || '')}</div></td>
      <td class="hidden lg:table-cell text-xs">${made || '<span class="text-slate-400">—</span>'}</td>
      <td class="text-right whitespace-nowrap">
        <button class="btn btn-ghost btn-sm" data-act="preview" data-nis="${esc(s.NIS)}">Pratinjau</button>
        <button class="btn btn-ghost btn-sm" data-act="generate" data-nis="${esc(s.NIS)}" title="Generate dengan bahasa & format terpilih">⚡</button>
        <button class="btn btn-ghost btn-sm" data-act="edit" data-nis="${esc(s.NIS)}" title="Edit identitas & nilai">✎</button>
        <button class="btn btn-danger btn-sm" data-act="delete" data-nis="${esc(s.NIS)}" title="Hapus siswa">✕</button>
      </td></tr>`;
  }).join('');
  $('chkAll').checked = !!(list.length && list.every((s) => state.selected.has(s.NIS)));
  $('selCount').textContent = state.selected.size;
  $('btnGenerate').disabled = !state.selected.size || state.busy;
  $('tableStats').textContent = state.tab === 'students' ? `${list.length} dari ${state.students.length} siswa` : `${visibleHistory().length} transkrip`;
}

async function onStudentAction(e) {
  const btn = e.target.closest('button[data-act]');
  if (!btn) return;
  const nis = btn.dataset.nis, act = btn.dataset.act;
  const st = state.students.find((s) => s.NIS === nis);
  if (act === 'preview') return openPreview(nis, getMode(), getFormat());
  if (act === 'generate') return generateMany([nis], getMode(), getFormat(), getJurusan());
  if (act === 'edit') return openEditor(nis, getMode(), getFormat(), getJurusan());
  if (act === 'delete') {
    const ok = await confirmBox('Hapus siswa?', `${st ? st.Nama_Siswa : nis} beserta seluruh nilainya akan dihapus dari Google Sheets.\nTranskrip yang sudah dibuat tetap ada di tab Riwayat.`, 'Hapus');
    if (!ok) return;
    try {
      const res = await withLoading('Menghapus…', () => call('apiDeleteStudent', nis));
      state.selected.delete(nis); setStudents(res.students); toast(`Dihapus (${res.gradesRemoved} nilai).`);
    } catch (err) { toast(errMsg(err), 'err'); }
  }
}

/* ---------------------------------------------------------------------------
 * 7. RIWAYAT TRANSKRIP (nama yang sudah dibuat transkripnya)
 * ------------------------------------------------------------------------- */
function visibleHistory() {
  const q = query();
  return q ? state.transcripts.filter((t) => [t.Nama_Siswa, t.NIS, t.Nomor_Transkrip].join(' ').toLowerCase().includes(q)) : state.transcripts;
}

function renderHistory() {
  $('histCount').textContent = state.transcripts.length;
  const list = visibleHistory();
  $('historyRows').innerHTML = !state.transcripts.length
    ? `<tr><td colspan="5" class="text-center py-12 text-slate-400">Belum ada transkrip yang dibuat.</td></tr>`
    : !list.length ? `<tr><td colspan="5" class="text-center py-10 text-slate-400">Tidak ada yang cocok.</td></tr>`
    : list.map((t) => `
      <tr>
        <td class="min-w-[200px]"><div class="font-semibold text-slate-900">${esc(t.Nama_Siswa)}</div><div class="text-xs text-slate-500">NIS ${esc(t.NIS)}</div></td>
        <td><span class="chip ${t.Format === 'RAPOR' ? 'chip-info' : 'chip-ok'}">${esc(comboLabel(t.Mode, t.Format))}</span>${t.Jurusan ? `<div class="text-[10px] text-indigo-700 mt-0.5">${esc(jurusanName(t.Jurusan))}</div>` : ''}</td>
        <td class="hidden md:table-cell text-xs">${esc(t.Nomor_Transkrip || '—')}</td>
        <td class="hidden lg:table-cell text-xs text-slate-600">${esc(fmtTime(t.Diperbarui))}<div class="text-[10px] text-slate-400">${Number(t.Jumlah_Generate) > 1 ? t.Jumlah_Generate + '× generate' : ''}${t.Catatan ? ' · ' + esc(t.Catatan) : ''}</div></td>
        <td class="text-right whitespace-nowrap">
          ${t.Doc_URL ? `<a class="btn btn-ghost btn-sm" href="${esc(t.Doc_URL)}" target="_blank" rel="noopener" title="Buka & edit langsung di Google Docs">Edit di Docs ↗</a>
          <button class="btn btn-ghost btn-sm" data-hact="refresh" data-id="${esc(t.ID)}" title="Ekspor ulang PDF & DOCX dari Google Docs yang sudah diedit">⟳ PDF/DOCX</button>` : ''}
          <button class="btn btn-ghost btn-sm" data-hact="edit" data-id="${esc(t.ID)}" title="Edit identitas & nilai lalu generate ulang">Edit data</button>
          <button class="btn btn-primary btn-sm" data-dl="${esc(t.PDF_ID)}">⬇ PDF</button>
          <button class="btn btn-ghost btn-sm" data-dl="${esc(t.DOCX_ID)}">⬇ DOCX</button>
          <button class="btn btn-ghost btn-sm" data-xl="1" data-nis="${esc(t.NIS)}" data-mode="${esc(t.Mode)}" data-format="${esc(t.Format)}"
            data-jur="${esc(t.Jurusan || 'NONE')}" data-name="${esc(excelBaseName(t.NIS, t.Nama_Siswa, t.Mode, t.Format))}" title="Excel disusun dari data nilai (bukan dari editan manual di Google Docs)">⬇ Excel</button>
          <button class="btn btn-danger btn-sm" data-hact="delete" data-id="${esc(t.ID)}" title="Hapus entri & file">✕</button>
        </td>
      </tr>`).join('');
  if (state.tab === 'history') $('tableStats').textContent = `${list.length} transkrip`;
}

function upsertTranscript(entry) {
  if (!entry) return;
  const i = state.transcripts.findIndex((t) => t.ID === entry.ID);
  entry.Diperbarui = new Date().toISOString();
  if (i >= 0) state.transcripts.splice(i, 1);
  state.transcripts.unshift(entry);
  const st = state.students.find((s) => s.NIS === String(entry.NIS));
  const code = entry.Mode + '-' + entry.Format;
  if (st && !(st.transkrip || []).includes(code)) st.transkrip = (st.transkrip || []).concat(code);
  renderHistory(); renderStudents();
}

async function onHistoryAction(e) {
  const dl = e.target.closest('button[data-dl], button[data-xl]');
  if (dl) return onDownloadClick(e);
  const btn = e.target.closest('button[data-hact]');
  if (!btn) return;
  const t = state.transcripts.find((x) => x.ID === btn.dataset.id);
  if (!t) return;
  if (btn.dataset.hact === 'refresh') {
    try {
      const res = await withLoading('Mengekspor ulang dari Google Docs…', () => call('apiRefreshFromDoc', t.ID), t.Nama_Siswa);
      upsertTranscript(res.entry); addResult(res); toast('PDF & DOCX diperbarui dari Google Docs.');
    } catch (err) { toast(errMsg(err), 'err', 8000); }
  } else if (btn.dataset.hact === 'edit') {
    openEditor(t.NIS, t.Mode, t.Format, t.Jurusan || 'NONE');
  } else if (btn.dataset.hact === 'delete') {
    if (!(await confirmBox('Hapus transkrip?', `${t.Nama_Siswa} · ${comboLabel(t.Mode, t.Format)}\nGoogle Docs, PDF, dan DOCX-nya dipindahkan ke Sampah Drive.`, 'Hapus'))) return;
    try {
      state.transcripts = await withLoading('Menghapus…', () => call('apiDeleteTranscript', t.ID));
      const st = state.students.find((s) => s.NIS === String(t.NIS));
      if (st) st.transkrip = (st.transkrip || []).filter((c) => c !== t.Mode + '-' + t.Format);
      renderHistory(); renderStudents(); toast('Transkrip dihapus.');
    } catch (err) { toast(errMsg(err), 'err'); }
  }
}

/* ---------------------------------------------------------------------------
 * 8. PRATINJAU
 * ------------------------------------------------------------------------- */
const COMBOS = [['TRI', 'TRI'], ['ID', 'GPA'], ['EN', 'GPA'], ['AR', 'GPA'], ['ID', 'RAPOR'], ['EN', 'RAPOR'], ['AR', 'RAPOR']];

async function openPreview(nis, mode, format) {
  state.preview = { nis, mode, format };
  $('previewModes').innerHTML = COMBOS.map(([m, f]) =>
    `<button type="button" data-m="${m}" data-f="${f}" class="px-2.5 py-1.5 rounded-md ${m === mode && f === format ? 'bg-white shadow text-brand-800' : 'text-slate-500 hover:text-slate-800'}">${esc(comboLabel(m, f))}</button>`).join('');
  $('previewModes').querySelectorAll('button').forEach((b) => b.addEventListener('click', () => openPreview(nis, b.dataset.m, b.dataset.f)));
  try {
    const d = await withLoading('Menyusun pratinjau…', () => call('apiGetTranscriptPreview', nis, mode, format, getJurusan()));
    state.preview.data = d;
    $('previewBody').innerHTML = renderPreview(d);
    $('previewBody').setAttribute('dir', d.rtl ? 'rtl' : 'ltr');
    openModal('modalPreview');
  } catch (e) { toast(errMsg(e), 'err'); }
}

function renderPreview(d) {
  const stack = (arr) => (Array.isArray(arr) ? arr : [arr]).filter((x) => x !== '' && x != null)
    .map((x) => `<div class="${isAr(x) ? 'ar text-[13px]' : ''}" ${isAr(x) ? 'dir="rtl"' : ''}>${esc(x)}</div>`).join('');
  const head = `
    <div class="text-center mb-4">
      ${d.title.map((x, i) => `<div class="${i === 0 ? 'text-xl font-extrabold text-brand-700 tracking-wide underline' : 'font-bold text-slate-700'} ${isAr(x) ? 'ar text-2xl' : ''}">${esc(x)}</div>`).join('')}
      <div class="text-xs text-slate-500 mt-1">${esc(d.nomorLabel)} ${esc(d.nomor || '—')}</div>
      <div class="text-[11px] mt-1 ${d.jurusan ? 'text-indigo-700 font-semibold' : 'text-slate-400'}">${d.jurusan ? 'Struktur mapel: ' + esc(d.jurusan) : 'Tanpa struktur kurikulum (semua mapel yang ada nilainya)'}</div>
      ${d.susunan ? `<div class="text-[11px] text-emerald-700 font-semibold">Urutan mapel mengikuti rapor (sheet Susunan_Mapel) — ${esc(d.susunan)}</div>` : '<div class="text-[11px] text-slate-400">Urutan mapel mengikuti file / kurikulum (SUSUNAN_MAPEL = FILE)</div>'}
    </div>
    <div class="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-1 text-sm mb-5 border border-brand-200 rounded-lg p-3 bg-brand-50/50">
      ${d.identity.map(([l, v]) => `<div class="flex gap-2"><div class="w-44 shrink-0 text-xs text-slate-500">${stack(l)}</div><div class="font-semibold">${esc(v || '—')}</div></div>`).join('')}
    </div>
    ${d.issues && d.issues.length ? `<div class="mb-4 text-xs rounded-lg bg-amber-50 text-amber-800 px-3 py-2">${d.issues.map(esc).join('<br>')}</div>` : ''}`;

  if (d.kind === 'RAPOR') {
    return head + `
      <table>
        <thead>
          <tr><th rowspan="2" class="w-10">${esc(d.head.no || 'No.')}</th><th rowspan="2" class="text-left">${esc(d.head.mapel)}</th>${d.head.kelas.map((k) => `<th colspan="2">${esc(k)}</th>`).join('')}</tr>
          <tr>${[1, 2, 3, 4, 5, 6].map((n) => `<th>${n}</th>`).join('')}</tr>
        </thead>
        <tbody>${d.rows.map((r) => r.grup ? `<tr class="grp"><td colspan="8">${esc(r.grup)}</td></tr>`
          : `<tr><td class="num">${esc(r.no || '')}</td><td>${esc(r.mapel)}</td>${r.n.map((v) => `<td class="num">${esc(v)}</td>`).join('')}</tr>`).join('')}
          <tr class="foot"><td colspan="2">${esc(d.head.avg)}</td>${d.avg.map((v) => `<td class="num">${esc(v)}</td>`).join('')}</tr>
        </tbody>
      </table>`;
  }
  const c = d.cols;
  const sem = (s) => `
    <div class="min-w-0"><table>
      <thead>
        <tr><th colspan="${d.showNo ? 7 : 6}">${stack(s.judul)}</th></tr>
        <tr>${d.showNo ? `<th>${stack(c.no)}</th>` : ''}<th>${stack(c.mapel)}</th><th>${stack(c.nilai)}</th><th>${stack(c.pred)}</th><th>${stack(c.bobot)}</th><th>${stack(c.sks)}</th><th>${stack(c.mutu)}</th></tr>
      </thead>
      <tbody>${s.rows.map((r) => `<tr>${d.showNo ? `<td class="num">${r.no}</td>` : ''}<td>${esc(r.mapel)}</td><td class="num">${esc(r.nilai)}</td><td class="num font-bold">${esc(r.pred)}</td><td class="num">${esc(r.bobot)}</td><td class="num">${esc(r.sks)}</td><td class="num">${esc(r.mutu)}</td></tr>`).join('')}</tbody>
      <tfoot>
        <tr><td colspan="${d.showNo ? 5 : 4}" class="text-center">${esc(d.jumlah)}</td><td class="num">${esc(s.totalSks)}</td><td class="num">${esc(s.totalMutu)}</td></tr>
        <tr><td colspan="${d.showNo ? 5 : 4}" class="text-center">${esc(d.ipsLabel)}</td><td colspan="2" class="num text-brand-700 text-sm">${esc(s.ips)}</td></tr>
      </tfoot>
    </table></div>`;
  return head + `
    <div class="grid grid-cols-1 2xl:grid-cols-2 gap-4">${d.semesters.map(sem).join('')}</div>
    <div class="mt-5 max-w-xl"><table><tbody>
      ${d.cumulative.map(([l, v], i) => `<tr><td class="font-semibold">${esc(l)}</td><td class="num ${i === 2 ? 'font-extrabold text-brand-700 text-base' : 'font-bold'}">${esc(v)}</td></tr>`).join('')}
    </tbody></table></div>`;
}

/* ---------------------------------------------------------------------------
 * 9. GENERATE & UNDUH
 * ------------------------------------------------------------------------- */
async function generateMany(nisList, mode, format, jurusan) {
  jurusan = jurusan || '';
  nisList = nisList.filter(Boolean);
  if (!nisList.length || state.busy) return;
  state.busy = true; renderStudents();
  const total = nisList.length;
  const bar = $('genProgressBar');
  $('genProgress').classList.toggle('hidden', total < 2);
  let ok = 0;
  for (let i = 0; i < total; i++) {
    const nis = nisList[i];
    const st = state.students.find((s) => s.NIS === nis) || { Nama_Siswa: nis };
    $('genProgressText').textContent = `${i + 1}/${total} · ${st.Nama_Siswa}`;
    const pct = Math.round((i / total) * 100); bar.style.width = pct + '%'; $('genProgressPct').textContent = pct + '%';
    try {
      const run = () => call('apiGenerate', nis, mode, format, jurusan);
      const res = total === 1 ? await withLoading('Membuat transkrip…', run, `${st.Nama_Siswa} · ${comboLabel(mode, format)} — biasanya 15–40 detik`) : await run();
      addResult(res); upsertTranscript(res.entry); ok++;
    } catch (e) { addResultError(st.Nama_Siswa, errMsg(e)); }
  }
  bar.style.width = '100%'; $('genProgressPct').textContent = '100%';
  setTimeout(() => $('genProgress').classList.add('hidden'), 1200);
  state.busy = false; renderStudents();
  toast(`${ok} dari ${total} transkrip berhasil dibuat.`, ok === total ? 'ok' : 'warn');
}

function addResult(r) {
  const el = document.createElement('div');
  el.className = 'border border-brand-200 rounded-xl p-3 bg-white';
  el.innerHTML = `
    <div class="flex items-start gap-2">
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm text-slate-900 truncate">${esc(r.nama)}</div>
        <div class="text-[11px] text-slate-500">${esc(comboLabel(r.mode, r.format))} · ${r.durationSec}s</div>
        ${r.jurusan ? `<div class="text-[11px] text-indigo-700">${esc(r.jurusan)}</div>` : ''}
      </div>
      <span class="chip chip-ok">Selesai</span>
    </div>
    <div class="mt-2 grid grid-cols-3 gap-2">
      <button class="btn btn-primary btn-sm" data-dl="${esc(r.pdf.id)}" type="button">⬇ PDF</button>
      <button class="btn btn-ghost btn-sm" data-dl="${esc(r.docx.id)}" type="button">⬇ DOCX</button>
      <button class="btn btn-ghost btn-sm" data-xl="1" data-nis="${esc(r.nis)}" data-mode="${esc(r.mode)}" data-format="${esc(r.format)}"
        data-jur="${esc((r.entry && r.entry.Jurusan) || 'NONE')}" data-name="${esc(String(r.pdf.name || '').replace(/\.pdf$/i, ''))}" type="button">⬇ Excel</button>
    </div>
    <div class="mt-1.5 flex flex-wrap gap-x-3 text-[11px]">
      ${r.doc ? `<a class="text-brand-600 font-semibold hover:underline" href="${esc(r.doc.url)}" target="_blank" rel="noopener">Edit di Google Docs ↗</a>` : ''}
      <a class="text-brand-600 hover:underline" href="${esc(r.pdf.url)}" target="_blank" rel="noopener">PDF di Drive</a>
      <a class="text-brand-600 hover:underline" href="${esc(r.docx.url)}" target="_blank" rel="noopener">DOCX di Drive</a>
    </div>
    ${r.warnings && r.warnings.length ? `<div class="mt-2 text-[11px] text-amber-700">${r.warnings.slice(0, 5).map(esc).join('<br>')}</div>` : ''}`;
  $('genResults').prepend(el);
}

function addResultError(name, msg) {
  const el = document.createElement('div');
  el.className = 'border border-red-200 bg-red-50 rounded-xl p-3 text-xs text-red-700';
  el.innerHTML = `<b>${esc(name)}</b><br>${esc(msg)}`;
  $('genResults').prepend(el);
}

async function onDownloadClick(e) {
  const xl = e.target.closest('button[data-xl]');
  if (xl) return downloadExcelFromButton(xl);
  const btn = e.target.closest('button[data-dl]');
  if (!btn || !btn.dataset.dl) return;
  const label = btn.textContent;
  btn.disabled = true; btn.innerHTML = '<span class="spinner spinner-sm"></span>';
  try { const f = await call('apiDownload', btn.dataset.dl); saveBase64(f.base64, f.mimeType, f.name); }
  catch (err) { toast(errMsg(err), 'err'); }
  finally { btn.disabled = false; btn.textContent = label; }
}

/* ---------------------------------------------------------------------------
 * 9b. UNDUH EXCEL (disusun di browser dengan SheetJS dari data transkrip)
 * ------------------------------------------------------------------------- */
function excelBaseName(nis, nama, mode, format) {
  const code = format === 'TRI' ? '3B' : `${mode}-${format}`;
  return `Transkrip_${code}_${nis}_${String(nama || '').replace(/[\\/:*?"<>|#%]+/g, '').replace(/\s+/g, '_').slice(0, 60)}`;
}

async function downloadExcelFromButton(btn) {
  const label = btn.textContent;
  btn.disabled = true; btn.innerHTML = '<span class="spinner spinner-sm"></span>';
  try {
    const d = await call('apiGetTranscriptPreview', btn.dataset.nis, btn.dataset.mode, btn.dataset.format, btn.dataset.jur || 'NONE');
    saveExcel(d, btn.dataset.name || excelBaseName(btn.dataset.nis, '', btn.dataset.mode, btn.dataset.format));
  } catch (err) { toast(errMsg(err), 'err'); }
  finally { btn.disabled = false; btn.textContent = label; }
}

/** Ubah "6,66" / "3.41" / "88" menjadi angka Excel (dengan format desimal yang sama). */
function xlCell(v) {
  if (v === null || v === undefined) return { v: '', t: 's' };
  const s = String(v).trim();
  if (/^-?\d+([.,]\d+)?$/.test(s)) {
    const dec = (s.split(/[.,]/)[1] || '').length;
    return { v: Number(s.replace(',', '.')), t: 'n', z: dec ? '0.' + '0'.repeat(dec) : '0' };
  }
  return { v: s, t: 's' };
}

function buildTranscriptSheet(d) {
  const rows = [], merges = [];
  const width = d.kind === 'RAPOR' ? 8 : (d.showNo ? 7 : 6);
  const line = (arr) => (Array.isArray(arr) ? arr : [arr]).filter((x) => x !== '' && x != null).join(' / ');
  const push = (cells, mergeAll) => {
    const r = rows.length;
    rows.push(cells.map((c) => (typeof c === 'object' && c !== null && 'v' in c ? c : xlCell(c))));
    if (mergeAll) merges.push({ s: { r, c: 0 }, e: { r, c: width - 1 } });
  };
  d.title.forEach((t) => push([t], true));
  push([`${d.nomorLabel} ${d.nomor || ''}`.trim()], true);
  if (d.jurusan) push(['Struktur mapel: ' + d.jurusan], true);
  push([]);
  d.identity.forEach(([label, value]) => {
    const r = rows.length;
    push([line(label), value || '']);
    merges.push({ s: { r, c: 1 }, e: { r, c: width - 1 } });
  });
  push([]);

  if (d.kind === 'RAPOR') {
    let r = rows.length;
    push([d.head.no || 'No.', d.head.mapel, d.head.kelas[0], '', d.head.kelas[1], '', d.head.kelas[2], '']);
    merges.push({ s: { r, c: 0 }, e: { r: r + 1, c: 0 } }, { s: { r, c: 1 }, e: { r: r + 1, c: 1 } },
      { s: { r, c: 2 }, e: { r, c: 3 } }, { s: { r, c: 4 }, e: { r, c: 5 } }, { s: { r, c: 6 }, e: { r, c: 7 } });
    push(['', '', 1, 2, 3, 4, 5, 6]);
    d.rows.forEach((x) => (x.grup ? push([x.grup], true) : push([x.no === undefined ? '' : Number(x.no) || x.no, x.mapel].concat(x.n))));
    r = rows.length;
    push([d.head.avg, ''].concat(d.avg));
    merges.push({ s: { r, c: 0 }, e: { r, c: 1 } });
  } else {
    const c = d.cols;
    d.semesters.forEach((sm) => {
      push([line(sm.judul)], true);
      push((d.showNo ? [line(c.no)] : []).concat([line(c.mapel), line(c.nilai), line(c.pred), line(c.bobot), line(c.sks), line(c.mutu)]));
      sm.rows.forEach((x) => push((d.showNo ? [x.no] : []).concat([x.mapel, x.nilai, x.pred, x.bobot, x.sks, x.mutu])));
      let r = rows.length;
      push([d.jumlah].concat(Array(width - 3).fill('')).concat([sm.totalSks, sm.totalMutu]));
      merges.push({ s: { r, c: 0 }, e: { r, c: width - 3 } });
      r = rows.length;
      push([d.ipsLabel].concat(Array(width - 3).fill('')).concat([sm.ips, '']));
      merges.push({ s: { r, c: 0 }, e: { r, c: width - 3 } }, { s: { r, c: width - 2 }, e: { r, c: width - 1 } });
      push([]);
    });
    d.cumulative.forEach(([label, value]) => {
      const r = rows.length;
      push([label, '', '', '', value]);
      merges.push({ s: { r, c: 0 }, e: { r, c: 3 } });
    });
  }
  if (d.issues && d.issues.length) { push([]); d.issues.forEach((i) => push(['Catatan: ' + i], true)); }

  const ws = {};
  let maxC = 0;
  rows.forEach((cells, r) => cells.forEach((cell, c) => {
    if (cell.v === '' && cell.t === 's') return;
    ws[XLSX.utils.encode_cell({ r, c })] = cell;
    maxC = Math.max(maxC, c);
  }));
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(rows.length - 1, 0), c: Math.max(maxC, width - 1) } });
  ws['!merges'] = merges;
  ws['!cols'] = d.kind === 'RAPOR' ? [{ wch: 5 }, { wch: 46 }].concat(Array(6).fill({ wch: 9 }))
    : (d.showNo ? [{ wch: 5 }] : []).concat([{ wch: 46 }, { wch: 9 }, { wch: 10 }, { wch: 9 }, { wch: 8 }, { wch: 10 }]);
  return ws;
}

function saveExcel(d, baseName) {
  if (typeof XLSX === 'undefined') { toast('Library SheetJS gagal dimuat.', 'err'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildTranscriptSheet(d), 'Transkrip');
  if (d.rtl) wb.Workbook = { Views: [{ RTL: true }] };
  XLSX.writeFile(wb, baseName + '.xlsx');
}

/* ---------------------------------------------------------------------------
 * 10. FORM IDENTITAS SISWA (tambah siswa baru)
 * ------------------------------------------------------------------------- */
const STUDENT_FIELDS = [
  { k: 'NIS', label: 'NIS *', req: true }, { k: 'NISN', label: 'NISN' },
  { k: 'Nama_Siswa', label: 'Nama Siswa (Latin) *', req: true, wide: true },
  { k: 'Nama_Arab', label: 'Nama huruf Arab (opsional)', wide: true, ar: true },
  { k: 'Jenis_Kelamin', label: 'Jenis Kelamin', type: 'gender' }, { k: 'Jurusan', label: 'Jurusan (struktur kurikulum)', type: 'jurusan' },
  { k: 'Nama_Ortu', label: 'Nama Orang Tua' },
  { k: 'Tempat_Lahir', label: 'Tempat Lahir' }, { k: 'Tanggal_Lahir', label: 'Tanggal Lahir', type: 'date' },
  { k: 'Masa_Studi', label: 'Masa Studi', ph: '2023/2024 – 2025/2026' }, { k: 'Catatan', label: 'Catatan' },
  { k: 'Nomor_Transkrip', label: 'Nomor Transkrip (3 Bahasa & GPA)', ph: '09.471/SMAIT/YKPI/V/2026' }, { k: 'Tanggal_Terbit', label: 'Tanggal Terbit (3 Bahasa & GPA)', type: 'date' },
  { k: 'Nomor_Rapor', label: 'Nomor Transkrip Rapor', ph: '09.498/SMAIT/YKPI/VI/2026' }, { k: 'Tanggal_Rapor', label: 'Tanggal Terbit Rapor', type: 'date' }
];

function fieldHtml(f, prefix, value) {
  const id = prefix + f.k;
  const v = esc(value == null ? '' : value);
  let input;
  if (f.type === 'jurusan') input = `<select id="${id}" class="input"><option value="">— tanpa jurusan —</option>${state.jurusan.map((j) => `<option value="${esc(j.Kode)}" ${j.Kode === value ? 'selected' : ''}>${esc(j.Nama)}</option>`).join('')}${value && !state.jurusan.some((j) => j.Kode === value) ? `<option value="${esc(value)}" selected>${esc(value)} (tidak aktif)</option>` : ''}</select>`;
  else if (f.type === 'gender') input = `<select id="${id}" class="input"><option value="">—</option><option value="L" ${value === 'L' ? 'selected' : ''}>Laki-laki</option><option value="P" ${value === 'P' ? 'selected' : ''}>Perempuan</option></select>`;
  else input = `<input id="${id}" class="input ${f.ar ? 'ar text-lg text-right' : ''}" ${f.ar ? 'dir="rtl" lang="ar"' : ''} type="${f.type === 'date' ? 'date' : 'text'}" value="${v}" ${f.req ? 'required' : ''} placeholder="${esc(f.ph || '')}">`;
  return `<div class="${f.wide ? 'sm:col-span-2' : ''}"><label class="label" for="${id}">${esc(f.label)}</label>${input}</div>`;
}

function renderStudentForm() { $('studentFields').innerHTML = STUDENT_FIELDS.map((f) => fieldHtml(f, 'f_', '')).join(''); }

function openStudentForm() {
  renderStudentForm();
  $('studentFormTitle').textContent = 'Tambah Siswa';
  openModal('modalStudent');
}

async function saveStudentForm(e) {
  e.preventDefault();
  const obj = {};
  STUDENT_FIELDS.forEach((f) => { obj[f.k] = $('f_' + f.k).value.trim(); });
  try {
    const res = await withLoading('Menyimpan…', () => call('apiSaveStudent', obj));
    setStudents(res.students); closeModal('modalStudent'); toast('Data siswa tersimpan.');
  } catch (err) { toast(errMsg(err), 'err'); }
}

/* ---------------------------------------------------------------------------
 * 11. EDIT TRANSKRIP (identitas + nilai) → simpan → generate ulang
 * ------------------------------------------------------------------------- */
async function openEditor(nis, mode, format, jurusan) {
  try {
    const d = await withLoading('Memuat data transkrip…', () => call('apiGetEditor', nis, format));
    state.editor = { nis, mode, format, jurusan: jurusan || '', data: d, sem: (d.semesters.find((s) => s.rows.length) || { sem: 1 }).sem };
    state.subjects = d.subjects || state.subjects;
    $('editorInfo').textContent = `${d.student.Nama_Siswa} · NIS ${nis} · ${comboLabel(mode, format)} · data nilai: ${d.kategori === 'RAPOR' ? 'Rapor' : 'GPA (dengan SKS)'}`;
    const note = $('editorNote');
    note.classList.toggle('hidden', !d.fallback);
    note.textContent = d.fallback ? 'Siswa ini belum punya data nilai rapor terpisah, jadi transkrip Rapor memakai data GPA. Perubahan di sini juga mengubah transkrip 3 Bahasa & GPA. Unggah DOCX/Excel rapor untuk memisahkannya.' : '';
    $('editorFields').innerHTML = STUDENT_FIELDS.filter((f) => f.k !== 'NIS' && f.k !== 'Catatan')
      .map((f) => fieldHtml(Object.assign({}, f, { wide: false }), 'e_', d.student[f.k])).join('');
    renderEditorGrades();
    openModal('modalEditor');
  } catch (e) { toast(errMsg(e), 'err'); }
}

function subjectOptions(selected) {
  const opts = state.subjects.map((s) => `<option value="${esc(s.Kode_Mapel)}" ${String(s.Kode_Mapel) === String(selected) ? 'selected' : ''}>${esc(s.Kode_Mapel)} — ${esc(s.Nama_ID)}</option>`);
  if (selected && !state.subjects.some((s) => String(s.Kode_Mapel) === String(selected))) opts.unshift(`<option value="${esc(selected)}" selected>${esc(selected)}</option>`);
  return `<option value="">— pilih mapel —</option>` + opts.join('');
}

function renderEditorGrades() {
  const ed = state.editor, d = ed.data;
  const rapor = d.kategori === 'RAPOR';
  $('editorSemTabs').innerHTML = d.semesters.map((s) =>
    `<button type="button" data-sem="${s.sem}" class="px-3 py-1 rounded-lg text-xs font-bold border ${s.sem === ed.sem ? 'bg-brand-700 text-white border-brand-700' : 'bg-white text-slate-600 border-slate-200'}">Sem ${s.sem}<span class="opacity-70"> (${s.rows.length})</span></button>`).join('');
  const s = d.semesters[ed.sem - 1];
  $('editorGrades').innerHTML = `
    <div class="flex flex-wrap items-end gap-3 mb-2">
      <div><label class="label" for="e_tahun">Tahun pelajaran semester ${s.sem}</label><input id="e_tahun" class="input input-sm w-40" value="${esc(s.tahun)}" placeholder="2023/2024"></div>
      <span class="text-xs text-slate-500">Urutan baris = urutan di transkrip. Kosongkan semua baris untuk menghapus semester ini.</span>
    </div>
    <div class="border border-slate-200 rounded-xl overflow-auto">
      <table class="data grade-grid">
        <thead><tr><th class="w-8">#</th><th>Mata pelajaran</th>${rapor ? '<th>Nama di transkrip (opsional)</th>' : ''}<th class="w-20">Nilai</th>${rapor ? '' : '<th class="w-16">SKS</th>'}<th class="w-28"></th></tr></thead>
        <tbody>${s.rows.map((r, i) => `
          <tr data-i="${i}">
            <td class="text-xs text-slate-400">${i + 1}</td>
            <td><select class="input input-sm" data-f="kode">${subjectOptions(r.kode)}</select></td>
            ${rapor ? `<td><input class="input input-sm" data-f="tampil" value="${esc(r.tampil)}"></td>` : ''}
            <td><input class="input input-sm text-center" data-f="nilai" inputmode="decimal" value="${esc(r.nilai)}"></td>
            ${rapor ? '' : `<td><input class="input input-sm text-center" data-f="sks" inputmode="numeric" value="${esc(r.sks)}"></td>`}
            <td class="text-right whitespace-nowrap">
              <button type="button" class="btn btn-ghost btn-sm" data-g="up" title="Naik">↑</button><button type="button" class="btn btn-ghost btn-sm" data-g="down" title="Turun">↓</button><button type="button" class="btn btn-danger btn-sm" data-g="del" title="Hapus baris">✕</button>
            </td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
    <button type="button" class="btn btn-ghost btn-sm mt-2" data-g="add">+ Tambah baris</button>`;
}

function readEditorGrid() {
  const ed = state.editor;
  if (!ed) return;
  const s = ed.data.semesters[ed.sem - 1];
  const tahun = $('e_tahun'); if (tahun) s.tahun = tahun.value.trim();
  document.querySelectorAll('#editorGrades tbody tr[data-i]').forEach((tr) => {
    const r = s.rows[Number(tr.dataset.i)];
    tr.querySelectorAll('[data-f]').forEach((inp) => { r[inp.dataset.f] = inp.value.trim(); });
  });
}

function onEditorGridClick(e) {
  const b = e.target.closest('button[data-g]');
  if (!b) return;
  readEditorGrid();
  const s = state.editor.data.semesters[state.editor.sem - 1];
  const tr = b.closest('tr');
  const i = tr ? Number(tr.dataset.i) : -1;
  if (b.dataset.g === 'add') s.rows.push({ kode: '', tampil: '', nilai: '', sks: state.editor.data.kategori === 'RAPOR' ? '' : 2 });
  if (b.dataset.g === 'del') s.rows.splice(i, 1);
  if (b.dataset.g === 'up' && i > 0) s.rows.splice(i - 1, 0, s.rows.splice(i, 1)[0]);
  if (b.dataset.g === 'down' && i < s.rows.length - 1) s.rows.splice(i + 1, 0, s.rows.splice(i, 1)[0]);
  renderEditorGrades();
}

async function saveEditor(andGenerate) {
  const ed = state.editor;
  readEditorGrid();
  const student = {};
  STUDENT_FIELDS.filter((f) => f.k !== 'NIS' && f.k !== 'Catatan').forEach((f) => { student[f.k] = $('e_' + f.k).value.trim(); });
  if (!student.Nama_Siswa) { toast('Nama siswa wajib diisi.', 'err'); return; }
  const payload = { nis: ed.nis, kategori: ed.data.kategori, student, semesters: ed.data.semesters };
  try {
    const res = await withLoading('Menyimpan perubahan…', () => call('apiSaveEditor', payload));
    setStudents(res.students);
    closeModal('modalEditor');
    // struktur mapel mengikuti kolom Jurusan di form ini (kosong = tanpa struktur kurikulum)
    if (andGenerate) await generateMany([ed.nis], ed.mode, ed.format, student.Jurusan || 'NONE');
    else toast('Perubahan tersimpan. Generate ulang agar transkrip ikut berubah.');
  } catch (e) { toast(errMsg(e), 'err', 9000); }
}

/* ---------------------------------------------------------------------------
 * 12. MATA PELAJARAN
 * ------------------------------------------------------------------------- */
function renderSubjects() {
  const kel = (v) => `<select class="input input-sm" data-f="Kelompok">${['', 'Wajib', 'Pilihan', 'Muatan Lokal'].map((k) => `<option ${String(v || '') === k ? 'selected' : ''} value="${k}">${k || '—'}</option>`).join('')}</select>`;
  const row = (s, isNew) => `
    <tr>
      <td><input class="input input-sm font-mono text-xs" data-f="Kode_Mapel" value="${esc(s.Kode_Mapel || '')}" ${isNew ? '' : 'readonly'} placeholder="KODE"></td>
      <td><input class="input input-sm" data-f="Nama_ID" value="${esc(s.Nama_ID || '')}"></td>
      <td><input class="input input-sm" data-f="Nama_EN" value="${esc(s.Nama_EN || '')}"></td>
      <td><input class="input input-sm" data-f="Nama_EN_GPA" value="${esc(s.Nama_EN_GPA || '')}"></td>
      <td><input class="input input-sm ar text-base text-right" dir="rtl" lang="ar" data-f="Nama_AR" value="${esc(s.Nama_AR || '')}"></td>
      <td>${kel(s.Kelompok)}</td>
      <td><input class="input input-sm" data-f="Alias" value="${esc(s.Alias || '')}"></td>
      <td><input class="input input-sm" data-f="Setara" value="${esc(s.Setara || '')}" placeholder="PABP|…"></td>
      <td><select class="input input-sm" data-f="Tampil_Transkrip"><option value="YA" ${String(s.Tampil_Transkrip || 'YA') !== 'TIDAK' ? 'selected' : ''}>Ya</option><option value="TIDAK" ${String(s.Tampil_Transkrip) === 'TIDAK' ? 'selected' : ''}>Tidak</option></select></td>
      <td><input class="input input-sm text-center" data-f="Urutan_Default" value="${esc(s.Urutan_Default == null ? '' : s.Urutan_Default)}"></td>
      <td class="text-right whitespace-nowrap">
        <button class="btn btn-primary btn-sm" data-sact="save" type="button">${isNew ? 'Tambah' : 'Simpan'}</button>
        ${isNew ? '' : '<button class="btn btn-danger btn-sm" data-sact="del" type="button">✕</button>'}
      </td>
    </tr>`;
  $('subjectRows').innerHTML = state.subjects.map((s) => row(s, false)).join('') + row({}, true);
}

async function onSubjectAction(e) {
  const btn = e.target.closest('button[data-sact]');
  if (!btn) return;
  const obj = {};
  btn.closest('tr').querySelectorAll('[data-f]').forEach((i) => { obj[i.dataset.f] = i.value.trim(); });
  try {
    if (btn.dataset.sact === 'del') {
      if (!(await confirmBox('Hapus mata pelajaran?', `${obj.Kode_Mapel} — ${obj.Nama_ID}`, 'Hapus'))) return;
      state.subjects = await withLoading('Menghapus…', () => call('apiDeleteSubject', obj.Kode_Mapel));
    } else {
      state.subjects = await withLoading('Menyimpan…', () => call('apiSaveSubject', obj));
    }
    renderSubjects(); toast('Mata pelajaran diperbarui.');
  } catch (err) { toast(errMsg(err), 'err', 7000); }
}
