// extra.js
// Sheet tambahan (tabel & grafik yang bentuknya khusus, tidak bisa memakai agg() biasa):
//   7. Perubahan per Kecamatan      <- GQREKLAS (Berubah / Tidak Berubah) x WADMKC
//   8. Perbandingan Penggunaan Tanah <- GNAME25 (2014) vs QNAME25 (2026)
//   9. Perubahan Terluas             <- GQNAME ("A menjadi B")
//  10. Reklasifikasi                 <- GREKLAS (2014) vs QREKLAS (2026)
// Tiap sheet = satu objek di array EXTRA di bagian bawah file ini:
//   build()  -> menghitung data (null = '(Data tidak ditemukan)')
//   html(d)  -> HTML tabel + kanvas grafik untuk tampilan web
//   chart()  -> membuat grafik Chart.js (boleh kosong jika tidak ada grafik)
//   excel()  -> menulis sheet Excel

let EXTRA_MODELS = [];
const PAL2 = PAL.concat(['#ff9da7', '#9c755f', '#bab0ab', '#59a14f']);

// ---------- Helper kecil ----------
const fd = v => fmt(Math.abs(v) < 0.005 ? 0 : v);       // hindari tampilan "-0.00"
const pc = (x, t) => (t ? x / t * 100 : 0);
const th = (t, o = '') => `<th ${o}>${t}</th>`;
const td = (v, c = '') => `<td class="${c}">${v}</td>`;
const cvTag = ex => `<div class=cv><canvas width=${ex.w} height=${ex.h}></canvas></div>`;
const kecOf = r => (COLS.includes('WADMKC') && r.WADMKC) || '-';
const area = r => parseFloat(r.LUASHA) || 0;

// jumlah luas per kategori untuk 1 kolom -> {kategori: luas}, atau null jika kolom tidak ada/kosong
function sumBy(f) {
  if (!ROWS.length || !COLS.includes(f)) return null;
  const m = {}; let any = false;
  for (const r of ROWS) { if (!r[f]) continue; any = true; m[r[f]] = (m[r[f]] || 0) + area(r); }
  return any ? m : null;
}

// ---------- Excel helpers ----------
// tulis sel header (bisa digabung) dengan gaya biru tua
function H(ws, r1, c1, r2, c2, v) {
  if (r1 !== r2 || c1 !== c2) ws.mergeCells(r1, c1, r2, c2);
  cell(ws, r1, c1, v);
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) hdr(ws.getCell(r, c));
}
// sisipkan grafik sebagai gambar PNG
function addChart(wb, ws, ex, d, col, row) {
  const cv = document.createElement('canvas'); cv.width = ex.w; cv.height = ex.h;
  const ch = ex.chart(cv, d, false), url = cv.toDataURL('image/png'); ch.destroy();
  ws.addImage(wb.addImage({ base64: url, extension: 'png' }), { tl: { col, row }, ext: { width: ex.w, height: ex.h } });
}

// ---------- Plugin label nilai untuk bar chart multi-seri ----------
const multiVals = {
  id: 'mv', afterDatasetsDraw(ch) {
    const c = ch.ctx; c.save(); c.font = '11px Arial'; c.textAlign = 'center'; c.fillStyle = '#222';
    ch.data.datasets.forEach((ds, i) => ch.getDatasetMeta(i).data.forEach((el, j) => {
      const v = ds.data[j]; c.fillText(fmt(v), el.x, v < 0 ? el.y + 13 : el.y - 5);
    })); c.restore();
  }
};
const baseOpt = (title, a, legend) => ({
  responsive: false, animation: a ? undefined : false, layout: { padding: { top: 18 } },
  plugins: { legend: { display: !!legend, position: legend, labels: { boxWidth: 10, font: { size: 10 } } }, title: { display: true, text: title, font: { size: 12 } } }
});

// =====================================================================
// SHEET 7 - Perubahan per Kecamatan
// =====================================================================
function build7() {
  if (!ROWS.length || !COLS.includes('GQREKLAS')) return null;
  const m = {}; let any = false;
  for (const r of ROWS) {
    const g = (r.GQREKLAS || '').toLowerCase(); if (!g) continue; any = true;
    const k = kecOf(r); m[k] ??= { b: 0, tb: 0 };
    if (g.startsWith('tidak')) m[k].tb += area(r); else m[k].b += area(r);   // "Tidak Berubah" vs "Berubah"
  }
  if (!any) return null;
  const rows = Object.keys(m).sort().map(k => ({ k, b: m[k].b, tb: m[k].tb, t: m[k].b + m[k].tb }));
  const tot = rows.reduce((a, r) => ({ b: a.b + r.b, tb: a.tb + r.tb, t: a.t + r.t }), { b: 0, tb: 0, t: 0 });
  return { rows, tot };
}
const html7 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=3')}${th('Kecamatan', 'rowspan=3')}${th('Perubahan Penggunaan Tanah', 'colspan=4')}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${th('Berubah', 'colspan=2')}${th('Tidak Berubah', 'colspan=2')}</tr><tr>${['Luas (Ha)', 'Luas (%)', 'Luas (Ha)', 'Luas (%)'].map(t => th(t)).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.k)}${td(fmt(r.b), 'n')}${td(fmt(pc(r.b, r.t)), 'n')}${td(fmt(r.tb), 'n')}${td(fmt(pc(r.tb, r.t)), 'n')}${td(fmt(r.t), 'n')}</tr>`);
  const t = d.tot;
  h += `<tr class=tot><td colspan=2 class=c>Total</td>${td(fmt(t.b), 'n')}${td(fmt(pc(t.b, t.t)), 'n')}${td(fmt(t.tb), 'n')}${td(fmt(pc(t.tb, t.t)), 'n')}${td(fmt(t.t), 'n')}</tr></table></div>${cvTag(ex)}</div>`;
  return h;
};
const chart7 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: {
    labels: d.rows.map(r => r.k), datasets: [
      { label: 'Berubah', data: d.rows.map(r => r.b), backgroundColor: '#ffc000' },
      { label: 'Tidak Berubah', data: d.rows.map(r => r.tb), backgroundColor: NAVY }]
  },
  options: { ...baseOpt('PERUBAHAN PENGGUNAAN TANAH PER KECAMATAN TAHUN 2014-2026 (HA)', a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel7(wb, ws, d, ex) {
  H(ws, 1, 1, 3, 1, 'No'); H(ws, 1, 2, 3, 2, 'Kecamatan'); H(ws, 1, 3, 1, 6, 'Perubahan Penggunaan Tanah');
  H(ws, 2, 3, 2, 4, 'Berubah'); H(ws, 2, 5, 2, 6, 'Tidak Berubah'); H(ws, 1, 7, 3, 7, 'Jumlah (Ha)');
  ['Luas (Ha)', 'Luas (%)', 'Luas (Ha)', 'Luas (%)'].forEach((t, i) => H(ws, 3, 3 + i, 3, 3 + i, t));
  d.rows.forEach((r, i) => {
    const n = 4 + i;
    cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, r.k);
    cell(ws, n, 3, r.b, { num: 1 }); cell(ws, n, 4, pc(r.b, r.t), { num: 1 });
    cell(ws, n, 5, r.tb, { num: 1 }); cell(ws, n, 6, pc(r.tb, r.t), { num: 1 }); cell(ws, n, 7, r.t, { num: 1, b: 1 });
  });
  const n = 4 + d.rows.length, t = d.tot;
  cell(ws, n, 1, 'Total', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  [t.b, pc(t.b, t.t), t.tb, pc(t.tb, t.t), t.t].forEach((v, i) => cell(ws, n, 3 + i, v, { num: 1, b: 1 }));
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 24; for (let c = 3; c <= 7; c++) ws.getColumn(c).width = 14;
  addChart(wb, ws, ex, d, 8, 0);
}

// =====================================================================
// SHEET 8 & 10 - Perbandingan 2 kolom (lama vs baru) + selisih
// =====================================================================
function buildCmp(fa, fb) {
  const a = sumBy(fa), b = sumBy(fb); if (!a || !b) return null;
  const cats = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort((x, y) => x.localeCompare(y));
  const rows = cats.map(c => ({ c, a: a[c] || 0, b: b[c] || 0, d: (b[c] || 0) - (a[c] || 0) }));
  const tot = rows.reduce((s, r) => ({ a: s.a + r.a, b: s.b + r.b, d: s.d + r.d }), { a: 0, b: 0, d: 0 });
  return { rows, tot };
}
const cmpTable = (d, title, h3) => {
  let h = `<table><tr>${th('No')}${th(title)}${th('Tahun 2014 (Ha)')}${th('Tahun 2026 (Ha)')}${th(h3)}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.c)}${td(fmt(r.a), 'n')}${td(fmt(r.b), 'n')}${td(fd(r.d), 'n')}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${td(fmt(d.tot.a), 'n')}${td(fmt(d.tot.b), 'n')}${td(fd(d.tot.d), 'n')}</tr></table>`;
};
function cmpExcel(ws, d, title, h3) {
  [['No', 1], [title, 2], ['Tahun 2014 (Ha)', 3], ['Tahun 2026 (Ha)', 4], [h3, 5]].forEach(([v, c]) => H(ws, 1, c, 1, c, v));
  d.rows.forEach((r, i) => {
    const n = 2 + i; cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, r.c);
    cell(ws, n, 3, r.a, { num: 1 }); cell(ws, n, 4, r.b, { num: 1 }); cell(ws, n, 5, Math.abs(r.d) < 0.005 ? 0 : r.d, { num: 1 });
  });
  const n = 2 + d.rows.length;
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  cell(ws, n, 3, d.tot.a, { num: 1, b: 1 }); cell(ws, n, 4, d.tot.b, { num: 1, b: 1 }); cell(ws, n, 5, Math.abs(d.tot.d) < 0.005 ? 0 : d.tot.d, { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 44; for (let c = 3; c <= 5; c++) ws.getColumn(c).width = 20;
  return n;
}
// --- Sheet 8 ---
const build8 = () => buildCmp('GNAME25', 'QNAME25');
const html8 = d => cmpTable(d, 'Penggunaan Tanah', 'Perubahan Penggunaan Tanah 2014-2026 (Ha)');
const excel8 = (wb, ws, d) => { cmpExcel(ws, d, 'Penggunaan Tanah', 'Perubahan Penggunaan Tanah 2014-2026 (Ha)'); };

// =====================================================================
// SHEET 9 - Perubahan terluas (GQNAME)
// =====================================================================
function build9() {
  if (!ROWS.length || !COLS.includes('GQNAME')) return null;
  const m = {};
  for (const r of ROWS) {
    const c = r.GQNAME; if (!c) continue;
    const p = c.split(' menjadi ');
    if (p.length === 2 && p[0].trim() === p[1].trim()) continue;      // "A menjadi A" = tidak berubah, dilewati
    m[c] = (m[c] || 0) + area(r);
  }
  const top = Object.entries(m).map(([c, t]) => ({ c, t })).sort((a, b) => b.t - a.t).slice(0, 5);
  if (!top.length) return null;
  return { top, T: top.reduce((s, x) => s + x.t, 0) };
}
const html9 = (d, ex) => `<div class=wrap><div><table><tr>${th('No')}${th('Perubahan Penggunaan Tanah 2014-2026 Terluas')}${th('Luas (Ha)')}</tr>`
  + d.top.map((x, i) => `<tr>${td(i + 1, 'c')}${td(x.c)}${td(fmt(x.t), 'n')}</tr>`).join('')
  + `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${td(fmt(d.T), 'n')}</tr></table></div>${cvTag(ex)}</div>`;
const chart9 = (cv, d, a) => mkChart(cv, { type: 'pie', ct: 'PERUBAHAN PENGGUNAAN TANAH DOMINAN TAHUN 2014-2026 (%)' }, { body: d.top }, a);
function excel9(wb, ws, d, ex) {
  H(ws, 1, 1, 1, 1, 'No'); H(ws, 1, 2, 1, 2, 'Perubahan Penggunaan Tanah 2014-2026 Terluas'); H(ws, 1, 3, 1, 3, 'Luas (Ha)');
  d.top.forEach((x, i) => { cell(ws, 2 + i, 1, i + 1, { ctr: 1 }); cell(ws, 2 + i, 2, x.c); cell(ws, 2 + i, 3, x.t, { num: 1 }); });
  const n = 2 + d.top.length;
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2); cell(ws, n, 3, d.T, { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 70; ws.getColumn(3).width = 16;
  addChart(wb, ws, ex, d, 4, 1);
}

// =====================================================================
// SHEET 10 - Reklasifikasi (GREKLAS vs QREKLAS)
// =====================================================================
const build10 = () => buildCmp('GREKLAS', 'QREKLAS');
const html10 = (d, ex) => {
  const side = `<table><tr>${th('Penggunaan Tanah Reklasifikasi')}${d.rows.map(r => th(r.c)).join('')}</tr>
   <tr><th>Perubahan Penggunaan Tanah 2014-2026 (Ha)</th>${d.rows.map(r => td(fd(r.d), 'n')).join('')}</tr></table>`;
  return `<div class=wrap><div>${cmpTable(d, 'Penggunaan Tanah Reklasifikasi', 'Perubahan Penggunaan Tanah 2014-2026 (Ha)')}</div><div class=side>${side}${cvTag(ex)}</div></div>`;
};
const chart10 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: ['Perubahan Penggunaan Tanah 2014-2026 (Ha)'], datasets: d.rows.map((r, i) => ({ label: r.c, data: [r.d], backgroundColor: PAL2[i % PAL2.length] })) },
  options: { ...baseOpt('PERUBAHAN PENGGUNAAN TANAH REKLASIFIKASI TAHUN 2014-2026 (HA)', a, 'right'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel10(wb, ws, d, ex) {
  cmpExcel(ws, d, 'Penggunaan Tanah Reklasifikasi', 'Perubahan Penggunaan Tanah 2014-2026 (Ha)');
  // tabel samping (transpose): kategori jadi kolom
  H(ws, 1, 7, 1, 7, 'Penggunaan Tanah Reklasifikasi'); H(ws, 2, 7, 2, 7, 'Perubahan Penggunaan Tanah 2014-2026 (Ha)');
  d.rows.forEach((r, i) => { H(ws, 1, 8 + i, 1, 8 + i, r.c); cell(ws, 2, 8 + i, Math.abs(r.d) < 0.005 ? 0 : r.d, { num: 1 }); });
  ws.getColumn(7).width = 40; for (let i = 0; i < d.rows.length; i++) ws.getColumn(8 + i).width = 16;
  addChart(wb, ws, ex, d, 6, 3);
}

// =====================================================================
// SHEET 28_Tabel_IV-17 - Potensi Cadangan Karbon
//   Kelas   : STD_C_G (2014) dan STD_C_Q (2026)
//   Nilai   : TON_C_G (2014) dan TON_C_Q (2026)
// =====================================================================
function sumByVal(catCol, valCol) {
  if (!ROWS.length || !COLS.includes(catCol) || !COLS.includes(valCol)) return null;
  const m = {}; let any = false;
  for (const r of ROWS) {
    const c = r[catCol]; if (!c) continue; any = true;
    m[c] = (m[c] || 0) + (parseFloat(r[valCol]) || 0);
  }
  return any ? m : null;
}
const ket = v => Math.abs(v) < 0.005 ? 'Tetap' : v > 0 ? 'Bertambah' : 'Berkurang';

function build17() {
  const a = sumByVal('STD_C_G', 'TON_C_G'), b = sumByVal('STD_C_Q', 'TON_C_Q');
  if (!a || !b) return null;
  const cats = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort((x, y) => x.localeCompare(y));
  const rows = cats.map(c => ({ c, a: a[c] || 0, b: b[c] || 0, d: (b[c] || 0) - (a[c] || 0) }));
  const tot = rows.reduce((s, r) => ({ a: s.a + r.a, b: s.b + r.b, d: s.d + r.d }), { a: 0, b: 0, d: 0 });
  return { rows, tot };
}
const html17 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=2')}${th('Kelas Penutup Lahan', 'rowspan=2')}${th('Potensi Cadangan Karbon (Ton)', 'colspan=3')}${th('Keterangan', 'rowspan=2')}</tr>
  <tr>${th('Tahun 2014')}${th('Tahun 2026')}${th('Perubahan')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.c)}${td(fmt(r.a), 'n')}${td(fmt(r.b), 'n')}${td(fd(r.d), 'n')}${td(ket(r.d), 'c')}</tr>`);
  const t = d.tot;
  h += `<tr class=tot><td colspan=2 class=c>Jumlah (Ton)</td>${td(fmt(t.a), 'n')}${td(fmt(t.b), 'n')}${td(fd(t.d), 'n')}${td(ket(t.d), 'c')}</tr></table></div>`;
  const side = `<table><tr>${th('Kelas Penutup Lahan')}${d.rows.map(r => th(r.c)).join('')}</tr>
   <tr><th>Perubahan</th>${d.rows.map(r => td(fd(r.d), 'n')).join('')}</tr></table>`;
  return h + `<div class=side>${side}${cvTag(ex)}</div></div>`;
};
const chart17 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: ['Perubahan (Ton)'], datasets: d.rows.map((r, i) => ({ label: r.c, data: [r.d], backgroundColor: PAL2[i % PAL2.length] })) },
  options: { ...baseOpt('PERUBAHAN POTENSI CADANGAN KARBON TAHUN 2014-2026 (TON)', a, 'right'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel17(wb, ws, d, ex) {
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Kelas Penutup Lahan'); H(ws, 1, 3, 1, 5, 'Potensi Cadangan Karbon (Ton)');
  H(ws, 2, 3, 2, 3, 'Tahun 2014'); H(ws, 2, 4, 2, 4, 'Tahun 2026'); H(ws, 2, 5, 2, 5, 'Perubahan'); H(ws, 1, 6, 2, 6, 'Keterangan');
  const z = v => Math.abs(v) < 0.005 ? 0 : v;
  d.rows.forEach((r, i) => {
    const n = 3 + i;
    cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, r.c);
    cell(ws, n, 3, r.a, { num: 1 }); cell(ws, n, 4, r.b, { num: 1 }); cell(ws, n, 5, z(r.d), { num: 1 });
    cell(ws, n, 6, ket(r.d), { ctr: 1 });
  });
  const n = 3 + d.rows.length, t = d.tot;
  cell(ws, n, 1, 'Jumlah (Ton)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  cell(ws, n, 3, t.a, { num: 1, b: 1 }); cell(ws, n, 4, t.b, { num: 1, b: 1 }); cell(ws, n, 5, z(t.d), { num: 1, b: 1 });
  cell(ws, n, 6, ket(t.d), { b: 1, ctr: 1 });
  H(ws, 1, 8, 1, 8, 'Kelas Penutup Lahan'); H(ws, 2, 8, 2, 8, 'Perubahan');
  d.rows.forEach((r, i) => { H(ws, 1, 9 + i, 1, 9 + i, r.c); cell(ws, 2, 9 + i, z(r.d), { num: 1 }); });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 32; for (let c = 3; c <= 5; c++) ws.getColumn(c).width = 16; ws.getColumn(6).width = 14;
  ws.getColumn(8).width = 24; for (let i = 0; i < d.rows.length; i++) ws.getColumn(9 + i).width = 18;
  addChart(wb, ws, ex, d, 7, 3);
}

// =====================================================================
// SHEET 25_Tabel_IV-14, 26_Tabel_IV-15, 27_Tabel_IV-16 - Ketersediaan Tanah
//   Sumber : kolom V_ARAHAN x kecamatan (WADMKC), luas dari LUASHA
// =====================================================================
// ---------- PENGATURAN (edit di sini) ----------
// IV-15: 2 jenis arahan Tanaman Pangan. 'key' = kata yang HARUS ada di teks V_ARAHAN (huruf kecil, tidak peka besar/kecil).
const ARAHAN_PANGAN = [
  { label: 'Tersedia untuk Tanaman Pangan dalam rangka optimalisasi penggunaan tanah', key: ['tanaman pangan', 'optimalisasi'] },
  { label: 'Tersedia untuk Tanaman Pangan sesuai tata ruang', key: ['tanaman pangan', 'tata ruang'] }
];
// IV-16: kelompok potensi sektor lain. Semua V_ARAHAN yang mengandung 'key' dimasukkan ke kelompok itu (1 baris per nilai V_ARAHAN).
// const SEKTOR_LAIN = [
//   { name: 'Potensi Perumahan', key: 'perumahan' },
//   { name: 'Potensi Perkebunan', key: 'perkebunan' },
//   { name: 'Potensi Hortikultura', key: 'hortikultura' },
//   { name: 'Potensi Peternakan', key: 'peternakan' },
//   { name: 'Potensi Pariwisata', key: 'pariwisata' },
//   { name: 'Potensi Perdagangan dan Jasa', key: 'perdagangan' },
//   { name: 'Potensi Perkantoran', key: 'perkantoran' },
//   { name: 'Potensi Industri', key: 'industri' },
//   { name: 'Potensi Pertambangan', key: 'pertambangan' },
//   { name: 'Potensi Fasilitas Umum dan Sosial', key: 'fasilitas umum' },
//   { name: 'Potensi Infrastruktur Perkotaan', key: 'infrastruktur' },
//   { name: 'Potensi Transportasi', key: 'transportasi' },
//   { name: 'Potensi Pertahanan dan Keamanan', key: 'pertahanan' }
// ];
const SEKTOR_LAIN = [
  { name: 'Potensi Pengembangan Sektor Perkebunan', key: 'perkebunan' },
  { name: 'Potensi Pengembangan Sektor Perumahan', key: 'perumahan' },
  { name: 'Potensi Pengembangan Sektor Industri', key: 'industri' },
  { name: 'Potensi Pengembangan Sektor Pertambangan', key: 'pertambangan' }
];
const COL_BAR = ['#ffc000', '#1f3864', '#70ad47', '#5b9bd5', '#ed7d31', '#a5a5a5', '#7030a0', '#00b0a0', '#c55a11', '#2e75b6', '#bf9000', '#548235', '#843c0c', '#44546a'];
const PIE_COL = ['#70ad47', '#5b9bd5', '#ffc000', '#ed7d31', '#7030a0', '#00b0a0', '#2e75b6', '#bf9000', '#843c0c', '#44546a', '#a5a5a5', '#1f3864', '#c55a11', '#548235', '#ff9da7'];
const SISA_COL = '#c00000';                                                           // warna "Sisa Luas"

// ---------- Helper ----------
const norm = s => String(s || '').toLowerCase().split(' ').filter(Boolean).join(' ');
const hasAll = (a, keys) => keys.every(k => norm(a).includes(k));
const sumArr = v => v.reduce((a, b) => a + b, 0);
// luas per (V_ARAHAN x kecamatan) -> {kecs, m:{arahan:{kec:luas}}} atau null
function arahanMatrix() {
  if (!ROWS.length || !COLS.includes('V_ARAHAN')) return null;
  const kecs = [...new Set(ROWS.map(kecOf))].sort(), m = {};
  for (const r of ROWS) {
    const a = (r.V_ARAHAN || '').trim(); if (!a) continue;
    const k = kecOf(r); (m[a] ??= {})[k] = (m[a][k] || 0) + area(r);
  }
  return Object.keys(m).length ? { kecs, m } : null;
}
const sumKec = (mx, names) => mx.kecs.map(k => names.reduce((s, a) => s + (mx.m[a][k] || 0), 0));
// nama wilayah (kota/kabupaten) diambil dari kolom WADMKK; jika ada >1 nilai, digabung dengan koma
function wilayah() {
  if (!ROWS.length || !COLS.includes('WADMKK')) return 'Wilayah';
  const m = {}; ROWS.forEach(r => { const w = (r.WADMKK || '').trim(); if (w) m[w] = (m[w] || 0) + 1; });
  const k = Object.keys(m).sort((a, b) => m[b] - m[a]);
  return k.length ? k.join(', ') : 'Wilayah';
}

const tdr = v => `<td class="n"><b>${fd(v)}</b></td>`;

// ---------- IV-15 : Potensi Pertanian Pangan ----------
function build15() {
  const mx = arahanMatrix(); if (!mx) return null;
  const all = Object.keys(mx.m); let found = false;
  const rows = ARAHAN_PANGAN.map(g => {
    const names = all.filter(a => hasAll(a, g.key)); if (names.length) found = true;
    const v = sumKec(mx, names); return { label: g.label, v, t: sumArr(v) };
  });
  if (!found) return null;
  const kt = mx.kecs.map((_, i) => rows.reduce((s, r) => s + r.v[i], 0));
  return { kecs: mx.kecs, rows, kt, T: sumArr(kt), wil: wilayah() };
}
const html15 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=2')}${th('Arahan Ketersediaan', 'rowspan=2')}${th('Kecamatan (Ha)', 'colspan=' + d.kecs.length)}${th(d.wil + ' (Ha)', 'rowspan=2')}</tr>
  <tr>${d.kecs.map(k => th(k)).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.label)}${r.v.map(v => td(fd(v), 'n')).join('')}${tdr(r.t)}</tr>`);
  h += `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${d.kt.map(v => td(fd(v), 'n')).join('')}${td(fd(d.T), 'n')}</tr></table></div><div class=side>${cvTag(ex)}</div></div>`;
  return h;
};
const chart15 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: d.kecs, datasets: d.rows.map((r, i) => ({ label: r.label, data: r.v, backgroundColor: COL_BAR[i % COL_BAR.length] })) },
  options: { ...baseOpt('KETERSEDIAAN TANAH UNTUK POTENSI PERTANIAN PANGAN TAHUN 2026 (HA)', a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel15(wb, ws, d, ex) {
  const nk = d.kecs.length, cT = 3 + nk;
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Arahan Ketersediaan'); H(ws, 1, 3, 1, 2 + nk, 'Kecamatan (Ha)');
  d.kecs.forEach((k, i) => H(ws, 2, 3 + i, 2, 3 + i, k)); H(ws, 1, cT, 2, cT, d.wil + ' (Ha)');
  d.rows.forEach((r, i) => {
    const n = 3 + i;
    cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, r.label).alignment = { vertical: 'middle', wrapText: true };
    r.v.forEach((v, j) => cell(ws, n, 3 + j, v, { num: 1 })); cell(ws, n, cT, r.t, { num: 1, b: 1 });
  });
  const n = 3 + d.rows.length;
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  d.kt.forEach((v, j) => cell(ws, n, 3 + j, v, { num: 1, b: 1 })); cell(ws, n, cT, d.T, { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 55; for (let c = 3; c <= cT; c++) ws.getColumn(c).width = 15;
  addChart(wb, ws, ex, d, cT + 1, 0);
}

// ---------- IV-16 : Potensi Sektor Lain ----------
function build16() {
  const mx = arahanMatrix(); if (!mx) return null;
  const all = Object.keys(mx.m), used = new Set(), groups = [];
  all.forEach(a => { if (ARAHAN_PANGAN.some(g => hasAll(a, g.key))) used.add(a); });   // milik IV-15, jangan dobel
  SEKTOR_LAIN.forEach(s => {
    const names = all.filter(a => !used.has(a) && norm(a).includes(s.key)).sort((x, y) => x.localeCompare(y));
    if (!names.length) return;
    names.forEach(a => used.add(a));
    const items = names.map(a => { const v = sumKec(mx, [a]); return { label: a, v, t: sumArr(v) }; });
    const sub = mx.kecs.map((_, i) => items.reduce((x, r) => x + r.v[i], 0));
    groups.push({ name: s.name, items, sub, t: sumArr(sub) });
  });
  if (!groups.length) return null;
  const kt = mx.kecs.map((_, i) => groups.reduce((x, g) => x + g.sub[i], 0));
  return { kecs: mx.kecs, groups, kt, T: sumArr(kt), wil: wilayah(), unused: all.filter(a => !used.has(a) && !norm(a).startsWith('tidak tersedia')).sort() };
}
const html16 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=2')}${th('Potensi Sektor Lain', 'rowspan=2')}${th('Arahan Ketersediaan', 'rowspan=2')}${th('Kecamatan (Ha)', 'colspan=' + d.kecs.length)}${th(d.wil + ' (Ha)', 'rowspan=2')}</tr>
  <tr>${d.kecs.map(k => th(k)).join('')}</tr>`;
  d.groups.forEach((g, gi) => {
    g.items.forEach((it, i) => h += `<tr>${i === 0 ? `<td class=c rowspan=${g.items.length + 1}>${gi + 1}</td><td class=c rowspan=${g.items.length + 1}>${g.name}</td>` : ''}${td(it.label)}${it.v.map(v => td(fd(v), 'n')).join('')}${tdr(it.t)}</tr>`);
    h += `<tr class=tot>${td(g.name + ' (Ha)')}${g.sub.map(v => td(fd(v), 'n')).join('')}${td(fd(g.t), 'n')}</tr>`;
  });
  h += `<tr class=tot><td colspan=3 class=c>Jumlah (Ha)</td>${d.kt.map(v => td(fd(v), 'n')).join('')}${td(fd(d.T), 'n')}</tr></table>`;
  if (d.unused.length) h += `<div class=note>Nilai V_ARAHAN yang tidak masuk IV-15 maupun IV-16: ${d.unused.join('; ')}</div>`;
  return h + `</div><div class=side>${cvTag(ex)}</div></div>`;
};
const chart16 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: d.kecs, datasets: d.groups.map((g, i) => ({ label: g.name, data: g.sub, backgroundColor: COL_BAR[i % COL_BAR.length] })) },
  options: { ...baseOpt('KETERSEDIAAN TANAH UNTUK POTENSI SEKTOR LAIN TAHUN 2026 (HA)', a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel16(wb, ws, d, ex) {
  const nk = d.kecs.length, cT = 4 + nk;
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Potensi Sektor Lain'); H(ws, 1, 3, 2, 3, 'Arahan Ketersediaan');
  H(ws, 1, 4, 1, 3 + nk, 'Kecamatan (Ha)'); d.kecs.forEach((k, i) => H(ws, 2, 4 + i, 2, 4 + i, k)); H(ws, 1, cT, 2, cT, d.wil + ' (Ha)');
  const mid = { horizontal: 'center', vertical: 'middle', wrapText: true }, wrap = { vertical: 'middle', wrapText: true };
  let r = 3;
  d.groups.forEach((g, gi) => {
    const r0 = r, rs = r0 + g.items.length;                       // rs = baris subtotal
    for (let x = r0; x <= rs; x++) { cell(ws, x, 1, null); cell(ws, x, 2, null); }
    ws.getCell(r0, 1).value = gi + 1; ws.getCell(r0, 1).alignment = mid;
    ws.getCell(r0, 2).value = g.name; ws.getCell(r0, 2).alignment = mid;
    ws.mergeCells(r0, 1, rs, 1); ws.mergeCells(r0, 2, rs, 2);
    g.items.forEach((it, i) => {
      const n = r0 + i;
      cell(ws, n, 3, it.label).alignment = wrap;
      it.v.forEach((v, j) => cell(ws, n, 4 + j, v, { num: 1 })); cell(ws, n, cT, it.t, { num: 1, b: 1 });
    });
    cell(ws, rs, 3, g.name + ' (Ha)', { b: 1 });
    g.sub.forEach((v, j) => cell(ws, rs, 4 + j, v, { num: 1, b: 1 })); cell(ws, rs, cT, g.t, { num: 1, b: 1 });
    r = rs + 1;
  });
  cell(ws, r, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, r, 2, null); cell(ws, r, 3, null); ws.mergeCells(r, 1, r, 3);
  d.kt.forEach((v, j) => cell(ws, r, 4 + j, v, { num: 1, b: 1 })); cell(ws, r, cT, d.T, { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 26; ws.getColumn(3).width = 60; for (let c = 4; c <= cT; c++) ws.getColumn(c).width = 15;
  addChart(wb, ws, ex, d, cT + 1, 0);
}

// ---------- IV-14 : Ringkasan Potensi Sektoral ----------
function build14() {
  const p = build15(), s = build16(); if (!p && !s) return null;
  const total = ROWS.reduce((a, r) => a + area(r), 0), wil = wilayah();   // luas seluruh wilayah (semua baris)           // luas seluruh wilayah (semua baris)
  const items = [{ c: 'Potensi Pertanian Pangan', t: p ? p.T : 0 }];
  if (s) s.groups.forEach(g => items.push({ c: g.name, t: g.t }));
  items.push({ c: 'Sisa Luas ' + wil, t: total - sumArr(items.map(x => x.t)) });
  items.forEach(x => x.p = pc(x.t, total));
  return { items, total, wil };
}
const html14 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No')}${th('Potensi Sektoral')}${th('Ha')}${th('% Wilayah')}</tr>`;
  d.items.forEach((x, i) => h += `<tr>${td(i + 1, 'c')}${td(x.c)}${td(fd(x.t), 'n')}${td(fd(x.p), 'n')}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${td(fd(d.total), 'n')}${td('100.00', 'n')}</tr></table></div><div class=side>${cvTag(ex)}</div></div>`;
};
const pieVals = {
  id: 'pv', afterDatasetsDraw(ch) {
    const c = ch.ctx, ds = ch.data.datasets[0], tot = sumArr(ds.data);
    c.save(); c.font = 'bold 12px Arial'; c.fillStyle = '#fff'; c.textAlign = 'center';
    ch.getDatasetMeta(0).data.forEach((el, i) => { const p = tot ? ds.data[i] / tot * 100 : 0; if (p < 2) return; const t = el.tooltipPosition(); c.fillText(p.toFixed(2) + '%', t.x, t.y + 4); });
    c.restore();
  }
};
const chart14 = (cv, d, a) => {
  const n = d.items.length;
  return new Chart(cv, {
    type: 'pie',
    data: { labels: d.items.map(x => x.c), datasets: [{ data: d.items.map(x => Math.max(0, x.t)), backgroundColor: d.items.map((_, i) => i === n - 1 ? SISA_COL : PIE_COL[i % PIE_COL.length]) }] },
    options: baseOpt('KETERSEDIAAN TANAH UNTUK POTENSI SEKTORAL TAHUN 2026 (%)', a, 'right'), plugins: [bgPlugin, pieVals]
  });
};
function excel14(wb, ws, d, ex) {
  H(ws, 1, 1, 1, 1, 'No'); H(ws, 1, 2, 1, 2, 'Potensi Sektoral'); H(ws, 1, 3, 1, 3, 'Ha'); H(ws, 1, 4, 1, 4, '% Wilayah');
  d.items.forEach((x, i) => {
    const n = 2 + i;
    cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, x.c); cell(ws, n, 3, x.t, { num: 1 }); cell(ws, n, 4, x.p, { num: 1 });
  });
  const n = 2 + d.items.length;
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  cell(ws, n, 3, d.total, { num: 1, b: 1 }); cell(ws, n, 4, 100, { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 40; ws.getColumn(3).width = 16; ws.getColumn(4).width = 14;
  addChart(wb, ws, ex, d, 5, 0);
}

// =====================================================================
// SHEET 24_Tabel_IV-12 - Arahan Ketersediaan Tanah per Kecamatan x Pola Ruang RDTR
//   Kecamatan = WADMKC | Pola Ruang = POLA_COL | Arahan = V_ARAHAN | Luas = LUASHA
//   Luas (%) = luas baris / jumlah seluruh baris tabel x 100
// =====================================================================
const POLA_COL = 'NAMOBJ';   // kolom Pola Ruang RDTR (ganti jika nama kolomnya lain)

// angka kecil tampil lebih banyak desimal: 2 desimal; jika hasilnya 0.00 -> 4 desimal; jika masih 0.0000 -> 5, 6, dst
const decA = v => { let n = 2; while (n < 8 && v > 0 && Number(v.toFixed(n)) === 0) n += n === 2 ? 2 : 1; return n; };
const fmtA = v => v.toLocaleString('en-US', { minimumFractionDigits: decA(v), maximumFractionDigits: decA(v) });
const nfA = v => '#,##0.' + '0'.repeat(decA(v));

function build12() {
  if (!ROWS.length || !['V_ARAHAN', POLA_COL].every(c => COLS.includes(c))) return null;
  const m = {}; let total = 0;
  for (const r of ROWS) {
    const p = (r[POLA_COL] || '').trim(), a = (r.V_ARAHAN || '').trim(); if (!p || !a) continue;
    const k = kecOf(r), v = area(r);
    m[k] ??= {}; m[k][p] ??= {}; m[k][p][a] = (m[k][p][a] || 0) + v; total += v;
  }
  if (!Object.keys(m).length) return null;
  const by = (x, y) => x.localeCompare(y);
  const kecs = Object.keys(m).sort().map(k => {
    const polas = Object.keys(m[k]).sort(by).map(p => ({ p, items: Object.keys(m[k][p]).sort(by).map(a => ({ a, v: m[k][p][a] })) }));
    return { k, polas, n: polas.reduce((s, q) => s + q.items.length, 0) };
  });
  return { kecs, total, wil: wilayah() };
}
const html12 = d => {
  const top = 'style="vertical-align:top"';
  let h = `<div class=raw-scroll><table><tr>${th('No')}${th('Kecamatan')}${th('Pola Ruang RDTR')}${th('Arahan Ketersediaan Tanah')}${th('Luas (Ha)')}${th('Luas (%)')}</tr>`;
  d.kecs.forEach((kc, ki) => kc.polas.forEach((q, qi) => q.items.forEach((it, i) => {
    h += '<tr>';
    if (qi === 0 && i === 0) h += `<td class=c rowspan=${kc.n} ${top}>${ki + 1}</td><td rowspan=${kc.n} ${top}>${kc.k}</td>`;
    if (i === 0) h += `<td class=c rowspan=${q.items.length} ${top}>${q.p}</td>`;
    h += `${td(it.a)}${td(fmtA(it.v), 'n')}${td(fmtA(pc(it.v, d.total)), 'n')}</tr>`;
  })));
  return h + `<tr class=tot><td colspan=4 class=c>${d.wil}</td>${td(fd(d.total), 'n')}${td('100.00', 'n')}</tr></table></div>`;
};
function excel12(wb, ws, d) {
  ['No', 'Kecamatan', 'Pola Ruang RDTR', 'Arahan Ketersediaan Tanah', 'Luas (Ha)', 'Luas (%)'].forEach((t, i) => H(ws, 1, i + 1, 1, i + 1, t));
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true }, top = { vertical: 'top', wrapText: true };
  const num = (r, c, v) => { const x = cell(ws, r, c, v, { num: 1 }); x.numFmt = nfA(v); };
  let r = 2;
  d.kecs.forEach((kc, ki) => {
    const k0 = r;
    kc.polas.forEach(q => {
      const p0 = r;
      q.items.forEach(it => {
        for (let c = 1; c <= 3; c++) cell(ws, r, c, null);
        cell(ws, r, 4, it.a); num(r, 5, it.v); num(r, 6, pc(it.v, d.total)); r++;
      });
      ws.getCell(p0, 3).value = q.p; ws.getCell(p0, 3).alignment = mid;
      if (r - 1 > p0) ws.mergeCells(p0, 3, r - 1, 3);
    });
    ws.getCell(k0, 1).value = ki + 1; ws.getCell(k0, 1).alignment = mid;
    ws.getCell(k0, 2).value = kc.k; ws.getCell(k0, 2).alignment = top;
    if (r - 1 > k0) { ws.mergeCells(k0, 1, r - 1, 1); ws.mergeCells(k0, 2, r - 1, 2); }
  });
  cell(ws, r, 1, d.wil, { b: 1, ctr: 1 }); for (let c = 2; c <= 4; c++) cell(ws, r, c, null); ws.mergeCells(r, 1, r, 4);
  cell(ws, r, 5, d.total, { num: 1, b: 1 }); cell(ws, r, 6, 100, { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 22; ws.getColumn(3).width = 36; ws.getColumn(4).width = 75; ws.getColumn(5).width = 14; ws.getColumn(6).width = 12;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
}

// =====================================================================
// SHEET 23_Tabel_IV-11 - Ketersediaan Tanah (Tersedia / Tidak Tersedia) per Kecamatan x Pola Ruang RDTR
//   Kecamatan = WADMKC | Pola Ruang = POLA_COL | Tersedia/Tidak Tersedia = VNAME | Luas = LUASHA
//   Luas (%) tiap baris = luas / Jumlah (Ha) baris itu x 100  (baris total: terhadap seluruh wilayah)
//   Butuh: POLA_COL dan norm() dari blok sebelumnya (IV-12 dan IV-14/15/16).
// =====================================================================
function build11() {
  const VCOL = COLS.find(c => c.toUpperCase() === 'VNAME');   // cocokkan tanpa peduli huruf besar/kecil
  if (!ROWS.length || ![VCOL, POLA_COL].every(c => COLS.includes(c))) return null;
  const m = {}; let any = false;
  for (const r of ROWS) {
    const p = (r[POLA_COL] || '').trim(), s = norm(r[VCOL]); if (!p || !s) continue;
    const i = s.startsWith('tidak') ? 1 : s.includes('tersedia') ? 0 : -1; if (i < 0) continue;   // 0 = Tersedia, 1 = Tidak Tersedia
    any = true; const k = kecOf(r);
    m[k] ??= {}; m[k][p] ??= [0, 0]; m[k][p][i] += area(r);
  }
  if (!any) return null;
  const mk = (a, b) => { const t = a + b; return { a, b, t, pa: pc(a, t), pb: pc(b, t) }; };
  const by = (x, y) => x.localeCompare(y);
  const kecs = Object.keys(m).sort().map(k => ({ k, rows: Object.keys(m[k]).sort(by).map(p => ({ p, ...mk(m[k][p][0], m[k][p][1]) })) }));
  const all = kecs.flatMap(kc => kc.rows);
  return { kecs, tot: mk(sumArr(all.map(r => r.a)), sumArr(all.map(r => r.b))), wil: wilayah() };
}
const html11 = d => {
  const top = 'style="vertical-align:top"';
  let h = `<table><tr>${th('No', 'rowspan=3')}${th('Kecamatan', 'rowspan=3')}${th('Pola Ruang RDTR', 'rowspan=3')}${th('Ketersediaan Tanah', 'colspan=4')}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${th('Tersedia', 'colspan=2')}${th('Tidak Tersedia', 'colspan=2')}</tr>
  <tr>${th('Luas (Ha)')}${th('Luas (%)')}${th('Luas (Ha)')}${th('Luas (%)')}</tr>`;
  const cells = r => `${td(fd(r.a), 'n')}${td(fd(r.pa), 'n')}${td(fd(r.b), 'n')}${td(fd(r.pb), 'n')}${tdr(r.t)}`;
  d.kecs.forEach((kc, ki) => kc.rows.forEach((r, i) => {
    h += '<tr>' + (i === 0 ? `<td class=c rowspan=${kc.rows.length} ${top}>${ki + 1}</td><td rowspan=${kc.rows.length} ${top}>${kc.k}</td>` : '') + td(r.p) + cells(r) + '</tr>';
  }));
  return h + `<tr class=tot><td colspan=3 class=c>${d.wil}</td>${cells(d.tot)}</tr></table>`;
};
function excel11(wb, ws, d) {
  H(ws, 1, 1, 3, 1, 'No'); H(ws, 1, 2, 3, 2, 'Kecamatan'); H(ws, 1, 3, 3, 3, 'Pola Ruang RDTR'); H(ws, 1, 4, 1, 7, 'Ketersediaan Tanah');
  H(ws, 2, 4, 2, 5, 'Tersedia'); H(ws, 2, 6, 2, 7, 'Tidak Tersedia');
  [4, 5, 6, 7].forEach((c, i) => H(ws, 3, c, 3, c, i % 2 ? 'Luas (%)' : 'Luas (Ha)')); H(ws, 1, 8, 3, 8, 'Jumlah (Ha)');
  const put = (r, x, b) => [x.a, x.pa, x.b, x.pb, x.t].forEach((v, i) => cell(ws, r, 4 + i, v, { num: 1, b: b || i === 4 }));
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true }, left = { vertical: 'top', wrapText: true };
  let r = 4;
  d.kecs.forEach((kc, ki) => {
    const k0 = r;
    kc.rows.forEach(x => { cell(ws, r, 1, null); cell(ws, r, 2, null); cell(ws, r, 3, x.p).alignment = left; put(r, x); r++; });
    ws.getCell(k0, 1).value = ki + 1; ws.getCell(k0, 1).alignment = mid;
    ws.getCell(k0, 2).value = kc.k; ws.getCell(k0, 2).alignment = left;
    if (r - 1 > k0) { ws.mergeCells(k0, 1, r - 1, 1); ws.mergeCells(k0, 2, r - 1, 2); }
  });
  cell(ws, r, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, r, 2, null); cell(ws, r, 3, null); ws.mergeCells(r, 1, r, 3); put(r, d.tot, true);
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 20; ws.getColumn(3).width = 38; for (let c = 4; c <= 7; c++) ws.getColumn(c).width = 12; ws.getColumn(8).width = 15;
  ws.views = [{ state: 'frozen', ySplit: 3 }];
}

// =====================================================================
// SHEET 21_Tabel_IV-9  - Ketersediaan Tanah per Kecamatan (+ grafik)
// SHEET 22_Tabel_IV-10 - Ketersediaan Tanah per Kecamatan x Penggunaan Tanah (QNAME25)
//   Tersedia / Tidak Tersedia diambil dari kolom Vname (huruf besar/kecil bebas)
// =====================================================================
const vCol = () => COLS.find(c => c.toUpperCase() === 'VNAME');
// 0 = Tersedia, 1 = Tidak Tersedia, -1 = tidak dikenali
const ketIdx = v => {
  const s = norm(String(v || '').replace(/[\s_\-]+/g, ' ')); if (!s) return -1;
  return /(^|\s)(tidak|tdk|belum)(\s|$)/.test(s) ? 1 : s.includes('tersedia') ? 0 : -1;
};
const ketMk = (a, b) => { const t = a + b; return { a, b, t, pa: pc(a, t), pb: pc(b, t) }; };

// ---------- IV-9 : per kecamatan ----------
function buildKetKec() {
  const V = vCol(); if (!ROWS.length || !V) return null;
  const m = {}; let any = false;
  for (const r of ROWS) {
    const i = ketIdx(r[V]); if (i < 0) continue;
    any = true; const k = kecOf(r); m[k] ??= [0, 0]; m[k][i] += area(r);
  }
  if (!any) return null;
  const rows = Object.keys(m).sort().map(k => ({ k, ...ketMk(m[k][0], m[k][1]) }));
  return { rows, tot: ketMk(sumArr(rows.map(r => r.a)), sumArr(rows.map(r => r.b))), wil: wilayah() };
}
const htmlKetKec = (d, ex) => {
  const cells = r => `${td(fd(r.a), 'n')}${td(fd(r.pa), 'n')}${td(fd(r.b), 'n')}${td(fd(r.pb), 'n')}${tdr(r.t)}`;
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=3')}${th('Kecamatan', 'rowspan=3')}${th('Ketersediaan Tanah', 'colspan=4')}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${th('Tersedia', 'colspan=2')}${th('Tidak Tersedia', 'colspan=2')}</tr>
  <tr>${th('Luas (Ha)')}${th('Luas (%)')}${th('Luas (Ha)')}${th('Luas (%)')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.k)}${cells(r)}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>${d.wil}</td>${cells(d.tot)}</tr></table></div>${cvTag(ex)}</div>`;
};
const chartKetKec = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: {
    labels: d.rows.map(r => r.k), datasets: [
      { label: 'Tersedia', data: d.rows.map(r => r.a), backgroundColor: '#ffc000' },
      { label: 'Tidak Tersedia', data: d.rows.map(r => r.b), backgroundColor: NAVY }]
  },
  options: { ...baseOpt(['KETERSEDIAAN TANAH PER KECAMATAN', d.wil.toUpperCase() + ' TAHUN 2026 (HA)'], a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excelKetKec(wb, ws, d, ex) {
  H(ws, 1, 1, 3, 1, 'No'); H(ws, 1, 2, 3, 2, 'Kecamatan'); H(ws, 1, 3, 1, 6, 'Ketersediaan Tanah');
  H(ws, 2, 3, 2, 4, 'Tersedia'); H(ws, 2, 5, 2, 6, 'Tidak Tersedia'); H(ws, 1, 7, 3, 7, 'Jumlah (Ha)');
  ['Luas (Ha)', 'Luas (%)', 'Luas (Ha)', 'Luas (%)'].forEach((t, i) => H(ws, 3, 3 + i, 3, 3 + i, t));
  const put = (n, x, b) => [x.a, x.pa, x.b, x.pb, x.t].forEach((v, i) => cell(ws, n, 3 + i, v, { num: 1, b: b || i === 4 }));
  d.rows.forEach((r, i) => { const n = 4 + i; cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, r.k); put(n, r); });
  const n = 4 + d.rows.length;
  cell(ws, n, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2); put(n, d.tot, true);
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 24; for (let c = 3; c <= 7; c++) ws.getColumn(c).width = 14;
  addChart(wb, ws, ex, d, 8, 0);
}

// ---------- IV-10 : per kecamatan x Penggunaan Tanah (kolom col, mis. QNAME25) ----------
function buildKetPola(col) {
  const V = vCol(); if (!ROWS.length || !V || !COLS.includes(col)) return null;
  const m = {}; let any = false;
  for (const r of ROWS) {
    const p = (r[col] || '').trim(), i = ketIdx(r[V]); if (!p || i < 0) continue;
    any = true; const k = kecOf(r); m[k] ??= {}; m[k][p] ??= [0, 0]; m[k][p][i] += area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  const kecs = Object.keys(m).sort().map(k => ({ k, rows: Object.keys(m[k]).sort(by).map(p => ({ p, ...ketMk(m[k][p][0], m[k][p][1]) })) }));
  const all = kecs.flatMap(kc => kc.rows);
  return { kecs, tot: ketMk(sumArr(all.map(r => r.a)), sumArr(all.map(r => r.b))), wil: wilayah() };
}
const htmlKetPola = label => d => {
  const top = 'style="vertical-align:top"';
  const cells = r => `${td(fd(r.a), 'n')}${td(fd(r.pa), 'n')}${td(fd(r.b), 'n')}${td(fd(r.pb), 'n')}${tdr(r.t)}`;
  let h = `<div class=raw-scroll><table><tr>${th('No', 'rowspan=3')}${th('Kecamatan', 'rowspan=3')}${th(label, 'rowspan=3')}${th('Ketersediaan Tanah', 'colspan=4')}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${th('Tersedia', 'colspan=2')}${th('Tidak Tersedia', 'colspan=2')}</tr>
  <tr>${th('Luas (Ha)')}${th('Luas (%)')}${th('Luas (Ha)')}${th('Luas (%)')}</tr>`;
  d.kecs.forEach((kc, ki) => kc.rows.forEach((r, i) => {
    h += '<tr>' + (i === 0 ? `<td class=c rowspan=${kc.rows.length} ${top}>${ki + 1}</td><td rowspan=${kc.rows.length} ${top}>${kc.k}</td>` : '') + td(r.p) + cells(r) + '</tr>';
  }));
  return h + `<tr class=tot><td colspan=3 class=c>${d.wil}</td>${cells(d.tot)}</tr></table></div>`;
};
const excelKetPola = label => (wb, ws, d) => {
  H(ws, 1, 1, 3, 1, 'No'); H(ws, 1, 2, 3, 2, 'Kecamatan'); H(ws, 1, 3, 3, 3, label); H(ws, 1, 4, 1, 7, 'Ketersediaan Tanah');
  H(ws, 2, 4, 2, 5, 'Tersedia'); H(ws, 2, 6, 2, 7, 'Tidak Tersedia');
  [4, 5, 6, 7].forEach((c, i) => H(ws, 3, c, 3, c, i % 2 ? 'Luas (%)' : 'Luas (Ha)')); H(ws, 1, 8, 3, 8, 'Jumlah (Ha)');
  const put = (r, x, b) => [x.a, x.pa, x.b, x.pb, x.t].forEach((v, i) => cell(ws, r, 4 + i, v, { num: 1, b: b || i === 4 }));
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true }, left = { vertical: 'top', wrapText: true };
  let r = 4;
  d.kecs.forEach((kc, ki) => {
    const k0 = r;
    kc.rows.forEach(x => { cell(ws, r, 1, null); cell(ws, r, 2, null); cell(ws, r, 3, x.p).alignment = left; put(r, x); r++; });
    ws.getCell(k0, 1).value = ki + 1; ws.getCell(k0, 1).alignment = mid;
    ws.getCell(k0, 2).value = kc.k; ws.getCell(k0, 2).alignment = left;
    if (r - 1 > k0) { ws.mergeCells(k0, 1, r - 1, 1); ws.mergeCells(k0, 2, r - 1, 2); }
  });
  cell(ws, r, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, r, 2, null); cell(ws, r, 3, null); ws.mergeCells(r, 1, r, 3); put(r, d.tot, true);
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 20; ws.getColumn(3).width = 38; for (let c = 4; c <= 7; c++) ws.getColumn(c).width = 12; ws.getColumn(8).width = 15;
  ws.views = [{ state: 'frozen', ySplit: 3 }];
};

// =====================================================================
// SHEET 20_Tabel_IV-8 - Kesesuaian Penggunaan Tanah terhadap RDTR per Kecamatan x Pola Ruang
// SHEET 19_Tabel_IV-7 - Matriks Penggunaan Tanah (QNAME25) x Arahan Fungsi Kawasan (NAMOBJ) = S / T / M
//   Kesesuaian diambil dari kolom KSPOLA (atau KS_POLA): Sesuai (S), Tidak Sesuai (T), Mendukung (M)
// =====================================================================
const MENDUKUNG_SESUAI = true;   // true: "Mendukung" ikut dihitung ke kolom "Sesuai" di IV-8; false: "Mendukung" tidak dihitung
const ksCol = () => COLS.find(c => ['KSPOLA', 'KS_POLA'].includes(c.toUpperCase()));
const ksCode = v => {            // -> 'S' | 'T' | 'M' | '' (tidak dikenali)
  const s = norm(String(v || '').replace(/[\s_\-]+/g, ' ')); if (!s) return '';
  if (/(^|\s)(tidak|tdk)(\s|$)/.test(s)) return 'T';
  if (s.includes('mendukung')) return 'M';
  return s.includes('sesuai') ? 'S' : '';
};
// const ksMk = (a, b) => { const t = a + b; return { a, b, t, pa: pc(a, t), pb: pc(b, t) }; };

// ---------- IV-8 ----------
// 3 kategori: Sesuai (S), Mendukung (M), Tidak Sesuai (T). x = Tidak Sesuai, t = jumlah.
const ksMk = (s, m, x) => { const t = s + m + x; return { s, m, x, t, ps: pc(s, t), pm: pc(m, t), px: pc(x, t) }; };
function buildKes() {
  const K = ksCol(); if (!ROWS.length || !K || !COLS.includes(POLA_COL)) return null;
  const g = {}; let any = false;
  for (const r of ROWS) {
    const p = (r[POLA_COL] || '').trim(), c = ksCode(r[K]); if (!p || !c) continue;
    const i = c === 'S' ? 0 : c === 'M' ? 1 : 2;
    any = true; const k = kecOf(r); g[k] ??= {}; g[k][p] ??= [0, 0, 0]; g[k][p][i] += area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  const kecs = Object.keys(g).sort().map(k => {
    const rows = Object.keys(g[k]).sort(by).map(p => ({ p, ...ksMk(...g[k][p]) }));
    return { k, rows, sub: ksMk(sumArr(rows.map(r => r.s)), sumArr(rows.map(r => r.m)), sumArr(rows.map(r => r.x))) };
  });
  const tot = ksMk(sumArr(kecs.map(c => c.sub.s)), sumArr(kecs.map(c => c.sub.m)), sumArr(kecs.map(c => c.sub.x)));
  return { kecs, tot, wil: wilayah() };
}
const ksCells = r => `${td(fd(r.s), 'n')}${td(fd(r.ps), 'n')}${td(fd(r.m), 'n')}${td(fd(r.pm), 'n')}${td(fd(r.x), 'n')}${td(fd(r.px), 'n')}${tdr(r.t)}`;
const ksHead = (kolom, rs) => `<tr>${th('No', 'rowspan=3')}${th('Kecamatan', 'rowspan=3')}${kolom ? th(kolom, 'rowspan=3') : ''}${th('Kesesuaian Penggunaan Tanah terhadap RTRW', 'colspan=6')}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${th('Sesuai', 'colspan=2')}${th('Mendukung', 'colspan=2')}${th('Tidak Sesuai', 'colspan=2')}</tr>
  <tr>${[1, 2, 3].map(() => th('Luas (Ha)') + th('Luas (%)')).join('')}</tr>`;
const htmlKes = (d, ex) => {
  const top = 'style="vertical-align:top"';
  // ringkasan per kecamatan + grafik (di atas), lalu tabel rinci
  let h = `<div class=wrap><div><table>${ksHead('')}`;
  d.kecs.forEach((kc, i) => h += `<tr>${td(i + 1, 'c')}${td(kc.k)}${ksCells(kc.sub)}</tr>`);
  h += `<tr class=tot><td colspan=2 class=c>${d.wil}</td>${ksCells(d.tot)}</tr></table></div>${cvTag(ex)}</div><br>`;
  h += `<div class=raw-scroll><table>${ksHead('Pola Ruang RTRW')}`;
  d.kecs.forEach((kc, ki) => {
    const span = kc.rows.length + 1;
    kc.rows.forEach((r, i) => {
      h += '<tr>' + (i === 0 ? `<td class=c rowspan=${span} ${top}>${ki + 1}</td><td rowspan=${span} ${top}>${kc.k}</td>` : '') + td(r.p) + ksCells(r) + '</tr>';
    });
    h += `<tr class=tot><td class=c>Jumlah</td>${ksCells(kc.sub)}</tr>`;
  });
  return h + `<tr class=tot><td colspan=3 class=c>${d.wil}</td>${ksCells(d.tot)}</tr></table></div>`;
};
const chartKes = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: {
    labels: d.kecs.map(k => k.k), datasets: [
      { label: 'Sesuai', data: d.kecs.map(k => k.sub.s), backgroundColor: NAVY },
      { label: 'Mendukung', data: d.kecs.map(k => k.sub.m), backgroundColor: '#ffc000' },
      { label: 'Tidak Sesuai', data: d.kecs.map(k => k.sub.x), backgroundColor: '#c00000' }]
  },
  options: { ...baseOpt(['KESESUAIAN PENGGUNAAN TANAH TERHADAP FUNGSI KAWASAN', 'PER KECAMATAN DI ' + d.wil.toUpperCase() + ' TAHUN 2026 (HA)'], a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excelKes(wb, ws, d, ex) {
  const heads = (c0, kolom) => {      // header 3 baris mulai kolom c0; kolom = true -> ada kolom Pola Ruang
    const o = kolom ? 3 : 2;          // jumlah kolom sebelum data angka
    H(ws, 1, c0, 3, c0, 'No'); H(ws, 1, c0 + 1, 3, c0 + 1, 'Kecamatan'); if (kolom) H(ws, 1, c0 + 2, 3, c0 + 2, 'Pola Ruang RTRW');
    H(ws, 1, c0 + o, 1, c0 + o + 5, 'Kesesuaian Penggunaan Tanah terhadap RTRW');
    ['Sesuai', 'Mendukung', 'Tidak Sesuai'].forEach((t, i) => H(ws, 2, c0 + o + i * 2, 2, c0 + o + i * 2 + 1, t));
    for (let i = 0; i < 6; i++) H(ws, 3, c0 + o + i, 3, c0 + o + i, i % 2 ? 'Luas (%)' : 'Luas (Ha)');
    H(ws, 1, c0 + o + 6, 3, c0 + o + 6, 'Jumlah (Ha)');
    return c0 + o;                    // kolom angka pertama
  };
  const put = (r, c, x, b) => [x.s, x.ps, x.m, x.pm, x.x, x.px, x.t].forEach((v, i) => cell(ws, r, c + i, v, { num: 1, b: b || i === 6 }));
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true }, left = { vertical: 'top', wrapText: true };
  // ---- tabel rinci (kiri): kolom 1..10 ----
  const n1 = heads(1, true);
  let r = 4;
  d.kecs.forEach((kc, ki) => {
    const k0 = r;
    kc.rows.forEach(x => { cell(ws, r, 1, null); cell(ws, r, 2, null); cell(ws, r, 3, x.p).alignment = left; put(r, n1, x); r++; });
    cell(ws, r, 1, null); cell(ws, r, 2, null); cell(ws, r, 3, 'Jumlah', { b: 1, ctr: 1 }); put(r, n1, kc.sub, true); r++;
    ws.getCell(k0, 1).value = ki + 1; ws.getCell(k0, 1).alignment = mid;
    ws.getCell(k0, 2).value = kc.k; ws.getCell(k0, 2).alignment = left;
    ws.mergeCells(k0, 1, r - 1, 1); ws.mergeCells(k0, 2, r - 1, 2);
  });
  cell(ws, r, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, r, 2, null); cell(ws, r, 3, null); ws.mergeCells(r, 1, r, 3); put(r, n1, d.tot, true);
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 20; ws.getColumn(3).width = 38; for (let c = 4; c <= 9; c++) ws.getColumn(c).width = 12; ws.getColumn(10).width = 15;
  // ---- ringkasan per kecamatan (kanan): mulai kolom 12 ----
  const c0 = 12, n2 = heads(c0, false);
  d.kecs.forEach((kc, i) => { const n = 4 + i; cell(ws, n, c0, i + 1, { ctr: 1 }); cell(ws, n, c0 + 1, kc.k); put(n, n2, kc.sub); });
  const n = 4 + d.kecs.length;
  cell(ws, n, c0, d.wil, { b: 1, ctr: 1 }); cell(ws, n, c0 + 1, null); ws.mergeCells(n, c0, n, c0 + 1); put(n, n2, d.tot, true);
  ws.getColumn(c0).width = 6; ws.getColumn(c0 + 1).width = 20; for (let c = c0 + 2; c <= c0 + 7; c++) ws.getColumn(c).width = 12; ws.getColumn(c0 + 8).width = 15;
  addChart(wb, ws, ex, d, c0 - 1, n + 1);   // grafik di bawah tabel ringkasan
  ws.views = [{ state: 'frozen', ySplit: 3 }];
}

// ---------- IV-7 : matriks S / T / M ----------
function buildMatKs() {
  const K = ksCol(); if (!ROWS.length || !K || !COLS.includes('QNAME25') || !COLS.includes(POLA_COL)) return null;
  const m = {}, cols = new Set(); let any = false;
  for (const r of ROWS) {
    const q = (r.QNAME25 || '').trim(), p = (r[POLA_COL] || '').trim(), c = ksCode(r[K]); if (!q || !p || !c) continue;
    any = true; cols.add(p); (m[q] ??= {})[p] ??= new Set(); m[q][p].add(c);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  return { cols: [...cols].sort(by), rows: Object.keys(m).sort(by).map(q => ({ q, v: m[q] })) };
}
const matCell = (v, p) => v[p] ? [...v[p]].join('/') : '-';   // jika satu pasangan punya >1 nilai, tampil mis. "S/T"
const htmlMatKs = d => {
  const vert = 'style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:normal;height:170px;min-width:30px;padding:6px 2px;text-align:left"';
  let h = `<div class=raw-scroll><table><tr>${th('No', 'rowspan=2')}${th('Penggunaan Tanah', 'rowspan=2')}${th('Arahan Fungsi Kawasan pada Rencana Detail Tata Ruang', `colspan=${d.cols.length}`)}</tr>
  <tr>${d.cols.map(p => th(p, vert)).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.q)}${d.cols.map(p => td(matCell(r.v, p), 'c')).join('')}</tr>`);
  return h + '</table></div><div class=note>S = Sesuai, T = Tidak Sesuai, M = Mendukung, - = tidak ada data</div>';
};
function excelMatKs(wb, ws, d) {
  const nc = d.cols.length;
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Penggunaan Tanah'); H(ws, 1, 3, 1, 2 + nc, 'Arahan Fungsi Kawasan pada Rencana Detail Tata Ruang');
  d.cols.forEach((p, j) => { H(ws, 2, 3 + j, 2, 3 + j, p); ws.getCell(2, 3 + j).alignment = { textRotation: 90, horizontal: 'center', vertical: 'middle', wrapText: true }; });
  ws.getRow(2).height = 150;
  d.rows.forEach((r, i) => {
    const n = 3 + i; cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, r.q);
    d.cols.forEach((p, j) => cell(ws, n, 3 + j, matCell(r.v, p), { ctr: 1 }));
  });
  const n = 4 + d.rows.length; ws.getCell(n, 2).value = 'S = Sesuai, T = Tidak Sesuai, M = Mendukung, - = tidak ada data';
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 42; for (let c = 3; c <= 2 + nc; c++) ws.getColumn(c).width = 6;
  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
}

// =====================================================================
// SHEET 16_Tabel_IV-4 - Perubahan Penggunaan Tanah (GQNAME) per Arahan Fungsi Kawasan dalam RTRW (NAMOBJ)
//   Luas (%) = luas baris / luas total arahan fungsi kawasan itu x 100
// SHEET 14_Tabel_IV-3 - Matriks Penggunaan Tanah Lama (GNAME25) x Penggunaan Tanah Baru (QNAME25), satuan Ha
//   Butuh: POLA_COL, fmtA, nfA (dari blok IV-12) dan fd, td, th, tdr, cell, H, hdr.
// =====================================================================
// ---------- IV-4 ----------
function buildGQ() {
  if (!ROWS.length || !['GQNAME', POLA_COL].every(c => COLS.includes(c))) return null;
  const m = {}; let any = false;
  for (const r of ROWS) {
    const p = (r[POLA_COL] || '').trim(), g = (r.GQNAME || '').trim(); if (!p || !g) continue;
    any = true; (m[p] ??= {})[g] = (m[p][g] || 0) + area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  return {
    groups: Object.keys(m).sort(by).map(p => {
      const items = Object.keys(m[p]).sort(by).map(g => ({ g, v: m[p][g] }));
      return { p, items, t: sumArr(items.map(i => i.v)) };
    })
  };
}
const htmlGQ = d => {
  const top = 'style="vertical-align:top"';
  let h = `<div class=raw-scroll><table><tr>${th('No')}${th('Arahan Fungsi Kawasan dalam RTRW')}${th('Luas (Ha)')}${th('Perubahan Penggunaan Tanah')}${th('Luas (Ha)')}${th('Luas (%)')}</tr>`;
  d.groups.forEach((gr, gi) => gr.items.forEach((it, i) => {
    h += '<tr>' + (i === 0 ? `<td class=c rowspan=${gr.items.length} ${top}>${gi + 1}</td><td class=c rowspan=${gr.items.length} ${top}>${gr.p}</td><td class=n rowspan=${gr.items.length} ${top}>${fmtA(gr.t)}</td>` : '') +
      `${td(it.g)}${td(fmtA(it.v), 'n')}${td(fmtA(pc(it.v, gr.t)), 'n')}</tr>`;
  }));
  return h + '</table></div>';
};
function excelGQ(wb, ws, d) {
  ['No', 'Arahan Fungsi Kawasan dalam RTRW', 'Luas (Ha)', 'Perubahan Penggunaan Tanah', 'Luas (Ha)', 'Luas (%)'].forEach((t, i) => H(ws, 1, i + 1, 1, i + 1, t));
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true };
  const num = (r, c, v) => { const x = cell(ws, r, c, v, { num: 1 }); x.numFmt = nfA(v); };
  let r = 2;
  d.groups.forEach((gr, gi) => {
    const g0 = r;
    gr.items.forEach(it => { for (let c = 1; c <= 3; c++) cell(ws, r, c, null); cell(ws, r, 4, it.g); num(r, 5, it.v); num(r, 6, pc(it.v, gr.t)); r++; });
    ws.getCell(g0, 1).value = gi + 1; ws.getCell(g0, 2).value = gr.p; ws.getCell(g0, 3).value = gr.t; ws.getCell(g0, 3).numFmt = nfA(gr.t);
    for (let c = 1; c <= 3; c++) ws.getCell(g0, c).alignment = mid;
    if (r - 1 > g0) for (let c = 1; c <= 3; c++) ws.mergeCells(g0, c, r - 1, c);
  });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 34; ws.getColumn(3).width = 14; ws.getColumn(4).width = 70; ws.getColumn(5).width = 14; ws.getColumn(6).width = 12;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
}

// ---------- IV-3 ----------
function buildMatLB() {
  if (!ROWS.length || !['GNAME25', 'QNAME25'].every(c => COLS.includes(c))) return null;
  const m = {}, cs = new Set(); let any = false;
  for (const r of ROWS) {
    const g = (r.GNAME25 || '').trim(), q = (r.QNAME25 || '').trim(); if (!g || !q) continue;
    any = true; cs.add(q); (m[g] ??= {})[q] = (m[g][q] || 0) + area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y), cols = [...cs].sort(by);
  const rows = Object.keys(m).sort(by).map(g => ({ g, v: m[g], t: sumArr(Object.values(m[g])) }));
  const kt = cols.map(q => sumArr(rows.map(r => r.v[q] || 0)));
  return { cols, rows, kt, T: sumArr(kt) };
}
const htmlMatLB = d => {
  const vert = 'style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:normal;height:170px;min-width:30px;padding:6px 2px;text-align:left"';
  let h = `<div style="overflow-x:auto"><table><tr>${th('No', 'rowspan=2')}${th('Penggunaan Tanah Lama (Tahun 2014)', 'rowspan=2')}${th('Penggunaan Tanah Baru (Tahun 2026)', `colspan=${d.cols.length}`)}${th('Jumlah (Ha)', 'rowspan=2')}</tr>
  <tr>${d.cols.map(q => th(q, vert)).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.g)}${d.cols.map(q => td(fd(r.v[q] || 0), 'n')).join('')}${tdr(r.t)}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${d.kt.map(v => tdr(v)).join('')}${tdr(d.T)}</tr></table></div>`;
};
function excelMatLB(wb, ws, d) {
  const nc = d.cols.length;
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Penggunaan Tanah Lama (Tahun 2014)'); H(ws, 1, 3, 1, 2 + nc, 'Penggunaan Tanah Baru (Tahun 2026)'); H(ws, 1, 3 + nc, 2, 3 + nc, 'Jumlah (Ha)');
  d.cols.forEach((q, j) => { H(ws, 2, 3 + j, 2, 3 + j, q); ws.getCell(2, 3 + j).alignment = { textRotation: 90, horizontal: 'center', vertical: 'middle', wrapText: true }; });
  ws.getRow(2).height = 150;
  d.rows.forEach((r, i) => {
    const n = 3 + i; cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, r.g);
    d.cols.forEach((q, j) => cell(ws, n, 3 + j, r.v[q] || 0, { num: 1 })); cell(ws, n, 3 + nc, r.t, { num: 1, b: 1 });
  });
  const n = 3 + d.rows.length;
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  d.kt.forEach((v, j) => cell(ws, n, 3 + j, v, { num: 1, b: 1 })); cell(ws, n, 3 + nc, d.T, { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 36; for (let c = 3; c <= 2 + nc; c++) ws.getColumn(c).width = 10; ws.getColumn(3 + nc).width = 13;
  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
}

// =====================================================================
// DAFTAR SHEET TAMBAHAN (urutan = urutan sheet 7, 8, 9, 10)
// =====================================================================
const EXTRA = [
  { tab: '12_Tabel_IV-1', build: build7, html: html7, chart: chart7, excel: excel7, w: 640, h: 340 },
  { tab: '13_Tabel_IV-2', build: build8, html: html8, chart: null, excel: excel8 },
  { tab: '15_Grafik_IV-1', build: build9, html: html9, chart: chart9, excel: excel9, w: 640, h: 340 },
  { tab: '17_Tabel_IV-5', build: build10, html: html10, chart: chart10, excel: excel10, w: 640, h: 340 },
  { tab: '28_Tabel_IV-17', build: build17, html: html17, chart: chart17, excel: excel17, w: 760, h: 380 },
  { tab: '25_Tabel_IV-14', build: build14, html: html14, chart: chart14, excel: excel14, w: 760, h: 340 },
  { tab: '26_Tabel_IV-15', build: build15, html: html15, chart: chart15, excel: excel15, w: 760, h: 400 },
  { tab: '27_Tabel_IV-16', build: build16, html: html16, chart: chart16, excel: excel16, w: 760, h: 400 },
  { tab: '24_Tabel_IV-12', build: build12, html: html12, chart: null, excel: excel12 },
  { tab: '23_Tabel_IV-11', build: build11, html: html11, chart: null, excel: excel11 },
  { tab: '22_Tabel_IV-10', build: () => buildKetPola('QNAME25'), html: htmlKetPola('Penggunaan Tanah'), chart: null, excel: excelKetPola('Penggunaan Tanah') },
  { tab: '21_Tabel_IV-9', build: buildKetKec, html: htmlKetKec, chart: chartKetKec, excel: excelKetKec, w: 640, h: 340 },
  { tab: '20_Tabel_IV-8', build: buildKes, html: htmlKes, chart: chartKes, excel: excelKes, w: 640, h: 340 },
  { tab: '19_Tabel_IV-7', build: buildMatKs, html: htmlMatKs, chart: null, excel: excelMatKs },
  { tab: '16_Tabel_IV-4', build: buildGQ, html: htmlGQ, chart: null, excel: excelGQ },
  { tab: '14_Tabel_IV-3', build: buildMatLB, html: htmlMatLB, chart: null, excel: excelMatLB }
];
