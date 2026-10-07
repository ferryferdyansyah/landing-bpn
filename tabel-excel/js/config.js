// config.js
// Pengaturan utama: daftar sheet (DEFS), warna, variabel data global, dan fungsi format angka.
// Mau ubah judul/tahun/kolom sumber? Cukup edit DEFS di bawah.

const DEFS=[
 {f:'GNAME25',tab:'06_Tabel_III-6',h:'Penggunaan Tanah Lama (Tahun 2014)',t5:'Penggunaan Tanah Lama (Tahun 2014)',ct:'LIMA PENGGUNAAN TANAH TERLUAS (HA)',type:'bar'},
 {f:'QNAME25',tab:'07_Tabel_III-7',h:'Penggunaan Tanah Baru (Tahun 2026)',t5:'Penggunaan Tanah Baru (Tahun 2026)',ct:'LIMA PENGGUNAAN TANAH TERLUAS (HA)',type:'bar'},
 {f:'ONAME25',tab:'08_Tabel_III-8',h:'Gambaran Umum Penguasaan Tanah Tahun 2026',ct:'GAMBARAN UMUM PENGUASAAN TANAH TAHUN 2026 (%)',type:'pie'},
 {f:'NAMOBJ',tab:'10_Tabel_III-10',h:'Rencana Pola Ruang pada RDTR/RTRW',t5:'Rencana Pola Ruang pada RDTR/RTRW Terbesar',ct:'LIMA RENCANA POLA RUANG',type:'bar'},
 {f:'FKWS',tab:'11_Tabel_III-11',h:'Kawasan Hutan',ct:'KAWASAN HUTAN TAHUN 2026 (%)',type:'pie'}
];
const PAL=['#4472c4','#ed7d31','#a5a5a5','#ffc000','#5b9bd5','#70ad47','#264478','#9e480e'];
const NAVY='#1f3864';
let ROWS=[],COLS=[],MODELS=[];
const fmt=n=>n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
