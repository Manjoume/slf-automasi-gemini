// ================================================================
// GOOGLE APPS SCRIPT — AI Agent Inspeksi SLF
// ================================================================
// CARA INSTALL:
// 1. Buka Google Sheets baru (sheets.new)
// 2. Klik Extensions → Apps Script
// 3. Hapus semua kode default → paste seluruh kode ini
// 4. Klik Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Authorize → Copy URL → paste ke aplikasi tab Pengaturan
// ================================================================

var SHEET_TEMUAN    = 'Temuan';
var SHEET_RINGKASAN = 'Ringkasan';
var SHEET_LOG       = 'Log Sync';

// ── Entry point POST ──────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();

    writeTemuan(ss, data);
    writeRingkasan(ss, data);
    writeLog(ss, data);

    return ok({ count: (data.findings || []).length });
  } catch(err) {
    return ok({ error: err.message });
  }
}

// ── Entry point GET (test koneksi) ────────────────────
function doGet(e) {
  return ok({ status: 'AI Agent Inspeksi SLF — aktif' });
}

function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// SHEET TEMUAN — satu sheet per proyek
// ================================================================
function writeTemuan(ss, data) {
  var name = 'Temuan — ' + clean(data.project || 'Proyek');
  var ws   = ss.getSheetByName(name) || ss.insertSheet(name);

  // Tulis header jika sheet baru / kosong
  if (ws.getLastRow() < 6) buildTemuanHeader(ws, data);

  // Hapus data lama mulai baris 7, tulis ulang semua
  var startRow = 7;
  if (ws.getLastRow() >= startRow) {
    ws.deleteRows(startRow, ws.getLastRow() - startRow + 1);
  }

  var findings = data.findings || [];
  if (!findings.length) return;

  // Tulis baris data
  var rows = findings.map(function(f) {
    return [
      f.no         || '',
      f.lokasi     || '',
      f.aspek      || '',
      f.uraian     || '',
      f.persyaratan|| '',
      f.kesesuaian || '',
      f.rekomendasi|| '',
      data.tanggal || '',
      (f.keterangan || '') + (f.urgensi ? ' — ' + f.urgensi : '')
    ];
  });

  ws.getRange(startRow, 1, rows.length, 9).setValues(rows);

  // Styling tiap baris
  rows.forEach(function(row, i) {
    var r   = ws.getRange(startRow + i, 1, 1, 9);
    var ks  = ws.getRange(startRow + i, 6);

    r.setFontSize(10).setVerticalAlignment('top').setWrap(true)
     .setBorder(true, true, true, true, true, true, '#dee2e6',
                SpreadsheetApp.BorderStyle.SOLID);

    r.setBackground(i % 2 === 0 ? '#ffffff' : '#f8f9fc');

    var sym = row[5] || '';
    if      (sym.indexOf('✘') !== -1) { ks.setBackground('#fdeaea').setFontColor('#b02020').setFontWeight('bold'); }
    else if (sym.indexOf('⚠') !== -1) { ks.setBackground('#fff8e1').setFontColor('#7a5a00').setFontWeight('bold'); }
    else if (sym.indexOf('✔') !== -1) { ks.setBackground('#e8f5ee').setFontColor('#1a7a4a').setFontWeight('bold'); }
  });

  // Baris ringkasan di bawah tabel
  var sr    = startRow + rows.length + 1;
  var ok_n  = findings.filter(function(f){ return (f.kesesuaian||'').indexOf('✔') !== -1 && (f.kesesuaian||'').indexOf('⚠') === -1; }).length;
  var ct_n  = findings.filter(function(f){ return (f.kesesuaian||'').indexOf('⚠') !== -1; }).length;
  var no_n  = findings.filter(function(f){ return (f.kesesuaian||'').indexOf('✘') !== -1; }).length;

  ws.getRange(sr, 1, 1, 9).setValues([['RINGKASAN:','','','','✔ Sesuai','⚠ Dgn Catatan','✘ Tidak Sesuai','TOTAL','']])
    .setFontWeight('bold').setBackground('#e8edf5').setFontSize(10);
  ws.getRange(sr+1, 1, 1, 9).setValues([['','','','',ok_n, ct_n, no_n, findings.length,'']]).setFontSize(10);
  ws.getRange(sr+1, 5).setBackground('#e8f5ee').setFontColor('#1a7a4a').setFontWeight('bold');
  ws.getRange(sr+1, 6).setBackground('#fff8e1').setFontColor('#7a5a00').setFontWeight('bold');
  ws.getRange(sr+1, 7).setBackground('#fdeaea').setFontColor('#b02020').setFontWeight('bold');
  ws.getRange(sr+1, 8).setFontWeight('bold');

  SpreadsheetApp.flush();
}

function buildTemuanHeader(ws, data) {
  var proj = data.project  || 'Proyek';
  var type = data.type     || 'SLF';
  var tgl  = data.tanggal  || '';
  var kons = data.konsultan|| '-';

  // Lebar kolom
  [40, 160, 150, 320, 350, 160, 320, 90, 240].forEach(function(w, i) {
    ws.setColumnWidth(i + 1, w);
  });

  // Baris 1 — judul utama
  ws.getRange('A1:I1').merge()
    .setValue('KAJIAN ' + type.toUpperCase() + ' — ' + proj.toUpperCase())
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#0f2744').setFontColor('#ffffff').setVerticalAlignment('middle');
  ws.setRowHeight(1, 36);

  // Baris 2 — subjudul
  ws.getRange('A2:I2').merge()
    .setValue('TABEL TEMUAN PEMERIKSAAN BANGUNAN GEDUNG')
    .setFontSize(11).setHorizontalAlignment('center')
    .setBackground('#1a3a5c').setFontColor('#cce0f0');
  ws.setRowHeight(2, 26);

  // Baris 3 — info
  ws.getRange('A3:I3').merge()
    .setValue('Tanggal: ' + tgl + '     Konsultan: ' + kons +
              '     Di-sync: ' + new Date().toLocaleString('id-ID'))
    .setFontSize(10).setHorizontalAlignment('center')
    .setBackground('#e8edf5').setFontColor('#1a2a3a');
  ws.setRowHeight(3, 22);

  // Baris 4 — legend
  ws.getRange('A4:I4').merge()
    .setValue('Status:   ✔ SESUAI = memenuhi persyaratan   |   ⚠ SESUAI DENGAN CATATAN = perlu perbaikan minor   |   ✘ TIDAK SESUAI = wajib tindak lanjut')
    .setFontSize(9).setHorizontalAlignment('center')
    .setBackground('#f5f7fa').setFontColor('#6b7a8d');
  ws.setRowHeight(4, 20);

  // Baris 5 — spacer
  ws.setRowHeight(5, 8);

  // Baris 6 — header kolom
  var headers = ['No.','Lokasi / Ruangan','Aspek Pemeriksaan','Uraian Temuan',
                 'Persyaratan Teknis','Kesesuaian','Rekomendasi Perbaikan','Tgl','Keterangan / Urgensi'];
  ws.getRange(6, 1, 1, 9).setValues([headers])
    .setFontSize(10).setFontWeight('bold')
    .setBackground('#0f2744').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  ws.setRowHeight(6, 32);

  ws.setFrozenRows(6);
  ws.setFrozenColumns(2);
}

// ================================================================
// SHEET RINGKASAN — dashboard semua proyek
// ================================================================
function writeRingkasan(ss, data) {
  var ws = ss.getSheetByName(SHEET_RINGKASAN) || ss.insertSheet(SHEET_RINGKASAN);
  var f  = data.findings || [];

  var ok_n  = f.filter(function(x){ return (x.kesesuaian||'').indexOf('✔') !== -1 && (x.kesesuaian||'').indexOf('⚠') === -1; }).length;
  var ct_n  = f.filter(function(x){ return (x.kesesuaian||'').indexOf('⚠') !== -1; }).length;
  var no_n  = f.filter(function(x){ return (x.kesesuaian||'').indexOf('✘') !== -1; }).length;
  var sg_n  = f.filter(function(x){ return x.urgensi === 'SEGERA'; }).length;
  var mn_n  = f.filter(function(x){ return x.urgensi === 'JANGKA MENENGAH'; }).length;

  // Header jika kosong
  if (ws.getLastRow() < 2) {
    ws.getRange('A1:I1').merge()
      .setValue('RINGKASAN INSPEKSI — SEMUA PROYEK')
      .setFontSize(13).setFontWeight('bold')
      .setBackground('#0f2744').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setRowHeight(1, 30);

    ws.getRange(2,1,1,9).setValues([['Proyek','Tanggal','Konsultan','Total','✔ Sesuai','⚠ Catatan','✘ Tidak Sesuai','⚡ Segera','Jangka Menengah']])
      .setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setFontSize(10);
    ws.setRowHeight(2, 27);

    [180,100,120,55,65,65,75,65,100].forEach(function(w,i){ ws.setColumnWidth(i+1,w); });
    ws.setFrozenRows(2);
  }

  var proj = data.project || '-';
  var row  = [proj, data.tanggal||'', data.konsultan||'-', f.length, ok_n, ct_n, no_n, sg_n, mn_n];
  var found = false;

  for (var r = 3; r <= ws.getLastRow(); r++) {
    if (ws.getRange(r,1).getValue() === proj) {
      ws.getRange(r,1,1,9).setValues([row]);
      styleRingkasanRow(ws, r, no_n, sg_n);
      found = true; break;
    }
  }
  if (!found) {
    var nr = Math.max(ws.getLastRow()+1, 3);
    ws.getRange(nr,1,1,9).setValues([row]);
    styleRingkasanRow(ws, nr, no_n, sg_n);
  }

  var dataRows = Math.max(ws.getLastRow()-2, 1);
  ws.getRange(3,1,dataRows,9)
    .setBorder(true,true,true,true,true,true,'#dee2e6',SpreadsheetApp.BorderStyle.SOLID)
    .setFontSize(10).setVerticalAlignment('middle');
  ws.getRange(3,1,dataRows,1).setFontWeight('bold');
}

function styleRingkasanRow(ws, r, no_n, sg_n) {
  ws.setRowHeight(r, 25);
  if (no_n > 0) { ws.getRange(r,7).setBackground('#fdeaea').setFontColor('#b02020').setFontWeight('bold'); }
  if (sg_n > 0) { ws.getRange(r,8).setBackground('#fdeaea').setFontColor('#b02020').setFontWeight('bold'); }
}

// ================================================================
// SHEET LOG — riwayat sync
// ================================================================
function writeLog(ss, data) {
  var ws = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);

  if (ws.getLastRow() < 1) {
    ws.getRange(1,1,1,6).setValues([['Waktu Sync','Proyek','Jenis Kajian','Konsultan','Jml Temuan','Status']])
      .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff').setFontSize(10);
    [140,180,160,130,70,70].forEach(function(w,i){ ws.setColumnWidth(i+1,w); });
    ws.setFrozenRows(1);
  }

  var nr = ws.getLastRow()+1;
  ws.getRange(nr,1,1,6).setValues([[
    new Date().toLocaleString('id-ID'),
    data.project  || '-',
    data.type     || '-',
    data.konsultan|| '-',
    (data.findings||[]).length,
    'OK'
  ]]).setFontSize(10)
    .setBorder(true,true,true,true,true,true,'#dee2e6',SpreadsheetApp.BorderStyle.SOLID);

  ws.getRange(nr,6).setBackground('#e8f5ee').setFontColor('#1a7a4a').setFontWeight('bold');
  if (nr%2===0) ws.getRange(nr,1,1,5).setBackground('#f8f9fc');
  ws.setRowHeight(nr, 22);
}

// ── Helper ────────────────────────────────────────────
function clean(s) {
  return (s||'').replace(/[\\\/\?\*\[\]:']/g,'').substring(0,25);
}
