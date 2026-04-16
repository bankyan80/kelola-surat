// ============================================================
// GOOGLE APPS SCRIPT - Backend untuk Administrasi Surat
// ============================================================
// CARA PASANG:
// 1. Buka Google Sheets baru
// 2. Klik Extensions > Apps Script
// 3. Hapus semua kode, lalu paste kode ini
// 4. Klik Deploy > New Deployment
// 5. Pilih type: Web app
// 6. Execute as: Me
// 7. Who has access: Anyone
// 8. Klik Deploy, lalu copy URL-nya
// 9. Masukkan URL tersebut ke pengaturan di aplikasi
// ============================================================

// Fungsi utama untuk menerima data via HTTP POST (Web App)
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action; // 'saveMasuk', 'saveKeluar', 'deleteMasuk', 'deleteKeluar', 'syncAll'

    if (action === 'syncAll') {
      // Sinkronisasi semua data sekaligus
      var result = {};
      if (data.suratMasuk) {
        result.masuk = syncSheet(data.suratMasuk, 'Surat Masuk', getHeadersMasuk());
      }
      if (data.suratKeluar) {
        result.keluar = syncSheet(data.suratKeluar, 'Surat Keluar', getHeadersKeluar());
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: 'Sinkronisasi berhasil',
        data: result
      })).setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'saveMasuk') {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: appendRow('Surat Masuk', getHeadersMasuk(), data.item)
      })).setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'saveKeluar') {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: appendRow('Surat Keluar', getHeadersKeluar(), data.item)
      })).setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'deleteMasuk' || action === 'deleteKeluar') {
      var sheetName = action === 'deleteMasuk' ? 'Surat Masuk' : 'Surat Keluar';
      var result = deleteRowByNoUrut(sheetName, data.noUrut);
      return ContentService.createTextOutput(JSON.stringify({
        success: result
      })).setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'ping') {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: 'Koneksi ke Google Sheets berhasil'
      })).setMimeType(ContentService.MimeType.JSON);

    } else {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Action tidak dikenali: ' + action
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi untuk menerima GET request (buka di browser untuk test)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Google Sheets API aktif. Kirim POST request untuk menyimpan data.',
    version: '1.0'
  })).setMimeType(ContentService.MimeType.JSON);
}

// Header untuk Surat Masuk
function getHeadersMasuk() {
  return ['No', 'No Surat', 'Tanggal Surat', 'Pengirim', 'Perihal', 'Kode Klasifikasi', 'Sifat', 'No Urut', 'Tanggal Diterima', 'Keterangan'];
}

// Header untuk Surat Keluar
function getHeadersKeluar() {
  return ['No', 'No Surat', 'Tanggal Surat', 'Tujuan', 'Perihal', 'Kode Klasifikasi', 'Sifat', 'Jenis Surat', 'No Urut', 'Tanggal Dikirim', 'Keterangan'];
}

// Sinkronisasi seluruh data ke sheet (replace all)
function syncSheet(items, sheetName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  // Buat sheet jika belum ada
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1b3a5c');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
  }

  // Hapus semua data lama (kecuali header)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  // Pastikan header ada
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  var existingHeaders = headerRange.getValues()[0];
  if (existingHeaders[0] !== headers[0]) {
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1b3a5c');
    headerRange.setFontColor('#ffffff');
  }

  // Tulis data baru
  if (items && items.length > 0) {
    var rows = items.map(function(item, index) {
      if (sheetName === 'Surat Masuk') {
        return [
          index + 1,
          item.noSurat || '',
          item.tglSurat || '',
          item.pengirim || '',
          item.perihal || '',
          item.kodeKlasifikasi || '',
          item.sifat || '',
          item.noUrut || '',
          item.tglDiterima || item.tglSurat || '',
          item.keterangan || ''
        ];
      } else {
        return [
          index + 1,
          item.noSurat || '',
          item.tglSurat || '',
          item.tujuan || '',
          item.perihal || '',
          item.kodeKlasifikasi || '',
          item.sifat || '',
          item.jenisSurat || '',
          item.noUrut || '',
          item.tglDikirim || item.tglSurat || '',
          item.keterangan || ''
        ];
      }
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Auto-resize kolom
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Freeze header
  sheet.setFrozenRows(1);

  return 'Berhasil: ' + (items ? items.length : 0) + ' data ditulis ke ' + sheetName;
}

// Tambah satu baris data
function appendRow(sheetName, headers, item) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  // Buat sheet jika belum ada
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1b3a5c');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  var lastRow = sheet.getLastRow();
  var no = lastRow; // nomor urut otomatis

  var row;
  if (sheetName === 'Surat Masuk') {
    row = [
      no,
      item.noSurat || '',
      item.tglSurat || '',
      item.pengirim || '',
      item.perihal || '',
      item.kodeKlasifikasi || '',
      item.sifat || '',
      item.noUrut || '',
      item.tglDiterima || item.tglSurat || '',
      item.keterangan || ''
    ];
  } else {
    row = [
      no,
      item.noSurat || '',
      item.tglSurat || '',
      item.tujuan || '',
      item.perihal || '',
      item.kodeKlasifikasi || '',
      item.sifat || '',
      item.jenisSurat || '',
      item.noUrut || '',
      item.tglDikirim || item.tglSurat || '',
      item.keterangan || ''
    ];
  }

  sheet.appendRow(row);

  // Auto-resize
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  return 'Data berhasil ditambahkan ke ' + sheetName;
}

// Hapus baris berdasarkan noUrut
function deleteRowByNoUrut(sheetName, noUrut) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // Cari dari bawah ke atas agar index tidak bergeser
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][7] == noUrut) { // kolom ke-8 (index 7) = No Urut
      sheet.deleteRow(i + 2); // +2 karena index mulai dari 0 dan header di baris 1
      return true;
    }
  }
  return false;
}

// Fungsi utilitas: setup sheet otomatis (jalankan manual dari Apps Script editor)
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hapus Sheet default "Sheet1" jika ada
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // Buat sheet Surat Masuk
  var masukHeaders = getHeadersMasuk();
  var masukSheet = ss.getSheetByName('Surat Masuk');
  if (!masukSheet) {
    masukSheet = ss.insertSheet('Surat Masuk');
  }
  masukSheet.getRange(1, 1, 1, masukHeaders.length).setValues([masukHeaders]);
  masukSheet.getRange(1, 1, 1, masukHeaders.length).setFontWeight('bold');
  masukSheet.getRange(1, 1, 1, masukHeaders.length).setBackground('#1b3a5c');
  masukSheet.getRange(1, 1, 1, masukHeaders.length).setFontColor('#ffffff');
  masukSheet.setFrozenRows(1);
  for (var i = 1; i <= masukHeaders.length; i++) {
    masukSheet.autoResizeColumn(i);
  }

  // Buat sheet Surat Keluar
  var keluarHeaders = getHeadersKeluar();
  var keluarSheet = ss.getSheetByName('Surat Keluar');
  if (!keluarSheet) {
    keluarSheet = ss.insertSheet('Surat Keluar');
  }
  keluarSheet.getRange(1, 1, 1, keluarHeaders.length).setValues([keluarHeaders]);
  keluarSheet.getRange(1, 1, 1, keluarHeaders.length).setFontWeight('bold');
  keluarSheet.getRange(1, 1, 1, keluarHeaders.length).setBackground('#1b3a5c');
  keluarSheet.getRange(1, 1, 1, keluarHeaders.length).setFontColor('#ffffff');
  keluarSheet.setFrozenRows(1);
  for (var j = 1; j <= keluarHeaders.length; j++) {
    keluarSheet.autoResizeColumn(j);
  }

  SpreadsheetApp.getUi().alert('Setup selesai! Sheet Surat Masuk dan Surat Keluar berhasil dibuat.');
}
