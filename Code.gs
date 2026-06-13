function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index_5')
    .setTitle('SIBDARA — Distribusi Darah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ID Spreadsheet yang diberikan oleh user
const SPREADSHEET_ID = 'masukuan id sheet kalian';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    ensureHeaders_(sheet);
  }
  return sheet;
}

// Pastikan header ada di sheet
function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Id', 'Jenis Permintaan', 'Tanggal Dropin', 'RS/Klinik Tujuan', 'Komponen', 'Golongan Darah', 'Rhesus', 'Jumlah', 'Jenis Pengimputan', 'Bulan', 'Tahun'
    ]);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#f3f3f3');
  }
}

// Fungsi pembantu untuk mengambil array data (Optimized)
function getSheetDataArray(sheetName) {
  try {
    var sheet = getSheet(sheetName);
    var data = sheet.getDataRange().getValues(); // getValues jauh lebih cepat daripada getDisplayValues
    if (data.length < 2) return [];
    
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rowData = data[i].slice(0, 11);
      if (rowData.join("").trim() === "") continue;
      
      // Konversi Date ke string YYYY-MM-DD agar aman dan sesuai zona waktu
      if (rowData[2] instanceof Date) {
        rowData[2] = Utilities.formatDate(rowData[2], Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      
      rowData[11] = i + 1; // Pastikan index selalu di posisi 11
      rows.push(rowData);

    }
    return rows;
  } catch (e) {
    console.error('Error in getSheetDataArray:', e);
    return [];
  }
}

// 1. Simpan ke DB KK Sari
function saveToKKSari(records) {
  try {
    var sheet = getSheet('DB KK Sari');
    var dataToAppend = [];
    records.forEach(function(rec) {
      dataToAppend.push([
        Utilities.getUuid(),
        rec.jenisPermintaan,
        rec.tanggal,
        rec.rsKlinik,
        rec.komponen,
        rec.golDarah,
        rec.rhesus,
        rec.jumlah,
        rec.jenis,
        rec.bulan,
        rec.tahun
      ]);
    });
    if (dataToAppend.length > 0) {
      var lastRow = getActualLastRow(sheet);
      sheet.getRange(lastRow + 1, 1, dataToAppend.length, 11).setValues(dataToAppend);
    }
    return { success: true, message: 'Data berhasil disimpan ke DB KK Sari' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 2. Ambil data DB KK Sari
function getKKSariData() {
  return { success: true, data: getSheetDataArray('DB KK Sari') };
}

// 3. Ambil data DB Distribusi
function getDistribusiData() {
  return { success: true, data: getSheetDataArray('DB Distribusi') };
}

// 4. Unified Data Function (Baru & Cepat)
function getUnifiedData(distMonths) {
  try {
    var kkData = getSheetDataArray('DB KK Sari');
    var distData = getSheetDataArray('DB Distribusi');
    
    // Generate dashboards from data already in memory (Avoid reading sheet again)
    var dashKK = processDashboardData(kkData, []);
    var dashDist = processDashboardData(distData, distMonths || []);
    
    return {
      success: true,
      kkData: kkData,
      distData: distData, // Kita kirim keduanya agar client selalu sinkron
      dashKK: dashKK,
      dashDist: dashDist
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Helper untuk memproses dashboard di memory (Sangat Cepat)
function processDashboardData(rows, filterMonths) {
  var dashboard = {};
  var DROPING_RS = ["ITD ZA", "MEURAXA", "RSIA", "RS Teuku Umar Calang", "UDD Kota Langsa", "UDD Kab. Tangerang", "UDD Kab. Pidie", "UDD Kab. Aceh Utara", "UDD. Kota Medan"];
  
  var monthFilterMap = {};
  if (filterMonths && filterMonths.length > 0) {
    filterMonths.forEach(function(m) { monthFilterMap[m] = true; });
  }

  rows.forEach(function(row) {
    var rs = row[3];
    var bulan = row[9];
    if (!rs) return;
    
    if (filterMonths && filterMonths.length > 0 && !monthFilterMap[bulan]) return;

    if (!dashboard[rs]) {
      dashboard[rs] = {
        tipe: DROPING_RS.indexOf(rs) !== -1 ? 'droping' : 'non-droping',
        permintaan: 0,
        pemenuhan: 0
      };
    }
    
    var jenis = row[8];
    var jml = parseInt(row[7]) || 0;
    
    if (jenis === 'Permintaan') dashboard[rs].permintaan += jml;
    else if (jenis === 'Pemenuhan') dashboard[rs].pemenuhan += jml;
  });
  
  return dashboard;
}

// Fungsi lama getDashboardData diupdate agar menggunakan logic baru yang lebih cepat
function getDashboardData(sheetName, filterMonths) {
  try {
    var rows = getSheetDataArray(sheetName);
    return { success: true, data: processDashboardData(rows, filterMonths) };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 5. Pindahkan data terpilih ke DB Distribusi
function moveToDistribusi(ids) {
  try {
    var sheetKK = getSheet('DB KK Sari');
    var sheetDist = getSheet('DB Distribusi');
    var rowIndices = ids.map(function(id) { return parseInt(id); }).sort(function(a, b) { return a - b; });
    
    var dataKK = sheetKK.getDataRange().getValues();
    var dataToAppend = [];
    var processedRowIndices = []; // Menyimpan baris yang berhasil divalidasi
    
    rowIndices.forEach(function(rowIdx) {
      var rowData = dataKK[rowIdx - 1]; // Array 0-indexed
      if (rowData) {
        var sliceToK = rowData.slice(0, 11); // Kolom A:K (index 0-10)
        
        // Validasi: pastikan semua kolom A:K terisi
        var isAllFilled = true;
        for (var i = 0; i < 11; i++) {
          if (sliceToK[i] === "" || sliceToK[i] === null || sliceToK[i] === undefined) {
            isAllFilled = false;
            break;
          }
        }
        
        if (isAllFilled) {
          dataToAppend.push(sliceToK);
          processedRowIndices.push(rowIdx);
        }
      }
    });
    
    if (dataToAppend.length > 0) {
      var lastDistRow = getActualLastRow(sheetDist);
      sheetDist.getRange(lastDistRow + 1, 1, dataToAppend.length, 11).setValues(dataToAppend);
      
      // Hapus konten kolom A:I (kolom 1 sampai 9) untuk baris yang berhasil dipindah
      processedRowIndices.forEach(function(rowIdx) {
        sheetKK.getRange(rowIdx, 1, 1, 9).clearContent(); // Hanya hapus datanya saja, baris tetap ada
      });
      
      var skipped = rowIndices.length - processedRowIndices.length;
      var msg = 'Data berhasil dipindahkan ke DB Distribusi.';
      if (skipped > 0) {
        msg += ' (' + skipped + ' baris gagal karena kolom A:K belum lengkap terisi)';
      }
      return { success: true, message: msg };
    } else {
      return { success: false, error: 'Gagal dipindahkan: Pastikan baris yang dipilih sudah terisi lengkap dari kolom A sampai K.' };
    }
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 6. Hapus baris dari DB KK Sari
function deleteFromKKSari(id) {
  try {
    var sheet = getSheet('DB KK Sari');
    sheet.deleteRow(parseInt(id));
    return { success: true, message: 'Data berhasil dihapus' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 7. Update baris di DB KK Sari
function updateKKSariRow(editId, rec) {
  try {
    var sheet = getSheet('DB KK Sari');
    var rowIdx = parseInt(editId);
    var uuid = sheet.getRange(rowIdx, 1).getValue();
    if (!uuid) uuid = Utilities.getUuid();
    var values = [uuid, rec.jenisPermintaan, rec.tanggal, rec.rsKlinik, rec.komponen, rec.golDarah, rec.rhesus, rec.jumlah, rec.jenis, rec.bulan, rec.tahun];
    sheet.getRange(rowIdx, 1, 1, 11).setValues([values]);
    return { success: true, message: 'Data berhasil diperbarui' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Mendapatkan baris terakhir yang benar-benar berisi data (menghindari baris kosong tapi berformat)
 */
function getActualLastRow(sheet) {
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i].join("").trim() !== "") return i + 1;
  }
  return 0;
}

// 8. Hapus baris dari DB Distribusi
function deleteFromDistribusi(id) {
  try {
    var sheet = getSheet('DB Distribusi');
    sheet.deleteRow(parseInt(id));
    return { success: true, message: 'Data berhasil dihapus' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 9. Update baris di DB Distribusi
function updateDistribusiRow(editId, rec) {
  try {
    var sheet = getSheet('DB Distribusi');
    var rowIdx = parseInt(editId);
    var uuid = sheet.getRange(rowIdx, 1).getValue();
    if (!uuid) uuid = Utilities.getUuid();
    var values = [uuid, rec.jenisPermintaan, rec.tanggal, rec.rsKlinik, rec.komponen, rec.golDarah, rec.rhesus, rec.jumlah, rec.jenis, rec.bulan, rec.tahun];
    sheet.getRange(rowIdx, 1, 1, 11).setValues([values]);
    return { success: true, message: 'Data berhasil diperbarui' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 10. Auto Generate ID (onEdit Trigger)
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'DB KK Sari') return;
  
  var colStart = e.range.getColumn();
  var colEnd = colStart + e.range.getNumColumns() - 1;
  var rowStart = e.range.getRow();
  var rowEnd = rowStart + e.range.getNumRows() - 1;
  
  // Jika kolom 'Jenis Permintaan' (Kolom 2) ikut diedit
  if (colStart <= 2 && colEnd >= 2) {
    if (rowStart <= 1) rowStart = 2; // Abaikan baris header
    if (rowStart > rowEnd) return;
    
    var idRange = sheet.getRange(rowStart, 1, rowEnd - rowStart + 1, 1);
    var jpRange = sheet.getRange(rowStart, 2, rowEnd - rowStart + 1, 1);
    
    var idValues = idRange.getValues();
    var jpValues = jpRange.getValues();
    
    var changed = false;
    for (var i = 0; i < jpValues.length; i++) {
      if (jpValues[i][0] !== "" && idValues[i][0] === "") {
        idValues[i][0] = Utilities.getUuid();
        changed = true;
      }
    }
    
    if (changed) {
      idRange.setValues(idValues);
    }
  }
}

// 11. Generate ID untuk data yang sudah ada sebelumnya
function generateMissingIds() {
  var sheet = getSheet('DB KK Sari');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 2); // Ambil kolom Id (1) dan Jenis Permintaan (2)
  var values = dataRange.getValues();
  
  var changed = false;
  for (var i = 0; i < values.length; i++) {
    var idValue = values[i][0];
    var jpValue = values[i][1];
    
    if (jpValue !== "" && idValue === "") {
      values[i][0] = Utilities.getUuid();
      changed = true;
    }
  }
  
  if (changed) {
    // Tulis balik hanya kolom Id
    var idOnlyValues = values.map(function(row) { return [row[0]]; });
    sheet.getRange(2, 1, lastRow - 1, 1).setValues(idOnlyValues);
  }
}

// 12. Tambahkan custom menu agar user bisa mengklik "Generate Missing IDs"
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sibdara Menu')
      .addItem('Generate Missing IDs (DB KK Sari)', 'generateMissingIds')
      .addToUi();
}
