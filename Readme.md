Panduan Menghubungkan Team Martabak POS ke Google Sheets

Untuk menjadikan Google Sheets sebagai database, ikuti langkah-langkah mudah berikut:

Langkah 1: Buat Spreadsheet Baru

Buka Google Sheets dan buat Spreadsheet Kosong baru.

Beri nama file spreadsheet Anda, misalnya: Database Martabak POS.

Langkah 2: Memasukkan Kode Apps Script

Pada menu atas Google Sheets, klik Ekstensi > Apps Script.

Tab baru akan terbuka. Hapus semua kode yang ada di editor.

Salin dan tempel kode Ultimate Anti-Error berikut ke dalam editor:

// ==========================================
// 1. JALANKAN FUNGSI INI PERTAMA KALI
// ==========================================
function setupAuth() {
  // Fungsi ini bertugas untuk meminta izin akses ke Google Akun Anda 
  // dan membuat struktur sheet dasar secara otomatis.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Tidak terhubung ke spreadsheet.");
  
  getOrCreateSheet(ss, 'Products', ['id', 'name', 'cost', 'price']);
  getOrCreateSheet(ss, 'Transactions', ['id', 'date', 'total', 'paymentMethod', 'kasir', 'items']);
  getOrCreateSheet(ss, 'Expenses', ['id', 'date', 'description', 'amount']);
  
  Logger.log("Setup Berhasil! Database siap digunakan.");
}

// ==========================================
// 2. FUNGSI UNTUK MEMBACA DATA (GET)
// ==========================================
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet tidak ditemukan.");

    var data = {
      products: getSheetData(ss, 'Products'),
      transactions: getSheetData(ss, 'Transactions'),
      expenses: getSheetData(ss, 'Expenses')
    };

    // Parse item transaksi dari JSON String ke Object dengan aman
    data.transactions = data.transactions.map(function(t) {
      if(t.items && typeof t.items === 'string') {
        try { t.items = JSON.parse(t.items); } catch(err) {}
      }
      return t;
    });

    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. FUNGSI UNTUK MENYIMPAN DATA (POST)
// ==========================================
function doPost(e) {
  // Proteksi 1: Mencegah error jika dijalankan langsung (Run) dari editor Apps Script
  if (!e || !e.postData) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Akses ditolak. Endpoint ini hanya menerima request POST dari aplikasi POS.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    // Proteksi 2: Parsing payload secara aman
    var payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseError) {
      throw new Error("Format data tidak valid. Harap kirimkan JSON.");
    }

    var action = payload.action;
    var data = payload.data || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet;

    // Routing Logic berdasarkan Action
    if (action === 'save_transaction') {
      sheet = getOrCreateSheet(ss, 'Transactions', ['id', 'date', 'total', 'paymentMethod', 'kasir', 'items']);
      sheet.appendRow([data.id, data.date, data.total, data.paymentMethod, data.kasir, JSON.stringify(data.items)]);
    }
    else if (action === 'save_expense') {
      sheet = getOrCreateSheet(ss, 'Expenses', ['id', 'date', 'description', 'amount']);
      sheet.appendRow([data.id, data.date, data.description, data.amount]);
    }
    else if (action === 'delete_expense') {
      sheet = getOrCreateSheet(ss, 'Expenses', ['id', 'date', 'description', 'amount']);
      deleteRowById(sheet, data.id);
    }
    else if (action === 'save_product') {
      sheet = getOrCreateSheet(ss, 'Products', ['id', 'name', 'cost', 'price']);
      var row = findRowById(sheet, data.id);
      if (row > 0) {
        // Update produk yang ada
        sheet.getRange(row, 2, 1, 3).setValues([[data.name, data.cost, data.price]]);
      } else {
        // Tambah produk baru
        sheet.appendRow([data.id, data.name, data.cost, data.price]);
      }
    }
    else if (action === 'delete_product') {
      sheet = getOrCreateSheet(ss, 'Products', ['id', 'name', 'cost', 'price']);
      deleteRowById(sheet, data.id);
    }
    else {
      throw new Error("Action '" + action + "' tidak dikenali.");
    }

    return ContentService.createTextOutput(JSON.stringify({status: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Untuk menangani Pre-flight CORS request secara diam-diam
function doOptions(e) {
  return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
}

// ==========================================
// --- FUNGSI PEMBANTU (HELPER) ---
// ==========================================

function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    // Format header menjadi tebal (bold)
    sheet.getRange("A1:F1").setFontWeight("bold");
  }
  return sheet;
}

function getSheetData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return []; 
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 

  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    var isEmptyRow = true; 
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
      if (data[i][j] !== '') isEmptyRow = false;
    }
    if (!isEmptyRow) result.push(obj);
  }
  return result;
}

function findRowById(sheet, id) {
  if (!sheet) return -1;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function deleteRowById(sheet, id) {
  var row = findRowById(sheet, id);
  if (row > 0) sheet.deleteRow(row);
}


Klik tombol Simpan (ikon disket) di bagian atas.

Langkah 3: Berikan Izin Akses Google (SANGAT PENTING)

Ini adalah langkah krusial agar tidak terjadi error:

Di bagian atas editor Apps Script, ada menu dropdown pilihan fungsi (biasanya tertulis doGet atau setupAuth).

Pilih fungsi setupAuth.

Klik tombol Jalankan (Run).

Google akan memunculkan pop-up "Otorisasi diperlukan" (Authorization required). Klik Tinjau Izin (Review Permissions).

Pilih akun Google Anda.

Akan muncul peringatan "Google belum memverifikasi aplikasi ini". Klik Lanjutan (Advanced) di bagian bawah.

Klik Buka project tak berjudul (tidak aman) / Go to untitled project (unsafe).

Klik Izinkan (Allow).

Jika di log bagian bawah tertulis "Setup Berhasil! Database siap digunakan", berarti database Anda sudah 100% aktif.

Langkah 4: Deploy (Terapkan) Menjadi Web App

Di pojok kanan atas, klik tombol biru Terapkan (Deploy) > Deployment Baru.

Klik ikon roda gigi ⚙️ di sebelah tulisan "Pilih jenis", pilih Aplikasi Web (Web app).

Isi kolom Deskripsi (misal: "Versi 3 Final").

Pada kolom Akses (Who has access), pilih Siapa saja (Anyone). (Sangat penting agar aplikasi bisa sinkronisasi data!)

Klik Terapkan (Deploy).

Akan muncul URL panjang (berakhiran /exec). Salin (Copy) URL tersebut.

Langkah 5: Masukkan URL ke Aplikasi

Buka file index.html aplikasi kasir Anda.

Cari baris const GOOGLE_SHEET_URL = ''; di bagian JavaScript.

Paste (tempel) URL yang Anda salin ke dalam tanda kutip tersebut.
Contoh: const GOOGLE_SHEET_URL = 'https://script.google.com/macros/s/AKfyc.../exec';

Selesai! Sekarang aplikasi kasir Anda sudah terhubung secara mulus dengan Google Sheets.