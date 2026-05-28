/**
 * 1. Menampilkan Halaman HTML
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Iuran Kas Warga')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 2. Inisialisasi Database
 * Fungsi ini membuat Spreadsheet baru secara otomatis jika standalone,
 * serta menyiapkan konfigurasi Tanggal_Mulai untuk perhitungan tunggakan otomatis.
 */
function initData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Penanganan jika skrip dijalankan secara standalone (ss bernilai null)
  if (!ss) {
    try {
      ss = SpreadsheetApp.create('Database Iuran Kas Warga');
      Logger.log('Spreadsheet baru berhasil dibuat karena tidak ada spreadsheet aktif yang terdeteksi.');
      Logger.log('URL Spreadsheet baru Anda: ' + ss.getUrl());
    } catch (err) {
      throw new Error("Gagal membuat Spreadsheet baru. Pastikan akun Anda memiliki izin akses Google Drive. Detail: " + err.message);
    }
  }
  
  // Setup Sheet "Users"
  let sheetUsers = ss.getSheetByName('Users');
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(['Username', 'Password', 'Role', 'Nama_Lengkap']);
    // Buat data dummy default
    sheetUsers.appendRow(['admin', 'admin123', 'admin', 'Bapak RT']);
    sheetUsers.appendRow(['warga', 'warga123', 'warga', 'Budi Santoso']);
    // Format Header
    sheetUsers.getRange("A1:D1").setFontWeight("bold").setBackground("#d1fae5");
  }

  // Setup Sheet "Kas"
  let sheetKas = ss.getSheetByName('Kas');
  if (!sheetKas) {
    sheetKas = ss.insertSheet('Kas');
    sheetKas.appendRow(['Tanggal', 'Jenis', 'Nominal', 'Keterangan', 'Nama_Warga']);
    sheetKas.getRange("A1:E1").setFontWeight("bold").setBackground("#d1fae5");
  }
  
  // Setup Sheet "Pengaturan" (Untuk konfigurasi iuran dan tanggal mulai)
  let sheetSet = ss.getSheetByName('Pengaturan');
  if (!sheetSet) {
    sheetSet = ss.insertSheet('Pengaturan');
    sheetSet.appendRow(['Kunci', 'Nilai']);
    sheetSet.appendRow(['Iuran_Bulanan', 50000]); // Default Iuran Rp 50.000 / bulan
    sheetSet.appendRow(['Tanggal_Mulai', '2026-01-01']); // Format: YYYY-MM-DD (Kas dimulai dari Januari 2026)
    sheetSet.getRange("A1:B1").setFontWeight("bold").setBackground("#d1fae5");
  } else {
    // Jika sheet sudah ada tapi belum ada baris Tanggal_Mulai, kita tambahkan
    const dataSet = sheetSet.getDataRange().getValues();
    let adaTanggalMulai = false;
    for(let i = 1; i < dataSet.length; i++) {
      if(dataSet[i][0] === 'Tanggal_Mulai') adaTanggalMulai = true;
    }
    if(!adaTanggalMulai) {
      sheetSet.appendRow(['Tanggal_Mulai', '2026-01-01']);
    }
  }

  // Hapus Sheet default bawaan "Sheet1" / "Sheet 1" jika ada agar rapi
  let defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // Tampilkan informasi sukses
  const msg = "Inisialisasi Database Berhasil!\n\n" +
              "Silakan cek Google Sheets Anda.\n" +
              "URL: " + ss.getUrl();
              
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {
    Logger.log(msg);
  }
}

/**
 * 3. Fungsi Autentikasi / Login
 */
function loginUser(username, password) {
  const ss = getDatabase();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  // Skip header (i = 1)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return {
        success: true,
        role: data[i][2],
        name: data[i][3],
        username: data[i][0]
      };
    }
  }
  return { success: false, message: "Username atau Password salah!" };
}

/**
 * 4. Fungsi Mengambil Data Warga (Dashboard)
 * Menghitung saldo, tunggakan secara otomatis/dinamis berdasarkan waktu berjalan
 */
function getWargaDashboard(username) {
  const ss = getDatabase();
  const sheetKas = ss.getSheetByName('Kas');
  const sheetUsers = ss.getSheetByName('Users');
  const sheetSet = ss.getSheetByName('Pengaturan');
  
  // A. Cari Nama Lengkap Warga
  let namaWarga = "";
  const usersData = sheetUsers.getDataRange().getValues();
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] == username) {
      namaWarga = usersData[i][3];
      break;
    }
  }

  const kasData = sheetKas.getDataRange().getValues();
  
  let totalPemasukan = 0;
  let totalPengeluaran = 0;
  let totalIuranSaya = 0;
  let historyKas = [];

  // B. Iterasi Data Kas untuk Saldo dan Riwayat
  for (let i = 1; i < kasData.length; i++) {
    let tgl = Utilities.formatDate(new Date(kasData[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    let jenis = kasData[i][1];
    let nominal = parseFloat(kasData[i][2]);
    let ket = kasData[i][3];
    let namaUser = kasData[i][4];

    // Kalkulasi Saldo Keseluruhan
    if (jenis == "Pemasukan") totalPemasukan += nominal;
    if (jenis == "Pengeluaran") totalPengeluaran += nominal;

    // Kalkulasi Total Pembayaran Iuran Saya
    if (jenis == "Pemasukan" && namaUser == namaWarga) {
      totalIuranSaya += nominal;
    }

    historyKas.push({
      tgl: tgl,
      jenis: jenis,
      nominal: nominal,
      ket: ket,
      warga: namaUser
    });
  }

  let saldoAkhir = totalPemasukan - totalPengeluaran;
  
  // C. Menghitung Tunggakan secara otomatis dan dinamis
  let iuranBulanan = 50000;
  let tanggalMulaiStr = '2026-01-01'; // Default backup jika tidak terbaca
  
  try {
    const setValues = sheetSet.getDataRange().getValues();
    for (let i = 1; i < setValues.length; i++) {
      if (setValues[i][0] === 'Iuran_Bulanan') {
        iuranBulanan = parseFloat(setValues[i][1]);
      }
      if (setValues[i][0] === 'Tanggal_Mulai') {
        // Jika bertipe Date di Google Sheets, konversi ke string YYYY-MM-DD
        if (setValues[i][1] instanceof Date) {
          tanggalMulaiStr = Utilities.formatDate(setValues[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          tanggalMulaiStr = setValues[i][1].toString();
        }
      }
    }
  } catch(e) {
    Logger.log("Gagal membaca pengaturan: " + e.toString());
  }
  
  // Menghitung selisih bulan dari Tanggal_Mulai sampai Hari Ini
  let tglMulai = new Date(tanggalMulaiStr);
  let tglHariIni = new Date();
  
  // Rumus menghitung jumlah bulan berjalan
  let tahunDiff = tglHariIni.getFullYear() - tglMulai.getFullYear();
  let bulanDiff = tglHariIni.getMonth() - tglMulai.getMonth();
  let jumlahBulanBerjalan = (tahunDiff * 12) + bulanDiff + 1; // +1 untuk menghitung bulan berjalan saat ini
  
  if (jumlahBulanBerjalan < 0) {
    jumlahBulanBerjalan = 0; // Jaga-jaga jika tanggal mulai diset di masa depan
  }
  
  let totalKewajiban = jumlahBulanBerjalan * iuranBulanan; 
  let tunggakan = totalKewajiban - totalIuranSaya;

  // D. 5 Arus Kas Terakhir
  let last5 = historyKas.reverse().slice(0, 5);

  return {
    saldo: saldoAkhir,
    tunggakan: tunggakan,
    last5: last5
  };
}

/**
 * 4b. Fungsi Mengambil Daftar Nama Warga terdaftar (untuk dropdown admin)
 */
function getWargaList() {
  const ss = getDatabase();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  let wargaList = [];
  
  // Mulai dari baris ke-2 (index 1) untuk melewatkan header
  for (let i = 1; i < data.length; i++) {
    let role = data[i][2];
    let nama = data[i][3];
    // Filter hanya user yang memiliki role 'warga'
    if (role === 'warga') {
      wargaList.push(nama);
    }
  }
  return wargaList;
}

/**
 * 5. Fungsi Admin Menyimpan Transaksi
 */
function simpanTransaksi(payload) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Kas');
    
    let tglObj = new Date(payload.tgl);
    
    sheet.appendRow([
      tglObj, 
      payload.jenis, 
      payload.nominal, 
      payload.ket, 
      payload.warga
    ]);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Helper: Mendapatkan spreadsheet database secara aman
 */
function getDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  
  // Jika standalone, cari file bernama 'Database Iuran Kas Warga' di Google Drive
  const files = DriveApp.getFilesByName('Database Iuran Kas Warga');
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.openById(file.getId());
  }
  
  throw new Error("Database Spreadsheet tidak ditemukan. Silakan jalankan fungsi 'initData' terlebih dahulu.");
}

