function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Iuran Kas Warga')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    try {
      ss = SpreadsheetApp.create('Database Iuran Kas Warga');
      Logger.log('Spreadsheet baru berhasil dibuat.');
    } catch (err) {
      throw new Error("Gagal membuat Spreadsheet: " + err.message);
    }
  }
  
  let sheetUsers = ss.getSheetByName('Users');
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(['Username', 'Password', 'Role', 'Nama_Lengkap']);
    sheetUsers.appendRow(['admin', 'admin123', 'admin', 'Bapak RT']);
    sheetUsers.appendRow(['warga', 'warga123', 'warga', 'Budi Santoso']);
    sheetUsers.getRange("A1:D1").setFontWeight("bold").setBackground("#d1fae5");
  }

  let sheetKas = ss.getSheetByName('Kas');
  if (!sheetKas) {
    sheetKas = ss.insertSheet('Kas');
    sheetKas.appendRow(['Tanggal', 'Jenis', 'Nominal', 'Keterangan', 'Nama_Warga']);
    sheetKas.getRange("A1:E1").setFontWeight("bold").setBackground("#d1fae5");
  }
  
  let sheetSet = ss.getSheetByName('Pengaturan');
  if (!sheetSet) {
    sheetSet = ss.insertSheet('Pengaturan');
    sheetSet.appendRow(['Kunci', 'Nilai']);
    sheetSet.appendRow(['Iuran_Bulanan', 50000]); 
    sheetSet.appendRow(['Tanggal_Mulai', '2026-01-01']); 
    sheetSet.getRange("A1:B1").setFontWeight("bold").setBackground("#d1fae5");
  }

  let defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  try { SpreadsheetApp.getUi().alert("Inisialisasi Database Berhasil!\n\nURL: " + ss.getUrl()); } 
  catch(e) { Logger.log("Berhasil inisialisasi"); }
}

function loginUser(username, password) {
  const ss = getDatabase();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return { success: true, role: data[i][2], name: data[i][3], username: data[i][0] };
    }
  }
  return { success: false, message: "Username atau Password salah!" };
}

// FUNGSI BARU: Mengambil list nama warga untuk form dropdown admin
function getWargaList() {
  const ss = getDatabase();
  const sheetUsers = ss.getSheetByName('Users');
  const data = sheetUsers.getDataRange().getValues();
  let list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === 'warga') {
      list.push(data[i][3]);
    }
  }
  return list;
}

// FUNGSI BARU: Dashboard Khusus Admin
function getAdminDashboard() {
  const ss = getDatabase();
  const sheetKas = ss.getSheetByName('Kas');
  const kasData = sheetKas.getDataRange().getValues();
  
  let totalPemasukan = 0;
  let totalPengeluaran = 0;
  let masukBulanIni = 0;
  let keluarBulanIni = 0;
  let historyKas = [];

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  for (let i = 1; i < kasData.length; i++) {
    let rawDate = new Date(kasData[i][0]);
    let tgl = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    let jenis = kasData[i][1];
    let nominal = parseFloat(kasData[i][2]);
    let ket = kasData[i][3];
    let namaUser = kasData[i][4];

    if (jenis == "Pemasukan") totalPemasukan += nominal;
    if (jenis == "Pengeluaran") totalPengeluaran += nominal;

    // Filter bulan ini
    if (rawDate.getMonth() === currentMonth && rawDate.getFullYear() === currentYear) {
      if (jenis == "Pemasukan") masukBulanIni += nominal;
      if (jenis == "Pengeluaran") keluarBulanIni += nominal;
    }

    historyKas.push({ tgl: tgl, jenis: jenis, nominal: nominal, ket: ket, warga: namaUser });
  }

  let saldoAkhir = totalPemasukan - totalPengeluaran;
  let last5 = historyKas.reverse().slice(0, 5);

  return {
    saldo: saldoAkhir,
    masukBulanIni: masukBulanIni,
    keluarBulanIni: keluarBulanIni,
    last5: last5
  };
}

function getWargaDashboard(username) {
  const ss = getDatabase();
  const sheetKas = ss.getSheetByName('Kas');
  const sheetUsers = ss.getSheetByName('Users');
  const sheetSet = ss.getSheetByName('Pengaturan');
  
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

  for (let i = 1; i < kasData.length; i++) {
    let tgl = Utilities.formatDate(new Date(kasData[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    let jenis = kasData[i][1];
    let nominal = parseFloat(kasData[i][2]);
    let ket = kasData[i][3];
    let namaUser = kasData[i][4];

    if (jenis == "Pemasukan") totalPemasukan += nominal;
    if (jenis == "Pengeluaran") totalPengeluaran += nominal;
    if (jenis == "Pemasukan" && namaUser == namaWarga) totalIuranSaya += nominal;

    historyKas.push({ tgl: tgl, jenis: jenis, nominal: nominal, ket: ket, warga: namaUser });
  }

  let saldoAkhir = totalPemasukan - totalPengeluaran;
  
  let iuranBulanan = 50000;
  let tanggalMulaiStr = '2026-01-01'; 
  try {
    const setValues = sheetSet.getDataRange().getValues();
    for (let i = 1; i < setValues.length; i++) {
      if (setValues[i][0] === 'Iuran_Bulanan') iuranBulanan = parseFloat(setValues[i][1]);
      if (setValues[i][0] === 'Tanggal_Mulai') {
        tanggalMulaiStr = (setValues[i][1] instanceof Date) ? Utilities.formatDate(setValues[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd") : setValues[i][1].toString();
      }
    }
  } catch(e) {}
  
  let tglMulai = new Date(tanggalMulaiStr);
  let tglHariIni = new Date();
  
  let tahunDiff = tglHariIni.getFullYear() - tglMulai.getFullYear();
  let bulanDiff = tglHariIni.getMonth() - tglMulai.getMonth();
  let jumlahBulanBerjalan = (tahunDiff * 12) + bulanDiff + 1; 
  if (jumlahBulanBerjalan < 0) jumlahBulanBerjalan = 0;
  
  let totalKewajiban = jumlahBulanBerjalan * iuranBulanan; 
  let tunggakan = totalKewajiban - totalIuranSaya;
  let last5 = historyKas.reverse().slice(0, 5);

  return { saldo: saldoAkhir, tunggakan: tunggakan, last5: last5 };
}

// FUNGSI BARU: Untuk Export PDF
function getLaporanFilter(startStr, endStr) {
  const ss = getDatabase();
  const sheetKas = ss.getSheetByName('Kas');
  const kasData = sheetKas.getDataRange().getValues();
  let historyKas = [];
  
  const startDate = new Date(startStr);
  startDate.setHours(0,0,0,0);
  const endDate = new Date(endStr);
  endDate.setHours(23,59,59,999);

  for (let i = 1; i < kasData.length; i++) {
    let rawDate = new Date(kasData[i][0]);
    if (rawDate >= startDate && rawDate <= endDate) {
      let tgl = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      historyKas.push({
        tgl: tgl,
        jenis: kasData[i][1],
        nominal: parseFloat(kasData[i][2]),
        ket: kasData[i][3],
        warga: kasData[i][4]
      });
    }
  }
  return historyKas;
}

function simpanTransaksi(payload) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Kas');
    let tglObj = new Date(payload.tgl);
    sheet.appendRow([tglObj, payload.jenis, payload.nominal, payload.ket, payload.warga]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  const files = DriveApp.getFilesByName('Database Iuran Kas Warga');
  if (files.hasNext()) return SpreadsheetApp.openById(files.next().getId());
  throw new Error("Database Spreadsheet tidak ditemukan.");
}


