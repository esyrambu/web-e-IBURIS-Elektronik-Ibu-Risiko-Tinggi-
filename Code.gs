/**
 * Aplikasi e-IBURIS - Deteksi Dini Risiko Tinggi Ibu Hamil
 * Backend Code - Google Apps Script (code.gs)
 * Berdasarkan standar Buku KIA 2024 dan Kartu Skor Poedji Rochjati (KSPR)
 */

// Fungsi untuk melayani file HTML
function doGet(e) {
  var page = e.parameter ? e.parameter.page : null;
  
  // Jika URL membuka halaman dashboard (?page=dashboard)
  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('Dashboard - e-IBURIS')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Halaman Utama / Landing Page (Default)
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('e-IBURIS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi pendukung untuk mengambil URL Script
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Menyiapkan Spreadsheet Database Otomatis
function setupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.create("Database_e_IBURIS");
    }
    
    // 1. Sheet Pengguna / Akun Sistem
    var sheetUsers = ss.getSheetByName("Pengguna") || ss.insertSheet("Pengguna");
    if (sheetUsers.getLastRow() === 0) {
      sheetUsers.appendRow(["Username", "Password", "Nama Lengkap", "Role", "Faskes", "Tanggal Dibuat"]);
      sheetUsers.appendRow(["admin", "admin123", "Bidan Administrator", "Bidan", "Rumah Sakit Umum Hoba Kalla", new Date()]);
      sheetUsers.appendRow(["bidan1", "bidan123", "Bdn. Theresia Rambu Melani Ma'laof Pedjaga,S.Tr.Keb", "Bidan", "Rumah Sakit Umum Hoba Kalla", new Date()]);
    }

    // 2. Sheet Data Ibu Hamil (Buku KIA 2024)
    var sheetIbu = ss.getSheetByName("Data_Ibu") || ss.insertSheet("Data_Ibu");
    if (sheetIbu.getLastRow() === 0) {
      sheetIbu.appendRow([
        "ID_Ibu", "NIK", "Nama_Ibu", "Nama_Suami", "Tanggal_Lahir", "Umur", 
        "Alamat", "No_HP", "Gol_Darah", "Pekerjaan", "Pendidikan", 
        "HPHT", "Taksiran_Persalinan", "G", "P", "A", "Tanggal_Daftar"
      ]);
      // Mock data awal
      sheetIbu.appendRow([
        "IBU-001", "3201234567890001", "Rina Rahmawati", "Budi Santoso", "1994-05-12", "32",
        "Jl. Kenanga No. 12, RT 02/RW 05", "081234567890", "A", "Ibu Rumah Tangga", "SMA",
        "2025-10-15", "2026-07-22", "2", "1", "0", new Date()
      ]);
      sheetIbu.appendRow([
        "IBU-002", "3201234567890002", "Siti Maryam", "Ahmad Fauzi", "2007-08-20", "18",
        "Kampung Melati RT 01/RW 03", "085712345678", "O", "Swasta", "SMP",
        "2026-01-05", "2026-10-12", "1", "0", "0", new Date()
      ]);
    }

    // 3. Sheet Skrining KSPR
    var sheetSkrining = ss.getSheetByName("Skrining_KSPR") || ss.insertSheet("Skrining_KSPR");
    if (sheetSkrining.getLastRow() === 0) {
      sheetSkrining.appendRow([
        "ID_Skrining", "ID_Ibu", "Tanggal_Skrining", "Skor_Awal", 
        "Faktor_1_Pilihan", "Faktor_2_Pilihan", "Faktor_3_Pilihan", 
        "Skor_Total", "Kategori_Risiko", "Bidan_Pemeriksa", "Rekomendasi"
      ]);
      sheetSkrining.appendRow([
        "SKR-001", "IBU-001", "2026-02-15", "2",
        "Terlalu lambat hamil I (nikah >= 4 th) (+4)", "", "",
        "6", "KRT (Risiko Tinggi)", "Bidan Administrator", "Edukasi intensif, kontrol rutin bidan"
      ]);
      sheetSkrining.appendRow([
        "SKR-002", "IBU-002", "2026-03-10", "2",
        "Terlalu muda (umur <= 16 th) (+4)", "Kurang Energi Kronis (KEK) (+4)", "",
        "10", "KRT (Risiko Tinggi)", "Siti Aminah, A.Md.Keb", "Rujuk konsultasi Dokter Spesialis Obgyn di T1"
      ]);
    }

    // 4. Sheet Kunjungan ANC
    var sheetANC = ss.getSheetByName("Kunjungan_ANC") || ss.insertSheet("Kunjungan_ANC");
    if (sheetANC.getLastRow() === 0) {
      sheetANC.appendRow([
        "ID_ANC", "ID_Ibu", "Tanggal_Kunjungan", "Usia_Kehamilan_Minggu", 
        "Berat_Badan", "Tekanan_Darah", "Tinggi_Fundus", "Denyut_Jantung_Janin", 
        "Status_Imunisasi_TT", "Tablet_Fe", "Keluhan", "Pemeriksa", "Rencana_Tindak_Lanjut"
      ]);
      sheetANC.appendRow([
        "ANC-001", "IBU-001", "2026-02-15", "17", "58", "110/80", "15", "142", "TT2", "Ya", "Mual ringan", "Bidan Administrator", "Konsumsi suplemen besi, kontrol 4 minggu lagi"
      ]);
    }

    // 5. Sheet Rujukan
    var sheetRujukan = ss.getSheetByName("Rujukan") || ss.insertSheet("Rujukan");
    if (sheetRujukan.getLastRow() === 0) {
      sheetRujukan.appendRow([
        "ID_Rujukan", "ID_Ibu", "Tanggal_Rujukan", "Faskes_Tujuan", 
        "Diagnosis_Penyerta", "Alasan_Rujukan", "Transportasi", "Status_Rujukan", "Catatan"
      ]);
      sheetRujukan.appendRow([
        "RUJ-001", "IBU-002", "2026-03-11", "RSUD Sayang Ibu", "Hamil risiko tinggi (Usia muda & KEK)", "Penanganan Dokter Spesialis Kebidanan", "Ambulans Rumah Sakit", "Selesai", "Rekomendasi melahirkan di RS"
      ]);
    }

    return { success: true, message: "Database e-IBURIS berhasil dibuat dan diinisialisasi!" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Fungsi Autentikasi Login Custom
function loginUser(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Pengguna");
    if (!sheet) return { success: false, message: "Database belum disetup. Jalankan fungsi setupDatabase() terlebih dahulu." };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === username.toLowerCase() && data[i][1].toString() === password) {
        return {
          success: true,
          user: {
            username: data[i][0],
            namaLengkap: data[i][2],
            role: data[i][3],
            faskes: data[i][4]
          }
        };
      }
    }
    return { success: false, message: "Username atau Password salah!" };
  } catch (error) {
    return { success: false, message: "Terjadi kesalahan sistem: " + error.toString() };
  }
}

// Fungsi mengambil semua Data Ibu Hamil
function getMothers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data_Ibu");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var keys = data[0];
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < keys.length; j++) {
        var val = data[i][j];
        if (val instanceof Date) {
          obj[keys[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          obj[keys[j]] = val;
        }
      }
      list.push(obj);
    }
    return list;
  } catch (error) {
    return [];
  }
}

// Fungsi menambah data Ibu Hamil Baru
function addMother(motherData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data_Ibu");
    if (!sheet) return { success: false, message: "Database Data_Ibu tidak ditemukan!" };
    
    // Generate ID_Ibu
    var lastRow = sheet.getLastRow();
    var idNum = lastRow;
    var idStr = "IBU-" + String(idNum).padStart(3, '0');
    
    var rowData = [
      idStr,
      motherData.nik,
      motherData.namaIbu,
      motherData.namaSuami,
      motherData.tanggalLahir,
      motherData.umur,
      motherData.alamat,
      motherData.noHp,
      motherData.golDarah,
      motherData.pekerjaan,
      motherData.pendidikan,
      motherData.hpht,
      motherData.taksiranPersalinan,
      motherData.g,
      motherData.p,
      motherData.a,
      new Date()
    ];
    
    sheet.appendRow(rowData);
    return { success: true, id: idStr, message: "Ibu hamil berhasil didaftarkan!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Fungsi menyimpan Skrining KSPR baru
function saveScreening(screeningData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Skrining_KSPR");
    if (!sheet) return { success: false, message: "Database Skrining_KSPR tidak ditemukan!" };
    
    var lastRow = sheet.getLastRow();
    var idStr = "SKR-" + String(lastRow).padStart(3, '0');
    
    sheet.appendRow([
      idStr,
      screeningData.idIbu,
      screeningData.tanggalSkrining,
      screeningData.skorAwal || "2",
      screeningData.faktor1 || "",
      screeningData.faktor2 || "",
      screeningData.faktor3 || "",
      screeningData.skorTotal,
      screeningData.kategoriRisiko,
      screeningData.bidanPemeriksa,
      screeningData.rekomendasi
    ]);
    
    return { success: true, id: idStr, message: "Skrining KSPR berhasil disimpan!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Mengambil seluruh data Skrining
function getScreenings() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Skrining_KSPR");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var keys = data[0];
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < keys.length; j++) {
        var val = data[i][j];
        if (val instanceof Date) {
          obj[keys[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          obj[keys[j]] = val;
        }
      }
      list.push(obj);
    }
    return list;
  } catch (error) {
    return [];
  }
}

// Fungsi menyimpan Kunjungan ANC baru
function saveANC(ancData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kunjungan_ANC");
    if (!sheet) return { success: false, message: "Database Kunjungan_ANC tidak ditemukan!" };
    
    var lastRow = sheet.getLastRow();
    var idStr = "ANC-" + String(lastRow).padStart(3, '0');
    
    sheet.appendRow([
      idStr,
      ancData.idIbu,
      ancData.tanggalKunjungan,
      ancData.usiaKehamilan,
      ancData.beratBadan,
      ancData.tekananDarah,
      ancData.tinggiFundus,
      ancData.djj,
      ancData.statusImunisasi,
      ancData.tabletFe,
      ancData.keluhan,
      ancData.pemeriksa,
      ancData.rtl
    ]);
    
    return { success: true, id: idStr, message: "Kunjungan ANC berhasil dicatat!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Mengambil seluruh data Kunjungan ANC
function getANCs() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kunjungan_ANC");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var keys = data[0];
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < keys.length; j++) {
        var val = data[i][j];
        if (val instanceof Date) {
          obj[keys[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          obj[keys[j]] = val;
        }
      }
      list.push(obj);
    }
    return list;
  } catch (error) {
    return [];
  }
}

// Fungsi menyimpan Rujukan
function saveReferral(refData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Rujukan Rumah Sakit");
    if (!sheet) return { success: false, message: "Database Rujukan tidak ditemukan!" };
    
    var lastRow = sheet.getLastRow();
    var idStr = "RUJ-" + String(lastRow).padStart(3, '0');
    
    sheet.appendRow([
      idStr,
      refData.idIbu,
      refData.tanggalRujukan,
      refData.faskesTujuan,
      refData.diagnosis,
      refData.alasan,
      refData.transportasi,
      refData.statusRujukan || "Aktif",
      refData.catatan || ""
    ]);
    
    return { success: true, id: idStr, message: "Surat rujukan berhasil diterbitkan!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Mengambil seluruh data Rujukan
function getReferrals() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Rujukan Rumah Sakit");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var keys = data[0];
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < keys.length; j++) {
        var val = data[i][j];
        if (val instanceof Date) {
          obj[keys[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          obj[keys[j]] = val;
        }
      }
      list.push(obj);
    }
    return list;
  } catch (error) {
    return [];
  }
}

// Update Pengaturan Akun Admin / Bidan
function saveSettings(username, settingsData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Pengguna");
    if (!sheet) return { success: false, message: "Database pengguna tidak ditemukan!" };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === username.toLowerCase()) {
        sheet.getCell(i + 1, 2).setValue(settingsData.password || data[i][1]);
        sheet.getCell(i + 1, 3).setValue(settingsData.namaLengkap || data[i][2]);
        sheet.getCell(i + 1, 5).setValue(settingsData.faskes || data[i][4]);
        return { success: true, message: "Pengaturan profil berhasil diperbarui!" };
      }
    }
    return { success: false, message: "User tidak ditemukan!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
