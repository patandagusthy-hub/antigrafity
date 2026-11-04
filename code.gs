/**
 * =========================================================================
 * GUSTHYPALINPATANDAEXAM - GOOGLE APPS SCRIPT BACKEND (Code.gs)
 * CBT EXAM ENGINE & GOOGLE SPREADSHEET REAL-TIME AUTO-SAVE DATABASE
 * =========================================================================
 * Versi          : 2.7.0-Enterprise Realtime
 * SPREADSHEET ID : 1fkrsi8HGxtSnd1_7xofU-LnNeKlqSRgttj5NYMHgVzo
 * FOLDER DRIVE ID: 1dkeihpLKEzQXIMfaHIAEAkj1Uq5vJCv-
 * 
 * FITUR UTAMA:
 * 1. Koneksi langsung ke Google Spreadsheet ID: 1fkrsi8HGxtSnd1_7xofU-LnNeKlqSRgttj5NYMHgVzo
 * 2. Inisialisasi otomatis tab sheet:
 *    - Sheet 'Hasil_Ujian' : [Timestamp, NISN, Nama_Siswa, Kelas, Skor_Akhir, Durasi, Status_Kecurangan, Detail_Jawaban_JSON]
 *    - Sheet 'Data_Siswa'  : [NISN, Nama_Siswa, Kelas, Password_Barcode, Status_Login]
 *    - Sheet 'Bank_Soal'   : [ID_Soal, Tipe_Soal, Kategori, Pertanyaan, Opsi_JSON, Kunci_Jawaban, Bobot]
 * 3. Concurrency-Safe (LockService 30.000 ms) untuk submit 1.000+ peserta serentak tanpa race condition.
 * 4. Fungsi real-time:
 *    - simpanHasilUjian(payload)
 *    - simpanDataSiswa(payload)
 *    - simpanBankSoal(payload)
 *    - ambilDataDariSheet()
 * 5. Dual-Storage Architecture: Snapshot JSON otomatis tersimpan ke Folder Drive 1dkeihpLKEzQXIMfaHIAEAkj1Uq5vJCv-
 * 6. Manajemen Kredensial Admin & Konfigurasi Fitur Aplikasi (CONFIG_APP_GPP.json).
 * =========================================================================
 */

// =============================================================================
// KONFIGURASI GLOBAL SISTEM & DATABASE
// =============================================================================
var SPREADSHEET_ID = "1fkrsi8HGxtSnd1_7xofU-LnNeKlqSRgttj5NYMHgVzo";
var TARGET_FOLDER_ID = "1dkeihpLKEzQXIMfaHIAEAkj1Uq5vJCv-";
var BACKUP_SUBFOLDER_NAME = "Backup_Siswa_JSON";
var CONFIG_APP_FILENAME = "CONFIG_APP_GPP.json";
var STATUS_LOGIN_FILENAME = "STATUS_LOGIN_SISWA.json";
var TIMEZONE = "Asia/Makassar"; // SMAN 1 Sebatik (WITA / Kalimantan Utara)
var LOCK_TIMEOUT_MS = 30000;   // 30 detik untuk antrean concurrency 1000+ siswa

// Backward compatibility alias
var CONFIG = {
  SPREADSHEET_ID: SPREADSHEET_ID,
  TARGET_FOLDER_ID: TARGET_FOLDER_ID,
  BACKUP_SUBFOLDER_NAME: BACKUP_SUBFOLDER_NAME,
  CONFIG_APP_FILENAME: CONFIG_APP_FILENAME,
  TIMEZONE: TIMEZONE,
  LOCK_TIMEOUT_MS: LOCK_TIMEOUT_MS,
  DEFAULT_ADMIN_USER: "admin",
  DEFAULT_ADMIN_PASS: "admin123"
};

// =============================================================================
// 1. INISIALISASI OTOMATIS DATABASE SPREADSHEET (setupDatabase)
// =============================================================================
/**
 * Mengecek apakah tab sheet database sudah ada; jika belum, otomatis membuat sheet beserta header dan formatnya.
 * a. Sheet 'Hasil_Ujian': [Timestamp, NISN, Nama_Siswa, Kelas, Skor_Akhir, Durasi, Status_Kecurangan, Detail_Jawaban_JSON]
 * b. Sheet 'Data_Siswa' : [NISN, Nama_Siswa, Kelas, Password_Barcode, Status_Login]
 * c. Sheet 'Bank_Soal'  : [ID_Soal, Tipe_Soal, Kategori, Pertanyaan, Opsi_JSON, Kunci_Jawaban, Bobot]
 */
function setupDatabase() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (err) {
    throw new Error("Gagal membuka Spreadsheet ID: " + SPREADSHEET_ID + ". Pastikan ID benar dan akun GAS memiliki izin Edit. Error: " + err.message);
  }

  var sheetsCreated = [];

  // 1. Sheet Hasil_Ujian
  var sheetHasil = ss.getSheetByName("Hasil_Ujian");
  if (!sheetHasil) {
    sheetHasil = ss.insertSheet("Hasil_Ujian");
    var headersHasil = ["Timestamp", "NISN", "Nama_Siswa", "Kelas", "Skor_Akhir", "Durasi", "Status_Kecurangan", "Detail_Jawaban_JSON"];
    sheetHasil.appendRow(headersHasil);
    var hrHasil = sheetHasil.getRange(1, 1, 1, headersHasil.length);
    hrHasil.setBackground("#0f172a"); // Dark Slate
    hrHasil.setFontColor("#38bdf8"); // Cyan 400
    hrHasil.setFontWeight("bold");
    hrHasil.setHorizontalAlignment("center");
    sheetHasil.setFrozenRows(1);
    sheetHasil.autoResizeColumns(1, headersHasil.length);
    sheetsCreated.push("Hasil_Ujian");
  }

  // 2. Sheet Data_Siswa
  var sheetSiswa = ss.getSheetByName("Data_Siswa");
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet("Data_Siswa");
    var headersSiswa = ["NISN", "Nama_Siswa", "Kelas", "Password_Barcode", "Status_Login"];
    sheetSiswa.appendRow(headersSiswa);
    var hrSiswa = sheetSiswa.getRange(1, 1, 1, headersSiswa.length);
    hrSiswa.setBackground("#064e3b"); // Dark Emerald
    hrSiswa.setFontColor("#34d399"); // Emerald 400
    hrSiswa.setFontWeight("bold");
    hrSiswa.setHorizontalAlignment("center");
    sheetSiswa.setFrozenRows(1);
    sheetSiswa.autoResizeColumns(1, headersSiswa.length);
    sheetsCreated.push("Data_Siswa");
  }

  // 3. Sheet Bank_Soal
  var sheetSoal = ss.getSheetByName("Bank_Soal");
  if (!sheetSoal) {
    sheetSoal = ss.insertSheet("Bank_Soal");
    var headersSoal = ["ID_Soal", "Tipe_Soal", "Kategori", "Pertanyaan", "Opsi_JSON", "Kunci_Jawaban", "Bobot"];
    sheetSoal.appendRow(headersSoal);
    var hrSoal = sheetSoal.getRange(1, 1, 1, headersSoal.length);
    hrSoal.setBackground("#1e1b4b"); // Dark Indigo
    hrSoal.setFontColor("#818cf8"); // Indigo 400
    hrSoal.setFontWeight("bold");
    hrSoal.setHorizontalAlignment("center");
    sheetSoal.setFrozenRows(1);
    sheetSoal.autoResizeColumns(1, headersSoal.length);
    sheetsCreated.push("Bank_Soal");
  }

  // 4. Pastikan Subfolder Backup Drive juga siap
  try {
    var folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    getOrCreateSubfolder(folder, BACKUP_SUBFOLDER_NAME);
  } catch (driveErr) {
    Logger.log("Peringatan Drive Folder: " + driveErr.message);
  }

  return {
    status: "success",
    message: "Inisialisasi database Spreadsheet CBT berhasil!",
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetUrl: ss.getUrl(),
    newlyCreated: sheetsCreated
  };
}

// =============================================================================
// 2. FUNGSI SIMPAN HASIL UJIAN (simpanHasilUjian)
// =============================================================================
/**
 * Menyimpan hasil pengerjaan ujian siswa langsung ke Google Spreadsheet sheet Hasil_Ujian
 * Dilengkapi LockService (waitLock 30000ms) dan sheet.appendRow() agar aman untuk 1000+ peserta serentak.
 */
function simpanHasilUjian(payload) {
  if (!payload) {
    return { status: "error", message: "Data payload ujian kosong" };
  }

  if (typeof payload === "string") {
    try { payload = JSON.parse(payload); } catch (e) {}
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
  } catch (lockErr) {
    return {
      status: "busy",
      message: "Server sedang memproses antrean peserta lain. Antrean lokal akan mencoba kembali dalam beberapa detik.",
      retryInSeconds: 3
    };
  }

  try {
    // 1. Buka spreadsheet dan ambil sheet Hasil_Ujian
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Hasil_Ujian");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("Hasil_Ujian");
    }

    // 2. Parsing & Normalisasi Data
    var timestamp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    var nisn = "'" + String(payload.nisn || "-").trim();
    var namaSiswa = String(payload.studentName || payload.nama || payload.name || "-");
    var kelas = String(payload.className || payload.kelas || payload.classId || "-");
    var skorAkhir = Number(payload.finalScore !== undefined ? payload.finalScore : (payload.score || 0));
    var durasi = payload.durationFormatted || (payload.durationSeconds ? (Math.floor(payload.durationSeconds / 60) + " menit " + (payload.durationSeconds % 60) + " detik") : (payload.durasi || "-"));

    var statusKecurangan = "";
    if (payload.violations && Array.isArray(payload.violations) && payload.violations.length > 0) {
      statusKecurangan = payload.violations.length + "x Terdeteksi (" + payload.violations.map(function(v) { return v.reason || v.type || "Pindah Tab"; }).join("; ") + ")";
    } else if (payload.violationsCount > 0) {
      statusKecurangan = payload.violationsCount + "x Terdeteksi";
    } else {
      statusKecurangan = "Tertib / Bersih";
    }

    var detailJawabanJson = "";
    if (typeof payload.answers === "string") {
      detailJawabanJson = payload.answers;
    } else {
      detailJawabanJson = JSON.stringify(payload.answers || payload.detailJawaban || {});
    }

    // 3. Simpan Baris ke Sheet Hasil_Ujian
    // [Timestamp, NISN, Nama_Siswa, Kelas, Skor_Akhir, Durasi, Status_Kecurangan, Detail_Jawaban_JSON]
    var rowData = [
      timestamp,
      nisn,
      namaSiswa,
      kelas,
      skorAkhir,
      durasi,
      statusKecurangan,
      detailJawabanJson
    ];
    sheet.appendRow(rowData);

    // 4. Simpan Snapshot JSON Backup ke Google Drive Folder ID 1dkeihpLKEzQXIMfaHIAEAkj1Uq5vJCv-
    var backupInfo = null;
    try {
      backupInfo = saveStudentJsonBackupToDrive(payload);
    } catch (bErr) {
      Logger.log("Peringatan simpan backup Drive: " + bErr.message);
    }

    return {
      status: "success",
      message: "Hasil ujian berhasil disimpan secara real-time ke Google Spreadsheet!",
      studentName: namaSiswa,
      nisn: payload.nisn,
      score: skorAkhir,
      timestamp: timestamp,
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetUrl: ss.getUrl(),
      backupFileId: backupInfo ? backupInfo.fileId : null
    };

  } catch (err) {
    Logger.log("Error simpanHasilUjian: " + err.toString());
    return {
      status: "error",
      message: "Gagal menyimpan hasil ujian ke Spreadsheet: " + err.message
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

// Backward compatibility alias
function submitExamResult(payload) {
  return simpanHasilUjian(payload);
}

// =============================================================================
// 3. FUNGSI SIMPAN DATA SISWA (simpanDataSiswa)
// =============================================================================
/**
 * Menyimpan data siswa ke Google Spreadsheet sheet Data_Siswa
 * Format Header: [NISN, Nama_Siswa, Kelas, Password_Barcode, Status_Login]
 */
function simpanDataSiswa(payload) {
  if (!payload) {
    return { status: "error", message: "Data siswa kosong" };
  }

  if (typeof payload === "string") {
    try { payload = JSON.parse(payload); } catch (e) {}
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
  } catch (lockErr) {
    return { status: "busy", message: "Server sibuk memproses antrean simpan data siswa." };
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Data_Siswa");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("Data_Siswa");
    }

    // Kasus 1: Tambah 1 siswa tunggal (appendOnly)
    if (payload.appendOnly && payload.student) {
      var s = payload.student;
      var singleRow = [
        "'" + String(s.nisn || "-").trim(),
        s.name || s.Nama_Siswa || "-",
        s.classId || s.kelas || s.Kelas || "-",
        s.token || s.Password_Barcode || s.password || `GPP-${s.classId || "X"}-${String(s.nisn || "").slice(-3)}`,
        (s.isLocked || s.Status_Login === "TERKUNCI") ? "TERKUNCI" : "AKTIF"
      ];
      sheet.appendRow(singleRow);
      return { status: "success", message: "1 data siswa berhasil disimpan ke Spreadsheet!", count: 1 };
    }

    // Kasus 2: Sinkronisasi daftar array siswa
    var studentsList = Array.isArray(payload) ? payload : (payload.students || [payload]);
    var rows = studentsList.map(function(s) {
      return [
        "'" + String(s.nisn || "-").trim(),
        s.name || s.Nama_Siswa || "-",
        s.classId || s.kelas || s.Kelas || "-",
        s.token || s.Password_Barcode || s.password || `GPP-${s.classId || "X"}-${String(s.nisn || "").slice(-3)}`,
        (s.isLocked || s.Status_Login === "TERKUNCI") ? "TERKUNCI" : "AKTIF"
      ];
    });

    if (payload.appendOnly) {
      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }
    } else {
      // Overwrite/Sync Seluruh Data (kosongkan baris setelah header baris 1)
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
      }
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, 5).setValues(rows);
      }
    }

    return {
      status: "success",
      message: "Berhasil menyinkronkan " + rows.length + " data siswa ke Google Spreadsheet!",
      count: rows.length,
      spreadsheetId: SPREADSHEET_ID
    };

  } catch (err) {
    Logger.log("Error simpanDataSiswa: " + err.toString());
    return { status: "error", message: "Gagal menyimpan data siswa: " + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Backward compatibility alias
function syncStudentsToSheet(studentsList) {
  return simpanDataSiswa(studentsList);
}

// =============================================================================
// 4. FUNGSI SIMPAN BANK SOAL (simpanBankSoal)
// =============================================================================
/**
 * Menyimpan data bank soal ke Google Spreadsheet sheet Bank_Soal
 * Format Header: [ID_Soal, Tipe_Soal, Kategori, Pertanyaan, Opsi_JSON, Kunci_Jawaban, Bobot]
 */
function simpanBankSoal(payload) {
  if (!payload) {
    return { status: "error", message: "Data bank soal kosong" };
  }

  if (typeof payload === "string") {
    try { payload = JSON.parse(payload); } catch (e) {}
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
  } catch (lockErr) {
    return { status: "busy", message: "Server sibuk memproses antrean simpan bank soal." };
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Bank_Soal");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("Bank_Soal");
    }

    var questionsList = Array.isArray(payload) ? payload : (payload.questions || [payload]);

    var rows = questionsList.map(function(q, idx) {
      var opsiStr = "";
      if (typeof q.options === "string") {
        opsiStr = q.options;
      } else if (q.options || q.pairs) {
        opsiStr = JSON.stringify(q.options || q.pairs);
      } else {
        opsiStr = "[]";
      }

      var kunciStr = "";
      if (Array.isArray(q.correctAnswer)) {
        kunciStr = q.correctAnswer.join(",");
      } else {
        kunciStr = String(q.correctAnswer !== undefined ? q.correctAnswer : (q.answer || "-"));
      }

      return [
        q.id || ("Q" + (idx + 1)),
        q.type || q.Tipe_Soal || "PILIHAN_GANDA",
        q.category || q.Kategori || "Umum",
        q.question || q.Pertanyaan || "",
        opsiStr,
        kunciStr,
        Number(q.points || q.Bobot || 10)
      ];
    });

    if (payload.appendOnly) {
      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }
    } else {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
      }
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, 7).setValues(rows);
      }
    }

    return {
      status: "success",
      message: "Berhasil menyimpan " + rows.length + " butir soal ke Google Spreadsheet!",
      count: rows.length,
      spreadsheetId: SPREADSHEET_ID
    };

  } catch (err) {
    Logger.log("Error simpanBankSoal: " + err.toString());
    return { status: "error", message: "Gagal menyimpan bank soal: " + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Backward compatibility alias
function syncQuestionsToSheet(questionsList) {
  return simpanBankSoal(questionsList);
}

// =============================================================================
// 5. FUNGSI TARIK DATA DARI SHEET (ambilDataDariSheet)
// =============================================================================
/**
 * Mengambil data siswa, bank soal, dan hasil ujian langsung dari Google Spreadsheet
 * Membaca sheet Data_Siswa, Bank_Soal, dan Hasil_Ujian
 */
function ambilDataDariSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    setupDatabase(); // Pastikan struktur sheet database siap

    // 1. Ambil Data Siswa dari sheet Data_Siswa
    var sheetSiswa = ss.getSheetByName("Data_Siswa");
    var students = [];
    if (sheetSiswa && sheetSiswa.getLastRow() > 1) {
      var valsSiswa = sheetSiswa.getRange(2, 1, sheetSiswa.getLastRow() - 1, 5).getValues();
      for (var i = 0; i < valsSiswa.length; i++) {
        var r = valsSiswa[i];
        var nisnStr = String(r[0]).replace(/^'/, "").trim();
        var namaStr = String(r[1]).trim();
        var kelasStr = String(r[2]).trim();
        var tokenStr = String(r[3]).trim();
        var statusStr = String(r[4]).trim();
        if (nisnStr || namaStr) {
          students.push({
            id: "S" + (1001 + i),
            nisn: nisnStr,
            name: namaStr,
            classId: kelasStr || "X-1",
            username: namaStr.toLowerCase().replace(/[^a-z0-9]/g, "").slice(0, 8),
            token: tokenStr || `GPP-${kelasStr}-${nisnStr.slice(-3)}`,
            isLocked: statusStr.toUpperCase() === "TERKUNCI",
            examStatus: "Belum Mulai",
            score: 0,
            violations: 0,
            violationsList: [],
            lastActivity: "-"
          });
        }
      }
    }

    // 2. Ambil Bank Soal dari sheet Bank_Soal
    var sheetSoal = ss.getSheetByName("Bank_Soal");
    var questions = [];
    if (sheetSoal && sheetSoal.getLastRow() > 1) {
      var valsSoal = sheetSoal.getRange(2, 1, sheetSoal.getLastRow() - 1, 7).getValues();
      for (var j = 0; j < valsSoal.length; j++) {
        var qRow = valsSoal[j];
        var qId = String(qRow[0] || ("Q" + (j + 1)));
        var qType = String(qRow[1] || "PILIHAN_GANDA");
        var qCat = String(qRow[2] || "Umum");
        var qText = String(qRow[3] || "");
        var qOpsiRaw = String(qRow[4] || "[]");
        var qKunci = String(qRow[5] || "");
        var qBobot = Number(qRow[6]) || 10;

        var parsedOptions = null;
        var parsedPairs = null;
        try {
          var parsed = JSON.parse(qOpsiRaw);
          if (qType === "MENJODOHKAN") {
            parsedPairs = parsed;
          } else {
            parsedOptions = parsed;
          }
        } catch (e) {
          parsedOptions = [];
        }

        var correctAnswer = qKunci;
        if (qType === "PILIHAN_GANDA_KOMPLEKS" && qKunci.indexOf(",") !== -1) {
          correctAnswer = qKunci.split(",").map(function(k) { return k.trim(); });
        }

        questions.push({
          id: qId,
          number: j + 1,
          type: qType,
          category: qCat,
          question: qText,
          options: parsedOptions,
          pairs: parsedPairs,
          correctAnswer: correctAnswer,
          points: qBobot,
          audioUrl: qType === "AUDIO_LISTENING" ? "https://actions.google.com/sounds/v1/science/digital_beeps.ogg" : null,
          audioMaxPlays: 2
        });
      }
    }

    // 3. Ambil Rekap Hasil Ujian dari sheet Hasil_Ujian
    var sheetHasil = ss.getSheetByName("Hasil_Ujian");
    var examResults = [];
    if (sheetHasil && sheetHasil.getLastRow() > 1) {
      var valsHasil = sheetHasil.getRange(2, 1, sheetHasil.getLastRow() - 1, 8).getValues();
      for (var k = 0; k < valsHasil.length; k++) {
        var hRow = valsHasil[k];
        examResults.push({
          timestamp: hRow[0],
          nisn: String(hRow[1]).replace(/^'/, ""),
          name: hRow[2],
          classId: hRow[3],
          score: Number(hRow[4]) || 0,
          duration: hRow[5],
          violations: hRow[6],
          answers: hRow[7]
        });
      }
    }

    return {
      status: "success",
      message: "Data berhasil ditarik dari Google Spreadsheet!",
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetUrl: ss.getUrl(),
      studentsCount: students.length,
      questionsCount: questions.length,
      examResultsCount: examResults.length,
      students: students,
      questions: questions,
      examResults: examResults
    };

  } catch (err) {
    Logger.log("Error ambilDataDariSheet: " + err.toString());
    return {
      status: "error",
      message: "Gagal mengambil data dari Google Spreadsheet: " + err.message,
      students: [],
      questions: [],
      examResults: []
    };
  }
}

// Backward compatibility alias
function getStudentDataFromSheet() {
  var res = ambilDataDariSheet();
  return { status: res.status, count: res.studentsCount, students: res.students };
}

function getQuestionsFromSheet() {
  var res = ambilDataDariSheet();
  return { status: res.status, count: res.questionsCount, questions: res.questions };
}

function getExamRecapGAS() {
  var res = ambilDataDariSheet();
  return { status: res.status, count: res.examResultsCount, data: res.examResults };
}

// =============================================================================
// 6. SNAPSHOT BACKUP JSON KE GOOGLE DRIVE
// =============================================================================
/**
 * Menyimpan snapshot backup JSON lengkap ke subfolder Drive 'Backup_Siswa_JSON'
 */
function saveStudentJsonBackupToDrive(payload) {
  var folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  var backupFolder = getOrCreateSubfolder(folder, BACKUP_SUBFOLDER_NAME);
  var nisn = String(payload.nisn || "UNKNOWN").replace(/[^a-zA-Z0-9_-]/g, "");
  var timeStr = Utilities.formatDate(new Date(), TIMEZONE, "yyyyMMdd_HHmmss");
  var fileName = "HASIL_UJIAN_" + nisn + "_" + timeStr + ".json";

  var jsonContent = JSON.stringify(payload, null, 2);
  var file = backupFolder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT);

  return {
    fileId: file.getId(),
    fileName: fileName,
    fileUrl: file.getUrl()
  };
}

/**
 * Mendapatkan atau membuat subfolder di dalam folder induk
 */
function getOrCreateSubfolder(parentFolder, subfolderName) {
  var folders = parentFolder.getFoldersByName(subfolderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(subfolderName);
  }
}

// =============================================================================
// 7. PENGATURAN FITUR & KREDENSIAL PASSWORD ADMIN
// =============================================================================
function getDefaultAppConfig() {
  return {
    admin: {
      username: CONFIG.DEFAULT_ADMIN_USER || "admin",
      password: CONFIG.DEFAULT_ADMIN_PASS || "admin123",
      requirePasswordOnRoleSwitch: true,
      lastUpdated: new Date().toISOString()
    },
    exam: {
      schoolName: "SMAN 1 SEBATIK",
      examTitle: "Penilaian Sumatif Komprehensif Semester 2026",
      bannerAnnouncement: "UJIAN AKTIF: Penilaian Sumatif Bersama 2026 • Real-Time Spreadsheet Sync Aktif",
      durationMinutes: 60,
      kkm: 75,
      shuffleQuestions: false,
      shuffleOptions: false,
      showScoreAfterExam: true,
      showDiscussionAfterExam: false,
      allowDoubtful: true,
      audioListeningMaxPlays: 2
    },
    security: {
      enforceFullscreen: true,
      detectTabSwitch: true,
      detectSplitScreen: true,
      blockContextMenu: true,
      blockCopyPaste: true,
      blockDevTools: true,
      maxViolations: 4,
      enableBeepAlarm: true
    },
    sync: {
      spreadsheetId: SPREADSHEET_ID,
      targetFolderId: TARGET_FOLDER_ID,
      staggeredDelayMin: 1,
      staggeredDelayMax: 4,
      enableLocalBackup: true
    }
  };
}

function getAppConfigGAS() {
  try {
    var folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    var fileName = CONFIG_APP_FILENAME;
    var files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      var file = files.next();
      var parsed = JSON.parse(file.getBlob().getDataAsString());
      return { status: "success", config: parsed, source: "drive" };
    } else {
      var defaultConfig = getDefaultAppConfig();
      folder.createFile(fileName, JSON.stringify(defaultConfig, null, 2), MimeType.PLAIN_TEXT);
      return { status: "success", config: defaultConfig, source: "default_created" };
    }
  } catch (e) {
    Logger.log("Error getAppConfigGAS: " + e.message);
    return { status: "fallback", config: getDefaultAppConfig(), error: e.message };
  }
}

function saveAppConfigGAS(newConfig) {
  if (!newConfig) return { status: "error", message: "Data konfigurasi kosong" };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { status: "busy", message: "Server sibuk" }; }

  try {
    var folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    var files = folder.getFilesByName(CONFIG_APP_FILENAME);
    var file = files.hasNext() ? files.next() : folder.createFile(CONFIG_APP_FILENAME, "{}", MimeType.PLAIN_TEXT);
    newConfig.lastSavedAt = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    file.setContent(JSON.stringify(newConfig, null, 2));
    lock.releaseLock();
    return { status: "success", message: "Pengaturan berhasil disimpan ke Google Drive", config: newConfig };
  } catch (err) {
    lock.releaseLock();
    return { status: "error", message: "Gagal menyimpan konfigurasi: " + err.message };
  }
}

function verifyAdminLoginGAS(username, password) {
  var confRes = getAppConfigGAS();
  var conf = confRes.config || getDefaultAppConfig();
  var expectedUser = (conf.admin && conf.admin.username) ? conf.admin.username : "admin";
  var expectedPass = (conf.admin && conf.admin.password) ? conf.admin.password : "admin123";

  if (String(username).trim() === expectedUser && String(password) === expectedPass) {
    return { status: "success", authenticated: true, message: "Autentikasi Admin Berhasil", username: expectedUser };
  } else {
    return { status: "error", authenticated: false, message: "Username atau Password Admin salah!" };
  }
}

function changeAdminPasswordGAS(oldPassword, newPassword) {
  var confRes = getAppConfigGAS();
  var conf = confRes.config || getDefaultAppConfig();
  var currentPass = (conf.admin && conf.admin.password) ? conf.admin.password : "admin123";

  if (String(oldPassword) !== currentPass) {
    return { status: "error", message: "Password lama tidak sesuai!" };
  }
  if (!newPassword || newPassword.length < 4) {
    return { status: "error", message: "Password baru minimal 4 karakter!" };
  }

  if (!conf.admin) conf.admin = {};
  conf.admin.password = String(newPassword);
  conf.admin.lastPasswordChangedAt = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");

  return saveAppConfigGAS(conf);
}

function resetStudentLoginGAS(nisn) {
  if (!nisn) return { status: "error", message: "NISN tidak boleh kosong" };
  try {
    var folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    var files = folder.getFilesByName(STATUS_LOGIN_FILENAME);
    var statusMap = {};
    var file = files.hasNext() ? files.next() : folder.createFile(STATUS_LOGIN_FILENAME, "{}", MimeType.PLAIN_TEXT);
    try { statusMap = JSON.parse(file.getBlob().getDataAsString()); } catch (e) { statusMap = {}; }

    statusMap[String(nisn).trim()] = {
      isLocked: false,
      lastResetAt: Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss"),
      resetBy: "Proctor / Admin CBT"
    };

    file.setContent(JSON.stringify(statusMap, null, 2));
    return { status: "success", message: "Status login NISN " + nisn + " berhasil di-reset (Unlocked).", nisn: nisn };
  } catch (err) {
    return { status: "error", message: "Gagal reset status login: " + err.message };
  }
}

// =============================================================================
// 8. HTTP ENDPOINT (doGet & doPost)
// =============================================================================
/**
 * Melayani UI Web App (HTMLService) atau REST API GET
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action).toLowerCase() : "serve_ui";

  if (action === "ping") {
    return createJsonResponse({
      status: "online",
      system: "GusthyPalinPatandaExam Engine",
      version: "2.7.0-Enterprise",
      spreadsheetId: SPREADSHEET_ID,
      folderId: TARGET_FOLDER_ID,
      timestamp: new Date().toISOString()
    });
  }

  if (action === "setup_db" || action === "init") {
    return createJsonResponse(setupDatabase());
  }

  if (action === "ambil_data" || action === "get_data") {
    return createJsonResponse(ambilDataDariSheet());
  }

  if (action === "get_config") {
    return createJsonResponse(getAppConfigGAS());
  }

  if (action === "verify_admin") {
    var u = (e.parameter && e.parameter.username) ? e.parameter.username : "";
    var p = (e.parameter && e.parameter.password) ? e.parameter.password : "";
    return createJsonResponse(verifyAdminLoginGAS(u, p));
  }

  if (action === "reset_login") {
    var nisn = (e.parameter && e.parameter.nisn) ? e.parameter.nisn : "";
    return createJsonResponse(resetStudentLoginGAS(nisn));
  }

  // Melayani file index.html jika diakses via URL Web App
  try {
    var template = HtmlService.createTemplateFromFile("index");
    return template.evaluate()
      .setTitle("GusthyPalinPatandaExam - CBT Anti-Kecurangan 2026")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (errHtml) {
    try {
      var templateAlt = HtmlService.createTemplateFromFile("Index");
      return templateAlt.evaluate()
        .setTitle("GusthyPalinPatandaExam - CBT Anti-Kecurangan 2026")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (errFallback) {
      return createJsonResponse({
        status: "ready",
        message: "GusthyPalinPatandaExam GAS Backend Aktif.",
        spreadsheetId: SPREADSHEET_ID,
        folderId: TARGET_FOLDER_ID,
        serverTime: new Date().toISOString()
      });
    }
  }
}

/**
 * Melayani HTTP POST Request (REST API & CORS Fetch)
 */
function doPost(e) {
  try {
    var payload;
    var rawData = e && e.postData ? e.postData.contents : "";
    if (rawData) {
      try { payload = JSON.parse(rawData); } catch (e) { payload = e.parameter; }
    } else {
      payload = (e && e.parameter) ? e.parameter : {};
    }

    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch (e) {}
    }

    var action = payload.action || (e && e.parameter ? e.parameter.action : null);

    if (action === "ping") {
      return createJsonResponse({ status: "online", spreadsheetId: SPREADSHEET_ID });
    }
    if (action === "setup_db") {
      return createJsonResponse(setupDatabase());
    }
    if (action === "ambil_data" || action === "ambilDataDariSheet") {
      return createJsonResponse(ambilDataDariSheet());
    }
    if (action === "simpanHasilUjian" || action === "submit_exam") {
      return createJsonResponse(simpanHasilUjian(payload));
    }
    if (action === "simpanDataSiswa" || action === "sync_students") {
      return createJsonResponse(simpanDataSiswa(payload));
    }
    if (action === "simpanBankSoal" || action === "sync_questions") {
      return createJsonResponse(simpanBankSoal(payload));
    }
    if (action === "get_config") {
      return createJsonResponse(getAppConfigGAS());
    }
    if (action === "save_config") {
      return createJsonResponse(saveAppConfigGAS(payload.config || payload));
    }
    if (action === "verify_admin") {
      return createJsonResponse(verifyAdminLoginGAS(payload.username, payload.password));
    }
    if (action === "change_password") {
      return createJsonResponse(changeAdminPasswordGAS(payload.oldPassword, payload.newPassword));
    }
    if (action === "reset_login") {
      return createJsonResponse(resetStudentLoginGAS(payload.nisn));
    }

    // Default POST: simpan hasil ujian
    return createJsonResponse(simpanHasilUjian(payload));

  } catch (error) {
    Logger.log("Error doPost: " + error.toString());
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// 9. MENU KUSTOM SPREADSHEET (onOpen)
// =============================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🎓 Gusthy CBT Exam")
    .addItem("🚀 Inisialisasi Database (setupDatabase)", "menuSetupDatabase")
    .addItem("🔄 Tarik Data Terkini (ambilDataDariSheet)", "menuAmbilData")
    .addSeparator()
    .addItem("🧪 Uji Koneksi Folder Google Drive", "menuTestConnection")
    .addItem("🔓 Buka Kunci Siswa (Reset Login NISN)", "menuPromptResetStudent")
    .addItem("🔑 Ubah Password Login Admin", "menuChangeAdminPassword")
    .addToUi();
}

function menuSetupDatabase() {
  var res = setupDatabase();
  SpreadsheetApp.getUi().alert("Setup Selesai:\n\n" + res.message + "\nID Spreadsheet: " + res.spreadsheetId);
}

function menuAmbilData() {
  var res = ambilDataDariSheet();
  SpreadsheetApp.getUi().alert("Tarik Data Berhasil:\n\nSiswa: " + res.studentsCount + "\nSoal: " + res.questionsCount + "\nHasil Ujian: " + res.examResultsCount);
}

function menuTestConnection() {
  try {
    var folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    SpreadsheetApp.getUi().alert("Koneksi Google Drive Berhasil!\n\nNama Folder: " + folder.getName() + "\nID Folder: " + folder.getId());
  } catch (e) {
    SpreadsheetApp.getUi().alert("Gagal koneksi Google Drive: " + e.message);
  }
}

function menuChangeAdminPassword() {
  var ui = SpreadsheetApp.getUi();
  var p1 = ui.prompt("Ubah Password Admin", "Masukkan Password Lama:", ui.ButtonSet.OK_CANCEL);
  if (p1.getSelectedButton() === ui.Button.OK) {
    var oldP = p1.getResponseText();
    var p2 = ui.prompt("Ubah Password Admin", "Masukkan Password Baru (Min 4 karakter):", ui.ButtonSet.OK_CANCEL);
    if (p2.getSelectedButton() === ui.Button.OK) {
      var newP = p2.getResponseText();
      var res = changeAdminPasswordGAS(oldP, newP);
      ui.alert(res.message);
    }
  }
}

function menuPromptResetStudent() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt("Reset Login Siswa Terkunci", "Masukkan NISN siswa yang ingin dibuka kuncinya:", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() === ui.Button.OK) {
    var nisn = prompt.getResponseText().trim();
    if (nisn) {
      var res = resetStudentLoginGAS(nisn);
      ui.alert(res.message);
    }
  }
}
