// =========================
// FILE: Code.gs - Sistem Izin Siswa PAMIT NGGIH
// Versi: 4.5.5 - FINAL FIX PDF + JAM KE 1-10
// =========================

var CONFIG = {
  DEVICE_ID: "d1fb0e3bf18cbaedc9ee2a656c33fd6a",
  TEMPLATE_DOC_ID: "1QL3-MZ74vreJpCy6A8JEFVnE1xWU04M_-aQfNl4jILs",
  FOLDER_ID: "1_yDRDqwQS5OhpaUpz1zOxNwmeNzCp5xP",

  SHEET_NAME: "Data Izin Siswa",
  DATA_SISWA_SHEET_NAME: "Data Siswa",
  GURU_PIKET_SHEET_NAME: "Guru Piket Beta",
  JAM_PELAJARAN_SHEET_NAME: "Jam Pelajaran",

  // ================= CONFIGURASI DATABASE UTAMA =================
  MASTER_DB_ID: "180hXN2MVeAORQrl_QMc9KrcUTSqVoYVl1KWgybJ5gQ8",
  
  // Nomor WA Admin untuk menerima Laporan Rekap Mingguan
  ADMIN_WA: "6285183218979", // GANTI DENGAN NOMOR ADMIN YANG SESUNGGUHNYA
  // ==========================================================

  ADMIN_PASSWORD: "admin123",
  MAX_STUDENTS: 36,
  WACENTER_API_KEY: "f4fff5e69fb130bd0dbf653c7ad4fc27",

  TZ: "Asia/Jakarta",
  SUBMIT_START: "07:00",
  SUBMIT_END: "17:00",

  // ==========================================
  // PEMETAKAN SPREADSHEET PER KELAS
  // ==========================================
  SPREADSHEET_MAP: {
    "X PPLG": "1R2si1qZmYkHDwfi9RFrbq1c2Ib-cKlGjH-8dGkj-Q6g",
    "X PM 1": "1z_SZPucE8FUAN8mICDvIDuquriHOhwHBRm1NBWbSPEM",
    "X PM 2": "1hwOMF41kmHyDcy5aJMW-s8JYECvgTRwc0ukUm1CgzBo",
    "X PM 3": "12QJmM5unIuYAGiW0zykp_e5h1l4iZ13yrIjlJ5EKP5M",
    "X MPLB 1": "1jKqIHZsXkWHaJW2fZ5ymX63yO6j-74m4mzPb4BRR0bw",
    "X MPLB 2": "1jXPRjhXqW5y8kxPkdAbhj6VodpmGEgf94FNTsFl_R5I",
    "X MPLB 3": "1MQVCCemxtQ45oC8pz1l84OrEWxX5geY-_5MN5SOUo7E",
    "X AKL 1": "1_-BFCbmD_UyetIC6IhMSPxr2ve_lXuxJ5jKYqDcjj7E",
    "X AKL 2": "1r8zQ0_21Q4uj3DSznen3Jia7FbmxRdzNmh492OfwQ_4",
    "X AKL 3": "16DOs55YHGmTLdS0PX5Bf9-xrsRzUPZZ3UkT7a06WUAw",

    "XI PPLG": "1jP8AaK8jx07vmPwbdxzbKCbZjhfOFciSrF3v_UAEeTo",
    "XI PM 1": "16MXzi7QTwpqTLzwhMQHi7sLcBns0W_LtpdIWL6vt4CY",
    "XI PM 2": "1TFtKw5q7u2px6ER_laTh5Ko6zgaLbkkA68mVhfgTotw",
    "XI PM 3": "1jdnr3QmVKKQN0Z2p8RJHeQLEUAoPnMw0lLXZM4q0y6o",
    "XI MPLB 1": "1AKqO9sqWLOnBn-IqHTjI_mj3R5cpS7XqhJBu6gxZ6HI",
    "XI MPLB 2": "1v42xDkpaHzQ6gKudPRCXQvyyYGlu_OKp_1u1f7zolu8",
    "XI MPLB 3": "17Sv1XCfidGplBWWFt36aoO5u-rcrWNvQmSe9kDl6ky4",
    "XI AKL 1": "1fFIiuqPsmUThpAWSfmwhlJlMA7kKyZ_Dp7t9TNMnijA",
    "XI AKL 2": "10GOxZm8wTwHMX9wJnHayU0Xu3-b5YdfCktujEc06ZS0",
    "XI AKL 3": "1JSUl93OmUYevzYOfUA15TTLXChrh8G7FR2GVxjexLrg",

    "XII PPLG": "1qCi6Ewvy6T9SLh6I25kl5ieeU2TE4AFkrsfyYV5fb88",
    "XII PM 1": "1Q1uowgCWRkASRQw_EjnQlGlb1tQ7L_HiGgC78KtokWI",
    "XII PM 2": "1ZVp-DpHv_3cwGqLeoo1j_Y-nCAgB9iNEOu8ASg3sY0k",
    "XII PM 3": "1bJvGvXApk3cxRNzB7FJsy0SR4pYN_z8UKMGVjGmFf-U",
    "XII MPLB 1": "1aGgOcO2UaeLvRyTG7eIuvGRaScnUccStrW_ztZpNzRk",
    "XII MPLB 2": "12EOcekhNX8iLTZQ4K7CEIHF0pz6wMRMkhKFx80HVk1E",
    "XII MPLB 3": "1y96UX4lqZq9hS47HzPoe2SE12VDoZ9rk0OgFX7gv0qY",
    "XII AKL 1": "19Y8ttXNbhGpHOw44Ne9xLytTGbKv260luDnsKSh1QT0",
    "XII AKL 2": "1A063ZdSYXyUqXdVCcwt0EEmik3hPad8iyY6aw4jEGKU",
    "XII AKL 3": "1_sT4cChxb0MReajFN8e5XZLyIjCsdnN3vlGKWcBnOgo"
  },

  // ==========================================
  // JAM PELAJARAN (LANGSUNG DI CONFIG)
  // ==========================================
  JAM_PELAJARAN: {
    "1": { mulai: "07:00", selesai: "07:45" },
    "2": { mulai: "07:45", selesai: "08:30" },
    "3": { mulai: "08:30", selesai: "09:15" },
    "4": { mulai: "09:15", selesai: "10:00" },
    "5": { mulai: "10:00", selesai: "10:45" },
    "6": { mulai: "10:45", selesai: "11:30" },
    "7": { mulai: "12:00", selesai: "12:45" },
    "8": { mulai: "12:45", selesai: "13:30" },
    "9": { mulai: "13:30", selesai: "14:15" },
    "10": { mulai: "14:15", selesai: "15:00" }
  }
};

// =========================
// HTML SERVE
// =========================
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.kode) {
      return HtmlService.createTemplateFromFile("konfirmasi")
        .evaluate()
        .setTitle("Konfirmasi Kepulangan")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (e && e.parameter && e.parameter.stats) {
      return HtmlService.createTemplateFromFile("stats")
        .evaluate()
        .setTitle("Statistik Izin Hari Ini")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (e && e.parameter && e.parameter.cekTerlambat) {
      return ContentService.createTextOutput(JSON.stringify(cekDanKirimNotifikasiTerlambat()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (e && e.parameter && e.parameter.setupTrigger) {
      return ContentService.createTextOutput(JSON.stringify(setupNotifikasiOtomatis()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (e && e.parameter && e.parameter.rekap) {
      return ContentService.createTextOutput(JSON.stringify(runWeeklyRecapToWa_()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (e && e.parameter && e.parameter.debug) {
      return ContentService.createTextOutput(JSON.stringify(debugPDF()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return HtmlService.createTemplateFromFile("index")
      .evaluate()
      .setTitle("Formulir Izin Siswa")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    Logger.log("doGet Error: " + err.message + "\n" + err.stack);
    return HtmlService.createHtmlOutput("<h1>Error</h1><p>" + err.message + "</p>");
  }
}

function include(filename) {
  try {
    if (!filename) return "<!-- include: filename kosong -->";
    return HtmlService.createHtmlOutputFromFile(String(filename)).getContent();
  } catch (e) {
    Logger.log("include Error (" + filename + "): " + e.message);
    return "<!-- include gagal: " + String(filename) + " -->";
  }
}

// =========================
// PING
// =========================
function ping() {
  return {
    ok: true,
    now: Utilities.formatDate(new Date(), CONFIG.TZ, "dd/MM/yyyy, HH:mm:ss"),
    app: "Izin Siswa v4.5.5"
  };
}

// =========================
// FUNGSI UNTUK MENDAPATKAN DROPDOWN JAM KE
// =========================
function getJamKeList() {
  return ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"];
}

// =========================
// KONVERSI JAM KE KE FORMAT WAKTU
// =========================
function konversiJamKeKeWaktu(jamKe) {
  if (!jamKe) return null;
  var jam = CONFIG.JAM_PELAJARAN[jamKe.toString()];
  return jam ? jam.mulai : null;
}

// =========================
// SHEET PICKERS
// =========================
function getMasterSS_() { 
  try { 
    return SpreadsheetApp.openById(CONFIG.MASTER_DB_ID); 
  } catch(e) { 
    Logger.log("Master DB Error: " + e.message); 
    return null; 
  } 
}

function getActiveSS_() { 
  return SpreadsheetApp.getActiveSpreadsheet(); 
}

function getLogSheet_(preferMaster) {
  var sheet = null;

  if (preferMaster) {
    try {
      var mss = getMasterSS_();
      if(mss) {
        sheet = mss.getSheetByName(CONFIG.SHEET_NAME);
        if (sheet) return { ss: mss, sheet: sheet, where: "MASTER" };
      }
    } catch (e) {}
  }

  try {
    var ass = getActiveSS_();
    sheet = ass.getSheetByName(CONFIG.SHEET_NAME);
    if (sheet) return { ss: ass, sheet: sheet, where: "ACTIVE" };
  } catch (e2) {}

  if (!preferMaster) {
    try {
      var mss2 = getMasterSS_();
      if(mss2) {
        sheet = mss2.getSheetByName(CONFIG.SHEET_NAME);
        if (sheet) return { ss: mss2, sheet: sheet, where: "MASTER" };
      }
    } catch (e3) {}
  }

  return { ss: null, sheet: null, where: "NONE" };
}

// =========================
// SHEET SETUP
// =========================
function createSheet(ss) {
  if (!ss) throw new Error("createSheet: Spreadsheet kosong");
  var sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sh) return sh;

  sh = ss.insertSheet(CONFIG.SHEET_NAME);

  var headers = [
    "Timestamp",
    "Keperluan",
    "Tanggal",
    "Jam Awal",
    "Jam Akhir",
    "Nama Guru",
    "Kode Konfirmasi",
    "Status",
    "Waktu Konfirmasi",
    "Jumlah Siswa",
    "Tingkat",
    "Kelas"
  ];

  for (var i = 1; i <= CONFIG.MAX_STUDENTS; i++) headers.push("Nama " + i);
  for (var j = 1; j <= CONFIG.MAX_STUDENTS; j++) headers.push("Konfirmasi " + j);
  for (var k = 1; k <= CONFIG.MAX_STUDENTS; k++) headers.push("Waktu Konfirmasi " + k);

  sh.appendRow(headers);
  sh.setFrozenRows(1);
  return sh;
}

function initLogSheet() {
  var pick = getLogSheet_(true);
  if (pick.sheet) return { success: true, sheet: CONFIG.SHEET_NAME, where: pick.where };

  var ss = getActiveSS_();
  createSheet(ss);
  return { success: true, sheet: CONFIG.SHEET_NAME, where: "ACTIVE(created)" };
}

function saveToSpreadsheet(sheet, data) {
  try {
    if (!Array.isArray(data)) throw new Error("Data row tidak valid (bukan array)");

    if (!sheet || typeof sheet.getLastRow !== "function") {
      var pick = getLogSheet_(true);
      if (pick.sheet) sheet = pick.sheet;
      if (!sheet) sheet = createSheet(getActiveSS_());
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 1) lastRow = 1;

    sheet.getRange(lastRow + 1, 1, 1, data.length).setValues([data]);

    sheet.getRange(lastRow + 1, 1, 1, data.length)
      .setBackground("#f0f9ff")
      .setBorder(true, true, true, true, true, true);

    return true;
  } catch (error) {
    Logger.log("saveToSpreadsheet Error: " + error.message + "\nStack: " + error.stack);
    throw new Error("Gagal menyimpan ke spreadsheet: " + error.message);
  }
}

// =========================
// HELPERS (umum)
// =========================
function getWebAppUrl_() {
  var url = "";
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  if (url) return url;
  try {
    var sp = PropertiesService.getScriptProperties();
    return sp.getProperty("WEB_APP_URL") || "";
  } catch (e2) { return ""; }
}

function normalizeHeader_(h) { 
  return (h || "").toString().toLowerCase().trim().replace(/\s+/g, " "); 
}

function findHeaderIndex_(headers, candidates) {
  var norm = headers.map(normalizeHeader_);
  for (var i = 0; i < candidates.length; i++) {
    var idx = norm.indexOf(normalizeHeader_(candidates[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

function toWa62_(raw) {
  var s = (raw || "").toString().trim();
  if (!s) return "";
  s = s.replace(/[^\d]/g, "");
  if (!s) return "";
  if (s.startsWith("0")) s = "62" + s.substring(1);
  if (s.startsWith("620")) s = "62" + s.substring(3);
  if (!/^62\d{8,15}$/.test(s)) return "";
  return s;
}

function getHariIndonesia_() {
  var hari = ["minggu", "senin", "selasa", "rabu", "kamis", "jumat", "sabtu"];
  var now = new Date();
  var wibDay = parseInt(Utilities.formatDate(now, CONFIG.TZ, "u"), 10); 
  var map = { 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 0 };
  return hari[map[wibDay]];
}

function getHariIndonesiaFromYmd_(ymd) {
  try {
    var s = String(ymd || "").trim();
    if (!s) return "";
    var parts = s.split("-");
    var d = (parts.length === 3) ? new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])) : new Date(s);
    if (isNaN(d.getTime())) return "";
    var u = parseInt(Utilities.formatDate(d, CONFIG.TZ, "u"), 10);
    var map = { 1: "senin", 2: "selasa", 3: "rabu", 4: "kamis", 5: "jumat", 6: "sabtu", 7: "minggu" };
    return map[u] || "";
  } catch (e) {
    return "";
  }
}

function generateConfirmationCode() {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var code = "";
  for (var i = 0; i < 6; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
  return code;
}

function safeFileName_(s) {
  return (s || "")
    .toString()
    .replace(/[\\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .substring(0, 60);
}

function escapeRegexForDoc_(s) { 
  return (s || "").toString().replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); 
}

function replaceLineByLabel_(body, label, value) {
  var val = (value == null) ? "" : value.toString();
  var pattern = "^\\s*" + escapeRegexForDoc_(label) + "\\s*:\\s*.*$";
  var replacement = label + " : " + val;
  try { if (body.findText(pattern)) body.replaceText(pattern, replacement); } catch (e) {}
}

function replaceExactText_(body, fromText, toText) {
  try {
    var patt = escapeRegexForDoc_(fromText);
    if (body.findText(patt)) body.replaceText(patt, toText);
  } catch (e) {}
}

function replaceAnglePlaceholderLoose_(body, key, value) {
  try {
    var val = (value == null) ? "" : String(value);
    var patt = "<\\s*" + escapeRegexForDoc_(key) + "\\s*>";
    if (body.findText(patt)) body.replaceText(patt, val);
  } catch (e) {}
}

function setParagraphTextUnbold_(paragraph, newText) {
  try {
    paragraph.setText(String(newText || ""));
    var t = paragraph.editAsText();
    if (t) t.setBold(false);
    return true;
  } catch (e) {
    return false;
  }
}

function replaceParagraphContainsUnbold_(body, containsText, newFullText) {
  try {
    var paras = body.getParagraphs();
    var needle = String(containsText || "");
    for (var i = 0; i < paras.length; i++) {
      var t = paras[i].getText();
      if (t && t.indexOf(needle) !== -1) {
        return setParagraphTextUnbold_(paras[i], newFullText);
      }
    }
  } catch (e) {}
  return false;
}

// =========================
// FORMAT TANGGAL & HARI
// =========================
function formatTanggalSurat_(dateString) {
  try {
    var s = String(dateString || "").trim();
    if (!s) return "";

    var d;
    var parts = s.split("-");
    if (parts.length === 3) d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    else d = new Date(s);
    if (isNaN(d.getTime())) return s;

    var bulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
    var ddStr = Utilities.formatDate(d, CONFIG.TZ, "dd");
    var mm = parseInt(Utilities.formatDate(d, CONFIG.TZ, "M"), 10) - 1;
    var yy = Utilities.formatDate(d, CONFIG.TZ, "yyyy");

    return ddStr + " " + bulan[mm] + " " + yy;
  } catch (e) {
    return String(dateString || "");
  }
}

function formatHariOnly_(dateString) {
  try {
    var s = String(dateString || "").trim();
    if (!s) return "";

    var d;
    var parts = s.split("-");
    if (parts.length === 3) d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    else d = new Date(s);
    if (isNaN(d.getTime())) return s;

    var hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"];
    var u = parseInt(Utilities.formatDate(d, CONFIG.TZ, "u"), 10);
    if (u >= 1 && u <= 5) return hari[u - 1];
    return "";
  } catch (e) {
    return String(dateString || "");
  }
}

function formatHariTanggal_(dateString) {
  try {
    var s = String(dateString || "").trim();
    if (!s) return "";

    var d;
    var parts = s.split("-");
    if (parts.length === 3) d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    else d = new Date(s);
    if (isNaN(d.getTime())) return s;

    var hari = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
    var bulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];

    var dayIndex = parseInt(Utilities.formatDate(d, CONFIG.TZ, "e"), 10) - 1;
    if (dayIndex < 0 || dayIndex > 6) dayIndex = d.getDay();

    var ddStr = Utilities.formatDate(d, CONFIG.TZ, "dd");
    var mm = parseInt(Utilities.formatDate(d, CONFIG.TZ, "M"), 10) - 1;
    var yy = Utilities.formatDate(d, CONFIG.TZ, "yyyy");

    var h = hari[dayIndex];
    if (h === "Jumat") h = "Jum'at";

    return h + ", " + ddStr + " " + bulan[mm] + " " + yy;
  } catch (e) {
    return String(dateString || "");
  }
}

// =========================
// FUNGSI UNTUK MENDAPATKAN RENTANG JAM KE
// =========================
function getRentangJamKe(jamAwalKe, jamAkhirKe) {
  if (!jamAwalKe || !jamAkhirKe) return "";
  return "Ke-" + jamAwalKe + " sampai dengan Ke-" + jamAkhirKe;
}

// =========================
// PDF GENERATION - VERSI PERBAIKAN
// =========================
function extractDriveId_(input) {
  var s = (input || "").toString().trim();
  if (!s) return "";
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s) && s.indexOf("http") !== 0) return s;
  var m = s.match(/\/d\/([-\w]{25,})/) || s.match(/[-\w]{25,}/);
  if (!m) return "";
  return (m[1] || m[0] || "").toString();
}

function getFileByIdSafe_(idOrUrl, label) {
  var id = extractDriveId_(idOrUrl);
  if (!id) throw new Error(label + " tidak valid (ID kosong / URL tidak terbaca)");
  try {
    var f = DriveApp.getFileById(id);
    f.getName();
    return f;
  } catch (e) {
    throw new Error(label + " tidak bisa diakses / tidak ditemukan. ID: " + id);
  }
}

function getFolderByIdSafe_(idOrUrl, label) {
  var id = extractDriveId_(idOrUrl);
  if (!id) throw new Error(label + " tidak valid (ID kosong / URL tidak terbaca)");
  try {
    var f = DriveApp.getFolderById(id);
    f.getName();
    return f;
  } catch (e) {
    throw new Error(label + " tidak bisa diakses / tidak ditemukan. ID: " + id);
  }
}

// =========================
// DEBUG PDF
// =========================
function debugPDF() {
  var result = {
    success: false,
    steps: [],
    error: null
  };
  
  try {
    result.steps.push("Memulai debug PDF...");
    
    // Cek akses ke template
    var templateId = CONFIG.TEMPLATE_DOC_ID;
    result.steps.push("Template ID: " + templateId);
    
    try {
      var template = DriveApp.getFileById(templateId);
      result.steps.push("✅ Template ditemukan: " + template.getName());
    } catch (e) {
      result.steps.push("❌ ERROR: Template tidak bisa diakses: " + e.message);
      result.error = e.message;
      return result;
    }
    
    // Cek akses ke folder
    var folderId = CONFIG.FOLDER_ID;
    result.steps.push("Folder ID: " + folderId);
    
    try {
      var folder = DriveApp.getFolderById(folderId);
      result.steps.push("✅ Folder ditemukan: " + folder.getName());
    } catch (e) {
      result.steps.push("⚠️ Folder tidak bisa diakses, akan gunakan root folder");
      folder = DriveApp.getRootFolder();
    }
    
    // Coba buat dokumen test
    try {
      var testDoc = template.makeCopy("TEST-IZIN-" + new Date().getTime(), folder);
      var docId = testDoc.getId();
      result.steps.push("✅ Berhasil membuat copy dokumen: " + docId);
      
      // Coba konversi ke PDF
      var pdfBlob = DriveApp.getFileById(docId).getBlob().getAs("application/pdf");
      result.steps.push("✅ Berhasil konversi ke PDF, ukuran: " + pdfBlob.getBytes().length + " bytes");
      
      // Hapus file test
      DriveApp.getFileById(docId).setTrashed(true);
      result.steps.push("✅ File test dihapus");
      
      result.success = true;
      
    } catch (e) {
      result.steps.push("❌ ERROR saat membuat/mengkonversi dokumen: " + e.message);
      result.error = e.message;
    }
    
  } catch (e) {
    result.error = e.message;
  }
  
  return result;
}

// =========================
// DOKUMEN + PDF - VERSI PERBAIKAN
// =========================
function generateDocAndPdf_(formData, studentData, kodeKonfirmasi) {
  try {
    formData = formData || {};
    studentData = Array.isArray(studentData) ? studentData : [];

    console.log("=== MEMULAI GENERATE PDF ===");
    
    // ===== CEK TEMPLATE =====
    var templateId = CONFIG.TEMPLATE_DOC_ID;
    console.log("Template ID: " + templateId);
    
    var templateFile;
    try {
      templateFile = DriveApp.getFileById(templateId);
      console.log("✅ Template ditemukan: " + templateFile.getName());
    } catch (e) {
      throw new Error("Template tidak ditemukan. Pastikan TEMPLATE_DOC_ID benar dan file bisa diakses. Error: " + e.message);
    }

    // ===== CEK FOLDER =====
    var folderId = CONFIG.FOLDER_ID;
    console.log("Folder ID: " + folderId);
    
    var folder;
    try {
      folder = DriveApp.getFolderById(folderId);
      console.log("✅ Folder ditemukan: " + folder.getName());
    } catch (e) {
      console.log("⚠️ Folder tidak bisa diakses, akan gunakan root folder");
      folder = DriveApp.getRootFolder();
    }

    // ===== SIAPKAN DATA =====
    var s1 = studentData[0] || {};
    var vKelas = String(s1.tingkat || "").trim();   
    var vJurusan = String(s1.jurusan || "").trim(); 

    var namaList = studentData
      .map(function (s) { return (s && s.nama) ? String(s.nama).trim() : ""; })
      .filter(function(n) { return n !== ""; });
    var vNamaGabung = namaList.join(", ");
    
    console.log("Nama siswa: " + vNamaGabung);
    console.log("Kelas: " + vKelas + " " + vJurusan);

    // ===== BUAT COPY DOKUMEN =====
    var timestamp = Utilities.formatDate(new Date(), CONFIG.TZ, "yyyyMMdd-HHmmss");
    var baseName = "Izin_" + 
      (s1.nama ? s1.nama.replace(/[^a-zA-Z0-9]/g, "_").substring(0, 30) : "Siswa") + 
      "_" + timestamp;
    
    console.log("Nama file: " + baseName);
    
    var docCopy;
    try {
      docCopy = templateFile.makeCopy(baseName, folder);
      console.log("✅ Berhasil membuat copy dokumen: " + docCopy.getId());
    } catch (e) {
      throw new Error("Gagal membuat copy template: " + e.message);
    }

    // ===== BUKA DOKUMEN =====
    var doc;
    try {
      doc = DocumentApp.openById(docCopy.getId());
      console.log("✅ Berhasil membuka dokumen");
    } catch (e) {
      throw new Error("Template bukan Google Docs atau tidak bisa dibuka: " + e.message);
    }

    // ===== EDIT DOKUMEN =====
    var body = doc.getBody();

    var vTanggalSurat = formatTanggalSurat_(formData.tanggal || "");
    var vHariOnly = formatHariOnly_(formData.tanggal || "");
    var jamAwalKe = formData.jamAwal || "";
    var jamAkhirKe = formData.jamAkhir || "";
    var vJamUntukSurat = getRentangJamKe(jamAwalKe, jamAkhirKe);
    var vKeperluan = String(formData.keperluan || "").trim();
    var vKelasProdi = (vKelas && vJurusan) ? (vKelas + " " + vJurusan) : (vKelas || vJurusan || "");

    console.log("Data untuk template:");
    console.log("- Nama: " + vNamaGabung);
    console.log("- Kelas: " + vKelasProdi);
    console.log("- Hari: " + vHariOnly);
    console.log("- Tanggal: " + vTanggalSurat);
    console.log("- Jam: " + vJamUntukSurat);
    console.log("- Keperluan: " + vKeperluan);

    // Ganti placeholder dengan berbagai metode
    try {
      // Metode 1: Replace text biasa
      body.replaceText("<nama>", vNamaGabung);
      body.replaceText("<kelasprogramkeahlian>", vKelasProdi);
      body.replaceText("<padahariini>", vHariOnly);
      body.replaceText("<tanggal>", vTanggalSurat);
      body.replaceText("<jampelajaran>", vJamUntukSurat);
      body.replaceText("<keperluan>", vKeperluan);
      
      // Metode 2: Dengan spasi
      body.replaceText("< nama >", vNamaGabung);
      body.replaceText("< kelasprogramkeahlian >", vKelasProdi);
      body.replaceText("< padahariini >", vHariOnly);
      body.replaceText("< tanggal >", vTanggalSurat);
      body.replaceText("< jampelajaran >", vJamUntukSurat);
      body.replaceText("< keperluan >", vKeperluan);
      
      // Metode 3: Tanpa kurung siku (langsung teks)
      body.replaceText("Nama :", "Nama : " + vNamaGabung);
      body.replaceText("Kelas Program Keahlian :", "Kelas Program Keahlian : " + vKelasProdi);
      body.replaceText("Pada Hari Ini :", "Pada Hari Ini : " + vHariOnly);
      body.replaceText("Tanggal :", "Tanggal : " + vTanggalSurat);
      body.replaceText("Jam Pelajaran :", "Jam Pelajaran : " + vJamUntukSurat);
      body.replaceText("Keperluan :", "Keperluan : " + vKeperluan);
      
      console.log("✅ Berhasil mengganti placeholder");
    } catch (e) {
      console.log("⚠️ Error saat mengganti placeholder: " + e.message);
    }

    doc.saveAndClose();
    console.log("✅ Dokumen disimpan");

    // ===== SET SHARING =====
    try {
      docCopy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      console.log("⚠️ Gagal set sharing: " + e.message);
    }

    var docUrl = "https://docs.google.com/document/d/" + docCopy.getId() + "/edit";

    // ===== KONVERSI KE PDF =====
    console.log("Mengkonversi ke PDF...");
    
    var pdfBlob;
    try {
      pdfBlob = DriveApp.getFileById(docCopy.getId()).getBlob().getAs("application/pdf");
      pdfBlob.setName(baseName + ".pdf");
      console.log("✅ Berhasil konversi ke PDF, ukuran: " + pdfBlob.getBytes().length + " bytes");
    } catch (e) {
      throw new Error("Gagal konversi ke PDF: " + e.message);
    }

    // ===== SIMPAN PDF =====
    var pdfFile;
    try {
      pdfFile = folder.createFile(pdfBlob);
      console.log("✅ PDF disimpan di folder: " + pdfFile.getName());
    } catch (e) {
      console.log("⚠️ Gagal simpan di folder, coba root: " + e.message);
      pdfFile = DriveApp.getRootFolder().createFile(pdfBlob);
      console.log("✅ PDF disimpan di root");
    }

    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    var pdfUrl = "https://drive.google.com/uc?id=" + pdfFile.getId();
    
    console.log("✅ PDF URL: " + pdfUrl);
    console.log("=== GENERATE PDF SELESAI ===");

    return { 
      docId: docCopy.getId(), 
      docUrl: docUrl, 
      pdfId: pdfFile.getId(), 
      pdfUrl: pdfUrl 
    };
    
  } catch (error) {
    console.log("❌ ERROR generateDocAndPdf_: " + error.message);
    console.log(error.stack);
    throw new Error("Gagal membuat dokumen/PDF: " + error.message);
  }
}

// =========================
// GURU PIKET
// =========================
function getGuruPiketList() {
  try {
    var ss = getMasterSS_();
    if (!ss) throw new Error("Master Database tidak bisa diakses");

    var sheet = ss.getSheetByName(CONFIG.GURU_PIKET_SHEET_NAME);
    if (!sheet) throw new Error("Sheet '" + CONFIG.GURU_PIKET_SHEET_NAME + "' not found");

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) throw new Error("Data guru piket kosong");

    var headers = data[0].map(function (h) { return (h || "").toString(); });

    var hariIndex = findHeaderIndex_(headers, ["HARI", "Hari", "hari"]);
    if (hariIndex === -1) throw new Error("Header 'HARI' tidak ditemukan");

    var guruIdx = [];
    var phoneIdx = [];

    for (var n = 1; n <= 9; n++) {
      var gi = findHeaderIndex_(headers, ["NAMA GURU " + n, "Nama Guru " + n, "GURU " + n, "Guru " + n]);
      var pi = findHeaderIndex_(headers, ["TELEPON GURU " + n, "Telepon Guru " + n, "NO HP GURU " + n, "No HP Guru " + n, "WA GURU " + n, "Wa Guru " + n]);
      guruIdx.push(gi);
      phoneIdx.push(pi);
    }

    var missing = [];
    for (var c = 1; c <= 9; c++) {
      if (guruIdx[c - 1] === -1) missing.push("NAMA GURU " + c);
      if (phoneIdx[c - 1] === -1) missing.push("TELEPON GURU " + c);
    }
    if (missing.length) throw new Error("Header guru piket tidak lengkap: " + missing.join(", "));

    var today = getHariIndonesia_();

    for (var i = 1; i < data.length; i++) {
      var hari = (data[i][hariIndex] || "").toString().trim().toLowerCase();
      if (hari !== today) continue;

      var gurus = [];
      for (var k = 0; k < 9; k++) {
        var nama = (data[i][guruIdx[k]] || "").toString().trim();
        var phoneRaw = (data[i][phoneIdx[k]] || "").toString().trim();

        if (!nama || nama === "-") continue;
        if (!phoneRaw || phoneRaw === "-" || phoneRaw.toUpperCase() === "FALSE") continue;

        var phone = toWa62_(phoneRaw);
        if (!phone) continue;

        gurus.push({ nama: nama, phone: phone });
      }

      return gurus;
    }

    return [];
  } catch (e) {
    Logger.log("getGuruPiketList Error: " + e.message);
    return [];
  }
}

function getAllGuruPiket_() {
  try {
    var ss = getMasterSS_();
    if (!ss) return [];

    var sheet = ss.getSheetByName(CONFIG.GURU_PIKET_SHEET_NAME);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return [];

    var guruMap = {};

    for (var i = 1; i < data.length; i++) {
      for (var n = 0; n < 10; n++) {
        var idxNama = 1 + (n * 2);
        var idxWa = 2 + (n * 2);

        var nama = (data[i][idxNama] || "").toString().trim();
        var phoneRaw = (data[i][idxWa] || "").toString().trim();

        if (!nama || nama === "-") continue;
        if (!phoneRaw || phoneRaw === "-") continue;

        var phone = toWa62_(phoneRaw);
        if (!phone) continue;

        if (!guruMap[nama]) {
          guruMap[nama] = {
            nama: nama,
            phone: phone
          };
        }
      }
    }

    var guruList = [];
    for (var key in guruMap) {
      guruList.push(guruMap[key]);
    }

    return guruList;

  } catch (e) {
    Logger.log("getAllGuruPiket_ Error: " + e.message);
    return [];
  }
}

function getGuruPiketPhoneByName_(namaGuru) {
  try {
    var name = (namaGuru || "").toString().trim();
    if (!name) return "";

    var semuaGuru = getAllGuruPiket_();
    for (var i = 0; i < semuaGuru.length; i++) {
      if (semuaGuru[i].nama === name) {
        return semuaGuru[i].phone;
      }
    }
    return "";
  } catch (e) {
    Logger.log("getGuruPiketPhoneByName_ Error: " + e.message);
    return "";
  }
}

// =========================
// DATA KELAS
// =========================
function getKelasList() {
  try {
    var kelasList = [];
    for (var kelas in CONFIG.SPREADSHEET_MAP) {
      if (CONFIG.SPREADSHEET_MAP.hasOwnProperty(kelas)) {
        kelasList.push(kelas);
      }
    }
    kelasList.sort();
    return kelasList;
  } catch (e) {
    Logger.log("getKelasList Error: " + e.message);
    return [];
  }
}

function getAngkatanList() { 
  return ["X", "XI", "XII"]; 
}

function getJurusanList(angkatan) {
  try {
    var fixed = ["AKL 1","AKL 2","AKL 3","MPLB 1","MPLB 2","MPLB 3","PM 1","PM 2","PM 3","PPLG"];
    var ss = getActiveSS_();
    var sh = ss.getSheetByName(CONFIG.DATA_SISWA_SHEET_NAME);
    if (!sh) throw new Error("Sheet '" + CONFIG.DATA_SISWA_SHEET_NAME + "' tidak ditemukan");

    var values = sh.getDataRange().getValues();
    var set = {};
    var aNeed = (angkatan || "").toString().trim().toUpperCase();

    for (var i = 1; i < values.length; i++) {
      var a = (values[i][0] || "").toString().trim().toUpperCase();
      var j = (values[i][1] || "").toString().trim();
      if (!a || !j) continue;
      if (a === aNeed) set[j] = true;
    }

    var found = Object.keys(set);
    found.sort(function (x, y) {
      var ix = fixed.indexOf(x), iy = fixed.indexOf(y);
      if (ix === -1 && iy === -1) return x.localeCompare(y);
      if (ix === -1) return 1;
      if (iy === -1) return -1;
      return ix - iy;
    });

    return found.length ? found : fixed;
  } catch (e) {
    return [];
  }
}

// =========================
// AMBIL NAMA SISWA BERDASARKAN KELAS
// =========================
function getNamaSiswaByKelas(kelasPenuh) {
  try {
    if (!kelasPenuh || kelasPenuh.trim() === "") {
      return [];
    }

    var kelasKey = kelasPenuh.trim();
    var ssId = CONFIG.SPREADSHEET_MAP[kelasKey];

    if (!ssId) {
      Logger.log("Kelas " + kelasKey + " tidak ada di SPREADSHEET_MAP");
      return [];
    }

    var classSS = SpreadsheetApp.openById(ssId);
    if (!classSS) return [];

    var sheet = classSS.getSheetByName("Data Siswa");
    if (!sheet) {
      sheet = classSS.getSheets()[0];
    }

    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return [];

    var namaCol = 1;
    var nisCol = 2;

    var out = [];
    var parts = kelasKey.split(" ");
    var tingkat = parts[0];
    var jurusan = parts.slice(1).join(" ");

    for (var i = 1; i < data.length; i++) {
      var rowNama = (data[i][namaCol] || "").toString().trim();
      var rowNis = (data[i][nisCol] || "").toString().trim();

      if (!rowNama) continue;

      out.push({
        nama: rowNama,
        nis: rowNis,
        tingkat: tingkat,
        jurusan: jurusan
      });
    }

    out.sort(function (a, b) {
      return (a.nama || "").localeCompare((b.nama || ""), "id");
    });

    return out;

  } catch (e) {
    Logger.log("getNamaSiswaByKelas Error: " + e.message);
    return [];
  }
}

function getNamaSiswaList(angkatan, jurusan) {
  try {
    var ss = getActiveSS_();
    var sh = ss.getSheetByName(CONFIG.DATA_SISWA_SHEET_NAME);
    if (!sh) throw new Error("Sheet '" + CONFIG.DATA_SISWA_SHEET_NAME + "' tidak ditemukan");

    var values = sh.getDataRange().getValues();
    var out = [];

    var aNeed = (angkatan || "").toString().trim().toUpperCase();
    var jNeed = (jurusan || "").toString().trim();

    for (var i = 1; i < values.length; i++) {
      var a = (values[i][0] || "").toString().trim().toUpperCase();
      var j = (values[i][1] || "").toString().trim();
      var nis = (values[i][2] || "").toString().trim();
      var nama = (values[i][3] || "").toString().trim();
      if (!a || !j || !nama) continue;
      if (a === aNeed && j === jNeed) out.push({ nama: nama, nis: nis });
    }

    out.sort(function (p, q) { return (p.nama || "").localeCompare((q.nama || ""), "id"); });
    return out;
  } catch (e) {
    return [];
  }
}

// =========================
// WA ORTU (AMBIL DARI FILE PER KELAS)
// =========================
function getNoWaOrtuList_(studentData) {
  studentData = Array.isArray(studentData) ? studentData : [];
  if (studentData.length === 0) return [];

  try {
    var firstSiswa = studentData[0] || {};
    var angkatan = (firstSiswa.tingkat || "").trim();
    var jurusan = (firstSiswa.jurusan || "").trim();
    
    var classKey = angkatan + " " + jurusan;
    var ssId = CONFIG.SPREADSHEET_MAP[classKey];

    if (ssId) {
      var classSS = SpreadsheetApp.openById(ssId);
      if (classSS) {
        var sheet = classSS.getSheets()[0]; 
        var list = getNoWaFromSheet_(sheet, studentData);
        if (list.length > 0) return list;
      }
    }

  } catch (e) {
    Logger.log("Gagal akses Spreadsheet per Kelas untuk WA Ortu: " + e.message);
  }

  try {
    var masterDb = getMasterSS_();
    if(masterDb) {
      var presensiSheet = masterDb.getSheetByName("Presensi");
      if (presensiSheet) {
        var list = getNoWaFromSheet_(presensiSheet, studentData);
        if (list.length > 0) return list;
      }
    }
  } catch (e) {}

  var ss = getActiveSS_();
  var sh = ss.getSheetByName(CONFIG.DATA_SISWA_SHEET_NAME);
  if (sh) {
    return getNoWaFromSheet_(sh, studentData);
  }

  return [];
}

function getNoWaFromSheet_(sheet, studentData) {
  studentData = Array.isArray(studentData) ? studentData : [];
  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  var headers = values[0].map(function (h) { return (h || "").toString(); });

  var nisCol  = findHeaderIndex_(headers, ["nis", "Nis", "NIS"]);
  var namaCol = findHeaderIndex_(headers, ["nama_siswa", "nama siswa", "Nama Siswa", "nama", "Nama"]);
  var waCol   = findHeaderIndex_(headers, ["no wa ortu", "No WA Ortu", "nowa ortu", "wa ortu", "no wa", "nomor wa ortu"]);

  if (nisCol === -1 && namaCol === -1) return [];
  if (waCol === -1) return [];

  var out = [];

  for (var i = 0; i < studentData.length; i++) {
    var s = studentData[i] || {};
    var targetNis = (s.nis || "").toString().trim();
    var targetNama = (s.nama || "").toString().trim();

    var rowIndex = -1;

    if (targetNis && nisCol !== -1) {
      for (var r = 1; r < values.length; r++) {
        if ((values[r][nisCol] || "").toString().trim() === targetNis) { rowIndex = r; break; }
      }
    }

    if (rowIndex === -1 && targetNama && namaCol !== -1) {
      for (var r2 = 1; r2 < values.length; r2++) {
        if ((values[r2][namaCol] || "").toString().trim() === targetNama) { rowIndex = r2; break; }
      }
    }

    if (rowIndex === -1) continue;

    var wa = toWa62_(values[rowIndex][waCol]);
    if (wa) out.push(wa);
  }

  return out;
}

// =========================
// WHATSAPP SENDER
// =========================
function sendWhatsAppMessage_(number, message) {
  try {
    var num = toWa62_(number);
    if (!num) return { success: false, message: "Nomor WhatsApp tidak valid." };

    var payload = {
      device_id: CONFIG.DEVICE_ID,
      number: num,
      message: message,
      api_key: CONFIG.WACENTER_API_KEY || null
    };

    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 30000
    };

    var response = UrlFetchApp.fetch("https://app.whacenter.com/api/send", options);
    var result = JSON.parse(response.getContentText() || "{}");

    if (result && result.status === true) return { success: true };
    var errorMsg = (result && result.message) ? result.message : "Gagal mengirim pesan";
    return { success: false, message: errorMsg };
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

// =========================
// NOTIFIKASI KONFIRMASI KEPULANGAN (HANYA 1 KALI)
// =========================

function sendKonfirmasiNotifToGuru(noHPGuru, namaSiswa, waktuKonfirmasi) {
  var message = 
    "*Notifikasi : Kepulangan Siswa*\n\n" +
    "Siswa atas nama *" + namaSiswa + "* telah melakukan konfirmasi kepulangan dan dinyatakan telah kembali.\n\n" +
    "Waktu Konfirmasi : " + waktuKonfirmasi + " WIB\n\n" +
    "Data kepulangan telah tercatat pada sistem.\n\n" +
    "Hormat kami,\n" +
    "Sistem Perizinan Sekolah\n" +
    "SMKN 9 Semarang\n\n" +
    "_Pesan ini dikirim secara otomatis oleh Sistem/Bot Resmi SMKN 9 Semarang._";
  
  return sendWhatsAppMessage_(noHPGuru, message);
}

// =========================
// MAIN PROCESS
// =========================
function prepareFormData(formData, studentData, kodeKonfirmasi) {
  formData = formData || {};
  studentData = Array.isArray(studentData) ? studentData : [];

  var timestamp = Utilities.formatDate(new Date(), CONFIG.TZ, "dd/MM/yyyy, HH.mm.ss");

  var jumlahSiswa = parseInt(formData.jumlahSiswa, 10) || (studentData.length || 1);
  if (jumlahSiswa < 1) jumlahSiswa = 1;
  if (jumlahSiswa > CONFIG.MAX_STUDENTS) jumlahSiswa = CONFIG.MAX_STUDENTS;

  var first = studentData[0] || {};
  var tingkat = first.tingkat ? first.tingkat : "";
  var kelas = first.jurusan ? first.jurusan : "";

  var row = [
    timestamp,
    formData.keperluan || "",
    formData.tanggal || "",
    formData.jamAwal || "",
    formData.jamAkhir || "",
    formData.namaGuru || "",
    (kodeKonfirmasi || ""),
    "Menunggu Konfirmasi",
    "",
    jumlahSiswa,
    tingkat,
    kelas
  ];

  for (var i = 0; i < CONFIG.MAX_STUDENTS; i++) {
    row.push((studentData[i] && studentData[i].nama) ? studentData[i].nama : "");
  }
  for (var j = 0; j < CONFIG.MAX_STUDENTS; j++) row.push("");
  for (var k = 0; k < CONFIG.MAX_STUDENTS; k++) row.push("");

  return row;
}

function processForm(formData) {
  try {
    formData = formData || {};
    if (!Object.keys(formData).length) throw new Error("Form data kosong.");

    var pick = getLogSheet_(true);
    var sheet = pick.sheet;

    if (!sheet) {
      var ass = getActiveSS_();
      sheet = createSheet(ass);
      pick = { ss: ass, sheet: sheet, where: "ACTIVE(created)" };
    }

    var studentData = [];
    var jumlahSiswa = parseInt(formData.jumlahSiswa, 10) || 1;
    if (jumlahSiswa < 1) jumlahSiswa = 1;
    if (jumlahSiswa > CONFIG.MAX_STUDENTS) jumlahSiswa = CONFIG.MAX_STUDENTS;

    for (var i = 1; i <= jumlahSiswa; i++) {
      studentData.push({
        nama: (formData["nama" + i] || "").toString().trim(),
        nis: (formData["nis" + i] || "").toString().trim(),
        tingkat: (formData["tingkat" + i] || formData["angkatan" + i] || "").toString().trim().toUpperCase(),
        jurusan: (formData["kelas" + i] || formData["jurusan" + i] || "").toString().trim()
      });
    }

    for (var v = 0; v < studentData.length; v++) {
      if (!studentData[v].nama || !studentData[v].tingkat || !studentData[v].jurusan) {
        throw new Error("Data siswa " + (v + 1) + " wajib lengkap.");
      }
    }

    var seen = {};
    for (var u = 0; u < studentData.length; u++) {
      var key = (studentData[u].nis || (studentData[u].nama + "-" + studentData[u].tingkat + "-" + studentData[u].jurusan)).toLowerCase();
      if (seen[key]) throw new Error("Terdapat siswa yang sama dalam daftar izin.");
      seen[key] = true;
    }

    var tglStr = (formData.tanggal || "").toString().trim();
    if (!tglStr) throw new Error("Tanggal tidak valid");
    var dParts = tglStr.split("-");
    if (dParts.length !== 3) throw new Error("Tanggal tidak valid");
    var submissionDate = new Date(Number(dParts[0]), Number(dParts[1]) - 1, Number(dParts[2]));
    if (isNaN(submissionDate.getTime())) throw new Error("Tanggal tidak valid");

    var dayWib = parseInt(Utilities.formatDate(submissionDate, CONFIG.TZ, "u"), 10);
    if (dayWib === 6 || dayWib === 7) throw new Error("Pengajuan izin tidak dapat dilakukan pada hari Sabtu atau Minggu.");

    // VALIDASI JAM KE
    var jamAwalKe = (formData.jamAwal || "").toString().trim();
    var jamAkhirKe = (formData.jamAkhir || "").toString().trim();

    if (!jamAwalKe || jamAwalKe === "") {
      throw new Error("Jam Keluar harus dipilih.");
    }
    if (!jamAkhirKe || jamAkhirKe === "") {
      throw new Error("Jam Kembali harus dipilih.");
    }

    var jamAwalInt = parseInt(jamAwalKe, 10);
    var jamAkhirInt = parseInt(jamAkhirKe, 10);
    
    if (isNaN(jamAwalInt) || jamAwalInt < 1 || jamAwalInt > 10) {
      throw new Error("Jam Keluar tidak valid. Pilih Jam ke 1-10.");
    }
    if (isNaN(jamAkhirInt) || jamAkhirInt < 1 || jamAkhirInt > 10) {
      throw new Error("Jam Kembali tidak valid. Pilih Jam ke 1-10.");
    }
    if (jamAwalInt >= jamAkhirInt) {
      throw new Error("Jam Kembali harus lebih akhir dari Jam Keluar.");
    }

    var guruPiketName = (formData.namaGuru || "").toString().trim();
    var guruPiketList = getGuruPiketList();
    if (!guruPiketList.length) throw new Error("Guru Piket hari ini tidak ditemukan.");
    
    var guruFound = false;
    var guruObj = null;
    for (var g = 0; g < guruPiketList.length; g++) {
      if (guruPiketList[g].nama === guruPiketName) {
        guruFound = true;
        guruObj = guruPiketList[g];
        break;
      }
    }

    if (!guruFound) {
      throw new Error("Guru Piket yang dipilih (" + guruPiketName + ") tidak valid.");
    }

    var noHPGuru = guruObj ? guruObj.phone : "";
    if (!noHPGuru) throw new Error("Nomor telepon Guru Piket tidak tersedia.");

    var kodeKonfirmasi = generateConfirmationCode();
    var row = prepareFormData(formData, studentData, kodeKonfirmasi);

    // Generate PDF
    var out = generateDocAndPdf_(formData, studentData, kodeKonfirmasi);
    var docUrl = out.docUrl;
    var pdfUrl = out.pdfUrl;

    saveToSpreadsheet(sheet, row);
    SpreadsheetApp.flush();

    var noWaOrtuList = getNoWaOrtuList_(studentData);

    sendNotifications_(formData, studentData, pdfUrl, kodeKonfirmasi, noHPGuru, noWaOrtuList, docUrl);

    return {
      success: true,
      pdfUrl: pdfUrl,
      docUrl: docUrl,
      message: "Permohonan izin berhasil dikirim.",
      savedTo: pick.where,
      waOrtuCount: noWaOrtuList.length,
      kodeKonfirmasi: kodeKonfirmasi
    };
  } catch (error) {
    Logger.log("processForm Error: " + error.message + "\n" + error.stack);
    throw new Error("Gagal memproses formulir: " + error.message);
  }
}

// =========================
// WA NOTIFICATIONS - DENGAN FORMAT JAM KE
// =========================
function sendNotifications_(formData, studentData, pdfUrl, kodeKonfirmasi, noHPGuru, noWaOrtuList, docUrl) {
  formData = formData || {};
  studentData = Array.isArray(studentData) ? studentData : [];

  var namaSiswa = studentData.map(function (s) { return s.nama; }).filter(Boolean).join(", ");
  var hariTanggal = formatHariTanggal_(formData.tanggal || "");
  
  var jamPelajaran = getRentangJamKe(formData.jamAwal, formData.jamAkhir);

  var baseUrl = getWebAppUrl_();
  var linkKonfirmasi = baseUrl ? (baseUrl + "?kode=" + encodeURIComponent(kodeKonfirmasi)) : "";

  // ==================== PESAN UNTUK GURU PIKET ====================
  var messageGuru = 
    "*Notifikasi : Pemberitahuan Izin Keluar Sekolah*\n\n" +
    "Yth. Bapak/Ibu Guru Piket\n" +
    "SMKN 9 Semarang\n\n" +
    "Dengan hormat,\n\n" +
    "Melalui pesan ini kami memberitahukan bahwa siswa berikut telah mengajukan dan memperoleh persetujuan izin keluar sekolah:\n\n" +
    "Nama Siswa : *" + namaSiswa + "*\n" +
    "Hari/Tanggal : " + hariTanggal + "\n" +
    "Jam Pelajaran : " + jamPelajaran + "\n" +
    "Keperluan : " + (formData.keperluan || "-") + "\n" +
    "Disetujui oleh : " + (formData.namaGuru || "") + "\n\n" +
    "Kode Konfirmasi : *" + kodeKonfirmasi + "*\n" +
    "Konfirmasi Kepulangan :\n" + linkKonfirmasi + "\n\n" +
    "Surat Izin Resmi (PDF) :\n" + pdfUrl + "\n\n" +
    "Demikian pemberitahuan ini kami sampaikan. Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.\n\n" +
    "Hormat kami,\n" +
    "Sistem Perizinan Sekolah\n" +
    "SMKN 9 Semarang\n\n" +
    "_Pesan ini dikirim secara otomatis oleh Sistem/Bot Resmi SMK NEGERI 9 SEMARANG._";

  // ==================== PESAN UNTUK ORANG TUA ====================
  var messageWali = 
    "*Notifikasi : Informasi Izin Keluar Sekolah*\n\n" +
    "Yth. Bapak/Ibu Wali Murid\n" +
    "*" + namaSiswa + "*\n\n" +
    "Dengan hormat,\n\n" +
    "Kami informasikan bahwa siswa dengan data berikut telah mendapatkan persetujuan izin keluar sekolah:\n\n" +
    "Nama Siswa : *" + namaSiswa + "*\n" +
    "Hari/Tanggal : " + hariTanggal + "\n" +
    "Jam Pelajaran : " + jamPelajaran + "\n" +
    "Keperluan : " + (formData.keperluan || "-") + "\n" +
    "Disetujui oleh : " + (formData.namaGuru || "") + "\n\n" +
    "Surat Izin Resmi (PDF) dapat diakses melalui tautan berikut:\n" + pdfUrl + "\n\n" +
    "Demikian informasi ini kami sampaikan. Terima kasih atas perhatian dan kerja sama Bapak/Ibu.\n\n" +
    "Hormat kami,\n" +
    "Sistem Perizinan Sekolah\n" +
    "SMKN 9 Semarang\n\n" +
    "_Pesan ini dikirim secara otomatis oleh Sistem/Bot Resmi SMKN 9 SEMARANG._";

  // Kirim ke Guru
  var guruResult = sendWhatsAppMessage_(noHPGuru, messageGuru);
  if (!guruResult.success) {
    Logger.log("Gagal kirim WA ke guru: " + guruResult.message);
  }

  // Kirim ke Orang Tua
  if (Array.isArray(noWaOrtuList) && noWaOrtuList.length > 0) {
    var sentCount = 0;
    for (var i = 0; i < noWaOrtuList.length; i++) {
      var noWa = noWaOrtuList[i];
      if (noWa && String(noWa).trim() !== "") {
        var ortResult = sendWhatsAppMessage_(noWa, messageWali);
        if (ortResult.success) sentCount++;
      }
    }
    Logger.log("Pesan WA dikirim ke " + sentCount + " orang tua");
  } else {
    Logger.log("Tidak ada nomor WA orang tua");
  }

  return true;
}

// =========================
// KONFIRMASI PER-NAMA
// =========================
function ensureConfirmColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return (h || "").toString(); });

  var nama1Index = findHeaderIndex_(headers, ["Nama 1", "NAMA 1"]);
  if (nama1Index === -1) throw new Error("Kolom 'Nama 1' tidak ditemukan.");

  var kon1Index = findHeaderIndex_(headers, ["Konfirmasi 1", "KONFIRMASI 1"]);
  var wkon1Index = findHeaderIndex_(headers, ["Waktu Konfirmasi 1", "WAKTU KONFIRMASI 1"]);

  if (kon1Index !== -1 && wkon1Index !== -1) {
    return {
      headers: headers,
      namaStart: nama1Index,
      konfirmasiStart: kon1Index,
      waktuKonfirmasiStart: wkon1Index,
      lastCol: lastCol
    };
  }

  var addHeaders = [];
  for (var i = 1; i <= CONFIG.MAX_STUDENTS; i++) addHeaders.push("Konfirmasi " + i);
  for (var j = 1; j <= CONFIG.MAX_STUDENTS; j++) addHeaders.push("Waktu Konfirmasi " + j);

  sheet.getRange(1, lastCol + 1, 1, addHeaders.length).setValues([addHeaders]);
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  var newLastCol = sheet.getLastColumn();
  var newHeaders = sheet.getRange(1, 1, 1, newLastCol).getDisplayValues()[0].map(function (h) { return (h || "").toString(); });

  var newKon1Index = findHeaderIndex_(newHeaders, ["Konfirmasi 1"]);
  var newWKon1Index = findHeaderIndex_(newHeaders, ["Waktu Konfirmasi 1"]);

  return {
    headers: newHeaders,
    namaStart: nama1Index,
    konfirmasiStart: newKon1Index,
    waktuKonfirmasiStart: newWKon1Index,
    lastCol: newLastCol
  };
}

function normalizeKode_(kode) {
  return (kode || "")
    .toString()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "")
    .slice(0, 6);
}

function getRowByKode_(sheet, kode, headers) {
  var kodeIndex = findHeaderIndex_(headers, ["Kode Konfirmasi", "KODE KONFIRMASI", "Kode"]);
  if (kodeIndex === -1) throw new Error("Kolom 'Kode Konfirmasi' tidak ditemukan.");

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var cell = sheet.getRange(2, kodeIndex + 1, lastRow - 1, 1)
    .createTextFinder(kode)
    .matchCase(false)
    .matchEntireCell(true)
    .findNext();

  if (!cell) return null;
  return cell.getRow();
}

function parseTanggalCellToYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TZ, "yyyy-MM-dd");
  return (v || "").toString().trim();
}

function isExpiredByTanggal_(tanggalYmd) {
  var t = (tanggalYmd || "").toString().trim();
  if (!t) return false;
  var todayYmd = Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd");
  return (t < todayYmd);
}

function getConfirmInfo(kode) {
  try {
    var k = normalizeKode_(kode);
    if (!k || k.length !== 6) return { success: false, message: "Kode tidak valid." };

    var pick = getLogSheet_(true);
    var sheet = pick.sheet;
    if (!sheet) return { success: false, message: "Sheet log tidak ditemukan." };

    var meta = ensureConfirmColumns_(sheet);
    var headers = meta.headers;

    var rowNum = getRowByKode_(sheet, k, headers);
    if (!rowNum) return { success: false, message: "Kode konfirmasi tidak ditemukan." };

    var row = sheet.getRange(rowNum, 1, 1, meta.lastCol).getValues()[0];

    var tanggalIndex = findHeaderIndex_(headers, ["Tanggal", "TANGGAL"]);
    var tanggalYmd = parseTanggalCellToYmd_(row[tanggalIndex]);

    var siswa = [];
    var total = 0;
    var done = 0;

    for (var i = 0; i < CONFIG.MAX_STUDENTS; i++) {
      var nama = (row[meta.namaStart + i] || "").toString().trim();
      if (!nama) continue;

      total++;

      var konfVal = (row[meta.konfirmasiStart + i] || "").toString().trim().toLowerCase();
      var confirmed = (konfVal === "selesai" || konfVal === "true" || konfVal === "ya" || konfVal === "ok");

      if (confirmed) done++;

      siswa.push({ idx: i + 1, nama: nama, confirmed: confirmed });
    }

    return {
      success: true,
      kode: k,
      tanggal: formatHariTanggal_(tanggalYmd),
      tanggalYmd: tanggalYmd,
      expired: isExpiredByTanggal_(tanggalYmd),
      total: total,
      done: done,
      siswa: siswa,
      where: pick.where
    };
  } catch (e) {
    return { success: false, message: "Gagal memuat info: " + e.message };
  }
}

function confirmReturn(kode, nama) {
  try {
    var k = normalizeKode_(kode);
    var n = (nama || "").toString().trim();

    if (!k || k.length !== 6) return { success: false, message: "Kode tidak valid." };
    if (!n) return { success: false, message: "Nama wajib dipilih." };

    var pick = getLogSheet_(true);
    var sheet = pick.sheet;
    if (!sheet) return { success: false, message: "Sheet log tidak ditemukan." };

    var meta = ensureConfirmColumns_(sheet);
    var headers = meta.headers;

    var rowNum = getRowByKode_(sheet, k, headers);
    if (!rowNum) return { success: false, message: "Kode konfirmasi tidak ditemukan." };

    var row = sheet.getRange(rowNum, 1, 1, meta.lastCol).getValues()[0];

    var tanggalIndex = findHeaderIndex_(headers, ["Tanggal", "TANGGAL"]);
    var statusIndex  = findHeaderIndex_(headers, ["Status", "STATUS"]);
    var confirmIndex = findHeaderIndex_(headers, ["Waktu Konfirmasi", "WAKTU KONFIRMASI"]);
    var namaGuruIndex = findHeaderIndex_(headers, ["Nama Guru", "NAMA GURU", "Guru", "GURU"]);

    var tanggalYmd = parseTanggalCellToYmd_(row[tanggalIndex]);
    var namaGuruRow = (namaGuruIndex !== -1) ? (row[namaGuruIndex] || "").toString().trim() : "";

    if (isExpiredByTanggal_(tanggalYmd)) {
      return { success: false, expired: true, message: "Konfirmasi sudah ditutup." };
    }

    var foundIdx = -1;
    for (var i = 0; i < CONFIG.MAX_STUDENTS; i++) {
      var namaCell = (row[meta.namaStart + i] || "").toString().trim();
      if (!namaCell) continue;
      if (namaCell === n) { foundIdx = i; break; }
    }
    if (foundIdx === -1) return { success: false, message: "Nama tidak ditemukan." };

    var konfCellVal = (row[meta.konfirmasiStart + foundIdx] || "").toString().trim().toLowerCase();
    var already = (konfCellVal === "selesai" || konfCellVal === "true" || konfCellVal === "ya" || konfCellVal === "ok");

    if (already) {
      return {
        success: true,
        alreadyConfirmed: true,
        message: "Sudah dikonfirmasi sebelumnya.",
        nama: n,
        tanggal: formatHariTanggal_(tanggalYmd)
      };
    }

    var waktuKonfirmasi = Utilities.formatDate(new Date(), CONFIG.TZ, "dd/MM/yyyy, HH.mm.ss");

    sheet.getRange(rowNum, meta.konfirmasiStart + foundIdx + 1).setValue("Selesai");
    sheet.getRange(rowNum, meta.waktuKonfirmasiStart + foundIdx + 1).setValue(waktuKonfirmasi);

    // ===== KIRIM NOTIFIKASI KE GURU (HANYA 1 KALI) =====
    if (namaGuruRow) {
      var noHPGuru = getGuruPiketPhoneByName_(namaGuruRow);
      if (noHPGuru) {
        var waktuFormatted = waktuKonfirmasi.replace(/,/g, '');
        sendKonfirmasiNotifToGuru(noHPGuru, n, waktuFormatted);
      }
    }
    // ===== END NOTIFIKASI =====

    var total = 0;
    var done = 0;
    for (var j = 0; j < CONFIG.MAX_STUDENTS; j++) {
      var namaJ = (sheet.getRange(rowNum, meta.namaStart + j + 1).getDisplayValue() || "").toString().trim();
      if (!namaJ) continue;
      total++;
      var konfJ = (sheet.getRange(rowNum, meta.konfirmasiStart + j + 1).getDisplayValue() || "").toString().trim().toLowerCase();
      var ok = (konfJ === "selesai" || konfJ === "true" || konfJ === "ya" || konfJ === "ok");
      if (ok) done++;
    }

    if (statusIndex !== -1 && confirmIndex !== -1) {
      if (total > 0 && done >= total) {
        sheet.getRange(rowNum, statusIndex + 1).setValue("Selesai");
        sheet.getRange(rowNum, confirmIndex + 1).setValue(waktuKonfirmasi);
        sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).setBackground("#ccffcc");
      } else {
        var st = sheet.getRange(rowNum, statusIndex + 1).getDisplayValue();
        if (!st || st.toString().trim() === "") sheet.getRange(rowNum, statusIndex + 1).setValue("Menunggu Konfirmasi");
      }
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      alreadyConfirmed: false,
      message: "Kepulangan berhasil dikonfirmasi.",
      nama: n,
      tanggal: formatHariTanggal_(tanggalYmd),
      waktuKonfirmasi: waktuKonfirmasi,
      where: pick.where
    };

  } catch (error) {
    Logger.log("confirmReturn Error: " + error.message + "\n" + (error.stack || ""));
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

// =========================
// CEK DAN KIRIM NOTIFIKASI KONFIRMASI TERLAMBAT
// =========================

function cekDanKirimNotifikasiTerlambat() {
  try {
    Logger.log("=== MEMULAI PENGECEKAN KONFIRMASI TERLAMBAT ===");
    
    var pick = getLogSheet_(true);
    var sheet = pick.sheet;
    if (!sheet) {
      return { success: false, message: "Sheet log tidak ditemukan" };
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      return { success: true, message: "Tidak ada data" };
    }

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = (data[0] || []).map(function (h) { return (h || "").toString(); });

    var tanggalIndex = findHeaderIndex_(headers, ["Tanggal"]);
    var statusIndex = findHeaderIndex_(headers, ["Status"]);
    var namaGuruIndex = findHeaderIndex_(headers, ["Nama Guru"]);
    var kodeIndex = findHeaderIndex_(headers, ["Kode Konfirmasi"]);
    var namaStartIndex = findHeaderIndex_(headers, ["Nama 1"]);
    var konfirmasiStartIndex = findHeaderIndex_(headers, ["Konfirmasi 1"]);

    if (tanggalIndex === -1 || statusIndex === -1 || namaGuruIndex === -1 || 
        kodeIndex === -1 || namaStartIndex === -1 || konfirmasiStartIndex === -1) {
      return { success: false, message: "Format spreadsheet tidak valid" };
    }

    var today = new Date();
    var todayYmd = Utilities.formatDate(today, CONFIG.TZ, "yyyy-MM-dd");

    var izinTerlambat = [];
    var izinSudahDikirim = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var tanggalIzin = parseTanggalCellToYmd_(row[tanggalIndex]);
      if (!tanggalIzin) continue;

      if (tanggalIzin >= todayYmd) continue;

      var status = (row[statusIndex] || "").toString().trim().toLowerCase();
      if (status.indexOf("menunggu") === -1) continue;

      var semuaSiswa = [];
      var siswaBelumKonfirmasi = [];
      
      for (var j = 0; j < CONFIG.MAX_STUDENTS; j++) {
        var namaSiswa = (row[namaStartIndex + j] || "").toString().trim();
        if (!namaSiswa) continue;
        
        semuaSiswa.push(namaSiswa);
        
        var konfirmasiStatus = (row[konfirmasiStartIndex + j] || "").toString().trim().toLowerCase();
        if (konfirmasiStatus !== "selesai" && konfirmasiStatus !== "true" && konfirmasiStatus !== "ya" && konfirmasiStatus !== "ok") {
          siswaBelumKonfirmasi.push(namaSiswa);
        }
      }

      if (siswaBelumKonfirmasi.length > 0) {
        var kode = (row[kodeIndex] || "").toString().trim();
        var namaGuru = (row[namaGuruIndex] || "").toString().trim();
        
        if (izinSudahDikirim.indexOf(kode) === -1) {
          izinTerlambat.push({
            rowNum: i + 1,
            tanggal: tanggalIzin,
            tanggalFormatted: formatHariTanggal_(tanggalIzin),
            kode: kode,
            namaGuru: namaGuru,
            semuaSiswa: semuaSiswa,
            siswaBelumKonfirmasi: siswaBelumKonfirmasi,
            jumlahBelum: siswaBelumKonfirmasi.length,
            jumlahTotal: semuaSiswa.length
          });
          
          izinSudahDikirim.push(kode);
        }
      }
    }

    if (izinTerlambat.length > 0) {
      var hasilKirim = kirimNotifikasiTerlambatKeGuruPiket(izinTerlambat);
      tandaiIzinTerlambat(sheet, izinTerlambat);
      
      return { 
        success: true, 
        message: "Notifikasi dikirim untuk " + izinTerlambat.length + " izin terlambat",
        data: izinTerlambat,
        hasilKirim: hasilKirim
      };
    } else {
      return { success: true, message: "Tidak ada izin terlambat" };
    }

  } catch (error) {
    Logger.log("cekDanKirimNotifikasiTerlambat Error: " + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

function kirimNotifikasiTerlambatKeGuruPiket(izinTerlambat) {
  try {
    var semuaGuruPiket = getAllGuruPiket_();
    
    if (semuaGuruPiket.length === 0) {
      return false;
    }

    var daftarIzinText = "";
    
    for (var i = 0; i < izinTerlambat.length; i++) {
      var izin = izinTerlambat[i];
      
      daftarIzinText += "\n" + (i+1) + ". *Tanggal: " + izin.tanggalFormatted + "*";
      daftarIzinText += "\n   Guru: " + izin.namaGuru;
      daftarIzinText += "\n   Total Siswa: " + izin.jumlahTotal;
      daftarIzinText += "\n   Belum Konfirmasi: " + izin.jumlahBelum;
      daftarIzinText += "\n   Nama Siswa Belum Konfirmasi:";
      
      for (var j = 0; j < izin.siswaBelumKonfirmasi.length; j++) {
        daftarIzinText += "\n   • " + izin.siswaBelumKonfirmasi[j];
      }
      
      daftarIzinText += "\n   Kode: " + izin.kode + "\n";
    }

    var messageHeader = 
      "*⚠️ NOTIFIKASI KONFIRMASI TERLAMBAT*\n\n" +
      "Ditemukan " + izinTerlambat.length + " izin yang masih belum dikonfirmasi kepulangannya melebihi batas waktu (" + 
      Utilities.formatDate(new Date(), CONFIG.TZ, "dd/MM/yyyy") + ").\n\n" +
      "*DAFTAR IZIN TERLAMBAT:*\n";

    var messageFooter = 
      "\n\nMohon segera ditindaklanjuti.\n" +
      "Terima kasih.\n\n" +
      "_Pesan ini dikirim otomatis oleh sistem._";

    var fullMessage = messageHeader + daftarIzinText + messageFooter;

    var nomorTerkirim = [];
    var nomorGagal = [];

    for (var g = 0; g < semuaGuruPiket.length; g++) {
      var guru = semuaGuruPiket[g];
      var noHP = guru.phone;
      
      if (noHP && noHP !== "-") {
        var result = sendWhatsAppMessage_(noHP, fullMessage);
        
        if (result.success) {
          nomorTerkirim.push(guru.nama);
        } else {
          nomorGagal.push(guru.nama);
        }
      }
    }

    simpanLogNotifikasiTerlambat(izinTerlambat, nomorTerkirim, nomorGagal);

    return {
      success: true,
      terkirim: nomorTerkirim.length,
      gagal: nomorGagal.length
    };

  } catch (error) {
    Logger.log("kirimNotifikasiTerlambatKeGuruPiket Error: " + error.message);
    return false;
  }
}

function tandaiIzinTerlambat(sheet, izinTerlambat) {
  try {
    for (var i = 0; i < izinTerlambat.length; i++) {
      var rowNum = izinTerlambat[i].rowNum;
      
      sheet.getRange(rowNum, 1, 1, sheet.getLastColumn())
        .setBackground("#ffcccc")
        .setNote("⚠️ Konfirmasi terlambat - Notifikasi telah dikirim pada " + 
                 Utilities.formatDate(new Date(), CONFIG.TZ, "dd/MM/yyyy HH:mm"));
    }
    
    SpreadsheetApp.flush();
    return true;
  } catch (e) {
    return false;
  }
}

function simpanLogNotifikasiTerlambat(izinTerlambat, nomorTerkirim, nomorGagal) {
  try {
    var ss = getMasterSS_();
    if (!ss) return false;

    var logSheet = ss.getSheetByName("Log Notifikasi Terlambat");
    
    if (!logSheet) {
      logSheet = ss.insertSheet("Log Notifikasi Terlambat");
      logSheet.appendRow([
        "Timestamp",
        "Jumlah Izin Terlambat",
        "Detail Izin",
        "Nomor Terkirim",
        "Jumlah Terkirim",
        "Nomor Gagal",
        "Jumlah Gagal",
        "Status"
      ]);
    }

    var detailIzin = "";
    for (var i = 0; i < izinTerlambat.length; i++) {
      var izin = izinTerlambat[i];
      detailIzin += izin.tanggal + " (" + izin.kode + "): " + izin.jumlahBelum + "/" + izin.jumlahTotal + " belum konfirmasi\n";
    }

    var timestamp = Utilities.formatDate(new Date(), CONFIG.TZ, "dd/MM/yyyy, HH.mm.ss");
    
    logSheet.appendRow([
      timestamp,
      izinTerlambat.length,
      detailIzin,
      nomorTerkirim.join(", "),
      nomorTerkirim.length,
      nomorGagal.join(", "),
      nomorGagal.length,
      "Terkirim: " + nomorTerkirim.length + ", Gagal: " + nomorGagal.length
    ]);

    return true;
  } catch (e) {
    return false;
  }
}

function setupNotifikasiOtomatis() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "cekDanKirimNotifikasiTerlambat") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    ScriptApp.newTrigger("cekDanKirimNotifikasiTerlambat")
      .timeBased()
      .atHour(20)
      .nearMinute(0)
      .inTimezone(CONFIG.TZ)
      .everyDays(1)
      .create();

    return { success: true, message: "Trigger notifikasi otomatis berhasil dibuat" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function runCekTerlambatManual() {
  return cekDanKirimNotifikasiTerlambat();
}

// =========================
// Statistik Harian
// =========================
function getTodayStats() {
  try {
    var pick = getLogSheet_(true);
    var sheet = pick.sheet;
    if (!sheet) return { error: "Sheet log tidak ditemukan" };

    var today = Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd");
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { success: true, totalToday: 0, pendingToday: 0, pendingNames: [], date: today };

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = (data[0] || []).map(function (h) { return (h || "").toString(); });

    var tanggalIndex = findHeaderIndex_(headers, ["Tanggal"]);
    var statusIndex  = findHeaderIndex_(headers, ["Status"]);
    var namaIndex    = findHeaderIndex_(headers, ["Nama 1"]);

    if (tanggalIndex === -1 || statusIndex === -1 || namaIndex === -1) {
      return { error: "Format spreadsheet tidak valid" };
    }

    var totalToday = 0, pendingToday = 0, pendingNames = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowDate = row[tanggalIndex];

      if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, CONFIG.TZ, "yyyy-MM-dd");
      rowDate = (rowDate || "").toString().trim();

      if (rowDate === today) {
        totalToday++;
        if (((row[statusIndex] || "").toString().trim()).toLowerCase().indexOf("menunggu") !== -1) {
          pendingToday++;
          pendingNames.push(row[namaIndex]);
        }
      }
    }

    return { success: true, totalToday: totalToday, pendingToday: pendingToday, pendingNames: pendingNames, date: today };
  } catch (error) {
    return { error: "Gagal mengambil statistik: " + error.message };
  }
}

function checkExpiredCodes() { 
  return { success: true, message: "Fungsi checkExpiredCodes berjalan" }; 
}

// =========================
// REKAP MINGGUAN
// =========================

function getMondayOfCurrentWeek_() {
  var now = new Date();
  var day = now.getDay() || 7; 
  if (day !== 1) now.setHours(-24 * (day - 1)); 
  return now;
}

function generateWeeklyRecapReport_() {
  try {
    var pick = getLogSheet_(true);
    var sheet = pick.sheet;
    if (!sheet) return "Error: Sheet log tidak ditemukan.";

    var monday = getMondayOfCurrentWeek_();
    var mondayYmd = Utilities.formatDate(monday, CONFIG.TZ, "yyyy-MM-dd");
    
    var friday = new Date(monday);
    friday.setDate(monday.getDate() + 4);
    var fridayYmd = Utilities.formatDate(friday, CONFIG.TZ, "yyyy-MM-dd");

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return "Belum ada data izin minggu ini.";

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = (data[0] || []).map(function (h) { return (h || "").toString(); });

    var tanggalIndex = findHeaderIndex_(headers, ["Tanggal"]);

    if (tanggalIndex === -1) return "Format sheet salah (butuh Tanggal).";

    var studentCounts = {};
    var totalEntries = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowDate = row[tanggalIndex];

      if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, CONFIG.TZ, "yyyy-MM-dd");
      rowDate = (rowDate || "").toString().trim();

      if (rowDate >= mondayYmd && rowDate <= fridayYmd) {
        totalEntries++;
        
        for (var n = 0; n < CONFIG.MAX_STUDENTS; n++) {
          var namaIdx = findHeaderIndex_(headers, ["Nama " + (n + 1)]);
          if (namaIdx !== -1) {
            var nama = (row[namaIdx] || "").toString().trim();
            if (nama) {
              if (!studentCounts[nama]) studentCounts[nama] = 0;
              studentCounts[nama]++;
            }
          }
        }
      }
    }

    if (totalEntries === 0) return "Tidak ada data izin minggu ini.";

    var sortedNames = Object.keys(studentCounts).sort(function (a, b) {
      return studentCounts[b] - studentCounts[a];
    });

    var report = "*LAPORAN REKAP IZIN SISWA (MINGGU INI)*\n";
    report += "Periode: " + Utilities.formatDate(monday, CONFIG.TZ, "dd MMM yyyy") + " - " + Utilities.formatDate(friday, CONFIG.TZ, "dd MMM yyyy") + "\n";
    report += "Total Pengajuan: " + totalEntries + "\n\n";
    report += "*Ranking Izin Terbanyak:*\n";

    for (var k = 0; k < sortedNames.length; k++) {
      var nm = sortedNames[k];
      report += (k + 1) + ". " + nm + " (" + studentCounts[nm] + "x)\n";
    }

    return report;

  } catch (e) {
    return "Gagal generate rekap: " + e.message;
  }
}

function runWeeklyRecapToWa_() {
  try {
    var report = generateWeeklyRecapReport_();
    if (report.indexOf("Error") !== -1 || report.indexOf("Format sheet salah") !== -1) {
      return { success: false, message: report };
    }

    var adminNum = CONFIG.ADMIN_WA;
    if (!adminNum || adminNum === "62xxxxxxxxxx") {
      return { success: false, message: "ADMIN_WA belum disetting" };
    }

    var res = sendWhatsAppMessage_(adminNum, report);
    return { success: true, message: "Rekap mingguan terkirim", result: res };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// =========================
// DEBUG
// =========================
function debugPiket() {
  try {
    var ss = getMasterSS_();
    var sheet = ss.getSheetByName(CONFIG.GURU_PIKET_SHEET_NAME);
    
    if (!sheet) {
      Logger.log("ERROR: Sheet '" + CONFIG.GURU_PIKET_SHEET_NAME + "' tidak ditemukan!");
      return;
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log("=== HASIL DIAGNOSA GURU KE-10 ===");
    
    var idxNama10 = -1;
    var idxWa10 = -1;

    for (var j = 0; j < headers.length; j++) {
      var h = String(headers[j]).trim().toLowerCase();
      
      if (h === "nama guru 10") idxNama10 = j;
      if (h === "wa guru 10") idxWa10 = j;
    }

    Logger.log("1. Index Header 'NAMA GURU 10': " + idxNama10);
    Logger.log("2. Index Header 'WA GURU 10':   " + idxWa10);

    var data = sheet.getDataRange().getValues();
    var rowRabu = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === "rabu") {
        rowRabu = i;
        break;
      }
    }

    if (rowRabu === -1) {
      Logger.log("ERROR: Baris 'Rabu' tidak ditemukan!");
      return;
    }

    Logger.log("3. Baris 'Rabu' ditemukan di index: " + rowRabu + " (Baris ke-" + (rowRabu + 1) + ")");
    Logger.log("--------------------------------------------------");
    
    if (idxNama10 === -1) {
      Logger.log("❌ PROBLEM DITEMUKAN: Header 'NAMA GURU 10' TIDAK DITEMUKAN.");
    } else {
      var namaGuru10 = data[rowRabu][idxNama10];
      Logger.log("✅ Isi Kolom NAMA GURU 10: " + namaGuru10);
    }

    if (idxWa10 === -1) {
      Logger.log("❌ PROBLEM DITEMUKAN: Header 'WA GURU 10' TIDAK DITEMUKAN.");
    } else {
      var waGuru10 = data[rowRabu][idxWa10];
      Logger.log("✅ Isi Kolom WA GURU 10:   " + waGuru10);
    }

  } catch (e) {
    Logger.log("Error Diagnosa: " + e.toString());
  }
}