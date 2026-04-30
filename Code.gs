/**
 * OWNER COMMAND CENTER - BPR KS
 * Versi: 181.0 (Base) + Add-on Coaching Interactive Table
 */

var DB_ID = '1fKi00TOdtqwlyP6Uk20sGbJHrm_dsQ5frlkFm6jMAzE';

// Embedded admin users (development bootstrap)
// NIK | Nama | Cabang | Password | Role_App | Status_User
var EMBEDDED_ADMIN_USERS = {
  '2510285': { nik: '2510285', nama: 'YULIUS PUJI LAKSONO', cabang: 'LEUWIPANJANG', password: '2510285', roleApp: 'ADMIN', statusUser: 'AKTIF' },
  '1234567': { nik: '1234567', nama: 'ADMIN1', cabang: 'LEUWIPANJANG', password: '1234567', roleApp: 'ADMIN', statusUser: 'AKTIF' },
  '1111111': { nik: '1111111', nama: 'ADMIN2', cabang: 'LEUWIPANJANG', password: '1111111', roleApp: 'ADMIN', statusUser: 'AKTIF' },
  '9999999': { nik: '9999999', nama: 'ADMIN3', cabang: 'LEUWIPANJANG', password: '9999999', roleApp: 'ADMIN', statusUser: 'AKTIF' }
};

function getEmbeddedAdminByNik(nik) {
  var key = String(nik || '').trim();
  if (!key) return null;
  var rec = EMBEDDED_ADMIN_USERS[key];
  if (!rec) return null;
  if (String(rec.statusUser || '').toUpperCase() !== 'AKTIF') return null;
  return rec;
}

function getDB() {
  return SpreadsheetApp.openById(DB_ID);
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Executive Dashboard - BPR KS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUserData() {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase();
    var ss = getDB();
    var sh = ss.getSheetByName('User_ID');
    var shEmp = ss.getSheetByName('DB_Karyawan') || ss.getSheetByName('Master_Karyawan');
    var data = sh ? sh.getDataRange().getValues() : [];
    var devAdminNik = '2510285';
    if (email.includes("yulius")) return { email: email, nama: "YULIUS PUJI LAKSONO", role: "Admin", cabang: "ALL", nik: devAdminNik };

    // Development override: NIK tertentu diperlakukan sebagai Admin
    // berdasarkan mapping email di DB_Karyawan. Jika Apps Script tidak
    // memberikan email aktif (kosong), tetap izinkan mode dev.
    if (shEmp && shEmp.getLastRow() > 1) {
      var empRows = shEmp.getDataRange().getValues();
      var hdr = empRows[0] || [];
      var idxNik = -1, idxNama = -1, idxCab = -1, idxEmail = -1;
      for (var c = 0; c < hdr.length; c++) {
        var lab = String(hdr[c] || '').toLowerCase();
        if (idxNik < 0 && (lab.indexOf('nik') !== -1 || lab.indexOf('npk') !== -1)) idxNik = c;
        if (idxNama < 0 && lab.indexOf('nama') !== -1) idxNama = c;
        if (idxCab < 0 && (lab.indexOf('cabang') !== -1 || lab.indexOf('branch') !== -1)) idxCab = c;
        if (idxEmail < 0 && lab.indexOf('email') !== -1) idxEmail = c;
      }
      if (idxNik < 0) idxNik = 1;
      if (idxNama < 0) idxNama = 2;
      if (idxCab < 0) idxCab = 5;
      if (idxEmail < 0) idxEmail = 11;

      for (var r = 1; r < empRows.length; r++) {
        var nik = String((empRows[r] && empRows[r][idxNik]) || '').trim();
        var empEmail = String((empRows[r] && empRows[r][idxEmail]) || '').trim().toLowerCase();
        var embeddedAdmin = getEmbeddedAdminByNik(nik);
        if (!embeddedAdmin) {
          // Backward compatibility for previous single dev override
          if (nik !== devAdminNik) continue;
          embeddedAdmin = { nik: nik, nama: String((empRows[r] && empRows[r][idxNama]) || 'ADMIN DEV').trim() || 'ADMIN DEV', cabang: String((empRows[r] && empRows[r][idxCab]) || 'ALL').trim() || 'ALL' };
        }
        if (email === '' || (empEmail && email === empEmail)) {
          return {
            email: email || empEmail,
            nama: embeddedAdmin.nama || String((empRows[r] && empRows[r][idxNama]) || 'ADMIN DEV').trim() || 'ADMIN DEV',
            role: 'Admin',
            cabang: embeddedAdmin.cabang || String((empRows[r] && empRows[r][idxCab]) || 'ALL').trim() || 'ALL',
            nik: embeddedAdmin.nik || nik
          };
        }
      }
    }

    if (email === "") return { email: email, nama: "ADMIN", role: "Admin", cabang: "ALL" };
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toLowerCase() === email) return { email: data[i][0], nama: data[i][1], role: data[i][3], cabang: data[i][4] };
    }
  } catch(e) {}
  return { nama: "ADMIN", role: "Admin", cabang: "ALL" };
}

function isJunkBranch(val) {
  if (!val) return true;
  var v = val.toString().toUpperCase();
  var junk = ["ITEM", "VOLUME", "CABANG", "STAFF", "DEBITUR", "PLAFOND", "APK NAIK", "APP BLM CAIR", "BELUM"];
  return junk.some(word => v.indexOf(word) > -1);
}

function getDashboardData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("db_v181_no_tele");
  if (cached) return JSON.parse(cached);

  var ss = getDB();
  var res = { u: getUserData(), achv: [], cair: [], pipeline: [], listCHome: [], listCPipe: [], listS: [] };

  try {
    ["Raw_Achv_CS", "Raw_Achv_SPV"].forEach(name => {
      var sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 1) return;
      var data = sh.getDataRange().getValues();
      var key = name.split('_')[2].toLowerCase();
      data.forEach(r => {
        if(!r[1] || r[1] === "CABANG") return;
        var tgl = (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : String(r[0]).substring(0,10);
        var cab = String(r[1]).trim();
        if (!isJunkBranch(cab) && res.listCHome.indexOf(cab) === -1) res.listCHome.push(cab);
        res.achv.push([tgl, cab, String(r[2]).trim(), key, Number(r[3])||0, Number(r[4])||0]);
      });
    });

    ["Raw_Cair_CS", "Raw_Cair_SPV"].forEach(name => {
      var sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 1) return;
      var data = sh.getDataRange().getValues();
      data.forEach(r => {
        if(!r[1] || r[1] === "CABANG") return;
        var tgl = (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : String(r[0]).substring(0,10);
        var mapP = [{n:"KABHTSP",i:3,v:4},{n:"KAB",i:5,v:6},{n:"KPLM",i:7,v:8},{n:"KABM",i:9,v:10},{n:"KPSM",i:11,v:12},{n:"KBMBL",i:13,v:14},{n:"KABEKS",i:15,v:16}];
        mapP.forEach(p => {
          var itm = Number(r[p.i]) || 0; var vol = Number(r[p.v]) || 0;
          if (itm > 0 || vol > 0) res.cair.push([tgl, String(r[1]).trim(), String(r[2]).trim(), p.n, itm, vol]);
        });
      });
    });

    var shP = ss.getSheetByName('Raw_Pipeline');
    if (shP && shP.getLastRow() >= 1) {
      var dataP = shP.getDataRange().getValues();
      dataP.forEach(r => {
        if (!r[0] || !r[2] || String(r[2]).toLowerCase().includes("debitur")) return;
        var rawB = r[1], msec;
        if (typeof rawB === 'number') msec = (rawB - 25569) * 86400 * 1000;
        else if (rawB instanceof Date) msec = rawB.getTime();
        else msec = new Date().getTime();
        var tglIn = (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : String(r[0]).substring(0,10);
        var deb = String(r[2]).trim(); var cab = String(r[3]).trim(); var stat = String(r[4]).trim(); var mix = String(r[5]).trim(); var kep = String(r[6]).trim(); var sal = String(r[7] || "-").trim(); var pla = Number(String(r[10]||"0").replace(/[^0-9.-]+/g,"")) || 0; 
        if (cab && !isJunkBranch(cab)) {
          if (res.listCPipe.indexOf(cab) === -1) res.listCPipe.push(cab);
          if (sal !== "-" && res.listS.indexOf(sal) === -1) res.listS.push(sal);
        }
        res.pipeline.push([tglIn, msec, deb, stat, mix, kep, sal, pla, cab]);
      });
    }
    cache.put("db_v181_no_tele", JSON.stringify(res), 300);
    return res;
  } catch(e) { return res; }
}

function processExcelData(rows, tgl, tipe) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = getDB();
    var map = { "1":"Raw_Achv_CS", "2":"Raw_Achv_SPV", "4":"Raw_Cair_CS", "5":"Raw_Cair_SPV", "7":"Raw_Pipeline" };
    if (!map[tipe]) return "❌ Jenis upload tidak dikenali.";
    var sh = ss.getSheetByName(map[tipe]) || ss.insertSheet(map[tipe]);
    var finalData = [];
    var clean = v => { if (!v || v === "-" || v === "" || v === 0) return 0; return parseFloat(v.toString().replace(/\./g, "").replace(/,/g, ".")) || 0; };
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]; if (!r[0] || r[0] === "CABANG" || r[0] === "TGL") continue;
      var row = [tgl];
      for (var c = 0; c < r.length; c++) { 
        var isNum = (tipe == "7" && (c == 9 || c == 10)) || (tipe != "7" && c >= 2); 
        row.push(isNum ? clean(r[c]) : r[c]); 
      }
      finalData.push(row);
    }
    if (finalData.length > 0) {
      var oldData = sh.getDataRange().getValues();
      var filtered = oldData.filter(r => (r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : String(r[0]).substring(0,10)) !== tgl);
      sh.clearContents();
      if (filtered.length > 0) sh.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
      sh.getRange(sh.getLastRow() + 1, 1, finalData.length, finalData[0].length).setValues(finalData);
      CacheService.getScriptCache().remove("db_v181_no_tele");
      return "✅ Berhasil.";
    }
  } catch(e) { return "❌ Error: " + e.message; } finally { lock.releaseLock(); }
}

/* --- ADD-ON: INTERACTIVE COACHING LOGIC --- */

function getCoachingFullData() {
  var ss = getDB();
  var shM = ss.getSheetByName('DB_Karyawan') || ss.getSheetByName('Master_Karyawan');
  var shC = ss.getSheetByName('Raw_Coaching') || ss.insertSheet('Raw_Coaching');
  var shP = ss.getSheetByName('Raw_Coaching_Progress') || ss.insertSheet('Raw_Coaching_Progress');
  var shR = ss.getSheetByName('DB_Region') || ss.getSheetByName('MASTER_REGION');
  var shS = ss.getSheetByName('MASTER_SATUAN');
  
  return {
    master: shM ? shM.getDataRange().getValues() : [],
    coaching: shC.getDataRange().getValues(),
    progress: shP.getDataRange().getValues(),
    regionMap: shR ? shR.getDataRange().getValues() : [],
    satuan: shS ? shS.getDataRange().getValues() : []
  };
}

function getEmployeeMasterData() {
  var ss = getDB();
  var sh = ss.getSheetByName('DB_Karyawan') || ss.getSheetByName('Master_Karyawan');
  if (!sh) return [];
  return sh.getDataRange().getValues();
}

function getDbKaryawanIndexMap(headers) {
  var normalizedHeaders = (headers || []).map(function(h) {
    return String(h || '').toLowerCase().trim();
  });

  function findIndex(candidates, fallback) {
    for (var i = 0; i < normalizedHeaders.length; i++) {
      var label = normalizedHeaders[i];
      var found = candidates.some(function(c) {
        return label.indexOf(String(c).toLowerCase()) !== -1;
      });
      if (found) return i;
    }
    return fallback;
  }

  // Default fallback mengikuti struktur DB_Karyawan (A..N)
  return {
    timestamp: findIndex(['timestamp', 'waktu'], 0),
    nik: findIndex(['nik', 'npk'], 1),
    nama: findIndex(['nama karyawan', 'nama'], 2),
    jabatan: findIndex(['jabatan', 'posisi', 'position'], 3),
    divisi: findIndex(['divisi', 'departemen', 'department'], 4),
    cabang: findIndex(['cabang', 'lokasi kerja', 'branch'], 5),
    roleUser: findIndex(['role user', 'role'], 6),
    nikAtasan: findIndex(['nik atasan langsung'], 7),
    namaAtasan: findIndex(['nama atasan langsung'], 8),
    nikAtasan2: findIndex(['nik atasan dari atasan langsung', 'nik atasan 2'], 9),
    namaAtasan2: findIndex(['nama atasan dari atasan langsung', 'nama atasan 2'], 10),
    email: findIndex(['email karyawan', 'email'], 11),
    wa: findIndex(['no wa', 'wa', 'whatsapp'], 12),
    statusKaryawan: findIndex(['status karyawan', 'status aktif', 'status'], 13)
  };
}

function getAdminEmployeeData() {
  var rows = getEmployeeMasterData();
  if (!rows || rows.length === 0) {
    return {
      headers: [],
      rows: [],
      indexMap: {},
      summary: { total: 0, active: 0, nonActive: 0, branches: 0, roles: 0, hasHierarchy: 0 }
    };
  }

  var headers = rows[0] || [];
  var dataRows = rows.slice(1);
  var idx = getDbKaryawanIndexMap(headers);

  var uniqueBranches = {};
  var uniqueRoles = {};
  var activeCount = 0;
  var nonActiveCount = 0;
  var hierarchyCount = 0;

  var normalizedRows = dataRows.map(function(r) {
    var branch = String((r && r[idx.cabang]) || '').trim().toUpperCase();
    var role = String((r && r[idx.roleUser]) || '').trim().toUpperCase();
    var status = String((r && r[idx.statusKaryawan]) || '').trim().toUpperCase();
    var nikAtasan = String((r && r[idx.nikAtasan]) || '').trim();
    var nikAtasan2 = String((r && r[idx.nikAtasan2]) || '').trim();

    if (branch) uniqueBranches[branch] = true;
    if (role) uniqueRoles[role] = true;
    if (status === 'AKTIF' || status === 'ACTIVE') activeCount++;
    else if (status) nonActiveCount++;
    if (nikAtasan || nikAtasan2) hierarchyCount++;

    return {
      timestamp: r[idx.timestamp] || '',
      nik: String(r[idx.nik] || '').trim(),
      nama: String(r[idx.nama] || '').trim(),
      jabatan: String(r[idx.jabatan] || '').trim(),
      divisi: String(r[idx.divisi] || '').trim(),
      cabang: String(r[idx.cabang] || '').trim(),
      roleUser: String(r[idx.roleUser] || '').trim(),
      nikAtasan: nikAtasan,
      namaAtasan: String(r[idx.namaAtasan] || '').trim(),
      nikAtasan2: nikAtasan2,
      namaAtasan2: String(r[idx.namaAtasan2] || '').trim(),
      email: String(r[idx.email] || '').trim(),
      wa: String(r[idx.wa] || '').trim(),
      statusKaryawan: String(r[idx.statusKaryawan] || '').trim()
    };
  });

  return {
    headers: headers,
    rows: normalizedRows,
    indexMap: idx,
    summary: {
      total: normalizedRows.length,
      active: activeCount,
      nonActive: nonActiveCount,
      branches: Object.keys(uniqueBranches).length,
      roles: Object.keys(uniqueRoles).length,
      hasHierarchy: hierarchyCount
    }
  };
}

function getEmployeeDirectoryData() {
  var rows = getEmployeeMasterData();
  if (!rows || rows.length === 0) {
    return { headers: [], records: [], byNik: {}, byName: {} };
  }

  var headers = rows[0] || [];
  var normalizedHeaders = headers.map(function(h) {
    return String(h || '').toLowerCase();
  });

  function headerIndex(candidates, fallback) {
    for (var i = 0; i < normalizedHeaders.length; i++) {
      var h = normalizedHeaders[i];
      var found = candidates.some(function(c) { return h.indexOf(c) !== -1; });
      if (found) return i;
    }
    return fallback;
  }

  var idx = {
    timestamp: headerIndex(['timestamp', 'waktu'], 0),
    nik: headerIndex(['nik', 'npk'], 1),
    nama: headerIndex(['nama karyawan', 'nama'], 2),
    jabatan: headerIndex(['jabatan', 'posisi', 'position'], 3),
    divisi: headerIndex(['divisi', 'departemen', 'department'], 4),
    cabang: headerIndex(['cabang', 'lokasi kerja', 'branch'], 5),
    role: headerIndex(['role user', 'role'], 6),
    nikAtasan: headerIndex(['nik atasan langsung', 'atasan langsung'], 7),
    namaAtasan: headerIndex(['nama atasan langsung'], 8),
    nikAtasan2: headerIndex(['nik atasan dari atasan langsung'], 9),
    namaAtasan2: headerIndex(['nama atasan dari atasan langsung'], 10),
    email: headerIndex(['email karyawan', 'email'], 11),
    wa: headerIndex(['no wa', 'wa', 'whatsapp'], 12),
    status: headerIndex(['status karyawan', 'status'], 13)
  };

  var records = [];
  var byNik = {};
  var byName = {};

  rows.slice(1).forEach(function(r) {
    var nik = String((r && r[idx.nik]) || '').trim();
    var nama = String((r && r[idx.nama]) || '').trim();
    if (!nik && !nama) return;

    var rec = {
      timestamp: r[idx.timestamp] || '',
      nik: nik,
      nama: nama,
      jabatan: String((r && r[idx.jabatan]) || '').trim(),
      divisi: String((r && r[idx.divisi]) || '').trim(),
      cabang: String((r && r[idx.cabang]) || '').trim(),
      roleUser: String((r && r[idx.role]) || '').trim(),
      nikAtasanLangsung: String((r && r[idx.nikAtasan]) || '').trim(),
      namaAtasanLangsung: String((r && r[idx.namaAtasan]) || '').trim(),
      nikAtasan2: String((r && r[idx.nikAtasan2]) || '').trim(),
      namaAtasan2: String((r && r[idx.namaAtasan2]) || '').trim(),
      email: String((r && r[idx.email]) || '').trim(),
      noWa: String((r && r[idx.wa]) || '').trim(),
      statusKaryawan: String((r && r[idx.status]) || '').trim()
    };

    records.push(rec);
    if (nik) byNik[nik] = rec;
    if (nama) byName[nama.toUpperCase()] = rec;
  });

  return {
    headers: headers,
    records: records,
    byNik: byNik,
    byName: byName
  };
}

function getUserRegistrationReferenceData() {
  var directory = getEmployeeDirectoryData();
  var records = directory.records || [];
  var uniqueBranches = {};

  var employees = records
    .filter(function(rec) {
      var status = String(rec.statusKaryawan || '').toUpperCase().trim();
      return status === '' || status === 'AKTIF' || status === 'ACTIVE';
    })
    .map(function(rec) {
      var cab = String(rec.cabang || '').trim();
      if (cab) uniqueBranches[cab.toUpperCase()] = cab;
      return {
        nik: String(rec.nik || '').trim(),
        nama: String(rec.nama || '').trim(),
        jabatan: String(rec.jabatan || '').trim(),
        cabang: cab
      };
    })
    .filter(function(rec) { return rec.nik && rec.nama; });

  var branches = Object.keys(uniqueBranches)
    .sort()
    .map(function(k) { return uniqueBranches[k]; });

  return {
    employees: employees,
    branches: branches
  };
}

function registerUserId(payload) {
  payload = payload || {};
  var nik = String(payload.nik || '').trim();
  var nama = String(payload.nama || '').trim();
  var jabatan = String(payload.jabatan || '').trim();
  var cabang = String(payload.cabang || '').trim();
  var password = String(payload.password || '');
  var rePassword = String(payload.rePassword || '');

  if (!nik || !nama || !jabatan || !cabang) return { ok: false, message: 'NIK, Nama, Jabatan, dan Cabang wajib diisi.' };
  if (!password || !rePassword) return { ok: false, message: 'Password dan Re Password wajib diisi.' };
  if (password !== rePassword) return { ok: false, message: 'Password dan Re Password tidak sama.' };
  if (password.length < 6) return { ok: false, message: 'Password minimal 6 karakter.' };

  var ref = getUserRegistrationReferenceData();
  var matched = (ref.employees || []).find(function(rec) {
    return String(rec.nik || '').trim() === nik && String(rec.nama || '').trim().toUpperCase() === nama.toUpperCase();
  });
  if (!matched) return { ok: false, message: 'NIK dan Nama tidak cocok dengan DB_Karyawan.' };

  var roleApp = deriveRoleFromJabatan(jabatan);
  var ss = getDB();
  var sh = ss.getSheetByName('User_ID') || ss.insertSheet('User_ID');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['NIK', 'NAMA', 'CABANG', 'PASSWORD', 'ROLE_APP', 'STATUS_USER', 'UPDATED_AT']);
  }

  var data = sh.getDataRange().getValues();
  var hdr = data[0] || [];
  var idxNik = 0;
  var idxNama = 1;
  var idxCab = 2;
  var idxPass = 3;
  var idxRole = 4;
  var idxStatus = 5;
  var idxUpdatedAt = 6;

  // Backward compatibility: jika header lama (email,nama,...), tetap upayakan update by nik di kolom 0.
  var targetRow = -1;
  for (var r = 1; r < data.length; r++) {
    if (String((data[r] && data[r][idxNik]) || '').trim() === nik) {
      targetRow = r + 1;
      break;
    }
  }

  var rowValues = [nik, nama, cabang, password, roleApp, 'AKTIF', new Date()];
  if (targetRow > 0) {
    sh.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sh.appendRow(rowValues);
  }

  return { ok: true, message: 'User ID berhasil disimpan.', roleApp: roleApp };
}

function deriveRoleFromJabatan(jabatan) {
  var j = String(jabatan || '').toUpperCase();
  if (j.indexOf('HEAD') !== -1) return 'HEAD';
  if (j.indexOf('BRANCH MANAGER') !== -1) return 'BM';
  if (j.indexOf('SUPERVISOR') !== -1) return 'SUPERVISOR';
  return 'STAFF';
}

function saveCoachingAction(obj) {
  var ss = getDB();
  var sh = ss.getSheetByName('Raw_Coaching') || ss.insertSheet('Raw_Coaching');
  var id = "CHG-" + new Date().getTime();

  var header = [
    "ID",
    "Tanggal",
    "Cabang",
    "Nama",
    "Root Cause",
    "Topic",
    "Target",
    "Status",
    "ParentID",
    "BM/RH",
    "Region",
    "LastPerf",
    "NIK",
    "Jabatan",
    "Satuan Target",
    "Target Date",
    "How",
    "Result",
    "Feedback"
  ];

  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
  }

  sh.appendRow([
    id,
    new Date(),
    obj.cabang || "",
    obj.staff || "",
    obj.root || "",
    obj.topic || "",
    obj.target || "",
    obj.status || "ONGOING",
    obj.parentId || "",
    obj.bm || "",
    obj.region || "",
    obj.lastPerf || "",
    obj.nik || "",
    obj.jabatan || "",
    obj.satuanTarget || "",
    obj.targetDate || "",
    obj.how || "",
    obj.result || "",
    obj.feedback || ""
  ]);
  return "Berhasil";
}

/** Kolom laporan: sama urutannya dengan header Raw_Coaching (untuk ekspor Excel/PDF). */
function getCoachingReportColumnKeys() {
  return [
    "ID",
    "Tanggal",
    "Cabang",
    "Nama",
    "Root Cause",
    "Topic",
    "Target",
    "Status",
    "ParentID",
    "BM/RH",
    "Region",
    "LastPerf",
    "NIK",
    "Jabatan",
    "Satuan Target",
    "Target Date",
    "How",
    "Result",
    "Feedback"
  ];
}

/**
 * Ambil email atasan per cabang dari DB_Region / MASTER_REGION.
 * Tambahkan kolom header yang mengandung "email" (mis. EMAIL_BM, EMAIL ATASAN).
 */
function resolveSupervisorEmailForBranch(cabangUpper) {
  var key = String(cabangUpper || "").trim().toUpperCase();
  if (!key) return "";
  var ss = getDB();
  var sh = ss.getSheetByName("DB_Region") || ss.getSheetByName("MASTER_REGION");
  if (!sh || sh.getLastRow() < 2) return "";
  var rows = sh.getDataRange().getValues();
  var hdr = rows[0];
  var idxCab = -1;
  var idxEmail = -1;
  for (var c = 0; c < hdr.length; c++) {
    var lab = String(hdr[c] || "").toLowerCase();
    if (lab.indexOf("cabang") !== -1 || lab.indexOf("branch") !== -1) idxCab = c;
    if (lab.indexOf("email") !== -1) idxEmail = c;
  }
  if (idxCab < 0) idxCab = 0;
  if (idxEmail < 0) return "";
  for (var i = 1; i < rows.length; i++) {
    var cab = String(rows[i][idxCab] || "").trim().toUpperCase();
    if (cab === key) {
      var em = String(rows[i][idxEmail] || "").trim();
      if (em.indexOf("@") !== -1) return em;
    }
  }
  return "";
}

/**
 * Kirim pengingat coaching ke email atasan cabang (jika kolom email di sheet region terisi),
 * atau ke email user aktif sebagai fallback.
 */
function sendCoachingReminderEmail(payload) {
  payload = payload || {};
  var cabang = String(payload.cabang || "").trim();
  var subject = String(payload.subject || "Pengingat Coaching BPR KS").trim();
  var body = String(payload.body || "").trim();
  if (!body) return { ok: false, cabang: cabang, to: "", message: "Isi pesan kosong." };

  var to = resolveSupervisorEmailForBranch(cabang.toUpperCase());
  if (!to) {
    try {
      to = Session.getActiveUser().getEmail();
    } catch (e) {
      to = "";
    }
  }
  if (!to) {
    return {
      ok: false,
      cabang: cabang,
      to: "",
      message: "Tidak ada alamat email (isi kolom Email di DB_Region/MASTER_REGION atau login Google)."
    };
  }

  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body
    });
    return { ok: true, cabang: cabang, to: to, message: "Email terkirim." };
  } catch (e) {
    return { ok: false, cabang: cabang, to: to, message: e.message || String(e) };
  }
}

/** Kirim beberapa pengingat sekaligus (satu email per item, mis. per cabang). */
function sendCoachingReminderBatch(items) {
  items = items || [];
  var results = [];
  for (var i = 0; i < items.length; i++) {
    results.push(sendCoachingReminderEmail(items[i] || {}));
  }
  return results;
}
