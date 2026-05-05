/**
 * OWNER COMMAND CENTER - BPR KS
 * Versi: 181.0 (Base) + Add-on Coaching Interactive Table
 */

var DB_ID = '1fKi00TOdtqwlyP6Uk20sGbJHrm_dsQ5frlkFm6jMAzE';

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
    var data = sh ? sh.getDataRange().getValues() : [];
    if (email.includes("yulius") || email === "") return { email: email, nama: "YULIUS PUJI LAKSONO", role: "Admin", cabang: "ALL" };
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

function parsePlafondRupiahForPipeline(val) {
  if (val == null || val === "") return 0;
  if (typeof val === "number" && !isNaN(val)) return Math.round(val);
  var s = String(val).trim().replace(/\s/g, "");
  if (!s) return 0;
  // format 1.234.567,89
  var decComma = s.match(/^([\d.]+),(\d{1,2})$/);
  if (decComma) {
    var intPart = decComma[1].replace(/\./g, "");
    return Math.round(parseFloat(intPart + "." + decComma[2]));
  }
  // format 1.234.567 atau 1234567
  var noSep = s.replace(/\./g, "").replace(/,/g, "");
  if (/^\d+$/.test(noSep)) return Math.round(parseInt(noSep, 10));
  var n = Number(s.replace(/\./g, "").replace(",", "."));
  return isNaN(n) ? 0 : Math.round(n);
}

function getRawPipelineColumnMap_(headerRow) {
  var map = { idxDate: 0, idxRawB: 1, idxDeb: 2, idxCab: 3, idxStat: 4, idxMix: 5, idxKep: 6, idxSal: 7, idxPla: 10, hasHeader: false };
  if (!headerRow || !headerRow.length) return map;
  var norm = headerRow.map(function(v) { return String(v || "").toLowerCase().trim(); });
  function findIdx(cands, fallback) {
    for (var i = 0; i < norm.length; i++) {
      var label = norm[i];
      if (!label) continue;
      for (var j = 0; j < cands.length; j++) {
        if (label.indexOf(cands[j]) !== -1) return i;
      }
    }
    return fallback;
  }
  var idxDeb = findIdx(["debitur", "nama debitur"], -1);
  var idxCab = findIdx(["cabang", "branch"], -1);
  var idxSal = findIdx(["sales", "nama sales", "marketing"], -1);
  var idxPla = findIdx(["plafond", "plafon", "limit"], -1);
  if (idxDeb >= 0 || idxCab >= 0 || idxSal >= 0 || idxPla >= 0) {
    map.hasHeader = true;
    map.idxDate = findIdx(["tgl", "tanggal", "date"], 0);
    map.idxRawB = findIdx(["aging", "serial", "msec", "hari"], 1);
    map.idxDeb = idxDeb >= 0 ? idxDeb : 2;
    map.idxCab = idxCab >= 0 ? idxCab : 3;
    map.idxStat = findIdx(["status"], 4);
    map.idxMix = findIdx(["mix", "produk", "product"], 5);
    map.idxKep = findIdx(["keputusan", "decision"], 6);
    map.idxSal = idxSal >= 0 ? idxSal : 7;
    map.idxPla = idxPla >= 0 ? idxPla : 10;
  }
  return map;
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
      var colMap = getRawPipelineColumnMap_(dataP[0] || []);
      var startRow = colMap.hasHeader ? 1 : 0;
      for (var pr = startRow; pr < dataP.length; pr++) {
        var r = dataP[pr];
        var debRaw = r[colMap.idxDeb];
        if (!r[colMap.idxDate] || !debRaw || String(debRaw).toLowerCase().indexOf("debitur") !== -1) continue;
        var rawB = r[colMap.idxRawB], msec;
        if (typeof rawB === 'number') msec = (rawB - 25569) * 86400 * 1000;
        else if (rawB instanceof Date) msec = rawB.getTime();
        else msec = new Date().getTime();
        var tglRaw = r[colMap.idxDate];
        var tglIn = (tglRaw instanceof Date) ? Utilities.formatDate(tglRaw, "GMT+7", "yyyy-MM-dd") : String(tglRaw).substring(0,10);
        var deb = String(r[colMap.idxDeb] || "").trim();
        var cab = String(r[colMap.idxCab] || "").trim();
        var stat = String(r[colMap.idxStat] || "").trim();
        var mix = String(r[colMap.idxMix] || "").trim();
        var kep = String(r[colMap.idxKep] || "").trim();
        var sal = String(r[colMap.idxSal] || "-").trim();
        var pla = parsePlafondRupiahForPipeline(r[colMap.idxPla]);
        if (cab && !isJunkBranch(cab)) {
          if (res.listCPipe.indexOf(cab) === -1) res.listCPipe.push(cab);
          if (sal !== "-" && res.listS.indexOf(sal) === -1) res.listS.push(sal);
        }
        res.pipeline.push([tglIn, msec, deb, stat, mix, kep, sal, pla, cab]);
      }
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
    var pipelineNumericCols = {};
    if (tipe == "7" && rows && rows.length) {
      var hdr = rows[0] || [];
      for (var hc = 0; hc < hdr.length; hc++) {
        var label = String(hdr[hc] || "").toLowerCase();
        if (label.indexOf("plafond") !== -1 || label.indexOf("plafon") !== -1 || label.indexOf("limit") !== -1) {
          pipelineNumericCols[hc] = true;
        }
      }
    }
    var clean = v => { if (!v || v === "-" || v === "" || v === 0) return 0; return parseFloat(v.toString().replace(/\./g, "").replace(/,/g, ".")) || 0; };
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]; if (!r[0] || r[0] === "CABANG" || r[0] === "TGL") continue;
      var row = [tgl];
      for (var c = 0; c < r.length; c++) { 
        var isNum = (tipe == "7" && ((pipelineNumericCols[c] === true) || (Object.keys(pipelineNumericCols).length === 0 && (c == 9 || c == 10)))) || (tipe != "7" && c >= 2); 
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
