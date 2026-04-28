/**
 * OWNER COMMAND CENTER - BPR KS
 * Versi: 181.0 (Base) + Add-on Coaching Interactive Table
 */

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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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

function getDashboardData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("db_v181_no_tele");
  if (cached) return JSON.parse(cached);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shM = ss.getSheetByName('DB_Karyawan') || ss.getSheetByName('Master_Karyawan');
  var shC = ss.getSheetByName('Raw_Coaching') || ss.insertSheet('Raw_Coaching');
  var shP = ss.getSheetByName('Raw_Coaching_Progress') || ss.insertSheet('Raw_Coaching_Progress');
  
  return {
    master: shM ? shM.getDataRange().getValues() : [],
    coaching: shC.getDataRange().getValues(),
    progress: shP.getDataRange().getValues()
  };
}

function saveCoachingAction(obj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Raw_Coaching');
  var id = "CHG-" + new Date().getTime();
  
  // Header: ID, Tanggal, Cabang, Nama, Root, Topic, Target, Status, ParentID, BM, Region, LastPerf
  sh.appendRow([
    id, 
    new Date(), 
    obj.cabang, 
    obj.staff, 
    obj.root, 
    obj.topic, 
    obj.target, 
    "ONGOING", 
    obj.parentId || "", 
    obj.bm || "", 
    obj.region || "", 
    obj.lastPerf || ""
  ]);
  return "Berhasil";
}
