/**
 * API Order Meals PT SSB (Google Apps Script, container-bound ke spreadsheet respons).
 * Langkah: 1) jalankan setupSheets() sekali, 2) Deploy sebagai Web app
 * (Execute as: Me, Who has access: Anyone), 3) tempel URL /exec ke index.html.
 */
const TZ = "Asia/Makassar";
const SH_INPUT = "Input_App", SH_PIN = "Admin_Departemen", SH_TITIK = "Titik_Departemen",
      SH_CFG = "Config", SH_MASTER = "Master";
const MAX_LINES = 150, MAX_QTY = 999;
const DAPUR = "DAPUR";   // nama khusus untuk login monitoring tim dapur (di sheet Admin_Departemen)

// [departemen, titik, peruntukan] diambil dari file Summary 24 Nov 2026. Silakan rapikan di sheet Titik_Departemen.
const SEED_TITIK = [
  ["Brimob", "Pos 4 (Jetty)", "Brimob"],
  ["Brimob", "Mess Brimob", "Brimob"],
  ["Eksplorasi", "Office Eksplorasi", "Supervisor"],
  ["Eksplorasi", "Office Eksplorasi", "Staff"],
  ["Engineering", "MCC Shelter", "MCC Dispatcher"],
  ["Engineering", "Office Engineering", "SPT Engineering"],
  ["Engineering", "Office Engineering", "SPV Short Term"],
  ["Engineering", "Office Engineering", "SPV Long Term"],
  ["Engineering", "Office Engineering", "SPV MCC"],
  ["Engineering", "Office Engineering", "FM PIT Control"],
  ["Engineering", "Office Engineering", "FM MCC"],
  ["Engineering", "Office Engineering", "Jr. Mineplan"],
  ["Engineering", "MCC Shelter", "FM PIT Control"],
  ["Engineering", "Mess HOD", "SPT Engineering"],
  ["Environment", "Mess HOD", "Superintendent"],
  ["Environment", "Office Environment", "Supervisor"],
  ["Environment", "Office Environment", "Foreman"],
  ["Environment", "Office Environment", "Crew"],
  ["Environment", "Pos Security Baru Jetty", "Operator Long Arm dan Pengawas"],
  ["Environment", "Nursery Baru", "Crew"],
  ["Klinik", "Klinik", "Dokter"],
  ["Klinik", "Klinik", "Perawat"],
  ["Logistik", "Kamar Pa Bedu", "SPT, SPV dan Foreman"],
  ["Logistik", "Loket Logistik", "Admin Shift 1"],
  ["Logistik", "Loket Logistik", "Crew Shift 1"],
  ["Logistik", "Loket Logistik", "BBC Solar"],
  ["Logistik", "Mess SANY Depan Logistik", "Vendor SANY"],
  ["Logistik", "Mess UT Depan Logistik", "Vendor United Tractor"],
  ["Logistik", "Kantin", "Crew Shift 2"],
  ["Logistik", "Mess Perempuan", "Admin Shift 2"],
  ["Logistik", "Loket Logistik", "Crew Shift 2"],
  ["Management", "Mess HOD", "KTT, WAKTT dan DPM Pak Budhi (masing-masing 1)"],
  ["Management", "Mess Tarik", "Chief dan Wakil Chief Security (masing-masing 1)"],
  ["Management", "Mess Pengawas (Mess Bawah)", "Danru Security"],
  ["Management", "Office KTT", "Staff Officer KTT"],
  ["Management", "Pos 2", "Security"],
  ["Management", "Kantin", "Security"],
  ["Management", "Pos 3", "Security"],
  ["Management", "Pos 4", "Security"],
  ["OHS", "Office OHS", "Supervisor"],
  ["OHS", "Office OHS", "Admin"],
  ["OHS", "Office OHS", "Foreman, Patrol dan Crew (termasuk SPL saat Breakfast)"],
  ["OHS", "Mess Putri", "Admin"],
  ["Plant", "Mess Bawah", "Tim Plant"],
  ["Plant", "Base Tire", "Tim Tire"],
  ["Plant", "Office Plant", "Pak Eko"],
  ["Plant", "Office Plant", "Tim Office Plant"],
  ["Plant", "Pit Stop", "Tim Pit Stop"],
  ["Plant", "PLTD Induk", "Tim Electric"],
  ["Plant", "Welder", "Tim Welder"],
  ["Plant", "Workshop", "Tim Workshop"],
  ["Plant", "Jetty", "Tim Tire Jetty"],
  ["Plant", "Jetty", "Mekanik Jetty"],
  ["Plant", "Jetty", "Tim Electric"],
  ["Plant", "Base Welder", "Tim Welder"],
  ["Produksi & MHR", "Elang 1", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Elang 2", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Kantin Lama", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Chef", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Garuda A", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Garuda B", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess HOD", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Hutan", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Klinik", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Langit", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Laundry", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Michat", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Putri", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Voly", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Mess Workshop", "Tim Produksi & MHR"],
  ["Produksi & MHR", "Kantin", "Penghuni Mess Atas"],
  ["Produksi & MHR", "Kantin", "Penghuni Mess Pengawas"],
  ["Produksi & MHR", "Jalur MHR", "Operator unit dan Pengawas"],
  ["Produksi & MHR", "Konstruksi EFO", "DT Hauling Waste, Quarry, Bulldozer dan Checker"],
  ["Produksi & MHR", "Pos Pantau Jetty", "Operator Barging dan Pengawas"],
  ["Produksi & MHR", "Pos Sample House Jetty", "DT Barging dan Checker"],
  ["Produksi & MHR", "Pos Sample House Long Hauling", "DT Long Hauling dan Checker"],
  ["Produksi & MHR", "Office Produksi", "Tim Produksi"],
  ["Produksi & MHR", "Shelter", "Pengawas, EXC Loading, EXC Trimming dan Support"],
  ["Produksi & MHR", "Fleet kiri dan kanan", "Operator, Pengawas dan Checker"],
  ["Produksi & MHR", "Sample House Tahisi 2", "DT Produksi dan Checker"],
  ["QAQC", "Pos Sampel ETO (gerbang masuk ETO)", "Grade Control ETO"],
  ["QAQC", "Pos Sampel EFO / Jetty", "Anindya A"],
  ["QAQC", "Mess Bawah (KKB)", "Anindya B"],
  ["QAQC", "Office QAQC", "Supervisor QAQC"],
  ["QAQC", "Office QAQC", "Supervisor Mine Geos"],
  ["QAQC", "Office QAQC", "Foreman dan Analis"],
  ["QAQC", "Office QAQC", "Grade Control EFO"],
  ["QAQC", "Office QAQC", "Helper Preparasi"],
  ["QAQC", "Sample House Tahisi 2", "QAQC PIT"],
  ["QAQC", "Pos Sampel EFO / Jetty", "Grade Control EFO"],
  ["QAQC", "Office QAQC", "Analis"],
  ["Survey", "Office Survey", "Supervisor"],
  ["Survey", "Office Survey", "Crew"],
  ["Survey", "Kantin", "Crew"],
  ["TRD", "Mess Foreman TRD", "Foreman TRD"],
  ["TRD", "Mess HOD", "Pak Isfardiansyah (SPT TRD)"]
];

/* ---------- Setup (jalankan sekali dari editor) ---------- */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty("SS_ID", ss.getId());
  ss.setSpreadsheetTimeZone(TZ);
  const master = ss.getSheetByName(SH_MASTER);
  if (!master) throw new Error("Tab 'Master' tidak ditemukan. Impor template v3 dulu.");

  let sh = ss.getSheetByName(SH_INPUT);
  if (!sh) {
    sh = ss.insertSheet(SH_INPUT);
    sh.getRange(1, 1, 1, 12).setValues([["Timestamp", "Nama Pemesan", "Departemen", "Tanggal Makan",
      "Periode Makan", "Titik Pengantaran / Layanan", "Peruntukan", "Packmeal", "Canteen", "Catatan",
      "Kunci Pesanan", "Sumber"]]).setFontWeight("bold");
    sh.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm");
    sh.getRange("D:D").setNumberFormat("dd/mm/yyyy");
  }

  sh = ss.getSheetByName(SH_PIN);
  if (!sh) {
    sh = ss.insertSheet(SH_PIN);
    const depts = master.getRange("A3:A27").getValues().map(r => String(r[0]).trim()).filter(String);
    const used = {};
    const rows = depts.map(d => {
      let p;
      do { p = String(Math.floor(100000 + Math.random() * 900000)); } while (used[p]);
      used[p] = 1;
      return [d, p];
    });
    sh.getRange("B:B").setNumberFormat("@");
    sh.getRange(1, 1, 1, 2).setValues([["Departemen", "PIN"]]).setFontWeight("bold");
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
    sh.getRange("D1").setValue("Sheet ini rahasia. Bagikan PIN hanya ke admin departemen terkait.");
  }

  sh = ss.getSheetByName(SH_TITIK);
  if (!sh) {
    sh = ss.insertSheet(SH_TITIK);
    sh.getRange(1, 1, 1, 3).setValues([["Departemen", "Titik Pengantaran / Layanan", "Peruntukan"]]).setFontWeight("bold");
    sh.getRange(2, 1, SEED_TITIK.length, 3).setValues(SEED_TITIK);
    sh.getRange("E1").setValue("Daftar titik langganan yang muncul otomatis di aplikasi. Titik harus sama persis dengan Master kolom F.");
  }

  sh = ss.getSheetByName(SH_CFG);
  if (!sh) {
    sh = ss.insertSheet(SH_CFG);
    sh.getRange(1, 1, 1, 4).setValues([["Periode", "Hari sebelum makan", "Jam batas", "Menit batas"]]).setFontWeight("bold");
    const per = masterPeriode_(ss);
    sh.getRange(2, 1, per.length, 4).setValues(per.map(p => [p, 0, 0, 0]));
    sh.getRange("F1").setValue("0, 0, 0 = pesanan untuk suatu tanggal bisa dibuat atau diubah sampai tanggal berganti (pukul 00:00 hari makan). Contoh 1, 14, 0 = ditutup H-1 pukul 14:00.");
  }
  migrasiV2();
  Logger.log("Setup selesai. PIN tiap departemen dan PIN dapur ada di sheet " + SH_PIN + ".");
}

/**
 * Pembaruan data v2 (aman dijalankan berulang):
 * - Departemen Klinik dihapus; titik Klinik (Dokter, Perawat) pindah ke OHS.
 * - Titik Produksi "Kantin (Penghuni Mess Atas/Pengawas)" diganti menjadi
 *   "Mess Atas Non Staff (Latimojong A-F)" dan "Mess Atas Staff (Semeru A&B)".
 * - Departemen GA ditambahkan (titik: Kantor GA, Kantor Utama).
 * - Titik "Lokasi lain (tulis di catatan)" untuk order khusus.
 * - PIN untuk GA dan login Monitoring Dapur dibuat.
 */
function migrasiV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty("SS_ID", ss.getId());
  const master = ss.getSheetByName(SH_MASTER);
  if (!master) throw new Error("Tab 'Master' tidak ditemukan.");
  const cmp = (a, b) => String(a).localeCompare(String(b));

  // 1) Master: departemen
  let depts = master.getRange("A3:A27").getValues().map(r => String(r[0]).trim()).filter(String)
    .filter(d => d !== "Klinik");
  if (depts.indexOf("GA") < 0) depts.push("GA");
  depts.sort(cmp);
  writeColumn_(master, 3, 1, 25, depts);

  // 2) Master: area
  let areas = master.getRange("B3:B27").getValues().map(r => String(r[0]).trim()).filter(String);
  if (areas.indexOf("Lainnya") < 0) areas.push("Lainnya");
  areas.sort(cmp);
  writeColumn_(master, 3, 2, 25, areas);

  // 3) Master: pasangan Area dan Titik
  const pairs = master.getRange("E3:F122").getValues()
    .filter(r => String(r[1]).trim() !== "").map(r => [String(r[0]).trim(), String(r[1]).trim()]);
  const have = {};
  pairs.forEach(p => { have[p[1]] = 1; });
  [["Mess", "Mess Atas Non Staff (Latimojong A-F)"], ["Mess", "Mess Atas Staff (Semeru A&B)"],
   ["Perkantoran", "Kantor GA"], ["Perkantoran", "Kantor Utama"],
   ["Lainnya", "Lokasi lain (tulis di catatan)"]].forEach(p => { if (!have[p[1]]) pairs.push(p); });
  const out = [];
  for (let i = 0; i < 120; i++) out.push(i < pairs.length ? pairs[i] : ["", ""]);
  master.getRange(3, 5, 120, 2).setValues(out);

  // 4) Admin_Departemen: hapus Klinik, tambah GA dan DAPUR
  const sp = ss.getSheetByName(SH_PIN);
  let pv = sp.getDataRange().getValues().slice(1).filter(r => String(r[0]).trim())
    .map(r => [String(r[0]).trim(), String(r[1]).trim()]);
  pv = pv.filter(r => r[0] !== "Klinik");
  const used = {};
  pv.forEach(r => { used[r[1]] = 1; });
  [ "GA", DAPUR ].forEach(name => {
    if (!pv.some(r => r[0] === name)) {
      let p;
      do { p = String(Math.floor(100000 + Math.random() * 900000)); } while (used[p]);
      used[p] = 1; pv.push([name, p]);
    }
  });
  sp.getRange("B:B").setNumberFormat("@");
  sp.getRange(2, 1, sp.getMaxRows() - 1, 2).clearContent();
  sp.getRange(2, 1, pv.length, 2).setValues(pv);

  // 5) Titik_Departemen
  const st = ss.getSheetByName(SH_TITIK);
  let rows = st.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => [String(r[0]).trim(), String(r[1]).trim(), String(r[2] || "").trim()]);
  rows = rows.filter(r => r[0] !== "Klinik");
  rows = rows.map(r => {
    if (r[0] === "Produksi & MHR" && r[1] === "Kantin") {
      if (r[2] === "Penghuni Mess Atas") return [r[0], "Mess Atas Non Staff (Latimojong A-F)", r[2]];
      if (r[2] === "Penghuni Mess Pengawas") return [r[0], "Mess Atas Staff (Semeru A&B)", r[2]];
    }
    return r;
  });
  const ensure = (d, t, p) => { if (!rows.some(r => r[0] === d && r[1] === t && r[2] === p)) rows.push([d, t, p]); };
  ensure("OHS", "Klinik", "Dokter"); ensure("OHS", "Klinik", "Perawat");
  ensure("GA", "Kantor GA", "Staff GA"); ensure("GA", "Kantor Utama", "Staff GA");
  rows = rows.map((r, i) => [r, i]).sort((a, b) => cmp(a[0][0], b[0][0]) || a[1] - b[1]).map(x => x[0]);
  st.getRange(2, 1, st.getMaxRows() - 1, 3).clearContent();
  st.getRange(2, 1, rows.length, 3).setValues(rows);

  Logger.log("Migrasi v2 selesai. Lihat PIN GA dan DAPUR di sheet " + SH_PIN + ".");
}

function writeColumn_(sh, row, col, height, values) {
  const v = [];
  for (let i = 0; i < height; i++) v.push([i < values.length ? values[i] : ""]);
  sh.getRange(row, col, height, 1).setValues(v);
}

/* ---------- Web app ---------- */
function doGet() { return json_({ ok: true, msg: "API Order Meals aktif" }); }

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const dept = deptByPin_(req.pin);
    if (!dept) return json_({ ok: false, error: "PIN salah." });
    if (dept === DAPUR) {                      // login tim dapur: hanya baca
      if (req.action === "login") return json_(loginDapur_());
      if (req.action === "monitor") return json_(monitor_(req.tanggal));
      return json_({ ok: false, error: "Akses ditolak." });
    }
    switch (req.action) {
      case "login":  return json_(login_(dept));
      case "load":   return json_(load_(dept, req.tanggal));
      case "last":   return json_(last_(dept, String(req.periode || ""), req.tanggal));
      case "submit": return json_(submit_(dept, req));
      default:       return json_({ ok: false, error: "Aksi tidak dikenal." });
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function login_(dept) {
  const v = ss_().getSheetByName(SH_TITIK).getDataRange().getValues();
  const langganan = [];
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === dept && String(v[i][1]).trim()) {
      langganan.push({ titik: String(v[i][1]).trim(), peruntukan: String(v[i][2] || "").trim() });
    }
  }
  return { ok: true, role: "admin", departemen: dept, periode: masterPeriode_(), titik: masterTitik_(), langganan: langganan };
}

function loginDapur_() {
  return { ok: true, role: "dapur", departemen: "Monitoring Dapur", periode: masterPeriode_() };
}

/** Semua pesanan satu tanggal, dikelompokkan per periode lalu per departemen (khusus tim dapur). */
function monitor_(tanggal) {
  checkTanggal_(tanggal);
  const now = new Date();
  const pers = masterPeriode_();
  const orders = {};
  pers.forEach(p => {
    const c = cutoffFor_(tanggal, p);
    orders[p] = { cutoff: fmtCutoff_(c), locked: now > c, depts: {} };
  });
  const area = {};
  masterTitik_().forEach(t => { area[t.titik] = t.area; });
  const sh = ss_().getSheetByName(SH_INPUT);
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 12).getValues().forEach(r => {
      const k = String(r[10]).split("|");
      if (k[1] !== tanggal || !orders[k[2]]) return;
      const bucket = orders[k[2]].depts;
      const e = bucket[k[0]] || (bucket[k[0]] = { nama: "", diperbarui: "", lines: [] });
      e.nama = String(r[1]);
      e.diperbarui = Utilities.formatDate(new Date(r[0]), TZ, "dd MMM HH:mm");
      const ln = lineOut_(r);
      ln.area = area[ln.titik] || "";
      e.lines.push(ln);
    });
  }
  const deptList = ss_().getSheetByName(SH_MASTER).getRange("A3:A27").getValues()
    .map(r => String(r[0]).trim()).filter(String);
  return { ok: true, tanggal: tanggal, departemen: deptList, periode: pers, orders: orders,
           diperbaruiPada: Utilities.formatDate(now, TZ, "HH:mm:ss") };
}

function load_(dept, tanggal) {
  checkTanggal_(tanggal);
  const now = new Date();
  const orders = {};
  masterPeriode_().forEach(p => {
    const c = cutoffFor_(tanggal, p);
    orders[p] = { lines: [], locked: now > c, cutoff: fmtCutoff_(c) };
  });
  const sh = ss_().getSheetByName(SH_INPUT);
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 12).getValues().forEach(r => {
      const k = String(r[10]).split("|");
      if (k[0] === dept && k[1] === tanggal && orders[k[2]]) orders[k[2]].lines.push(lineOut_(r));
    });
  }
  return { ok: true, tanggal: tanggal, orders: orders };
}

function last_(dept, periode, tanggal) {
  checkTanggal_(tanggal);
  const sh = ss_().getSheetByName(SH_INPUT);
  const last = sh.getLastRow();
  let best = "", lines = [];
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 12).getValues().forEach(r => {
      const k = String(r[10]).split("|");
      if (k[0] !== dept || k[2] !== periode || !(k[1] < tanggal)) return;
      if (k[1] > best) { best = k[1]; lines = []; }
      if (k[1] === best) lines.push(lineOut_(r));
    });
  }
  return { ok: true, tanggal: best, lines: lines };
}

function submit_(dept, req) {
  checkTanggal_(req.tanggal);
  const nama = norm_(req.nama, 60);
  if (!nama) throw new Error("Nama pemesan wajib diisi.");
  const orders = Array.isArray(req.orders) ? req.orders : [];
  if (!orders.length) throw new Error("Tidak ada pesanan untuk dikirim.");

  const pers = masterPeriode_();
  const titikOk = {};
  masterTitik_().forEach(t => { titikOk[t.titik] = 1; });
  const tgl = Utilities.parseDate(req.tanggal, TZ, "yyyy-MM-dd");

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = ss_().getSheetByName(SH_INPUT);
    const now = new Date();
    const results = [];
    orders.forEach(o => {
      const periode = String(o.periode || "");
      try {
        if (pers.indexOf(periode) < 0) throw new Error("Periode tidak dikenal.");
        if (now > cutoffFor_(req.tanggal, periode)) throw new Error("Batas waktu pemesanan sudah lewat.");
        const lines = Array.isArray(o.lines) ? o.lines : [];
        if (lines.length > MAX_LINES) throw new Error("Terlalu banyak baris (maks " + MAX_LINES + ").");
        const key = [dept, req.tanggal, periode].join("|");
        const rows = [];
        lines.forEach(l => {
          const pm = qty_(l.packmeal), cn = qty_(l.canteen);
          if (pm + cn <= 0) return;
          const titik = String(l.titik || "").trim();
          if (!titikOk[titik]) throw new Error("Titik tidak ada di Master: " + titik);
          const per = norm_(l.peruntukan, 100);
          if (!per) throw new Error("Peruntukan kosong untuk titik " + titik);
          rows.push([now, nama, dept, tgl, periode, titik, per, pm, cn, norm_(l.catatan, 200), key, "App"]);
        });
        removeKey_(sh, key);                 // pesanan lama (dept + tanggal + periode) diganti
        if (rows.length) {
          const start = Math.max(sh.getLastRow(), 1) + 1;
          sh.getRange(start, 1, rows.length, 12).setValues(rows);
        }
        results.push({ periode: periode, ok: true, baris: rows.length });
      } catch (err) {
        results.push({ periode: periode, ok: false, error: String((err && err.message) || err) });
      }
    });
    return { ok: true, results: results };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Helper ---------- */
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty("SS_ID");
  if (!id) throw new Error("Server belum disiapkan. Jalankan setupSheets() dulu.");
  return SpreadsheetApp.openById(id);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function deptByPin_(pin) {
  pin = String(pin || "").trim();
  if (!/^\d{4,8}$/.test(pin)) return null;
  const v = ss_().getSheetByName(SH_PIN).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][1]).trim() === pin) return String(v[i][0]).trim();
  }
  return null;
}
function masterPeriode_(ss) {
  return (ss || ss_()).getSheetByName(SH_MASTER).getRange("C3:C6").getValues()
    .map(r => String(r[0]).trim()).filter(String);
}
function masterTitik_() {
  return ss_().getSheetByName(SH_MASTER).getRange("E3:F122").getValues()
    .filter(r => String(r[1]).trim() !== "")
    .map(r => ({ area: String(r[0]).trim(), titik: String(r[1]).trim() }));
}
function cutoffFor_(tanggal, periode) {
  const rows = ss_().getSheetByName(SH_CFG).getDataRange().getValues();
  let hari = 0, jam = 0, menit = 0;   // bawaan: pesanan terbuka sampai tanggal berganti
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === periode) {
      const h = Number(rows[i][1]), j = Number(rows[i][2]), m = Number(rows[i][3]);
      if (isFinite(h)) hari = h;
      if (isFinite(j)) jam = j;
      if (isFinite(m)) menit = m;
      break;
    }
  }
  const base = Utilities.parseDate(tanggal, TZ, "yyyy-MM-dd");
  return new Date(base.getTime() - hari * 86400000 + (jam * 60 + menit) * 60000);
}
function fmtCutoff_(d) {
  // batas tepat pukul 00:00 ditampilkan sebagai 23:59 hari sebelumnya agar tidak membingungkan
  if (Utilities.formatDate(d, TZ, "HH:mm") === "00:00") d = new Date(d.getTime() - 60000);
  return Utilities.formatDate(d, TZ, "dd MMM yyyy HH:mm");
}
function checkTanggal_(t) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(t || ""))) throw new Error("Format tanggal tidak valid.");
  const d = Utilities.parseDate(t, TZ, "yyyy-MM-dd");
  const now = new Date();
  if (d.getTime() > now.getTime() + 45 * 86400000) throw new Error("Tanggal terlalu jauh ke depan.");
  if (d.getTime() < now.getTime() - 400 * 86400000) throw new Error("Tanggal terlalu lama.");
}
function qty_(v) {
  const n = Number(v || 0);
  if (!isFinite(n) || n < 0 || n > MAX_QTY || Math.floor(n) !== n) {
    throw new Error("Jumlah harus bilangan bulat 0 sampai " + MAX_QTY + ".");
  }
  return n;
}
function norm_(s, max) {
  s = String(s == null ? "" : s).replace(/[\r\n\t]+/g, " ").trim();
  if (s.length > max) s = s.substring(0, max);
  if (/^[=+\-@]/.test(s)) s = " " + s;      // cegah teks terbaca sebagai rumus
  return s;
}
function lineOut_(r) {
  return { titik: String(r[5]), peruntukan: String(r[6]), packmeal: Number(r[7]) || 0,
           canteen: Number(r[8]) || 0, catatan: String(r[9] || "") };
}
function removeKey_(sh, key) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const keys = sh.getRange(2, 11, last - 1, 1).getValues();
  for (let i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i][0]) === key) sh.deleteRow(i + 2);
  }
}
