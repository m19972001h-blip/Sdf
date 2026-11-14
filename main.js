const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require("electron");
const path = require("path"), fs = require("fs"), crypto = require("crypto");
const { execFile } = require("child_process");
const XLSX = require("xlsx"), initSqlJs = require("sql.js");

let win, db, SQL, cfg, cfgPath, dbPath;
const sha = t => crypto.createHash("sha256").update(String(t)).digest("hex");
const tick = () => new Promise(r => setImmediate(r));
const SCHEMA = `
CREATE TABLE IF NOT EXISTS files(id INTEGER PRIMARY KEY,name TEXT,path TEXT,rows INTEGER,at TEXT);
CREATE TABLE IF NOT EXISTS people(id INTEGER PRIMARY KEY,fid INTEGER,r INTEGER,c TEXT,name TEXT,norm TEXT,w1 TEXT,qada TEXT,sex TEXT);
CREATE INDEX IF NOT EXISTS i_norm ON people(norm);
CREATE INDEX IF NOT EXISTS i_w1 ON people(w1);
CREATE INDEX IF NOT EXISTS i_fid ON people(fid);
CREATE TABLE IF NOT EXISTS log(id INTEGER PRIMARY KEY,at TEXT,msg TEXT);`;

// ---------- توحيد النص العربي ----------
function norm(t) {
  let s = String(t).trim().replace(/[\u064B-\u065F\u0670\u0640]/g, "");
  s = s.replace(/[إأآٱ]/g, "ا").replace(/ى/g, "ي").replace(/ة/g, "ه").replace(/ؤ/g, "و").replace(/ئ/g, "ي");
  s = s.replace(/[^\u0600-\u06FFa-zA-Z0-9\s]/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
  return s.replace(/(^|\s)(عبد|ابو|ابي|ذو) (\S+)/g, "$1$2$3");
}
function lev(a, b) {
  if (a === b) return 0; const m = a.length, n = b.length; if (!m) return n; if (!n) return m;
  let p = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) { const c = [i]; for (let j = 1; j <= n; j++) c[j] = Math.min(p[j] + 1, c[j - 1] + 1, p[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); p = c; }
  return p[n];
}
const eq = (a, b) => a === b || (a.length >= 6 && b.length >= 6 && lev(a, b) <= 1);
function level(p, q) {
  let k = 0; const m = Math.min(4, p.length, q.length);
  while (k < m && eq(p[k], q[k])) k++;
  return [0, 0, 40, 60, 80][k];
}
const sexOf = v => { v = String(v).trim().toLowerCase(); return !v ? "" : /^(ذ|m)/.test(v) ? "M" : /^(ان|أن|إن|f)/.test(v) ? "F" : ""; };
const qadaOf = v => { const s = String(v).split(/[-–—،,\/]/).pop().replace(/\s+/g, " ").trim(); return s; };

// ---------- قاعدة البيانات (SQLite عبر sql.js) ----------
function q(sql, p = []) { const s = db.prepare(sql); s.bind(p); const o = []; while (s.step()) o.push(s.getAsObject()); s.free(); return o; }
function run(sql, p = []) { db.run(sql, p); }
function save() { fs.writeFileSync(dbPath, Buffer.from(db.export())); }
function log(msg) { run("INSERT INTO log(at,msg) VALUES(?,?)", [new Date().toISOString(), msg]); }
function loadCfg() {
  cfgPath = path.join(app.getPath("userData"), "config.json");
  try { cfg = JSON.parse(fs.readFileSync(cfgPath, "utf8")); } catch (e) { cfg = {}; }
  cfg.dir = cfg.dir || app.getPath("userData");
  cfg.hash = cfg.hash || sha("1891");
  if (cfg.must === undefined) cfg.must = true;
  try { fs.mkdirSync(cfg.dir, { recursive: true }); } catch (e) { cfg.dir = app.getPath("userData"); }
}
const saveCfg = () => fs.writeFileSync(cfgPath, JSON.stringify(cfg));
function loadDb() {
  dbPath = path.join(cfg.dir, "people.sqlite");
  db = fs.existsSync(dbPath) ? new SQL.Database(fs.readFileSync(dbPath)) : new SQL.Database();
  db.run(SCHEMA);
}
function W(o) {
  const c = [], p = [];
  if (o.fid) { c.push("p.fid=?"); p.push(+o.fid); }
  if (o.qada) { c.push("p.qada=?"); p.push(o.qada); }
  if (o.sex) { c.push("p.sex=?"); p.push(o.sex); }
  return [c, p];
}

// ---------- الحسابات ----------
ipcMain.handle("login", (_, { u, p }) => ({ ok: u === "fqar" && sha(p) === cfg.hash, must: cfg.must }));
ipcMain.handle("chpw", (_, { o, n }) => {
  if (sha(o) !== cfg.hash) return "كلمة المرور الحالية غير صحيحة";
  if (!n || n.length < 6 || n === "1891") return "كلمة المرور الجديدة ضعيفة";
  cfg.hash = sha(n); cfg.must = false; saveCfg(); return "ok";
});

// ---------- المعلومات والإحصاءات ----------
ipcMain.handle("info", () => ({
  dir: cfg.dir,
  files: q("SELECT * FROM files ORDER BY id DESC"),
  log: q("SELECT at,msg FROM log ORDER BY id DESC LIMIT 60"),
  qs: q("SELECT DISTINCT qada k FROM people WHERE qada!='' ORDER BY qada LIMIT 500").map(x => x.k)
}));
ipcMain.handle("stats", (_, fid) => {
  const w = fid ? " WHERE fid=" + (+fid) : "", a = fid ? " AND fid=" + (+fid) : "";
  return {
    total: q("SELECT COUNT(*) n FROM people" + w)[0].n,
    qada: q("SELECT qada k,COUNT(*) n FROM people WHERE qada!=''" + a + " GROUP BY qada ORDER BY n DESC LIMIT 12"),
    sex: q("SELECT sex k,COUNT(*) n FROM people WHERE sex!=''" + a + " GROUP BY sex")
  };
});

// ---------- استيراد ملفات Excel (غير محدود) مع شريط تقدم ----------
ipcMain.handle("import", async () => {
  const r = await dialog.showOpenDialog(win, { properties: ["openFile", "multiSelections"], filters: [{ name: "Excel", extensions: ["xlsx", "xls", "xlsm", "csv"] }] });
  if (r.canceled) return 0;
  const N = r.filePaths.length;
  for (let k = 0; k < N; k++) {
    const fp = r.filePaths[k], base = k / N, share = 1 / N;
    const prog = (p, m) => win.webContents.send("prog", { p: Math.round((base + share * p) * 100), m: m + " (" + (k + 1) + "/" + N + ")" });
    try {
      prog(0, "قراءة " + path.basename(fp)); await tick();
      const wb = XLSX.readFile(fp), ws = wb.Sheets[wb.SheetNames[0]];
      const rng = XLSX.utils.decode_range(ws["!ref"] || "A1");
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      prog(0.2, "تحليل الأعمدة"); await tick();
      const H = (rows[0] || []).map(x => String(x).trim());
      const find = (...ks) => { for (const key of ks) { const i = H.findIndex(h => h.includes(key)); if (i >= 0) return i; } return -1; };
      let ci = find("الاسم", "اسم", "name"); if (ci < 0) ci = 0;
      const qi = find("القضاء", "السكن", "العنوان", "المنطقة", "address"), si = find("الجنس", "النوع", "gender", "sex");
      run("INSERT INTO files(name,path,rows,at) VALUES(?,?,?,?)", [path.basename(fp), fp, 0, new Date().toISOString()]);
      const fid = q("SELECT last_insert_rowid() id")[0].id;
      const st = db.prepare("INSERT INTO people(fid,r,c,name,norm,w1,qada,sex) VALUES(?,?,?,?,?,?,?,?)");
      const col = XLSX.utils.encode_col(rng.s.c + ci); let cnt = 0;
      for (let i = 1; i < rows.length; i += 2000) {
        db.run("BEGIN");
        for (let j = i; j < Math.min(i + 2000, rows.length); j++) {
          const v = rows[j][ci]; if (v === "" || v === undefined) continue;
          const n = norm(v); if (!n) continue;
          st.run([fid, rng.s.r + j + 1, col, String(v).trim(), n, n.split(" ")[0], qi >= 0 ? qadaOf(rows[j][qi]) : "", si >= 0 ? sexOf(rows[j][si]) : ""]); cnt++;
        }
        db.run("COMMIT");
        prog(0.2 + 0.8 * Math.min(1, i / rows.length), "إدخال الأسماء"); await tick();
      }
      st.free(); run("UPDATE files SET rows=? WHERE id=?", [cnt, fid]);
      log("استيراد الملف " + path.basename(fp) + " (" + cnt + " اسم)");
    } catch (e) { log("فشل استيراد " + path.basename(fp) + ": " + e.message); }
  }
  save(); return N;
});
ipcMain.handle("remove", (_, id) => {
  const f = q("SELECT name FROM files WHERE id=?", [id])[0];
  run("DELETE FROM people WHERE fid=?", [id]); run("DELETE FROM files WHERE id=?", [id]);
  log("حذف الملف " + (f ? f.name : id)); save(); return 1;
});

// ---------- البحث الذكي ----------
ipcMain.handle("search", (_, o) => {
  const [c, p] = W(o);
  for (const t of norm(o.q || "").split(" ").filter(Boolean)) { c.push("p.norm LIKE ?"); p.push("%" + t + "%"); }
  const w = c.length ? " WHERE " + c.join(" AND ") : "";
  const total = q("SELECT COUNT(*) n FROM people p" + w, p)[0].n;
  const rows = q("SELECT p.fid,p.r,p.c,p.name,p.qada,p.sex,f.name fn FROM people p JOIN files f ON f.id=p.fid" + w + " ORDER BY p.name LIMIT 200 OFFSET ?", [...p, o.off || 0]);
  return { total, rows };
});

// ---------- المقارنة (40 / 60 / 80) ----------
ipcMain.handle("cmp", (_, o) => {
  const get = fid => { const [c, p] = W({ fid, qada: o.qada, sex: o.sex }); return q("SELECT p.id,p.fid,p.r,p.c,p.name,p.norm,p.qada,p.sex FROM people p WHERE " + c.join(" AND "), p).map(r => (r.p = r.norm.split(" "), r)); };
  const A = get(o.a), B = o.a === o.b ? A : get(o.b), idx = new Map();
  B.forEach(r => { if (!idx.has(r.p[0])) idx.set(r.p[0], []); idx.get(r.p[0]).push(r); });
  const names = {}; q("SELECT id,name FROM files").forEach(f => names[f.id] = f.name);
  const out = [];
  for (const x of A) for (const y of idx.get(x.p[0]) || []) {
    if (x.id === y.id || (o.a === o.b && x.id > y.id)) continue;
    const L = level(x.p, y.p);
    if (L >= o.lv) { out.push({ s: L, a: { n: x.name, fid: x.fid, f: names[x.fid], r: x.r, c: x.c, q: x.qada }, b: { n: y.name, fid: y.fid, f: names[y.fid], r: y.r, c: y.c, q: y.qada } }); if (out.length >= 20000) return out.sort((m, n) => n.s - m.s); }
  }
  return out.sort((m, n) => n.s - m.s);
});

// ---------- الانتقال إلى الخلية في Excel ----------
ipcMain.handle("go", (_, { fid, row, col }) => {
  const f = q("SELECT path FROM files WHERE id=?", [fid])[0];
  if (!f || !fs.existsSync(f.path)) return "الملف الأصلي غير موجود في مساره";
  const cell = String(col).replace(/[^A-Z]/g, "") + (+row);
  const ps = `$x=New-Object -ComObject Excel.Application;$x.Visible=$true;$w=$x.Workbooks.Open('${f.path.replace(/'/g, "''")}');$s=$w.Sheets.Item(1);$s.Activate();$s.Range('${cell}').Select()`;
  return new Promise(res => execFile("powershell", ["-NoProfile", "-Command", ps], e => {
    if (e) { shell.openPath(f.path); res("فُتح الملف، والخلية: " + cell); } else res("تم الانتقال إلى الخلية " + cell);
  }));
});

// ---------- التصدير (Excel / PDF) ----------
ipcMain.handle("exp", async (_, { kind, head, rows }) => {
  const r = await dialog.showSaveDialog(win, { defaultPath: "النتائج." + kind, filters: [{ name: kind.toUpperCase(), extensions: [kind] }] });
  if (r.canceled) return "";
  if (kind === "xlsx") {
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([head, ...rows]), "النتائج");
    wb.Workbook = { Views: [{ RTL: true }] }; XLSX.writeFile(wb, r.filePath);
  } else {
    const e = s => String(s).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
    const html = `<meta charset="utf-8"><body dir="rtl" style="font-family:Tahoma;font-size:11px"><table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;width:100%"><tr>${head.map(h => `<th>${e(h)}</th>`).join("")}</tr>${rows.map(x => `<tr>${x.map(v => `<td>${e(v)}</td>`).join("")}</tr>`).join("")}</table>`;
    const w = new BrowserWindow({ show: false });
    await w.loadURL("data:text/html;charset=utf-8," + encodeURIComponent(html));
    fs.writeFileSync(r.filePath, await w.webContents.printToPDF({ landscape: true })); w.destroy();
  }
  return r.filePath;
});

// ---------- مجلد القاعدة والنسخ الاحتياطي ----------
ipcMain.handle("dir", async () => {
  const r = await dialog.showOpenDialog(win, { properties: ["openDirectory", "createDirectory"] });
  if (r.canceled) return null;
  save(); const old = dbPath; cfg.dir = r.filePaths[0]; saveCfg();
  const nu = path.join(cfg.dir, "people.sqlite");
  if (!fs.existsSync(nu)) fs.copyFileSync(old, nu);
  loadDb(); log("تغيير مجلد قاعدة البيانات"); save(); return cfg.dir;
});
ipcMain.handle("backup", async () => {
  const r = await dialog.showSaveDialog(win, { defaultPath: "نسخة_احتياطية.sqlite" });
  if (r.canceled) return ""; save(); fs.copyFileSync(dbPath, r.filePath); return r.filePath;
});
ipcMain.handle("restore", async () => {
  const r = await dialog.showOpenDialog(win, { properties: ["openFile"], filters: [{ name: "SQLite", extensions: ["sqlite", "db"] }] });
  if (r.canceled) return 0;
  fs.copyFileSync(r.filePaths[0], dbPath); loadDb(); log("استعادة نسخة احتياطية"); save(); return 1;
});

// ---------- النافذة ----------
function createWindow() {
  win = new BrowserWindow({
    width: 1200, height: 820, minWidth: 360, minHeight: 500, autoHideMenuBar: true, title: "معلومات شخصية",
    webPreferences: { preload: path.join(__dirname, "preload.js"), contextIsolation: true }
  });
  Menu.setApplicationMenu(Menu.buildFromTemplate([{ role: "editMenu" }, { role: "viewMenu" }]));
  win.webContents.on("context-menu", (e, p) => {
    Menu.buildFromTemplate([
      { role: "cut", label: "قص", enabled: p.editFlags.canCut },
      { role: "copy", label: "نسخ", enabled: p.editFlags.canCopy },
      { role: "paste", label: "لصق", enabled: p.editFlags.canPaste }
    ]).popup();
  });
  win.loadFile(path.join(__dirname, "desktop.html"));
}
app.whenReady().then(async () => { loadCfg(); SQL = await initSqlJs(); loadDb(); createWindow(); });
app.on("window-all-closed", () => app.quit());
