"""The page shown in the browser (one self-contained HTML string, no internet needed)."""

PAGE = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel to PDF</title>
<style>
:root{
  --bg:#f4f6f9;--card:#fff;--text:#1d2330;--muted:#667085;--line:#e1e5ec;
  --accent:#0b5cad;--accent-text:#fff;--ok:#1b7f3b;--fail:#b3261e;--run:#0b5cad;--skip:#6b6b6b;
  --warn-bg:#fff4ce;--warn-text:#5c4400;--err-bg:#fde8e7;--err-text:#7a1913;--info-bg:#e8f1fb;--info-text:#0b3d75;
  --row-hover:#f7f9fc;
}
@media (prefers-color-scheme: dark){:root{
  --bg:#14171c;--card:#1d2128;--text:#e8eaee;--muted:#9aa3b2;--line:#2c323c;
  --accent:#4c9aff;--accent-text:#0b1220;--ok:#4cc27a;--fail:#ff8a80;--run:#6db3ff;--skip:#a0a7b4;
  --warn-bg:#3a3010;--warn-text:#ffe08a;--err-bg:#40191a;--err-text:#ffb4ab;--info-bg:#162a42;--info-text:#a9d0ff;
  --row-hover:#232830;
}}
*{box-sizing:border-box}
html,body{margin:0}
body{background:var(--bg);color:var(--text);font:15px/1.45 system-ui,"Segoe UI","Khmer UI","Leelawadee UI","Noto Sans Khmer",sans-serif}
.wrap{max-width:980px;margin:0 auto;padding:20px 16px 40px}
header{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:14px}
h1{font-size:22px;margin:0}
h1 small{font-weight:400;color:var(--muted);font-size:13px;margin-left:8px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:14px}
.btn{font:inherit;border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:8px;padding:7px 14px;cursor:pointer}
.btn:hover:not(:disabled){border-color:var(--accent)}
.btn:disabled{opacity:.45;cursor:not-allowed}
.btn.primary{background:var(--accent);color:var(--accent-text);border-color:var(--accent);font-weight:600;padding:10px 22px;font-size:16px}
.btn.link{border:none;background:none;color:var(--accent);padding:2px 6px;text-decoration:underline}
.btn.x{border:none;background:none;color:var(--muted);padding:2px 8px;font-size:16px}
.bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.spacer{flex:1}
.banner{border-radius:10px;padding:10px 14px;margin-bottom:12px;position:relative}
.banner.warn{background:var(--warn-bg);color:var(--warn-text)}
.banner.error{background:var(--err-bg);color:var(--err-text)}
.banner.info{background:var(--info-bg);color:var(--info-text)}
.banner b{display:block;margin-bottom:4px;padding-right:28px}
.banner ul{margin:4px 0 0;padding-left:20px}
.banner .hint{margin-top:6px;opacity:.9}
.banner .close{position:absolute;top:6px;right:8px;border:none;background:none;color:inherit;font-size:18px;cursor:pointer}
.drop{border:2px dashed var(--line);border-radius:12px;padding:22px;text-align:center;color:var(--muted);margin-top:12px;transition:.15s}
.drop.over{border-color:var(--accent);background:var(--info-bg);color:var(--info-text)}
.pathrow{display:flex;gap:8px;margin-top:12px}
input[type=text],select{font:inherit;color:var(--text);background:var(--card);border:1px solid var(--line);border-radius:8px;padding:7px 10px}
.pathrow input{flex:1;min-width:0}
.tablebox{max-height:340px;overflow:auto;margin-top:12px;border:1px solid var(--line);border-radius:10px}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top}
th{position:sticky;top:0;background:var(--card);font-size:13px;color:var(--muted);z-index:1}
tr:last-child td{border-bottom:none}
tbody tr:hover{background:var(--row-hover)}
td.name{word-break:break-all}
td.time,th.time{text-align:right;white-space:nowrap}
td.act{white-space:nowrap;text-align:right}
.s-waiting{color:var(--muted)}.s-running{color:var(--run)}.s-done{color:var(--ok)}
.s-failed{color:var(--fail)}.s-skipped,.s-cancelled{color:var(--skip)}
pre.raw{margin:6px 0 0;white-space:pre-wrap;word-break:break-word;font:12px/1.4 ui-monospace,Consolas,monospace;color:var(--muted);max-height:160px;overflow:auto}
.empty{padding:26px;text-align:center;color:var(--muted)}
fieldset{border:1px solid var(--line);border-radius:10px;margin:0 0 14px;padding:12px 14px}
legend{padding:0 6px;color:var(--muted);font-size:13px}
.opt{display:grid;grid-template-columns:auto 1fr;gap:10px 12px;align-items:center}
.opt label.row{grid-column:1/-1;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.opt input[type=text]{flex:1;min-width:180px}
progress{width:100%;height:10px;accent-color:var(--accent)}
.summary{margin:6px 0 12px;color:var(--muted)}
.actions{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.note{font-size:13px;color:var(--muted);margin-top:8px}
#closed{position:fixed;inset:0;background:var(--bg);display:none;align-items:center;justify-content:center;text-align:center;z-index:10;padding:20px}
#closed div{max-width:420px}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Excel to PDF<small>v__VERSION__ &middot; runs on this PC, nothing is uploaded to the internet</small></h1>
    <button class="btn" id="btnQuit" title="Close the program">Quit app</button>
  </header>

  <div id="banners"></div>
  <div id="notices"></div>

  <div class="card">
    <div class="bar">
      <button class="btn" id="btnFiles">Add files&hellip;</button>
      <button class="btn" id="btnFolder">Add folder&hellip;</button>
      <button class="btn" id="btnClear">Clear list</button>
      <span class="spacer"></span>
      <span id="uploadLine" class="note" style="margin:0"></span>
    </div>
    <input type="file" id="inFiles" accept=".xlsx,.xlsm" multiple hidden>
    <input type="file" id="inFolder" webkitdirectory hidden>
    <div class="drop" id="drop">Drag Excel files or a folder here</div>
    <div class="pathrow">
      <input type="text" id="pathBox" placeholder="Or type a file or folder path on this PC, for example D:\reports" spellcheck="false">
      <button class="btn" id="btnPath">Add path</button>
    </div>
    <div class="tablebox">
      <table>
        <thead><tr><th>Excel file</th><th>Status</th><th class="time">Time</th><th></th></tr></thead>
        <tbody id="rows"></tbody>
      </table>
    </div>
  </div>

  <div class="card">
    <fieldset>
      <legend>Options</legend>
      <div class="opt">
        <label class="row"><input type="radio" name="outmode" value="folder"> Save PDFs in this folder:
          <input type="text" id="outDir" spellcheck="false"></label>
        <label class="row"><input type="radio" name="outmode" value="same"> Save PDF next to the Excel file
          <span class="note" style="margin:0">(only for files added by path; uploaded files always use the folder above)</span></label>
        <span>Khmer font in PDF:</span><select id="khmer"></select>
        <span>Speed:</span>
        <select id="speed">
          <option>Auto (recommended)</option>
          <option>Faster (3 at once)</option>
          <option>Safer (1 at a time)</option>
        </select>
        <label class="row"><input type="checkbox" id="skip"> Skip files that already have an up-to-date PDF</label>
      </div>
    </fieldset>

    <progress id="prog" value="0" max="1"></progress>
    <div class="summary" id="summary">Ready.</div>
    <div class="actions">
      <button class="btn primary" id="btnGo">Convert to PDF</button>
      <button class="btn" id="btnCancel" disabled>Cancel</button>
      <span class="spacer"></span>
      <a class="btn" id="btnZip" href="/api/zip" download style="text-decoration:none">Download all (ZIP)</a>
      <button class="btn" id="btnOpen">Open output folder</button>
    </div>
    <div class="note" id="engineLine"></div>
  </div>
</div>

<div id="closed"><div><h2>The app is closed</h2><p>You can close this tab. Start excel2pdf.exe again to use it.</p></div></div>

<script>
"use strict";
const $ = s => document.querySelector(s);
const cid = Math.random().toString(36).slice(2) + Date.now().toString(36);
let S = null, loaded = false, uploading = false, fails = 0, closed = false, lastKey = "";
let localNotices = [], localSeq = 0;
const openDetails = new Set();

/* ---------------------------------------------------------------- helpers */
function el(tag, cls, text){
  const e = document.createElement(tag);
  if(cls) e.className = cls;
  if(text !== undefined) e.textContent = text;
  return e;
}
async function api(path, body){
  const r = await fetch(path, {method:"POST", headers:{"Content-Type":"application/json"},
                               body: JSON.stringify(body || {})});
  let j = {};
  try{ j = await r.json(); }catch(e){}
  if(!r.ok && !j.error) j.error = "Request failed (" + r.status + ").";
  return j;
}
function addLocal(level, title, rows, hint){
  localNotices.push({id:"L" + (++localSeq), level, title, rows:(rows||[]).map(r => ({name:r[0], reason:r[1]})), hint:hint||"", more:0});
  renderNotices();
}
function rejectReason(name){
  const n = name.toLowerCase();
  if(n.startsWith("~$")) return "Temporary Excel lock file (the file is open in Excel).";
  if(!(n.endsWith(".xlsx") || n.endsWith(".xlsm"))) return "Unsupported file type. Only .xlsx and .xlsm are accepted.";
  return "";
}

/* ---------------------------------------------------------------- rendering */
function renderNotices(){
  const box = $("#notices");
  box.textContent = "";
  const all = (S ? S.notices : []).concat(localNotices);
  for(const n of all){
    const b = el("div", "banner " + n.level);
    const close = el("button", "close", "\u00d7");
    close.title = "Dismiss";
    close.onclick = async () => {
      if(String(n.id).startsWith("L")) localNotices = localNotices.filter(x => x.id !== n.id);
      else await api("/api/dismiss", {id:n.id});
      await refresh();
      renderNotices();
    };
    b.appendChild(close);
    b.appendChild(el("b", "", n.title));
    if(n.rows && n.rows.length){
      const ul = el("ul");
      for(const r of n.rows){
        const li = el("li");
        if(r.name){ li.appendChild(el("strong", "", r.name)); li.appendChild(document.createTextNode(r.reason ? ": " + r.reason : "")); }
        else li.textContent = r.reason;
        ul.appendChild(li);
      }
      if(n.more) ul.appendChild(el("li", "", "... and " + n.more + " more."));
      b.appendChild(ul);
    }
    if(n.hint) b.appendChild(el("div", "hint", n.hint));
    box.appendChild(b);
  }
}

function renderBanners(){
  const box = $("#banners");
  box.textContent = "";
  if(S.engine_problem){
    const b = el("div", "banner error"); b.textContent = S.engine_problem; box.appendChild(b);
  }
  if(!S.khmer_fonts.length){
    const b = el("div", "banner warn");
    b.textContent = "No Khmer font found on this PC. Khmer text may look broken in the PDF. " +
                    "Install a Khmer font (for example Khmer OS System or Noto Sans Khmer).";
    box.appendChild(b);
  }
}

function renderRows(){
  const tb = $("#rows");
  tb.textContent = "";
  if(!S.items.length){
    const tr = el("tr"), td = el("td", "empty", "No files yet. Add files, drop them here, or type a path.");
    td.colSpan = 4; tr.appendChild(td); tb.appendChild(tr);
    return;
  }
  const running = S.run.running;
  for(const it of S.items){
    const tr = el("tr");
    const name = el("td", "name", it.name);
    name.title = it.kind === "path" ? "Read from this PC" : "Uploaded copy";
    const st = el("td", "s-" + it.status, it.text);
    if(it.status === "failed" && it.raw && openDetails.has(it.id)) st.appendChild(el("pre", "raw", it.raw));
    const tm = el("td", "time", it.time || "");
    const act = el("td", "act");
    if(it.pdf){
      const a = el("a", "btn link", "PDF"); a.href = "/api/download?id=" + encodeURIComponent(it.id); a.download = "";
      act.appendChild(a);
    }
    if(it.status === "failed" && it.raw){
      const d = el("button", "btn link", openDetails.has(it.id) ? "Hide details" : "Details");
      d.onclick = () => { openDetails.has(it.id) ? openDetails.delete(it.id) : openDetails.add(it.id); lastKey = ""; renderAll(); };
      act.appendChild(d);
    }
    const x = el("button", "btn x", "\u2715"); x.title = "Remove from list"; x.disabled = running;
    x.onclick = async () => { await api("/api/remove", {id:it.id}); await refresh(); };
    act.appendChild(x);
    tr.append(name, st, tm, act);
    tb.appendChild(tr);
  }
}

function renderControls(){
  const r = S.run, running = r.running;
  const busy = running || uploading;
  for(const id of ["#btnFiles", "#btnFolder", "#btnClear", "#btnPath"]) $(id).disabled = busy;
  $("#btnGo").disabled = busy || !S.items.length;
  $("#btnCancel").disabled = !running;
  const anyPdf = S.items.some(i => i.pdf);
  const zip = $("#btnZip");
  zip.style.opacity = anyPdf ? "1" : ".45";
  zip.style.pointerEvents = anyPdf ? "auto" : "none";
  const p = $("#prog");
  p.max = Math.max(1, r.total); p.value = r.done;
  let text = "Ready.";
  if(running) text = r.done + " of " + r.total + " finished...";
  else if(r.finished) text = r.cancelled ? "Cancelled." :
      "Finished: " + r.ok + " done, " + r.fail + " failed, in " + Math.round(r.elapsed) + " s.";
  $("#summary").textContent = text;
  $("#engineLine").textContent = "Engine: " + S.engine_name + " \u00b7 PDFs are saved to: " + S.settings.out_dir;
}

function loadSettingsOnce(){
  if(loaded) return;
  loaded = true;
  const s = S.settings;
  document.querySelectorAll("input[name=outmode]").forEach(r => r.checked = (r.value === s.out_mode));
  $("#outDir").value = s.out_dir;
  const k = $("#khmer");
  for(const f of ["Auto (detect)"].concat(S.khmer_fonts)){ const o = el("option", "", f); o.value = f; k.appendChild(o); }
  if(![...k.options].some(o => o.value === s.khmer)){ const o = el("option", "", s.khmer); o.value = s.khmer; k.appendChild(o); }
  k.value = s.khmer;
  $("#speed").value = s.speed;
  $("#skip").checked = s.skip;
}

function renderAll(){
  if(!S) return;
  loadSettingsOnce();
  const key = JSON.stringify([S.items, S.run, S.notices, S.engine_problem, uploading, [...openDetails], localNotices.length]);
  if(key === lastKey) return;
  lastKey = key;
  renderBanners(); renderNotices(); renderRows(); renderControls();
}

async function refresh(){
  const r = await fetch("/api/state?cid=" + cid, {cache:"no-store"});
  if(!r.ok) throw new Error("state");
  S = await r.json();
  renderAll();
}
async function poll(){
  if(closed) return;
  try{ await refresh(); fails = 0; }
  catch(e){ if(++fails >= 3){ closed = true; $("#closed").style.display = "flex"; return; } }
  setTimeout(poll, S && S.run.running ? 600 : 1500);
}

/* ---------------------------------------------------------------- adding files */
function uploadOne(f, onp){
  return new Promise(res => {
    const x = new XMLHttpRequest();
    x.open("POST", "/api/upload?name=" + encodeURIComponent(f.name) + "&mtime=" + f.lastModified);
    x.setRequestHeader("Content-Type", "application/octet-stream");
    x.upload.onprogress = e => { if(e.lengthComputable) onp(Math.round(100 * e.loaded / e.total)); };
    x.onload = () => {
      let j = {};
      try{ j = JSON.parse(x.responseText); }catch(e){}
      res(x.status === 200 && j.ok ? {ok:true} : {ok:false, reason:j.reason || j.error || "Upload failed."});
    };
    x.onerror = () => res({ok:false, reason:"Upload failed (connection lost)."});
    x.send(f);
  });
}

async function addFiles(list){          // list = [{file, folder}]
  const good = [], bad = [];
  let ignored = 0;
  for(const x of list){
    const why = rejectReason(x.file.name);
    if(!why) good.push(x.file);
    else if(x.folder) ignored++;
    else bad.push([x.file.name, why]);
  }
  if(ignored) addLocal("info", ignored + " other file(s) in the folder were ignored (not .xlsx or .xlsm, or temporary ~$ files).");
  if(!good.length){
    if(bad.length) addLocal("warn", "These files were skipped and not added:", bad);
    else if(!ignored) addLocal("info", "No Excel files were found.");
    return;
  }
  uploading = true; lastKey = ""; renderAll();
  for(let i = 0; i < good.length; i++){
    const f = good[i];
    const res = await uploadOne(f, p => { $("#uploadLine").textContent = "Adding " + (i + 1) + " of " + good.length + ": " + f.name + " (" + p + "%)"; });
    if(!res.ok) bad.push([f.name, res.reason]);
  }
  $("#uploadLine").textContent = "";
  uploading = false; lastKey = "";
  if(bad.length) addLocal("warn", "These files were skipped and not added:", bad);
  await refresh();
}

function readEntries(dir){
  return new Promise((res, rej) => {
    const out = [], rd = dir.createReader();
    (function more(){ rd.readEntries(b => { if(!b.length) return res(out); out.push(...b); more(); }, rej); })();
  });
}
async function walk(entry, out){
  if(entry.isFile) await new Promise(r => entry.file(f => { out.push({file:f, folder:true}); r(); }, () => r()));
  else if(entry.isDirectory) for(const e of await readEntries(entry)) await walk(e, out);
}
async function fromDrop(dt){
  const out = [];
  if(dt.items && dt.items.length && dt.items[0].webkitGetAsEntry){
    const entries = [];
    for(const it of dt.items){ const e = it.webkitGetAsEntry && it.webkitGetAsEntry(); if(e) entries.push(e); }
    for(const e of entries){
      if(e.isFile) await new Promise(r => e.file(f => { out.push({file:f, folder:false}); r(); }, () => r()));
      else await walk(e, out);
    }
  } else for(const f of dt.files) out.push({file:f, folder:false});
  return out;
}

/* ---------------------------------------------------------------- events */
$("#btnFiles").onclick = () => $("#inFiles").click();
$("#btnFolder").onclick = () => $("#inFolder").click();
$("#inFiles").onchange = e => { addFiles([...e.target.files].map(f => ({file:f, folder:false}))); e.target.value = ""; };
$("#inFolder").onchange = e => { addFiles([...e.target.files].map(f => ({file:f, folder:true}))); e.target.value = ""; };

const drop = $("#drop");
["dragenter", "dragover"].forEach(t => window.addEventListener(t, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach(t => window.addEventListener(t, e => { e.preventDefault(); if(t === "drop" || e.target === document.documentElement) drop.classList.remove("over"); }));
window.addEventListener("drop", async e => { if(!uploading && !(S && S.run.running)) addFiles(await fromDrop(e.dataTransfer)); });

async function addPath(){
  const box = $("#pathBox");
  if(!box.value.trim()){ addLocal("info", "Please type a file or folder path first."); return; }
  const j = await api("/api/add_path", {path:box.value});
  if(j.error) addLocal("error", j.error); else box.value = "";
  await refresh();
}
$("#btnPath").onclick = addPath;
$("#pathBox").addEventListener("keydown", e => { if(e.key === "Enter") addPath(); });

$("#btnClear").onclick = async () => { await api("/api/clear"); openDetails.clear(); await refresh(); };

async function saveSettings(){
  const mode = document.querySelector("input[name=outmode]:checked");
  const j = await api("/api/settings", {out_mode:mode ? mode.value : "folder", out_dir:$("#outDir").value,
                                        khmer:$("#khmer").value, speed:$("#speed").value, skip:$("#skip").checked});
  if(j.error) addLocal("error", j.error);
  await refresh();
}
document.querySelectorAll("input[name=outmode]").forEach(r => r.onchange = saveSettings);
for(const id of ["#outDir", "#khmer", "#speed", "#skip"]) $(id).onchange = saveSettings;

$("#btnGo").onclick = async () => {
  await saveSettings();
  const j = await api("/api/convert");
  if(j.error) addLocal("error", j.error);
  await refresh();
};
$("#btnCancel").onclick = async () => { await api("/api/cancel"); await refresh(); };
$("#btnOpen").onclick = async () => { const j = await api("/api/open_folder"); if(j.error) addLocal("error", j.error); };
$("#btnQuit").onclick = async () => {
  if(S && S.run.running && !confirm("Files are still being converted. Quit anyway?")) return;
  await api("/api/quit");
  closed = true; $("#closed").style.display = "flex";
};
window.addEventListener("pagehide", () => { try{ navigator.sendBeacon("/api/bye?cid=" + cid); }catch(e){} });

poll();
</script>
</body>
</html>
'''
