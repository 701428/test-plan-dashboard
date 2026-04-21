"""
Flask — Test Plan Dashboard
• /          → dashboard UI
• /data      → all sheets as JSON (reads active Excel from disk)
• /upload    → replace active Excel (saved to disk, shared across workers)
• /effort    → HW sheet tests with mandays for effort calculator
"""

from pathlib import Path
from flask import Flask, jsonify, render_template_string, request
import openpyxl

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB

BASE        = Path(__file__).parent
DEFAULT_XL  = BASE / "Abbreviation of Test Plan_1P _ Better Meter.xlsx"
UPLOADED_XL = BASE / "uploaded_data.xlsx"   # persists on disk → all workers see it

SPECIAL = {
    "Abbreviation of Test Plan_HW": {"hdr": [2, 3], "data": 4},
}
DISPLAY = {
    "Priority Test":                   "Priority Test",
    "Abbreviation of Test Plan_HW":    "HW Test Plan",
    "Abbreviation of Test Plan_FW":    "FW Test Plan",
    "Abbreviation of TP_Module HW":    "Module HW",
    "Abbreviation of TP_CommsTesting": "Comms Testing",
}


def active_xl():
    return UPLOADED_XL if UPLOADED_XL.exists() else DEFAULT_XL


def clean(v):
    return "" if v is None else str(v).replace("\n", " ").strip()


def find_hdr(rows):
    for i, r in enumerate(rows[:5]):
        if sum(1 for v in r if v is not None and str(v).strip()) >= 2:
            return i
    return 0


def read_excel():
    path = active_xl()
    wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out  = {}
    for name in wb.sheetnames:
        ws   = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        display = DISPLAY.get(name, name)
        if name in SPECIAL:
            cfg  = SPECIAL[name]
            hdr_rows = [rows[i] for i in cfg["hdr"] if i < len(rows)]
            ncols = max(len(r) for r in hdr_rows)
            hdrs  = []
            for c in range(ncols):
                parts = []
                for r in hdr_rows:
                    v = clean(r[c]) if c < len(r) else ""
                    if v and v.lower() not in ("none","false","true"):
                        parts.append(v)
                hdrs.append(" / ".join(parts) if parts else f"Col{c+1}")
            data_rows = rows[cfg["data"]:]
        else:
            hi   = find_hdr(rows)
            hdrs = [clean(v) or f"Col{i+1}" for i,v in enumerate(rows[hi])]
            data_rows = rows[hi+1:]
        data = [[clean(v) for v in r] for r in data_rows if any(v is not None and str(v).strip() for v in r)]
        out[display] = {"headers": hdrs, "rows": data}
    wb.close()
    return out


def read_effort():
    """
    Parse HW sheet for effort calculator.
    Returns list of items:
      { id, group, name, level, mandays, is_group_header }
    """
    path = active_xl()
    wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "Abbreviation of Test Plan_HW" not in wb.sheetnames:
        wb.close()
        return []
    ws   = wb["Abbreviation of Test Plan_HW"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    items = []
    current_group = ""
    for row in rows[4:]:   # data starts row index 4
        sno     = row[0]
        name    = clean(row[1])
        level   = row[2]
        mandays = row[23]   # "Man Days Required for the test"

        if not name:
            continue

        # Group header rows have a letter in col 0 (A, B, C…)
        if isinstance(sno, str) and sno.strip().isalpha():
            current_group = name
            items.append({
                "id": f"grp_{sno.strip()}",
                "group": name,
                "name": name,
                "level": None,
                "mandays": None,
                "is_group": True,
            })
        elif isinstance(sno, (int, float)):
            md = None
            if mandays is not None:
                try:
                    md = float(mandays)
                except (ValueError, TypeError):
                    md = None
            lv = None
            if level is not None:
                try:
                    lv = int(level)
                except (ValueError, TypeError):
                    lv = None
            items.append({
                "id": f"test_{int(sno)}_{len(items)}",
                "group": current_group,
                "name": name,
                "level": lv,
                "mandays": md,
                "is_group": False,
            })
    return items


# ─── HTML ─────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Test Plan Dashboard</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#0f1117;color:#e2e8f0;min-height:100vh}

/* ── top bar ── */
.top-bar{background:linear-gradient(135deg,#1e293b,#0f172a);border-bottom:1px solid #334155;
  padding:13px 22px;display:flex;align-items:center;gap:12px}
.logo{width:36px;height:36px;background:linear-gradient(135deg,#6366f1,#8b5cf6);
  border-radius:9px;display:grid;place-items:center;font-size:17px;color:#fff;font-weight:700;flex-shrink:0}
.tb-text{flex:1}
.tb-text h1{font-size:1.05rem;font-weight:700;color:#f1f5f9;display:flex;align-items:center;gap:7px}
.badge{font-size:.6rem;background:#166534;color:#4ade80;border:1px solid #22543d;
  padding:2px 7px;border-radius:20px;font-weight:700}
.tb-text .sub{font-size:.72rem;color:#475569;margin-top:2px}
.tb-actions{display:flex;align-items:center;gap:8px}
.ls{font-size:.69rem;color:#475569}

/* ── tabs ── */
.tabs{display:flex;gap:2px;background:#0f1117;border-bottom:1px solid #334155;padding:0 22px}
.tab{padding:10px 18px;font-size:.82rem;font-weight:600;color:#64748b;cursor:pointer;
  border-bottom:2px solid transparent;transition:all .2s;white-space:nowrap}
.tab:hover{color:#a5b4fc}
.tab.active{color:#a5b4fc;border-bottom-color:#6366f1}

/* ── buttons ── */
.btn{border-radius:8px;padding:6px 13px;font-size:.78rem;cursor:pointer;
  display:flex;align-items:center;gap:5px;transition:all .2s;
  border:1px solid #334155;background:#1e293b;color:#94a3b8;white-space:nowrap}
.btn:hover{background:#334155;color:#e2e8f0;border-color:#6366f1}
.btn-pur{background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;
  border-color:transparent;box-shadow:0 2px 10px rgba(99,102,241,.4)}
.btn-pur:hover{opacity:.88;border-color:transparent}
.spin{animation:spin .7s linear infinite;display:inline-block}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── upload overlay ── */
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);
  z-index:200;align-items:center;justify-content:center}
.overlay.show{display:flex}
.upload-box{background:#1e293b;border:2px dashed #6366f1;border-radius:16px;
  padding:44px 52px;text-align:center;max-width:400px;width:92%;position:relative}
.upload-box.drag{background:#1e1b4b;border-color:#a5b4fc}
.upload-box h2{color:#f1f5f9;margin:12px 0 6px;font-size:1.05rem}
.upload-box p{color:#64748b;font-size:.82rem;margin-bottom:22px}
.close-btn{position:absolute;top:12px;right:14px;background:none;border:none;
  color:#475569;font-size:1.3rem;cursor:pointer;padding:2px 7px;border-radius:5px}
.close-btn:hover{color:#e2e8f0;background:#334155}
.prog-bar{height:5px;background:#334155;border-radius:3px;margin-top:14px;overflow:hidden;display:none}
.prog-fill{height:100%;background:linear-gradient(90deg,#6366f1,#8b5cf6);
  width:0%;transition:width .3s;border-radius:3px}
.prog-txt{font-size:.78rem;color:#94a3b8;margin-top:7px;min-height:20px}

/* ── layout ── */
.layout{display:flex;height:calc(100vh - 107px)}

/* ── sidebar ── */
.sidebar{width:220px;min-width:220px;background:#1e293b;border-right:1px solid #334155;
  display:flex;flex-direction:column;padding:14px 10px;gap:4px;overflow-y:auto}
.sl{font-size:.62rem;font-weight:700;letter-spacing:.12em;color:#475569;
  text-transform:uppercase;padding:0 8px 8px}
.sbtn{background:transparent;border:1px solid transparent;border-radius:9px;
  padding:9px 11px;cursor:pointer;text-align:left;color:#94a3b8;font-size:.8rem;
  font-weight:500;transition:all .18s;line-height:1.4;width:100%}
.sbtn:hover{background:#334155;color:#e2e8f0;border-color:#475569}
.sbtn.active{background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;
  border-color:transparent;box-shadow:0 3px 12px rgba(99,102,241,.4)}
.cnt{display:inline-block;font-size:.62rem;background:rgba(255,255,255,.18);
  border-radius:20px;padding:1px 6px;margin-left:4px;vertical-align:middle}

/* ── main ── */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden}
.toolbar{background:#1a2235;border-bottom:1px solid #334155;padding:10px 18px;
  display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.toolbar h2{font-size:.9rem;font-weight:600;color:#f1f5f9;flex:1}
.chip{background:#334155;border-radius:7px;padding:3px 10px;font-size:.7rem;color:#94a3b8}
.chip span{color:#a5b4fc;font-weight:600}
.sw{position:relative}
.sw input{background:#0f1117;border:1px solid #334155;border-radius:8px;
  padding:5px 11px 5px 28px;color:#e2e8f0;font-size:.8rem;width:190px;
  outline:none;transition:border-color .2s}
.sw input:focus{border-color:#6366f1}
.si{position:absolute;left:8px;top:50%;transform:translateY(-50%);
  color:#475569;pointer-events:none;font-size:11px}

.table-wrap{flex:1;overflow:auto;padding:14px 16px}
table{width:100%;border-collapse:collapse;font-size:.77rem}
thead tr{position:sticky;top:0;z-index:10}
thead th{background:#1e293b;border:1px solid #334155;padding:8px 10px;
  font-weight:600;color:#a5b4fc;white-space:nowrap;font-size:.71rem;
  text-transform:uppercase;letter-spacing:.04em;text-align:left}
tbody tr:nth-child(even){background:#141b2d}
tbody tr:hover{background:#1e3a5f}
tbody td{border:1px solid #1e293b;padding:7px 10px;color:#cbd5e1;
  max-width:250px;word-break:break-word;vertical-align:top}
.ec{color:#334155;font-style:italic;font-size:.69rem}

/* ── effort calc ── */
.effort-wrap{flex:1;display:flex;overflow:hidden}
.effort-list{flex:1;overflow-y:auto;padding:16px}
.effort-summary{width:280px;min-width:280px;background:#1e293b;
  border-left:1px solid #334155;padding:20px;display:flex;flex-direction:column;gap:14px;overflow-y:auto}

.group-hdr{background:#1e1b4b;border-radius:8px;padding:8px 12px;
  font-size:.78rem;font-weight:700;color:#a5b4fc;margin:10px 0 4px;
  display:flex;align-items:center;gap:8px;cursor:pointer;user-select:none}
.group-hdr:first-child{margin-top:0}
.group-hdr .garrow{transition:transform .2s;display:inline-block}
.group-hdr.collapsed .garrow{transform:rotate(-90deg)}

.test-row{display:flex;align-items:flex-start;gap:10px;padding:8px 10px;
  border-radius:8px;transition:background .15s;cursor:pointer;margin-bottom:2px}
.test-row:hover{background:#1e293b}
.test-row input[type=checkbox]{width:15px;height:15px;accent-color:#6366f1;
  cursor:pointer;flex-shrink:0;margin-top:2px}
.test-row .tname{flex:1;font-size:.8rem;color:#cbd5e1;line-height:1.4}
.test-row .tlevel{font-size:.67rem;padding:1px 6px;border-radius:4px;
  font-weight:600;flex-shrink:0;margin-top:2px}
.lv1{background:#7f1d1d;color:#fca5a5}
.lv2{background:#78350f;color:#fcd34d}
.lv3{background:#14532d;color:#86efac}
.test-row .tdays{font-size:.77rem;color:#a5b4fc;font-weight:600;
  flex-shrink:0;min-width:48px;text-align:right;margin-top:2px}

/* summary panel */
.sum-title{font-size:.7rem;font-weight:700;letter-spacing:.1em;
  color:#475569;text-transform:uppercase}
.sum-big{font-size:2.6rem;font-weight:800;color:#a5b4fc;line-height:1}
.sum-unit{font-size:.75rem;color:#64748b;margin-top:2px}
.sum-divider{border:none;border-top:1px solid #334155}
.sum-row{display:flex;justify-content:space-between;align-items:center;
  font-size:.78rem}
.sum-row .label{color:#64748b}
.sum-row .val{color:#e2e8f0;font-weight:600}
.sum-row .val.hi{color:#f87171}
.sum-row .val.mid{color:#fcd34d}
.sum-row .val.lo{color:#86efac}
.sel-count{font-size:.78rem;color:#64748b}
.clear-btn{width:100%;border-radius:8px;padding:8px;font-size:.8rem;
  cursor:pointer;border:1px solid #334155;background:#0f1117;
  color:#94a3b8;transition:all .2s;text-align:center}
.clear-btn:hover{border-color:#f87171;color:#f87171}
.effort-toolbar{background:#1a2235;border-bottom:1px solid #334155;
  padding:10px 16px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.effort-toolbar h2{font-size:.9rem;font-weight:600;color:#f1f5f9;flex:1}
.filter-btns{display:flex;gap:5px}
.fbtn{padding:4px 10px;font-size:.72rem;border-radius:6px;cursor:pointer;
  border:1px solid #334155;background:transparent;color:#64748b;transition:all .18s}
.fbtn:hover{border-color:#6366f1;color:#a5b4fc}
.fbtn.on{background:#312e81;border-color:#6366f1;color:#a5b4fc}

.state-box{flex:1;display:flex;flex-direction:column;align-items:center;
  justify-content:center;gap:12px;color:#475569;text-align:center;padding:40px}
.state-box .big{font-size:2.6rem}
.state-box p{font-size:.86rem;line-height:1.6}
.err{color:#f87171}

::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:#0f1117}
::-webkit-scrollbar-thumb{background:#334155;border-radius:3px}
</style>
</head>
<body>

<!-- Upload overlay -->
<div class="overlay" id="overlay" onclick="overlayClick(event)">
  <div class="upload-box" id="upload-box"
       ondragover="doDragOver(event)" ondragleave="doDragLeave()" ondrop="doDrop(event)">
    <button class="close-btn" onclick="closeUpload()">✕</button>
    <div style="font-size:2.2rem">📂</div>
    <h2>Upload Excel File</h2>
    <p>Drag & drop your .xlsx here or click below</p>
    <button class="btn btn-pur" onclick="document.getElementById('fi').click()">
      📎 Choose File
    </button>
    <input type="file" id="fi" style="display:none" accept=".xlsx,.xls" onchange="doUpload(this.files[0])"/>
    <div class="prog-bar" id="prog-bar"><div class="prog-fill" id="prog-fill"></div></div>
    <div class="prog-txt" id="prog-txt"></div>
  </div>
</div>

<!-- Top bar -->
<div class="top-bar">
  <div class="logo">T</div>
  <div class="tb-text">
    <h1>Test Plan Dashboard <span class="badge">LIVE</span></h1>
    <div class="sub" id="sub">Loading…</div>
  </div>
  <div class="tb-actions">
    <span class="ls" id="ls"></span>
    <button class="btn" onclick="loadAll()"><span id="si">🔄</span> Refresh</button>
    <button class="btn btn-pur" onclick="openUpload()">⬆ Upload Excel</button>
  </div>
</div>

<!-- Tabs -->
<div class="tabs">
  <div class="tab active" id="tab-data" onclick="switchTab('data')">📊 Data View</div>
  <div class="tab" id="tab-effort" onclick="switchTab('effort')">⏱ Effort Calculator</div>
</div>

<!-- ══ DATA TAB ══ -->
<div class="layout" id="pane-data">
  <nav class="sidebar" id="sidebar"><div class="sl">Sheets</div></nav>
  <div class="main">
    <div class="state-box" id="data-state"><div class="big">⏳</div><p>Loading…</p></div>
    <div id="data-content" style="display:none;flex-direction:column;flex:1;overflow:hidden">
      <div class="toolbar">
        <h2 id="sheet-title"></h2>
        <div style="display:flex;gap:7px">
          <div class="chip">Rows <span id="rc">0</span></div>
          <div class="chip">Cols <span id="cc">0</span></div>
        </div>
        <div class="sw">
          <span class="si">🔍</span>
          <input type="text" id="q" placeholder="Search…" oninput="filter()"/>
        </div>
      </div>
      <div class="table-wrap">
        <table><thead id="thead"></thead><tbody id="tbody"></tbody></table>
      </div>
    </div>
  </div>
</div>

<!-- ══ EFFORT TAB ══ -->
<div class="layout" id="pane-effort" style="display:none">
  <div class="main">
    <div class="state-box" id="effort-state" style="display:none"><div class="big">⏳</div><p>Loading…</p></div>
    <div style="display:flex;flex-direction:column;flex:1;overflow:hidden">
      <div class="effort-toolbar">
        <h2>HW Test Effort Calculator</h2>
        <div class="sw">
          <span class="si">🔍</span>
          <input type="text" id="eq" placeholder="Search tests…" oninput="filterEffort()"/>
        </div>
        <div class="filter-btns">
          <button class="fbtn on" id="fl-all" onclick="setLevelFilter(0)">All</button>
          <button class="fbtn" id="fl-1" onclick="setLevelFilter(1)" style="color:#fca5a5">Critical</button>
          <button class="fbtn" id="fl-2" onclick="setLevelFilter(2)" style="color:#fcd34d">Hard</button>
          <button class="fbtn" id="fl-3" onclick="setLevelFilter(3)" style="color:#86efac">Others</button>
        </div>
      </div>
      <div class="effort-wrap">
        <div class="effort-list" id="effort-list"></div>
        <div class="effort-summary">
          <div class="sum-title">Total Effort</div>
          <div class="sum-big" id="sum-total">0</div>
          <div class="sum-unit">Man Days</div>
          <hr class="sum-divider"/>
          <div class="sum-row"><span class="label">Tests selected</span><span class="val" id="sum-sel">0</span></div>
          <div class="sum-row"><span class="label">Critical</span><span class="val hi" id="sum-l1">0 days</span></div>
          <div class="sum-row"><span class="label">Hard to pass</span><span class="val mid" id="sum-l2">0 days</span></div>
          <div class="sum-row"><span class="label">Others</span><span class="val lo" id="sum-l3">0 days</span></div>
          <hr class="sum-divider"/>
          <div class="sum-row"><span class="label">Tests with no days</span><span class="val" id="sum-nd">0</span></div>
          <hr class="sum-divider"/>
          <button class="clear-btn" onclick="clearAll()">✕ Clear Selection</button>
        </div>
      </div>
    </div>
  </div>
</div>

<script>
/* ════════════════════════ state ════════════════════════ */
let DATA={}, ROWS=[], ACTIVE=null;
let EFFORT=[], selected=new Set(), lvFilter=0;
const ICONS={"Priority Test":"⭐","HW Test Plan":"🔧","FW Test Plan":"💾","Module HW":"📡","Comms Testing":"📶"};

/* ════════════════════════ tabs ════════════════════════ */
function switchTab(t){
  document.getElementById('pane-data').style.display   = t==='data'  ?'flex':'none';
  document.getElementById('pane-effort').style.display = t==='effort'?'flex':'none';
  document.getElementById('tab-data').classList.toggle('active',  t==='data');
  document.getElementById('tab-effort').classList.toggle('active', t==='effort');
}

/* ════════════════════════ data load ════════════════════════ */
async function fetchWithRetry(url, retries=5, delayMs=6000){
  for(let i=0;i<retries;i++){
    try{
      const ctrl=new AbortController();
      const tid=setTimeout(()=>ctrl.abort(), 55000); // 55s timeout
      const r=await fetch(url,{signal:ctrl.signal});
      clearTimeout(tid);
      if(r.ok) return r;
      throw new Error('HTTP '+r.status);
    }catch(e){
      if(i===retries-1) throw e;
      const wait=delayMs*(i+1);
      setWakeMsg(`Server is waking up… retrying in ${wait/1000}s (${i+1}/${retries})`);
      await new Promise(res=>setTimeout(res,wait));
    }
  }
}

function setWakeMsg(msg){
  document.getElementById('data-state').innerHTML=
    '<div class="big">⏳</div><p style="color:#a5b4fc">'+msg+'</p>';
  document.getElementById('data-state').style.display='flex';
  document.getElementById('data-content').style.display='none';
}

async function loadAll(){
  spin(true);
  await Promise.all([loadData(), loadEffort()]);
  spin(false);
}

async function loadData(){
  try{
    setWakeMsg('Connecting to server…');
    const r=await fetchWithRetry('/data?t='+Date.now());
    const j=await r.json();
    DATA=j.data;
    document.getElementById('sub').textContent=j.filename+' · Real-time from Excel';
    document.getElementById('ls').textContent='Synced '+new Date().toLocaleTimeString();
    buildSidebar();
    if(ACTIVE&&DATA[ACTIVE]) showSheet(ACTIVE);
    else if(Object.keys(DATA).length){ACTIVE=Object.keys(DATA)[0];buildSidebar();showSheet(ACTIVE);}
  }catch(e){
    document.getElementById('data-state').innerHTML=
      '<div class="big">❌</div><p style="color:#f87171">Could not reach server.<br><small>'+e.message+'</small></p>'+
      '<button class="btn btn-pur" onclick="loadAll()" style="margin-top:8px">🔄 Try Again</button>';
    document.getElementById('data-state').style.display='flex';
    document.getElementById('data-content').style.display='none';
  }
}

async function loadEffort(){
  try{
    const r=await fetchWithRetry('/effort?t='+Date.now());
    EFFORT=await r.json();
    renderEffort();
  }catch(e){
    document.getElementById('effort-state').innerHTML=
      '<div class="big">❌</div><p style="color:#f87171">'+e.message+'</p>';
    document.getElementById('effort-state').style.display='flex';
  }
}

/* ════════════════════════ sidebar ════════════════════════ */
function buildSidebar(){
  const sb=document.getElementById('sidebar');
  sb.innerHTML='<div class="sl">Sheets</div>';
  Object.keys(DATA).forEach(s=>{
    const b=document.createElement('button');
    b.className='sbtn'+(s===ACTIVE?' active':'');
    b.innerHTML=(ICONS[s]||'📄')+' '+s+'<span class="cnt">'+DATA[s].rows.length+'</span>';
    b.onclick=()=>{ACTIVE=s;buildSidebar();showSheet(s);};
    sb.appendChild(b);
  });
}

function showSheet(s){
  const{headers,rows:r}=DATA[s];ROWS=r;
  document.getElementById('sheet-title').textContent=s;
  document.getElementById('rc').textContent=r.length;
  document.getElementById('cc').textContent=headers.length;
  document.getElementById('q').value='';
  document.getElementById('thead').innerHTML='<tr>'+headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr>';
  renderRows(r);
  document.getElementById('data-state').style.display='none';
  document.getElementById('data-content').style.display='flex';
}

function renderRows(data){
  document.getElementById('tbody').innerHTML=data.map(r=>
    '<tr>'+r.map(c=>c?'<td>'+esc(c)+'</td>':'<td><span class="ec">—</span></td>').join('')+'</tr>'
  ).join('');
  document.getElementById('rc').textContent=data.length;
}

function filter(){
  const q=document.getElementById('q').value.toLowerCase().trim();
  renderRows(q?ROWS.filter(r=>r.some(c=>c.toLowerCase().includes(q))):ROWS);
}

/* ════════════════════════ effort ════════════════════════ */
function renderEffort(){
  const q=document.getElementById('eq').value.toLowerCase().trim();
  const list=document.getElementById('effort-list');
  list.innerHTML='';
  let curGrp=null, grpEl=null, grpBody=null, grpVisible=true;

  EFFORT.forEach(item=>{
    // level filter
    if(!item.is_group && lvFilter>0 && item.level!==lvFilter) return;
    // search filter
    if(q && !item.name.toLowerCase().includes(q) && !item.group.toLowerCase().includes(q)) return;

    if(item.is_group){
      // Group header
      const g=document.createElement('div');
      g.className='group-hdr';
      g.dataset.grp=item.id;
      g.innerHTML='<span class="garrow">▾</span> '+esc(item.name);
      g.onclick=()=>toggleGroup(item.id);
      list.appendChild(g);
      const body=document.createElement('div');
      body.id='gbody_'+item.id;
      list.appendChild(body);
      curGrp=item.id; grpEl=g; grpBody=body;
    } else {
      const lv=item.level;
      const lvClass=lv===1?'lv1':lv===2?'lv2':'lv3';
      const lvLabel=lv===1?'Critical':lv===2?'Hard':'Others';
      const md=item.mandays!=null?item.mandays.toFixed(1)+' d':'—';
      const checked=selected.has(item.id)?'checked':'';
      const row=document.createElement('div');
      row.className='test-row';
      row.innerHTML=
        '<input type="checkbox" id="cb_'+item.id+'" '+checked+' onchange="toggle(\''+item.id+'\')"/>'+
        '<label class="tname" for="cb_'+item.id+'">'+esc(item.name)+'</label>'+
        (lv?'<span class="tlevel '+lvClass+'">'+lvLabel+'</span>':'')+
        '<span class="tdays">'+md+'</span>';
      (grpBody||list).appendChild(row);
    }
  });
  updateSummary();
}

function toggle(id){
  if(selected.has(id)) selected.delete(id); else selected.add(id);
  updateSummary();
}

function toggleGroup(gid){
  const hdr=document.querySelector('.group-hdr[data-grp="'+gid+'"]');
  const body=document.getElementById('gbody_'+gid);
  if(!hdr||!body) return;
  hdr.classList.toggle('collapsed');
  body.style.display=hdr.classList.contains('collapsed')?'none':'';
}

function clearAll(){
  selected.clear();
  document.querySelectorAll('.effort-list input[type=checkbox]').forEach(cb=>cb.checked=false);
  updateSummary();
}

function updateSummary(){
  const sel=EFFORT.filter(t=>!t.is_group&&selected.has(t.id));
  const total=sel.reduce((s,t)=>s+(t.mandays||0),0);
  const l1=sel.filter(t=>t.level===1).reduce((s,t)=>s+(t.mandays||0),0);
  const l2=sel.filter(t=>t.level===2).reduce((s,t)=>s+(t.mandays||0),0);
  const l3=sel.filter(t=>t.level===3).reduce((s,t)=>s+(t.mandays||0),0);
  const nd=sel.filter(t=>t.mandays==null).length;
  document.getElementById('sum-total').textContent=total.toFixed(1);
  document.getElementById('sum-sel').textContent=sel.length;
  document.getElementById('sum-l1').textContent=l1.toFixed(1)+' days';
  document.getElementById('sum-l2').textContent=l2.toFixed(1)+' days';
  document.getElementById('sum-l3').textContent=l3.toFixed(1)+' days';
  document.getElementById('sum-nd').textContent=nd;
}

function filterEffort(){ renderEffort(); }

function setLevelFilter(lv){
  lvFilter=lv;
  ['fl-all','fl-1','fl-2','fl-3'].forEach((id,i)=>
    document.getElementById(id).classList.toggle('on', i===lv));
  renderEffort();
}

/* ════════════════════════ upload ════════════════════════ */
function openUpload(){ document.getElementById('overlay').classList.add('show'); resetUpload(); }
function closeUpload(){ document.getElementById('overlay').classList.remove('show'); }
function overlayClick(e){ if(e.target===document.getElementById('overlay')) closeUpload(); }
function doDragOver(e){ e.preventDefault(); document.getElementById('upload-box').classList.add('drag'); }
function doDragLeave(){ document.getElementById('upload-box').classList.remove('drag'); }
function doDrop(e){ e.preventDefault(); doDragLeave(); if(e.dataTransfer.files[0]) doUpload(e.dataTransfer.files[0]); }

async function doUpload(file){
  if(!file) return;
  if(!file.name.match(/\.xlsx?$/i)){ alert('Please upload an .xlsx or .xls file'); return; }
  const bar=document.getElementById('prog-bar');
  const fill=document.getElementById('prog-fill');
  const txt=document.getElementById('prog-txt');
  bar.style.display='block'; fill.style.width='30%'; txt.textContent='Uploading '+file.name+'…';
  const fd=new FormData(); fd.append('file',file);
  try{
    fill.style.width='65%';
    const r=await fetch('/upload',{method:'POST',body:fd});
    fill.style.width='90%';
    if(!r.ok) throw new Error(await r.text());
    fill.style.width='100%';
    txt.innerHTML='<span style="color:#4ade80">✓ Uploaded! Refreshing…</span>';
    setTimeout(async()=>{ await loadAll(); closeUpload(); },700);
  }catch(e){
    txt.innerHTML='<span style="color:#f87171">✗ '+e.message+'</span>';
    fill.style.width='0%';
  }
}

function resetUpload(){
  document.getElementById('prog-bar').style.display='none';
  document.getElementById('prog-fill').style.width='0%';
  document.getElementById('prog-txt').textContent='';
  document.getElementById('fi').value='';
}

/* ════════════════════════ helpers ════════════════════════ */
function spin(on){ document.getElementById('si').classList.toggle('spin',on); }
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

/* ════════════════════════ init ════════════════════════ */
loadAll();
setInterval(loadAll, 30000);
// Keep server awake — ping every 10 min so Render free tier doesn't sleep
setInterval(()=>fetch('/ping').catch(()=>{}), 10*60*1000);
</script>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/data")
def data():
    try:
        return jsonify({"filename": active_xl().name, "data": read_excel()})
    except Exception as e:
        return str(e), 500


@app.route("/effort")
def effort():
    try:
        return jsonify(read_effort())
    except Exception as e:
        return str(e), 500


@app.route("/ping")
def ping():
    return "ok", 200


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return "No file provided", 400
    f = request.files["file"]
    if not f.filename:
        return "Empty filename", 400
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return "Only .xlsx / .xls files are supported", 400
    # Save directly to BASE (same folder, all workers read same disk path)
    save_path = UPLOADED_XL
    f.save(str(save_path))
    try:
        wb = openpyxl.load_workbook(save_path, read_only=True)
        wb.close()
    except Exception:
        save_path.unlink(missing_ok=True)
        return "Invalid or corrupt Excel file", 400
    return jsonify({"ok": True, "filename": f.filename})


if __name__ == "__main__":
    app.run(debug=True, port=8765)
