"""
Flask app — Test Plan Dashboard
- Reads Excel file on every /data request
- POST /upload  → replace the Excel file without git/redeploy
- Auto-detects all sheets and their headers
"""

import io
import os
from pathlib import Path
from flask import Flask, jsonify, render_template_string, request

import openpyxl

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB max upload

BASE = Path(__file__).parent
DEFAULT_EXCEL = BASE / "Abbreviation of Test Plan_1P _ Better Meter.xlsx"

# Runtime-replaceable path (updated on upload)
_excel_path = [DEFAULT_EXCEL]

# Known sheets with special header structure
SPECIAL_SHEETS = {
    "Abbreviation of Test Plan_HW": {"header_rows": [2, 3], "data_start": 4},  # 0-indexed
}
DISPLAY_NAMES = {
    "Priority Test":                   "Priority Test",
    "Abbreviation of Test Plan_HW":    "HW Test Plan",
    "Abbreviation of Test Plan_FW":    "FW Test Plan",
    "Abbreviation of TP_Module HW":    "Module HW",
    "Abbreviation of TP_CommsTesting": "Comms Testing",
}


def clean(val):
    if val is None:
        return ""
    return str(val).replace("\n", " ").strip()


def find_header_row(rows):
    """Find the first row that looks like a header (non-empty, mostly strings)."""
    for i, row in enumerate(rows[:5]):
        non_empty = [v for v in row if v is not None and str(v).strip()]
        if len(non_empty) >= 2:
            return i
    return 0


def read_excel():
    path = _excel_path[0]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    result = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            continue

        display = DISPLAY_NAMES.get(sheet_name, sheet_name)

        # Special multi-row header handling
        if sheet_name in SPECIAL_SHEETS:
            cfg = SPECIAL_SHEETS[sheet_name]
            combined = []
            header_row_sets = [all_rows[i] for i in cfg["header_rows"] if i < len(all_rows)]
            max_cols = max(len(r) for r in header_row_sets)
            for col in range(max_cols):
                parts = []
                for row in header_row_sets:
                    v = clean(row[col]) if col < len(row) else ""
                    if v and v.lower() not in ("none", "false", "true"):
                        parts.append(v)
                combined.append(" / ".join(parts) if parts else f"Col{col+1}")
            headers = combined
            data_start = cfg["data_start"]
        else:
            hdr_idx = find_header_row(all_rows)
            headers = [clean(v) or f"Col{i+1}" for i, v in enumerate(all_rows[hdr_idx])]
            data_start = hdr_idx + 1

        rows = []
        for raw in all_rows[data_start:]:
            cleaned = [clean(v) for v in raw]
            if any(cleaned):
                rows.append(cleaned)

        result[display] = {"headers": headers, "rows": rows}

    wb.close()
    return result


# ─── HTML ────────────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Test Plan Dashboard</title>
<style>
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',system-ui,sans-serif;background:#0f1117;color:#e2e8f0;min-height:100vh}

  /* Top bar */
  .top-bar{background:linear-gradient(135deg,#1e293b,#0f172a);border-bottom:1px solid #334155;
    padding:14px 24px;display:flex;align-items:center;gap:12px}
  .logo{width:38px;height:38px;background:linear-gradient(135deg,#6366f1,#8b5cf6);
    border-radius:10px;display:grid;place-items:center;font-size:18px;color:#fff;font-weight:700;flex-shrink:0}
  .top-bar-text{flex:1}
  .top-bar h1{font-size:1.1rem;font-weight:700;color:#f1f5f9;display:flex;align-items:center;gap:8px}
  .live-badge{font-size:.62rem;background:#166534;color:#4ade80;border:1px solid #166534;
    padding:2px 8px;border-radius:20px;font-weight:600}
  .subtitle{font-size:.74rem;color:#64748b;margin-top:2px}
  .top-actions{display:flex;align-items:center;gap:8px}
  .last-sync{font-size:.7rem;color:#475569}

  /* Buttons */
  .btn{border-radius:8px;padding:7px 14px;font-size:.8rem;cursor:pointer;
    display:flex;align-items:center;gap:6px;transition:all .2s;border:1px solid #334155;
    background:#1e293b;color:#94a3b8;white-space:nowrap}
  .btn:hover{background:#334155;color:#e2e8f0;border-color:#6366f1}
  .btn-upload{background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;
    border-color:transparent;box-shadow:0 2px 10px rgba(99,102,241,.4)}
  .btn-upload:hover{opacity:.9;border-color:transparent}
  .spin{animation:spin .7s linear infinite;display:inline-block}
  @keyframes spin{to{transform:rotate(360deg)}}

  /* Upload drop zone */
  .upload-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);
    z-index:100;align-items:center;justify-content:center}
  .upload-overlay.show{display:flex}
  .upload-box{background:#1e293b;border:2px dashed #6366f1;border-radius:16px;
    padding:48px 56px;text-align:center;max-width:420px;width:90%}
  .upload-box h2{color:#f1f5f9;font-size:1.1rem;margin-bottom:8px}
  .upload-box p{color:#64748b;font-size:.85rem;margin-bottom:24px}
  .upload-box.drag-over{background:#1e1b4b;border-color:#a5b4fc}
  .file-input{display:none}
  .upload-progress{display:none;margin-top:16px}
  .progress-bar{height:6px;background:#334155;border-radius:3px;overflow:hidden;margin-top:8px}
  .progress-fill{height:100%;background:linear-gradient(90deg,#6366f1,#8b5cf6);
    width:0%;transition:width .3s;border-radius:3px}
  .upload-status{font-size:.82rem;color:#94a3b8;margin-top:8px}
  .upload-close{position:absolute;top:16px;right:16px;background:none;border:none;
    color:#475569;font-size:1.4rem;cursor:pointer;padding:4px 8px;border-radius:6px}
  .upload-close:hover{color:#e2e8f0;background:#334155}

  /* Layout */
  .layout{display:flex;height:calc(100vh - 68px)}

  /* Sidebar */
  .sidebar{width:224px;min-width:224px;background:#1e293b;border-right:1px solid #334155;
    display:flex;flex-direction:column;padding:14px 10px;gap:4px;overflow-y:auto}
  .sidebar-label{font-size:.63rem;font-weight:700;letter-spacing:.12em;color:#475569;
    text-transform:uppercase;padding:0 8px 10px}
  .sheet-btn{background:transparent;border:1px solid transparent;border-radius:10px;
    padding:10px 12px;cursor:pointer;text-align:left;color:#94a3b8;font-size:.81rem;
    font-weight:500;transition:all .18s;line-height:1.4;width:100%}
  .sheet-btn:hover{background:#334155;color:#e2e8f0;border-color:#475569}
  .sheet-btn.active{background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;
    border-color:transparent;box-shadow:0 4px 14px rgba(99,102,241,.4)}
  .cnt{display:inline-block;font-size:.63rem;background:rgba(255,255,255,.18);
    border-radius:20px;padding:1px 7px;margin-left:4px;vertical-align:middle}

  /* Main */
  .main{flex:1;display:flex;flex-direction:column;overflow:hidden}
  .toolbar{background:#1a2235;border-bottom:1px solid #334155;padding:11px 20px;
    display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .toolbar h2{font-size:.93rem;font-weight:600;color:#f1f5f9;flex:1}
  .chips{display:flex;gap:7px}
  .chip{background:#334155;border-radius:7px;padding:3px 10px;font-size:.71rem;color:#94a3b8}
  .chip span{color:#a5b4fc;font-weight:600}
  .search-wrap{position:relative}
  .search-wrap input{background:#0f1117;border:1px solid #334155;border-radius:8px;
    padding:6px 12px 6px 30px;color:#e2e8f0;font-size:.81rem;width:200px;
    outline:none;transition:border-color .2s}
  .search-wrap input:focus{border-color:#6366f1}
  .si{position:absolute;left:8px;top:50%;transform:translateY(-50%);
    color:#475569;pointer-events:none;font-size:12px}

  /* Table */
  .table-wrap{flex:1;overflow:auto;padding:14px 18px}
  table{width:100%;border-collapse:collapse;font-size:.78rem}
  thead tr{position:sticky;top:0;z-index:10}
  thead th{background:#1e293b;border:1px solid #334155;padding:9px 11px;
    font-weight:600;color:#a5b4fc;white-space:nowrap;font-size:.72rem;
    text-transform:uppercase;letter-spacing:.04em;text-align:left}
  tbody tr:nth-child(even){background:#141b2d}
  tbody tr:hover{background:#1e3a5f}
  tbody td{border:1px solid #1e293b;padding:7px 11px;color:#cbd5e1;
    max-width:260px;word-break:break-word;vertical-align:top}
  .ec{color:#334155;font-style:italic;font-size:.7rem}

  /* States */
  .state-box{flex:1;display:flex;flex-direction:column;align-items:center;
    justify-content:center;gap:12px;color:#475569;text-align:center;padding:40px}
  .state-box .big{font-size:2.8rem}
  .state-box p{font-size:.88rem;line-height:1.6}
  .err{color:#f87171}
  .ok{color:#4ade80}

  ::-webkit-scrollbar{width:6px;height:6px}
  ::-webkit-scrollbar-track{background:#0f1117}
  ::-webkit-scrollbar-thumb{background:#334155;border-radius:4px}
</style>
</head>
<body>

<!-- Upload overlay -->
<div class="upload-overlay" id="upload-overlay" onclick="closeUpload(event)">
  <div class="upload-box" id="upload-box"
       ondragover="onDragOver(event)" ondragleave="onDragLeave(event)" ondrop="onDrop(event)">
    <button class="upload-close" onclick="closeUploadBtn()">✕</button>
    <div style="font-size:2.4rem;margin-bottom:12px">📂</div>
    <h2>Upload Excel File</h2>
    <p>Drag & drop your .xlsx file here,<br>or click to browse</p>
    <button class="btn btn-upload" onclick="document.getElementById('file-input').click()">
      📎 Choose File
    </button>
    <input type="file" id="file-input" class="file-input" accept=".xlsx,.xls" onchange="uploadFile(this.files[0])"/>
    <div class="upload-progress" id="upload-progress">
      <div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div>
      <div class="upload-status" id="upload-status">Uploading…</div>
    </div>
  </div>
</div>

<!-- Top bar -->
<div class="top-bar">
  <div class="logo">T</div>
  <div class="top-bar-text">
    <h1>Test Plan Dashboard <span class="live-badge">LIVE</span></h1>
    <div class="subtitle" id="subtitle">Loading…</div>
  </div>
  <div class="top-actions">
    <span class="last-sync" id="last-sync"></span>
    <button class="btn" onclick="loadData()"><span id="spin-icon">🔄</span> Refresh</button>
    <button class="btn btn-upload" onclick="openUpload()">⬆ Upload Excel</button>
  </div>
</div>

<!-- Layout -->
<div class="layout">
  <nav class="sidebar" id="sidebar"><div class="sidebar-label">Sheets</div></nav>
  <div class="main">
    <div class="state-box" id="state-box">
      <div class="big">⏳</div><p>Loading data from Excel…</p>
    </div>
    <div id="content" style="display:none;flex-direction:column;flex:1;overflow:hidden;">
      <div class="toolbar">
        <h2 id="sheet-title"></h2>
        <div class="chips">
          <div class="chip">Rows <span id="row-count">0</span></div>
          <div class="chip">Cols <span id="col-count">0</span></div>
        </div>
        <div class="search-wrap">
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

<script>
let DATA={}, rows=[], active=null;

/* ── Data loading ── */
async function loadData(){
  const icon=document.getElementById('spin-icon');
  icon.classList.add('spin');
  try{
    const r=await fetch('/data?t='+Date.now());
    if(!r.ok) throw new Error(await r.text());
    const json=await r.json();
    DATA=json.data;
    document.getElementById('subtitle').textContent=json.filename+' · Real-time from Excel';
    document.getElementById('last-sync').textContent='Synced '+new Date().toLocaleTimeString();
    buildSidebar();
    if(active&&DATA[active]) showSheet(active);
    else if(Object.keys(DATA).length){active=Object.keys(DATA)[0];buildSidebar();showSheet(active);}
  }catch(e){
    document.getElementById('state-box').innerHTML=
      '<div class="big">❌</div><p class="err">Cannot read Excel.<br><small>'+e.message+'</small></p>';
    document.getElementById('state-box').style.display='flex';
    document.getElementById('content').style.display='none';
  }finally{icon.classList.remove('spin');}
}

function buildSidebar(){
  const sb=document.getElementById('sidebar');
  sb.innerHTML='<div class="sidebar-label">Sheets</div>';
  Object.keys(DATA).forEach(s=>{
    const b=document.createElement('button');
    b.className='sheet-btn'+(s===active?' active':'');
    const icon=({Priority Test:'⭐','HW Test Plan':'🔧','FW Test Plan':'💾','Module HW':'📡','Comms Testing':'📶'}[s])||'📄';
    b.innerHTML=icon+' '+s+'<span class="cnt">'+DATA[s].rows.length+'</span>';
    b.onclick=()=>{active=s;buildSidebar();showSheet(s);};
    sb.appendChild(b);
  });
}

function showSheet(s){
  const{headers,rows:r}=DATA[s];rows=r;
  document.getElementById('sheet-title').textContent=s;
  document.getElementById('row-count').textContent=r.length;
  document.getElementById('col-count').textContent=headers.length;
  document.getElementById('q').value='';
  document.getElementById('thead').innerHTML='<tr>'+headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr>';
  render(r);
  document.getElementById('state-box').style.display='none';
  document.getElementById('content').style.display='flex';
}

function render(data){
  document.getElementById('tbody').innerHTML=data.map(r=>
    '<tr>'+r.map(c=>c?'<td>'+esc(c)+'</td>':'<td><span class="ec">—</span></td>').join('')+'</tr>'
  ).join('');
  document.getElementById('row-count').textContent=data.length;
}

function filter(){
  const q=document.getElementById('q').value.toLowerCase().trim();
  render(q?rows.filter(r=>r.some(c=>c.toLowerCase().includes(q))):rows);
}

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

/* ── Upload ── */
function openUpload(){document.getElementById('upload-overlay').classList.add('show');}
function closeUploadBtn(){document.getElementById('upload-overlay').classList.remove('show');resetUploadUI();}
function closeUpload(e){if(e.target===document.getElementById('upload-overlay'))closeUploadBtn();}

function onDragOver(e){e.preventDefault();document.getElementById('upload-box').classList.add('drag-over');}
function onDragLeave(){document.getElementById('upload-box').classList.remove('drag-over');}
function onDrop(e){
  e.preventDefault();
  document.getElementById('upload-box').classList.remove('drag-over');
  const f=e.dataTransfer.files[0];
  if(f) uploadFile(f);
}

async function uploadFile(file){
  if(!file||(!file.name.endsWith('.xlsx')&&!file.name.endsWith('.xls'))){
    alert('Please select an Excel file (.xlsx or .xls)');return;
  }
  const prog=document.getElementById('upload-progress');
  const fill=document.getElementById('progress-fill');
  const status=document.getElementById('upload-status');
  prog.style.display='block';
  fill.style.width='30%';
  status.textContent='Uploading '+file.name+'…';

  const form=new FormData();
  form.append('file',file);
  try{
    fill.style.width='60%';
    const r=await fetch('/upload',{method:'POST',body:form});
    fill.style.width='90%';
    if(!r.ok) throw new Error(await r.text());
    fill.style.width='100%';
    status.innerHTML='<span style="color:#4ade80">✓ Uploaded! Refreshing…</span>';
    setTimeout(async()=>{
      await loadData();
      closeUploadBtn();
    },800);
  }catch(e){
    status.innerHTML='<span style="color:#f87171">✗ '+e.message+'</span>';
    fill.style.width='0%';
  }
}

function resetUploadUI(){
  document.getElementById('upload-progress').style.display='none';
  document.getElementById('progress-fill').style.width='0%';
  document.getElementById('upload-status').textContent='';
  document.getElementById('file-input').value='';
}

loadData();
setInterval(loadData,30000);
</script>
</body>
</html>"""


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/data")
def data():
    try:
        result = read_excel()
        filename = Path(_excel_path[0]).name
        return jsonify({"filename": filename, "data": result})
    except Exception as e:
        return str(e), 500


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return "No file provided", 400
    f = request.files["file"]
    if not f.filename:
        return "Empty filename", 400
    if not (f.filename.endswith(".xlsx") or f.filename.endswith(".xls")):
        return "Only .xlsx / .xls files are supported", 400

    # Save to /tmp (persists for the life of the server process)
    save_path = Path("/tmp") / f.filename
    f.save(str(save_path))

    # Validate it's a real Excel file
    try:
        wb = openpyxl.load_workbook(save_path, read_only=True)
        wb.close()
    except Exception:
        save_path.unlink(missing_ok=True)
        return "Invalid Excel file", 400

    _excel_path[0] = save_path
    return jsonify({"ok": True, "filename": f.filename})


if __name__ == "__main__":
    app.run(debug=True, port=8765)
