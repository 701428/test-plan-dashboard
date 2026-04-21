"""
Flask app — serves the Test Plan Dashboard.
Reads the Excel file on every /data request (real-time).
"""

import json
from pathlib import Path
from flask import Flask, jsonify, render_template_string

import openpyxl

app = Flask(__name__)

BASE = Path(__file__).parent
EXCEL_FILE = BASE / "Abbreviation of Test Plan_1P _ Better Meter.xlsx"

SHEETS = {
    "Priority Test":                   ("Priority Test",  1, 2),
    "Abbreviation of Test Plan_HW":    ("HW Test Plan",   3, 5),
    "Abbreviation of Test Plan_FW":    ("FW Test Plan",   1, 2),
    "Abbreviation of TP_Module HW":    ("Module HW",      2, 3),
    "Abbreviation of TP_CommsTesting": ("Comms Testing",  2, 3),
}


def clean(val):
    if val is None:
        return ""
    return str(val).replace("\n", " ").strip()


def read_excel():
    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    result = {}

    for orig, (display, hdr_row, data_row) in SHEETS.items():
        if orig not in wb.sheetnames:
            continue
        ws = wb[orig]
        all_rows = list(ws.iter_rows(values_only=True))

        if orig == "Abbreviation of Test Plan_HW":
            row3 = [clean(v) for v in all_rows[2]]
            row4 = [clean(v) for v in all_rows[3]]
            headers = []
            for a, b in zip(row3, row4):
                parts = [x for x in [a, b] if x and x.lower() not in ("none", "false")]
                headers.append(" / ".join(parts) if parts else "")
        else:
            headers = [clean(v) for v in all_rows[hdr_row - 1]]

        rows = []
        for raw_row in all_rows[data_row - 1:]:
            cleaned = [clean(v) for v in raw_row]
            if any(cleaned):
                rows.append(cleaned)

        result[display] = {"headers": headers, "rows": rows}

    wb.close()
    return result


HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Test Plan Dashboard</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: #0f1117; color: #e2e8f0; min-height: 100vh; }

  .top-bar { background: linear-gradient(135deg,#1e293b,#0f172a); border-bottom:1px solid #334155;
    padding:16px 28px; display:flex; align-items:center; gap:14px; }
  .logo { width:40px;height:40px;background:linear-gradient(135deg,#6366f1,#8b5cf6);
    border-radius:10px;display:grid;place-items:center;font-size:20px;color:#fff;font-weight:700;flex-shrink:0; }
  .top-bar-text { flex:1; }
  .top-bar h1 { font-size:1.15rem;font-weight:700;color:#f1f5f9;
    display:flex;align-items:center;gap:8px; }
  .live-badge { font-size:.65rem;background:#166534;color:#4ade80;border:1px solid #166534;
    padding:2px 8px;border-radius:20px;font-weight:600;letter-spacing:.04em; }
  .top-bar .subtitle { font-size:.76rem;color:#64748b;margin-top:3px; }
  .sync-bar { display:flex;align-items:center;gap:10px; }
  .last-sync { font-size:.72rem;color:#475569; }
  .sync-btn { background:#1e293b;border:1px solid #334155;border-radius:8px;
    padding:7px 14px;color:#94a3b8;font-size:.8rem;cursor:pointer;
    display:flex;align-items:center;gap:6px;transition:all .2s; }
  .sync-btn:hover { background:#334155;color:#e2e8f0;border-color:#6366f1; }
  .spin { animation:spin .7s linear infinite;display:inline-block; }
  @keyframes spin { to { transform:rotate(360deg); } }

  .layout { display:flex; height:calc(100vh - 70px); }

  .sidebar { width:228px;min-width:228px;background:#1e293b;border-right:1px solid #334155;
    display:flex;flex-direction:column;padding:16px 12px;gap:5px;overflow-y:auto; }
  .sidebar-label { font-size:.65rem;font-weight:700;letter-spacing:.12em;color:#475569;
    text-transform:uppercase;padding:0 8px 10px; }
  .sheet-btn { background:transparent;border:1px solid transparent;border-radius:10px;
    padding:10px 13px;cursor:pointer;text-align:left;color:#94a3b8;font-size:.82rem;
    font-weight:500;transition:all .18s;line-height:1.4;width:100%; }
  .sheet-btn:hover { background:#334155;color:#e2e8f0;border-color:#475569; }
  .sheet-btn.active { background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;
    border-color:transparent;box-shadow:0 4px 14px rgba(99,102,241,.4); }
  .sheet-btn .cnt { display:inline-block;font-size:.65rem;background:rgba(255,255,255,.18);
    border-radius:20px;padding:1px 7px;margin-left:5px;vertical-align:middle; }

  .main { flex:1;display:flex;flex-direction:column;overflow:hidden; }
  .toolbar { background:#1a2235;border-bottom:1px solid #334155;padding:12px 22px;
    display:flex;align-items:center;gap:12px;flex-wrap:wrap; }
  .toolbar h2 { font-size:.95rem;font-weight:600;color:#f1f5f9;flex:1; }
  .chips { display:flex;gap:8px; }
  .chip { background:#334155;border-radius:7px;padding:4px 11px;font-size:.72rem;color:#94a3b8; }
  .chip span { color:#a5b4fc;font-weight:600; }
  .search-wrap { position:relative; }
  .search-wrap input { background:#0f1117;border:1px solid #334155;border-radius:8px;
    padding:6px 12px 6px 32px;color:#e2e8f0;font-size:.82rem;width:210px;
    outline:none;transition:border-color .2s; }
  .search-wrap input:focus { border-color:#6366f1; }
  .si { position:absolute;left:9px;top:50%;transform:translateY(-50%);color:#475569;pointer-events:none;font-size:13px; }

  .table-wrap { flex:1;overflow:auto;padding:16px 20px; }
  table { width:100%;border-collapse:collapse;font-size:.79rem; }
  thead tr { position:sticky;top:0;z-index:10; }
  thead th { background:#1e293b;border:1px solid #334155;padding:9px 12px;
    font-weight:600;color:#a5b4fc;white-space:nowrap;font-size:.73rem;
    text-transform:uppercase;letter-spacing:.04em;text-align:left; }
  tbody tr:nth-child(even) { background:#141b2d; }
  tbody tr:hover { background:#1e3a5f; }
  tbody td { border:1px solid #1e293b;padding:8px 12px;color:#cbd5e1;
    max-width:260px;word-break:break-word;vertical-align:top; }
  .ec { color:#334155;font-style:italic;font-size:.7rem; }

  .state-box { flex:1;display:flex;flex-direction:column;align-items:center;
    justify-content:center;gap:12px;color:#475569;text-align:center;padding:40px; }
  .state-box .big { font-size:2.8rem; }
  .state-box p { font-size:.88rem;line-height:1.6; }
  .err { color:#f87171; }

  ::-webkit-scrollbar { width:7px;height:7px; }
  ::-webkit-scrollbar-track { background:#0f1117; }
  ::-webkit-scrollbar-thumb { background:#334155;border-radius:4px; }
  ::-webkit-scrollbar-thumb:hover { background:#475569; }
</style>
</head>
<body>

<div class="top-bar">
  <div class="logo">T</div>
  <div class="top-bar-text">
    <h1>Test Plan Dashboard <span class="live-badge">LIVE</span></h1>
    <div class="subtitle">Abbreviation of Test Plan — 1P / Better Meter &nbsp;·&nbsp; Real-time from Excel</div>
  </div>
  <div class="sync-bar">
    <span class="last-sync" id="last-sync"></span>
    <button class="sync-btn" id="sync-btn" onclick="loadData(true)">
      <span id="spin-icon">🔄</span> Refresh
    </button>
  </div>
</div>

<div class="layout">
  <nav class="sidebar" id="sidebar">
    <div class="sidebar-label">Sheets</div>
  </nav>
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
const ICONS={"Priority Test":"⭐","HW Test Plan":"🔧","FW Test Plan":"💾","Module HW":"📡","Comms Testing":"📶"};
let DATA={}, rows=[], active=null;

async function loadData(manual=false) {
  const icon = document.getElementById('spin-icon');
  icon.classList.add('spin');
  try {
    const r = await fetch('/data?t='+Date.now());
    if (!r.ok) throw new Error(await r.text());
    DATA = await r.json();
    buildSidebar();
    document.getElementById('last-sync').textContent = 'Synced '+new Date().toLocaleTimeString();
    if (active && DATA[active]) showSheet(active);
    else if (Object.keys(DATA).length) { active=Object.keys(DATA)[0]; buildSidebar(); showSheet(active); }
  } catch(e) {
    document.getElementById('state-box').innerHTML =
      '<div class="big">❌</div><p class="err">Cannot read Excel file.<br><small>'+e.message+'</small></p>';
    document.getElementById('state-box').style.display='flex';
    document.getElementById('content').style.display='none';
  } finally { icon.classList.remove('spin'); }
}

function buildSidebar(){
  const sb=document.getElementById('sidebar');
  sb.innerHTML='<div class="sidebar-label">Sheets</div>';
  Object.keys(DATA).forEach(s=>{
    const b=document.createElement('button');
    b.className='sheet-btn'+(s===active?' active':'');
    b.innerHTML=`${ICONS[s]||'📄'} ${s}<span class="cnt">${DATA[s].rows.length}</span>`;
    b.onclick=()=>{ active=s; buildSidebar(); showSheet(s); };
    sb.appendChild(b);
  });
}

function showSheet(s){
  const {headers,rows:r}=DATA[s]; rows=r;
  document.getElementById('sheet-title').textContent=s;
  document.getElementById('row-count').textContent=r.length;
  document.getElementById('col-count').textContent=headers.length;
  document.getElementById('q').value='';
  document.getElementById('thead').innerHTML='<tr>'+headers.map(h=>`<th>${esc(h)}</th>`).join('')+'</tr>';
  render(r);
  document.getElementById('state-box').style.display='none';
  document.getElementById('content').style.display='flex';
}

function render(data){
  document.getElementById('tbody').innerHTML=data.map(r=>
    '<tr>'+r.map(c=>c?`<td>${esc(c)}</td>`:`<td><span class="ec">—</span></td>`).join('')+'</tr>'
  ).join('');
  document.getElementById('row-count').textContent=data.length;
}

function filter(){
  const q=document.getElementById('q').value.toLowerCase().trim();
  render(q?rows.filter(r=>r.some(c=>c.toLowerCase().includes(q))):rows);
}

function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

loadData();
setInterval(loadData, 30000);
</script>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/data")
def data():
    try:
        return jsonify(read_excel())
    except Exception as e:
        return str(e), 500


if __name__ == "__main__":
    app.run(debug=True, port=8765)
