#!/usr/bin/env python3
"""
Excel SQL Explorer
==================
Query Excel/CSV files with SQL — CLI + Web Interface (offline, no internet needed).

Usage:
  Web mode (default):   python excel_sql.py
  CLI mode:             python excel_sql.py --cli
  Custom port/host:     python excel_sql.py --port 8080 --host 0.0.0.0
"""

import argparse
import io
import json
import os
import re
import sqlite3
import sys
import threading
import traceback
import webbrowser
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template_string, request, send_file

# ─────────────────────────────────────────────
#  In-memory state
# ─────────────────────────────────────────────
DB_CONN: sqlite3.Connection | None = None
LOADED_FILE: str = ""
SHEETS: list[str] = []
QUERY_HISTORY: list[dict] = []

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB upload limit


# ─────────────────────────────────────────────
#  Core logic
# ─────────────────────────────────────────────

def sanitize_name(name: str) -> str:
    """Convert a sheet/column name to a safe SQL identifier."""
    name = re.sub(r"[^\w]", "_", str(name).strip())
    if name and name[0].isdigit():
        name = "t_" + name
    return name or "col"


def load_file(filepath: str) -> dict:
    """Load Excel or CSV into an in-memory SQLite DB. Returns info dict."""
    global DB_CONN, LOADED_FILE, SHEETS

    ext = Path(filepath).suffix.lower()
    try:
        if ext == ".csv":
            sheets = {"data": pd.read_csv(filepath)}
        elif ext in (".xlsx", ".xls", ".xlsm", ".xlsb", ".ods"):
            sheets = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
        else:
            return {"ok": False, "error": f"Unsupported file type: {ext}"}
    except Exception as e:
        return {"ok": False, "error": str(e)}

    DB_CONN = sqlite3.connect(":memory:", check_same_thread=False)
    SHEETS = []
    table_info = []

    for raw_name, df in sheets.items():
        table_name = sanitize_name(raw_name)
        # Sanitize column names
        df.columns = [sanitize_name(c) + (f"_{i}" if list(df.columns).count(c) > 1 else "")
                      for i, c in enumerate(df.columns)]
        df.to_sql(table_name, DB_CONN, if_exists="replace", index=False)
        SHEETS.append(table_name)
        table_info.append({
            "name": table_name,
            "original": str(raw_name),
            "rows": len(df),
            "columns": list(df.columns),
        })

    LOADED_FILE = os.path.basename(filepath)
    return {"ok": True, "tables": table_info, "file": LOADED_FILE}


def run_query(sql: str, limit: int = 500) -> dict:
    """Execute SQL and return results as dict."""
    global DB_CONN, QUERY_HISTORY

    if DB_CONN is None:
        return {"ok": False, "error": "No file loaded. Upload an Excel/CSV file first."}
    if not sql.strip():
        return {"ok": False, "error": "Empty query."}

    try:
        cursor = DB_CONN.cursor()
        cursor.execute(sql)

        is_select = sql.strip().upper().startswith("SELECT")
        if is_select:
            rows = cursor.fetchmany(limit)
            cols = [d[0] for d in cursor.description] if cursor.description else []
            total = len(rows)
            entry = {
                "sql": sql,
                "ok": True,
                "columns": cols,
                "rows": [list(r) for r in rows],
                "count": total,
                "truncated": total >= limit,
            }
        else:
            DB_CONN.commit()
            entry = {
                "sql": sql,
                "ok": True,
                "columns": [],
                "rows": [],
                "count": cursor.rowcount,
                "message": f"Query OK. {cursor.rowcount} row(s) affected.",
            }

        QUERY_HISTORY.insert(0, {"sql": sql, "ok": True})
        if len(QUERY_HISTORY) > 50:
            QUERY_HISTORY.pop()
        return entry

    except Exception as e:
        entry = {"sql": sql, "ok": False, "error": str(e)}
        QUERY_HISTORY.insert(0, {"sql": sql, "ok": False})
        return entry


def get_schema() -> list[dict]:
    """Return table schema from the loaded DB."""
    if DB_CONN is None:
        return []
    cursor = DB_CONN.cursor()
    result = []
    for table in SHEETS:
        cursor.execute(f"PRAGMA table_info({table})")
        cols = [{"name": row[1], "type": row[2]} for row in cursor.fetchall()]
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        row_count = cursor.fetchone()[0]
        result.append({"table": table, "columns": cols, "rows": row_count})
    return result


# ─────────────────────────────────────────────
#  Flask Routes
# ─────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/api/upload", methods=["POST"])
def api_upload():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file part"})
    f = request.files["file"]
    if not f.filename:
        return jsonify({"ok": False, "error": "No file selected"})

    tmp_path = f"/tmp/{f.filename}"
    f.save(tmp_path)
    result = load_file(tmp_path)
    return jsonify(result)


@app.route("/api/query", methods=["POST"])
def api_query():
    data = request.get_json(force=True)
    sql = data.get("sql", "")
    limit = int(data.get("limit", 500))
    result = run_query(sql, limit)
    return jsonify(result)


@app.route("/api/schema")
def api_schema():
    return jsonify(get_schema())


@app.route("/api/history")
def api_history():
    return jsonify(QUERY_HISTORY[:20])


@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.get_json(force=True)
    sql = data.get("sql", "")
    result = run_query(sql, limit=100_000)
    if not result["ok"]:
        return jsonify(result), 400
    df = pd.DataFrame(result["rows"], columns=result["columns"])
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    buf.seek(0)
    return send_file(
        io.BytesIO(buf.getvalue().encode()),
        mimetype="text/csv",
        as_attachment=True,
        download_name="query_result.csv",
    )


# ─────────────────────────────────────────────
#  HTML Template (single-file embed)
# ─────────────────────────────────────────────

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>Excel SQL Explorer</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
  :root {
    --bg:       #0d1117;
    --surface:  #161b22;
    --surface2: #1c2230;
    --border:   #2a3444;
    --accent:   #f0b429;
    --accent2:  #e05c3a;
    --green:    #3ddc97;
    --text:     #e6edf3;
    --muted:    #7d8fa1;
    --danger:   #e05c3a;
    --radius:   10px;
    --mono:     'JetBrains Mono', monospace;
    --sans:     'Sora', sans-serif;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html, body { height: 100%; background: var(--bg); color: var(--text); font-family: var(--sans); font-size: 14px; }

  /* ── Layout ── */
  .app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

  header {
    display: flex; align-items: center; gap: 12px; padding: 0 20px;
    height: 52px; background: var(--surface); border-bottom: 1px solid var(--border);
    flex-shrink: 0;
  }
  .logo { display: flex; align-items: center; gap: 8px; }
  .logo-icon {
    width: 30px; height: 30px; border-radius: 7px;
    background: linear-gradient(135deg, var(--accent) 0%, var(--accent2) 100%);
    display: flex; align-items: center; justify-content: center;
    font-size: 16px;
  }
  .logo-text { font-size: 15px; font-weight: 700; letter-spacing: -0.3px; }
  .logo-text span { color: var(--accent); }
  header .spacer { flex: 1; }
  .file-badge {
    font-family: var(--mono); font-size: 11px; color: var(--green);
    background: rgba(61,220,151,.1); border: 1px solid rgba(61,220,151,.25);
    padding: 3px 10px; border-radius: 20px; display: none; max-width: 220px;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }

  .main { display: flex; flex: 1; overflow: hidden; }

  /* ── Sidebar ── */
  .sidebar {
    width: 240px; background: var(--surface); border-right: 1px solid var(--border);
    display: flex; flex-direction: column; overflow: hidden; flex-shrink: 0;
    transition: width .25s;
  }
  .sidebar-head {
    padding: 14px 16px 10px; font-size: 11px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 1px; color: var(--muted);
    border-bottom: 1px solid var(--border);
  }
  .schema-area { flex: 1; overflow-y: auto; padding: 8px 0; }
  .table-entry { padding: 6px 16px; cursor: pointer; user-select: none; }
  .table-entry:hover .tname { color: var(--accent); }
  .tname {
    font-family: var(--mono); font-size: 12px; font-weight: 500;
    color: var(--text); display: flex; align-items: center; gap: 6px;
  }
  .tname::before { content: '▶'; font-size: 9px; color: var(--muted); }
  .tname.open::before { content: '▼'; }
  .trows { font-size: 10px; color: var(--muted); margin-top: 1px; margin-left: 14px; }
  .col-list { display: none; padding: 4px 0 4px 14px; }
  .col-list.visible { display: block; }
  .col-item {
    font-family: var(--mono); font-size: 11px; color: var(--muted);
    padding: 2px 0; display: flex; gap: 8px; align-items: center;
  }
  .col-item span { color: var(--green); font-size: 10px; }
  .sidebar-history { border-top: 1px solid var(--border); padding: 8px 0; }
  .hist-head { padding: 8px 16px 4px; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: var(--muted); }
  .hist-item {
    padding: 5px 16px; font-family: var(--mono); font-size: 11px; color: var(--muted);
    cursor: pointer; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }
  .hist-item:hover { color: var(--accent); }
  .hist-item.err { color: var(--danger); }

  /* ── Editor + Results panel ── */
  .center { flex: 1; display: flex; flex-direction: column; overflow: hidden; min-width: 0; }

  .upload-zone {
    flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center;
    gap: 16px; padding: 40px;
  }
  .upload-box {
    width: 100%; max-width: 440px; border: 2px dashed var(--border); border-radius: 16px;
    padding: 40px 30px; text-align: center; cursor: pointer; transition: all .2s;
    background: var(--surface);
  }
  .upload-box:hover, .upload-box.drag { border-color: var(--accent); background: rgba(240,180,41,.04); }
  .upload-icon { font-size: 42px; margin-bottom: 12px; }
  .upload-title { font-size: 17px; font-weight: 700; margin-bottom: 6px; }
  .upload-sub { color: var(--muted); font-size: 13px; }
  .upload-exts { margin-top: 16px; display: flex; gap: 8px; justify-content: center; flex-wrap: wrap; }
  .ext-badge {
    font-family: var(--mono); font-size: 11px; background: var(--surface2);
    border: 1px solid var(--border); border-radius: 5px; padding: 3px 8px; color: var(--muted);
  }
  #fileInput { display: none; }

  .editor-panel { display: none; flex-direction: column; flex: 1; overflow: hidden; }

  .sql-area {
    padding: 12px 16px; border-bottom: 1px solid var(--border);
    background: var(--surface); flex-shrink: 0;
  }
  .sql-label {
    font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase;
    letter-spacing: 1px; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center;
  }
  .sql-actions { display: flex; gap: 8px; }
  textarea#sqlInput {
    width: 100%; background: var(--bg); color: var(--text);
    border: 1px solid var(--border); border-radius: var(--radius);
    font-family: var(--mono); font-size: 13px; padding: 12px 14px;
    resize: vertical; min-height: 100px; max-height: 260px;
    outline: none; transition: border .2s;
  }
  textarea#sqlInput:focus { border-color: var(--accent); }
  .run-bar { display: flex; gap: 8px; margin-top: 10px; align-items: center; flex-wrap: wrap; }

  .btn {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 8px 16px; border-radius: 8px; font-family: var(--sans);
    font-size: 13px; font-weight: 600; cursor: pointer; border: none; transition: all .15s;
  }
  .btn-primary { background: var(--accent); color: #0d1117; }
  .btn-primary:hover { background: #f7c94e; }
  .btn-ghost {
    background: transparent; color: var(--muted);
    border: 1px solid var(--border);
  }
  .btn-ghost:hover { color: var(--text); border-color: var(--muted); }
  .btn-danger { background: var(--danger); color: #fff; }
  .btn-danger:hover { background: #ea7558; }
  .btn:disabled { opacity: .4; cursor: not-allowed; }

  .limit-select {
    font-family: var(--mono); font-size: 12px; background: var(--bg);
    color: var(--muted); border: 1px solid var(--border); border-radius: 6px;
    padding: 7px 10px; cursor: pointer;
  }
  .limit-select:focus { outline: none; border-color: var(--accent); }

  .results-area { flex: 1; display: flex; flex-direction: column; overflow: hidden; }

  .results-header {
    padding: 8px 16px; background: var(--surface); border-bottom: 1px solid var(--border);
    display: flex; align-items: center; gap: 10px; font-size: 12px;
    color: var(--muted); flex-wrap: wrap; flex-shrink: 0;
  }
  .result-stat { display: flex; align-items: center; gap: 4px; }
  .stat-val { color: var(--accent); font-family: var(--mono); font-weight: 600; }
  .result-spacer { flex: 1; }

  .table-wrap { flex: 1; overflow: auto; }
  table { width: max-content; min-width: 100%; border-collapse: collapse; }
  thead tr { position: sticky; top: 0; z-index: 2; }
  thead th {
    background: var(--surface2); padding: 9px 14px; text-align: left;
    font-family: var(--mono); font-size: 11px; font-weight: 600;
    color: var(--muted); border-bottom: 2px solid var(--border);
    white-space: nowrap; user-select: none; letter-spacing: .3px;
  }
  tbody tr { border-bottom: 1px solid rgba(42,52,68,.6); }
  tbody tr:hover { background: rgba(240,180,41,.04); }
  tbody td {
    padding: 7px 14px; font-family: var(--mono); font-size: 12px;
    color: var(--text); white-space: nowrap; max-width: 300px;
    overflow: hidden; text-overflow: ellipsis;
  }
  td.null { color: var(--muted); font-style: italic; }
  td.num { color: #82aaff; }
  td.str { color: var(--green); }

  .error-box {
    margin: 20px; padding: 16px 20px; background: rgba(224,92,58,.1);
    border: 1px solid rgba(224,92,58,.3); border-radius: var(--radius);
    color: var(--danger); font-family: var(--mono); font-size: 13px;
  }
  .empty-state {
    flex: 1; display: flex; align-items: center; justify-content: center;
    color: var(--muted); flex-direction: column; gap: 8px; padding: 40px;
  }
  .empty-state .big { font-size: 36px; }

  .spinner {
    width: 18px; height: 18px; border: 2px solid var(--border);
    border-top-color: var(--accent); border-radius: 50%;
    animation: spin .6s linear infinite; display: none;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  .status-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--muted); }
  .status-dot.ok { background: var(--green); }
  .status-dot.err { background: var(--danger); }

  /* ── Mobile toggle ── */
  .sidebar-toggle {
    display: none; background: none; border: none; color: var(--muted);
    font-size: 20px; cursor: pointer; padding: 4px 8px;
  }
  .overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.5); z-index: 10; }

  /* ── Truncated notice ── */
  .truncate-notice {
    text-align: center; padding: 10px; font-size: 12px; color: var(--accent);
    background: rgba(240,180,41,.05); border-top: 1px solid var(--border);
    font-family: var(--mono);
  }

  /* ── Responsive ── */
  @media (max-width: 700px) {
    .sidebar {
      position: fixed; left: 0; top: 52px; bottom: 0; z-index: 20;
      transform: translateX(-100%); transition: transform .25s;
      width: 260px;
    }
    .sidebar.open { transform: translateX(0); }
    .overlay.open { display: block; }
    .sidebar-toggle { display: block; }
    .sql-area { padding: 10px 12px; }
    textarea#sqlInput { min-height: 80px; font-size: 12px; }
    .logo-text { font-size: 13px; }
    .file-badge { max-width: 130px; font-size: 10px; }
  }
</style>
</head>
<body>
<div class="app">
  <!-- Header -->
  <header>
    <button class="sidebar-toggle" id="sidebarToggle">☰</button>
    <div class="logo">
      <div class="logo-icon">⚡</div>
      <div class="logo-text">Excel<span>SQL</span></div>
    </div>
    <div class="spacer"></div>
    <div class="file-badge" id="fileBadge"></div>
    <label for="fileInput" class="btn btn-ghost" style="padding:6px 12px;font-size:12px;cursor:pointer">
      📂 Load File
    </label>
    <input type="file" id="fileInput" accept=".xlsx,.xls,.xlsm,.xlsb,.ods,.csv">
  </header>

  <div class="main">
    <!-- Sidebar: Schema + History -->
    <aside class="sidebar" id="sidebar">
      <div class="sidebar-head">Schema</div>
      <div class="schema-area" id="schemaArea">
        <div style="padding:20px 16px;color:var(--muted);font-size:12px">
          Upload a file to see schema
        </div>
      </div>
      <div class="sidebar-history">
        <div class="hist-head">History</div>
        <div id="historyList"></div>
      </div>
    </aside>
    <div class="overlay" id="overlay"></div>

    <!-- Center: Editor + Results -->
    <div class="center" id="center">

      <!-- Upload zone -->
      <div class="upload-zone" id="uploadZone">
        <div class="upload-box" id="dropZone">
          <div class="upload-icon">📊</div>
          <div class="upload-title">Drop your Excel or CSV here</div>
          <div class="upload-sub">or click the "Load File" button above</div>
          <div class="upload-exts">
            <span class="ext-badge">.xlsx</span>
            <span class="ext-badge">.xls</span>
            <span class="ext-badge">.xlsm</span>
            <span class="ext-badge">.ods</span>
            <span class="ext-badge">.csv</span>
          </div>
        </div>
      </div>

      <!-- Editor + Results (hidden until file loaded) -->
      <div class="editor-panel" id="editorPanel">

        <div class="sql-area">
          <div class="sql-label">
            <span>SQL Query</span>
            <div class="sql-actions">
              <button class="btn btn-ghost" style="padding:4px 10px;font-size:11px" id="clearBtn">Clear</button>
            </div>
          </div>
          <textarea id="sqlInput" placeholder="SELECT * FROM your_table LIMIT 10" spellcheck="false"></textarea>
          <div class="run-bar">
            <button class="btn btn-primary" id="runBtn">▶ Run Query</button>
            <button class="btn btn-ghost" id="exportBtn">⬇ Export CSV</button>
            <select class="limit-select" id="limitSelect">
              <option value="100">100 rows</option>
              <option value="500" selected>500 rows</option>
              <option value="1000">1 000 rows</option>
              <option value="5000">5 000 rows</option>
              <option value="50000">All rows</option>
            </select>
            <div class="spinner" id="spinner"></div>
            <div style="flex:1"></div>
            <div style="font-size:11px;color:var(--muted)">Ctrl+Enter to run</div>
          </div>
        </div>

        <div class="results-area" id="resultsArea">
          <div class="empty-state">
            <div class="big">💡</div>
            <div>Write a query above and press <strong>Run</strong></div>
          </div>
        </div>

      </div>
    </div>
  </div>
</div>

<script>
const $ = id => document.getElementById(id);

// ── State ──
let currentSQL = '';

// ── Upload handling ──
$('fileInput').addEventListener('change', e => {
  if (e.target.files[0]) uploadFile(e.target.files[0]);
});

$('dropZone').addEventListener('click', () => $('fileInput').click());

['dragenter','dragover'].forEach(ev =>
  $('dropZone').addEventListener(ev, e => { e.preventDefault(); $('dropZone').classList.add('drag'); })
);
['dragleave','drop'].forEach(ev =>
  $('dropZone').addEventListener(ev, e => { e.preventDefault(); $('dropZone').classList.remove('drag'); })
);
$('dropZone').addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (file) uploadFile(file);
});

async function uploadFile(file) {
  const fd = new FormData();
  fd.append('file', file);
  showSpinner(true);
  try {
    const res = await fetch('/api/upload', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.ok) {
      $('fileBadge').textContent = '● ' + data.file;
      $('fileBadge').style.display = 'block';
      $('uploadZone').style.display = 'none';
      $('editorPanel').style.display = 'flex';
      loadSchema();
      // Auto-populate first SELECT
      if (data.tables && data.tables.length > 0) {
        $('sqlInput').value = `SELECT * FROM ${data.tables[0].name} LIMIT 20`;
      }
    } else {
      alert('Error loading file: ' + data.error);
    }
  } catch(e) { alert('Upload failed: ' + e); }
  showSpinner(false);
}

// ── Schema ──
async function loadSchema() {
  const res = await fetch('/api/schema');
  const tables = await res.json();
  const el = $('schemaArea');
  if (!tables.length) { el.innerHTML = '<div style="padding:16px;color:var(--muted);font-size:12px">No tables found</div>'; return; }
  el.innerHTML = tables.map(t => `
    <div class="table-entry">
      <div class="tname" onclick="toggleCols(this,'${t.table}')" data-table="${t.table}">
        ${t.table}
      </div>
      <div class="trows">${t.rows.toLocaleString()} rows · ${t.columns.length} cols</div>
      <div class="col-list" id="cols_${t.table}">
        ${t.columns.map(c => `<div class="col-item">${c.name}<span>${c.type||'TEXT'}</span></div>`).join('')}
      </div>
    </div>
  `).join('');
}

function toggleCols(el, table) {
  const list = document.getElementById('cols_' + table);
  const open = list.classList.toggle('visible');
  el.classList.toggle('open', open);
  if (open) {
    // Double-click to insert table name into query
    el.ondblclick = () => insertAtCursor($('sqlInput'), table);
  }
}

function insertAtCursor(el, text) {
  const s = el.selectionStart, e = el.selectionEnd;
  el.value = el.value.slice(0, s) + text + el.value.slice(e);
  el.setSelectionRange(s + text.length, s + text.length);
  el.focus();
}

// ── Query ──
$('runBtn').addEventListener('click', runQuery);
$('clearBtn').addEventListener('click', () => { $('sqlInput').value = ''; $('sqlInput').focus(); });

$('sqlInput').addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); runQuery(); }
  // Tab to indent
  if (e.key === 'Tab') {
    e.preventDefault();
    insertAtCursor($('sqlInput'), '    ');
  }
});

async function runQuery() {
  const sql = $('sqlInput').value.trim();
  if (!sql) return;
  currentSQL = sql;
  showSpinner(true);
  $('runBtn').disabled = true;
  try {
    const limit = parseInt($('limitSelect').value);
    const res = await fetch('/api/query', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ sql, limit })
    });
    const data = await res.json();
    renderResults(data);
    loadHistory();
  } catch(e) { renderError('Network error: ' + e); }
  showSpinner(false);
  $('runBtn').disabled = false;
}

function renderResults(data) {
  const el = $('resultsArea');
  if (!data.ok) { el.innerHTML = `<div class="error-box">❌ ${escHtml(data.error)}</div>`; return; }
  if (!data.columns || data.columns.length === 0) {
    el.innerHTML = `<div class="empty-state"><div class="big">✅</div><div>${escHtml(data.message || 'Query executed.')}</div></div>`;
    return;
  }
  const time = new Date().toLocaleTimeString();
  let html = `
    <div class="results-header">
      <div class="status-dot ok"></div>
      <div class="result-stat"><span class="stat-val">${data.count.toLocaleString()}</span>&nbsp;rows</div>
      <div class="result-stat"><span class="stat-val">${data.columns.length}</span>&nbsp;cols</div>
      <div class="result-spacer"></div>
      <div style="font-size:11px;font-family:var(--mono)">${time}</div>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr>${data.columns.map(c => `<th>${escHtml(c)}</th>`).join('')}</tr></thead>
        <tbody>
          ${data.rows.map(row => `<tr>${row.map(cell => renderCell(cell)).join('')}</tr>`).join('')}
        </tbody>
      </table>
    </div>
  `;
  if (data.truncated) {
    html += `<div class="truncate-notice">⚠ Showing first ${data.count.toLocaleString()} rows. Use LIMIT or change row limit to see more.</div>`;
  }
  el.innerHTML = html;
}

function renderCell(val) {
  if (val === null || val === undefined || val === '') return `<td class="null">NULL</td>`;
  const escaped = escHtml(String(val));
  if (typeof val === 'number') return `<td class="num">${escaped}</td>`;
  return `<td class="str">${escaped}</td>`;
}

function renderError(msg) {
  $('resultsArea').innerHTML = `<div class="error-box">❌ ${escHtml(msg)}</div>`;
}

// ── Export ──
$('exportBtn').addEventListener('click', async () => {
  const sql = $('sqlInput').value.trim();
  if (!sql) return;
  const res = await fetch('/api/export', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ sql })
  });
  if (!res.ok) { alert('Export failed'); return; }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'query_result.csv'; a.click();
  URL.revokeObjectURL(url);
});

// ── History ──
async function loadHistory() {
  const res = await fetch('/api/history');
  const hist = await res.json();
  const el = $('historyList');
  if (!hist.length) { el.innerHTML = ''; return; }
  el.innerHTML = hist.slice(0, 10).map(h =>
    `<div class="hist-item ${h.ok?'':'err'}" title="${escHtml(h.sql)}"
      onclick="$('sqlInput').value=${JSON.stringify(h.sql)}">${escHtml(truncate(h.sql, 40))}</div>`
  ).join('');
}

// ── Sidebar mobile ──
$('sidebarToggle').addEventListener('click', () => {
  $('sidebar').classList.toggle('open');
  $('overlay').classList.toggle('open');
});
$('overlay').addEventListener('click', () => {
  $('sidebar').classList.remove('open');
  $('overlay').classList.remove('open');
});

// ── Helpers ──
function showSpinner(on) { $('spinner').style.display = on ? 'block' : 'none'; }
function escHtml(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function truncate(s, n) { return s.length > n ? s.slice(0, n) + '…' : s; }
</script>
</body>
</html>
"""


# ─────────────────────────────────────────────
#  CLI Mode
# ─────────────────────────────────────────────

def cli_mode():
    print("\n╔══════════════════════════════════════╗")
    print("║       Excel SQL Explorer  (CLI)      ║")
    print("╚══════════════════════════════════════╝\n")

    filepath = input("📂 Enter path to Excel/CSV file: ").strip().strip('"').strip("'")
    if not os.path.exists(filepath):
        print(f"❌  File not found: {filepath}")
        sys.exit(1)

    print(f"\n⏳ Loading {filepath} …")
    result = load_file(filepath)
    if not result["ok"]:
        print(f"❌  {result['error']}")
        sys.exit(1)

    print(f"\n✅  Loaded: {result['file']}")
    for t in result["tables"]:
        print(f"   📋 {t['name']}  ·  {t['rows']} rows  ·  {len(t['columns'])} columns")
        print(f"      columns: {', '.join(t['columns'][:8])}{'…' if len(t['columns'])>8 else ''}")

    print("\n💡 Type SQL queries. Commands: .schema  .history  .quit\n")

    while True:
        try:
            sql = input("sql> ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n👋  Bye!")
            break

        if not sql:
            continue
        if sql.lower() in (".quit", ".exit", "quit", "exit"):
            print("👋  Bye!")
            break
        if sql.lower() == ".schema":
            for t in get_schema():
                print(f"\n  Table: {t['table']}  ({t['rows']} rows)")
                for c in t["columns"]:
                    print(f"    {c['name']}  {c['type']}")
            continue
        if sql.lower() == ".history":
            for i, h in enumerate(QUERY_HISTORY[:10], 1):
                mark = "✓" if h["ok"] else "✗"
                print(f"  {mark} {i}. {h['sql'][:80]}")
            continue

        result = run_query(sql)
        if not result["ok"]:
            print(f"❌  {result['error']}")
        elif not result["columns"]:
            print(f"✅  {result.get('message', 'OK')}")
        else:
            cols = result["columns"]
            rows = result["rows"]
            # Calculate widths
            widths = [len(c) for c in cols]
            for row in rows:
                for i, v in enumerate(row):
                    widths[i] = min(max(widths[i], len(str(v) if v is not None else "NULL")), 30)
            sep = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
            fmt = "| " + " | ".join("{:<{w}}" for w in widths) + " |"
            print(sep)
            print(fmt.format(*cols, w=0).replace("w=0",""))  # header
            # Actually format properly:
            header_cells = [c[:w].ljust(w) for c, w in zip(cols, widths)]
            print("| " + " | ".join(header_cells) + " |")
            print(sep)
            for row in rows:
                cells = [str(v if v is not None else "NULL")[:w].ljust(w) for v, w in zip(row, widths)]
                print("| " + " | ".join(cells) + " |")
            print(sep)
            note = f"  {result['count']} row(s)"
            if result.get("truncated"):
                note += f"  [TRUNCATED at 500 — add LIMIT to query]"
            print(note)


# ─────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Excel SQL Explorer")
    parser.add_argument("--cli", action="store_true", help="Run in CLI (terminal) mode")
    parser.add_argument("--host", default="0.0.0.0", help="Web server host (default: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=5050, help="Web server port (default: 5050)")
    parser.add_argument("--no-browser", action="store_true", help="Don't auto-open browser")
    args = parser.parse_args()

    if args.cli:
        cli_mode()
    else:
        print("\n╔══════════════════════════════════════════╗")
        print("║     Excel SQL Explorer  ·  Web Mode      ║")
        print("╚══════════════════════════════════════════╝")
        print(f"\n  Local:    http://localhost:{args.port}")
        print(f"  Network:  http://<your-ip>:{args.port}")
        print(f"\n  Open the Local URL in any browser (PC or mobile on same WiFi)")
        print("  Press Ctrl+C to stop\n")
        if not args.no_browser:
            threading.Timer(0.8, lambda: webbrowser.open(f"http://localhost:{args.port}")).start()
        app.run(host=args.host, port=args.port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
