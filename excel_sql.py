#!/usr/bin/env python3
"""
Excel SQL Explorer  v3.1
========================
Query Excel/CSV files with SQL  +  Natural Language → SQL via LOCAL AI (Ollama).

100% offline — no API keys, no internet, no cloud.

Setup:
  pip install flask pandas openpyxl

Local AI (optional):
  1. Install Ollama:  https://ollama.com
  2. ollama pull llama3          (or: deepseek-coder / mistral / phi3 / codellama)
  3. ollama serve                (starts on port 11434)

Run:
  Web mode:   python excel_sql.py
  CLI mode:   python excel_sql.py --cli
  Custom:     python excel_sql.py --port 8080
"""

import argparse
import io
import json
import os
import re
import sqlite3
import sys
import tempfile
import threading
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template_string, request, send_file

# ─────────────────────────────────────────────
#  State
# ─────────────────────────────────────────────
DB_CONN: sqlite3.Connection | None = None
LOADED_FILE: str = ""
SHEETS: list[str] = []
QUERY_HISTORY: list[dict] = []
NL_HISTORY: list[dict] = []

OLLAMA_URL: str = "http://localhost:11434"
SELECTED_MODEL: str = ""

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024   # 100 MB


# ─────────────────────────────────────────────
#  Global error handlers — always return JSON
# ─────────────────────────────────────────────

@app.errorhandler(400)
def bad_request(e):
    return jsonify({"ok": False, "error": f"Bad request: {e}"}), 400

@app.errorhandler(404)
def not_found(e):
    return jsonify({"ok": False, "error": "Endpoint not found"}), 404

@app.errorhandler(413)
def too_large(e):
    return jsonify({"ok": False, "error": "File too large (max 100 MB)"}), 413

@app.errorhandler(500)
def server_error(e):
    return jsonify({"ok": False, "error": f"Server error: {e}"}), 500

@app.errorhandler(Exception)
def unhandled(e):
    return jsonify({"ok": False, "error": str(e)}), 500


# ─────────────────────────────────────────────
#  File loading
# ─────────────────────────────────────────────

def sanitize_name(name: str) -> str:
    name = re.sub(r"[^\w]", "_", str(name).strip())
    if name and name[0].isdigit():
        name = "t_" + name
    return name or "col"


def load_file(filepath: str) -> dict:
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
        return {"ok": False, "error": f"Could not read file: {e}"}

    DB_CONN = sqlite3.connect(":memory:", check_same_thread=False)
    SHEETS = []
    table_info = []

    for raw_name, df in sheets.items():
        table_name = sanitize_name(raw_name)
        seen: dict[str, int] = {}
        new_cols = []
        for c in df.columns:
            sc = sanitize_name(c)
            count = seen.get(sc, 0)
            new_cols.append(sc if count == 0 else f"{sc}_{count}")
            seen[sc] = count + 1
        df.columns = new_cols
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


# ─────────────────────────────────────────────
#  SQL execution
# ─────────────────────────────────────────────

def run_query(sql: str, limit: int = 500) -> dict:
    global DB_CONN, QUERY_HISTORY
    if DB_CONN is None:
        return {"ok": False, "error": "No file loaded yet."}
    if not sql.strip():
        return {"ok": False, "error": "Empty query."}
    try:
        cursor = DB_CONN.cursor()
        cursor.execute(sql)
        upper = sql.strip().upper()
        if upper.startswith("SELECT") or upper.startswith("WITH"):
            rows = cursor.fetchmany(limit)
            cols = [d[0] for d in cursor.description] if cursor.description else []
            entry = {"sql": sql, "ok": True, "columns": cols,
                     "rows": [list(r) for r in rows],
                     "count": len(rows), "truncated": len(rows) >= limit}
        else:
            DB_CONN.commit()
            entry = {"sql": sql, "ok": True, "columns": [], "rows": [],
                     "count": cursor.rowcount,
                     "message": f"Query OK. {cursor.rowcount} row(s) affected."}
        QUERY_HISTORY.insert(0, {"sql": sql, "ok": True})
        if len(QUERY_HISTORY) > 50:
            QUERY_HISTORY.pop()
        return entry
    except Exception as e:
        QUERY_HISTORY.insert(0, {"sql": sql, "ok": False})
        return {"sql": sql, "ok": False, "error": str(e)}


def get_schema() -> list[dict]:
    if DB_CONN is None:
        return []
    cursor = DB_CONN.cursor()
    result = []
    for table in SHEETS:
        cursor.execute(f"PRAGMA table_info({table})")
        cols = [{"name": r[1], "type": r[2] or "TEXT"} for r in cursor.fetchall()]
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        result.append({"table": table, "columns": cols, "rows": cursor.fetchone()[0]})
    return result


# ─────────────────────────────────────────────
#  Ollama: Local AI
# ─────────────────────────────────────────────

def _http_post(url: str, payload: dict, timeout: int = 5) -> dict:
    data = json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data,
                                 headers={"Content-Type": "application/json"},
                                 method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read())

def _http_get(url: str, timeout: int = 4) -> dict:
    with urllib.request.urlopen(url, timeout=timeout) as resp:
        return json.loads(resp.read())


def ollama_status() -> dict:
    global SELECTED_MODEL
    try:
        data = _http_get(f"{OLLAMA_URL}/api/tags", timeout=3)
        models = [m["name"] for m in data.get("models", [])]
        if not SELECTED_MODEL and models:
            preferred = ["deepseek-coder", "codellama", "llama3", "llama3.1",
                         "mistral", "phi3", "qwen2", "gemma2"]
            SELECTED_MODEL = next(
                (m for p in preferred for m in models if p in m.lower()),
                models[0]
            )
        return {"running": True, "models": models, "selected": SELECTED_MODEL}
    except Exception as e:
        return {"running": False, "models": [], "selected": "", "error": str(e)}


def build_schema_prompt() -> str:
    schema = get_schema()
    if not schema:
        return "No tables loaded."
    lines = ["SQLite database schema:\n"]
    for t in schema:
        col_defs = ", ".join(f"{c['name']} {c['type']}" for c in t["columns"])
        lines.append(f"CREATE TABLE {t['table']} ({col_defs});  -- {t['rows']:,} rows")
    return "\n".join(lines)


def nl_to_sql(question: str) -> dict:
    global NL_HISTORY, SELECTED_MODEL

    status = ollama_status()
    if not status["running"]:
        return {"ok": False, "error":
            "Ollama is not running.\n"
            "Start it with: ollama serve\n"
            "Install from: https://ollama.com"}
    if not status["models"]:
        return {"ok": False, "error":
            "No local models found. Pull one with:\n"
            "  ollama pull llama3\n"
            "  ollama pull deepseek-coder"}

    model = SELECTED_MODEL or status["models"][0]
    schema_text = build_schema_prompt()

    prompt = (
        "You are an expert SQLite query writer.\n"
        "Convert the user question into a valid SQLite SELECT query.\n\n"
        f"{schema_text}\n\n"
        "Rules:\n"
        "- Output ONLY the raw SQL query. No explanation. No markdown. No backticks.\n"
        "- Use only tables and columns from the schema above.\n"
        "- Default to LIMIT 500 unless user asks for all or a specific count.\n"
        "- Use proper SQLite syntax.\n"
        "- If it cannot be answered with the schema, output: ERROR: <reason>\n\n"
        f"User question: {question}\n\n"
        "SQL:"
    )

    try:
        result = _http_post(
            f"{OLLAMA_URL}/api/generate",
            {"model": model, "prompt": prompt, "stream": False,
             "options": {"temperature": 0.1, "num_predict": 512}},
            timeout=120,
        )
        sql = result.get("response", "").strip()

        # Strip accidental markdown fences
        sql = re.sub(r"^```sql\s*", "", sql, flags=re.IGNORECASE)
        sql = re.sub(r"^```\s*", "", sql)
        sql = re.sub(r"\s*```$", "", sql).strip()

        # Take only the first statement if model outputs extra text
        if "\n\n" in sql:
            sql = sql.split("\n\n")[0].strip()

        if sql.upper().startswith("ERROR:"):
            NL_HISTORY.insert(0, {"question": question, "ok": False, "sql": ""})
            return {"ok": False, "error": sql[6:].strip()}

        NL_HISTORY.insert(0, {"question": question, "ok": True, "sql": sql})
        if len(NL_HISTORY) > 20:
            NL_HISTORY.pop()
        return {"ok": True, "sql": sql, "model": model}

    except urllib.error.URLError:
        return {"ok": False, "error": "Cannot reach Ollama. Run: ollama serve"}
    except Exception as e:
        return {"ok": False, "error": f"Ollama error: {e}"}


# ─────────────────────────────────────────────
#  Flask Routes
# ─────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/api/upload", methods=["POST"])
def api_upload():
    try:
        if "file" not in request.files:
            return jsonify({"ok": False, "error": "No file in request."})
        f = request.files["file"]
        if not f.filename:
            return jsonify({"ok": False, "error": "No filename."})

        # Cross-platform safe temp path (works on Windows, Mac, Linux)
        safe_name = re.sub(r"[^\w.\-]", "_", f.filename)
        tmp_path = os.path.join(tempfile.gettempdir(), safe_name)
        f.save(tmp_path)

        result = load_file(tmp_path)

        # Clean up temp file
        try:
            os.remove(tmp_path)
        except OSError:
            pass

        return jsonify(result)

    except Exception as e:
        return jsonify({"ok": False, "error": f"Upload error: {e}"}), 500


@app.route("/api/query", methods=["POST"])
def api_query():
    try:
        data = request.get_json(force=True, silent=True) or {}
        return jsonify(run_query(data.get("sql", ""), int(data.get("limit", 500))))
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/schema")
def api_schema():
    return jsonify(get_schema())


@app.route("/api/history")
def api_history():
    return jsonify(QUERY_HISTORY[:20])


@app.route("/api/nl_history")
def api_nl_history():
    return jsonify(NL_HISTORY[:10])


@app.route("/api/nl2sql", methods=["POST"])
def api_nl2sql():
    try:
        data = request.get_json(force=True, silent=True) or {}
        question = data.get("question", "").strip()
        if not question:
            return jsonify({"ok": False, "error": "Empty question."})
        return jsonify(nl_to_sql(question))
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/ollama_status")
def api_ollama_status():
    return jsonify(ollama_status())


@app.route("/api/set_model", methods=["POST"])
def api_set_model():
    global SELECTED_MODEL
    data = request.get_json(force=True, silent=True) or {}
    SELECTED_MODEL = data.get("model", "").strip()
    return jsonify({"ok": True, "model": SELECTED_MODEL})


@app.route("/api/set_ollama_url", methods=["POST"])
def api_set_ollama_url():
    global OLLAMA_URL
    data = request.get_json(force=True, silent=True) or {}
    OLLAMA_URL = data.get("url", "http://localhost:11434").rstrip("/")
    return jsonify({"ok": True})


@app.route("/api/export", methods=["POST"])
def api_export():
    try:
        data = request.get_json(force=True, silent=True) or {}
        result = run_query(data.get("sql", ""), limit=100_000)
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
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ─────────────────────────────────────────────
#  HTML Template
# ─────────────────────────────────────────────

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>Excel SQL Explorer</title>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#0d1117;--surface:#161b22;--surface2:#1c2230;--border:#2a3444;
  --accent:#f0b429;--accent2:#e05c3a;--green:#3ddc97;
  --teal:#22d3ee;--teal-dim:rgba(34,211,238,.1);--teal-border:rgba(34,211,238,.28);
  --text:#e6edf3;--muted:#7d8fa1;--danger:#e05c3a;
  --radius:10px;--mono:'JetBrains Mono',monospace;--sans:'Sora',sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px}
.app{display:flex;flex-direction:column;height:100vh;overflow:hidden}
header{display:flex;align-items:center;gap:10px;padding:0 16px;height:52px;
  background:var(--surface);border-bottom:1px solid var(--border);flex-shrink:0}
.logo{display:flex;align-items:center;gap:8px}
.logo-icon{width:30px;height:30px;border-radius:7px;
  background:linear-gradient(135deg,var(--accent),var(--accent2));
  display:flex;align-items:center;justify-content:center;font-size:16px}
.logo-text{font-size:15px;font-weight:700;letter-spacing:-.3px}
.logo-text span{color:var(--accent)}
.local-badge{font-size:10px;font-weight:700;background:var(--teal-dim);
  border:1px solid var(--teal-border);color:var(--teal);
  padding:2px 8px;border-radius:20px;letter-spacing:.5px}
.spacer{flex:1}
.file-badge{font-family:var(--mono);font-size:11px;color:var(--green);
  background:rgba(61,220,151,.1);border:1px solid rgba(61,220,151,.25);
  padding:3px 10px;border-radius:20px;display:none;max-width:180px;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.btn{display:inline-flex;align-items:center;gap:5px;padding:6px 13px;border-radius:8px;
  font-family:var(--sans);font-size:12px;font-weight:600;cursor:pointer;border:none;transition:all .15s}
.btn-primary{background:var(--accent);color:#0d1117}
.btn-primary:hover{background:#f7c94e}
.btn-ai{background:var(--teal-dim);color:var(--teal);border:1px solid var(--teal-border)}
.btn-ai:hover{background:rgba(34,211,238,.18)}
.btn-ghost{background:transparent;color:var(--muted);border:1px solid var(--border)}
.btn-ghost:hover{color:var(--text);border-color:var(--muted)}
.btn:disabled{opacity:.4;cursor:not-allowed}
.main{display:flex;flex:1;overflow:hidden}
.sidebar{width:230px;background:var(--surface);border-right:1px solid var(--border);
  display:flex;flex-direction:column;overflow:hidden;flex-shrink:0}
.sidebar-head{padding:11px 14px 8px;font-size:10px;font-weight:700;
  text-transform:uppercase;letter-spacing:1px;color:var(--muted);border-bottom:1px solid var(--border)}
.schema-area{flex:1;overflow-y:auto;padding:5px 0}
.table-entry{padding:4px 13px;cursor:pointer;user-select:none}
.table-entry:hover .tname{color:var(--accent)}
.tname{font-family:var(--mono);font-size:11px;font-weight:600;color:var(--text);
  display:flex;align-items:center;gap:5px}
.tname::before{content:'▶';font-size:8px;color:var(--muted)}
.tname.open::before{content:'▼'}
.trows{font-size:10px;color:var(--muted);margin:1px 0 2px 14px}
.col-list{display:none;padding:2px 0 3px 14px}
.col-list.visible{display:block}
.col-item{font-family:var(--mono);font-size:10px;color:var(--muted);
  padding:1px 0;display:flex;gap:6px;align-items:center}
.col-item span{color:var(--green);font-size:9px}
.sidebar-section{border-top:1px solid var(--border);padding:5px 0}
.s-head{padding:5px 14px 2px;font-size:10px;font-weight:700;
  text-transform:uppercase;letter-spacing:1px;color:var(--muted)}
.hist-item{padding:3px 13px;font-family:var(--mono);font-size:10px;color:var(--muted);
  cursor:pointer;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hist-item:hover{color:var(--accent)}
.hist-item.err{color:var(--danger)}
.hist-item.ai{color:var(--teal)}
.center{flex:1;display:flex;flex-direction:column;overflow:hidden;min-width:0}
.upload-zone{flex:1;display:flex;flex-direction:column;align-items:center;
  justify-content:center;gap:14px;padding:32px}
.upload-box{width:100%;max-width:440px;border:2px dashed var(--border);border-radius:16px;
  padding:36px 28px;text-align:center;cursor:pointer;transition:all .2s;background:var(--surface)}
.upload-box:hover,.upload-box.drag{border-color:var(--accent);background:rgba(240,180,41,.04)}
.upload-icon{font-size:40px;margin-bottom:10px}
.upload-title{font-size:16px;font-weight:700;margin-bottom:5px}
.upload-sub{color:var(--muted);font-size:12px}
.upload-exts{margin-top:12px;display:flex;gap:6px;justify-content:center;flex-wrap:wrap}
.ext-badge{font-family:var(--mono);font-size:10px;background:var(--surface2);
  border:1px solid var(--border);border-radius:5px;padding:2px 7px;color:var(--muted)}
#fileInput{display:none}
.editor-panel{display:none;flex-direction:column;flex:1;overflow:hidden}
.ai-panel{padding:11px 15px;border-bottom:1px solid var(--teal-border);flex-shrink:0;
  background:linear-gradient(135deg,rgba(34,211,238,.05),rgba(34,211,238,.02))}
.ai-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;
  color:var(--teal);margin-bottom:7px;display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.model-pill{background:var(--teal-dim);border:1px solid var(--teal-border);color:var(--teal);
  font-family:var(--mono);font-size:10px;padding:2px 8px;border-radius:12px;font-weight:400;
  cursor:pointer}
.model-pill:hover{background:rgba(34,211,238,.2)}
.ai-input-row{display:flex;gap:8px;align-items:flex-start}
textarea#nlInput{flex:1;background:rgba(34,211,238,.05);color:var(--text);
  border:1px solid var(--teal-border);border-radius:var(--radius);
  font-family:var(--sans);font-size:13px;padding:9px 12px;
  resize:none;height:40px;min-height:40px;max-height:110px;
  outline:none;transition:all .2s;line-height:1.5}
textarea#nlInput:focus{border-color:var(--teal);height:76px}
textarea#nlInput::placeholder{color:rgba(34,211,238,.35)}
textarea#nlInput:disabled{opacity:.4}
.ai-spinner{width:14px;height:14px;border:2px solid var(--teal-border);
  border-top-color:var(--teal);border-radius:50%;
  animation:spin .6s linear infinite;display:none;margin-top:13px;flex-shrink:0}
.sql-preview{margin-top:8px;background:rgba(34,211,238,.05);
  border:1px solid var(--teal-border);border-radius:8px;
  padding:10px 12px;font-family:var(--mono);font-size:12px;color:var(--teal);
  display:none;white-space:pre-wrap;word-break:break-all}
.preview-model{font-size:10px;color:rgba(34,211,238,.5);margin-bottom:4px;font-family:var(--sans)}
.preview-actions{display:flex;gap:6px;margin-top:7px}
.ai-status{margin-top:6px;font-size:10px;display:flex;align-items:center;gap:5px}
.status-led{width:6px;height:6px;border-radius:50%;background:var(--muted);flex-shrink:0}
.status-led.on{background:var(--green)}
.status-led.off{background:var(--danger)}
.status-led.pulse{background:var(--teal);animation:blink 1.2s infinite}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
.sql-area{padding:9px 15px;border-bottom:1px solid var(--border);
  background:var(--surface);flex-shrink:0}
.sql-label{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;
  letter-spacing:1px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center}
textarea#sqlInput{width:100%;background:var(--bg);color:var(--text);
  border:1px solid var(--border);border-radius:var(--radius);
  font-family:var(--mono);font-size:12px;padding:9px 12px;
  resize:vertical;min-height:84px;max-height:200px;outline:none;transition:border .2s}
textarea#sqlInput:focus{border-color:var(--accent)}
.run-bar{display:flex;gap:7px;margin-top:8px;align-items:center;flex-wrap:wrap}
.limit-select{font-family:var(--mono);font-size:11px;background:var(--bg);
  color:var(--muted);border:1px solid var(--border);border-radius:6px;padding:5px 8px;cursor:pointer}
.spinner{width:15px;height:15px;border:2px solid var(--border);
  border-top-color:var(--accent);border-radius:50%;
  animation:spin .6s linear infinite;display:none}
@keyframes spin{to{transform:rotate(360deg)}}
.results-area{flex:1;display:flex;flex-direction:column;overflow:hidden}
.results-header{padding:6px 14px;background:var(--surface);border-bottom:1px solid var(--border);
  display:flex;align-items:center;gap:8px;font-size:11px;color:var(--muted);flex-wrap:wrap;flex-shrink:0}
.stat-val{color:var(--accent);font-family:var(--mono);font-weight:700}
.result-spacer{flex:1}
.table-wrap{flex:1;overflow:auto}
table{width:max-content;min-width:100%;border-collapse:collapse}
thead tr{position:sticky;top:0;z-index:2}
thead th{background:var(--surface2);padding:7px 12px;text-align:left;
  font-family:var(--mono);font-size:10px;font-weight:700;color:var(--muted);
  border-bottom:2px solid var(--border);white-space:nowrap}
tbody tr{border-bottom:1px solid rgba(42,52,68,.5)}
tbody tr:hover{background:rgba(240,180,41,.04)}
tbody td{padding:5px 12px;font-family:var(--mono);font-size:11px;color:var(--text);
  white-space:nowrap;max-width:280px;overflow:hidden;text-overflow:ellipsis}
td.null{color:var(--muted);font-style:italic}
td.num{color:#82aaff}
.error-box{margin:16px;padding:13px 17px;background:rgba(224,92,58,.1);
  border:1px solid rgba(224,92,58,.3);border-radius:var(--radius);
  color:var(--danger);font-family:var(--mono);font-size:12px;white-space:pre-wrap}
.empty-state{flex:1;display:flex;align-items:center;justify-content:center;
  color:var(--muted);flex-direction:column;gap:8px;padding:40px;text-align:center}
.empty-state .big{font-size:34px}
.ok-dot{width:7px;height:7px;border-radius:50%;background:var(--green)}
.truncate-notice{text-align:center;padding:8px;font-size:11px;color:var(--accent);
  background:rgba(240,180,41,.05);border-top:1px solid var(--border);font-family:var(--mono)}
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:100;
  align-items:center;justify-content:center}
.modal-bg.open{display:flex}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:14px;
  padding:22px;width:92%;max-width:460px;box-shadow:0 24px 60px rgba(0,0,0,.5)}
.modal h3{font-size:16px;font-weight:700;margin-bottom:3px}
.modal .sub{color:var(--muted);font-size:12px;margin-bottom:18px}
.info-box{background:var(--teal-dim);border:1px solid var(--teal-border);border-radius:8px;
  padding:10px 13px;font-size:11px;color:rgba(34,211,238,.85);margin-bottom:16px;line-height:1.8}
.info-box code{font-family:var(--mono);background:rgba(34,211,238,.15);
  border-radius:4px;padding:1px 5px;font-size:10px}
.form-group{margin-bottom:13px}
.form-label{font-size:10px;font-weight:700;color:var(--muted);
  text-transform:uppercase;letter-spacing:.8px;margin-bottom:5px;display:block}
.form-input{width:100%;background:var(--bg);color:var(--text);border:1px solid var(--border);
  border-radius:8px;font-family:var(--mono);font-size:12px;padding:8px 11px;outline:none}
.form-input:focus{border-color:var(--teal)}
select.form-input option{background:var(--surface)}
.modal-status{margin-top:4px;font-size:11px;font-family:var(--mono)}
.modal-status.ok{color:var(--green)}
.modal-status.err{color:var(--danger)}
.modal-actions{display:flex;gap:8px;justify-content:flex-end;margin-top:18px}
.divider{border:none;border-top:1px solid var(--border);margin:16px 0}
.sidebar-toggle{display:none;background:none;border:none;color:var(--muted);
  font-size:20px;cursor:pointer;padding:3px 6px}
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:10}
@media(max-width:680px){
  .sidebar{position:fixed;left:0;top:52px;bottom:0;z-index:20;
    transform:translateX(-100%);transition:transform .25s;width:240px}
  .sidebar.open{transform:translateX(0)}
  .overlay.open{display:block}
  .sidebar-toggle{display:block}
  .ai-panel,.sql-area{padding:9px 11px}
  textarea#sqlInput{min-height:64px;font-size:11px}
  textarea#nlInput{font-size:12px}
  .logo-text{font-size:13px}
  .file-badge{max-width:110px}
}
</style>
</head>
<body>
<div class="app">
  <header>
    <button class="sidebar-toggle" id="sidebarToggle">☰</button>
    <div class="logo">
      <div class="logo-icon">⚡</div>
      <div class="logo-text">Excel<span>SQL</span></div>
    </div>
    <span class="local-badge">LOCAL AI</span>
    <div class="spacer"></div>
    <div class="file-badge" id="fileBadge"></div>
    <button class="btn btn-ghost" id="settingsBtn" style="padding:5px 11px">⚙ Settings</button>
    <label for="fileInput" class="btn btn-ghost" style="padding:5px 11px;cursor:pointer">📂 Load File</label>
    <input type="file" id="fileInput" accept=".xlsx,.xls,.xlsm,.xlsb,.ods,.csv">
  </header>

  <div class="main">
    <aside class="sidebar" id="sidebar">
      <div class="sidebar-head">Schema</div>
      <div class="schema-area" id="schemaArea">
        <div style="padding:16px 13px;color:var(--muted);font-size:11px">Upload a file to see schema</div>
      </div>
      <div class="sidebar-section">
        <div class="s-head">SQL History</div>
        <div id="historyList"></div>
      </div>
      <div class="sidebar-section">
        <div class="s-head" style="color:var(--teal)">⬡ AI History</div>
        <div id="aiHistoryList"></div>
      </div>
    </aside>
    <div class="overlay" id="overlay"></div>

    <div class="center">
      <div class="upload-zone" id="uploadZone">
        <div class="upload-box" id="dropZone">
          <div class="upload-icon">📊</div>
          <div class="upload-title">Drop your Excel or CSV here</div>
          <div class="upload-sub">or click <strong>Load File</strong> in the header</div>
          <div class="upload-exts">
            <span class="ext-badge">.xlsx</span><span class="ext-badge">.xls</span>
            <span class="ext-badge">.xlsm</span><span class="ext-badge">.ods</span>
            <span class="ext-badge">.csv</span>
          </div>
        </div>
      </div>

      <div class="editor-panel" id="editorPanel">

        <!-- AI Panel -->
        <div class="ai-panel">
          <div class="ai-label">
            ⬡ Ask your data in plain English
            <span class="model-pill" id="modelPill" onclick="openSettings()" title="Click to change model">—</span>
          </div>
          <div class="ai-input-row">
            <textarea id="nlInput" placeholder="e.g.  Show top 10 products by total sales…" rows="1"></textarea>
            <button class="btn btn-ai" id="generateBtn">⬡ Generate SQL</button>
            <div class="ai-spinner" id="aiSpinner"></div>
          </div>
          <div class="sql-preview" id="sqlPreview">
            <div class="preview-model" id="previewModel"></div>
            <div id="previewSQL"></div>
          </div>
          <div class="preview-actions" id="previewActions" style="display:none">
            <button class="btn btn-primary" id="useAndRunBtn" style="font-size:11px;padding:5px 12px">▶ Use &amp; Run</button>
            <button class="btn btn-ghost" id="useOnlyBtn" style="font-size:11px;padding:5px 12px">Copy to Editor</button>
          </div>
          <div class="ai-status">
            <div class="status-led" id="ollamaLed"></div>
            <span id="ollamaStatusText" style="font-size:10px;color:var(--muted)">Checking Ollama…</span>
          </div>
        </div>

        <!-- SQL Editor -->
        <div class="sql-area">
          <div class="sql-label">
            <span>SQL Query</span>
            <button class="btn btn-ghost" id="clearBtn" style="padding:3px 8px;font-size:10px">Clear</button>
          </div>
          <textarea id="sqlInput" placeholder="SELECT * FROM your_table LIMIT 10" spellcheck="false"></textarea>
          <div class="run-bar">
            <button class="btn btn-primary" id="runBtn">▶ Run</button>
            <button class="btn btn-ghost" id="exportBtn">⬇ CSV</button>
            <select class="limit-select" id="limitSelect">
              <option value="100">100 rows</option>
              <option value="500" selected>500 rows</option>
              <option value="1000">1 000 rows</option>
              <option value="5000">5 000 rows</option>
              <option value="50000">All rows</option>
            </select>
            <div class="spinner" id="spinner"></div>
            <div style="flex:1"></div>
            <span style="font-size:10px;color:var(--muted)">Ctrl+Enter to run</span>
          </div>
        </div>

        <!-- Results -->
        <div class="results-area" id="resultsArea">
          <div class="empty-state">
            <div class="big">💡</div>
            <div>Ask in plain English — or write SQL directly</div>
            <div style="font-size:11px;margin-top:4px">Ctrl+Enter runs the query</div>
          </div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- Settings Modal -->
<div class="modal-bg" id="settingsModal">
  <div class="modal">
    <h3>⚙ Settings — Local AI</h3>
    <div class="sub">Connect to your local Ollama instance for AI-powered SQL generation</div>
    <div class="info-box">
      <strong>Quick setup:</strong><br>
      1. Install: <a href="https://ollama.com" target="_blank" style="color:var(--teal)">ollama.com</a><br>
      2. Start server: <code>ollama serve</code><br>
      3. Pull a model: <code>ollama pull llama3</code> &nbsp;or&nbsp; <code>ollama pull deepseek-coder</code><br>
      Models will appear in the list below automatically ✓
    </div>
    <div class="form-group">
      <label class="form-label">Ollama URL</label>
      <input type="text" class="form-input" id="ollamaUrlInput" value="http://localhost:11434">
    </div>
    <div class="form-group">
      <label class="form-label">Active Model</label>
      <select class="form-input" id="modelSelect">
        <option value="">— click Refresh to load models —</option>
      </select>
      <div class="modal-status" id="ollamaModalStatus"></div>
    </div>
    <hr class="divider">
    <div style="font-size:11px;color:var(--muted)">
      <strong style="color:var(--text)">Best models for SQL:</strong>
      deepseek-coder &nbsp;·&nbsp; codellama &nbsp;·&nbsp; llama3.1 &nbsp;·&nbsp; mistral &nbsp;·&nbsp; phi3
    </div>
    <div class="modal-actions">
      <button class="btn btn-ghost" id="refreshModelsBtn">↻ Refresh</button>
      <button class="btn btn-ghost" id="closeSettings">Cancel</button>
      <button class="btn btn-ai" id="saveSettings">Save</button>
    </div>
  </div>
</div>

<script>
const $ = id => document.getElementById(id);
let generatedSQL = '';

// ── Upload ──
$('fileInput').addEventListener('change', e => { if (e.target.files[0]) uploadFile(e.target.files[0]); });
$('dropZone').addEventListener('click', () => $('fileInput').click());
['dragenter','dragover'].forEach(ev => $('dropZone').addEventListener(ev, e => { e.preventDefault(); $('dropZone').classList.add('drag'); }));
['dragleave','drop'].forEach(ev => $('dropZone').addEventListener(ev, e => { e.preventDefault(); $('dropZone').classList.remove('drag'); }));
$('dropZone').addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) uploadFile(f); });

async function uploadFile(file) {
  const fd = new FormData();
  fd.append('file', file);
  showSpinner(true);
  try {
    const res = await fetch('/api/upload', { method: 'POST', body: fd });
    // Guard: always parse as text first, then try JSON
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); }
    catch(e) { throw new Error('Server returned non-JSON: ' + text.slice(0, 200)); }

    if (data.ok) {
      $('fileBadge').textContent = '● ' + data.file;
      $('fileBadge').style.display = 'block';
      $('uploadZone').style.display = 'none';
      $('editorPanel').style.display = 'flex';
      loadSchema();
      if (data.tables?.length)
        $('sqlInput').value = 'SELECT * FROM ' + data.tables[0].name + ' LIMIT 20';
    } else {
      alert('Could not load file:\n' + data.error);
    }
  } catch(e) {
    alert('Upload failed:\n' + e.message);
  }
  showSpinner(false);
}

// ── Schema ──
async function loadSchema() {
  const tables = await safeGet('/api/schema');
  const el = $('schemaArea');
  if (!tables || !tables.length) {
    el.innerHTML = '<div style="padding:14px;color:var(--muted);font-size:11px">No tables</div>';
    return;
  }
  el.innerHTML = tables.map(t => `
    <div class="table-entry">
      <div class="tname" onclick="toggleCols(this,'${t.table}')">${t.table}</div>
      <div class="trows">${t.rows.toLocaleString()} rows &middot; ${t.columns.length} cols</div>
      <div class="col-list" id="cols_${t.table}">
        ${t.columns.map(c => '<div class="col-item">' + c.name + '<span>' + c.type + '</span></div>').join('')}
      </div>
    </div>`).join('');
}
function toggleCols(el, t) {
  document.getElementById('cols_' + t).classList.toggle('visible');
  el.classList.toggle('open');
}

// ── AI Generate ──
$('generateBtn').addEventListener('click', generateSQL);
$('nlInput').addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); generateSQL(); }
});

async function generateSQL() {
  const question = $('nlInput').value.trim();
  if (!question) return;
  $('generateBtn').disabled = true;
  $('aiSpinner').style.display = 'block';
  $('sqlPreview').style.display = 'none';
  $('previewActions').style.display = 'none';

  const data = await safePost('/api/nl2sql', { question });
  if (data && data.ok) {
    generatedSQL = data.sql;
    $('previewModel').textContent = '⬡ ' + (data.model || 'local model');
    $('previewSQL').textContent = data.sql;
    $('previewSQL').style.color = 'var(--teal)';
    $('sqlPreview').style.borderColor = 'var(--teal-border)';
    $('sqlPreview').style.display = 'block';
    $('previewActions').style.display = 'flex';
    loadAIHistory();
  } else if (data) {
    $('previewModel').textContent = '';
    $('previewSQL').textContent = '❌  ' + data.error;
    $('previewSQL').style.color = 'var(--danger)';
    $('sqlPreview').style.borderColor = 'rgba(224,92,58,.3)';
    $('sqlPreview').style.display = 'block';
  }
  $('generateBtn').disabled = false;
  $('aiSpinner').style.display = 'none';
}

$('useAndRunBtn').addEventListener('click', () => {
  if (!generatedSQL) return;
  $('sqlInput').value = generatedSQL;
  runQuery();
});
$('useOnlyBtn').addEventListener('click', () => {
  if (!generatedSQL) return;
  $('sqlInput').value = generatedSQL;
  $('sqlInput').focus();
});

// ── SQL ──
$('runBtn').addEventListener('click', runQuery);
$('clearBtn').addEventListener('click', () => { $('sqlInput').value = ''; $('sqlInput').focus(); });
$('sqlInput').addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); runQuery(); }
  if (e.key === 'Tab') { e.preventDefault(); insertAt($('sqlInput'), '    '); }
});

async function runQuery() {
  const sql = $('sqlInput').value.trim();
  if (!sql) return;
  showSpinner(true); $('runBtn').disabled = true;
  const data = await safePost('/api/query', { sql, limit: parseInt($('limitSelect').value) });
  if (data) renderResults(data); else renderError('Request failed.');
  loadHistory();
  showSpinner(false); $('runBtn').disabled = false;
}

function renderResults(data) {
  const el = $('resultsArea');
  if (!data.ok) { el.innerHTML = '<div class="error-box">❌ ' + esc(data.error) + '</div>'; return; }
  if (!data.columns?.length) {
    el.innerHTML = '<div class="empty-state"><div class="big">✅</div><div>' + esc(data.message || 'Done.') + '</div></div>';
    return;
  }
  let html = '<div class="results-header">'
    + '<div class="ok-dot"></div>'
    + '<span class="stat-val">' + data.count.toLocaleString() + '</span>&nbsp;rows&nbsp;'
    + '<span class="stat-val">' + data.columns.length + '</span>&nbsp;cols'
    + '<div class="result-spacer"></div>'
    + '<span style="font-size:10px;font-family:var(--mono)">' + new Date().toLocaleTimeString() + '</span>'
    + '</div>'
    + '<div class="table-wrap"><table>'
    + '<thead><tr>' + data.columns.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr></thead>'
    + '<tbody>' + data.rows.map(r => '<tr>' + r.map(v => cell(v)).join('') + '</tr>').join('') + '</tbody>'
    + '</table></div>';
  if (data.truncated)
    html += '<div class="truncate-notice">⚠ Showing first ' + data.count.toLocaleString() + ' rows — change limit or add LIMIT clause</div>';
  el.innerHTML = html;
}
function cell(v) {
  if (v === null || v === undefined || v === '') return '<td class="null">NULL</td>';
  return typeof v === 'number'
    ? '<td class="num">' + esc(String(v)) + '</td>'
    : '<td>' + esc(String(v)) + '</td>';
}
function renderError(msg) { $('resultsArea').innerHTML = '<div class="error-box">❌ ' + esc(msg) + '</div>'; }

// ── Export ──
$('exportBtn').addEventListener('click', async () => {
  const sql = $('sqlInput').value.trim(); if (!sql) return;
  const res = await fetch('/api/export', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({sql}) });
  if (!res.ok) { alert('Export failed'); return; }
  const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(await res.blob()), download: 'result.csv' });
  a.click();
});

// ── Histories ──
async function loadHistory() {
  const h = await safeGet('/api/history') || [];
  $('historyList').innerHTML = h.slice(0, 8).map(x =>
    '<div class="hist-item ' + (x.ok ? '' : 'err') + '" title="' + esc(x.sql) + '"'
    + ' onclick="$(\'sqlInput\').value=' + JSON.stringify(x.sql) + '">' + esc(trunc(x.sql, 34)) + '</div>'
  ).join('');
}
async function loadAIHistory() {
  const h = await safeGet('/api/nl_history') || [];
  $('aiHistoryList').innerHTML = h.slice(0, 6).map(x =>
    '<div class="hist-item ai" title="' + esc(x.question) + '"'
    + ' onclick="$(\'nlInput\').value=' + JSON.stringify(x.question) + '">' + esc(trunc(x.question, 34)) + '</div>'
  ).join('');
}

// ── Ollama status ──
async function checkOllama() {
  const led = $('ollamaLed'), txt = $('ollamaStatusText'), pill = $('modelPill');
  led.className = 'status-led pulse';
  const s = await safeGet('/api/ollama_status');
  if (!s) return;
  if (s.running && s.models.length) {
    led.className = 'status-led on';
    txt.textContent = 'Ollama ready · ' + s.models.length + ' model(s) available';
    txt.style.color = 'var(--green)';
    pill.textContent = s.selected || s.models[0];
    $('generateBtn').disabled = false;
    $('nlInput').disabled = false;
  } else if (s.running) {
    led.className = 'status-led off';
    txt.textContent = 'Ollama running but no models — run: ollama pull llama3';
    txt.style.color = 'var(--danger)';
    pill.textContent = 'no model';
    $('generateBtn').disabled = true;
  } else {
    led.className = 'status-led off';
    txt.textContent = 'Ollama not found — install from ollama.com then: ollama serve';
    txt.style.color = 'var(--danger)';
    pill.textContent = 'offline';
    $('generateBtn').disabled = true;
    $('nlInput').disabled = true;
  }
}
checkOllama();
setInterval(checkOllama, 15000);

// ── Settings ──
function openSettings() { $('settingsModal').classList.add('open'); refreshModelsUI(); }
$('settingsBtn').addEventListener('click', openSettings);
$('closeSettings').addEventListener('click', () => $('settingsModal').classList.remove('open'));
$('settingsModal').addEventListener('click', e => { if (e.target === $('settingsModal')) $('settingsModal').classList.remove('open'); });
$('refreshModelsBtn').addEventListener('click', refreshModelsUI);

async function refreshModelsUI() {
  const st = $('ollamaModalStatus');
  st.textContent = 'Checking…'; st.className = 'modal-status';
  const s = await safeGet('/api/ollama_status');
  if (!s) { st.textContent = 'Error'; st.className = 'modal-status err'; return; }
  const sel = $('modelSelect');
  if (s.models.length) {
    sel.innerHTML = s.models.map(m => '<option value="' + esc(m) + '"' + (m === s.selected ? ' selected' : '') + '>' + m + '</option>').join('');
    st.textContent = '✓ ' + s.models.length + ' model(s) found';
    st.className = 'modal-status ok';
  } else {
    sel.innerHTML = '<option value="">No models found</option>';
    st.textContent = s.running ? '⚠ Ollama running but no models pulled' : '✗ Ollama not running — start with: ollama serve';
    st.className = 'modal-status err';
  }
}

$('saveSettings').addEventListener('click', async () => {
  const url = $('ollamaUrlInput').value.trim();
  const model = $('modelSelect').value;
  if (url) await safePost('/api/set_ollama_url', { url });
  if (model) await safePost('/api/set_model', { model });
  $('settingsModal').classList.remove('open');
  checkOllama();
});

// ── Mobile sidebar ──
$('sidebarToggle').addEventListener('click', () => { $('sidebar').classList.toggle('open'); $('overlay').classList.toggle('open'); });
$('overlay').addEventListener('click', () => { $('sidebar').classList.remove('open'); $('overlay').classList.remove('open'); });

// ── Safe HTTP helpers (always parse text first, never crash on HTML responses) ──
async function safeGet(url) {
  try {
    const res = await fetch(url);
    const text = await res.text();
    return JSON.parse(text);
  } catch(e) { console.error('GET', url, e); return null; }
}
async function safePost(url, body) {
  try {
    const res = await fetch(url, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body) });
    const text = await res.text();
    return JSON.parse(text);
  } catch(e) { console.error('POST', url, e); return null; }
}

function showSpinner(on) { $('spinner').style.display = on ? 'block' : 'none'; }
function insertAt(el, t) {
  const s = el.selectionStart, e = el.selectionEnd;
  el.value = el.value.slice(0, s) + t + el.value.slice(e);
  el.setSelectionRange(s + t.length, s + t.length); el.focus();
}
function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function trunc(s, n) { return s.length > n ? s.slice(0, n) + '…' : s; }
</script>
</body>
</html>
"""


# ─────────────────────────────────────────────
#  CLI Mode
# ─────────────────────────────────────────────

def cli_mode():
    print("\n╔══════════════════════════════════════════╗")
    print("║   Excel SQL Explorer  v3.1  (CLI)        ║")
    print("║   Local AI powered by Ollama             ║")
    print("╚══════════════════════════════════════════╝\n")

    status = ollama_status()
    if status["running"]:
        print(f"✦  Ollama: running  |  models: {', '.join(status['models'][:5])}")
        print(f"   Active model: {status.get('selected','—')}")
    else:
        print("⚠  Ollama not running. AI (.ask) is disabled.")
        print("   Install: https://ollama.com  →  ollama serve\n")

    filepath = input("\n📂 Enter path to Excel/CSV file: ").strip().strip('"').strip("'")
    if not os.path.exists(filepath):
        print(f"❌  File not found: {filepath}")
        sys.exit(1)

    print("⏳ Loading …")
    result = load_file(filepath)
    if not result["ok"]:
        print(f"❌  {result['error']}")
        sys.exit(1)

    print(f"\n✅  Loaded: {result['file']}")
    for t in result["tables"]:
        print(f"   📋 {t['name']:<22} {t['rows']:>8,} rows  ·  {len(t['columns'])} cols")
        print(f"      {', '.join(t['columns'][:7])}{'…' if len(t['columns'])>7 else ''}")

    print("\n💡 Commands:  .ask <question>  .schema  .history  .models  .quit\n")

    while True:
        try:
            user_input = input("sql> ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n👋  Bye!")
            break

        if not user_input:
            continue
        if user_input.lower() in (".quit", ".exit", "quit", "exit"):
            print("👋  Bye!")
            break
        if user_input.lower() == ".schema":
            for t in get_schema():
                print(f"\n  Table: {t['table']}  ({t['rows']:,} rows)")
                for c in t["columns"]:
                    print(f"    {c['name']:<28} {c['type']}")
            continue
        if user_input.lower() == ".history":
            for i, h in enumerate(QUERY_HISTORY[:10], 1):
                print(f"  {'✓' if h['ok'] else '✗'} {i:2}. {h['sql'][:78]}")
            continue
        if user_input.lower() == ".models":
            s = ollama_status()
            if s["running"]:
                print(f"  Ollama running. Models available:")
                for m in s["models"]:
                    print(f"    {'▶' if m==SELECTED_MODEL else '·'} {m}")
            else:
                print("  ✗ Ollama not running")
            continue
        if user_input.lower().startswith(".ask "):
            question = user_input[5:].strip()
            if not question:
                continue
            print("  ⬡ Generating SQL …")
            r = nl_to_sql(question)
            if r["ok"]:
                print(f"\n  Model : {r.get('model','?')}")
                print(f"  SQL   :\n{r['sql']}\n")
                ans = input("  Run this? [Y/n]: ").strip().lower()
                if ans in ("", "y", "yes"):
                    user_input = r["sql"]
                else:
                    continue
            else:
                print(f"  ❌  {r['error']}")
                continue

        result = run_query(user_input)
        if not result["ok"]:
            print(f"❌  {result['error']}")
        elif not result["columns"]:
            print(f"✅  {result.get('message', 'OK')}")
        else:
            cols, rows = result["columns"], result["rows"]
            widths = [len(c) for c in cols]
            for row in rows:
                for i, v in enumerate(row):
                    widths[i] = min(max(widths[i], len(str(v if v is not None else "NULL"))), 32)
            sep = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
            print(sep)
            print("| " + " | ".join(c[:w].ljust(w) for c, w in zip(cols, widths)) + " |")
            print(sep)
            for row in rows:
                cells = [str(v if v is not None else "NULL")[:w].ljust(w) for v, w in zip(row, widths)]
                print("| " + " | ".join(cells) + " |")
            print(sep)
            suffix = "  [TRUNCATED — add LIMIT]" if result.get("truncated") else ""
            print(f"  {result['count']:,} row(s){suffix}")


# ─────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Excel SQL Explorer v3.1 — Local AI")
    parser.add_argument("--cli", action="store_true", help="Run in CLI mode")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=5050)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--ollama-url", default="http://localhost:11434")
    args = parser.parse_args()

    global OLLAMA_URL
    OLLAMA_URL = args.ollama_url.rstrip("/")

    if args.cli:
        cli_mode()
    else:
        s = ollama_status()
        print("\n╔══════════════════════════════════════════╗")
        print("║   Excel SQL Explorer  v3.1  ·  Web Mode  ║")
        print("╚══════════════════════════════════════════╝")
        print(f"\n  Local:   http://localhost:{args.port}")
        print(f"  Network: http://<your-ip>:{args.port}")
        if s["running"]:
            print(f"  Ollama:  ✦ {s.get('selected','?')} ready")
        else:
            print("  Ollama:  ✗ Not detected  →  ollama.com  →  ollama serve")
        print("\n  Press Ctrl+C to stop\n")
        if not args.no_browser:
            threading.Timer(0.9, lambda: webbrowser.open(f"http://localhost:{args.port}")).start()
        app.run(host=args.host, port=args.port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
