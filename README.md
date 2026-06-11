# Excel Natural Language to SQL Explorer For Excel Locally⚡

Query Excel and CSV files using SQL — works fully **offline**, no internet required.

## Features

- Upload `.xlsx`, `.xls`, `.xlsm`, `.ods`, `.csv` files
- Each sheet automatically becomes a SQL table
- Full SQL support via SQLite (SELECT, WHERE, GROUP BY, ORDER BY, JOINs, aggregates…)
- Export results to CSV
- Query history in the sidebar
- Works on **PC browser** and **mobile browser** (same Wi-Fi)
- Also works in **terminal/CLI mode**

---

## Setup (one time)

```bash
pip install -r requirements.txt
```

---

## Usage

### 🌐 Web Mode (recommended — works on PC + mobile)

```bash
python excel_sql.py
```

Then open in browser:
- **PC:**     `http://localhost:5050`
- **Mobile:** `http://<your-computer-ip>:5050`  ← find your IP via `ipconfig` (Windows) or `ifconfig` (Mac/Linux)

### 💻 CLI / Terminal Mode

```bash
python excel_sql.py --cli
```

You'll be prompted for a file path, then can type SQL queries directly.

### Options

| Flag | Description |
|------|-------------|
| `--cli` | Run in terminal mode |
| `--port 8080` | Change web port (default: 5050) |
| `--host 127.0.0.1` | Restrict to localhost only |
| `--no-browser` | Don't auto-open browser |

---

## SQL Tips

```sql
-- View all data
SELECT * FROM Sheet1 LIMIT 20

-- Filter and sort
SELECT Name, Sales FROM Sheet1
WHERE Sales > 10000
ORDER BY Sales DESC

-- Aggregate
SELECT Region, SUM(Revenue) as Total
FROM Sheet1
GROUP BY Region

-- Multiple sheets (if Excel has many)
SELECT a.*, b.Category
FROM Sheet1 a
JOIN Sheet2 b ON a.id = b.id
```

### CLI Commands

| Command | Action |
|---------|--------|
| `.schema` | Show all tables and columns |
| `.history` | Show recent queries |
| `.quit` | Exit |
| `Ctrl+C` | Exit |

---

## Notes

- All data is loaded into memory — nothing is written to disk
- Files up to ~100 MB work well; very large files may be slow
- Column and sheet names with spaces/special chars are auto-sanitized (spaces → `_`)
