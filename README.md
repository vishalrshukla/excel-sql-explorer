# Natural Language to SQL Explorer for Excel Locally⚡

Query Excel and CSV files using SQL — works fully **offline**, no internet required. Perfect for analyzing spreadsheets without cloud uploads or external dependencies.

Convert natural language queries to SQL using **local Ollama models** — no API keys, no cloud dependence.

## Table of Contents

- [Features](#features)
- [Quick Start](#quick-start)
- [Requirements](#requirements)
- [Setup (one time)](#setup-one-time)
- [Ollama Setup](#ollama-setup)
- [Usage](#usage)
- [SQL Examples](#sql-examples)
- [CLI Commands](#cli-commands)
- [Limitations & Notes](#limitations--notes)
- [Troubleshooting](#troubleshooting)

---

## Features

- ✅ Upload `.xlsx`, `.xls`, `.xlsm`, `.ods`, `.csv` files
- ✅ Each sheet automatically becomes a SQL table
- ✅ Full SQL support via SQLite (SELECT, WHERE, GROUP BY, ORDER BY, JOINs, aggregates…)
- ✅ **Convert natural language to SQL** using local Ollama models (no internet/API keys needed)
- ✅ Export results to CSV
- ✅ Query history in the sidebar
- ✅ Works on **PC browser** and **mobile browser** (same Wi-Fi)
- ✅ **Terminal/CLI mode** for headless environments
- ✅ **100% offline** — data never leaves your machine

---

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Start Ollama service (see Ollama Setup section below)
ollama serve

# 3. In another terminal, start the web server
python excel_sql.py

# 4. Open in browser
# PC:     http://localhost:5050
# Mobile: http://<your-computer-ip>:5050
```

Done! Upload a file and start querying.

---

## Requirements

- **Python 3.8+**
- **Ollama** (for local NLP-to-SQL translation)
- Dependencies: Flask, openpyxl, pandas, sqlite3, requests (included in requirements.txt)
- Modern web browser (Chrome, Firefox, Safari, Edge)

---

## Setup (one time)

```bash
pip install -r requirements.txt
```

---

## Ollama Setup

Ollama enables local, offline AI models for converting natural language to SQL queries. No internet connection or API keys required.

### Install Ollama

**Windows:**
1. Download from [ollama.ai](https://ollama.ai)
2. Run the installer
3. Ollama will start automatically

**Mac:**
```bash
brew install ollama
```

**Linux:**
```bash
curl https://ollama.ai/install.sh | sh
```

### Start Ollama Service

Ollama must be running in the background for the application to work.

```bash
ollama serve
```

This starts the Ollama service on `http://localhost:11434` (default).

### Pull/Install Models

Download a model suitable for SQL generation. **Recommended models:**

**Option 1: Mistral (Recommended - Fast & Accurate)**
```bash
ollama pull mistral
```

**Option 2: Neural Chat (Good balance)**
```bash
ollama pull neural-chat
```

**Option 3: Llama 2 (Larger, slower)**
```bash
ollama pull llama2
```

**Option 4: CodeLlama (Best for SQL - if you have GPU/enough RAM)**
```bash
ollama pull codellama
```

### Verify Installation

```bash
# Check if Ollama is running
curl http://localhost:11434/api/tags

# Test a model
ollama run mistral "Convert 'Show me all sales over 1000' to SQL"
```

### System Requirements

- **Minimum RAM:** 8 GB (recommended 16 GB for better performance)
- **Storage:** 4-7 GB per model (Mistral: 4GB, Llama2: 3.8GB, CodeLlama: 3.8GB)
- **GPU (optional):** NVIDIA CUDA or Apple Metal for faster inference (auto-detected by Ollama)

### Troubleshooting Ollama

**"Connection refused" error:**
- Make sure Ollama is running: `ollama serve`
- Check it's listening on `localhost:11434`: `curl http://localhost:11434/api/tags`

**Model won't load / Out of memory:**
- Reduce context size or use a smaller model (Mistral instead of Llama2)
- Close other applications to free up RAM

**Slow responses:**
- Consider using a GPU (NVIDIA or Apple Silicon)
- Reduce model size or use Mistral instead of larger models

---

## Usage

### 🌐 Web Mode (recommended — works on PC + mobile)

```bash
python excel_sql.py
```

Then open in browser:
- **PC:**     `http://localhost:5050`
- **Mobile:** `http://<your-computer-ip>:5050`

**Find your computer's IP:**
- **Windows:** Run `ipconfig` in Command Prompt, look for "IPv4 Address"
- **Mac/Linux:** Run `ifconfig` in Terminal, look for "inet" address (usually `192.168.x.x`)

### Using Natural Language to SQL

1. Upload your Excel or CSV file
2. Type a natural language query in the query box (e.g., "Show me all sales over 1000 sorted by date")
3. The system uses your local Ollama model to convert it to SQL
4. Results appear instantly

### 💻 CLI / Terminal Mode

```bash
python excel_sql.py --cli
```

You'll be prompted for a file path, then can type SQL queries or natural language questions directly.

### Options

| Flag | Description |
|------|-------------|
| `--cli` | Run in terminal mode |
| `--port 8080` | Change web port (default: 5050) |
| `--host 127.0.0.1` | Restrict to localhost only (disable mobile access) |
| `--no-browser` | Don't auto-open browser |
| `--ollama-url http://localhost:11434` | Custom Ollama service URL |

**Example:**
```bash
python excel_sql.py --port 3000 --no-browser --ollama-url http://192.168.1.100:11434
```

---

## SQL Examples

### View all data
```sql
SELECT * FROM Sheet1 LIMIT 20
```

### Filter and sort
```sql
SELECT Name, Sales FROM Sheet1
WHERE Sales > 10000
ORDER BY Sales DESC
```

### Aggregate (GROUP BY)
```sql
SELECT Region, SUM(Revenue) as Total, COUNT(*) as Count
FROM Sheet1
GROUP BY Region
ORDER BY Total DESC
```

### JOIN multiple sheets
```sql
SELECT a.*, b.Category
FROM Sheet1 a
JOIN Sheet2 b ON a.id = b.id
WHERE a.Status = 'Active'
```

### Complex aggregation
```sql
SELECT 
  Region,
  AVG(Sales) as AvgSales,
  MAX(Sales) as MaxSales,
  MIN(Sales) as MinSales
FROM Sheet1
WHERE Year = 2024
GROUP BY Region
```

---

## CLI Commands

| Command | Action |
|---------|--------|
| `.schema` | Show all tables and columns |
| `.history` | Show recent queries |
| `.quit` | Exit |
| `Ctrl+C` | Exit |

---

## Limitations & Notes

### Data Loading
- All data is loaded into memory — nothing is written to disk
- **Recommended max file size:** ~100 MB (very large files may be slow)
- For files >100 MB, consider splitting into smaller sheets or using CLI mode

### Column & Sheet Name Handling
- Spaces and special characters are auto-sanitized:
  - `"Sales $"` → `Sales_`
  - `"Q1 2024"` → `Q1_2024`
  - `"Last/Name"` → `Last_Name`
- Use the sanitized names in your SQL queries

### Natural Language to SQL
- Accuracy depends on the Ollama model used
- Complex queries may require refinement or manual SQL editing
- Works best with clear, specific questions
- Requires Ollama service to be running

### SQLite Limitations
- No window functions (e.g., `ROW_NUMBER()`, `RANK()`)
- No recursive CTEs
- Limited string functions compared to other databases
- No native date arithmetic (use Python before querying)

### Security
- **Local-only:** All processing happens on your machine
- No data is transmitted to external servers
- No internet connection required (except for downloading Ollama models initially)
- Safe to use with sensitive/confidential spreadsheets
- AI models run locally — no queries sent to cloud services

---

## Troubleshooting

### Port 5050 already in use
**Error:** `Address already in use`

**Solution:** Use a different port:
```bash
python excel_sql.py --port 8080
```

### Can't access from mobile browser
**Problem:** Mobile shows "connection refused" or timeout

**Solutions:**
1. Verify both devices are on the **same Wi-Fi network**
2. Use the correct computer IP (not localhost) — check with `ipconfig` or `ifconfig`
3. Check firewall isn't blocking the port
4. Try `--host 0.0.0.0` explicitly:
   ```bash
   python excel_sql.py --host 0.0.0.0
   ```

### Query returns no results or errors
- Check table name with `.schema` command
- Verify column names are exact (spaces sanitized to `_`)
- Enclose quoted identifiers: `` `Column Name` `` (if needed)

### File upload fails
- Check file format is supported (`.xlsx`, `.xls`, `.xlsm`, `.ods`, `.csv`)
- Ensure file is not corrupted
- Try re-saving the file in Excel or LibreOffice
- Check file size is under 100 MB

### Large file is very slow
- Try filtering data with a `WHERE` clause first
- Split large files into smaller sheets
- Use CLI mode for better performance on massive datasets

### "Column not found" error
- Use `.schema` to view exact column names (spaces are `_`)
- Column names are case-sensitive in SQLite
- Check for leading/trailing spaces in column headers

### Natural Language to SQL not working
- Verify Ollama is running: `curl http://localhost:11434/api/tags`
- Check that a model is installed: `ollama list`
- Ensure the Ollama URL is correct (default: `http://localhost:11434`)
- Try rephrasing your question more clearly
- Check the browser console for error messages

---

## How It Works

1. **Upload** → File is read into memory as a SQLite in-memory database
2. **Auto-table creation** → Each sheet becomes a table with the sheet name
3. **Natural Language (Optional)** → Local Ollama model converts your question to SQL
4. **Query** → Your SQL runs directly against SQLite
5. **Export** → Results can be saved to CSV

All processing is local; your data never leaves your machine.

---

## Contributing

Found a bug or have a feature idea? Open an issue or submit a pull request!

---

## License

MIT License — see LICENSE file for details.
