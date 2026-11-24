# ExcelSQL Studio 2.0

A local-first, browser-based SQLite workbench with a DBeaver-inspired interface. It can open existing SQLite databases, browse tables and views, run SQL, and import Excel (`.xlsx`) or CSV files into SQLite tables.

Everything runs on your machine: the Python Flask server binds to `127.0.0.1` by default, and every API request is protected by a per-session token injected into the page.

## Features

- **Database connection**
  - Open an existing `.sqlite` / `.db` file by path.
  - Upload a SQLite file (opened as a writable copy inside the workspace).
  - Create a new empty database.
  - Generate a deterministic demo database with customers, products, orders, and a view.
  - Recent database list with missing-file detection.
- **Schema explorer**
  - Searchable table/view tree with column details and row counts.
  - Double-click a table to browse its data.
  - Context actions: view data, generate `SELECT`, copy quoted column names.
- **SQL workspace**
  - Multiple query tabs with syntax highlighting, line numbers, and autocomplete.
  - Multi-statement execution with per-statement result sets.
  - Query history and result CSV export.
  - `Ctrl+Enter` to run, `Ctrl+T` for a new query tab.
- **Data browser**
  - Sortable, searchable, paginated grid for tables and views.
  - Cell inspector for long or binary-looking values.
  - Per-table CSV export.
- **Import wizard**
  - XLSX and CSV preview, sheet selection, header toggle, and table-name editing.
  - Import into the active database, or automatically create `imported-data.sqlite`.
  - Append or replace modes.
- **Interface**
  - Dark and light themes, resizable editor/results panes, drag-and-drop file opening,
    toasts, modals, and a responsive layout for smaller screens.

## Requirements

- Python 3.10 or newer
- Flask 3.x (install with the command below)
- A modern browser with ES modules support

No pandas, openpyxl, SQLAlchemy, Node build step, or frontend package manager is required. The XLSX reader is implemented with the Python standard library.

## Run

From this `ai_version` directory:

```bash
python3 -m pip install -r requirements.txt
python3 app.py
```

The app opens `http://127.0.0.1:8765/` in your browser. If that port is busy, the next free port is used and printed in the terminal.

Convenience scripts:

```bash
./run.sh                 # Unix/macOS; installs Flask if missing
./run.sh --no-browser    # print the URL without opening a browser
run.bat                  # Windows
```

Useful options:

```bash
python3 app.py --db /path/to/database.sqlite
python3 app.py --db /path/to/database.sqlite --read-only
python3 app.py --port 8899 --no-browser
```

Press `Ctrl+C` in the terminal to stop the server.

## Desktop app (pywebview + PyInstaller)

The browser UI can run in a native window without changing the frontend. It still uses the same local Flask server; only the window shell is native.

### Run the desktop window during development

```bash
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
python -m pip install -r requirements-desktop.txt
python desktop.py
```

The build scripts create and use `.venv/` automatically. On Debian/Ubuntu, install `python3-venv` first if virtual environment creation fails. `desktop.py` picks a free local port, starts Flask in a background thread, and opens a native window. Closing the window stops the server.

### Build an executable

```bash
./build_desktop.sh        # Unix/macOS
build_desktop.bat         # Windows
```

The output is in `dist/ExcelSQL Studio/`. To debug startup problems, remove `--windowed` from the build script so terminal output remains visible.

Optional: place `assets/app.ico` (Windows) or `assets/app.icns` (macOS) in the project and the build script will use it automatically.

### Desktop-specific notes

- In desktop mode the workspace moves to a user-writable data directory:
  - Windows: `%APPDATA%/ExcelSQL Studio`
  - macOS: `~/Library/Application Support/ExcelSQL Studio`
  - Linux: `$XDG_DATA_HOME/ExcelSQL Studio` or `~/.local/share/ExcelSQL Studio`
- Windows may require the Edge WebView2 runtime.
- Linux may require WebKitGTK packages for pywebview.
- Native folder pickers currently come from `tkinter`; some Linux distributions need `python3-tk` installed.

## Workspace

Runtime files live in `ai_version/.workspace/` and are ignored by Git:

```text
.workspace/
  databases/   # demo, created, and uploaded-copy databases
  uploads/     # temporary spreadsheet/SQLite uploads
  config.json  # recent database list
```

Delete this folder to reset the workspace. It is recreated automatically on the next start.

## Project layout

```text
ai_version/
  app.py                     Flask app, routes, session token, startup logic
  desktop.py                 Native window launcher (pywebview)
  build_desktop.sh/.bat      PyInstaller build scripts
  requirements-desktop.txt   Desktop packaging dependencies
  core/
    database.py              SQLite schema/query/browse/export helpers
    spreadsheet.py           CSV/XLSX preview and import orchestration
    xlsx.py                  Dependency-free XLSX parser
    demo.py                  Demo database generator
  static/
    css/                     Theme, layout, workspace, and modal styles
    js/                      App shell, API wrapper, editor, grid, utilities
  templates/index.html       Single-page HTML shell and SVG icon sprite
```

## Security notes

- The server is intended for local use only. Do not expose it to a public network.
- API calls require an `X-App-Token` header whose value is rendered into the page at request time; the token is regenerated whenever the server restarts.
- Uploaded SQLite databases are copied before opening, so the original file is never modified. Databases opened by path are editable in place unless read-only mode is selected.

## Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| `Ctrl/⌘ + O` | Open database |
| `Ctrl/⌘ + I` | Import spreadsheet |
| `Ctrl/⌘ + T` | New query tab |
| `Ctrl + Enter` | Run current query |
| `Ctrl/⌘ + Shift + H` | Toggle query history |
| `Alt + W` | Close current tab |
| `F1` | Help |
