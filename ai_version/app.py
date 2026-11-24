"""ExcelSQL Studio 2.0 - a local-first SQLite & spreadsheet workbench.

Run with:  python app.py            (opens the browser automatically)
           python app.py --db path/to/database.sqlite
"""

from __future__ import annotations

import argparse
import hmac
import json
import os
import re
import secrets
import socket
import sqlite3
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path
from urllib.parse import quote

from flask import (
    Flask,
    Response,
    jsonify,
    render_template,
    request,
    stream_with_context,
)
from werkzeug.exceptions import HTTPException, RequestEntityTooLarge

from core.database import (
    Database,
    DatabaseError,
    sanitize_identifier,
    validate_sqlite_file,
)
from core.demo import create_demo_database
from core.spreadsheet import import_sheet, is_spreadsheet, open_spreadsheet, preview_sheet
from core.xlsx import XlsxReader

APP_VERSION = "2.0.0"
BASE_DIR = Path(__file__).resolve().parent
IS_FROZEN = getattr(sys, "frozen", False)
BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
MAX_UPLOAD_BYTES = 512 * 1024 * 1024


def resource_path(relative: str) -> Path:
    """Locate bundled templates/static assets in source or PyInstaller mode."""
    return BUNDLE_DIR / relative


def _user_data_root() -> Path:
    if sys.platform == "win32":
        return Path(os.environ.get("APPDATA") or (Path.home() / "AppData" / "Roaming"))
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support"
    return Path(os.environ.get("XDG_DATA_HOME") or (Path.home() / ".local" / "share"))


def workspace_dir() -> Path:
    """Use a stable user-writable location when packaged as a desktop app."""
    path = _user_data_root() / "ExcelSQL Studio" if IS_FROZEN else BASE_DIR / ".workspace"
    path.mkdir(parents=True, exist_ok=True)
    return path


WORKSPACE = workspace_dir()
DB_DIR = WORKSPACE / "databases"
UPLOAD_DIR = WORKSPACE / "uploads"
CONFIG_FILE = WORKSPACE / "config.json"
DEMO_FILE = DB_DIR / "demo.sqlite"

for folder in (DB_DIR, UPLOAD_DIR):
    folder.mkdir(parents=True, exist_ok=True)

TOKEN = secrets.token_urlsafe(32)

app = Flask(
    __name__,
    static_folder=str(resource_path("static")),
    template_folder=str(resource_path("templates")),
)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
app.json.sort_keys = False


class ApiError(Exception):
    def __init__(self, message: str, status: int = 400, code: str = "ERROR"):
        super().__init__(message)
        self.message = message
        self.status = status
        self.code = code


STATE_LOCK = threading.Lock()
STATE: dict = {"db": None}
UPLOADS: dict[str, dict] = {}


# --------------------------------------------------------------------- utils
def _display_name(filename: str) -> str:
    name = os.path.basename((filename or "").replace("\\", "/"))
    name = re.sub(r"[\x00-\x1f\x7f]", "", name).strip()
    return name[:120] or "upload"


def load_config() -> dict:
    try:
        with CONFIG_FILE.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, dict):
            return {}
        return data
    except (OSError, json.JSONDecodeError):
        return {}


def save_config(config: dict) -> None:
    tmp = CONFIG_FILE.with_suffix(".tmp")
    try:
        with tmp.open("w", encoding="utf-8") as fh:
            json.dump(config, fh, indent=2)
        tmp.replace(CONFIG_FILE)
    except OSError:
        pass


def add_recent(path: str) -> None:
    config = load_config()
    recent = [
        entry for entry in config.get("recent", []) if entry.get("path") != path
    ]
    recent.insert(0, {"path": path, "lastOpened": int(time.time() * 1000)})
    config["recent"] = recent[:12]
    save_config(config)


def recent_payload() -> list[dict]:
    result = []
    for entry in load_config().get("recent", [])[:12]:
        path = entry.get("path", "")
        if not path:
            continue
        result.append(
            {
                "path": path,
                "name": Path(path).name,
                "exists": Path(path).is_file(),
                "lastOpened": entry.get("lastOpened", 0),
            }
        )
    return result


def set_db(db: Database | None) -> None:
    with STATE_LOCK:
        STATE["db"] = db


def get_db() -> Database | None:
    with STATE_LOCK:
        return STATE["db"]


def require_db() -> Database:
    db = get_db()
    if db is None:
        raise ApiError("No database is open. Open a SQLite file or import data first.", 409, "NO_DATABASE")
    return db


def status_payload() -> dict:
    db = get_db()
    if db is not None and not Path(db.path).is_file():
        # The active file may be deleted or moved from outside the app.
        set_db(None)
        db = None
    return {
        "connected": db is not None,
        "database": db.info() if db is not None else None,
        "sqliteVersion": sqlite3.sqlite_version,
        "recent": recent_payload(),
        "app": {"version": APP_VERSION, "workspace": str(WORKSPACE)},
    }


def _resolve_db_path(raw: str) -> Path:
    text = (raw or "").strip().strip('"')
    if not text:
        raise ApiError("Please provide a database path.", 400, "INVALID_PATH")
    path = Path(text).expanduser()
    if not path.is_absolute():
        path = DB_DIR / path
    return path.resolve()


def _open_db(path: Path, read_only: bool) -> Database:
    db = Database(path, read_only=read_only)
    set_db(db)
    add_recent(db.path)
    return db


def _cleanup_uploads() -> None:
    cutoff = time.time() - 24 * 3600
    for item in UPLOAD_DIR.glob("*"):
        try:
            if item.is_file() and item.stat().st_mtime < cutoff:
                item.unlink()
        except OSError:
            pass


_cleanup_uploads()


# ------------------------------------------------------------------ security
@app.before_request
def verify_token():
    if request.path.startswith("/api/"):
        supplied = request.headers.get("X-App-Token", "")
        if not supplied or not hmac.compare_digest(supplied, TOKEN):
            raise ApiError("Invalid application token. Reload the page.", 403, "BAD_TOKEN")


@app.after_request
def security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "no-referrer"
    if request.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store"
    return response


@app.errorhandler(ApiError)
def handle_api_error(exc: ApiError):
    return jsonify({"error": exc.message, "code": exc.code}), exc.status


@app.errorhandler(DatabaseError)
def handle_database_error(exc: DatabaseError):
    return jsonify({"error": exc.message, "code": exc.code}), 400


@app.errorhandler(RequestEntityTooLarge)
def handle_too_large(exc: RequestEntityTooLarge):
    return (
        jsonify(
            {
                "error": "The uploaded file is too large. Maximum size is 512 MB.",
                "code": "FILE_TOO_LARGE",
            }
        ),
        413,
    )


@app.errorhandler(HTTPException)
def handle_http_error(exc: HTTPException):
    if request.path.startswith("/api/"):
        return jsonify({"error": exc.description, "code": exc.name}), exc.code
    return exc


# --------------------------------------------------------------------- pages
@app.get("/")
def index():
    response = app.make_response(
        render_template(
            "index.html",
            token=TOKEN,
            version=APP_VERSION,
            sqlite_version=sqlite3.sqlite_version,
        )
    )
    response.headers["Cache-Control"] = "no-store"
    return response


# ----------------------------------------------------------------------- api
@app.get("/api/status")
def api_status():
    return jsonify(status_payload())


@app.get("/api/schema")
def api_schema():
    db = require_db()
    schema = db.schema()
    return jsonify(
        {
            "tables": schema["tables"],
            "views": schema["views"],
            "database": db.info(),
        }
    )


@app.post("/api/database/open")
def api_open_database():
    data = request.get_json(silent=True) or {}
    path = _resolve_db_path(data.get("path", ""))
    if not path.exists():
        raise ApiError(f"Database not found: {path}", 404, "NOT_FOUND")
    db = _open_db(path, bool(data.get("readOnly")))
    return jsonify({"ok": True, "database": db.info(), "status": status_payload()})


@app.post("/api/database/new")
def api_new_database():
    data = request.get_json(silent=True) or {}
    raw = (data.get("path") or "").strip()
    if not raw:
        raise ApiError("Please provide a file name or path.", 400, "INVALID_PATH")
    path = _resolve_db_path(raw)
    if path.suffix == "":
        path = path.with_suffix(".db")
    if path.exists():
        raise ApiError(
            f"A file already exists at {path}. Open it instead.", 409, "ALREADY_EXISTS"
        )
    if not path.parent.exists():
        raise ApiError(
            f"The folder does not exist: {path.parent}", 400, "NO_FOLDER"
        )
    try:
        sqlite3.connect(str(path)).execute("VACUUM").connection.close()
    except sqlite3.Error as exc:
        path.unlink(missing_ok=True)
        raise ApiError(f"Could not create database: {exc}", 400, "CREATE_FAILED") from exc
    db = _open_db(path, False)
    return jsonify({"ok": True, "database": db.info(), "status": status_payload()})


@app.post("/api/database/demo")
def api_demo_database():
    create_demo_database(str(DEMO_FILE))
    db = _open_db(DEMO_FILE, False)
    return jsonify({"ok": True, "database": db.info(), "status": status_payload()})


@app.post("/api/database/close")
def api_close_database():
    set_db(None)
    return jsonify({"ok": True, "status": status_payload()})


@app.post("/api/database/upload")
def api_upload_database():
    uploaded = request.files.get("file")
    if uploaded is None:
        raise ApiError("No file was uploaded.", 400, "NO_FILE")
    original = _display_name(uploaded.filename or "database.db")
    ext = Path(original).suffix.lower()
    if ext == ".xls":
        raise ApiError(
            "Legacy .xls files are not supported. Save as .xlsx or .csv first.",
            400,
            "UNSUPPORTED_FILE",
        )
    stamp = time.strftime("%Y%m%d-%H%M%S")
    stem = sanitize_identifier(Path(original).stem, "database")
    target = DB_DIR / f"{stem}-copy-{stamp}.db"
    counter = 2
    while target.exists():
        target = DB_DIR / f"{stem}-copy-{stamp}-{counter}.db"
        counter += 1
    uploaded.save(target)
    try:
        db = _open_db(target, False)
    except Exception:
        target.unlink(missing_ok=True)
        raise
    return jsonify(
        {
            "ok": True,
            "copied": True,
            "originalName": original,
            "database": db.info(),
            "status": status_payload(),
        }
    )


@app.post("/api/database/recents/remove")
def api_remove_recent():
    data = request.get_json(silent=True) or {}
    path = (data.get("path") or "").strip()
    config = load_config()
    config["recent"] = [
        entry for entry in config.get("recent", []) if entry.get("path") != path
    ]
    save_config(config)
    return jsonify({"ok": True, "recent": recent_payload()})


DIALOG_SCRIPT = r"""
import sys
import tkinter as tk
from tkinter import filedialog

mode = sys.argv[1]
title = sys.argv[2]
initialdir = sys.argv[3] or None
initialfile = sys.argv[4] or None
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
root.update_idletasks()
filetypes = [
    ("SQLite databases", "*.db *.sqlite *.sqlite3 *.db3"),
    ("All files", "*.*"),
]
if mode == "save":
    path = filedialog.asksaveasfilename(
        title=title, initialdir=initialdir, initialfile=initialfile,
        defaultextension=".db", filetypes=filetypes,
    )
else:
    path = filedialog.askopenfilename(
        title=title, initialdir=initialdir, filetypes=filetypes,
    )
root.destroy()
sys.stdout.write(path or "")
"""


@app.post("/api/dialog")
def api_dialog():
    data = request.get_json(silent=True) or {}
    mode = "save" if data.get("mode") == "save" else "open"
    title = "Create SQLite database" if mode == "save" else "Open SQLite database"
    initialdir = str(Path.home())
    initialfile = data.get("suggestedName") or ""
    command = [
        sys.executable,
        "-c",
        DIALOG_SCRIPT,
        mode,
        title,
        initialdir,
        initialfile,
    ]
    kwargs = {"capture_output": True, "text": True, "timeout": 600}
    if os.name == "nt":
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        proc = subprocess.run(command, **kwargs)
    except subprocess.TimeoutExpired as exc:
        raise ApiError("The file dialog timed out.", 408, "DIALOG_TIMEOUT") from exc
    if proc.returncode != 0:
        detail = (proc.stderr or "").strip().splitlines()
        message = detail[-1] if detail else "Could not open the system file dialog."
        raise ApiError(
            f"{message} You can type the file path manually instead.",
            500,
            "DIALOG_FAILED",
        )
    path = (proc.stdout or "").strip()
    return jsonify({"path": path or None})


@app.get("/api/table/data")
def api_table_data():
    db = require_db()
    name = request.args.get("name", "")
    try:
        limit = int(request.args.get("limit", 100))
        offset = int(request.args.get("offset", 0))
    except ValueError as exc:
        raise ApiError("Invalid pagination values.", 400, "INVALID_PAGINATION") from exc
    return jsonify(
        db.read_table(
            name,
            limit=limit,
            offset=offset,
            order_by=request.args.get("orderBy") or None,
            order_dir=request.args.get("orderDir", "asc"),
            search=request.args.get("search", ""),
        )
    )


@app.post("/api/query")
def api_query():
    db = require_db()
    data = request.get_json(silent=True) or {}
    sql = data.get("sql", "")
    max_rows = data.get("maxRows", 5000)
    try:
        max_rows = int(max_rows)
    except (TypeError, ValueError):
        max_rows = 5000
    result = db.run_query(sql, max_rows=max_rows)
    result["database"] = db.info()
    return jsonify(result)


@app.get("/api/export/table")
def api_export_table():
    db = require_db()
    name = request.args.get("name", "")
    # Validate up-front so errors are JSON, not broken downloads.
    db.read_table(name, limit=1, offset=0)
    filename = sanitize_identifier(name, "table") + ".csv"

    def generate():
        yield "\ufeff"
        yield from db.iter_table_csv(name)

    response = Response(stream_with_context(generate()), mimetype="text/csv")
    response.headers["Content-Disposition"] = (
        f"attachment; filename=\"{filename}\"; filename*=UTF-8''{quote(filename)}"
    )
    response.headers["Cache-Control"] = "no-store"
    return response


# ------------------------------------------------------------------- imports
@app.post("/api/import/upload")
def api_import_upload():
    uploaded = request.files.get("file")
    if uploaded is None:
        raise ApiError("No file was uploaded.", 400, "NO_FILE")
    original = _display_name(uploaded.filename or "data.xlsx")
    ext = Path(original).suffix.lower()
    if ext == ".xls":
        raise ApiError(
            "Legacy .xls files are not supported. Save the workbook as .xlsx "
            "or export it to CSV, then try again.",
            400,
            "UNSUPPORTED_FILE",
        )
    if not is_spreadsheet(original):
        raise ApiError(
            "Unsupported file. Choose an .xlsx, .xlsm, .csv, .tsv or .txt file.",
            400,
            "UNSUPPORTED_FILE",
        )
    file_id = secrets.token_hex(12)
    target = UPLOAD_DIR / f"{file_id}{ext}"
    uploaded.save(target)
    csv_sheet = None
    if ext in (".csv", ".tsv", ".txt"):
        csv_sheet = Path(original).stem or "Sheet1"
    try:
        spreadsheet = open_spreadsheet(str(target), csv_sheet_name=csv_sheet)
        sheets = spreadsheet.sheet_names()
        spreadsheet.close()
    except Exception as exc:  # noqa: BLE001
        target.unlink(missing_ok=True)
        raise ApiError(f"Could not read file: {exc}", 400, "READ_FAILED") from exc
    UPLOADS[file_id] = {
        "path": str(target),
        "name": original,
        "size": target.stat().st_size,
        "ext": ext,
        "sheets": sheets,
        "csvSheet": csv_sheet,
        "created": time.time(),
    }
    return jsonify(
        {
            "ok": True,
            "fileId": file_id,
            "name": original,
            "size": UPLOADS[file_id]["size"],
            "sheets": sheets,
            "kind": "xlsx" if ext not in (".csv", ".tsv", ".txt") else "csv",
        }
    )


@app.post("/api/import/preview")
def api_import_preview():
    data = request.get_json(silent=True) or {}
    upload = UPLOADS.get(data.get("fileId", ""))
    if upload is None:
        raise ApiError("Upload expired. Please choose the file again.", 404, "UPLOAD_EXPIRED")
    sheet = data.get("sheet") or upload["sheets"][0]
    if sheet not in upload["sheets"]:
        raise ApiError(f"Sheet not found: {sheet}", 404, "SHEET_NOT_FOUND")
    preview = preview_sheet(
        upload["path"],
        sheet,
        limit=8,
        csv_sheet_name=upload.get("csvSheet"),
    )
    preview.update({"sheet": sheet, "sheets": upload["sheets"]})
    return jsonify(preview)


@app.post("/api/import/run")
def api_import_run():
    data = request.get_json(silent=True) or {}
    upload = UPLOADS.get(data.get("fileId", ""))
    if upload is None:
        raise ApiError("Upload expired. Please choose the file again.", 404, "UPLOAD_EXPIRED")
    sheets = data.get("sheets") or []
    if not sheets:
        raise ApiError("Select at least one sheet to import.", 400, "NO_SHEETS")
    has_header = bool(data.get("hasHeader", True))
    mode = "append" if data.get("mode") == "append" else "replace"

    db = get_db()
    created_database = False
    if db is None:
        imports_path = DB_DIR / "imported-data.sqlite"
        db = Database(imports_path, read_only=False)
        set_db(db)
        add_recent(db.path)
        created_database = True
    if db.read_only:
        raise ApiError(
            "The current database is open in read-only mode. Reopen it writable to import data.",
            400,
            "READ_ONLY",
        )

    used: set[str] = set()
    imported: list[dict] = []
    errors: list[dict] = []
    for item in sheets:
        sheet = item.get("sheet", "")
        if sheet not in upload["sheets"]:
            errors.append({"sheet": sheet, "error": "Sheet not found."})
            continue
        base = sanitize_identifier(item.get("table") or sheet, "imported")
        table = base
        counter = 2
        while table.lower() in used:
            table = f"{base[:56]}_{counter}"
            counter += 1
        used.add(table.lower())
        try:
            summary = import_sheet(
                db.path,
                upload["path"],
                sheet,
                table,
                mode=mode,
                has_header=has_header,
                csv_sheet_name=upload.get("csvSheet"),
            )
            imported.append(summary)
        except Exception as exc:  # noqa: BLE001
            errors.append({"sheet": sheet, "error": str(exc)})

    schema = db.schema() if imported else {"tables": [], "views": []}
    return jsonify(
        {
            "ok": not errors,
            "imported": imported,
            "errors": errors,
            "createdDatabase": created_database,
            "database": db.info(),
            "status": status_payload(),
            "schema": schema,
        }
    )


@app.post("/api/import/forget")
def api_import_forget():
    data = request.get_json(silent=True) or {}
    file_id = data.get("fileId", "")
    upload = UPLOADS.pop(file_id, None)
    if upload:
        try:
            Path(upload["path"]).unlink(missing_ok=True)
        except OSError:
            pass
    return jsonify({"ok": True})


# --------------------------------------------------------------------- start
def _port_is_free(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind((host, port))
            return True
        except OSError:
            return False


def main() -> None:
    parser = argparse.ArgumentParser(description="ExcelSQL Studio 2.0")
    parser.add_argument("--db", help="SQLite database to open at startup")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--read-only", action="store_true")
    args = parser.parse_args()

    if args.db:
        try:
            path = _resolve_db_path(args.db)
            validate_sqlite_file(path)
            _open_db(path, args.read_only)
            print(f"  Opened database: {path}")
        except Exception as exc:  # noqa: BLE001
            print(f"  Could not open database: {exc}")

    port = args.port
    for candidate in range(args.port, args.port + 30):
        if _port_is_free(args.host, candidate):
            port = candidate
            break
    else:
        print("No free port found.")
        return

    url = f"http://{args.host}:{port}/"
    print()
    print("  ExcelSQL Studio 2.0")
    print("  ─────────────────────────────────────────────")
    print(f"  Local URL : {url}")
    print(f"  Workspace : {WORKSPACE}")
    print("  Press Ctrl+C to stop the server.")
    print()
    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    app.run(host=args.host, port=port, debug=False, threaded=True, use_reloader=False)


if __name__ == "__main__":
    main()
