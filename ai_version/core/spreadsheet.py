"""Spreadsheet reading and import into SQLite (CSV / TSV / XLSX)."""

from __future__ import annotations

import csv
import io
import os
import re
import sqlite3
from pathlib import Path
from typing import Any, Iterator

from .database import DatabaseError, quote_identifier
from .xlsx import XlsxReader

csv.field_size_limit(20_000_000)

XLSX_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
CSV_EXTENSIONS = {".csv", ".tsv", ".txt"}
SUPPORTED_EXTENSIONS = XLSX_EXTENSIONS | CSV_EXTENSIONS

INTEGER_RE = re.compile(r"^[+-]?\d+$")
FLOAT_RE = re.compile(r"^[+-]?(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?$")
SAMPLE_ROWS = 500
INSERT_CHUNK = 1000


def is_spreadsheet(path: str | os.PathLike[str]) -> bool:
    return Path(path).suffix.lower() in SUPPORTED_EXTENSIONS


def _preview_cell(value: Any) -> Any:
    if isinstance(value, str) and len(value) > 300:
        return value[:300] + f" … [{len(value):,} chars]"
    if isinstance(value, (bytes, bytearray)):
        return f"<BLOB {len(value):,} bytes>"
    return value


def _clean_row(row: list[Any]) -> list[Any]:
    values = [None if value == "" else value for value in row]
    while values and values[-1] is None:
        values.pop()
    return values


def _detect_csv(path: str, ext: str) -> tuple[str, str]:
    with open(path, "rb") as fh:
        raw = fh.read(131_072)
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        encoding = "utf-16"
    elif raw.startswith(b"\xef\xbb\xbf"):
        encoding = "utf-8-sig"
    else:
        encoding = "utf-8"
        try:
            raw.decode("utf-8")
        except UnicodeDecodeError:
            for candidate in ("cp1252", "latin-1"):
                try:
                    raw.decode(candidate)
                    encoding = candidate
                    break
                except UnicodeDecodeError:
                    continue

    if ext == ".tsv":
        delimiter = "\t"
    else:
        sample = raw.decode(encoding, errors="replace")[:16_384]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            delimiter = dialect.delimiter
        except csv.Error:
            delimiter = ","
    return encoding, delimiter


class Spreadsheet:
    """Unified reader for CSV-like files and XLSX workbooks."""

    def __init__(self, path: str, csv_sheet_name: str | None = None):
        self.path = path
        self.ext = Path(path).suffix.lower()
        if self.ext not in SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Unsupported file type: {self.ext or 'unknown'}. "
                "Use .xlsx, .xlsm, .csv, .tsv or .txt."
            )
        self._xlsx: XlsxReader | None = None
        self._csv_sheet: str | None = None
        if self.ext in XLSX_EXTENSIONS:
            self._xlsx = XlsxReader(path)
        else:
            self._csv_sheet = csv_sheet_name or Path(path).stem or "Sheet1"

    def sheet_names(self) -> list[str]:
        if self._xlsx is not None:
            return self._xlsx.sheet_names()
        return [self._csv_sheet or "Sheet1"]

    def iter_rows(self, sheet: str) -> Iterator[list[Any]]:
        if self._xlsx is not None:
            yield from self._xlsx.iter_rows(sheet)
            return
        yield from self._iter_csv(sheet)

    def _iter_csv(self, sheet: str) -> Iterator[list[Any]]:
        if sheet != self._csv_sheet:
            raise ValueError(f"Sheet not found: {sheet}")
        encoding, delimiter = _detect_csv(self.path, self.ext)
        with open(self.path, "r", encoding=encoding, errors="replace", newline="") as fh:
            reader = csv.reader(fh, delimiter=delimiter)
            for row in reader:
                cleaned = _clean_row(row)
                if cleaned:
                    yield cleaned

    def close(self) -> None:
        if self._xlsx is not None:
            self._xlsx.close()
            self._xlsx = None


def open_spreadsheet(path: str, csv_sheet_name: str | None = None) -> Spreadsheet:
    try:
        return Spreadsheet(path, csv_sheet_name=csv_sheet_name)
    except ValueError:
        raise
    except Exception as exc:  # noqa: BLE001 - surface a useful message
        raise ValueError(f"Could not read spreadsheet: {exc}") from exc


def preview_sheet(
    path: str, sheet: str, limit: int = 8, csv_sheet_name: str | None = None
) -> dict[str, Any]:
    spreadsheet = open_spreadsheet(path, csv_sheet_name=csv_sheet_name)
    rows: list[list[Any]] = []
    width = 0
    try:
        for row in spreadsheet.iter_rows(sheet):
            cleaned = _clean_row(row)
            if not cleaned:
                continue
            width = max(width, len(cleaned))
            rows.append([_preview_cell(value) for value in cleaned])
            if len(rows) >= limit:
                break
    finally:
        spreadsheet.close()
    rows = [row + [None] * (width - len(row)) for row in rows]
    return {"rows": rows, "width": width}


def _make_headers(raw: list[Any], width: int) -> list[str]:
    headers: list[str] = []
    used: set[str] = set()
    for index in range(width):
        value = raw[index] if index < len(raw) else None
        name = str(value).strip() if value is not None else ""
        name = re.sub(r"[^\w$]+", "_", name, flags=re.UNICODE)
        name = re.sub(r"_+", "_", name).strip("_")
        if not name:
            name = f"column_{index + 1}"
        if name[0].isdigit():
            name = "c_" + name
        name = name[:60]
        base = name
        counter = 2
        while name.lower() in used:
            name = f"{base[:56]}_{counter}"
            counter += 1
        used.add(name.lower())
        headers.append(name)
    return headers


def _infer_types(rows: list[list[Any]], width: int) -> list[str]:
    kinds = [set() for _ in range(width)]
    for row in rows:
        for index in range(min(width, len(row))):
            value = row[index]
            if value is None:
                continue
            if isinstance(value, bool):
                kinds[index].add("int")
            elif isinstance(value, int):
                kinds[index].add("int")
            elif isinstance(value, float):
                kinds[index].add("real")
            elif isinstance(value, str):
                text = value.strip()
                if not text:
                    continue
                if INTEGER_RE.match(text):
                    kinds[index].add("int")
                elif FLOAT_RE.match(text):
                    kinds[index].add("real")
                else:
                    kinds[index].add("text")
            else:
                kinds[index].add("text")
    types: list[str] = []
    for kind_set in kinds:
        if not kind_set:
            types.append("TEXT")
        elif "text" in kind_set:
            types.append("TEXT")
        elif "real" in kind_set:
            types.append("REAL")
        else:
            types.append("INTEGER")
    return types


def _coerce(value: Any, column_type: str) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        return 1 if value else 0
    if column_type == "INTEGER":
        if isinstance(value, int):
            return value
        if isinstance(value, float) and value.is_integer():
            return int(value)
        if isinstance(value, str) and INTEGER_RE.match(value.strip()):
            try:
                return int(value.strip())
            except ValueError:
                return value
        return value
    if column_type == "REAL":
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            text = value.strip()
            if FLOAT_RE.match(text):
                try:
                    return float(text)
                except ValueError:
                    return value
        return value
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float)):
        return str(value)
    return str(value)


def table_exists(conn: sqlite3.Connection, name: str) -> str | None:
    row = conn.execute(
        "SELECT type FROM sqlite_master WHERE type IN ('table','view') "
        "AND name = ? COLLATE NOCASE",
        (name,),
    ).fetchone()
    return row[0] if row else None


def import_sheet(
    db_path: str,
    source_path: str,
    sheet: str,
    table_name: str,
    mode: str = "replace",
    has_header: bool = True,
    csv_sheet_name: str | None = None,
) -> dict[str, Any]:
    """Import one sheet into a SQLite database. Returns a summary."""

    from .database import sanitize_identifier

    table = sanitize_identifier(table_name or sheet, "imported")
    if mode not in ("replace", "append"):
        raise DatabaseError("Import mode must be 'replace' or 'append'.", "INVALID_MODE")

    spreadsheet = open_spreadsheet(source_path, csv_sheet_name=csv_sheet_name)
    conn: sqlite3.Connection | None = None
    try:
        iterator = spreadsheet.iter_rows(sheet)
        first = next(iterator, None)
        if first is None:
            raise DatabaseError(f"Sheet '{sheet}' is empty.", "EMPTY_SHEET")
        first = _clean_row(first)

        # Buffer a sample so type inference and column widths are reliable.
        buffer: list[list[Any]] = []
        if has_header:
            raw_header = first
            try:
                first_data = _clean_row(next(iterator))
            except StopIteration:
                first_data = None
            if first_data is not None:
                buffer.append(first_data)
        else:
            raw_header = []
            buffer.append(first)

        for row in iterator:
            cleaned = _clean_row(row)
            if cleaned:
                buffer.append(cleaned)
            if len(buffer) >= SAMPLE_ROWS:
                break

        width = max(
            [len(raw_header)] + [len(row) for row in buffer] + [1]
        )
        headers = _make_headers(raw_header, width)
        if not has_header:
            headers = [f"column_{i + 1}" for i in range(width)]

        buffered_rows = [
            (row + [None] * (width - len(row)))[:width] for row in buffer
        ]
        types = _infer_types(buffered_rows, width)

        conn = sqlite3.connect(db_path, timeout=10)
        conn.execute("PRAGMA busy_timeout = 8000")
        conn.execute("PRAGMA foreign_keys = ON")
        existing_type = table_exists(conn, table)

        if existing_type == "view":
            raise DatabaseError(
                f"'{table}' is a view and cannot be replaced.", "IS_VIEW"
            )

        if existing_type and mode == "append":
            existing_cols = [
                row[1] for row in conn.execute(f"PRAGMA table_info({quote_identifier(table)})")
            ]
            if len(existing_cols) != width:
                raise DatabaseError(
                    f"Cannot append: '{table}' has {len(existing_cols)} columns "
                    f"but this sheet has {width}.",
                    "COLUMN_MISMATCH",
                )
            target_columns = existing_cols
            created = False
        else:
            if existing_type:
                conn.execute(f"DROP TABLE {quote_identifier(table)}")
            columns_sql = ", ".join(
                f"{quote_identifier(name)} {sql_type}"
                for name, sql_type in zip(headers, types)
            )
            conn.execute(f"CREATE TABLE {quote_identifier(table)} ({columns_sql})")
            target_columns = headers
            created = True

        column_sql = ", ".join(quote_identifier(name) for name in target_columns)
        placeholders = ", ".join("?" for _ in target_columns)
        insert_sql = (
            f"INSERT INTO {quote_identifier(table)} ({column_sql}) "
            f"VALUES ({placeholders})"
        )

        def row_values(row: list[Any]) -> list[Any]:
            padded = (row + [None] * (width - len(row)))[:width]
            return [_coerce(value, types[index]) for index, value in enumerate(padded)]

        inserted = 0
        chunk = [row_values(row) for row in buffered_rows]
        if chunk:
            conn.executemany(insert_sql, chunk)
            inserted += len(chunk)

        pending: list[list[Any]] = []
        for row in iterator:
            cleaned = _clean_row(row)
            if not cleaned:
                continue
            pending.append(row_values(cleaned))
            if len(pending) >= INSERT_CHUNK:
                conn.executemany(insert_sql, pending)
                inserted += len(pending)
                pending = []
        if pending:
            conn.executemany(insert_sql, pending)
            inserted += len(pending)

        conn.commit()
        return {
            "sheet": sheet,
            "table": table,
            "columns": len(target_columns),
            "rows": inserted,
            "mode": "replace" if created else "append",
            "created": created,
        }
    except sqlite3.Error as exc:
        if conn is not None:
            conn.rollback()
        raise DatabaseError(f"Import failed: {exc}", "IMPORT_FAILED") from exc
    finally:
        spreadsheet.close()
        if conn is not None:
            conn.close()
