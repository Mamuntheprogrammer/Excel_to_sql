"""SQLite database access layer for ExcelSQL Studio 2.0.

Everything here uses the Python standard library only.  A fresh connection is
opened per request so long-running queries never block the UI and so imports
can safely run in parallel with browsing.
"""

from __future__ import annotations

import csv
import io
import math
import os
import re
import sqlite3
import time
from pathlib import Path
from typing import Any, Iterable

SQLITE_MAGIC = b"SQLite format 3\x00"
QUERY_TIMEOUT_SECONDS = 120
COUNT_TIMEOUT_SECONDS = 15
MAX_TEXT_CHARS = 20_000

INTEGER_RE = re.compile(r"^[+-]?\d+$")
FLOAT_RE = re.compile(
    r"^[+-]?(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?$"
)


class DatabaseError(Exception):
    """A user-facing database error."""

    def __init__(self, message: str, code: str = "DB_ERROR"):
        super().__init__(message)
        self.message = message
        self.code = code


def quote_identifier(name: str) -> str:
    if not isinstance(name, str) or not name:
        raise DatabaseError("Invalid SQL identifier.", "INVALID_IDENTIFIER")
    return '"' + name.replace('"', '""') + '"'


def validate_sqlite_file(path: str | os.PathLike[str]) -> Path:
    p = Path(path).expanduser()
    if not p.exists():
        raise DatabaseError(f"File not found: {p}", "NOT_FOUND")
    if p.is_dir():
        raise DatabaseError("The selected path is a folder, not a database file.", "IS_DIRECTORY")
    try:
        size = p.stat().st_size
        with p.open("rb") as fh:
            header = fh.read(16)
    except OSError as exc:
        raise DatabaseError(f"Cannot read file: {exc}", "READ_ERROR") from exc
    if size > 0 and not header.startswith(SQLITE_MAGIC):
        raise DatabaseError(
            "That file is not a SQLite database. Supported types: .db, .sqlite, .sqlite3",
            "NOT_SQLITE",
        )
    return p.resolve()


def _sanitize_number(value: int | float) -> int | float | str:
    if isinstance(value, float):
        if not math.isfinite(value):
            return str(value)
        return value
    if isinstance(value, int) and not (-(2**53 - 1) <= value <= 2**53 - 1):
        return str(value)
    return value


def serialize_value(value: Any) -> Any:
    """Convert a sqlite value to something JSON-safe and UI-friendly."""
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return _sanitize_number(value)
    if isinstance(value, (bytes, bytearray, memoryview)):
        raw = bytes(value)
        return f"<BLOB {len(raw):,} bytes>"
    if isinstance(value, str):
        if len(value) > MAX_TEXT_CHARS:
            return value[:MAX_TEXT_CHARS] + f" … [truncated, {len(value):,} chars]"
        return value
    return str(value)


def sanitize_identifier(value: str, fallback: str = "value") -> str:
    text = (value or "").strip()
    text = re.sub(r"[^\w$]+", "_", text, flags=re.UNICODE)
    text = re.sub(r"_+", "_", text).strip("_")
    if not text:
        text = fallback
    if text[0].isdigit():
        text = "t_" + text
    return text[:60]


def _strip_comments(sql: str) -> str:
    """Best-effort removal used only to detect comment-only statements."""
    sql = re.sub(r"/\*.*?\*/", " ", sql, flags=re.S)
    sql = re.sub(r"--[^\n]*", " ", sql)
    return sql.strip()


def split_statements(sql: str) -> list[str]:
    """Split a SQL script on complete statements, respecting strings/comments."""
    statements: list[str] = []
    buffer = ""
    for char in sql:
        buffer += char
        if char == ";" and sqlite3.complete_statement(buffer):
            statements.append(buffer.strip())
            buffer = ""
    if buffer.strip():
        statements.append(buffer.strip())
    return [stmt for stmt in statements if _strip_comments(stmt)]


class Database:
    def __init__(self, path: str | os.PathLike[str], read_only: bool = False):
        resolved = Path(path).expanduser().resolve()
        self.path = str(resolved)
        self.name = resolved.name
        self.read_only = bool(read_only)
        if not resolved.exists():
            raise DatabaseError(f"File not found: {resolved}", "NOT_FOUND")
        if not self.read_only:
            validate_sqlite_file(resolved)
        elif resolved.stat().st_size == 0:
            raise DatabaseError("This database file is empty.", "EMPTY_DB")

        # Verify with an actual connection (also validates the SQLite header).
        try:
            conn = self.connect()
            conn.execute("SELECT count(*) FROM sqlite_master").fetchone()
            conn.close()
        except sqlite3.Error as exc:
            raise DatabaseError(f"Unable to open database: {exc}", "OPEN_FAILED") from exc

    # ------------------------------------------------------------------ conn
    def connect(self) -> sqlite3.Connection:
        if self.read_only:
            uri = Path(self.path).as_uri() + "?mode=ro"
            conn = sqlite3.connect(uri, uri=True, timeout=10)
            conn.execute("PRAGMA query_only = ON")
        else:
            conn = sqlite3.connect(self.path, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA busy_timeout = 8000")
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def info(self) -> dict[str, Any]:
        stat = os.stat(self.path)
        return {
            "name": self.name,
            "path": self.path,
            "readOnly": self.read_only,
            "size": stat.st_size,
            "modified": int(stat.st_mtime * 1000),
        }

    # ---------------------------------------------------------------- schema
    def schema(self) -> dict[str, Any]:
        tables: list[dict[str, Any]] = []
        views: list[dict[str, Any]] = []
        conn = self.connect()
        try:
            master = conn.execute(
                """
                SELECT name, type, sql
                FROM sqlite_master
                WHERE type IN ('table', 'view') AND name NOT LIKE 'sqlite_%'
                ORDER BY name COLLATE NOCASE
                """
            ).fetchall()
            for row in master:
                name = row["name"]
                kind = row["type"]
                quoted = quote_identifier(name)
                try:
                    cols = conn.execute(f"PRAGMA table_xinfo({quoted})").fetchall()
                except sqlite3.Error:
                    cols = conn.execute(f"PRAGMA table_info({quoted})").fetchall()
                columns = []
                for col in cols:
                    keys = col.keys()
                    hidden = col["hidden"] if "hidden" in keys else 0
                    if hidden in (1, 2):  # hidden / generated columns
                        continue
                    columns.append(
                        {
                            "name": col["name"],
                            "type": (col["type"] or "").upper() or "TEXT",
                            "notNull": bool(col["notnull"]),
                            "pk": int(col["pk"] or 0),
                            "default": col["dflt_value"],
                        }
                    )
                indexes = []
                if kind == "table":
                    try:
                        for idx in conn.execute(f"PRAGMA index_list({quoted})").fetchall():
                            indexes.append(
                                {
                                    "name": idx["name"],
                                    "unique": bool(idx["unique"]),
                                    "origin": idx["origin"],
                                }
                            )
                    except sqlite3.Error:
                        pass
                sql = row["sql"] or ""
                entry = {
                    "name": name,
                    "type": kind,
                    "columns": columns,
                    "indexes": indexes,
                    "hasRowid": "WITHOUT ROWID" not in sql.upper(),
                    "sql": sql,
                }
                if kind == "table":
                    tables.append(entry)
                else:
                    views.append(entry)
        finally:
            conn.close()
        return {"tables": tables, "views": views}

    def _find_relation(self, name: str) -> dict[str, Any]:
        if not name:
            raise DatabaseError("Table name is required.", "INVALID_NAME")
        for group in ("tables", "views"):
            for item in self.schema()[group]:
                if item["name"] == name:
                    return item
        raise DatabaseError(f"Table or view not found: {name}", "NOT_FOUND")

    # ------------------------------------------------------------- data view
    def read_table(
        self,
        name: str,
        limit: int = 100,
        offset: int = 0,
        order_by: str | None = None,
        order_dir: str = "asc",
        search: str = "",
    ) -> dict[str, Any]:
        relation = self._find_relation(name)
        column_names = [c["name"] for c in relation["columns"]]
        limit = max(1, min(int(limit or 100), 1000))
        offset = max(0, int(offset or 0))

        where_sql = ""
        params: list[Any] = []
        search = (search or "").strip()
        if search:
            parts = [
                f"instr(lower(CAST({quote_identifier(col)} AS TEXT)), lower(?)) > 0"
                for col in column_names
            ]
            where_sql = " WHERE " + " OR ".join(parts)
            params = [search] * len(column_names)

        order_sql = ""
        if order_by:
            allowed = {c.lower(): c for c in column_names}
            actual = allowed.get(order_by.lower())
            if not actual:
                raise DatabaseError(f"Invalid sort column: {order_by}", "INVALID_ORDER")
            direction = "DESC" if str(order_dir).lower() == "desc" else "ASC"
            order_sql = f" ORDER BY {quote_identifier(actual)} {direction}"

        table_sql = quote_identifier(relation["name"])
        start = time.perf_counter()
        conn = self.connect()
        try:
            total: int | None = None
            try:
                deadline = time.monotonic() + COUNT_TIMEOUT_SECONDS
                conn.set_progress_handler(
                    lambda: 1 if time.monotonic() > deadline else 0, 50_000
                )
                total = conn.execute(
                    f"SELECT COUNT(*) FROM {table_sql}{where_sql}", params
                ).fetchone()[0]
            except sqlite3.Error:
                total = None
            finally:
                conn.set_progress_handler(None, 0)

            cursor = conn.execute(
                f"SELECT * FROM {table_sql}{where_sql}{order_sql} LIMIT ? OFFSET ?",
                [*params, limit, offset],
            )
            raw_rows = cursor.fetchall()
            columns = []
            for idx, col in enumerate(relation["columns"]):
                columns.append(
                    {
                        "name": col["name"],
                        "type": col["type"],
                        "pk": col["pk"],
                    }
                )
            rows = [[serialize_value(v) for v in row] for row in raw_rows]
        finally:
            conn.close()
        elapsed = round((time.perf_counter() - start) * 1000, 1)
        return {
            "name": relation["name"],
            "kind": relation["type"],
            "columns": columns,
            "rows": rows,
            "total": total,
            "limit": limit,
            "offset": offset,
            "elapsedMs": elapsed,
            "search": search,
            "orderBy": order_by,
            "orderDir": "desc" if str(order_dir).lower() == "desc" else "asc",
        }

    # ---------------------------------------------------------------- query
    def run_query(self, sql: str, max_rows: int = 5000) -> dict[str, Any]:
        sql = (sql or "").strip()
        if not sql:
            raise DatabaseError("Query is empty.", "EMPTY_QUERY")
        statements = split_statements(sql)
        if not statements:
            raise DatabaseError("Query is empty (only comments found).", "EMPTY_QUERY")
        max_rows = max(1, min(int(max_rows or 5000), 100_000))

        results: list[dict[str, Any]] = []
        error: dict[str, Any] | None = None
        overall_start = time.perf_counter()
        conn = self.connect()
        try:
            deadline = time.monotonic() + QUERY_TIMEOUT_SECONDS
            conn.set_progress_handler(
                lambda: 1 if time.monotonic() > deadline else 0, 100_000
            )
            for index, statement in enumerate(statements):
                start = time.perf_counter()
                try:
                    cursor = conn.execute(statement)
                    elapsed = round((time.perf_counter() - start) * 1000, 1)
                    if cursor.description:
                        columns = [desc[0] for desc in cursor.description]
                        raw = cursor.fetchmany(max_rows + 1)
                        truncated = len(raw) > max_rows
                        raw = raw[:max_rows]
                        results.append(
                            {
                                "kind": "rows",
                                "statement": statement,
                                "index": index,
                                "columns": [{"name": c} for c in columns],
                                "rows": [
                                    [serialize_value(v) for v in row] for row in raw
                                ],
                                "rowCount": len(raw),
                                "truncated": truncated,
                                "elapsedMs": elapsed,
                            }
                        )
                    else:
                        results.append(
                            {
                                "kind": "exec",
                                "statement": statement,
                                "index": index,
                                "changes": max(cursor.rowcount, 0),
                                "lastRowId": cursor.lastrowid,
                                "elapsedMs": elapsed,
                                "rowCount": 0,
                                "truncated": False,
                            }
                        )
                    if conn.in_transaction:
                        conn.commit()
                except sqlite3.Error as exc:
                    elapsed = round((time.perf_counter() - start) * 1000, 1)
                    message = str(exc)
                    if "interrupted" in message.lower():
                        message = (
                            f"Query cancelled after {QUERY_TIMEOUT_SECONDS}s. "
                            "Try adding a LIMIT or a WHERE clause."
                        )
                    error = {
                        "message": message,
                        "code": getattr(exc, "sqlite_errorname", None) or "SQL_ERROR",
                        "statementIndex": index,
                        "statement": statement,
                        "elapsedMs": elapsed,
                    }
                    break
        finally:
            conn.set_progress_handler(None, 0)
            conn.close()

        return {
            "ok": error is None,
            "results": results,
            "error": error,
            "elapsedMs": round((time.perf_counter() - overall_start) * 1000, 1),
            "readOnly": self.read_only,
        }

    # --------------------------------------------------------------- export
    def iter_table_csv(self, name: str) -> Iterable[str]:
        relation = self._find_relation(name)
        conn = self.connect()
        try:
            cursor = conn.execute(f"SELECT * FROM {quote_identifier(relation['name'])}")
            headers = [col[0] for col in cursor.description]
            buffer = io.StringIO()
            writer = csv.writer(buffer, lineterminator="\n")
            writer.writerow(headers)
            yield buffer.getvalue()
            while True:
                rows = cursor.fetchmany(500)
                if not rows:
                    break
                buffer.seek(0)
                buffer.truncate(0)
                for row in rows:
                    writer.writerow(
                        ["" if value is None else value for value in row]
                    )
                yield buffer.getvalue()
        finally:
            conn.close()
