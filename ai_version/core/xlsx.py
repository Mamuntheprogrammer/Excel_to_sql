"""A small, dependency-free XLSX reader.

Only what a data browser needs: sheet names and cell values (including shared
strings, inline strings, booleans and dates).  Formulas use their cached value.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from datetime import datetime, timedelta
from typing import Any, Iterator
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

BUILTIN_DATE_FORMAT_IDS = {
    14, 15, 16, 17, 18, 19, 20, 21, 22,
    27, 30, 36, 45, 46, 47, 50, 57,
}


def _q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def column_index_from_ref(ref: str) -> int:
    index = 0
    for char in ref:
        if char.isalpha():
            index = index * 26 + (ord(char.upper()) - 64)
        else:
            break
    return max(0, index - 1)


def _is_date_format(code: str) -> bool:
    # Remove quoted literals, escapes and bracket sections like [Red]/[$-409].
    cleaned = re.sub(r'"[^"]*"', "", code or "")
    cleaned = re.sub(r"\\.", "", cleaned)
    cleaned = re.sub(r"\[[^\]]*\]", "", cleaned)
    return bool(re.search(r"[ymdhs]", cleaned.lower()))


def _excel_serial_to_text(value: float, date_1904: bool = False) -> str | None:
    if date_1904:
        value += 1462
    if value < 0 or value > 2_958_465:
        return None
    days = int(value)
    fraction = value - days
    base = datetime(1899, 12, 30)
    if days == 0 and fraction:
        seconds = int(round(fraction * 86400))
        hours, rem = divmod(seconds, 3600)
        minutes, seconds = divmod(rem, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    moment = base + timedelta(days=days, seconds=round(fraction * 86400))
    if abs(fraction) < 1e-9:
        return moment.date().isoformat()
    if moment.microsecond == 0:
        return moment.strftime("%Y-%m-%d %H:%M:%S")
    return moment.strftime("%Y-%m-%d %H:%M:%S.%f").rstrip("0").rstrip(".")


class XlsxReader:
    def __init__(self, path: str):
        try:
            self.zf = zipfile.ZipFile(path)
        except (zipfile.BadZipFile, OSError) as exc:
            raise ValueError("This file is not a valid .xlsx workbook.") from exc
        self.shared_strings: list[str] = []
        self.style_is_date: list[bool] = []
        self.date_1904 = False
        self.sheets: list[dict[str, str]] = []
        self._load_workbook()
        self._load_shared_strings()
        self._load_styles()

    def close(self) -> None:
        self.zf.close()

    # ------------------------------------------------------------- metadata
    def _load_workbook(self) -> None:
        try:
            workbook = ET.fromstring(self.zf.read("xl/workbook.xml"))
        except KeyError as exc:
            raise ValueError("The workbook is missing xl/workbook.xml.") from exc

        props = workbook.find(_q(NS_MAIN, "workbookPr"))
        if props is not None:
            self.date_1904 = props.get("date1904") in ("1", "true")

        rel_targets: dict[str, str] = {}
        try:
            rels = ET.fromstring(self.zf.read("xl/_rels/workbook.xml.rels"))
            for rel in rels.findall(_q(NS_PKG_REL, "Relationship")):
                rid = rel.get("Id")
                target = rel.get("Target")
                if rid and target:
                    rel_targets[rid] = self._resolve_target(target)
        except KeyError:
            pass

        sheets_el = workbook.find(_q(NS_MAIN, "sheets"))
        if sheets_el is None:
            raise ValueError("The workbook contains no worksheets.")
        for position, sheet in enumerate(sheets_el.findall(_q(NS_MAIN, "sheet"))):
            name = sheet.get("name") or f"Sheet{position + 1}"
            rid = sheet.get(_q(NS_REL, "id"))
            target = rel_targets.get(rid or "")
            if not target:
                target = f"xl/worksheets/sheet{position + 1}.xml"
            self.sheets.append({"name": name, "path": target})

    @staticmethod
    def _resolve_target(target: str) -> str:
        target = target.replace("\\", "/")
        if target.startswith("/"):
            return target.lstrip("/")
        if target.startswith("xl/"):
            return posixpath.normpath(target)
        return posixpath.normpath(posixpath.join("xl", target))

    def _load_shared_strings(self) -> None:
        try:
            fh = self.zf.open("xl/sharedStrings.xml")
        except KeyError:
            return
        with fh:
            for _event, elem in ET.iterparse(fh, events=("end",)):
                if elem.tag == _q(NS_MAIN, "si"):
                    self.shared_strings.append(self._rich_text(elem))
                    elem.clear()

    @staticmethod
    def _rich_text(elem: ET.Element) -> str:
        return "".join(
            node.text or "" for node in elem.iter(_q(NS_MAIN, "t"))
        )

    def _load_styles(self) -> None:
        try:
            styles = ET.fromstring(self.zf.read("xl/styles.xml"))
        except KeyError:
            return
        custom_formats: dict[int, str] = {}
        numfmts = styles.find(_q(NS_MAIN, "numFmts"))
        if numfmts is not None:
            for fmt in numfmts.findall(_q(NS_MAIN, "numFmt")):
                try:
                    custom_formats[int(fmt.get("numFmtId", "0"))] = fmt.get(
                        "formatCode", ""
                    )
                except ValueError:
                    continue
        cell_xfs = styles.find(_q(NS_MAIN, "cellXfs"))
        if cell_xfs is None:
            return
        for xf in cell_xfs.findall(_q(NS_MAIN, "xf")):
            try:
                fmt_id = int(xf.get("numFmtId", "0"))
            except ValueError:
                fmt_id = 0
            if fmt_id in BUILTIN_DATE_FORMAT_IDS:
                self.style_is_date.append(True)
            elif fmt_id in custom_formats:
                self.style_is_date.append(_is_date_format(custom_formats[fmt_id]))
            else:
                self.style_is_date.append(False)

    # ---------------------------------------------------------------- sheets
    def sheet_names(self) -> list[str]:
        return [sheet["name"] for sheet in self.sheets]

    def _sheet_path(self, name: str) -> str:
        for sheet in self.sheets:
            if sheet["name"] == name:
                return sheet["path"]
        raise ValueError(f"Sheet not found: {name}")

    def iter_rows(self, sheet_name: str) -> Iterator[list[Any]]:
        path = self._sheet_path(sheet_name)
        try:
            fh = self.zf.open(path)
        except KeyError as exc:
            raise ValueError(f"Worksheet data is missing for '{sheet_name}'.") from exc

        with fh:
            context = ET.iterparse(fh, events=("start", "end"))
            root: ET.Element | None = None
            for event, elem in context:
                if event == "start":
                    if root is None:
                        root = elem
                    continue
                if elem.tag == _q(NS_MAIN, "row"):
                    row = self._parse_row(elem)
                    if any(value is not None for value in row):
                        yield row
                    elem.clear()
                    if root is not None:
                        root.clear()
                elif elem.tag == _q(NS_MAIN, "sheetData"):
                    elem.clear()

    def _parse_row(self, row_elem: ET.Element) -> list[Any]:
        values: dict[int, Any] = {}
        next_index = 0
        for cell in row_elem.findall(_q(NS_MAIN, "c")):
            ref = cell.get("r")
            index = column_index_from_ref(ref) if ref else next_index
            next_index = index + 1
            cell_type = cell.get("t")
            style_raw = cell.get("s")

            if cell_type == "inlineStr":
                inline = cell.find(_q(NS_MAIN, "is"))
                value: Any = self._rich_text(inline) if inline is not None else None
            else:
                v_el = cell.find(_q(NS_MAIN, "v"))
                raw = v_el.text if v_el is not None else None
                value = self._convert_cell(raw, cell_type, style_raw)

            if value == "":
                value = None
            if value is not None:
                values[index] = value

        if not values:
            return []
        max_index = max(values)
        return [values.get(i) for i in range(max_index + 1)]

    def _convert_cell(
        self, raw: str | None, cell_type: str | None, style_raw: str | None
    ) -> Any:
        if raw is None:
            return None
        if cell_type == "s":
            try:
                return self.shared_strings[int(raw)]
            except (ValueError, IndexError):
                return ""
        if cell_type in ("str", "e"):
            return raw
        if cell_type == "b":
            return 1 if raw.strip() in ("1", "true", "TRUE") else 0
        if cell_type == "d":
            return raw  # ISO date value

        try:
            number = float(raw)
        except ValueError:
            return raw

        is_date = False
        if style_raw is not None:
            try:
                style_index = int(style_raw)
                is_date = (
                    0 <= style_index < len(self.style_is_date)
                    and self.style_is_date[style_index]
                )
            except ValueError:
                is_date = False

        if is_date:
            converted = _excel_serial_to_text(number, self.date_1904)
            if converted is not None:
                return converted

        if number.is_integer() and abs(number) < 10**15:
            return int(number)
        return number
