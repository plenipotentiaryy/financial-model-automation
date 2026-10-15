"""Structured import with real cell/row/JSON locators and no formula execution."""

import csv
import json
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .mapping import alias_lookup, map_headers
from .schemas import Issue, ModelError, RawRow, Source


def read_file(path: Path, aliases: dict, sheet: str | None = None) -> tuple[list[RawRow], list[Issue]]:
    path = path.resolve()
    lookup = alias_lookup(aliases)
    rows, issues = [], []

    def append(headers, values, row_number, sheet_name=None, json_index=None):
        mapped, unknown = map_headers(headers, lookup)
        if unknown and not any(i.code == "UNKNOWN_COLUMNS" for i in issues):
            issues.append(Issue("UNKNOWN_COLUMNS", f"Ignored undocumented columns: {', '.join(unknown)}", source=str(path)))
        row_values, sources = {}, {}
        for index, canonical in mapped.items():
            value = values[index] if index < len(values) else None
            header = headers[index]
            if json_index is not None:
                locator = f"$[{json_index}][{json.dumps(header, ensure_ascii=False)}]"
            elif sheet_name is not None:
                locator = f"{sheet_name}!{get_column_letter(index + 1)}{row_number}"
            else:
                locator = f"row {row_number}, column {header}"
            row_values[canonical] = value
            sources[canonical] = Source(str(path), sheet_name, row_number, header, locator, value)
        rows.append(RawRow(row_values, sources))

    try:
        if path.suffix.lower() == ".csv":
            with path.open(newline="", encoding="utf-8-sig") as stream:
                reader = csv.reader(stream, strict=True)
                headers = next(reader, None)
                if not headers:
                    raise ModelError("Empty CSV: missing header")
                map_headers(headers, lookup)
                for number, values in enumerate(reader, 2):
                    if not values or all(v == "" for v in values):
                        continue
                    if len(values) != len(headers):
                        raise ModelError(f"CSV row {number}: {len(values)} fields, expected {len(headers)}")
                    append(headers, values, number)
        elif path.suffix.lower() == ".json":
            def invalid_constant(value):
                raise ModelError(f"Nonstandard JSON numeric constant: {value}")

            def object_pairs(pairs):
                obj = {}
                for key, value in pairs:
                    if key in obj:
                        raise ModelError(f"Duplicate JSON field: {key}")
                    obj[key] = value
                return obj

            # Preserve JSON decimal lexemes; don't round monetary amounts via float.
            data = json.loads(path.read_text(encoding="utf-8"), parse_float=str, parse_constant=invalid_constant, object_pairs_hook=object_pairs)
            if not isinstance(data, list) or not data:
                raise ModelError("JSON must be a nonempty array of row objects")
            for index, obj in enumerate(data):
                if not isinstance(obj, dict):
                    raise ModelError(f"JSON $[{index}] must be an object")
                append(list(obj), list(obj.values()), index + 1, json_index=index)
        elif path.suffix.lower() == ".xlsx":
            wb = load_workbook(path, read_only=True, data_only=False)
            try:
                nonempty = [ws for ws in wb if any(cell.value is not None for row in ws for cell in row)]
                if sheet is not None:
                    if sheet not in wb.sheetnames:
                        raise ModelError(f"Unknown worksheet: {sheet}")
                    ws = wb[sheet]
                elif len(nonempty) != 1:
                    raise ModelError("XLSX requires --sheet when there is not exactly one nonempty worksheet")
                else:
                    ws = nonempty[0]
                iterator = ws.iter_rows()
                header_cells = list(next(iterator, []))
                while header_cells and header_cells[-1].value is None:
                    header_cells.pop()
                headers = [c.value for c in header_cells]
                map_headers(headers, lookup)
                for number, cells in enumerate(iterator, 2):
                    if all(c.value is None for c in cells):
                        continue
                    if any(c.value is not None for c in cells[len(headers):]):
                        raise ModelError(f"Unexpected data beyond headers at row {number}")
                    values = [c.value for c in cells[:len(headers)]]
                    # Formula rejection is attributed to the affected company below.
                    append(headers, values, number, ws.title)
            finally:
                wb.close()
        else:
            raise ModelError(f"Unsupported input type {path.suffix}; use CSV, JSON or XLSX")
    except (OSError, ValueError, csv.Error, KeyError, TypeError, IndexError, SyntaxError) as exc:
        raise ModelError(f"Cannot import {path.name}: {exc}") from exc
    if not rows:
        raise ModelError(f"No data rows in {path.name}")
    return rows, issues
