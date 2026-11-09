# converter.py
"""
Excel (.xlsx / .xlsm) -> MySQL INSERT statements.

Features
--------
* Streaming + batching (handles 100k+ rows, flat memory use)
* DROP TABLE / CREATE TABLE / BATCH INSERT options
* Table-name validation + auto-suggestion
* Live preview generator for the GUI's SQL panel
"""

import os
import re
import math
import decimal
import datetime
import openpyxl
from openpyxl.utils import get_column_letter

MAX_IDENT = 64          # MySQL identifier length limit
MAX_BATCH = 5000        # upper bound for rows per INSERT


class ConversionCancelled(Exception):
    """Raised when the user cancels a running conversion."""


# ---------------------------------------------------------------------------
# Value escaping
# ---------------------------------------------------------------------------
def quote_ident(name):
    """Backtick-quote a MySQL identifier (doubling any backticks inside)."""
    return "`" + str(name).replace("`", "``") + "`"


def escape_value(val):
    """
    Format a cell value as a MySQL literal, changing nothing:
    text is kept byte-for-byte (no trimming), numbers keep their exact digits.
    """
    if val is None:
        return "NULL"
    if isinstance(val, bool):
        return "'TRUE'" if val else "'FALSE'"      # what Excel shows
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return "NULL"
        if val.is_integer() and abs(val) < 1e15:
            return str(int(val))                    # 5.0 -> 5 (as in Excel)
        s = repr(val)
        if "e" in s or "E" in s:                    # never emit 1e-05 style
            s = format(decimal.Decimal(s), "f")
        return s
    if isinstance(val, decimal.Decimal):
        return format(val, "f") if val.is_finite() else "NULL"
    if isinstance(val, datetime.datetime):
        if val.time() == datetime.time(0, 0):
            return f"'{val.date().isoformat()}'"   # pure date cell
        return f"'{val.isoformat(sep=' ')}'"
    if isinstance(val, datetime.date):
        return f"'{val.isoformat()}'"
    if isinstance(val, datetime.time):
        return f"'{val.isoformat()}'"
    if isinstance(val, datetime.timedelta):
        secs = int(val.total_seconds())
        sign = "-" if secs < 0 else ""
        secs = abs(secs)
        return f"'{sign}{secs // 3600:02d}:{(secs % 3600) // 60:02d}:{secs % 60:02d}'"

    s = str(val)                                    # NO .strip()
    s = (s.replace("\\", "\\\\")
          .replace("'", "\\'")
          .replace("\x00", "\\0")
          .replace("\x1a", "\\Z")
          .replace("\r", "\\r")
          .replace("\n", "\\n"))
    return f"'{s}'"


# ---------------------------------------------------------------------------
# Excel metadata helpers
# ---------------------------------------------------------------------------
def _sheet_total(ws):
    """Data rows (excluding header). max_row can be None in read-only mode."""
    mr = ws.max_row
    return max(mr - 1, 0) if isinstance(mr, int) else 0


def get_sheet_names(file_path):
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def count_rows(file_path, sheet_name):
    """Return number of data rows (excluding header); 0 if unknown."""
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' not found")
        return _sheet_total(wb[sheet_name])
    finally:
        wb.close()


def parse_header(raw_header):
    """
    Return (col_names, col_indices, notes).

    Every column from the first one up to the last non-blank header is kept
    (blank headers inside that range get the name column_<Excel letter>).
    Names are only changed where MySQL forces it (duplicates ignoring case,
    > 64 chars, trailing spaces); each change is listed in `notes`.
    """
    last = -1
    for i, c in enumerate(raw_header):
        if c is not None and str(c).strip() != "":
            last = i

    names, indices, notes, used = [], [], [], set()
    for i in range(last + 1):
        c = raw_header[i]
        if c is None or str(c).strip() == "":
            orig, base = None, f"column_{get_column_letter(i + 1).lower()}"
        else:
            orig = str(c)
            base = orig.rstrip()          # MySQL forbids trailing spaces
        base = base[:MAX_IDENT]
        cand, k = base, 1
        while cand.lower() in used:
            k += 1
            suffix = f"_{k}"
            cand = base[:MAX_IDENT - len(suffix)] + suffix
        used.add(cand.lower())
        names.append(cand)
        indices.append(i)
        if orig != cand:
            shown = "(blank)" if orig is None else " ".join(orig.split())
            notes.append(f"-- NOTE: Excel column {get_column_letter(i + 1)} "
                         f"header '{shown}' -> `{cand}`")
    return names, indices, notes


def iter_data_rows(rows, col_indices):
    """
    Yield each data row as a list of values for col_indices.
    Blank rows between data are KEPT (as all-NULL rows); only trailing
    blank rows at the end of the sheet are dropped.
    """
    pending = 0
    width = len(col_indices)
    for row in rows:
        vals = [row[i] if i < len(row) else None for i in col_indices]
        if all(v is None for v in vals):
            pending += 1
            continue
        for _ in range(pending):
            yield [None] * width
        pending = 0
        yield vals


def load_preview(file_path, sheet_name, max_rows=100):
    """
    Load first `max_rows` rows + header for the Excel preview widget.
    Returns (headers: list[str], rows: list[list[Any]], total_rows: int).
    """
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        total = _sheet_total(ws)

        it = ws.iter_rows(values_only=True)
        try:
            raw_header = next(it)
        except StopIteration:
            return [], [], 0

        headers, idx, _notes = parse_header(raw_header)

        rows = []
        for _, row in zip(range(max_rows), it):
            rows.append([row[i] if i < len(row) else None for i in idx])
        return headers, rows, total
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Table-name validation & suggestion
# ---------------------------------------------------------------------------
_VALID_TABLE_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")

_RESERVED = {
    "select", "insert", "update", "delete", "from", "where", "table",
    "database", "order", "group", "by", "join", "user", "index",
    "key", "primary", "foreign", "drop", "create", "alter", "values",
}


def validate_table_name(name: str):
    """
    Returns (is_valid: bool, message: str).
    Rules: letters/digits/underscore, must not start with a digit, <= 64 chars.
    """
    if not name:
        return False, "Table name is required."
    if not _VALID_TABLE_RE.match(name):
        return False, ("Only letters, digits, and underscores are allowed, "
                       "and the name cannot start with a digit.")
    if len(name) > MAX_IDENT:
        return False, "MySQL table names must be <= 64 characters."
    if name.lower() in _RESERVED:
        return True, (f"'{name}' is a MySQL reserved word - it will be quoted, "
                      f"but consider renaming.")
    return True, ""


def suggest_table_name(file_path: str, sheet_name: str) -> str:
    """
    Derive a safe table name from file + sheet names.
    'Sales Data 2024.xlsx' + 'Q1' -> 'sales_data_2024_q1'
    """
    base = ""
    if file_path:
        base = os.path.splitext(os.path.basename(file_path))[0]
    parts = [base, sheet_name] if sheet_name else [base]
    raw = "_".join(p for p in parts if p)
    clean = re.sub(r"[^A-Za-z0-9_]+", "_", raw).strip("_").lower()
    if not clean:
        clean = "my_table"
    if clean[0].isdigit():
        clean = "t_" + clean
    return clean[:MAX_IDENT]


# ---------------------------------------------------------------------------
# SQL header (DROP / CREATE)
# ---------------------------------------------------------------------------
def generate_sql_header(table_name, col_names,
                        drop_table=False, create_table=False):
    """Return the DROP/CREATE preamble (or empty string)."""
    parts = []
    t = quote_ident(table_name)
    if drop_table:
        parts.append(f"DROP TABLE IF EXISTS {t};")

    if create_table:
        # LONGTEXT: holds any cell (Excel allows 32,767 chars) with no cut-off.
        # utf8mb4: keeps Sinhala / Tamil / emoji etc. exactly.
        cols_def = [f"{quote_ident(n)} LONGTEXT" for n in col_names]
        parts.append(
            f"CREATE TABLE {t} (\n  "
            + ",\n  ".join(cols_def)
            + "\n) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 "
              "COLLATE=utf8mb4_unicode_ci;"
        )
    return "\n".join(parts)


def _clamp_batch(batch_size):
    try:
        b = int(batch_size)
    except (TypeError, ValueError):
        b = 500
    return max(1, min(b, MAX_BATCH))


# ---------------------------------------------------------------------------
# Streaming generator (used for full conversion / file write)
# ---------------------------------------------------------------------------
def excel_to_sql_stream(file_path, sheet_name, table_name,
                        batch_size=500,
                        drop_table=False, create_table=False,
                        batch_insert=True,
                        progress_cb=None,
                        cancel_event=None):
    """Yield (sql_chunk, rows_processed_so_far)."""
    wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        rows = ws.iter_rows(values_only=True)

        try:
            header = next(rows)
        except StopIteration:
            raise ValueError("Worksheet is empty")

        col_names, col_indices, notes = parse_header(header)
        if not col_names:
            raise ValueError("No valid header columns detected")

        t = quote_ident(table_name)
        col_str = ", ".join(quote_ident(c) for c in col_names)

        yield "SET NAMES utf8mb4;\nSET FOREIGN_KEY_CHECKS=0;\n\n", 0
        if notes:
            yield "\n".join(notes) + "\n\n", 0
        hdr = generate_sql_header(table_name, col_names, drop_table, create_table)
        if hdr:
            yield hdr + "\n\n", 0

        batch_size = _clamp_batch(batch_size) if batch_insert else 1

        batch = []
        processed = 0

        for vals_raw in iter_data_rows(rows, col_indices):
            if cancel_event is not None and cancel_event.is_set():
                raise ConversionCancelled("Cancelled by user")

            vals = [escape_value(v) for v in vals_raw]
            batch.append(f"({', '.join(vals)})")

            if len(batch) >= batch_size:
                processed += len(batch)
                stmt = (f"INSERT INTO {t} ({col_str}) VALUES\n"
                        + ",\n".join(batch) + ";\n\n")
                batch.clear()
                yield stmt, processed
                if progress_cb:
                    progress_cb(processed)

        if batch:
            processed += len(batch)
            stmt = (f"INSERT INTO {t} ({col_str}) VALUES\n"
                    + ",\n".join(batch) + ";\n")
            yield stmt, processed
            if progress_cb:
                progress_cb(processed)

        yield "\nSET FOREIGN_KEY_CHECKS=1;\n", processed

    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Write to .sql file
# ---------------------------------------------------------------------------
def excel_to_sql_file(file_path, sheet_name, table_name, output_path,
                      batch_size=500,
                      drop_table=False, create_table=False,
                      batch_insert=True,
                      progress_cb=None,
                      cancel_event=None):
    """Convert and write directly to output_path. Returns row count."""
    total = 0
    ok = False
    try:
        with open(output_path, "w", encoding="utf-8", newline="\n") as f:
            for chunk, processed in excel_to_sql_stream(
                    file_path, sheet_name, table_name,
                    batch_size=batch_size,
                    drop_table=drop_table,
                    create_table=create_table,
                    batch_insert=batch_insert,
                    progress_cb=progress_cb,
                    cancel_event=cancel_event):
                f.write(chunk)
                total = processed
        ok = True
    finally:
        if not ok:                      # cancelled or failed: no half-written file
            try:
                os.remove(output_path)
            except OSError:
                pass
    return total


# ---------------------------------------------------------------------------
# Live preview generator (feeds the GUI's SQL panel)
# ---------------------------------------------------------------------------
def generate_live_sql_preview(file_path, sheet_name, table_name,
                              batch_size=500,
                              drop_table=False, create_table=False,
                              batch_insert=True,
                              preview_rows=100):
    """
    Build a short SQL preview (max `preview_rows` data rows) for the GUI.
    """
    if not file_path or not sheet_name or not table_name:
        return "-- Load an Excel file and enter a table name to see SQL here."

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return f"-- Worksheet '{sheet_name}' not found."

        ws = wb[sheet_name]
        rows = ws.iter_rows(values_only=True)

        try:
            header = next(rows)
        except StopIteration:
            return "-- Worksheet is empty."

        col_names, col_indices, notes = parse_header(header)
        if not col_names:
            return "-- No valid header columns found."

        t = quote_ident(table_name)
        col_str = ", ".join(quote_ident(c) for c in col_names)

        parts = []
        hdr = generate_sql_header(table_name, col_names, drop_table, create_table)
        if hdr:
            parts.append(hdr)

        values = []
        limit = preview_rows if batch_insert else min(preview_rows, 3)

        if notes:
            parts.append("\n".join(notes))

        for vals_raw in iter_data_rows(rows, col_indices):
            if len(values) >= limit:
                break
            vals = [escape_value(v) for v in vals_raw]
            values.append(f"({', '.join(vals)})")

        if values:
            if batch_insert:
                bs = _clamp_batch(batch_size)
                for k in range(0, len(values), bs):
                    parts.append(
                        f"INSERT INTO {t} ({col_str}) VALUES\n"
                        + ",\n".join(values[k:k + bs]) + ";"
                    )
            else:
                for v in values:
                    parts.append(f"INSERT INTO {t} ({col_str}) VALUES {v};")

        parts.append(f"-- Preview of first {len(values)} rows. "
                     f"Click 'Convert & Preview' or 'Save .sql file' "
                     f"for the full output.")
        return "\n\n".join(parts)

    finally:
        wb.close()