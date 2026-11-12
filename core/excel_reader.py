"""
Excel reader / writer — works on ANY open Excel workbook,
no VBA required in the target file.

Strategy:
    - On Mac, use AppleScript (osascript) for detection and workbook/sheet
      listing (same proven approach as core/excel_bridge.py). This avoids
      xlwings's flaky Excel detection inside Streamlit on macOS.
    - For reading cell values and writing audit columns, use xlwings.
      xlwings needs to be installed AND macOS must grant Automation
      permissions for Excel (System Settings → Privacy & Security →
      Automation → Terminal/Python → tick "Microsoft Excel").
"""

from __future__ import annotations

import platform
import subprocess
from typing import Optional

IS_MAC = platform.system() == "Darwin"
IS_WIN = platform.system() == "Windows"


# ─────────────────────────────────────────────────────────────
def _xw():
    """Internal: return the xlwings module if installed, else None."""
    try:
        import xlwings as xw
        return xw
    except ImportError:
        return None


# Public alias so other modules don't need to import a private name
get_xlwings = _xw


# ─────────────────────────────────────────────────────────────
#  AppleScript helpers (Mac) — same pattern as core/excel_bridge.py
# ─────────────────────────────────────────────────────────────
def _osa(script: str, timeout: int = 25) -> tuple[bool, str]:
    try:
        proc = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=timeout,
        )
        if proc.returncode == 0:
            return True, proc.stdout.strip()
        return False, (proc.stderr or proc.stdout).strip()
    except Exception as exc:
        return False, str(exc)


def _mac_excel_running() -> bool:
    ok, out = _osa(
        'tell application "System Events" to '
        '(name of processes) contains "Microsoft Excel"',
        timeout=5,
    )
    return ok and out.lower() == "true"


def _mac_workbook_names() -> list[str]:
    if not _mac_excel_running():
        return []
    script = (
        'tell application "Microsoft Excel"\n'
        '    set bookList to ""\n'
        '    repeat with wb in workbooks\n'
        '        set bookList to bookList & (name of wb) & "|||"\n'
        '    end repeat\n'
        '    return bookList\n'
        'end tell'
    )
    ok, out = _osa(script, timeout=10)
    if not ok:
        return []
    return [n for n in out.split("|||") if n]


def _mac_sheet_names(workbook_name: str) -> list[str]:
    safe = workbook_name.replace('"', '')
    script = (
        'tell application "Microsoft Excel"\n'
        '    set sheetList to ""\n'
        f'    repeat with sh in sheets of workbook "{safe}"\n'
        '        set sheetList to sheetList & (name of sh) & "|||"\n'
        '    end repeat\n'
        '    return sheetList\n'
        'end tell'
    )
    ok, out = _osa(script, timeout=10)
    if not ok:
        return []
    return [n for n in out.split("|||") if n]


# ─────────────────────────────────────────────────────────────
#  Public API: detection / listing
# ─────────────────────────────────────────────────────────────
def xlwings_available() -> bool:
    return _xw() is not None


def excel_running() -> bool:
    """
    Use xlwings on both Mac & Windows. xlwings's `xw.apps` is the
    most reliable cross-platform way once Automation permission is granted.
    Falls back to AppleScript on Mac if xlwings fails.
    """
    xw = _xw()
    if xw is not None:
        try:
            return len(list(xw.apps)) > 0
        except Exception:
            pass
    if IS_MAC:
        return _mac_excel_running()
    return False


def get_open_workbooks() -> list[str]:
    """Names of all open Excel workbooks (xlwings primary, AppleScript fallback)."""
    xw = _xw()
    if xw is not None:
        try:
            names = [b.name for app in xw.apps for b in app.books]
            if names:
                return names
        except Exception:
            pass
    if IS_MAC:
        return _mac_workbook_names()
    return []


def get_sheets(workbook_name: str) -> list[str]:
    """Sheet names for a workbook (xlwings primary, AppleScript fallback)."""
    xw = _xw()
    if xw is not None:
        try:
            for app in xw.apps:
                for book in app.books:
                    if book.name == workbook_name:
                        return [sh.name for sh in book.sheets]
        except Exception:
            pass
    if IS_MAC:
        return _mac_sheet_names(workbook_name)
    return []


# ─────────────────────────────────────────────────────────────
#  xlwings access for read/write — needs Automation permissions
# ─────────────────────────────────────────────────────────────
def _find_book(xw, name: str):
    for app in xw.apps:
        for book in app.books:
            if book.name == name:
                return book
    return None


# ─────────────────────────────────────────────────────────────
#  Read sheet data
# ─────────────────────────────────────────────────────────────
def read_sheet_as_text(
    workbook_name: str,
    sheet_name: str,
    char_budget: int = 30000,
    max_rows: int = 500,
    target_col: Optional[int] = None,
) -> tuple[str, list[int], int]:
    """
    Returns (text, row_map, header_row).

    If `target_col` (1-based Excel column index) is given, only rows whose
    cell in that column has a real value (non-blank, not '-', not '0') are
    included in the output text and the row_map.
    Returns ("", [], 0) on failure.
    """
    xw = _xw()
    if xw is None:
        return "", [], 0
    book = _find_book(xw, workbook_name)
    if book is None:
        return "", [], 0

    try:
        sheet = book.sheets[sheet_name]
        used = sheet.used_range
        if used is None:
            return "", [], 0
        raw = used.value
        if not raw:
            return "", [], 0
    except Exception:
        return "", [], 0

    # Normalise to 2-D list
    if not isinstance(raw[0], (list, tuple)):
        raw = [raw]

    first_excel_row = used.row
    first_excel_col = used.column

    # Detect header row: row with the MOST non-blank cells in the first 15 rows
    # (that's the "deepest" sub-header — better than picking a sparse super-header)
    n_to_scan = min(15, len(raw))
    header_idx = 0
    best_count = 0
    for i in range(n_to_scan):
        non_blank = sum(1 for v in raw[i] if v is not None and str(v).strip())
        if non_blank > best_count:
            best_count = non_blank
            header_idx = i

    header_row_excel = first_excel_row + header_idx

    def cell_str(v):
        if v is None:
            return ""
        s = str(v)
        if s.endswith(".0") and s[:-2].lstrip("-").isdigit():
            s = s[:-2]
        return s.strip()

    # Position of the target column inside each row (0-based within `row`)
    target_idx = None
    if target_col is not None:
        target_idx = target_col - first_excel_col

    def is_blank_target(v: str) -> bool:
        s = (v or "").strip()
        if not s:
            return True
        if s in ("-", "0", "0.0", "0.00"):
            return True
        try:
            return float(s) == 0
        except (ValueError, TypeError):
            return False

    header_parts = [cell_str(v) for v in raw[header_idx] if cell_str(v)]
    text = "HEADER: " + " | ".join(header_parts) + "\n"

    row_map: list[int] = []
    data_row_num = 0

    for i in range(header_idx + 1, len(raw)):
        if data_row_num >= max_rows:
            text += "[...truncated — increase MAX_ROWS if needed...]\n"
            break
        row = raw[i]
        cells = [cell_str(v) for v in row]
        if not any(cells):
            continue

        # Filter on target column if requested
        if target_idx is not None:
            if target_idx < 0 or target_idx >= len(row):
                continue
            tval = cell_str(row[target_idx])
            if is_blank_target(tval):
                continue

        data_row_num += 1
        excel_row = first_excel_row + i
        row_map.append(excel_row)
        text += f"ROW {data_row_num}: " + " | ".join(cells) + "\n"
        if len(text) > char_budget:
            text = text[:char_budget] + "\n[...truncated...]\n"
            break

    return text, row_map, header_row_excel


# ─────────────────────────────────────────────────────────────
#  Column picker — show the user which columns exist + samples
# ─────────────────────────────────────────────────────────────
def _col_letter(n: int) -> str:
    """1-based column index → 'A', 'B', ..., 'Z', 'AA', ..."""
    out = ""
    while n > 0:
        n -= 1
        out = chr(65 + n % 26) + out
        n //= 26
    return out


def get_audit_column_options(workbook_name: str, sheet_name: str) -> tuple[list[dict], int]:
    """
    Inspect a sheet and return a user-friendly list of columns.

    Each option:
        {
            'col_idx':  int,    # 1-based Excel column index
            'header':   str,    # combined super-header > sub-header
            'label':    str,    # 'F — Quantity > This Bill (e.g. 0.088, 0.238)'
        }

    Returns (options, header_row_excel).
    """
    xw = _xw()
    if xw is None:
        return [], 0
    book = _find_book(xw, workbook_name)
    if book is None:
        return [], 0
    try:
        sheet = book.sheets[sheet_name]
        used = sheet.used_range
        if used is None:
            return [], 0
        raw = used.value
        if not raw:
            return [], 0
    except Exception:
        return [], 0

    if not isinstance(raw[0], (list, tuple)):
        raw = [raw]

    first_excel_row = used.row
    first_excel_col = used.column

    # Find the deepest header row (most non-blank cells in first 15 rows)
    n_to_scan = min(15, len(raw))
    header_idx = 0
    best = 0
    for i in range(n_to_scan):
        non_blank = sum(1 for v in raw[i] if v is not None and str(v).strip())
        if non_blank > best:
            best = non_blank
            header_idx = i

    header_row_excel = first_excel_row + header_idx

    def cell_str(v):
        if v is None:
            return ""
        s = str(v)
        if s.endswith(".0") and s[:-2].lstrip("-").isdigit():
            s = s[:-2]
        return s.strip()

    options: list[dict] = []
    header_row = raw[header_idx]
    n_cols = max((len(r) for r in raw if r is not None), default=0)

    for c in range(n_cols):
        # Sub-header (the detected header row)
        sub = cell_str(header_row[c]) if c < len(header_row) else ""

        # Super-header — look 1 and 2 rows above for merged groupings
        super_parts: list[str] = []
        for off in (1, 2):
            if header_idx - off < 0:
                break
            r_above = raw[header_idx - off]
            if c < len(r_above):
                v = cell_str(r_above[c])
                if v and v != sub:
                    super_parts.insert(0, v)

        full_header = " > ".join(super_parts + [sub]) if sub else " > ".join(super_parts)

        # Sample values — first 3 non-blank values in data rows
        samples: list[str] = []
        for i in range(header_idx + 1, min(header_idx + 25, len(raw))):
            r = raw[i]
            if c < len(r):
                v = cell_str(r[c])
                if v and v not in ("-",):
                    samples.append(v)
            if len(samples) >= 3:
                break

        if not full_header and not samples:
            continue

        col_idx = first_excel_col + c
        letter = _col_letter(col_idx)
        sample_text = ", ".join(samples) if samples else "no data"
        # Trim very long sample values
        if len(sample_text) > 60:
            sample_text = sample_text[:57] + "…"

        options.append({
            "col_idx": col_idx,
            "header":  full_header or "(no header)",
            "label":   f"{letter} — {full_header or '(no header)'}   [e.g. {sample_text}]",
        })

    return options, header_row_excel


# ─────────────────────────────────────────────────────────────
#  Write audit column back into Excel
# ─────────────────────────────────────────────────────────────
def write_audit_column(
    workbook_name: str,
    sheet_name: str,
    header_row_excel: int,
    row_map: list[int],
    results: list[tuple[str, bool]],
    column_title: str = "LLM Audit",
) -> dict:
    """
    Write audit results into the first empty column to the right of used data.
    Returns { ok: bool, col: int, error: str }
    """
    xw = _xw()
    if xw is None:
        return {"ok": False, "col": 0, "error": "xlwings not installed."}

    try:
        book = _find_book(xw, workbook_name)
        if book is None:
            return {"ok": False, "col": 0, "error": f"Workbook '{workbook_name}' not found."}
        sheet = book.sheets[sheet_name]
        used = sheet.used_range
        last_col = used.last_cell.column
        audit_col = last_col + 1

        hdr = sheet.cells(header_row_excel, audit_col)
        hdr.value = column_title
        hdr.font.bold = True
        hdr.color = (255, 242, 204)

        GREEN, RED = (198, 239, 206), (255, 199, 206)
        GREEN_FG, RED_FG = (0, 97, 0), (156, 0, 6)

        for i, excel_row in enumerate(row_map):
            if i >= len(results):
                break
            text, is_ok = results[i]
            cell = sheet.cells(excel_row, audit_col)
            cell.value = text
            cell.color = GREEN if is_ok else RED
            cell.font.color = GREEN_FG if is_ok else RED_FG

        # Width
        try:
            sheet.range(sheet.cells(header_row_excel, audit_col)).column_width = 55
        except Exception:
            pass

        return {"ok": True, "col": audit_col, "error": ""}
    except Exception as exc:
        return {"ok": False, "col": 0, "error": str(exc)}


# ─────────────────────────────────────────────────────────────
#  Permission diagnostic
# ─────────────────────────────────────────────────────────────
def xlwings_can_read() -> tuple[bool, str]:
    """
    Try to actually connect to Excel via xlwings. Returns (ok, hint).
    """
    xw = _xw()
    if xw is None:
        return False, "xlwings not installed. Run: pip install xlwings"
    try:
        apps = list(xw.apps)
        if not apps:
            return False, ("xlwings can't see Excel. On macOS, grant Automation permission: "
                           "System Settings → Privacy & Security → Automation → "
                           "Terminal (or your shell) → enable 'Microsoft Excel'. "
                           "Then restart Streamlit.")
        return True, ""
    except Exception as exc:
        return False, f"xlwings error: {exc}"
