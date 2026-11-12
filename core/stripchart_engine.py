"""
StripChart engine — pure-Python port of StripChart.bas.

No Excel / xlwings / AppleScript dependency. Reads an .xlsx with openpyxl,
generates a strip-chart sheet (with overlap & missing-lower-layer audit),
and returns the resulting workbook as bytes for download.

Public entry point:
    generate_stripchart(input_xlsx_bytes, cfg) -> output_xlsx_bytes
"""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Constants — match StripChart.bas
# ---------------------------------------------------------------------------
COLS_PER_BLOCK = 50
C_DATA_START = 3   # first chainage column = C

# Hex (no alpha) colors for openpyxl PatternFill / Font.
COLOR_RED = "FF0000"
COLOR_PURPLE = "800080"
COLOR_OK_GREEN = "008000"
COLOR_HEADER_GREY = "E6E6E6"

PALETTE = [
    "D4A574",  # tan
    "A8C77A",  # light green
    "7AA8C7",  # sky blue
    "C77AA8",  # pink
    "FFC107",  # gold
    "7AC7C7",  # light teal
    "C77A7A",  # salmon
]

FROM_KEYWORDS = {"from", "from chainage", "ch from", "from ch"}
TO_KEYWORDS = {"to", "to chainage", "ch to", "to ch"}
SIDE_KEYWORDS = {"side", "lhs/rhs", "rhs/lhs"}

RHS_SIDES = {"RHS", "BHS", "BS", "BOTH"}
LHS_SIDES = {"LHS", "BHS", "BS", "BOTH"}


# ---------------------------------------------------------------------------
# Configuration dataclass
# ---------------------------------------------------------------------------
@dataclass
class Slot:
    """One layer slot. `sheet_name` is the data sheet inside the source workbook."""
    sheet_name: str


@dataclass
class StripChartConfig:
    from_m: int                      # chart start chainage in metres
    to_m: int                        # chart end chainage in metres
    interval_m: int                  # metres per column
    assume_km: bool                  # numeric From/To values are in km
    slots: list[Slot]                # in selection order; first = bottom layer
    strip_sheet_name: str = "StripChart"
    log: list[str] = field(default_factory=list)  # populated during run


# ---------------------------------------------------------------------------
# Chainage parsing
# ---------------------------------------------------------------------------
def parse_chainage(value, assume_km: bool = True) -> Optional[int]:
    """
    Parse a chainage in any of these forms:
        "152+780"  -> 152780
        "152.78"   -> 152780  (decimal km)
        152780     -> 152780
        152.78     -> 152780  (numeric, smart unit detection)
    Returns None on failure.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        n = float(value)
        # Smart unit detection: small numbers are km, large are already metres.
        if n < 1000 and assume_km:
            return int(round(n * 1000))
        return int(round(n))
    s = str(value).strip()
    if not s:
        return None

    if "+" in s:
        try:
            km_str, m_str = s.split("+", 1)
            return int(km_str.strip()) * 1000 + int(m_str.strip())
        except ValueError:
            return None

    if "." in s:
        try:
            km_str, frac_str = s.split(".", 1)
            frac_str = (frac_str + "000")[:3]
            return int(km_str.strip()) * 1000 + int(frac_str)
        except ValueError:
            return None

    try:
        n = float(s)
        if n < 1000 and assume_km:
            return int(round(n * 1000))
        return int(round(n))
    except ValueError:
        return None


def format_chainage(metres: int) -> str:
    return f"{metres // 1000}+{metres % 1000:03d}"


# ---------------------------------------------------------------------------
# Header / data-row auto-detect (replaces StripChart.bas FindHeaderColumns)
# ---------------------------------------------------------------------------
def find_header_columns(ws, scan_rows: int = 30) -> tuple[int, int, int]:
    """Returns (from_col, to_col, side_col) — 1-based indices. Defaults E/F/G."""
    from_col, to_col, side_col = 5, 6, 7
    found_from = found_to = found_side = False

    for r in range(1, scan_rows + 1):
        for c in range(1, 27):
            try:
                v = ws.cell(row=r, column=c).value
            except Exception:
                continue
            if v is None:
                continue
            t = str(v).strip().lower()
            if not t:
                continue
            if not found_from and t in FROM_KEYWORDS:
                from_col = c
                found_from = True
            if not found_to and t in TO_KEYWORDS:
                to_col = c
                found_to = True
            if not found_side and t in SIDE_KEYWORDS:
                side_col = c
                found_side = True
        if found_from and found_to and found_side:
            break
    return from_col, to_col, side_col


def find_data_start_row(ws, scan_rows: int = 30) -> Optional[int]:
    """First row whose column A holds a numeric value."""
    for r in range(1, scan_rows + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        if isinstance(v, (int, float)):
            return r
        try:
            float(str(v).strip())
            return r
        except ValueError:
            continue
    return None


def find_last_data_row(ws, start_row: int) -> int:
    last = ws.max_row
    while last >= start_row:
        if ws.cell(row=last, column=1).value not in (None, ""):
            return last
        last -= 1
    return start_row - 1


# ---------------------------------------------------------------------------
# Layer color resolution (tab color -> A1 fill -> palette)
# ---------------------------------------------------------------------------
def _hex_from_rgb_attr(rgb) -> Optional[str]:
    """openpyxl can return an RGB attr or string. Normalise to a 6-char hex."""
    if rgb is None:
        return None
    s = rgb if isinstance(rgb, str) else getattr(rgb, "rgb", None) or getattr(rgb, "value", None)
    if not s:
        return None
    s = str(s).strip().upper()
    if len(s) == 8:        # ARGB -> drop alpha
        s = s[2:]
    if len(s) != 6:
        return None
    if s in ("FFFFFF", "000000"):
        return None
    return s


def layer_color_for(ws, idx: int) -> str:
    # 1) sheet tab color
    try:
        tab = ws.sheet_properties.tabColor
        if tab is not None:
            hex_ = _hex_from_rgb_attr(tab.rgb if hasattr(tab, "rgb") else tab)
            if hex_:
                return hex_
    except Exception:
        pass
    # 2) cell A1 fill color
    try:
        fill = ws.cell(row=1, column=1).fill
        fg = fill.fgColor if fill is not None else None
        if fg is not None:
            hex_ = _hex_from_rgb_attr(getattr(fg, "rgb", None))
            if hex_:
                return hex_
    except Exception:
        pass
    # 3) palette by index (1-based)
    return PALETTE[(idx - 1) % len(PALETTE)]


# ---------------------------------------------------------------------------
# Block layout (header / RHS rows / axis / LHS rows + borders)
# ---------------------------------------------------------------------------
def _solid_fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


def build_one_block(ws_strip, block_top: int, start_m: int, interval_m: int,
                    last_col: int, layer_names: list[str], layer_colors: list[str],
                    n_layers: int) -> None:
    hdr = block_top
    rhs_top = block_top + 1
    rhs_bot = block_top + n_layers
    axis = block_top + n_layers + 1
    lhs_top = block_top + n_layers + 2
    lhs_bot = block_top + 2 * n_layers + 1

    # Row heights
    ws_strip.row_dimensions[hdr].height = 22
    for r in range(rhs_top, lhs_bot + 1):
        ws_strip.row_dimensions[r].height = 64 if r == axis else 18

    # Header row: SIDE | LAYER | merged CHAINAGE
    ws_strip.cell(row=hdr, column=1, value="SIDE")
    ws_strip.cell(row=hdr, column=2, value="LAYER")
    ws_strip.merge_cells(start_row=hdr, start_column=C_DATA_START,
                         end_row=hdr, end_column=last_col)
    chain_cell = ws_strip.cell(row=hdr, column=C_DATA_START, value="CHAINAGE")
    chain_cell.alignment = Alignment(horizontal="center", vertical="center")
    chain_cell.font = Font(bold=True, size=14)
    for c in (1, 2):
        cell = ws_strip.cell(row=hdr, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # SIDE labels (merged vertically), with rotated text
    ws_strip.merge_cells(start_row=rhs_top, start_column=1, end_row=rhs_bot, end_column=1)
    rhs_lbl = ws_strip.cell(row=rhs_top, column=1, value="RHS")
    rhs_lbl.font = Font(bold=True)
    rhs_lbl.alignment = Alignment(horizontal="center", vertical="center", textRotation=90)

    ws_strip.merge_cells(start_row=lhs_top, start_column=1, end_row=lhs_bot, end_column=1)
    lhs_lbl = ws_strip.cell(row=lhs_top, column=1, value="LHS")
    lhs_lbl.font = Font(bold=True)
    lhs_lbl.alignment = Alignment(horizontal="center", vertical="center", textRotation=90)

    # Axis label (rotated chainage range) in column A on the axis row
    end_m = start_m + COLS_PER_BLOCK * interval_m
    axis_lbl_cell = ws_strip.cell(row=axis, column=1,
                                   value=f"{format_chainage(start_m)} to {format_chainage(end_m)}")
    axis_lbl_cell.font = Font(bold=True, size=9)
    axis_lbl_cell.alignment = Alignment(horizontal="center", vertical="center", textRotation=90)

    # LAYER column: order = first-selected near axis (RHS bottom / LHS top)
    for i in range(1, n_layers + 1):
        # RHS row i (top to bottom, i=1 at top) shows layerNames(N - i + 1)
        cell = ws_strip.cell(row=rhs_top + i - 1, column=2,
                              value=layer_names[n_layers - i])
        cell.fill = _solid_fill(layer_colors[n_layers - i])
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

        # LHS row i (top to bottom, i=1 at top) shows layerNames(i - 1)
        cell = ws_strip.cell(row=lhs_top + i - 1, column=2, value=layer_names[i - 1])
        cell.fill = _solid_fill(layer_colors[i - 1])
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Axis row chainage labels (rotated)
    for col_off in range(COLS_PER_BLOCK):
        m = start_m + col_off * interval_m
        cell = ws_strip.cell(row=axis, column=C_DATA_START + col_off,
                              value=format_chainage(m))
        cell.font = Font(size=8)
        cell.alignment = Alignment(horizontal="center", vertical="center", textRotation=90)

    # Light grey thin borders across the whole block
    thin = Side(style="thin", color="AAAAAA")
    bord = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(hdr, lhs_bot + 1):
        for c in range(1, last_col + 1):
            ws_strip.cell(row=r, column=c).border = bord


# ---------------------------------------------------------------------------
# Audit helpers — block-bounded range detection
# ---------------------------------------------------------------------------
def ranges_in_block(arr: list[list[bool]], layer_idx: int,
                     start_col: int, end_col: int, total_cols: int,
                     from_m: int, interval_m: int) -> str:
    if total_cols <= 0:
        return ""
    if end_col >= total_cols:
        end_col = total_cols - 1

    out_parts: list[str] = []
    in_range = False
    start_c = start_col
    for c in range(start_col, end_col + 1):
        if arr[layer_idx][c]:
            if not in_range:
                start_c = c
                in_range = True
        else:
            if in_range:
                out_parts.append(
                    f"Km {format_chainage(from_m + start_c * interval_m)}"
                    f" to Km {format_chainage(from_m + c * interval_m)}"
                )
                in_range = False
    if in_range:
        out_parts.append(
            f"Km {format_chainage(from_m + start_c * interval_m)}"
            f" to Km {format_chainage(from_m + (end_col + 1) * interval_m)}"
        )
    return "; ".join(out_parts)


def missing_lower_in_block(painted: list[list[bool]], upper_idx: int,
                            start_col: int, end_col: int, total_cols: int,
                            from_m: int, interval_m: int) -> str:
    if upper_idx < 2 or total_cols <= 0:
        return ""
    if end_col >= total_cols:
        end_col = total_cols - 1

    out_parts: list[str] = []
    in_range = False
    start_c = start_col
    for c in range(start_col, end_col + 1):
        is_issue = painted[upper_idx][c] and not painted[upper_idx - 1][c]
        if is_issue:
            if not in_range:
                start_c = c
                in_range = True
        else:
            if in_range:
                out_parts.append(
                    f"Km {format_chainage(from_m + start_c * interval_m)}"
                    f" to Km {format_chainage(from_m + c * interval_m)}"
                )
                in_range = False
    if in_range:
        out_parts.append(
            f"Km {format_chainage(from_m + start_c * interval_m)}"
            f" to Km {format_chainage(from_m + (end_col + 1) * interval_m)}"
        )
    return "; ".join(out_parts)


def has_any_paint(arr: list[list[bool]], layer_idx: int,
                   start_col: int, end_col: int, total_cols: int) -> bool:
    if end_col >= total_cols:
        end_col = total_cols - 1
    return any(arr[layer_idx][c] for c in range(start_col, end_col + 1))


def write_audit(ws_strip, row: int, col: int, msg: str, is_issue: bool) -> None:
    cell = ws_strip.cell(row=row, column=col, value=msg)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.font = Font(
        bold=True,
        color=COLOR_RED if is_issue else COLOR_OK_GREEN,
    )


def paint_missing_in_block(ws_strip, painted: list[list[bool]], upper_idx: int,
                            n_layers: int, block_rows: int, is_rhs: bool,
                            start_col: int, end_col: int, total_cols: int,
                            color_hex: str) -> None:
    if upper_idx < 2:
        return
    if end_col >= total_cols:
        end_col = total_cols - 1
    for c in range(start_col, end_col + 1):
        if painted[upper_idx][c] and not painted[upper_idx - 1][c]:
            block_no = c // COLS_PER_BLOCK
            col_in_block = c % COLS_PER_BLOCK
            block_top = block_no * block_rows + 1
            cell_col = C_DATA_START + col_in_block
            cell_row = (block_top + (n_layers - upper_idx + 1) if is_rhs
                        else block_top + n_layers + 1 + upper_idx)
            ws_strip.cell(row=cell_row, column=cell_col).fill = _solid_fill(color_hex)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def list_data_sheets(input_xlsx_bytes: bytes) -> list[str]:
    """Return non-hidden, non-internal sheet names from the workbook."""
    wb = openpyxl.load_workbook(BytesIO(input_xlsx_bytes), read_only=True, data_only=True)
    return [s for s in wb.sheetnames
            if not s.startswith("_") and "-Strip" not in s and s != "StripChart"]


# ---------------------------------------------------------------------------
# HTML preview renderer — converts a sheet to a styled HTML <table>
# Honours: merged ranges, cell fills, font color/bold/size, alignment,
# vertical-text rotation, column widths, row heights.
# ---------------------------------------------------------------------------
def _hex_from_color_obj(c) -> Optional[str]:
    """Best-effort extraction of an RGB hex string from an openpyxl color."""
    if c is None:
        return None
    # Skip indexed / theme colors — only handle explicit rgb values
    if getattr(c, "type", None) not in (None, "rgb"):
        return None
    rgb = getattr(c, "rgb", None)
    if rgb is None or not isinstance(rgb, str):
        return None
    s = rgb.upper()
    if s == "00000000":
        return None
    # openpyxl stores ARGB; the leading byte is padding, not transparency.
    if len(s) == 8:
        s = s[2:]
    if len(s) != 6:
        return None
    return s


def _escape_html(s: str) -> str:
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;"))


def render_sheet_to_html(xlsx_bytes: bytes, sheet_name: str,
                          max_rows: Optional[int] = None) -> str:
    """
    Build a self-contained HTML <table> that visually approximates the
    given sheet. Returns just the table markup (no <html>/<body>).
    """
    wb = openpyxl.load_workbook(BytesIO(xlsx_bytes), data_only=True)
    if sheet_name not in wb.sheetnames:
        return f"<p>Sheet '{_escape_html(sheet_name)}' not found.</p>"
    ws = wb[sheet_name]

    # Map merged ranges: top-left → (rowspan, colspan); other cells → "skip"
    merged_map: dict[tuple[int, int], object] = {}
    for mr in ws.merged_cells.ranges:
        min_r, min_c = mr.min_row, mr.min_col
        max_r, max_c = mr.max_row, mr.max_col
        merged_map[(min_r, min_c)] = (max_r - min_r + 1, max_c - min_c + 1)
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                if (r, c) != (min_r, min_c):
                    merged_map[(r, c)] = "skip"

    last_row = ws.max_row
    last_col = ws.max_column
    if max_rows is not None:
        last_row = min(last_row, max_rows)

    parts: list[str] = []
    parts.append('<div style="overflow:auto; border:1px solid #d0d4dc; '
                 'border-radius:8px; max-height:680px;">')
    parts.append(
        '<table style="border-collapse:collapse; table-layout:fixed; '
        'font-family:-apple-system,Segoe UI,Roboto,sans-serif; font-size:11px;">'
    )

    # Column widths via <colgroup> (Excel char-units ≈ 7px)
    parts.append("<colgroup>")
    for c in range(1, last_col + 1):
        letter = get_column_letter(c)
        cw = ws.column_dimensions[letter].width
        if cw is None:
            cw = 8
        px = max(int(cw * 7), 12)
        parts.append(f'<col style="width:{px}px">')
    parts.append("</colgroup>")

    for r in range(1, last_row + 1):
        rh = ws.row_dimensions[r].height
        if rh is None:
            rh = 15
        # Excel row height is in points; ~1.33 px per point
        h_px = max(int(rh * 1.33), 14)
        parts.append(f'<tr style="height:{h_px}px;">')

        for c in range(1, last_col + 1):
            entry = merged_map.get((r, c))
            if entry == "skip":
                continue
            cell = ws.cell(row=r, column=c)

            colspan = rowspan = 1
            if isinstance(entry, tuple):
                rowspan, colspan = entry

            value = cell.value
            text = "" if value is None else _escape_html(str(value))

            styles: list[str] = [
                "border:1px solid #e8eaf0",
                "padding:1px 3px",
                "overflow:hidden",
                "box-sizing:border-box",
            ]

            # Fill color
            fill = cell.fill
            if fill is not None and getattr(fill, "fill_type", None) == "solid":
                bg = _hex_from_color_obj(fill.fgColor)
                if bg:
                    styles.append(f"background-color:#{bg}")

            # Font
            font = cell.font
            if font is not None:
                if font.bold:
                    styles.append("font-weight:600")
                if font.size:
                    styles.append(f"font-size:{int(font.size)}px")
                fc = _hex_from_color_obj(font.color)
                if fc:
                    styles.append(f"color:#{fc}")

            # Alignment
            align = cell.alignment
            wrap_text = False
            rotation = 0
            if align is not None:
                if align.horizontal:
                    styles.append(f"text-align:{align.horizontal}")
                if align.vertical:
                    va_map = {"center": "middle", "top": "top", "bottom": "bottom"}
                    styles.append(f"vertical-align:{va_map.get(align.vertical, 'middle')}")
                else:
                    styles.append("vertical-align:middle")
                if align.wrap_text:
                    wrap_text = True
                if align.text_rotation:
                    rotation = int(align.text_rotation)

            if wrap_text:
                styles.append("white-space:normal; word-break:break-word")

            attrs = ""
            if colspan > 1:
                attrs += f' colspan="{colspan}"'
            if rowspan > 1:
                attrs += f' rowspan="{rowspan}"'

            inner = text
            if rotation:
                # writing-mode renders proper vertical text without the
                # transform-then-clip problem of `transform: rotate`.
                # vertical-rl + rotate(180deg) reads bottom→top (Excel +90).
                if rotation == 90 or rotation == -90 or rotation == 270:
                    inner = (
                        '<div style="writing-mode:vertical-rl; '
                        'transform:rotate(180deg); white-space:nowrap; '
                        'display:inline-block;">'
                        f'{text}</div>'
                    )
                else:
                    inner = (
                        f'<div style="display:inline-block; '
                        f'transform:rotate(-{rotation}deg); white-space:nowrap;">'
                        f'{text}</div>'
                    )

            parts.append(
                f'<td style="{"; ".join(styles)};"{attrs}>{inner}</td>'
            )
        parts.append("</tr>")

    parts.append("</table></div>")
    return "".join(parts)


def generate_stripchart(input_xlsx_bytes: bytes, cfg: StripChartConfig) -> bytes:
    """
    Generate a strip chart in the workbook and return the new .xlsx as bytes.
    The original sheets are preserved; a new (or replaced) strip sheet is added.
    """
    log = cfg.log
    wb = openpyxl.load_workbook(BytesIO(input_xlsx_bytes), data_only=True)
    log.append(f"START | sheets={len(wb.sheetnames)}")
    log.append(f"Range: {cfg.from_m} to {cfg.to_m} | interval={cfg.interval_m} | assumeKm={cfg.assume_km}")

    # --- Validate inputs ---
    if cfg.to_m <= cfg.from_m:
        raise ValueError("To chainage must be greater than From chainage.")
    if cfg.interval_m <= 0:
        raise ValueError("Interval must be greater than zero.")
    if not cfg.slots:
        raise ValueError("At least one layer slot is required.")

    # --- Resolve layers ---
    layers: list[tuple[object, str, str]] = []  # (worksheet, name, color_hex)
    for slot_idx, slot in enumerate(cfg.slots, start=1):
        if not slot.sheet_name:
            continue
        if slot.sheet_name not in wb.sheetnames:
            log.append(f"SKIP slot {slot_idx}: data sheet '{slot.sheet_name}' not found")
            continue
        ws_data = wb[slot.sheet_name]
        color = layer_color_for(ws_data, slot_idx)
        layers.append((ws_data, slot.sheet_name, color))
        log.append(f"Layer {slot_idx}: {slot.sheet_name} (color=#{color})")

    n = len(layers)
    if n == 0:
        raise ValueError("No valid layers configured.")

    # --- Validate strip-sheet name doesn't collide with data sheet names ---
    layer_names = [name for _, name, _ in layers]
    if cfg.strip_sheet_name.lower() in (n.lower() for n in layer_names):
        raise ValueError(
            f"Strip sheet name '{cfg.strip_sheet_name}' matches a data sheet — pick a different name."
        )

    # --- Compute layout ---
    block_rows = 1 + n + 1 + n + 1
    metres_per_block = COLS_PER_BLOCK * cfg.interval_m
    num_blocks = -(-(cfg.to_m - cfg.from_m) // metres_per_block)  # ceil
    total_cols = num_blocks * COLS_PER_BLOCK
    last_col = C_DATA_START + COLS_PER_BLOCK - 1

    # --- Recreate strip sheet ---
    if cfg.strip_sheet_name in wb.sheetnames:
        del wb[cfg.strip_sheet_name]
    ws_strip = wb.create_sheet(title=cfg.strip_sheet_name)

    # --- Column widths ---
    ws_strip.column_dimensions["A"].width = 5
    ws_strip.column_dimensions["B"].width = 11
    for c in range(C_DATA_START, last_col + 1):
        ws_strip.column_dimensions[get_column_letter(c)].width = 3.2

    # --- Build all blocks ---
    layer_colors = [c for _, _, c in layers]
    for b in range(num_blocks):
        block_top = b * block_rows + 1
        block_start_m = cfg.from_m + b * metres_per_block
        build_one_block(ws_strip, block_top, block_start_m, cfg.interval_m,
                         last_col, layer_names, layer_colors, n)
    log.append(f"Built {num_blocks} block(s), {n} layer(s) per block")

    # --- Tracking grids for audit (1-indexed by layer) ---
    painted_rhs = [[False] * total_cols for _ in range(n + 1)]
    painted_lhs = [[False] * total_cols for _ in range(n + 1)]
    overlap_rhs = [[False] * total_cols for _ in range(n + 1)]
    overlap_lhs = [[False] * total_cols for _ in range(n + 1)]

    # --- Paint each layer ---
    total_painted = 0
    for layer_idx, (ws_data, name, color) in enumerate(layers, start=1):
        from_col, to_col, side_col = find_header_columns(ws_data)
        start_row = find_data_start_row(ws_data)
        if start_row is None:
            log.append(f"  [{name}] no numeric SL.NO found — sheet skipped")
            continue
        last_row = find_last_data_row(ws_data, start_row)
        log.append(
            f"  [{name}] startRow={start_row} | cols: From={get_column_letter(from_col)} "
            f"To={get_column_letter(to_col)} Side={get_column_letter(side_col)}"
        )

        attempted = skipped = painted = 0
        for r in range(start_row, last_row + 1):
            attempted += 1
            f_val = ws_data.cell(row=r, column=from_col).value
            t_val = ws_data.cell(row=r, column=to_col).value
            f_m = parse_chainage(f_val, cfg.assume_km)
            t_m = parse_chainage(t_val, cfg.assume_km)
            if f_m is None or t_m is None or t_m <= f_m:
                skipped += 1
                continue
            if t_m <= cfg.from_m or f_m >= cfg.to_m:
                skipped += 1
                continue

            side_raw = ws_data.cell(row=r, column=side_col).value
            side = (str(side_raw) if side_raw is not None else "").strip().upper()
            if side not in RHS_SIDES and side not in LHS_SIDES:
                skipped += 1
                continue
            is_rhs = side in RHS_SIDES
            is_lhs = side in LHS_SIDES

            first_col_idx = (f_m - cfg.from_m) // cfg.interval_m
            last_col_idx = (t_m - cfg.from_m - 1) // cfg.interval_m
            if first_col_idx < 0:
                first_col_idx = 0
            if last_col_idx >= total_cols:
                last_col_idx = total_cols - 1
            if last_col_idx < first_col_idx:
                continue

            for col_idx in range(first_col_idx, last_col_idx + 1):
                block_no = col_idx // COLS_PER_BLOCK
                col_in_block = col_idx % COLS_PER_BLOCK
                block_top = block_no * block_rows + 1
                cell_col = C_DATA_START + col_in_block

                if is_rhs:
                    cell_row = block_top + (n - layer_idx + 1)
                    if painted_rhs[layer_idx][col_idx]:
                        ws_strip.cell(row=cell_row, column=cell_col).fill = _solid_fill(COLOR_RED)
                        overlap_rhs[layer_idx][col_idx] = True
                    else:
                        ws_strip.cell(row=cell_row, column=cell_col).fill = _solid_fill(color)
                        painted_rhs[layer_idx][col_idx] = True
                    painted += 1
                if is_lhs:
                    cell_row = block_top + n + 1 + layer_idx
                    if painted_lhs[layer_idx][col_idx]:
                        ws_strip.cell(row=cell_row, column=cell_col).fill = _solid_fill(COLOR_RED)
                        overlap_lhs[layer_idx][col_idx] = True
                    else:
                        ws_strip.cell(row=cell_row, column=cell_col).fill = _solid_fill(color)
                        painted_lhs[layer_idx][col_idx] = True
                    painted += 1

        total_painted += painted
        log.append(f"Layer {layer_idx} [{name}]: rows_seen={attempted} "
                   f"skipped={skipped} cells_painted={painted}")

    # --- Audit pass: per-layer per-side per-block, write to col (last_col+2) ---
    audit_col = last_col + 2
    ws_strip.column_dimensions[get_column_letter(audit_col)].width = 70

    issue_count = 0
    for b in range(num_blocks):
        block_start_col = b * COLS_PER_BLOCK
        block_end_col = block_start_col + COLS_PER_BLOCK - 1
        block_top = b * block_rows + 1

        for layer_idx in range(1, n + 1):
            rhs_row = block_top + (n - layer_idx + 1)
            lhs_row = block_top + n + 1 + layer_idx

            # Build per-side audit text
            rhs_msg_parts: list[str] = []
            ov = ranges_in_block(overlap_rhs, layer_idx, block_start_col, block_end_col,
                                  total_cols, cfg.from_m, cfg.interval_m)
            if ov:
                rhs_msg_parts.append(f"Overlap: {ov}")
            if layer_idx >= 2:
                ml = missing_lower_in_block(painted_rhs, layer_idx,
                                             block_start_col, block_end_col,
                                             total_cols, cfg.from_m, cfg.interval_m)
                if ml:
                    rhs_msg_parts.append(
                        f"Missing lower '{layer_names[layer_idx - 2]}' at {ml}"
                    )
            rhs_msg = " | ".join(rhs_msg_parts)

            lhs_msg_parts: list[str] = []
            ov = ranges_in_block(overlap_lhs, layer_idx, block_start_col, block_end_col,
                                  total_cols, cfg.from_m, cfg.interval_m)
            if ov:
                lhs_msg_parts.append(f"Overlap: {ov}")
            if layer_idx >= 2:
                ml = missing_lower_in_block(painted_lhs, layer_idx,
                                             block_start_col, block_end_col,
                                             total_cols, cfg.from_m, cfg.interval_m)
                if ml:
                    lhs_msg_parts.append(
                        f"Missing lower '{layer_names[layer_idx - 2]}' at {ml}"
                    )
            lhs_msg = " | ".join(lhs_msg_parts)

            # Paint missing-lower cells purple (overrides existing fill)
            if layer_idx >= 2:
                paint_missing_in_block(ws_strip, painted_rhs, layer_idx, n, block_rows,
                                        True, block_start_col, block_end_col,
                                        total_cols, COLOR_PURPLE)
                paint_missing_in_block(ws_strip, painted_lhs, layer_idx, n, block_rows,
                                        False, block_start_col, block_end_col,
                                        total_cols, COLOR_PURPLE)

            if rhs_msg:
                write_audit(ws_strip, rhs_row, audit_col, rhs_msg, True)
                issue_count += 1
            elif has_any_paint(painted_rhs, layer_idx, block_start_col, block_end_col, total_cols):
                write_audit(ws_strip, rhs_row, audit_col, "OK", False)

            if lhs_msg:
                write_audit(ws_strip, lhs_row, audit_col, lhs_msg, True)
                issue_count += 1
            elif has_any_paint(painted_lhs, layer_idx, block_start_col, block_end_col, total_cols):
                write_audit(ws_strip, lhs_row, audit_col, "OK", False)

    log.append(f"DONE | layers={n} | blocks={num_blocks} | "
               f"total_painted={total_painted} | audit_issues={issue_count}")

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
