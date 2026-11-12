"""
StripChart Plot tab — LIVE Excel version.

Workflow:
  1. User opens their .xlsx file (and StripChart_macros.xlsm) in Excel
  2. Dashboard lists open workbooks via xlwings
  3. User picks workbook, enters From / To / Interval, fills up to 7 slots
  4. Dashboard writes a hidden "_SC_Cfg" sheet into the chosen workbook
  5. Dashboard invokes StripChart_Run via direct AppleScript (osascript)
  6. The macro reads config, paints the combined strip chart, runs the audit
"""

from __future__ import annotations

import os
import platform
import subprocess
import tempfile
from typing import Optional

import streamlit as st

from core.excel_reader import (get_open_workbooks, get_sheets,
                                excel_running, xlwings_available, _xw)
from ui.theme import section_title, alert

IS_MAC = platform.system() == "Darwin"
IS_WIN = platform.system() == "Windows"

CFG_SHEET = "_SC_Cfg"
NUM_SLOTS = 7


# ─────────────────────────────────────────────────────────────
#  Chainage parsing  (km+m or km.m → total metres)
#    "185+000" → 185000     "185+5"   → 185005
#    "185.000" → 185000     "185.5"   → 185500   (decimal km)
#    "185"     → 185000
# ─────────────────────────────────────────────────────────────
def parse_chainage(s: str) -> Optional[int]:
    s = (s or "").strip()
    if not s:
        return None

    if "+" in s:
        try:
            km_str, m_str = s.split("+", 1)
            return int(km_str) * 1000 + int(m_str)
        except ValueError:
            return None

    if "." in s:
        try:
            km_str, frac_str = s.split(".", 1)
            frac_str = (frac_str + "000")[:3]
            return int(km_str) * 1000 + int(frac_str)
        except ValueError:
            return None

    try:
        return int(s) * 1000
    except ValueError:
        return None


def fmt_chainage(metres: int) -> str:
    return f"{metres // 1000}+{metres % 1000:03d}"


# ─────────────────────────────────────────────────────────────
#  Macro invocation  (Mac → osascript; Windows → xlwings)
# ─────────────────────────────────────────────────────────────
def _run_vba(workbook_name: str, macro: str) -> tuple[bool, str]:
    """Invoke the VBA macro. Returns (ok, error_message).

    The macro must live in an open .xlsm. We qualify the macro name with
    the macro workbook so Excel finds it regardless of which workbook is
    active. The macro itself locates the user's data workbook (the one
    with _SC_Cfg) by scanning open books.
    """
    macro_book_name = None
    xw = _xw()
    if xw is not None:
        try:
            app = xw.apps.active
            if app is not None:
                for b in app.books:
                    if b.name.lower().endswith(".xlsm"):
                        macro_book_name = b.name
                        try:
                            b.activate()
                        except Exception:
                            pass
                        break
        except Exception:
            pass

    if IS_MAC:
        macro_escaped = macro.replace('"', '\\"')
        book_escaped = (macro_book_name or "").replace('"', '\\"')

        # Activate the macro .xlsm via AppleScript (xlwings's activate is
        # unreliable on Mac), then call the macro by its unqualified name —
        # which is the form Excel for Mac handles best. Wrap with try and
        # a 30-minute timeout.
        activate_line = ""
        if book_escaped:
            activate_line = (
                f'            activate object workbook "{book_escaped}"\n'
            )

        script = (
            'tell application "Microsoft Excel"\n'
            '    with timeout of 1800 seconds\n'
            '        try\n'
            f'{activate_line}'
            f'            run VB Macro "{macro_escaped}"\n'
            '            return "OK"\n'
            '        on error errMsg number errNum\n'
            '            return "ERR " & errNum & ": " & errMsg\n'
            '        end try\n'
            '    end timeout\n'
            'end tell\n'
        )

        script_path = None
        try:
            with tempfile.NamedTemporaryFile(
                "w", suffix=".applescript", delete=False
            ) as f:
                f.write(script)
                script_path = f.name

            result = subprocess.run(
                ["osascript", script_path],
                capture_output=True, text=True, timeout=1800,
            )
        except subprocess.TimeoutExpired:
            return False, "Macro call timed out after 30 minutes."
        except Exception as e:
            return False, f"osascript subprocess failed: {e}"
        finally:
            if script_path and os.path.exists(script_path):
                try:
                    os.unlink(script_path)
                except OSError:
                    pass

        stdout = (result.stdout or "").strip()
        stderr = (result.stderr or "").strip()
        if result.returncode != 0:
            return False, (
                f"AppleScript error: {stderr or stdout or 'unknown'}\n\n"
                f"Script attempted:\n{script.strip()}"
            )
        if stdout.startswith("ERR"):
            return False, (
                f"Excel reported: {stdout}\n\n"
                f"Script attempted:\n{script.strip()}"
            )
        return True, ""

    # Windows fallback — use xlwings
    if xw is None:
        return False, "xlwings not installed."
    try:
        app = xw.apps.active
        if app is None:
            return False, "No active Excel app."
        macro_fn = app.macro(macro)
        macro_fn()
        return True, ""
    except Exception as e:
        return False, str(e)


# ─────────────────────────────────────────────────────────────
#  Write config sheet, save workbook, run macro, read log
# ─────────────────────────────────────────────────────────────
def _execute(workbook_name: str, from_m: int, to_m: int, interval_m: int,
             slots: list[tuple[str, str]], assume_km: bool) -> dict:
    xw = _xw()
    if xw is None:
        return {"ok": False, "error": "xlwings not installed."}

    book = None
    for app in xw.apps:
        for b in app.books:
            if b.name == workbook_name:
                book = b
                break
        if book:
            break
    if book is None:
        return {"ok": False, "error": f"Workbook '{workbook_name}' not found."}

    # 1. Create / reset hidden config sheet
    cfg = None
    for sh in book.sheets:
        if sh.name == CFG_SHEET:
            cfg = sh
            break
    if cfg is None:
        cfg = book.sheets.add(name=CFG_SHEET, after=book.sheets[-1])
    else:
        cfg.clear()

    # 2. Write config values
    cfg.range("A1").value = "from_m"
    cfg.range("B1").value = from_m
    cfg.range("A2").value = "to_m"
    cfg.range("B2").value = to_m
    cfg.range("A3").value = "interval_m"
    cfg.range("B3").value = interval_m
    cfg.range("A4").value = "assume_km"
    cfg.range("B4").value = "1" if assume_km else "0"

    for i, (ds, ss) in enumerate(slots):
        row = 5 + i
        cfg.range(f"A{row}").value = i + 1
        cfg.range(f"B{row}").value = ds or ""
        cfg.range(f"C{row}").value = ss or ""

    try:
        cfg.api.Visible = 0  # xlSheetHidden
    except Exception:
        pass

    # 3. Activate the target workbook so macro sees correct ActiveWorkbook
    try:
        book.activate()
    except Exception:
        pass

    # 4. Save (xlsx is fine — only the macro file needs to be .xlsm)
    try:
        book.save()
    except Exception as e:
        return {"ok": False, "error": f"Could not save workbook: {e}"}

    # 5. Snapshot sheets before macro
    before = {sh.name for sh in book.sheets}

    # 6. List open .xlsm files for the error message
    available_macros_in = []
    for app2 in xw.apps:
        for b2 in app2.books:
            if b2.name.lower().endswith(".xlsm"):
                available_macros_in.append(b2.name)

    ok, err = _run_vba(workbook_name, "StripChart_Run")
    if not ok:
        return {"ok": False, "error": err,
                "diag": {"xlsm_files_open": available_macros_in}}

    # 7. Read macro log from col Z of _SC_Cfg
    log_lines: list[str] = []
    try:
        cfg2 = book.sheets[CFG_SHEET]
        zcol = cfg2.range("Z1:Z200").value
        if isinstance(zcol, list):
            log_lines = [str(v) for v in zcol if v is not None and str(v).strip()]
        elif zcol:
            log_lines = [str(zcol)]
    except Exception:
        pass

    after = {sh.name for sh in book.sheets}
    new_sheets = sorted(after - before)

    macro_done = any(ln.startswith("DONE") for ln in log_lines)
    macro_aborted = any(ln.startswith("ABORT") or ln.startswith("FATAL")
                        for ln in log_lines)

    if macro_aborted:
        return {
            "ok": False,
            "error": "Macro aborted:\n\n" + "\n".join(f"  • {ln}" for ln in log_lines),
            "diag": {"xlsm_files_open": available_macros_in,
                     "log": log_lines, "new_sheets": new_sheets},
        }

    if not macro_done and not new_sheets:
        msg = "Macro ran but produced no result.\n\n"
        if log_lines:
            msg += "Macro log:\n" + "\n".join(f"  • {ln}" for ln in log_lines)
        else:
            msg += ("Could not read macro log. Likely causes:\n"
                    "  1. Master .xlsm with StripChart_Run isn't open\n"
                    "  2. Macro 'StripChart_Run' signature is outdated — re-import the latest StripChart.bas")
        return {
            "ok": False, "error": msg,
            "diag": {"xlsm_files_open": available_macros_in,
                     "log": log_lines, "new_sheets": []},
        }

    return {"ok": True, "error": "",
            "diag": {"log": log_lines, "new_sheets": new_sheets}}


# ─────────────────────────────────────────────────────────────
#  Streamlit UI
# ─────────────────────────────────────────────────────────────
def _ensure_state():
    if "sc_slots" not in st.session_state:
        st.session_state.sc_slots = [("", "")] * NUM_SLOTS


def render_stripchart_tab() -> None:
    _ensure_state()

    # ── Workbook picker ─────────────────────────────────────
    section_title("Select Workbook",
                  hint="Pick the workbook you have open in Excel")

    if not xlwings_available():
        alert("err", "xlwings not installed. Run: <code>.venv/bin/pip install xlwings</code>",
              icon="❌")
        return
    if not excel_running():
        alert("warn", "Excel is not running. Open your file in Excel and try again.",
              icon="⚠️")
        return

    books = get_open_workbooks()
    if not books:
        alert("warn", "No open workbooks. Open your file in Excel and refresh.",
              icon="⚠️")
        return

    c1, c2 = st.columns([5, 1])
    with c1:
        wb = st.selectbox("Workbook", options=books, label_visibility="collapsed",
                           key="sc_workbook")
    with c2:
        st.write("")
        if st.button("🔄 Refresh", use_container_width=True, key="sc_refresh"):
            st.rerun()

    sheets = get_sheets(wb)

    st.write("")

    # ── Range inputs ────────────────────────────────────────
    section_title("Chainage Range",
                  hint="Both formats accepted: 185+000 or 185.000 (= 185000 m)")

    cA, cB, cC = st.columns(3)
    with cA:
        from_str = st.text_input("From",   value="99+700", key="sc_from",
                                  placeholder="e.g. 99+700 or 99.700")
    with cB:
        to_str   = st.text_input("To",     value="119+015", key="sc_to",
                                  placeholder="e.g. 119+015 or 119.015")
    with cC:
        interval = st.number_input("Interval (metres)",
                                    min_value=1, max_value=1000,
                                    value=10, step=10, key="sc_interval")

    from_m = parse_chainage(from_str)
    to_m   = parse_chainage(to_str)

    valid_range = (from_m is not None and to_m is not None and to_m > from_m)
    if from_str and from_m is None:
        alert("err", f"Cannot parse From chainage: '{from_str}'", icon="❌")
    if to_str and to_m is None:
        alert("err", f"Cannot parse To chainage: '{to_str}'", icon="❌")
    if from_m is not None and to_m is not None and to_m <= from_m:
        alert("err", "To chainage must be greater than From chainage.", icon="❌")

    if valid_range:
        total_m = to_m - from_m
        cols_per_block = 50
        metres_per_block = cols_per_block * interval
        num_blocks = (total_m + metres_per_block - 1) // metres_per_block
        st.markdown(
            f"<div class='card' style='padding:10px 14px;color:var(--ink-700);'>"
            f"📐 Range: <b>{fmt_chainage(from_m)}</b> → <b>{fmt_chainage(to_m)}</b> "
            f"({total_m/1000:.3f} km) · "
            f"interval <b>{interval} m</b> · "
            f"<b>{cols_per_block} columns</b>/block · "
            f"<b>{num_blocks}</b> block(s) stacked vertically"
            f"</div>",
            unsafe_allow_html=True,
        )

    assume_km = st.checkbox(
        "Data-sheet From/To values are in kilometres (e.g. 152.78 = 152.78 km)",
        value=True, key="sc_assume_km",
        help="Numbers ≥ 1000 are auto-detected as metres regardless of this setting."
    )

    st.write("")

    # ── Slot configuration ──────────────────────────────────
    section_title("Layer Slots",
                  hint="Up to 7 sheets — leave a slot blank to skip it. First selected = bottom layer.")

    sheet_options = [""] + [s for s in sheets if not s.startswith("_") and "-Strip" not in s]
    new_slots: list[tuple[str, str]] = []

    h1, h2, h3 = st.columns([0.5, 3, 3])
    with h2:
        st.caption("Data sheet (first selected = bottom layer)")
    with h3:
        st.caption("Strip sheet name (Slot 1 used as combined chart name)")

    for i in range(NUM_SLOTS):
        cur_ds, cur_ss = (st.session_state.sc_slots[i]
                          if i < len(st.session_state.sc_slots) else ("", ""))
        slot_col1, slot_col2, slot_col3 = st.columns(
            [0.5, 3, 3], vertical_alignment="center"
        )
        with slot_col1:
            st.markdown(
                f"<div style='color:var(--ink-500);font-weight:600;'>Slot {i+1}</div>",
                unsafe_allow_html=True,
            )
        with slot_col2:
            ds_idx = sheet_options.index(cur_ds) if cur_ds in sheet_options else 0
            ds = st.selectbox(f"Data sheet slot {i+1}", options=sheet_options,
                               index=ds_idx, key=f"sc_ds_{i}",
                               label_visibility="collapsed")
        with slot_col3:
            default_ss = (ds + "-Strip") if ds and not cur_ss else cur_ss
            ss = st.text_input(f"Strip name slot {i+1}",
                                value=default_ss, key=f"sc_ss_{i}",
                                placeholder="e.g. HW-Strip",
                                label_visibility="collapsed")
        new_slots.append((ds, ss))

    st.session_state.sc_slots = new_slots

    populated = [(ds, ss) for ds, ss in new_slots if ds]
    if populated:
        st.markdown(
            f"<div class='card' style='padding:10px 14px;color:var(--ink-700);font-size:13px;'>"
            f"✅ {len(populated)} slot(s) will be processed: "
            + ", ".join(f"<b>{ds}</b> → <i>{ss or ds + '-Strip'}</i>"
                        for ds, ss in populated) +
            f"</div>",
            unsafe_allow_html=True,
        )

    st.write("")

    # ── Run button ──────────────────────────────────────────
    if st.button("🚀 Run StripChart Plot", type="primary", use_container_width=False,
                 key="sc_run"):
        if not valid_range:
            alert("err", "Fix the chainage range above first.", icon="❌")
            return
        if not populated:
            alert("warn", "Configure at least one slot.", icon="⚠️")
            return
        full_slots = [(ds, ss) for ds, ss in new_slots]

        with st.spinner("Writing config + running macro in Excel…"):
            res = _execute(wb, from_m, to_m, int(interval),
                            full_slots, assume_km)

        diag = res.get("diag", {})
        xlsm_open = diag.get("xlsm_files_open", [])

        if res["ok"]:
            new_sheets = diag.get("new_sheets", [])
            log = diag.get("log", [])
            if any("total_painted=0" in ln for ln in log):
                alert("warn",
                      f"Nothing was painted — no data rows fall inside "
                      f"<b>{fmt_chainage(from_m)}</b> → <b>{fmt_chainage(to_m)}</b>. "
                      f"Check the From/To range matches this project's chainages.",
                      icon="⚠️")
            alert("ok",
                  f"✅ Strip chart generated in <b>{wb}</b>." +
                  (f" New sheet(s): " + ", ".join(f"<code>{s}</code>" for s in new_sheets)
                   if new_sheets else " (existing strip sheet rebuilt)"),
                  icon="")
            if log:
                with st.expander("📜 Macro log"):
                    for line in log:
                        st.code(line, language=None)
        else:
            err = res.get("error", "Unknown error")
            extra = ""
            if xlsm_open:
                extra = f"<br/><br/>📂 Open .xlsm files (must contain the macro): {', '.join(xlsm_open) or '— none —'}"
            else:
                extra = "<br/><br/>⚠️ <b>NO .xlsm file is currently open in Excel.</b> Open StripChart_macros.xlsm first."
            alert("err",
                  f"<pre style='white-space:pre-wrap;'>{err}</pre>{extra}",
                  icon="❌")
